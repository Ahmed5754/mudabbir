"""Encrypted credential storage for Mudabbir.

Changes:
  - 2026-02-06: Initial implementation — Fernet encryption with machine-derived PBKDF2 key.

Stores API keys and tokens in ~/.Mudabbir/secrets.enc instead of plaintext config.json.
Encryption key derived from machine identity (hostname + MAC + username) so the encrypted
file only works on the same machine/user. Salt stored in ~/.Mudabbir/.salt.
"""

import base64
import hashlib
import json
import logging
import os
import platform
import uuid
from functools import lru_cache
from pathlib import Path

from cryptography.fernet import Fernet, InvalidToken

logger = logging.getLogger(__name__)

# Fields that are considered secrets and must be stored encrypted.
SECRET_FIELDS: frozenset[str] = frozenset(
    {
        "telegram_bot_token",
        "openai_api_key",
        "anthropic_api_key",
        "openai_compatible_api_key",
        "discord_bot_token",
        "slack_bot_token",
        "slack_app_token",
        "whatsapp_access_token",
        "whatsapp_verify_token",
        "tavily_api_key",
        "brave_search_api_key",
        "parallel_api_key",
        "elevenlabs_api_key",
        "google_api_key",
        "google_oauth_client_id",
        "google_oauth_client_secret",
        "spotify_client_id",
        "spotify_client_secret",
        "matrix_access_token",
        "matrix_password",
        "teams_app_id",
        "teams_app_password",
        "gchat_service_account_key",
        "sarvam_api_key",
    }
)


def _ensure_permissions(path: Path, mode: int = 0o600) -> None:
    """Set strict file permissions (owner read/write only)."""
    if not path.exists():
        return
    try:
        path.chmod(mode)
    except OSError:
        # Windows doesn't support chmod the same way — skip silently
        pass


def _ensure_dir_permissions(path: Path) -> None:
    """Set strict directory permissions (owner rwx only)."""
    _ensure_permissions(path, mode=0o700)


class CredentialStore:
    """Encrypted credential store backed by Fernet + PBKDF2.

    Storage:
      - ~/.Mudabbir/secrets.enc  (Fernet-encrypted JSON)
      - ~/.Mudabbir/.salt        (16-byte random salt, auto-generated)

    The encryption key is derived from:
      platform.node() + uuid.getnode() + os.getlogin()
    so the file is bound to the current machine and user account.
    """

    def __init__(self, config_dir: Path | None = None):
        if config_dir is None:
            config_dir = Path.home() / ".Mudabbir"
        self._config_dir = config_dir
        self._secrets_path = config_dir / "secrets.enc"
        self._salt_path = config_dir / ".salt"
        self._cache: dict[str, str] | None = None

    def _get_machine_id(self) -> str:
        """Return a persistent machine identifier.

        Tries (in order):
          1. /etc/machine-id  (Linux — systemd)
          2. /var/lib/dbus/machine-id  (Linux — older dbus)
          3. platform.node()  (hostname — fallback)

        uuid.getnode() is intentionally NOT used because it returns a
        random MAC on systems without a discoverable NIC, producing a
        different value on every process start.
        """
        for p in ("/etc/machine-id", "/var/lib/dbus/machine-id"):
            try:
                mid = Path(p).read_text().strip()
                if mid:
                    return mid
            except OSError:
                continue
        return platform.node()

    def _identity_candidates(self) -> list[bytes]:
        """Build candidate identities for decrypt compatibility.

        Why multiple:
        - os.getlogin() can vary/fail in headless or elevated sessions.
        - USER/USERNAME env vars may differ across launch contexts.

        We try a small ordered set and re-encrypt with the primary identity
        after successful fallback decrypt.
        """
        machine_id = self._get_machine_id()
        host_name = platform.node()
        mac_int = uuid.getnode()
        mac_variants = [str(mac_int), f"{mac_int:012x}"]
        candidates: list[str] = []

        # Primary identity (current behavior).
        try:
            login_name = os.getlogin()
            if login_name:
                candidates.append(f"{machine_id}|{login_name}")
        except OSError:
            pass

        # Common fallbacks across shells/services.
        for env_key in ("USER", "USERNAME", "LOGNAME"):
            value = os.environ.get(env_key, "")
            if value:
                candidates.append(f"{machine_id}|{value}")

        # Legacy format (v0.4.2 and earlier): "hostname|mac|username"
        # Keep this for backward compatibility with existing secrets.enc files.
        user_names: list[str] = []
        for env_key in ("USER", "USERNAME", "LOGNAME"):
            value = os.environ.get(env_key, "")
            if value and value not in user_names:
                user_names.append(value)
        try:
            login_name = os.getlogin()
            if login_name and login_name not in user_names:
                user_names.insert(0, login_name)
        except OSError:
            pass
        for uname in user_names:
            for mac_value in mac_variants:
                candidates.append(f"{host_name}|{mac_value}|{uname}")

        # App-name fallbacks (rebrand compatibility).
        # Keep old names so existing secrets.enc can still decrypt after renaming.
        candidates.append(f"{machine_id}|Mudabbir")
        candidates.append(f"{machine_id}|mudabbir")
        candidates.append(f"{machine_id}|Mudabbir")
        candidates.append(f"{machine_id}|pocketclaw")

        # De-duplicate while preserving order.
        deduped: list[bytes] = []
        seen: set[str] = set()
        for c in candidates:
            key = c.strip()
            if not key or key in seen:
                continue
            seen.add(key)
            deduped.append(key.encode("utf-8"))
        return deduped

    def _derive_key_for_identity(self, identity: bytes) -> bytes:
        """Derive a Fernet key from a specific machine identity candidate.

        Uses hashlib.pbkdf2_hmac for maximum compatibility across OpenSSL/
        cryptography builds on Windows while preserving key derivation output.
        """
        salt = self._get_or_create_salt()
        raw_key = hashlib.pbkdf2_hmac("sha256", identity, salt, 480_000, dklen=32)
        return base64.urlsafe_b64encode(raw_key)

    def _get_or_create_salt(self) -> bytes:
        """Load existing salt or generate a new 16-byte salt."""
        self._config_dir.mkdir(parents=True, exist_ok=True)
        _ensure_dir_permissions(self._config_dir)

        if self._salt_path.exists():
            salt = self._salt_path.read_bytes()
            if len(salt) >= 16:
                return salt[:16]

        salt = os.urandom(16)
        self._salt_path.write_bytes(salt)
        _ensure_permissions(self._salt_path)
        return salt

    def _derive_key(self) -> bytes:
        """Derive a Fernet key from machine identity + salt via PBKDF2."""
        candidates = self._identity_candidates()
        primary = candidates[0] if candidates else b"Mudabbir"
        return self._derive_key_for_identity(primary)

    def _load(self) -> dict[str, str]:
        """Decrypt and load secrets from disk."""
        if self._cache is not None:
            return self._cache

        if not self._secrets_path.exists():
            self._cache = {}
            return self._cache

        encrypted = self._secrets_path.read_bytes()
        last_exc: Exception | None = None

        for idx, identity in enumerate(self._identity_candidates()):
            try:
                fernet = Fernet(self._derive_key_for_identity(identity))
                decrypted = fernet.decrypt(encrypted)
                data = json.loads(decrypted)
                if not isinstance(data, dict):
                    raise json.JSONDecodeError("not-an-object", str(data), 0)
                self._cache = data

                # If a fallback identity worked, heal by re-encrypting with primary.
                if idx > 0:
                    logger.warning(
                        "Recovered secrets.enc using fallback identity candidate (%d). "
                        "Re-encrypting with current identity.",
                        idx + 1,
                    )
                    try:
                        self._save(dict(data))
                    except Exception:
                        pass
                return self._cache
            except (InvalidToken, json.JSONDecodeError, Exception) as exc:
                last_exc = exc
                continue

        if last_exc is None:
            reason = "unknown error"
        else:
            msg = str(last_exc).strip()
            reason = f"{type(last_exc).__name__}: {msg}" if msg else type(last_exc).__name__
        logger.warning(
            "Failed to decrypt secrets.enc (machine changed? corrupted?): %s. "
            "Starting with empty credential store.",
            reason,
        )
        self._cache = {}

        return self._cache

    def _save(self, data: dict[str, str]) -> None:
        """Encrypt and write secrets to disk."""
        self._config_dir.mkdir(parents=True, exist_ok=True)
        _ensure_dir_permissions(self._config_dir)

        fernet = Fernet(self._derive_key())
        plaintext = json.dumps(data).encode("utf-8")
        encrypted = fernet.encrypt(plaintext)
        self._secrets_path.write_bytes(encrypted)
        _ensure_permissions(self._secrets_path)
        self._cache = data

    def get(self, name: str) -> str | None:
        """Get a secret by name. Returns None if not found."""
        data = self._load()
        return data.get(name)

    def set(self, name: str, value: str) -> None:
        """Store a secret."""
        data = self._load()
        data[name] = value
        self._save(data)

    def delete(self, name: str) -> None:
        """Remove a secret."""
        data = self._load()
        if name in data:
            del data[name]
            self._save(data)

    def get_all(self) -> dict[str, str]:
        """Get a copy of all stored secrets."""
        return dict(self._load())

    def clear_cache(self) -> None:
        """Force re-read from disk on next access."""
        self._cache = None


@lru_cache
def get_credential_store() -> CredentialStore:
    """Get the singleton CredentialStore instance."""
    return CredentialStore()
