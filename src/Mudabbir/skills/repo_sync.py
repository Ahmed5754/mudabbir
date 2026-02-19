"""Sync external GitHub skill repositories into Mudabbir skill directory."""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urlparse

import yaml

logger = logging.getLogger(__name__)

_SKIP_DIR_NAMES = {
    ".git",
    ".hg",
    ".svn",
    ".venv",
    "venv",
    "node_modules",
    "__pycache__",
    "dist",
    "build",
}

_WINDOWS_HINTS = (
    "windows",
    "win32",
    "powershell",
    "active directory",
    "group policy",
    "uia",
    "desktop",
    "registry",
    "defender",
    "task scheduler",
    "event log",
    "credential",
    "privilege",
    "uac",
    "microsoft",
    "cmd.exe",
    "terminal",
)


@dataclass
class RepoSyncResult:
    source: str
    repo: str
    installed: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    generated: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def installed_count(self) -> int:
        return len(self.installed)


def _slug(value: str, max_len: int = 64) -> str:
    text = re.sub(r"[^a-z0-9]+", "-", value.casefold()).strip("-")
    text = re.sub(r"-{2,}", "-", text)
    if not text:
        text = "skill"
    if len(text) <= max_len:
        return text
    digest = hashlib.sha1(text.encode("utf-8")).hexdigest()[:8]
    base = text[: max(8, max_len - 9)].rstrip("-")
    return f"{base}-{digest}"


def _parse_source(source: str) -> tuple[str, str]:
    src = (source or "").strip()
    if not src:
        raise ValueError("Empty source")

    if src.startswith("http://") or src.startswith("https://"):
        parsed = urlparse(src)
        if "github.com" not in parsed.netloc.casefold():
            raise ValueError("Only github.com sources are supported")
        parts = [p for p in parsed.path.split("/") if p]
        if len(parts) < 2:
            raise ValueError("Invalid GitHub URL format")
        return parts[0], parts[1].removesuffix(".git")

    parts = [p for p in src.split("/") if p]
    if len(parts) < 2:
        raise ValueError("Source must be owner/repo or full GitHub URL")
    return parts[0], parts[1].removesuffix(".git")


def _iter_skill_dirs(repo_root: Path) -> list[Path]:
    skill_dirs: list[Path] = []
    for skill_md in repo_root.rglob("SKILL.md"):
        try:
            rel_parts = skill_md.relative_to(repo_root).parts
        except Exception:
            continue
        if any(part in _SKIP_DIR_NAMES for part in rel_parts):
            continue
        skill_dirs.append(skill_md.parent)
    return skill_dirs


def _is_windows_skill(skill_dir: Path, repo_root: Path) -> bool:
    rel = str(skill_dir.relative_to(repo_root)).replace("\\", "/").casefold()
    try:
        content = (skill_dir / "SKILL.md").read_text(encoding="utf-8", errors="ignore")
    except Exception:
        content = ""
    haystack = f"{rel}\n{content[:7000]}".casefold()
    return any(token in haystack for token in _WINDOWS_HINTS)


def _rewrite_skill_name(skill_md_path: Path, skill_name: str) -> None:
    text = skill_md_path.read_text(encoding="utf-8", errors="ignore")
    match = re.match(r"^---\s*\n(.*?)\n---\s*\n(.*)$", text, re.DOTALL)
    if match:
        try:
            frontmatter = yaml.safe_load(match.group(1)) or {}
            body = match.group(2).lstrip()
        except Exception:
            frontmatter = {}
            body = text
    else:
        frontmatter = {}
        body = text

    frontmatter["name"] = skill_name
    if "description" not in frontmatter:
        frontmatter["description"] = f"Imported skill: {skill_name}"

    yaml_text = yaml.safe_dump(frontmatter, allow_unicode=True, sort_keys=False).strip()
    patched = f"---\n{yaml_text}\n---\n\n{body.rstrip()}\n"
    skill_md_path.write_text(patched, encoding="utf-8")


def _install_skill_dir(
    *,
    owner: str,
    repo: str,
    repo_root: Path,
    skill_dir: Path,
    dest_root: Path,
) -> str:
    rel = str(skill_dir.relative_to(repo_root)).replace("\\", "/")
    rel_slug = _slug(rel, max_len=42)
    repo_slug = _slug(f"{owner}-{repo}", max_len=18)
    installed_name = _slug(f"{repo_slug}-{rel_slug}", max_len=64)

    destination = dest_root / installed_name
    if destination.exists():
        shutil.rmtree(destination)
    shutil.copytree(skill_dir, destination)

    skill_md = destination / "SKILL.md"
    if skill_md.exists():
        _rewrite_skill_name(skill_md, installed_name)
    return installed_name


def _generate_bridge_skill(
    *,
    owner: str,
    repo: str,
    repo_root: Path,
    dest_root: Path,
) -> str:
    repo_slug = _slug(f"{owner}-{repo}", max_len=48)
    skill_name = _slug(f"{repo_slug}-bridge", max_len=64)
    destination = dest_root / skill_name
    destination.mkdir(parents=True, exist_ok=True)

    readme = ""
    for candidate in ("README.md", "README.MD", "readme.md", "README.rst"):
        p = repo_root / candidate
        if p.exists():
            readme = p.read_text(encoding="utf-8", errors="ignore")
            break
    excerpt = "\n".join(readme.splitlines()[:140]).strip()
    if not excerpt:
        excerpt = "No README content was found in this repository."

    body = (
        f"You are using a bridge skill generated from `{owner}/{repo}`.\n\n"
        "Repository context (excerpt):\n\n"
        f"{excerpt}\n\n"
        "Instructions:\n"
        "- Use this repository knowledge to answer user questions about setup and usage.\n"
        "- If code execution is required, inspect repository docs/scripts before acting.\n"
        "- If unclear, ask for the specific file or scenario."
    )
    content = (
        "---\n"
        f"name: {skill_name}\n"
        f"description: Bridge skill generated from {owner}/{repo}\n"
        "user-invocable: true\n"
        "---\n\n"
        f"{body}\n"
    )
    (destination / "SKILL.md").write_text(content, encoding="utf-8")
    return skill_name


def sync_skill_repositories(
    sources: list[str],
    *,
    windows_only: bool = False,
    dest_root: Path | None = None,
) -> dict:
    """Clone repositories and install discovered skills into the local skill folder."""
    install_root = dest_root or (Path.home() / ".agents" / "skills")
    install_root.mkdir(parents=True, exist_ok=True)

    results: list[RepoSyncResult] = []
    total_installed = 0

    for source in sources:
        try:
            owner, repo = _parse_source(source)
        except Exception as e:
            results.append(
                RepoSyncResult(
                    source=source,
                    repo="",
                    errors=[f"invalid source: {e}"],
                )
            )
            continue

        result = RepoSyncResult(source=source, repo=f"{owner}/{repo}")
        results.append(result)

        with tempfile.TemporaryDirectory(prefix="Mudabbir_skill_sync_") as tmpdir:
            clone_dir = Path(tmpdir) / "repo"
            clone_cmd = [
                "git",
                "clone",
                "--depth=1",
                f"https://github.com/{owner}/{repo}.git",
                str(clone_dir),
            ]
            proc = subprocess.run(clone_cmd, capture_output=True, text=True, timeout=240)
            if proc.returncode != 0:
                err = (proc.stderr or proc.stdout or "clone failed").strip()
                result.errors.append(err)
                continue

            skill_dirs = _iter_skill_dirs(clone_dir)
            if windows_only:
                windows_dirs = [d for d in skill_dirs if _is_windows_skill(d, clone_dir)]
                result.skipped.extend(
                    str(d.relative_to(clone_dir)).replace("\\", "/")
                    for d in skill_dirs
                    if d not in windows_dirs
                )
                skill_dirs = windows_dirs

            if not skill_dirs:
                generated = _generate_bridge_skill(
                    owner=owner,
                    repo=repo,
                    repo_root=clone_dir,
                    dest_root=install_root,
                )
                result.generated.append(generated)
                total_installed += 1
                continue

            for skill_dir in sorted(skill_dirs, key=lambda p: str(p)):
                try:
                    installed_name = _install_skill_dir(
                        owner=owner,
                        repo=repo,
                        repo_root=clone_dir,
                        skill_dir=skill_dir,
                        dest_root=install_root,
                    )
                    result.installed.append(installed_name)
                    total_installed += 1
                except Exception as e:
                    rel_skill_path = str(skill_dir.relative_to(clone_dir)).replace("\\", "/")
                    result.errors.append(
                        f"{rel_skill_path}: {e}"
                    )

    return {
        "status": "ok",
        "dest": str(install_root),
        "windows_only": bool(windows_only),
        "total_installed": total_installed,
        "results": [
            {
                "source": r.source,
                "repo": r.repo,
                "installed_count": r.installed_count,
                "installed": r.installed,
                "generated": r.generated,
                "skipped_count": len(r.skipped),
                "errors": r.errors,
            }
            for r in results
        ],
    }


def _main() -> int:
    parser = argparse.ArgumentParser(description="Sync external skills into Mudabbir skill folder")
    parser.add_argument(
        "--source",
        action="append",
        required=True,
        help="GitHub source (owner/repo or full URL). Repeat for multiple repos.",
    )
    parser.add_argument(
        "--windows-only",
        action="store_true",
        help="Install only Windows-related skills when SKILL.md exists.",
    )
    parser.add_argument(
        "--dest",
        default="",
        help="Destination skills directory (default: ~/.agents/skills).",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output full JSON report.",
    )
    args = parser.parse_args()

    dest = Path(args.dest).expanduser().resolve() if args.dest else None
    report = sync_skill_repositories(
        sources=args.source,
        windows_only=bool(args.windows_only),
        dest_root=dest,
    )

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(
            f"Synced skills to {report['dest']} | installed: {report['total_installed']} | windows_only={report['windows_only']}"
        )
        for item in report["results"]:
            repo = item.get("repo") or item.get("source")
            print(
                f"- {repo}: installed={item.get('installed_count', 0)}, generated={len(item.get('generated', []))}, errors={len(item.get('errors', []))}"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
