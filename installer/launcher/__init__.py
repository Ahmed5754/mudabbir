# Mudabbir Desktop Launcher
# Thin wrapper that bootstraps Python/venv, installs mudabbir via pip,
# runs the server, and provides a system tray icon for non-technical users.
# Created: 2026-02-10

try:
    from importlib.metadata import version as _meta_version

    __version__ = _meta_version("mudabbir-launcher")
except Exception:
    # Fallback: read from MUDABBIR_VERSION env (set during build) or hardcoded
    import os

    __version__ = os.environ.get("MUDABBIR_VERSION", "0.1.0")
