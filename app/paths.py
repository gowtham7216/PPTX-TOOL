"""
SlideGuard Production Path & Resource Resolution.
Ensures reliable asset and workspace location both in development and packaged PyInstaller EXE environments.
"""

import sys
import os
from pathlib import Path


def is_frozen() -> bool:
    """Returns True if running inside a PyInstaller frozen executable."""
    return getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")


def get_bundle_dir() -> Path:
    """
    Returns the root directory where application code and bundled resources reside.
    - If packaged: sys._MEIPASS
    - If development: repo root (parent of app/)
    """
    if is_frozen():
        return Path(sys._MEIPASS).resolve()
    # Development mode: directory containing the project (parent of app/)
    return Path(__file__).resolve().parent.parent


def get_user_data_dir() -> Path:
    """
    Returns a persistent, user-writable data directory on Windows.
    Defaults to %LOCALAPPDATA%\\SlideGuard, falling back to ~/.slideguard.
    """
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base = Path(local_app_data) / "SlideGuard"
    else:
        base = Path.home() / ".slideguard"
    base.mkdir(parents=True, exist_ok=True)
    return base


def get_default_workspace_dir() -> Path:
    """
    Returns default directory for job workspaces.
    Uses local repo workspace_runs in dev mode if writable, otherwise user data dir.
    """
    if not is_frozen():
        dev_ws = Path(__file__).resolve().parent.parent / "workspace_runs"
        try:
            dev_ws.mkdir(parents=True, exist_ok=True)
            return dev_ws
        except Exception:
            pass
    ws = get_user_data_dir() / "workspaces"
    ws.mkdir(parents=True, exist_ok=True)
    return ws


def get_default_log_dir() -> Path:
    """Returns default directory for rotating production logs."""
    logs = get_user_data_dir() / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    return logs


def resolve_resource_path(relative_path: str) -> Path:
    """Resolves a relative resource path against the application bundle root."""
    return get_bundle_dir() / relative_path
