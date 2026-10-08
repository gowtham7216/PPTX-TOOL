"""
SlideGuard Production Configuration Management.
Separates configuration from source code and runtime behavior.
Ensures zero hardcoded secrets and robust fallback mechanisms.
"""

import os
from dataclasses import dataclass
from typing import Optional
from pathlib import Path
from app.paths import get_default_workspace_dir, get_default_log_dir


@dataclass
class AppConfig:
    """Runtime application configuration."""
    app_version: str = "1.0.0"
    gemini_api_key: Optional[str] = None
    gemini_model: str = "gemini-3.5"
    log_level: str = "INFO"
    max_log_size_bytes: int = 5 * 1024 * 1024  # 5 MB
    backup_log_count: int = 3
    workspace_dir: Path = get_default_workspace_dir()
    log_dir: Path = get_default_log_dir()
    output_suffix: str = "_SlideGuard_Remediated"
    report_suffix: str = "_SlideGuard_Report"
    enable_ai_authoring: bool = True

    @classmethod
    def load_from_env(cls) -> "AppConfig":
        """Loads configuration from environment variables with safe defaults."""
        api_key = (
            os.environ.get("GEMINI_API_KEY") or
            os.environ.get("GOOGLE_API_KEY") or
            os.environ.get("SLIDEGUARD_GEMINI_KEY")
        )
        # Strip potential surrounding quotes/whitespace
        if api_key:
            api_key = api_key.strip().strip("'\"")

        raw_level = os.environ.get("SLIDEGUARD_LOG_LEVEL", "INFO").upper()
        log_level = raw_level if raw_level in ("DEBUG", "INFO", "WARNING", "ERROR") else "INFO"

        return cls(
            gemini_api_key=api_key or None,
            gemini_model=os.environ.get("SLIDEGUARD_GEMINI_MODEL", "gemini-3.5"),
            log_level=log_level,
            enable_ai_authoring=os.environ.get("SLIDEGUARD_DISABLE_AI", "0") not in ("1", "true", "TRUE"),
        )

    def mask_key(self) -> str:
        """Returns masked API key for logging or UI display without exposing the secret."""
        if not self.gemini_api_key:
            return "Not Configured (Offline Fallback Active)"
        if len(self.gemini_api_key) <= 8:
            return "****"
        return f"{self.gemini_api_key[:4]}...{self.gemini_api_key[-4:]}"
