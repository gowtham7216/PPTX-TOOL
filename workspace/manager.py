"""
SlideGuard Workspace Manager.
Handles isolated job directories, temporary copies, job-specific logging, and safe cleanup.
"""

import shutil
import uuid
import logging
from pathlib import Path
from typing import Optional
from datetime import datetime

from models.job import JobContext
from models.enums import JobStage
from engine.common.interfaces import IWorkspaceManager


class WorkspaceManager(IWorkspaceManager):
    """
    Manages isolated job workspaces to ensure the user's original PPTX
    is NEVER modified directly and each execution run is fully sandboxed.
    """

    def __init__(self, base_workspace_dir: Optional[Path] = None):
        if base_workspace_dir is None:
            from app.paths import get_default_workspace_dir
            self.base_dir = get_default_workspace_dir()
        else:
            self.base_dir = Path(base_workspace_dir)
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def create_job(self, source_file: Path) -> JobContext:
        """
        Sets up an isolated workspace for a new processing job.
        Copies the source PPTX into the working folder.
        """
        import stat
        import os

        source_path = Path(source_file).resolve()
        if not source_path.exists():
            raise FileNotFoundError(f"Source presentation does not exist: {source_path}")

        unique_id = f"job_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
        job_dir = self.base_dir / unique_id

        # Setup standard directory structure
        original_dir = job_dir / "original"
        working_dir = job_dir / "working"
        candidate_dir = job_dir / "candidate"
        output_dir = job_dir / "output"
        reports_dir = job_dir / "reports"
        logs_dir = job_dir / "logs"

        for directory in [original_dir, working_dir, candidate_dir, output_dir, reports_dir, logs_dir]:
            directory.mkdir(parents=True, exist_ok=True)

        # Make copy of original in original_dir and working copy in working_dir
        archived_original = original_dir / source_path.name
        working_copy = working_dir / f"working_{source_path.name}"
        candidate_copy = candidate_dir / f"candidate_{source_path.name}"
        output_file = output_dir / f"remediated_{source_path.name}"

        shutil.copy2(source_path, archived_original)
        shutil.copy2(source_path, working_copy)

        # Ensure working copy is writable even if source file was read-only
        try:
            os.chmod(working_copy, stat.S_IWRITE | stat.S_IREAD)
        except Exception:
            pass

        # Setup dedicated job logger
        self._setup_job_logger(job_id=unique_id, logs_dir=logs_dir)

        context = JobContext(
            job_id=unique_id,
            workspace_dir=job_dir,
            source_file=source_path,
            original_file=archived_original,
            working_file=working_copy,
            candidate_file=candidate_copy,
            output_file=output_file,
            reports_dir=reports_dir,
            logs_dir=logs_dir,
            stage=JobStage.INITIALIZED,
        )

        logger = logging.getLogger(f"slideguard.{unique_id}")
        logger.info("Initialized job %s for %s", unique_id, source_path.name)
        return context


    def cleanup_job(self, job_context: JobContext, preserve_on_debug: bool = True) -> None:
        """
        Cleans up temporary job files.
        If preserve_on_debug is True, leaves the files for inspection.
        """
        if preserve_on_debug:
            return

        import gc
        # Explicitly close and detach logger handlers to release Windows file locks
        job_logger = logging.getLogger(f"slideguard.{job_context.job_id}")
        for handler in list(job_logger.handlers):
            try:
                handler.close()
                job_logger.removeHandler(handler)
            except Exception:
                pass

        # Close any handler in logging manager pointing to logs_dir
        if job_context.logs_dir and job_context.logs_dir.exists():
            log_file_str = str(job_context.logs_dir.resolve()).lower()
            for lg_name, lg in list(logging.root.manager.loggerDict.items()):
                if hasattr(lg, "handlers"):
                    for h in list(lg.handlers):
                        base_fn = getattr(h, "baseFilename", "")
                        if base_fn and log_file_str in str(Path(base_fn).resolve()).lower():
                            try:
                                h.close()
                                lg.removeHandler(h)
                            except Exception:
                                pass

        gc.collect()

        if job_context.workspace_dir.exists():
            try:
                shutil.rmtree(job_context.workspace_dir, ignore_errors=True)
            except Exception:
                pass

    def _setup_job_logger(self, job_id: str, logs_dir: Path) -> logging.Logger:
        """Configures an isolated logger for the job."""
        logger = logging.getLogger(f"slideguard.{job_id}")
        logger.setLevel(logging.DEBUG)

        log_file = logs_dir / "job.log"
        handler = logging.FileHandler(log_file, encoding="utf-8")
        formatter = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] [job=%(name)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        return logger
