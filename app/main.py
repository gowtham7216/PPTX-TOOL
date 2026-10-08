"""
SlideGuard Application Main Entry Point.
Launches the desktop GUI application or runs CLI batch processing.
"""

import sys
import os
import logging
from pathlib import Path

# Safe standard stream handling in frozen GUI environments
if sys.platform == "win32":
    try:
        import ctypes
        # If launched from terminal, attach to existing console
        if ctypes.windll.kernel32.AttachConsole(-1):
            if sys.stdout is None or not getattr(sys.stdout, "isatty", lambda: False)():
                sys.stdout = open("CONOUT$", "w", encoding="utf-8", errors="replace")
            if sys.stderr is None or not getattr(sys.stderr, "isatty", lambda: False)():
                sys.stderr = open("CONOUT$", "w", encoding="utf-8", errors="replace")
    except Exception:
        pass

    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(errors="replace")
    except Exception:
        pass

if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8", errors="replace")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8", errors="replace")

# Ensure repository root is on PYTHONPATH
ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from app.version import VERSION, APP_TITLE
from app.logging_config import setup_production_logging
from app.controller.app_controller import AppController


def run_cli_process(input_path_str: str) -> int:
    """Headless batch processing mode for smoke tests and automation."""
    logger = logging.getLogger("slideguard.cli")
    path = Path(input_path_str).resolve()
    if not path.exists():
        sys.stderr.write(f"Error: File not found: {path}\n")
        return 1

    try:
        controller = AppController()
        logger.info("Initializing job for %s...", path.name)
        ctx = controller.initialize_job(path)

        logger.info("Analyzing %d slides...", ctx.total_slides)
        ctx = controller.execute_analysis(ctx)

        logger.info("Analysis complete: %d issues detected. Score: %.1f/100",
                    ctx.metrics.total_issues, ctx.metrics.accessibility_score)

        ctx.user_approved_remediation = True
        logger.info("Applying safe remediation pipeline...")
        ctx = controller.execute_remediation(ctx)

        logger.info("Remediation complete. Remediated file: %s", ctx.final_output_file)
        logger.info("Report file: %s", ctx.final_report_file)

        logger.info("Running QA verification...")
        ctx = controller.run_qa_verification(ctx)
        qa_status = ctx.qa_result.overall_status.value if ctx.qa_result else "UNKNOWN"
        logger.info("QA Status: %s", qa_status)

        out_msg = f"SUCCESS: {ctx.final_output_file}\n"
        try:
            sys.stdout.write(out_msg)
            sys.stdout.flush()
        except UnicodeEncodeError:
            enc = getattr(sys.stdout, "encoding", "utf-8") or "utf-8"
            sys.stdout.write(out_msg.encode(enc, errors="replace").decode(enc))
            sys.stdout.flush()
        return 0
    except Exception as e:
        logger.error("CLI processing failed: %s", e, exc_info=True)
        err_msg = f"ERROR: {e}\n"
        try:
            sys.stderr.write(err_msg)
            sys.stderr.flush()
        except UnicodeEncodeError:
            enc = getattr(sys.stderr, "encoding", "utf-8") or "utf-8"
            sys.stderr.write(err_msg.encode(enc, errors="replace").decode(enc))
            sys.stderr.flush()
        return 1


def main():
    if "--version" in sys.argv or "-v" in sys.argv:
        sys.stdout.write(f"SlideGuard {VERSION}\n")
        sys.stdout.flush()
        sys.exit(0)

    log_file = setup_production_logging()
    logger = logging.getLogger("slideguard.main")
    logger.info("Starting SlideGuard (Version %s)... Log file: %s", VERSION, log_file)

    if "--process" in sys.argv:
        idx = sys.argv.index("--process")
        if idx + 1 < len(sys.argv):
            target_file = sys.argv[idx + 1]
            code = run_cli_process(target_file)
            sys.exit(code)
        else:
            sys.stderr.write("Usage: SlideGuard --process <file.pptx>\n")
            sys.exit(1)

    import tkinter as tk
    from app.ui.main_window import SlideGuardMainWindow

    try:
        root = tk.Tk()
        root.title(f"{APP_TITLE} (v{VERSION})")
        controller = AppController()
        app = SlideGuardMainWindow(root, controller=controller)

        # If launched with --test-mode, destroy after 500ms to verify GUI launch without hanging
        if "--test-mode" in sys.argv:
            logger.info("Test mode detected: GUI initialized successfully. Exiting cleanly.")
            root.after(500, root.destroy)

        root.mainloop()
    except Exception as e:
        logger.critical("Unhandled fatal exception in application main: %s", e, exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
