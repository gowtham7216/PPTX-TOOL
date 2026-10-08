"""
SlideGuard Presentation Writer.
Saves presentations to disk and ensures output integrity.
"""

from pathlib import Path
import logging
from pptx import Presentation

from models.presentation import PresentationModel
from engine.common.interfaces import IPresentationWriter

logger = logging.getLogger(__name__)


class PresentationWriter(IPresentationWriter):
    """Safely saves presentation files."""

    def write(self, presentation: PresentationModel, output_path: Path) -> None:
        target = Path(output_path).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)

        if presentation.raw_presentation is None:
            raise ValueError("No raw presentation object attached to PresentationModel.")

        try:
            presentation.raw_presentation.save(str(target))
            logger.info("Successfully wrote presentation to %s", target)
        except Exception as e:
            logger.error("Failed to save presentation to %s: %s", target, e)
            raise IOError(f"Could not write presentation to {target}: {e}") from e

        # Quick validation that the file is non-empty and readable
        if not target.exists() or target.stat().st_size == 0:
            raise IOError(f"Written file {target} is empty or missing.")
