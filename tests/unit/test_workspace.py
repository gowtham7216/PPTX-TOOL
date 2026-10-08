"""
Unit tests for SlideGuard WorkspaceManager.
Validates isolated sandbox creation, working copy duplication, and safe cleanup.
"""

from pathlib import Path
from workspace.manager import WorkspaceManager
from models.enums import JobStage


def test_workspace_isolation(sample_presentation_path, isolated_workspace):
    manager = WorkspaceManager(base_workspace_dir=isolated_workspace)
    context = manager.create_job(sample_presentation_path)

    # Check that directories exist
    assert context.workspace_dir.exists()
    assert context.original_file.exists()
    assert context.working_file.exists()
    assert context.reports_dir.exists()
    assert context.logs_dir.exists()

    # Check stage
    assert context.stage == JobStage.INITIALIZED

    # Working copy must not be the same file object as the source
    assert context.working_file != sample_presentation_path
    assert context.working_file.name.startswith("working_")

    # Safe cleanup
    manager.cleanup_job(context, preserve_on_debug=False)
    assert not context.workspace_dir.exists()
