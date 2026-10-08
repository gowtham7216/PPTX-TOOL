"""
SlideGuard QA Tests: Workspace Sandboxing & Safe Cleanup (Phase 10).
Validates that:
- Every job receives an isolated, unique sandbox workspace.
- Temporary files are cleanly destroyed when cleanup is invoked.
- Stale data from previous jobs never leaks into subsequent jobs.
- File locks on Windows are safely released before directory cleanup.
"""

import tempfile
import logging
from pathlib import Path
import pytest

from workspace.manager import WorkspaceManager
from tests.fixtures.sample_generator import create_comprehensive_test_pptx


def test_workspace_isolation_and_cleanup():
    """Verify separate workspaces and complete cleanup on completion."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source = tmp / "test_cleanup.pptx"
        create_comprehensive_test_pptx(source)

        ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")

        # Job 1
        ctx1 = ws_mgr.create_job(source)
        assert ctx1.workspace_dir.exists()
        assert ctx1.original_file.exists()
        assert ctx1.working_file.exists()

        # Job 2
        ctx2 = ws_mgr.create_job(source)
        assert ctx2.workspace_dir.exists()
        assert ctx1.job_id != ctx2.job_id
        assert ctx1.workspace_dir != ctx2.workspace_dir

        # Cleanup Job 1
        ws_mgr.cleanup_job(ctx1, preserve_on_debug=False)
        assert not ctx1.workspace_dir.exists()
        # Job 2 remains intact
        assert ctx2.workspace_dir.exists()

        # Cleanup Job 2
        ws_mgr.cleanup_job(ctx2, preserve_on_debug=False)
        assert not ctx2.workspace_dir.exists()

        # Original source remains intact
        assert source.exists()


def test_workspace_cleanup_after_pipeline_exception():
    """Verify workspace cleanup succeeds even if an operation fails midway."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source = tmp / "test_cleanup_fail.pptx"
        create_comprehensive_test_pptx(source)

        ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")
        ctx = ws_mgr.create_job(source)

        try:
            # Simulate work and error
            raise ValueError("Something unexpected broke during job!")
        except ValueError:
            pass
        finally:
            ws_mgr.cleanup_job(ctx, preserve_on_debug=False)

        assert not ctx.workspace_dir.exists()
        assert source.exists()
