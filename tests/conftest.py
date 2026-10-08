"""
Pytest configuration and shared fixtures for SlideGuard.
"""

import pytest
import shutil
from pathlib import Path
from tests.fixtures.sample_generator import create_comprehensive_test_pptx


@pytest.fixture(scope="session")
def sample_presentation_path(tmp_path_factory) -> Path:
    """Creates a temporary comprehensive PPTX file for the entire test session."""
    tmp_dir = tmp_path_factory.mktemp("slideguard_fixtures")
    pptx_path = tmp_dir / "sample_presentation.pptx"
    create_comprehensive_test_pptx(pptx_path)
    return pptx_path


@pytest.fixture
def isolated_workspace(tmp_path) -> Path:
    """Provides a fresh isolated directory for workspace tests."""
    ws_dir = tmp_path / "test_workspace"
    ws_dir.mkdir(parents=True, exist_ok=True)
    yield ws_dir
    if ws_dir.exists():
        shutil.rmtree(ws_dir, ignore_errors=True)
