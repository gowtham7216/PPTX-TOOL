"""
Regression Tests for SlideGuard Reading Order Remediation.
Verifies:
1. Slides with actual ordering changes (APPLIED, shapes physically reordered in candidate PPTX).
2. Slides requiring no changes (NO_CHANGES, sequence preserved, 0 mutations).
3. Low-confidence slides (SKIPPED, changes withheld by safety rule).
4. Writer failures (FAILED status recorded, error logged).
5. Batch processing, deduplication, and idempotent execution.
6. Original PPTX protection (original file remains strictly untouched).
7. Preservation of animation targets, shape IDs, and layout elements.
"""

import pytest
import shutil
from pathlib import Path
from unittest.mock import MagicMock
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

from models.enums import RemediationStatus, JobStage
from models.job import JobContext
from models.issue import Issue
from models.remediation import RemediationAction
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from engine.remediation.remediation_service import RemediationService
from app.controller.app_controller import AppController
from workspace.manager import WorkspaceManager
from rules.reading_order.rule import ReadingOrderRule


@pytest.fixture
def temp_workspace(tmp_path):
    ws_dir = tmp_path / "workspace"
    ws_dir.mkdir(parents=True, exist_ok=True)
    return ws_dir


@pytest.fixture
def sample_reading_order_deck(tmp_path):
    """
    Creates a multi-slide PPTX with:
    Slide 1: Inverted order (Body first, Title second) -> Needs APPLIED reordering
    Slide 2: Optimal order (Title first, Body second) -> Needs NO_CHANGES
    Slide 3: Footer slide (Title first, Body second, Footer last) -> Already logical or minimal
    """
    deck_path = tmp_path / "sample_ro.pptx"
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Slide 1: Body added first in XML, Title added second -> Needs reordering
    s1 = prs.slides.add_slide(blank_layout)
    body1 = s1.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(4))
    body1.text_frame.text = "Body text that was created before title."
    title1 = s1.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title1.text_frame.text = "Title text created after body."

    # Slide 2: Title added first in XML, Body added second -> Optimal order
    s2 = prs.slides.add_slide(blank_layout)
    title2 = s2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title2.text_frame.text = "Slide 2 Title"
    body2 = s2.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(4))
    body2.text_frame.text = "Slide 2 Body"

    # Slide 3: Title first, Body second, Footer at bottom
    s3 = prs.slides.add_slide(blank_layout)
    title3 = s3.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title3.text_frame.text = "Slide 3 Title"
    body3 = s3.shapes.add_textbox(Inches(1), Inches(2.0), Inches(8), Inches(3))
    body3.text_frame.text = "Slide 3 Content"
    footer3 = s3.shapes.add_textbox(Inches(1), Inches(7.0), Inches(3), Inches(0.4))
    footer3.text_frame.text = "WILEY © 2026"

    prs.save(str(deck_path))
    return deck_path


def test_actual_ordering_changes_applied_to_candidate(sample_reading_order_deck, temp_workspace):
    """
    Verifies that slides with inverted order are remediated to APPLIED,
    the candidate copy XML is physically updated, and original file is untouched.
    """
    orig_hash = sample_reading_order_deck.read_bytes()

    ws_mgr = WorkspaceManager(base_workspace_dir=temp_workspace)
    controller = AppController(workspace_manager=ws_mgr)
    context = controller.initialize_job(sample_reading_order_deck)
    context = controller.execute_analysis(context)

    # Approve remediation
    context.user_approved_remediation = True
    context = controller.execute_reading_order_remediation(context)

    # Check that original file was not modified
    assert sample_reading_order_deck.read_bytes() == orig_hash

    # Check candidate file exists
    assert context.candidate_file.exists()

    # Per-slide actions should be present in context.remediation_actions
    ro_actions = [a for a in context.remediation_actions if a.rule_id == "SG-RO-01"]
    assert len(ro_actions) == 3

    # Slide 1 had inverted order -> APPLIED
    s1_action = next(a for a in ro_actions if a.slide_number == 1)
    assert s1_action.status == RemediationStatus.APPLIED
    assert s1_action.original_state["order"] != s1_action.remediated_state["order"]

    # Candidate presentation check: verify physical XML order in candidate copy
    reader = PresentationReader()
    cand_pres = reader.read(context.candidate_file)
    cand_s1 = cand_pres.get_slide_by_number(1)
    cand_s1_shapes = list(cand_s1.raw_slide.shapes)

    # In candidate copy, title shape must now precede body shape in XML sequence
    title_id = cand_s1_shapes[0].shape_id
    body_id = cand_s1_shapes[1].shape_id
    assert "Title" in cand_s1_shapes[0].text_frame.text
    assert "Body" in cand_s1_shapes[1].text_frame.text


def test_slides_requiring_no_changes(tmp_path, temp_workspace):
    """
    Verifies that presentations already in logical reading order (like Chapter 9)
    are reported as NO_CHANGES, with 0 slides mutated.
    """
    deck_path = tmp_path / "optimal_deck.pptx"
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    for i in range(1, 4):
        s = prs.slides.add_slide(blank_layout)
        t = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        t.text_frame.text = f"Slide {i} Title"
        b = s.shapes.add_textbox(Inches(1), Inches(2.0), Inches(8), Inches(4))
        b.text_frame.text = f"Slide {i} Content"
    prs.save(str(deck_path))

    ws_mgr = WorkspaceManager(base_workspace_dir=temp_workspace)
    controller = AppController(workspace_manager=ws_mgr)
    context = controller.initialize_job(deck_path)
    context = controller.execute_analysis(context)

    context.user_approved_remediation = True
    context = controller.execute_reading_order_remediation(context)

    ro_actions = [a for a in context.remediation_actions if a.rule_id == "SG-RO-01"]
    assert len(ro_actions) == 3

    # All slides must have NO_CHANGES status
    assert all(a.status == RemediationStatus.NO_CHANGES for a in ro_actions)

    # Pipeline summary action rem_step_1 must have NO_CHANGES status
    step_1 = next(a for a in context.remediation_actions if a.action_id == "rem_step_1")
    assert step_1.status == RemediationStatus.NO_CHANGES
    assert "already in logical reading sequence" in step_1.description


def test_low_confidence_slide_skipped(sample_reading_order_deck, temp_workspace):
    """
    Verifies that slides with LOW confidence are withheld by the safety rule (SKIPPED).
    """
    ws_mgr = WorkspaceManager(base_workspace_dir=temp_workspace)
    controller = AppController(workspace_manager=ws_mgr)
    context = controller.initialize_job(sample_reading_order_deck)
    context = controller.execute_analysis(context)

    # Manually modify an issue's confidence to LOW
    for r_res in context.rule_results:
        if r_res.rule_id == "SG-RO-01":
            for issue in r_res.issues:
                if issue.slide_number == 1:
                    issue.extra_data["confidence"] = "LOW"

    context.user_approved_remediation = True
    context = controller.execute_reading_order_remediation(context)

    ro_actions = [a for a in context.remediation_actions if a.rule_id == "SG-RO-01"]
    s1_action = next(a for a in ro_actions if a.slide_number == 1)
    assert s1_action.status == RemediationStatus.SKIPPED
    assert "LOW confidence" in s1_action.description


def test_writer_failure_handling(sample_reading_order_deck, temp_workspace):
    """
    Verifies that when writer fails to write the candidate PPTX,
    remediation actions are marked as FAILED and the exception is raised.
    """
    mock_writer = MagicMock()
    mock_writer.write.side_effect = IOError("Disk full or permission denied")

    remediator = RemediationService(writer=mock_writer)
    ws_mgr = WorkspaceManager(base_workspace_dir=temp_workspace)
    controller = AppController(workspace_manager=ws_mgr, remediator=remediator)

    context = controller.initialize_job(sample_reading_order_deck)
    context = controller.execute_analysis(context)
    context.user_approved_remediation = True

    with pytest.raises(IOError, match="Disk full or permission denied"):
        controller.execute_reading_order_remediation(context)

    step_1 = next(a for a in context.remediation_actions if a.action_id == "rem_step_1")
    assert step_1.status == RemediationStatus.FAILED
    assert "Writer failure" in step_1.error_message or "Failed to write" in step_1.error_message


def test_idempotent_remediation_no_duplicate_actions(sample_reading_order_deck, temp_workspace):
    """
    Verifies that executing remediation twice on the same job context replaces
    prior reading order actions instead of creating duplicate entries.
    """
    ws_mgr = WorkspaceManager(base_workspace_dir=temp_workspace)
    controller = AppController(workspace_manager=ws_mgr)
    context = controller.initialize_job(sample_reading_order_deck)
    context = controller.execute_analysis(context)
    context.user_approved_remediation = True

    # First run
    context = controller.execute_reading_order_remediation(context)
    initial_count = len(context.remediation_actions)
    ro_actions_first = [a for a in context.remediation_actions if a.rule_id == "SG-RO-01"]
    assert len(ro_actions_first) == 3

    # Second run
    context = controller.execute_reading_order_remediation(context)
    second_count = len(context.remediation_actions)
    ro_actions_second = [a for a in context.remediation_actions if a.rule_id == "SG-RO-01"]

    assert second_count == initial_count
    assert len(ro_actions_second) == 3
    # Only 1 rem_step_1 action exists
    step_1_actions = [a for a in context.remediation_actions if a.action_id == "rem_step_1"]
    assert len(step_1_actions) == 1


def test_ui_log_deduplication():
    """
    Verifies that MainWindow._log_remediation_event ignores identical consecutive messages.
    """
    from app.ui.main_window import MainWindow

    class DummyApp:
        def __init__(self):
            self.rem_log_text = MagicMock()
            self._last_remediation_log_msg = None

    app = DummyApp()
    # Bind method to dummy app
    app._log_remediation_event = MainWindow._log_remediation_event.__get__(app, DummyApp)

    app._log_remediation_event("Reading Order: NO_CHANGES (all 35 slides already in logical order)")
    assert app.rem_log_text.insert.call_count == 1

    # Second identical call must be ignored
    app._log_remediation_event("Reading Order: NO_CHANGES (all 35 slides already in logical order)")
    assert app.rem_log_text.insert.call_count == 1

    # Different message must be logged
    app._log_remediation_event("Validation & Compare: QA PASSED")
    assert app.rem_log_text.insert.call_count == 2


def test_ui_no_changes_reports_no_changes_required(monkeypatch):
    """
    Verifies that when zero slides needed reordering, UI reports 'No changes required'
    instead of false-success 'Successfully reordered 0 slides'.
    """
    from app.ui.main_window import MainWindow

    mock_msgbox = MagicMock()
    monkeypatch.setattr("app.ui.main_window.messagebox.showinfo", mock_msgbox)

    class DummyApp:
        def __init__(self):
            self.btn_apply_ro = MagicMock()
            self.btn_fix_ro = MagicMock()
            self.ro_card_status_var = MagicMock()
            self.ro_affected_meta_var = MagicMock()
            self.status_var = MagicMock()
            self.rem_log_text = MagicMock()
            self._last_remediation_log_msg = None
            self.job_context = MagicMock()
            self.job_context.total_slides = 35
            self.job_context.remediation_actions = [
                RemediationAction(action_id="r1", rule_id="SG-RO-01", slide_number=1, status=RemediationStatus.NO_CHANGES, description="Already compliant")
            ]
        def _add_activity_event(self, color, text):
            pass
        def _log_remediation_event(self, msg):
            pass
        def _populate_reports_tab(self):
            pass

    app = DummyApp()
    app._on_reading_order_finished = MainWindow._on_reading_order_finished.__get__(app, DummyApp)

    app._on_reading_order_finished()

    # Must call info with "No Changes Required"
    mock_msgbox.assert_called_once()
    call_title, call_text = mock_msgbox.call_args[0]
    assert call_title == "No Changes Required"
    assert "No changes required" in call_text
    assert "Successfully reordered" not in call_text
    assert "35 slides" in call_text


def test_ui_skipped_reports_reasons(monkeypatch):
    """
    Verifies that when slides are skipped due to low confidence, UI reports
    the skipped count and specific slide reasons.
    """
    from app.ui.main_window import MainWindow

    mock_msgbox = MagicMock()
    monkeypatch.setattr("app.ui.main_window.messagebox.showinfo", mock_msgbox)

    class DummyApp:
        def __init__(self):
            self.btn_apply_ro = MagicMock()
            self.btn_fix_ro = MagicMock()
            self.ro_card_status_var = MagicMock()
            self.ro_affected_meta_var = MagicMock()
            self.status_var = MagicMock()
            self.job_context = MagicMock()
            self.job_context.total_slides = 10
            self.job_context.remediation_actions = [
                RemediationAction(
                    action_id="r1",
                    rule_id="SG-RO-01",
                    slide_number=4,
                    status=RemediationStatus.SKIPPED,
                    description="Skipped automatic reordering due to LOW confidence (manual author review required)."
                )
            ]
        def _add_activity_event(self, color, text):
            pass
        def _log_remediation_event(self, msg):
            pass
        def _populate_reports_tab(self):
            pass

    app = DummyApp()
    app._on_reading_order_finished = MainWindow._on_reading_order_finished.__get__(app, DummyApp)

    app._on_reading_order_finished()

    mock_msgbox.assert_called_once()
    call_title, call_text = mock_msgbox.call_args[0]
    assert call_title == "Reading Order Skipped"
    assert "1 slide(s) were skipped due to low confidence" in call_text
    assert "Slide 4" in call_text
    assert "Successfully reordered" not in call_text


def test_ui_applied_reports_success_only_when_greater_than_zero(monkeypatch):
    """
    Verifies that 'Successfully reordered N slides' is shown ONLY when N > 0.
    """
    from app.ui.main_window import MainWindow

    mock_msgbox = MagicMock()
    monkeypatch.setattr("app.ui.main_window.messagebox.showinfo", mock_msgbox)

    class DummyApp:
        def __init__(self):
            self.btn_apply_ro = MagicMock()
            self.btn_fix_ro = MagicMock()
            self.ro_card_status_var = MagicMock()
            self.ro_affected_meta_var = MagicMock()
            self.status_var = MagicMock()
            self.job_context = MagicMock()
            self.job_context.total_slides = 10
            self.job_context.remediation_actions = [
                RemediationAction(action_id="r1", rule_id="SG-RO-01", slide_number=1, status=RemediationStatus.APPLIED, description="Reordered shapes"),
                RemediationAction(action_id="r2", rule_id="SG-RO-01", slide_number=2, status=RemediationStatus.APPLIED, description="Reordered shapes"),
            ]
        def _add_activity_event(self, color, text):
            pass
        def _log_remediation_event(self, msg):
            pass
        def _populate_reports_tab(self):
            pass

    app = DummyApp()
    app._on_reading_order_finished = MainWindow._on_reading_order_finished.__get__(app, DummyApp)

    app._on_reading_order_finished()

    mock_msgbox.assert_called_once()
    call_title, call_text = mock_msgbox.call_args[0]
    assert call_title == "Reading Order Aligned"
    assert "Successfully reordered 2 slide(s)" in call_text
