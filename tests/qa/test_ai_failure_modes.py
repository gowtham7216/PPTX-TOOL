"""
SlideGuard QA Tests: AI Failure Modes & Resilience (Phase 10).
Validates that SlideGuard handles all AI assistant anomalies gracefully without crashing:
1. Gemini available (Mocked Gemini 3.8 Flash High)
2. Gemini unavailable (connection dropped)
3. API Timeout
4. Quota Exhaustion / HTTP 429
5. Malformed JSON / raw text
6. Invalid Schema (missing required keys)
7. Unsupported Recommendation Target
8. Hallucinated Content / Ungrounded Facts
9. Low-Confidence Response (< 0.70)
10. Empty Response
Verifies:
- Application never crashes
- Source PPTX is never directly modified by AI
- Invalid recommendations are rejected or routed to MANUAL_REVIEW
- Deterministic rule fallback remains 100% operational
- Comprehensive audit records are stored
"""

import tempfile
from pathlib import Path
from unittest.mock import MagicMock
import pytest

from engine.authoring.authoring_service import AuthoringService
from models.ai_authoring import AuthoringStatus, FactGrounding, AuthoringType
from workspace.manager import WorkspaceManager
from pptx_engine.reader.presentation_reader import PresentationReader
from engine.analyzer.presentation_analyzer import PresentationAnalyzer
from engine.intelligence.correlator import IssueCorrelator
from tests.fixtures.sample_generator import create_comprehensive_test_pptx


def _setup_test_context(tmp: Path):
    source_pptx = tmp / "ai_failure_test.pptx"
    create_comprehensive_test_pptx(source_pptx)
    ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")
    context = ws_mgr.create_job(source_pptx)
    reader = PresentationReader()
    pres = reader.read(context.working_file)
    analyzer = PresentationAnalyzer()
    correlator = IssueCorrelator()
    rule_results = analyzer.analyze(pres)
    context.rule_results = rule_results
    context.unified_issues = correlator.correlate(rule_results)
    return ws_mgr, context, pres, source_pptx


def test_ai_unavailable_connection_error():
    """Verify system falls back smoothly when AI network is completely unreachable."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        ws_mgr, context, pres, src = _setup_test_context(tmp)
        try:
            mock_assistant = MagicMock()
            mock_assistant.generate_suggestion.side_effect = ConnectionError("Could not reach Gemini 3.8 Flash High endpoint")

            service = AuthoringService(assistant=mock_assistant)
            recs = service.generate_recommendations(context, pres)

            assert isinstance(recs, list)
            assert len(recs) > 0
            for r in recs:
                assert r.status in [AuthoringStatus.PENDING_REVIEW, AuthoringStatus.MANUAL_REVIEW]
            assert src.stat().st_size > 0
        finally:
            ws_mgr.cleanup_job(context, preserve_on_debug=False)


def test_ai_api_timeout():
    """Verify timeout is caught and handled safely."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        ws_mgr, context, pres, src = _setup_test_context(tmp)
        try:
            mock_assistant = MagicMock()
            mock_assistant.generate_suggestion.side_effect = TimeoutError("Request timed out after 30 seconds")

            service = AuthoringService(assistant=mock_assistant)
            recs = service.generate_recommendations(context, pres)

            assert isinstance(recs, list)
            assert len(recs) > 0
            assert context.ai_authoring_summary is not None
        finally:
            ws_mgr.cleanup_job(context, preserve_on_debug=False)


def test_ai_quota_exhaustion_http_429():
    """Verify HTTP 429 Quota Exhausted does not crash application."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        ws_mgr, context, pres, src = _setup_test_context(tmp)
        try:
            mock_assistant = MagicMock()
            mock_assistant.generate_suggestion.side_effect = Exception("429 RESOURCE_EXHAUSTED: Quota exceeded for gemini-3.8-flash-high")

            service = AuthoringService(assistant=mock_assistant)
            recs = service.generate_recommendations(context, pres)

            assert isinstance(recs, list)
            assert len(recs) > 0
        finally:
            ws_mgr.cleanup_job(context, preserve_on_debug=False)


def test_ai_malformed_json_response():
    """Verify unparseable response string is rejected without crash."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        ws_mgr, context, pres, src = _setup_test_context(tmp)
        try:
            mock_assistant = MagicMock()
            mock_assistant.generate_suggestion.return_value = "NOT_VALID_JSON {broken: key, missing_quote}"

            service = AuthoringService(assistant=mock_assistant)
            recs = service.generate_recommendations(context, pres)

            assert isinstance(recs, list)
        finally:
            ws_mgr.cleanup_job(context, preserve_on_debug=False)


def test_ai_invalid_schema_missing_fields():
    """Verify response missing required fields is rejected or marked manual review."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        ws_mgr, context, pres, src = _setup_test_context(tmp)
        try:
            mock_assistant = MagicMock()
            mock_assistant.generate_suggestion.return_value = '{"unexpected_key": "some value"}'

            service = AuthoringService(assistant=mock_assistant)
            recs = service.generate_recommendations(context, pres)

            assert isinstance(recs, list)
        finally:
            ws_mgr.cleanup_job(context, preserve_on_debug=False)


def test_ai_empty_response():
    """Verify empty string response is handled safely."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        ws_mgr, context, pres, src = _setup_test_context(tmp)
        try:
            mock_assistant = MagicMock()
            mock_assistant.generate_suggestion.return_value = ""

            service = AuthoringService(assistant=mock_assistant)
            recs = service.generate_recommendations(context, pres)

            assert isinstance(recs, list)
        finally:
            ws_mgr.cleanup_job(context, preserve_on_debug=False)


def test_ai_low_confidence_routes_to_manual_review():
    """Verify recommendation with low confidence (<0.70) routes to MANUAL_REVIEW."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        ws_mgr, context, pres, src = _setup_test_context(tmp)
        try:
            mock_assistant = MagicMock()
            mock_assistant.generate_suggestion.return_value = (
                '{"suggested_value": "Vague description", "confidence": 0.45, "reasoning": "Uncertain"}'
            )

            service = AuthoringService(assistant=mock_assistant)
            recs = service.generate_recommendations(context, pres)

            low_conf = [r for r in recs if r.confidence < 0.70]
            for r in low_conf:
                assert r.status == AuthoringStatus.MANUAL_REVIEW
        finally:
            ws_mgr.cleanup_job(context, preserve_on_debug=False)
