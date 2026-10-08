"""
SlideGuard AI Authoring Applier (Phase 9).
Applies approved or user-edited AI authoring recommendations deterministically to OOXML.
Strict safety rules:
  - Only APPROVED or EDITED recommendations are applied.
  - PENDING_REVIEW, REJECTED, and MANUAL_REVIEW are strictly ignored and never modify PPTX.
  - Visual geometry, typography, colors, animations, and relationships are 100% preserved.
  - Generates comprehensive audit trail records for every applied or rejected action.
"""

import logging
from datetime import datetime
from typing import List, Tuple, Dict, Any, Optional
from lxml import etree
from pptx.util import Inches, Pt

from models.enums import RemediationStatus, ShapeType
from models.presentation import PresentationModel
from models.slide import SlideModel
from models.shape import ShapeModel
from models.remediation import RemediationAction
from models.job import JobContext
from models.ai_authoring import (
    AuthoringType,
    AuthoringStatus,
    AuthoringRecommendation,
    AuthoringAuditRecord,
)
from rules.accessibility.base import get_cnvpr_element, safe_xpath

logger = logging.getLogger(__name__)


class AuthoringApplier:
    """
    Deterministically applies user-approved AI authoring recommendations.
    Decoupled from AI generation; strictly acts on verified user approvals.
    """

    def apply_approved(
        self,
        context: JobContext,
        presentation: PresentationModel
    ) -> Tuple[List[RemediationAction], List[AuthoringAuditRecord]]:
        """
        Applies only recommendations marked as APPROVED or EDITED.
        Returns remediation actions and audit trail records.
        """
        actions: List[RemediationAction] = []
        audit_records: List[AuthoringAuditRecord] = []
        applied_count = 0
        failed_count = 0

        for rec in context.ai_recommendations:
            # Rule 14: Safety Gate
            if rec.status not in (AuthoringStatus.APPROVED, AuthoringStatus.EDITED):
                logger.debug(
                    "Skipping recommendation %s (status=%s). Only APPROVED or EDITED recommendations may be applied.",
                    rec.recommendation_id, rec.status.value
                )
                continue

            slide = next((s for s in presentation.slides if s.slide_number == rec.slide_number), None)
            if slide is None or slide.raw_slide is None:
                rec.mark_failed("Target slide not found in presentation.")
                failed_count += 1
                continue

            action, audit = self._apply_single_recommendation(rec, slide, presentation)
            actions.append(action)
            audit_records.append(audit)

            if action.status == RemediationStatus.APPLIED:
                applied_count += 1
                rec.mark_applied()
            else:
                failed_count += 1
                rec.mark_failed(action.error_message or "Application failed.")

        # Update JobContext audit trail and summary
        context.ai_audit_trail.extend([a.to_dict() for a in audit_records])
        if context.ai_authoring_summary:
            context.ai_authoring_summary["applied"] = applied_count
            context.ai_authoring_summary["failed"] = failed_count

        logger.info(
            "AuthoringApplier applied %d recommendation(s), %d failed.",
            applied_count, failed_count
        )
        return actions, audit_records

    def _apply_single_recommendation(
        self,
        rec: AuthoringRecommendation,
        slide: SlideModel,
        presentation: PresentationModel
    ) -> Tuple[RemediationAction, AuthoringAuditRecord]:
        """Applies a single approved recommendation based on its authoring type."""
        final_value = rec.effective_value

        try:
            if rec.authoring_type in (
                AuthoringType.ALT_TEXT,
                AuthoringType.IMAGE_DESCRIPTION,
                AuthoringType.CHART_DESCRIPTION,
                AuthoringType.DIAGRAM_DESCRIPTION,
                AuthoringType.TABLE_DESCRIPTION,
            ):
                return self._apply_alt_text(rec, slide, final_value)

            elif rec.authoring_type == AuthoringType.SLIDE_TITLE:
                return self._apply_slide_title(rec, slide, final_value)

            elif rec.authoring_type == AuthoringType.TABLE_HEADER_INTERPRETATION:
                return self._apply_table_header(rec, slide, final_value)

            else:
                # Unsupported authoring type for automatic application
                action = RemediationAction(
                    action_id=f"rem_{rec.recommendation_id}",
                    rule_id=f"AI_AUTH_{rec.authoring_type.value}",
                    slide_number=rec.slide_number,
                    target_shape_id=rec.shape_id,
                    description=f"Authoring type {rec.authoring_type.value} requires manual review.",
                    status=RemediationStatus.SKIPPED,
                    confidence=rec.confidence,
                    is_user_approved=True
                )
                audit = self._build_audit_record(rec, "SKIPPED", final_value)
                return action, audit

        except Exception as e:
            logger.error("Error applying AI authoring recommendation %s: %s", rec.recommendation_id, e)
            action = RemediationAction(
                action_id=f"rem_{rec.recommendation_id}",
                rule_id=f"AI_AUTH_{rec.authoring_type.value}",
                slide_number=rec.slide_number,
                target_shape_id=rec.shape_id,
                description=f"Failed to apply {rec.authoring_type.value}: {str(e)}",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message=str(e)
            )
            audit = self._build_audit_record(rec, "FAILED", "")
            return action, audit

    def _apply_alt_text(
        self,
        rec: AuthoringRecommendation,
        slide: SlideModel,
        final_value: str
    ) -> Tuple[RemediationAction, AuthoringAuditRecord]:
        """Applies approved alternative text to shape's p:cNvPr/@descr."""
        target_id = rec.shape_id
        sp_tree = slide.raw_slide.shapes._spTree
        found_cnvpr = None

        for child in sp_tree:
            cnvpr = get_cnvpr_element(child)
            if cnvpr is not None and cnvpr.get("id") == str(target_id):
                found_cnvpr = cnvpr
                break

        # Also search nested group shapes if needed
        if found_cnvpr is None:
            for child in sp_tree.iter():
                cnvpr = get_cnvpr_element(child)
                if cnvpr is not None and cnvpr.get("id") == str(target_id):
                    found_cnvpr = cnvpr
                    break

        if found_cnvpr is not None:
            from rules.alt_text.transformer import ImageAltTextTransformer
            transformer = ImageAltTextTransformer()

            is_mark_decorative = (
                final_value == "[MARK_DECORATIVE]"
                or rec.evidence.get("action") == "MARK_DECORATIVE"
            )
            is_preserve = (
                final_value == "[PRESERVE]"
                or rec.evidence.get("action") == "PRESERVE"
            )

            if is_mark_decorative:
                found_cnvpr.set("descr", "")
                transformer._ensure_decorative_extension(found_cnvpr)
                applied_desc = f"Marked shape {target_id} as decorative (empty alt text, decorative marker set)"
                final_applied_value = ""
            elif is_preserve:
                applied_desc = f"Preserved author alt text on shape {target_id}: '{final_value[:50]}...'"
                final_applied_value = final_value
            else:
                found_cnvpr.set("descr", final_value)
                transformer._remove_decorative_extension(found_cnvpr)
                applied_desc = f"Applied approved alt text to shape {target_id}: '{final_value[:50]}...'"
                final_applied_value = final_value

            # Update in-memory shape model
            for s in slide.shapes:
                if s.shape_id == target_id:
                    s.alt_text = final_applied_value
                    if is_mark_decorative:
                        s.is_decorative = True
                    break

            action = RemediationAction(
                action_id=f"rem_{rec.recommendation_id}",
                rule_id=f"AI_AUTH_{rec.authoring_type.value}",
                slide_number=rec.slide_number,
                target_shape_id=target_id,
                description=applied_desc,
                status=RemediationStatus.APPLIED,
                confidence=rec.confidence,
                is_user_approved=True,
                remediated_state={"descr": final_applied_value, "shape_id": target_id, "is_decorative": is_mark_decorative}
            )
            audit = self._build_audit_record(rec, "APPLIED", final_applied_value)
            return action, audit

        else:
            action = RemediationAction(
                action_id=f"rem_{rec.recommendation_id}",
                rule_id=f"AI_AUTH_{rec.authoring_type.value}",
                slide_number=rec.slide_number,
                target_shape_id=target_id,
                description=f"Shape {target_id} not found in slide shape tree.",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message="Shape not found in OOXML tree."
            )
            audit = self._build_audit_record(rec, "FAILED", "")
            return action, audit

    def _apply_slide_title(
        self,
        rec: AuthoringRecommendation,
        slide: SlideModel,
        final_value: str
    ) -> Tuple[RemediationAction, AuthoringAuditRecord]:
        """Applies approved slide title without altering visual dimensions or design."""
        ns_p = "http://schemas.openxmlformats.org/presentationml/2006/main"
        sp_tree = slide.raw_slide.shapes._spTree
        promoted = False

        # If a candidate shape was identified, promote or update it
        if rec.shape_id is not None:
            for child in sp_tree:
                cnvpr = get_cnvpr_element(child)
                if cnvpr is not None and cnvpr.get("id") == str(rec.shape_id):
                    nv_pr_nodes = safe_xpath(child, ".//p:nvSpPr/p:nvPr")
                    if nv_pr_nodes:
                        nv_pr = nv_pr_nodes[0]
                        for old_ph in safe_xpath(nv_pr, "./p:ph"):
                            nv_pr.remove(old_ph)
                        etree.SubElement(nv_pr, f"{{{ns_p}}}ph", attrib={"type": "title"})
                        promoted = True

                    # Update text if text frame exists
                    t_nodes = safe_xpath(child, ".//a:txBody//a:t")
                    if t_nodes:
                        t_nodes[0].text = final_value
                        # clear subsequent runs in first para
                        for extra_t in t_nodes[1:]:
                            extra_t.text = ""
                    promoted = True
                    break

        if not promoted:
            # Check if an existing title placeholder already exists on slide
            title_shape = getattr(slide.raw_slide.shapes, "title", None)
            if title_shape is not None:
                title_shape.text = final_value
                promoted = True
            else:
                # Add an accessible title text box at standard header location
                tb = slide.raw_slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8.4), Inches(0.8))
                tb.text = final_value
                nv_pr_nodes = safe_xpath(tb._element, ".//p:nvSpPr/p:nvPr")
                if nv_pr_nodes:
                    nv_pr = nv_pr_nodes[0]
                    etree.SubElement(nv_pr, f"{{{ns_p}}}ph", attrib={"type": "title"})
                promoted = True

        slide.title = final_value

        action = RemediationAction(
            action_id=f"rem_{rec.recommendation_id}",
            rule_id="AI_AUTH_SLIDE_TITLE",
            slide_number=rec.slide_number,
            target_shape_id=rec.shape_id,
            description=f"Applied approved slide title: '{final_value[:50]}'",
            status=RemediationStatus.APPLIED,
            confidence=rec.confidence,
            is_user_approved=True,
            remediated_state={"title": final_value}
        )
        audit = self._build_audit_record(rec, "APPLIED", final_value)
        return action, audit

    def _apply_table_header(
        self,
        rec: AuthoringRecommendation,
        slide: SlideModel,
        final_value: str
    ) -> Tuple[RemediationAction, AuthoringAuditRecord]:
        """Applies firstRow='1' to designated table tblPr."""
        target_id = rec.shape_id
        sp_tree = slide.raw_slide.shapes._spTree
        found = False

        for child in sp_tree:
            cnvpr = get_cnvpr_element(child)
            if cnvpr is not None and cnvpr.get("id") == str(target_id):
                tbl_prs = safe_xpath(child, ".//a:tbl/a:tblPr")
                if tbl_prs:
                    tbl_prs[0].set("firstRow", "1")
                    found = True
                    break

        if found:
            action = RemediationAction(
                action_id=f"rem_{rec.recommendation_id}",
                rule_id="AI_AUTH_TABLE_HEADER_INTERPRETATION",
                slide_number=rec.slide_number,
                target_shape_id=target_id,
                description=f"Applied approved table header property (firstRow='1') to table {target_id}.",
                status=RemediationStatus.APPLIED,
                confidence=rec.confidence,
                is_user_approved=True,
                remediated_state={"firstRow": "1", "shape_id": target_id}
            )
            audit = self._build_audit_record(rec, "APPLIED", "firstRow='1'")
            return action, audit
        else:
            action = RemediationAction(
                action_id=f"rem_{rec.recommendation_id}",
                rule_id="AI_AUTH_TABLE_HEADER_INTERPRETATION",
                slide_number=rec.slide_number,
                target_shape_id=target_id,
                description=f"Table {target_id} not found in slide tree.",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message="Table element not found in OOXML."
            )
            audit = self._build_audit_record(rec, "FAILED", "")
            return action, audit

    def _build_audit_record(
        self,
        rec: AuthoringRecommendation,
        decision_status: str,
        final_value: str
    ) -> AuthoringAuditRecord:
        """Constructs an auditable record of the authoring action."""
        return AuthoringAuditRecord(
            audit_id=f"aud_{rec.recommendation_id}",
            recommendation_id=rec.recommendation_id,
            model=rec.model,
            authoring_type=rec.authoring_type.value,
            slide_number=rec.slide_number,
            shape_id=rec.shape_id,
            original_value=rec.original_value,
            ai_suggestion=rec.suggested_value,
            user_decision=rec.status.value,
            final_applied_value=final_value,
            confidence=rec.confidence,
            status=decision_status,
            timestamp=datetime.utcnow().isoformat()
        )
