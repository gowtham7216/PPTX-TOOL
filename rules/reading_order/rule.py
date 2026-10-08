"""
SlideGuard Reading Order Accessibility Rule.
Inspects presentation reading order and semantic layout.
Ensures footer content is excluded from normal sequence and logical visual flow is maintained.
Implements safe, in-place XML node reordering without deleting or recreating elements.
"""

from typing import List, Optional, Dict, Any
import logging

from rules.base import BaseAccessibilityRule
from models.enums import RuleBehavior, IssueSeverity, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.remediation import RemediationAction
from rules.reading_order.semantic_classifier import SemanticRoleClassifier, SemanticItem, SemanticRole
from rules.reading_order.layout_analyzer import LayoutRegionAnalyzer
from rules.reading_order.order_resolver import ReadingOrderResolver, ReadingOrderResult
from rules.reading_order.xml_reorderer import XMLReadingOrderReorderer

logger = logging.getLogger(__name__)


class ReadingOrderRule(BaseAccessibilityRule):
    """
    Validates logical reading order of shapes on a slide.
    Ensures titles are read first and bottom footer items (e.g. 'WILEY')
    do not interfere with main reading content.
    """

    def __init__(
        self,
        classifier: Optional[SemanticRoleClassifier] = None,
        layout_analyzer: Optional[LayoutRegionAnalyzer] = None,
        resolver: Optional[ReadingOrderResolver] = None,
        reorderer: Optional[XMLReadingOrderReorderer] = None,
        assistant: Optional[Any] = None
    ):
        self.classifier = classifier or SemanticRoleClassifier()
        self.layout_analyzer = layout_analyzer or LayoutRegionAnalyzer()
        self.resolver = resolver or ReadingOrderResolver(layout_analyzer=self.layout_analyzer, assistant=assistant)
        self.reorderer = reorderer or XMLReadingOrderReorderer()


    @property
    def rule_id(self) -> str:
        return "SG-RO-01"

    @property
    def name(self) -> str:
        return "Reading Order Logical Sequence"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.SAFE_AUTO_REMEDIATE

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A)"

    @property
    def description(self) -> str:
        return "Verifies that reading order follows logical content hierarchy and footers are separated."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        if len(slide.shapes) <= 1:
            return issues

        # 1. Classify semantic items on slide
        raw_slide = slide.raw_slide
        if raw_slide is None:
            return issues

        classified_items: List[SemanticItem] = []
        raw_shapes = list(raw_slide.shapes)
        shape_map = {s.shape_id: s for s in raw_shapes}

        for z_idx, shape in enumerate(raw_shapes):
            s_elem = shape._element
            item = self.classifier.classify_shape(
                shape_elem=s_elem,
                shape_obj=shape,
                slide_h_pt=presentation.slide_height_pt,
                slide_w_pt=presentation.slide_width_pt,
                z_order=z_idx,
                surrounding_items=classified_items
            )
            classified_items.append(item)

        # 2. Resolve reading order
        result: ReadingOrderResult = self.resolver.resolve_reading_order(
            slide_number=slide.slide_number,
            items=classified_items,
            slide_w_pt=presentation.slide_width_pt,
            slide_h_pt=presentation.slide_height_pt
        )

        # 3. If reordering is required, emit an Issue
        if result.changes_required:
            issues.append(
                Issue(
                    issue_id=f"RO_ORDER_{slide.slide_number}",
                    rule_id=self.rule_id,
                    rule_name=self.name,
                    slide_number=slide.slide_number,
                    severity=IssueSeverity.WARNING,
                    wcag_mapping=self.wcag_mapping,
                    description=(
                        f"Slide {slide.slide_number} reading sequence does not match logical hierarchy. "
                        f"Current: {result.original_order} -> Proposed: {result.proposed_order}."
                    ),
                    recommendation="Reorder top-level shapes into proposed logical reading sequence.",
                    behavior=self.behavior,
                    extra_data=result.to_dict()
                )
            )

        return issues

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """
        Executes physical in-place reordering of elements within <p:spTree>.
        Preserves all shape IDs and animation target references.
        """
        action_id = f"rem_{issue.issue_id}"
        extra = issue.extra_data or {}
        proposed_order = extra.get("proposed_order", [])
        confidence = extra.get("confidence", "MEDIUM")

        if not proposed_order:
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description="No proposed reading order provided in issue.",
                status=RemediationStatus.SKIPPED,
                confidence=0.0,
                error_message="Missing proposed order data."
            )

        # Low confidence guard: do not auto-reorder ambiguous layouts
        if confidence == "LOW":
            logger.warning("Skipping automatic reorder for slide %d due to LOW confidence.", slide.slide_number)
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description="Skipped automatic reordering due to LOW confidence (manual author review required).",
                status=RemediationStatus.SKIPPED,
                confidence=0.3,
                error_message="Low confidence layout; automatic reordering withheld."
            )

        if slide.raw_slide is None or not hasattr(slide.raw_slide, "_element"):
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description="No raw slide element attached to SlideModel.",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message="Missing slide XML element."
            )

        # Perform in-place XML reordering
        slide_elem = slide.raw_slide._element
        success, message, resulting_order = self.reorderer.reorder_sp_tree(
            slide_elem=slide_elem,
            proposed_order=proposed_order
        )

        if not success:
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"In-place XML reordering failed: {message}",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message=message
            )

        # Animation safety verification: verify all valid animation targets are preserved
        orig_ids_set = set(extra.get("original_order", []))
        anim_targets_preserved = self._verify_animation_targets(slide_elem, orig_ids_set)
        if not anim_targets_preserved:
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description="Animation target verification failed post-reordering.",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message="Animation target shape ID was invalidated."
            )

        return RemediationAction(
            action_id=action_id,
            rule_id=self.rule_id,
            slide_number=slide.slide_number,
            description=f"Reordered shapes in-place to sequence {resulting_order}.",
            status=RemediationStatus.APPLIED,
            original_state={"order": extra.get("original_order", [])},
            remediated_state={"order": resulting_order},
            confidence=0.95 if confidence == "HIGH" else 0.80,
            is_user_approved=True
        )

    def _verify_animation_targets(self, slide_elem, orig_shape_ids: Optional[set] = None) -> bool:
        """Verifies that all shape IDs targeted by <p:timing> before reordering still exist on the slide."""
        timings = slide_elem.xpath(".//p:timing")
        if not timings:
            return True

        sp_tgts = timings[0].xpath(".//p:spTgt")
        if not sp_tgts:
            return True

        # Collect current shape IDs on slide
        current_ids = set()
        cnv_prs = slide_elem.xpath(".//p:cNvPr")
        for cp in cnv_prs:
            s_id = cp.get("id")
            if s_id and s_id.isdigit():
                current_ids.add(int(s_id))

        valid_orig_ids = orig_shape_ids if orig_shape_ids is not None else current_ids
        for tgt in sp_tgts:
            spid = tgt.get("spid")
            if spid and spid.isdigit():
                val = int(spid)
                if val in valid_orig_ids and val not in current_ids:
                    logger.error("Animation target spid=%s lost during reordering!", spid)
                    return False

        return True

