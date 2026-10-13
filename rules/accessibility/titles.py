"""
SlideGuard Slide Title Accessibility Rule (Phase 7).
Validates that every slide possesses an identifiable, accessible title.
Safely promotes obvious title-like text boxes to native title placeholders
while strictly preserving visual geometry and character-for-character text.
"""

from typing import List, Optional
from lxml import etree

from rules.base import BaseAccessibilityRule
from rules.accessibility.base import get_cnvpr_element, safe_xpath
from models.enums import RuleBehavior, IssueSeverity, ShapeType, AccessibilityStatus, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.shape import ShapeModel
from models.remediation import RemediationAction


class SlideTitleRule(BaseAccessibilityRule):
    """
    Validates presence of slide titles.
    Every slide in an accessible presentation must have a clear title for outline navigation.
    """

    @property
    def rule_id(self) -> str:
        return "SG-TTL-01"

    @property
    def name(self) -> str:
        return "Slide Title Presence"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.SAFE_AUTO_REMEDIATE

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 2.4.2 Page Titled (Level A)"

    @property
    def description(self) -> str:
        return "Ensures that every slide has an accessible title for navigation."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        # 1. Check if slide has a valid native/placeholder title
        if slide.title and slide.title.strip():
            return []

        # 2. Check for obvious title-like standalone text box
        candidate = self._find_title_candidate(slide)

        if candidate is not None:
            candidate_text = ""
            for p in candidate.paragraphs:
                p_text = "".join(r.text for r in p.runs).strip()
                if p_text:
                    candidate_text = p_text
                    break

            issues.append(
                Issue(
                    issue_id=f"TTL_PROMOTABLE_{slide.slide_number}_{candidate.shape_id}",
                    rule_id=self.rule_id,
                    rule_name=self.name,
                    slide_number=slide.slide_number,
                    severity=IssueSeverity.WARNING,
                    shape_id=candidate.shape_id,
                    shape_name=candidate.name,
                    wcag_mapping=self.wcag_mapping,
                    description=f"Slide {slide.slide_number} has title-like text box '{candidate_text[:30]}' lacking native title placeholder semantics.",
                    recommendation="Promote text box to native title placeholder while preserving layout and geometry.",
                    behavior=self.behavior,
                    status=AccessibilityStatus.SAFE_TO_REMEDIATE,
                    confidence="HIGH",
                    explanation=f"Text box is near slide top (top={candidate.top_pt:.1f}pt) with concise heading text ({len(candidate_text)} chars).",
                    extra_data={
                        "candidate_shape_id": candidate.shape_id,
                        "candidate_text": candidate_text
                    }
                )
            )
        else:
            # 3. Missing title with no obvious candidate
            issues.append(
                Issue(
                    issue_id=f"TTL_MISSING_{slide.slide_number}",
                    rule_id=self.rule_id,
                    rule_name=self.name,
                    slide_number=slide.slide_number,
                    severity=IssueSeverity.ERROR,
                    wcag_mapping=self.wcag_mapping,
                    description=f"Slide {slide.slide_number} has no defined title.",
                    recommendation="Add a descriptive title to the slide title placeholder.",
                    behavior=RuleBehavior.DETECT_ONLY,
                    status=AccessibilityStatus.MANUAL_REVIEW,
                    confidence="HIGH",
                    explanation="No native title placeholder or prominent heading text box detected on slide.",
                )
            )

        return issues

    def _find_title_candidate(self, slide: SlideModel) -> Optional[ShapeModel]:
        """Finds prominent text box near top of slide suitable for title promotion."""
        candidates = []
        for s in slide.shapes:
            if s.shape_type in (ShapeType.TEXT_BOX, ShapeType.AUTO_SHAPE) and s.has_text:
                # Text should be near top of slide (e.g. top_pt <= 160)
                if s.top_pt <= 160.0:
                    total_chars = sum(len(r.text) for p in s.paragraphs for r in p.runs)
                    # Reasonable heading length
                    if 1 <= total_chars <= 140:
                        candidates.append((s.top_pt, s))

        if candidates:
            # Pick highest shape on slide
            candidates.sort(key=lambda x: x[0])
            return candidates[0][1]
        return None

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """Promotes title-like text box to native title placeholder in-place."""
        extra = issue.extra_data or {}
        candidate_id = issue.shape_id or extra.get("candidate_shape_id")

        if candidate_id is None:
            # Cannot invent title text
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description="Cannot automatically invent title text for untitled slide. Preserving original content.",
                status=RemediationStatus.SKIPPED,
                confidence=1.0,
                is_user_approved=True
            )

        if slide.raw_slide is None:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description="Raw slide element not accessible for title promotion.",
                status=RemediationStatus.FAILED,
                confidence=0.0
            )

        try:
            ns_p = "http://schemas.openxmlformats.org/presentationml/2006/main"
            sp_tree = slide.raw_slide.shapes._spTree
            promoted = False

            for child in sp_tree:
                cnvpr = get_cnvpr_element(child)
                if cnvpr is not None and cnvpr.get("id") == str(candidate_id):
                    from rules.placeholders.placeholder_engine_v2 import PlaceholderEngineV2
                    engine = PlaceholderEngineV2()
                    res = engine.ensure_placeholder(child, "title", sp_tree=sp_tree)
                    promoted = res.inserted or (res.reason == "existing-placeholder-preserved")
                    break

            if promoted:
                # Update in-memory slide title
                cand_text = extra.get("candidate_text", "")
                if cand_text:
                    slide.title = cand_text

                return RemediationAction(
                    action_id=f"rem_{issue.issue_id}",
                    rule_id=self.rule_id,
                    slide_number=slide.slide_number,
                    target_shape_id=candidate_id,
                    description=f"Promoted text box {candidate_id} to native title placeholder in-place.",
                    status=RemediationStatus.APPLIED,
                    confidence=0.96,
                    is_user_approved=True
                )
            else:
                return RemediationAction(
                    action_id=f"rem_{issue.issue_id}",
                    rule_id=self.rule_id,
                    slide_number=slide.slide_number,
                    target_shape_id=candidate_id,
                    description=f"Candidate shape {candidate_id} not found in slide shape tree.",
                    status=RemediationStatus.FAILED,
                    confidence=0.0
                )
        except Exception as e:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                target_shape_id=candidate_id,
                description=f"Title promotion failed: {str(e)}",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message=str(e)
            )
