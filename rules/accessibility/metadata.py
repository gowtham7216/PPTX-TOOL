"""
SlideGuard Document Metadata Rule (Phase 7).
Validates presence of essential document properties (Title, Author).
Safely sets title/author ONLY if explicitly supplied via user configuration.
"""

from typing import List, Optional
from rules.base import BaseAccessibilityRule
from models.enums import RuleBehavior, IssueSeverity, AccessibilityStatus, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.remediation import RemediationAction


class DocumentMetadataRule(BaseAccessibilityRule):
    """
    Validates presentation metadata.
    Screen readers utilize document title metadata to present file identity.
    """

    def __init__(
        self,
        configured_title: Optional[str] = None,
        configured_author: Optional[str] = None
    ):
        self.configured_title = configured_title
        self.configured_author = configured_author

    @property
    def rule_id(self) -> str:
        return "SG-META-01"

    @property
    def name(self) -> str:
        return "Document Metadata Information"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.SAFE_AUTO_REMEDIATE if (self.configured_title or self.configured_author) else RuleBehavior.DETECT_ONLY

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 2.4.2 Page Titled (Level A)"

    @property
    def description(self) -> str:
        return "Ensures presentation has a title and author in core metadata properties."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        if slide.slide_number == 1:
            title = presentation.metadata.title
            if not title or not title.strip():
                can_auto = bool(self.configured_title)
                issues.append(
                    Issue(
                        issue_id="META_NO_TITLE",
                        rule_id=self.rule_id,
                        rule_name=self.name,
                        slide_number=1,
                        severity=IssueSeverity.WARNING,
                        wcag_mapping=self.wcag_mapping,
                        description="Presentation file properties lack a document title.",
                        recommendation="Set a descriptive document title in Presentation Properties.",
                        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE if can_auto else RuleBehavior.DETECT_ONLY,
                        status=AccessibilityStatus.SAFE_TO_REMEDIATE if can_auto else AccessibilityStatus.MANUAL_REVIEW,
                        confidence="HIGH",
                        explanation="Document core property dc:title is missing or whitespace only.",
                        extra_data={"configured_title": self.configured_title}
                    )
                )

        return issues

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """Sets document title/author ONLY if explicitly supplied via user configuration."""
        title_to_set = self.configured_title or (issue.extra_data or {}).get("configured_title")

        if not title_to_set:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=1,
                description="No document title provided in configuration; preserving original properties without assumption.",
                status=RemediationStatus.SKIPPED,
                confidence=1.0,
                is_user_approved=True
            )

        try:
            presentation.metadata.title = title_to_set
            # Also update raw pptx core_properties if available
            if presentation.raw_presentation is not None:
                presentation.raw_presentation.core_properties.title = title_to_set
                if self.configured_author:
                    presentation.raw_presentation.core_properties.author = self.configured_author

            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=1,
                description=f"Set document title to configured value '{title_to_set}' in core properties.",
                status=RemediationStatus.APPLIED,
                confidence=1.0,
                is_user_approved=True
            )
        except Exception as e:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=1,
                description=f"Failed to update document metadata: {str(e)}",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message=str(e)
            )
