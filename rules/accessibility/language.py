"""
SlideGuard Document Language Rule (Phase 7).
Validates presence of document language settings for screen reader pronunciation.
Safely sets default language ONLY if explicitly specified by user configuration.
"""

from typing import List, Optional
from rules.base import BaseAccessibilityRule
from models.enums import RuleBehavior, IssueSeverity, AccessibilityStatus, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.remediation import RemediationAction


class LanguageRule(BaseAccessibilityRule):
    """
    Validates document language settings.
    Ensures screen readers can select the correct synthetic speech synthesizer.
    """

    def __init__(self, configured_language: Optional[str] = None):
        self.configured_language = configured_language

    @property
    def rule_id(self) -> str:
        return "SG-LNG-01"

    @property
    def name(self) -> str:
        return "Language of Document"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.SAFE_AUTO_REMEDIATE if self.configured_language else RuleBehavior.DETECT_ONLY

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 3.1.1 Language of Page (Level A)"

    @property
    def description(self) -> str:
        return "Verifies that presentation core metadata specifies a document language."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        # Only evaluate on first slide to avoid duplicate document-level warnings
        if slide.slide_number == 1:
            lang = presentation.metadata.language

            if not lang:
                can_auto = bool(self.configured_language)
                issues.append(
                    Issue(
                        issue_id="LNG_MISSING_DOC",
                        rule_id=self.rule_id,
                        rule_name=self.name,
                        slide_number=1,
                        severity=IssueSeverity.WARNING,
                        wcag_mapping=self.wcag_mapping,
                        description="Presentation document metadata does not declare a default language.",
                        recommendation="Set default presentation language (e.g. 'en-US') in document properties.",
                        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE if can_auto else RuleBehavior.DETECT_ONLY,
                        status=AccessibilityStatus.SAFE_TO_REMEDIATE if can_auto else AccessibilityStatus.MANUAL_REVIEW,
                        confidence="HIGH",
                        explanation="No language code declared in presentation metadata or default text styles.",
                        extra_data={"configured_language": self.configured_language}
                    )
                )

        return issues

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """Sets default language ONLY if explicitly configured by user."""
        lang_to_set = self.configured_language or (issue.extra_data or {}).get("configured_language")

        if not lang_to_set:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=1,
                description="No default language explicitly configured; preserving document without assumption.",
                status=RemediationStatus.SKIPPED,
                confidence=1.0,
                is_user_approved=True
            )

        try:
            presentation.metadata.language = lang_to_set
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=1,
                description=f"Set document language to configured value '{lang_to_set}' in presentation metadata.",
                status=RemediationStatus.APPLIED,
                confidence=1.0,
                is_user_approved=True
            )
        except Exception as e:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=1,
                description=f"Failed to set document language: {str(e)}",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message=str(e)
            )
