"""
SlideGuard Remediation Service.
Coordinates the strictly ordered 6-step remediation pipeline.
Step 1 (Reading Order) is fully implemented; Steps 2-6 remain reserved for subsequent phases.
Requires explicit user approval before execution.
"""

from typing import List, Optional, Callable, Set, TYPE_CHECKING
import logging

from models.presentation import PresentationModel
from models.remediation import RemediationAction
from models.enums import RemediationStatus, JobStage, RuleBehavior, AccessibilityStatus
from models.job import JobContext
from pptx_engine.writer.presentation_writer import PresentationWriter

if TYPE_CHECKING:
    from rules.reading_order.rule import ReadingOrderRule
    from rules.ungrouping.ungroup_rule import SelectiveUngroupingRule
    from rules.placeholders.rule import PlaceholderRule

logger = logging.getLogger(__name__)


class RemediationService:
    """
    Executes presentation remediation strictly following the ordered pipeline:
      STEP 1: Reading Order (Active in Phase 3)
      STEP 2: Ungrouping (Phase 4)
      STEP 3: Placeholder Conversion (Phase 5)
      STEP 4: Lists (Phase 6)
      STEP 5: Accessibility Validation (Phase 7)
      STEP 6: Originality Verification (Phase 8)
    """

    def __init__(
        self,
        writer: Optional[PresentationWriter] = None,
        reading_order_rule: Optional["ReadingOrderRule"] = None,
        ungroup_rule: Optional["SelectiveUngroupingRule"] = None,
        placeholder_rule: Optional["PlaceholderRule"] = None,
        list_rule: Optional["ListRule"] = None
    ):
        self.writer = writer or PresentationWriter()
        if reading_order_rule is not None:
            self.ro_rule = reading_order_rule
        else:
            from rules.reading_order.rule import ReadingOrderRule
            self.ro_rule = ReadingOrderRule()

        if ungroup_rule is not None:
            self.ungroup_rule = ungroup_rule
        else:
            from rules.ungrouping.ungroup_rule import SelectiveUngroupingRule
            self.ungroup_rule = SelectiveUngroupingRule()

        if placeholder_rule is not None:
            self.ph_rule = placeholder_rule
        else:
            from rules.placeholders.rule import PlaceholderRule
            self.ph_rule = PlaceholderRule()

        if list_rule is not None:
            self.list_rule = list_rule
        else:
            from rules.lists.rule import ListRule
            self.list_rule = ListRule()

    def remediate(
        self,
        context: JobContext,
        presentation: PresentationModel,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> List[RemediationAction]:
        if not context.user_approved_remediation:
            raise PermissionError("Remediation cannot proceed without explicit user approval.")

        context.stage = JobStage.REMEDIATING
        logger.info("Starting ordered remediation pipeline for job %s", context.job_id)

        actions: List[RemediationAction] = []
        total_steps = 6

        # =====================================================================
        # STEP 1: Reading Order Remediation
        # =====================================================================
        if progress_callback:
            progress_callback(1, total_steps, "Executing STEP 1: Reading Order Alignment...")

        logger.info("Executing STEP 1: Reading Order Alignment")
        ro_actions = self._execute_reading_order_step(context, presentation)
        reordered_count = sum(1 for a in ro_actions if a.status == RemediationStatus.APPLIED)

        step_1_action = RemediationAction(
            action_id="rem_step_1",
            rule_id="PIPELINE_STEP_1_READING_ORDER",
            slide_number=0,
            description=f"STEP 1: Reading Order Alignment. Evaluated and processed {len(presentation.slides)} slides in-place.",
            status=RemediationStatus.APPLIED if reordered_count > 0 else RemediationStatus.SKIPPED,
            is_user_approved=True,
            confidence=0.95
        )
        actions.append(step_1_action)

        # =====================================================================
        # STEP 2: Selective Ungrouping Remediation (Phase 4)
        # =====================================================================
        if progress_callback:
            progress_callback(2, total_steps, "Executing STEP 2: Selective Ungrouping...")

        logger.info("Executing STEP 2: Selective Ungrouping")
        ungroup_actions = self._execute_ungrouping_step(context, presentation)
        ungrouped_count = sum(1 for a in ungroup_actions if a.status == RemediationStatus.APPLIED)

        step_2_action = RemediationAction(
            action_id="rem_step_2",
            rule_id="PIPELINE_STEP_2_UNGROUPING",
            slide_number=0,
            description=f"STEP 2: Selective Ungrouping. Evaluated and processed {len(presentation.slides)} slides in-place.",
            status=RemediationStatus.APPLIED if ungrouped_count > 0 else RemediationStatus.SKIPPED,
            is_user_approved=True,
            confidence=0.95
        )
        actions.append(step_2_action)

        # =====================================================================
        # STEP 3: Native Placeholder Conversion Remediation (Phase 5)
        # =====================================================================
        if progress_callback:
            progress_callback(3, total_steps, "Executing STEP 3: Native Placeholder Conversion...")

        logger.info("Executing STEP 3: Native Placeholder Conversion")
        ph_actions = self._execute_placeholder_step(context, presentation)
        converted_count = sum(1 for a in ph_actions if a.status == RemediationStatus.APPLIED)

        step_3_action = RemediationAction(
            action_id="rem_step_3",
            rule_id="PIPELINE_STEP_3_PLACEHOLDERS",
            slide_number=0,
            description=f"STEP 3: Native Placeholder Conversion. Evaluated and processed {len(presentation.slides)} slides in-place.",
            status=RemediationStatus.APPLIED if converted_count > 0 else RemediationStatus.SKIPPED,
            is_user_approved=True,
            confidence=0.95,
            remediated_state=context.task_4_placeholder_report or {}
        )
        actions.append(step_3_action)

        # =====================================================================
        # STEP 4: List Semantics & Numbering Preservation (Phase 6)
        # =====================================================================
        if progress_callback:
            progress_callback(4, total_steps, "Executing STEP 4: List Semantics & Numbering...")

        logger.info("Executing STEP 4: List Semantics & Numbering")
        list_actions = self._execute_list_step(context, presentation)
        list_remediated_count = sum(1 for a in list_actions if a.status == RemediationStatus.APPLIED)

        step_4_action = RemediationAction(
            action_id="rem_step_4",
            rule_id="PIPELINE_STEP_4_LISTS",
            slide_number=0,
            description=(
                f"STEP 4: List Semantics & Numbering. Remediated {list_remediated_count} list items in-place."
                if list_remediated_count > 0
                else "STEP 4: List Semantics & Numbering. No list remediation required; lists compliant or preserved."
            ),
            status=RemediationStatus.APPLIED if list_remediated_count > 0 else RemediationStatus.SKIPPED,
            is_user_approved=True,
            confidence=0.95
        )
        actions.append(step_4_action)

        # =====================================================================
        # STEP 5: Accessibility Validation & Safe Remediation (Phase 7)
        # =====================================================================
        if progress_callback:
            progress_callback(5, total_steps, "Executing STEP 5: Accessibility Validation & Safe Remediation...")

        logger.info("Executing STEP 5: Accessibility Validation & Safe Remediation")
        access_actions = self._execute_accessibility_step(context, presentation)
        access_remediated_count = sum(1 for a in access_actions if a.status == RemediationStatus.APPLIED)

        step_5_action = RemediationAction(
            action_id="rem_step_5",
            rule_id="PIPELINE_STEP_5_ACCESSIBILITY",
            slide_number=0,
            description=(
                f"STEP 5: Accessibility Validation & Safe Remediation. Safely remediated {access_remediated_count} issues in-place."
                if access_remediated_count > 0
                else "STEP 5: Accessibility Validation & Safe Remediation. No safe remediation required or issues flagged for manual review."
            ),
            status=RemediationStatus.APPLIED if access_remediated_count > 0 else RemediationStatus.SKIPPED,
            is_user_approved=True,
            confidence=0.95
        )
        actions.append(step_5_action)

        # =====================================================================
        # STEP 6: Final Validation & Issue Reconciliation (Phase 8)
        # =====================================================================
        if progress_callback:
            progress_callback(6, total_steps, "Executing STEP 6: Final Validation & Issue Reconciliation...")

        logger.info("Executing STEP 6: Final Validation & Issue Reconciliation")
        reconciliation_summary = self._execute_reconciliation_step(context, presentation, actions)

        step_6_action = RemediationAction(
            action_id="rem_step_6",
            rule_id="PIPELINE_STEP_6_RECONCILIATION",
            slide_number=0,
            description=(
                f"STEP 6: Final Validation & Issue Reconciliation. "
                f"Reconciled {reconciliation_summary.total_pre} pre-remediation issues: "
                f"{reconciliation_summary.resolved_count} resolved, "
                f"{reconciliation_summary.remediated_count} remediated, "
                f"{reconciliation_summary.unchanged_count} unchanged, "
                f"{reconciliation_summary.new_count} new."
            ),
            status=RemediationStatus.APPLIED,
            is_user_approved=True,
            confidence=1.0,
            remediated_state=reconciliation_summary.to_dict()
        )
        actions.append(step_6_action)

        # Save candidate output to isolated candidate directory
        if context.candidate_file and presentation.raw_presentation:
            try:
                self.writer.write(presentation, context.candidate_file)
                logger.info("Saved candidate working copy to %s", context.candidate_file)
            except Exception as e:
                logger.error("Failed to write candidate presentation: %s", e)

        context.remediation_actions = actions
        context.stage = JobStage.REMEDIATED
        return actions

    def _execute_reading_order_step(
        self,
        context: JobContext,
        presentation: PresentationModel
    ) -> List[RemediationAction]:
        """Runs Reading Order remediation across all slides with issues."""
        step_actions = []

        # Find reading order issues from context.rule_results
        ro_issues_by_slide = {}
        for r_res in context.rule_results:
            if r_res.rule_id == self.ro_rule.rule_id:
                for issue in r_res.issues:
                    ro_issues_by_slide[issue.slide_number] = issue

        reordered_count = 0
        for slide in presentation.slides:
            if slide.slide_number in ro_issues_by_slide:
                issue = ro_issues_by_slide[slide.slide_number]
                action = self.ro_rule.remediate_issue(slide, presentation, issue)
                step_actions.append(action)
                if action.status == RemediationStatus.APPLIED:
                    reordered_count += 1
            else:
                # Slide already had optimal reading order
                step_actions.append(
                    RemediationAction(
                        action_id=f"rem_ro_noop_{slide.slide_number}",
                        rule_id=self.ro_rule.rule_id,
                        slide_number=slide.slide_number,
                        description="Slide reading order is already logically compliant. Preserving original sequence.",
                        status=RemediationStatus.SKIPPED,
                        confidence=1.0,
                        is_user_approved=True
                    )
                )

        logger.info("Reading order remediation completed: %d slides reordered in-place.", reordered_count)
        return step_actions

    def _execute_ungrouping_step(
        self,
        context: JobContext,
        presentation: PresentationModel
    ) -> List[RemediationAction]:
        """Runs Selective Ungrouping remediation across all slides with issues."""
        step_actions = []

        # Find ungrouping issues from context.rule_results
        ungroup_issues_by_slide = {}
        for r_res in context.rule_results:
            if r_res.rule_id == self.ungroup_rule.rule_id:
                for issue in r_res.issues:
                    ungroup_issues_by_slide.setdefault(issue.slide_number, []).append(issue)

        ungrouped_count = 0
        affected_slides: Set[int] = set()
        for slide in presentation.slides:
            if slide.slide_number in ungroup_issues_by_slide:
                for issue in ungroup_issues_by_slide[slide.slide_number]:
                    action = self.ungroup_rule.remediate_issue(slide, presentation, issue)
                    step_actions.append(action)
                    if action.status == RemediationStatus.APPLIED:
                        ungrouped_count += 1
                        affected_slides.add(slide.slide_number)
            else:
                step_actions.append(
                    RemediationAction(
                        action_id=f"rem_ungroup_noop_{slide.slide_number}",
                        rule_id=self.ungroup_rule.rule_id,
                        slide_number=slide.slide_number,
                        description="Slide has no groups requiring ungrouping.",
                        status=RemediationStatus.SKIPPED,
                        confidence=1.0,
                        is_user_approved=True
                    )
                )

        if affected_slides:
            self._refresh_modified_slides(presentation, affected_slides)

        logger.info("Selective ungrouping completed: %d groups processed/ungrouped.", ungrouped_count)
        return step_actions

    def _execute_placeholder_step(
        self,
        context: JobContext,
        presentation: PresentationModel
    ) -> List[RemediationAction]:
        """Runs Placeholder conversion remediation across all slides with issues."""
        step_actions = []

        ph_issues_by_slide = {}
        for r_res in context.rule_results:
            if r_res.rule_id == self.ph_rule.rule_id:
                for issue in r_res.issues:
                    ph_issues_by_slide.setdefault(issue.slide_number, []).append(issue)

        converted_count = 0
        ph_affected_slides: Set[int] = set()
        for slide in presentation.slides:
            if slide.slide_number in ph_issues_by_slide:
                for issue in ph_issues_by_slide[slide.slide_number]:
                    action = self.ph_rule.remediate_issue(slide, presentation, issue)
                    step_actions.append(action)
                    if action.status == RemediationStatus.APPLIED:
                        converted_count += 1
                        ph_affected_slides.add(slide.slide_number)
            else:
                step_actions.append(
                    RemediationAction(
                        action_id=f"rem_ph_noop_{slide.slide_number}",
                        rule_id=self.ph_rule.rule_id,
                        slide_number=slide.slide_number,
                        description="Slide placeholders are already compliant or preserved.",
                        status=RemediationStatus.SKIPPED,
                        confidence=1.0,
                        is_user_approved=True
                    )
                )

        if ph_affected_slides:
            self._refresh_modified_slides(presentation, ph_affected_slides)

        # Generate Option C comprehensive placeholder audit report across all slides
        try:
            task_4_report = self.ph_rule.generate_presentation_audit_report(presentation)
            context.task_4_placeholder_report = task_4_report
        except Exception as e:
            logger.warning("Failed to generate task_4_placeholder_report: %s", e)

        logger.info("Placeholder conversion completed: %d shapes converted/remediated in-place.", converted_count)
        return step_actions

    def _execute_list_step(
        self,
        context: JobContext,
        presentation: PresentationModel
    ) -> List[RemediationAction]:
        """Runs List semantics remediation across all slides with issues."""
        step_actions = []

        list_issues_by_slide = {}
        for r_res in context.rule_results:
            if r_res.rule_id == self.list_rule.rule_id:
                for issue in r_res.issues:
                    list_issues_by_slide.setdefault(issue.slide_number, []).append(issue)

        remediated_count = 0
        for slide in presentation.slides:
            if slide.slide_number in list_issues_by_slide:
                for issue in list_issues_by_slide[slide.slide_number]:
                    action = self.list_rule.remediate_issue(slide, presentation, issue)
                    step_actions.append(action)
                    if action.status == RemediationStatus.APPLIED:
                        remediated_count += 1
            else:
                step_actions.append(
                    RemediationAction(
                        action_id=f"rem_list_noop_{slide.slide_number}",
                        rule_id=self.list_rule.rule_id,
                        slide_number=slide.slide_number,
                        description="Slide has no list issues requiring remediation.",
                        status=RemediationStatus.SKIPPED,
                        confidence=1.0,
                        is_user_approved=True
                    )
                )

        logger.info("List semantics step completed: %d items remediated in-place.", remediated_count)
        return step_actions

    def _execute_accessibility_step(
        self,
        context: JobContext,
        presentation: PresentationModel
    ) -> List[RemediationAction]:
        """Runs Phase 7 safe accessibility remediations across slides."""
        step_actions = []

        from rules.accessibility import (
            DecorativeRule,
            SlideTitleRule,
            TableAccessibilityRule,
            LanguageRule,
            DocumentMetadataRule
        )

        rules_by_id = {
            "SG-DEC-01": DecorativeRule(),
            "SG-TTL-01": SlideTitleRule(),
            "SG-TBL-01": TableAccessibilityRule(),
            "SG-LNG-01": LanguageRule(),
            "SG-META-01": DocumentMetadataRule(),
        }

        access_issues_by_slide = {}
        for r_res in context.rule_results:
            if r_res.rule_id in rules_by_id:
                for issue in r_res.issues:
                    if issue.behavior == RuleBehavior.SAFE_AUTO_REMEDIATE and issue.status == AccessibilityStatus.SAFE_TO_REMEDIATE:
                        access_issues_by_slide.setdefault(issue.slide_number, []).append(issue)

        remediated_count = 0
        access_affected_slides: Set[int] = set()
        for slide in presentation.slides:
            if slide.slide_number in access_issues_by_slide:
                for issue in access_issues_by_slide[slide.slide_number]:
                    rule = rules_by_id.get(issue.rule_id)
                    if rule:
                        try:
                            action = rule.remediate_issue(slide, presentation, issue)
                            step_actions.append(action)
                            if action.status == RemediationStatus.APPLIED:
                                remediated_count += 1
                                access_affected_slides.add(slide.slide_number)
                        except Exception as e:
                            logger.error("Accessibility remediation error on slide %d rule %s: %s",
                                         slide.slide_number, issue.rule_id, e)
                            step_actions.append(
                                RemediationAction(
                                    action_id=f"rem_access_fail_{issue.issue_id}",
                                    rule_id=issue.rule_id,
                                    slide_number=slide.slide_number,
                                    target_shape_id=issue.shape_id,
                                    description=f"Remediation failed: {str(e)}",
                                    status=RemediationStatus.FAILED,
                                    confidence=0.0,
                                    error_message=str(e)
                                )
                            )
            else:
                step_actions.append(
                    RemediationAction(
                        action_id=f"rem_access_noop_{slide.slide_number}",
                        rule_id="PIPELINE_STEP_5_ACCESSIBILITY",
                        slide_number=slide.slide_number,
                        description="Slide has no safe accessibility issues requiring automatic remediation.",
                        status=RemediationStatus.SKIPPED,
                        confidence=1.0,
                        is_user_approved=True
                    )
                )

        # Phase 9: Apply user-approved AI authoring recommendations
        if getattr(context, "ai_recommendations", None):
            from engine.authoring.authoring_applier import AuthoringApplier
            applier = AuthoringApplier()
            ai_actions, _ = applier.apply_approved(context, presentation)
            for act in ai_actions:
                step_actions.append(act)
                if act.status == RemediationStatus.APPLIED:
                    remediated_count += 1
                    access_affected_slides.add(act.slide_number)

        if access_affected_slides:
            self._refresh_modified_slides(presentation, access_affected_slides)

        logger.info("Accessibility safe remediation step completed: %d items remediated in-place.", remediated_count)
        return step_actions


    def _refresh_modified_slides(self, presentation: PresentationModel, slide_numbers: Set[int]):
        """Refreshes in-memory slide models from raw OOXML after structural mutation."""
        from pptx_engine.reader.presentation_reader import PresentationReader
        reader = PresentationReader()
        for slide in presentation.slides:
            if slide.slide_number in slide_numbers and slide.raw_slide is not None:
                try:
                    refreshed = reader._parse_slide(
                        slide.raw_slide,
                        slide.slide_number,
                        presentation.slide_width_pt,
                        presentation.slide_height_pt
                    )
                    slide.shapes = refreshed.shapes
                    slide.title = refreshed.title
                    slide.title_shape_id = refreshed.title_shape_id
                    slide.has_footer = refreshed.has_footer
                    slide.footer_text = refreshed.footer_text
                    slide.footer_shape_ids = refreshed.footer_shape_ids
                    logger.debug("Refreshed in-memory SlideModel for slide %d following structural change.", slide.slide_number)
                except Exception as e:
                    logger.warning("Could not refresh SlideModel for slide %d: %s", slide.slide_number, e)

    def _execute_reconciliation_step(
        self,
        context: JobContext,
        presentation: PresentationModel,
        applied_actions: List[RemediationAction]
    ):
        """Runs post-remediation analysis, issue correlation, and before/after reconciliation."""
        from engine.analyzer.presentation_analyzer import PresentationAnalyzer
        from engine.intelligence.correlator import IssueCorrelator
        from engine.intelligence.reconciliation import IssueReconciliationService

        correlator = IssueCorrelator()

        # 1. Ensure pre-remediation unified issues are populated
        if not context.unified_issues and context.rule_results:
            context.unified_issues = correlator.correlate(context.rule_results)
            context.pre_remediation_score = context.metrics.accessibility_score

        # 2. Run post-remediation analysis on the remediated presentation
        analyzer = PresentationAnalyzer()
        post_rule_results = analyzer.analyze(presentation)

        # 3. Correlate post-remediation issues
        post_unified_issues = correlator.correlate(post_rule_results)

        # 4. Reconcile pre-remediation vs post-remediation issues
        reconciler = IssueReconciliationService()
        reconciliation_summary = reconciler.reconcile(
            pre_issues=context.unified_issues,
            post_issues=post_unified_issues,
            remediation_actions=applied_actions
        )

        # 5. Compute post-remediation quality metrics
        post_metrics = analyzer.compute_metrics(post_rule_results, presentation.total_slides)
        context.post_remediation_score = post_metrics.quality_score_v2
        context.intelligence_summary = reconciliation_summary.to_dict()

        return reconciliation_summary


