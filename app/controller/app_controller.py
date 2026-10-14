"""
SlideGuard Application Controller.
Coordinates workflow between the desktop UI, workspace manager, analysis engine,
remediation pipeline, verification, and reporting.
"""

from pathlib import Path
from typing import Optional, Callable, Any, List
import logging

from models.job import JobContext
from models.enums import JobStage
from workspace.manager import WorkspaceManager
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from engine.analyzer.presentation_analyzer import PresentationAnalyzer
from engine.remediation.remediation_service import RemediationService
from engine.verification.verification_service import VerificationService
from engine.comparison.comparison_service import ComparisonService
from reports.generator import ReportGenerator

logger = logging.getLogger(__name__)


class AppController:
    """
    Decouples UI entirely from PowerPoint parsing and processing.
    Directs all lifecycle steps across isolated workspaces.
    """

    def __init__(
        self,
        workspace_manager: Optional[WorkspaceManager] = None,
        reader: Optional[PresentationReader] = None,
        writer: Optional[PresentationWriter] = None,
        analyzer: Optional[PresentationAnalyzer] = None,
        remediator: Optional[RemediationService] = None,
        verifier: Optional[VerificationService] = None,
        comparer: Optional[ComparisonService] = None,
        reporter: Optional[ReportGenerator] = None,
    ):
        self.workspace_mgr = workspace_manager or WorkspaceManager()
        self.reader = reader or PresentationReader()
        self.writer = writer or PresentationWriter()
        self.analyzer = analyzer or PresentationAnalyzer()
        self.remediator = remediator or RemediationService(writer=self.writer)
        self.verifier = verifier or VerificationService()
        self.comparer = comparer or ComparisonService()
        self.reporter = reporter or ReportGenerator()
        self.current_presentation = None

    def detect_presentation_fonts(self, presentation: Optional[Any]) -> List[str]:
        """Scans all slides, shapes, and runs to detect original fonts used in presentation."""
        detected = set()
        if not presentation:
            return []
        for slide in presentation.slides:
            for shape in slide.shapes:
                for para in shape.paragraphs:
                    for run in para.runs:
                        if run.font_name:
                            detected.add(run.font_name.strip())
        return sorted(list(detected))

    def initialize_job(self, file_path: Path) -> JobContext:
        """Sets up sandboxed workspace and creates an isolated working copy."""
        logger.info("Initializing job for %s", file_path)
        context = self.workspace_mgr.create_job(file_path)

        # Parse initial presentation structure
        self.current_presentation = self.reader.read(context.working_file)
        context.total_slides = self.current_presentation.total_slides
        context.slides_analyzed = 0
        context.detected_fonts = self.detect_presentation_fonts(self.current_presentation)
        return context

    def execute_analysis(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> JobContext:
        """Executes full rule-based accessibility analysis."""
        if self.current_presentation is None:
            self.current_presentation = self.reader.read(context.working_file)
        context.detected_fonts = self.detect_presentation_fonts(self.current_presentation)

        context.stage = JobStage.ANALYZING
        rule_results = self.analyzer.analyze(
            self.current_presentation,
            progress_callback=progress_callback
        )

        context.rule_results = rule_results
        context.slides_analyzed = self.current_presentation.total_slides
        context.metrics = self.analyzer.compute_metrics(
            rule_results,
            total_slides=self.current_presentation.total_slides
        )

        # Phase 8: Context-aware issue correlation and pre-remediation quality score
        from engine.intelligence.correlator import IssueCorrelator
        correlator = IssueCorrelator()
        context.unified_issues = correlator.correlate(rule_results)
        context.pre_remediation_score = context.metrics.quality_score_v2
        context.pre_accessibility_score = context.metrics.accessibility_score

        # Task 4 Comprehensive Placeholder Baseline Audit
        try:
            from rules.placeholders.rule import PlaceholderRule
            ph_rule = PlaceholderRule()
            context.task_4_placeholder_report = ph_rule.generate_presentation_audit_report(self.current_presentation)
        except Exception as err:
            logger.warning("Could not generate baseline placeholder audit: %s", err)

        # Phase 9: AI-Assisted Accessibility Authoring recommendations
        from engine.authoring.authoring_service import AuthoringService
        authoring_service = AuthoringService()
        authoring_service.generate_recommendations(context, self.current_presentation)

        context.stage = JobStage.ANALYZED

        # Generate audit reports
        if context.reports_dir:
            self.reporter.generate_json_report(context, context.reports_dir / "analysis_report.json")
            self.reporter.generate_markdown_report(context, context.reports_dir / "analysis_report.md")

        return context

    def generate_ai_authoring(
        self,
        context: JobContext,
        assistant: Optional[Any] = None
    ) -> JobContext:
        """Explicitly generates AI accessibility authoring recommendations for human review."""
        if self.current_presentation is None:
            self.current_presentation = self.reader.read(context.working_file)

        from engine.authoring.authoring_service import AuthoringService
        service = AuthoringService(assistant=assistant)
        service.generate_recommendations(context, self.current_presentation)
        return context

    def apply_approved_authoring(self, context: JobContext) -> JobContext:
        """Applies only user-approved or edited AI authoring recommendations to candidate copy."""
        if self.current_presentation is None:
            self.current_presentation = self.reader.read(context.working_file)

        from engine.authoring.authoring_applier import AuthoringApplier
        applier = AuthoringApplier()
        actions, audit = applier.apply_approved(context, self.current_presentation)
        context.remediation_actions.extend(actions)
        return context


    def execute_remediation(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        skip_empty_text: bool = False
    ) -> JobContext:
        """
        Executes remediation on candidate working copy.
        Strictly requires context.user_approved_remediation == True.
        """
        if not context.user_approved_remediation:
            raise PermissionError("User approval is required before running remediation.")

        if self.current_presentation is None:
            self.current_presentation = self.reader.read(context.working_file)

        # Execute remediation pipeline
        actions = self.remediator.remediate(
            context,
            self.current_presentation,
            progress_callback=progress_callback,
            skip_empty_text=skip_empty_text
        )
        context.remediation_actions = actions
        return self._finalize_candidate_and_export(context)

    def execute_reading_order_remediation(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> JobContext:
        """Executes independent reading order remediation and saves candidate copy."""
        if not context.user_approved_remediation:
            raise PermissionError("User approval is required before running remediation.")

        if self.current_presentation is None:
            if context.candidate_file and context.candidate_file.exists():
                self.current_presentation = self.reader.read(context.candidate_file)
            else:
                self.current_presentation = self.reader.read(context.working_file)

        if progress_callback:
            progress_callback(1, 2, "Aligning reading order across slides...")

        self.remediator.remediate_reading_order(context, self.current_presentation)
        return self._finalize_candidate_and_export(context)

    def execute_outline_order_remediation(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> JobContext:
        """Executes independent Outline Order alignment so Outline View matches Reading Order, and saves candidate copy."""
        if not context.user_approved_remediation:
            raise PermissionError("User approval is required before running remediation.")

        if self.current_presentation is None:
            if context.candidate_file and context.candidate_file.exists():
                self.current_presentation = self.reader.read(context.candidate_file)
            else:
                self.current_presentation = self.reader.read(context.working_file)

        if progress_callback:
            progress_callback(1, 2, "Aligning Outline View order to match Reading Order...")

        self.remediator.remediate_outline_order(context, self.current_presentation)
        return self._finalize_candidate_and_export(context)

    def execute_placeholder_remediation(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        skip_empty_text: bool = False
    ) -> JobContext:
        """Executes independent native placeholder conversion and saves candidate copy."""
        if not context.user_approved_remediation:
            raise PermissionError("User approval is required before running remediation.")

        if self.current_presentation is None:
            if context.candidate_file and context.candidate_file.exists():
                self.current_presentation = self.reader.read(context.candidate_file)
            else:
                self.current_presentation = self.reader.read(context.working_file)

        if progress_callback:
            progress_callback(1, 2, "Converting shapes to native placeholders...")

        self.remediator.remediate_placeholders(context, self.current_presentation, skip_empty_text=skip_empty_text)
        return self._finalize_candidate_and_export(context)

    def execute_alt_text_remediation(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> JobContext:
        """Executes independent alt text and decorative remediation and saves candidate copy."""
        if not context.user_approved_remediation:
            raise PermissionError("User approval is required before running remediation.")

        if self.current_presentation is None:
            if context.candidate_file and context.candidate_file.exists():
                self.current_presentation = self.reader.read(context.candidate_file)
            else:
                self.current_presentation = self.reader.read(context.working_file)

        if progress_callback:
            progress_callback(1, 2, "Applying alt text and marking decorative visuals...")

        self.remediator.remediate_alt_text_and_decorative(context, self.current_presentation)
        return self._finalize_candidate_and_export(context)

    def execute_ungrouping_remediation(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> JobContext:
        """Executes independent selective ungrouping remediation and saves candidate copy."""
        if not context.user_approved_remediation:
            raise PermissionError("User approval is required before running remediation.")

        if self.current_presentation is None:
            if context.candidate_file and context.candidate_file.exists():
                self.current_presentation = self.reader.read(context.candidate_file)
            else:
                self.current_presentation = self.reader.read(context.working_file)

        if progress_callback:
            progress_callback(1, 2, "Selectively ungrouping shapes for assistive access...")

        self.remediator.remediate_selective_ungrouping(context, self.current_presentation)
        return self._finalize_candidate_and_export(context)

    def execute_list_remediation(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> JobContext:
        """Executes independent list semantics remediation and saves candidate copy."""
        if not context.user_approved_remediation:
            raise PermissionError("User approval is required before running remediation.")

        if self.current_presentation is None:
            if context.candidate_file and context.candidate_file.exists():
                self.current_presentation = self.reader.read(context.candidate_file)
            else:
                self.current_presentation = self.reader.read(context.working_file)

        if progress_callback:
            progress_callback(1, 2, "Normalizing list semantics and numbering hierarchy...")

        self.remediator.remediate_list_semantics(context, self.current_presentation)
        return self._finalize_candidate_and_export(context)

    def generate_reports(self, context: JobContext) -> JobContext:
        """Explicitly generates or updates audit reports in reports folder and export destination."""
        if context.reports_dir:
            stage_val = getattr(context.stage, "value", str(context.stage))
            final_stages = {
                JobStage.REMEDIATED,
                JobStage.VERIFYING,
                JobStage.VERIFIED,
                JobStage.COMPLETED,
                "REMEDIATED",
                "VERIFYING",
                "VERIFIED",
                "COMPLETED",
            }
            report_name = "final_report" if (context.stage in final_stages or stage_val in final_stages) else "analysis_report"
            self.reporter.generate_json_report(context, context.reports_dir / f"{report_name}.json")
            self.reporter.generate_markdown_report(context, context.reports_dir / f"{report_name}.md")

            try:
                source_p = getattr(context, "source_file", None) or context.original_file
                source_resolved = source_p.resolve() if source_p else None

                # Reuse existing final_report_file if set, belongs to this job, and is not the source file
                final_rep = context.final_report_file
                is_safe_rep = (
                    final_rep is not None
                    and str(final_rep).strip() != ""
                    and (source_resolved is None or final_rep.resolve() != source_resolved)
                )

                if not is_safe_rep:
                    dest_dir = source_p.parent if (source_p and source_p.parent.exists()) else context.workspace_dir
                    stem = source_p.stem if source_p else "presentation"
                    final_rep = self.get_collision_safe_path(
                        base_dir=dest_dir,
                        stem=stem,
                        suffix="_SlideGuard_Report",
                        ext=".md",
                        forbidden_paths=[source_p] if source_p else None
                    )

                md_source = context.reports_dir / f"{report_name}.md"
                if md_source.exists():
                    import shutil
                    if source_resolved and final_rep.resolve() == source_resolved:
                        logger.error("Attempted to overwrite source presentation with report: %s. Aborting report export.", source_p)
                    else:
                        shutil.copy2(md_source, final_rep)
                        context.final_report_file = final_rep
                        logger.info("Exported audit report: %s", final_rep)
            except Exception as err:
                logger.warning("Report export warning: %s", err)
        return context

    def run_validation_and_compare(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> JobContext:
        """Executes verification, structural comparison, Phase 10 QA evaluation, and production export."""
        if self.current_presentation is None and context.working_file:
            self.current_presentation = self.reader.read(context.working_file)

        # If candidate file doesn't exist yet, copy working file as base candidate
        if context.candidate_file and not context.candidate_file.exists() and context.working_file.exists():
            import shutil
            shutil.copy2(context.working_file, context.candidate_file)

        if progress_callback:
            progress_callback(1, 3, "Verifying candidate integrity...")

        return self._finalize_candidate_and_export(context)

    def _export_candidate_pptx(self, context: JobContext):
        """Safely exports candidate PPTX to output directory with collision avoidance."""
        if context.candidate_file and context.candidate_file.exists():
            try:
                import shutil
                source_p = getattr(context, "source_file", None) or context.original_file
                source_resolved = source_p.resolve() if source_p else None

                # Reuse existing final_output_file if set, belongs to this job, and is not the source file
                final_pptx = context.final_output_file
                is_safe_pptx = (
                    final_pptx is not None
                    and str(final_pptx).strip() != ""
                    and (source_resolved is None or final_pptx.resolve() != source_resolved)
                )

                if not is_safe_pptx:
                    dest_dir = source_p.parent if (source_p and source_p.parent.exists()) else context.workspace_dir
                    stem = source_p.stem if source_p else "presentation"
                    final_pptx = self.get_collision_safe_path(
                        base_dir=dest_dir,
                        stem=stem,
                        suffix="_SlideGuard_Remediated",
                        ext=".pptx",
                        forbidden_paths=[source_p] if source_p else None
                    )

                if source_resolved and final_pptx.resolve() == source_resolved:
                    logger.error("Attempted to overwrite source presentation: %s. Aborting export.", source_p)
                else:
                    shutil.copy2(context.candidate_file, final_pptx)
                    context.final_output_file = final_pptx
                    logger.info("Exported production remediated presentation: %s", final_pptx)
            except Exception as e:
                logger.warning("Intermediate candidate export warning: %s", e)

    def _finalize_candidate_and_export(self, context: JobContext) -> JobContext:
        """Common candidate verification, structural comparison, QA evaluation, and report export."""
        # Verify candidate output
        if context.candidate_file and context.candidate_file.exists():
            context.stage = JobStage.VERIFYING
            verification = self.verifier.verify_file(
                context.candidate_file,
                expected_slide_count=context.total_slides
            )
            context.verification_result = verification

            # Parse candidate presentation for structural comparison
            candidate_presentation = self.reader.read(context.candidate_file)

            # Collect approved/remediated shape IDs for intentional change tracking
            approved_shape_ids = set()
            for r_res in context.rule_results:
                for issue in r_res.issues:
                    if issue.shape_id:
                        approved_shape_ids.add(issue.shape_id)

            comparison = self.comparer.compare(
                original=self.current_presentation,
                remediated=candidate_presentation,
                intentional_shape_ids=list(approved_shape_ids)
            )
            context.comparison_result = comparison

            # Re-analyze candidate presentation to verify actual resulting accessibility state
            post_rule_results = self.analyzer.analyze(candidate_presentation)
            post_metrics = self.analyzer.compute_metrics(post_rule_results, context.total_slides)
            context.post_rule_results = post_rule_results
            context.post_metrics = post_metrics
            context.post_remediation_score = post_metrics.quality_score_v2
            context.post_accessibility_score = post_metrics.accessibility_score

            # Run issue reconciliation comparing pre-issues against candidate findings
            from engine.intelligence.correlator import IssueCorrelator
            from engine.intelligence.reconciliation import IssueReconciliationService
            correlator = IssueCorrelator()
            if not context.unified_issues and context.rule_results:
                context.unified_issues = correlator.correlate(context.rule_results)

            post_unified_issues = correlator.correlate(post_rule_results)
            raw_pre_issues = [issue for res in context.rule_results for issue in res.issues]
            raw_post_issues = [issue for res in post_rule_results for issue in res.issues]
            reconciler = IssueReconciliationService()
            reconciliation_actions = getattr(context, "detailed_remediation_actions", None) or context.remediation_actions
            reconciliation_summary = reconciler.reconcile(
                pre_issues=context.unified_issues,
                post_issues=post_unified_issues,
                remediation_actions=reconciliation_actions,
                raw_pre_issues=raw_pre_issues,
                raw_post_issues=raw_post_issues,
            )

            # Update verification & remediation statuses in context.rule_results
            from collections import Counter
            post_issue_counts = Counter(
                (issue.slide_number, issue.shape_id, res.rule_id)
                for res in post_rule_results for issue in res.issues
            )
            from models.enums import AccessibilityStatus
            for res in context.rule_results:
                for issue in res.issues:
                    key = (issue.slide_number, issue.shape_id, res.rule_id)
                    if post_issue_counts[key] > 0:
                        post_issue_counts[key] -= 1
                        issue.verification_status = "VERIFIED"
                    else:
                        issue.status = AccessibilityStatus.REMEDIATED
                        issue.verification_status = "VERIFIED"

            raw_pre_count = len(raw_pre_issues)
            raw_post_count = len(raw_post_issues)
            raw_remediated_count = sum(1 for res in context.rule_results for issue in res.issues if issue.status == AccessibilityStatus.REMEDIATED)
            raw_unresolved_count = len([i for res in context.rule_results for i in res.issues if i.status != AccessibilityStatus.REMEDIATED])

            reconciliation_summary.raw_issues_pre = raw_pre_count
            reconciliation_summary.raw_issues_post = raw_post_count
            reconciliation_summary.raw_issues_remediated = raw_remediated_count
            context.intelligence_summary = reconciliation_summary.to_dict()

            # Update authoritative metrics on context.metrics
            context.metrics.raw_remediated_issues = raw_remediated_count
            context.metrics.remediated_issues = reconciliation_summary.remediated_count
            context.metrics.unresolved_issues = reconciliation_summary.total_post
            context.metrics.manual_review_issues = reconciliation_summary.manual_review_count
            context.metrics.correlated_issues_count = reconciliation_summary.total_pre

            # Phase 10: Run comprehensive QA checks on remediated candidate
            try:
                from engine.qa.qa_service import QAService
                qa_service = QAService(
                    reader=self.reader,
                    writer=self.writer,
                    analyzer=self.analyzer,
                    comparer=self.comparer
                )
                context.qa_result = qa_service.evaluate_job_context(context)
            except Exception as e:
                logger.warning("Phase 10 QA evaluation warning: %s", e)

        context.stage = JobStage.COMPLETED

        # Generate final reports in workspace
        if context.reports_dir:
            self.reporter.generate_json_report(context, context.reports_dir / "final_report.json")
            self.reporter.generate_markdown_report(context, context.reports_dir / "final_report.md")

        # Phase 11 Production Output Management:
        try:
            source_p = getattr(context, "source_file", None) or context.original_file
            source_resolved = source_p.resolve() if source_p else None

            # Reuse existing final_output_file if set, belongs to this job, and is not the source file
            final_pptx = context.final_output_file
            is_safe_pptx = (
                final_pptx is not None
                and str(final_pptx).strip() != ""
                and (source_resolved is None or final_pptx.resolve() != source_resolved)
            )

            if not is_safe_pptx:
                dest_dir = source_p.parent if (source_p and source_p.parent.exists()) else context.workspace_dir
                stem = source_p.stem if source_p else "presentation"
                final_pptx = self.get_collision_safe_path(
                    base_dir=dest_dir,
                    stem=stem,
                    suffix="_SlideGuard_Remediated",
                    ext=".pptx",
                    forbidden_paths=[source_p] if source_p else None
                )

            if context.candidate_file and context.candidate_file.exists():
                import shutil
                if source_resolved and final_pptx.resolve() == source_resolved:
                    logger.error("Attempted to overwrite source presentation: %s. Aborting export.", source_p)
                else:
                    shutil.copy2(context.candidate_file, final_pptx)
                    context.final_output_file = final_pptx
                    logger.info("Exported production remediated presentation: %s", final_pptx)

            # Reuse existing final_report_file if set, belongs to this job, and is not the source file
            final_rep = context.final_report_file
            is_safe_rep = (
                final_rep is not None
                and str(final_rep).strip() != ""
                and (source_resolved is None or final_rep.resolve() != source_resolved)
            )

            if not is_safe_rep:
                dest_dir = source_p.parent if (source_p and source_p.parent.exists()) else context.workspace_dir
                stem = source_p.stem if source_p else "presentation"
                final_rep = self.get_collision_safe_path(
                    base_dir=dest_dir,
                    stem=stem,
                    suffix="_SlideGuard_Report",
                    ext=".md",
                    forbidden_paths=[source_p] if source_p else None
                )

            if context.reports_dir and (context.reports_dir / "final_report.md").exists():
                import shutil
                if source_resolved and final_rep.resolve() == source_resolved:
                    logger.error("Attempted to overwrite source presentation with report: %s. Aborting report export.", source_p)
                else:
                    shutil.copy2(context.reports_dir / "final_report.md", final_rep)
                    context.final_report_file = final_rep
                    logger.info("Exported production remediation report: %s", final_rep)
        except Exception as err:
            logger.warning("Production export non-fatal warning: %s", err)

        return context

    @staticmethod
    def get_collision_safe_path(
        base_dir: Path,
        stem: str,
        suffix: str,
        ext: str,
        forbidden_paths: Optional[List[Path]] = None
    ) -> Path:
        """
        Generates a non-conflicting path:
        {stem}{suffix}{ext}
        If already exists or forbidden, appends ' (1)', ' (2)', etc.
        """
        forbidden = [p.resolve() for p in forbidden_paths] if forbidden_paths else []
        candidate = base_dir / f"{stem}{suffix}{ext}"
        if not candidate.exists() and candidate.resolve() not in forbidden:
            return candidate

        idx = 1
        while True:
            candidate = base_dir / f"{stem}{suffix} ({idx}){ext}"
            if not candidate.exists() and candidate.resolve() not in forbidden:
                return candidate
            idx += 1

    def run_qa_verification(self, context: JobContext) -> JobContext:
        """Executes comprehensive QA evaluation across all Phase 10 validation dimensions."""
        from engine.qa.qa_service import QAService
        qa_service = QAService(
            reader=self.reader,
            writer=self.writer,
            analyzer=self.analyzer,
            comparer=self.comparer
        )
        context.qa_result = qa_service.evaluate_job_context(context)
        return context


