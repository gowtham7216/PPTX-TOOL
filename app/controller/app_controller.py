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

    def initialize_job(self, file_path: Path) -> JobContext:
        """Sets up sandboxed workspace and creates an isolated working copy."""
        logger.info("Initializing job for %s", file_path)
        context = self.workspace_mgr.create_job(file_path)

        # Parse initial presentation structure
        self.current_presentation = self.reader.read(context.working_file)
        context.total_slides = self.current_presentation.total_slides
        context.slides_analyzed = 0
        return context

    def execute_analysis(
        self,
        context: JobContext,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> JobContext:
        """Executes full rule-based accessibility analysis."""
        if self.current_presentation is None:
            self.current_presentation = self.reader.read(context.working_file)

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
        progress_callback: Optional[Callable[[int, int, str], None]] = None
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
            progress_callback=progress_callback
        )
        context.remediation_actions = actions

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
        # Export remediated PPTX and Report into destination directory with collision protection
        try:
            source_p = getattr(context, "source_file", None) or context.original_file
            dest_dir = source_p.parent if source_p.parent.exists() else context.workspace_dir
            stem = source_p.stem

            final_pptx = self.get_collision_safe_path(
                base_dir=dest_dir,
                stem=stem,
                suffix="_SlideGuard_Remediated",
                ext=".pptx",
                forbidden_paths=[source_p]
            )

            # Copy candidate to final production output
            if context.candidate_file and context.candidate_file.exists():
                import shutil
                shutil.copy2(context.candidate_file, final_pptx)
                context.final_output_file = final_pptx
                logger.info("Exported production remediated presentation: %s", final_pptx)

            # Export final report alongside output
            if context.reports_dir and (context.reports_dir / "final_report.md").exists():
                final_rep = self.get_collision_safe_path(
                    base_dir=dest_dir,
                    stem=stem,
                    suffix="_SlideGuard_Report",
                    ext=".md",
                    forbidden_paths=[source_p]
                )
                import shutil
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


