"""
SlideGuard Quality & Accessibility Scoring Engine.
Calculates documented, auditable accessibility scores based on rule evaluation results.
"""

from typing import List
from models.rule_result import RuleResult
from models.job import QualityMetrics
from models.enums import RuleStatus, IssueSeverity
from engine.common.interfaces import IScoringEngine


class AccessibilityScorer(IScoringEngine):
    """
    Transparent scoring engine.
    Calculates metrics by evaluating rule pass/fail ratios and applying
    documented severity deductions per slide.
    
    Formula:
      Weighted Deductions = (Critical * 15 + Error * 8 + Warning * 2)
      Normalized Deduction = min(100.0, Weighted Deductions / max(1, total_slides))
      Accessibility Score = max(0.0, 100.0 - Normalized Deduction)
    """

    def __init__(
        self,
        critical_weight: float = 15.0,
        error_weight: float = 8.0,
        warning_weight: float = 2.0,
        info_weight: float = 0.0
    ):
        self.critical_weight = critical_weight
        self.error_weight = error_weight
        self.warning_weight = warning_weight
        self.info_weight = info_weight

        # Rule importance factors for Quality Score V2
        self.rule_importance = {
            "SG-TTL-01": 1.25,
            "SG-RO-01": 1.20,
            "SG-ALT-01": 1.20,
            "SG-TBL-01": 1.10,
            "SG-PH-01": 1.05,
            "SG-CNT-01": 1.00,
            "SG-LNG-01": 1.00,
            "SG-META-01": 0.95,
            "SG-LST-01": 0.90,
            "SG-GRP-01": 0.85,
            "SG-DEC-01": 0.80,
            "SG-RDB-01": 0.80,
            "SG-TXT-01": 0.75,
        }

    def calculate_metrics(self, rule_results: List[RuleResult], total_slides: int = 1) -> QualityMetrics:
        metrics = QualityMetrics(total_rules_evaluated=len(rule_results))

        raw_v2_deductions = 0.0

        for res in rule_results:
            if res.status == RuleStatus.PASSED:
                metrics.passed_rules += 1
            elif res.status == RuleStatus.FAILED:
                metrics.failed_rules += 1
            elif res.status == RuleStatus.WARNING:
                metrics.warning_rules += 1

            rule_factor = self.rule_importance.get(res.rule_id, 1.0)

            for issue in res.issues:
                metrics.total_issues += 1
                sev_weight = 0.0
                if issue.severity == IssueSeverity.CRITICAL:
                    metrics.critical_issues += 1
                    sev_weight = self.critical_weight
                elif issue.severity == IssueSeverity.ERROR:
                    metrics.error_issues += 1
                    sev_weight = self.error_weight
                elif issue.severity == IssueSeverity.WARNING:
                    metrics.warning_issues += 1
                    sev_weight = self.warning_weight
                elif issue.severity == IssueSeverity.INFO:
                    metrics.info_issues += 1
                    sev_weight = self.info_weight

                # Confidence factor for V2
                conf = getattr(issue, "confidence", "HIGH")
                conf_factor = 1.0 if conf == "HIGH" else (0.75 if conf in ("MEDIUM", "MODERATE") else 0.5)

                # Status factor for V2
                st = getattr(issue, "status", None)
                st_val = st.value if hasattr(st, "value") else str(st)
                status_factor = 1.0
                if st_val == "REMEDIATED":
                    status_factor = 0.15
                    metrics.remediated_issues += 1
                elif st_val == "MANUAL_REVIEW":
                    metrics.manual_review_issues += 1
                    status_factor = 0.90
                elif st_val == "UNSUPPORTED":
                    metrics.unsupported_issues += 1
                    status_factor = 1.10
                else:
                    metrics.unresolved_issues += 1

                if conf == "HIGH":
                    metrics.high_confidence_issues += 1
                elif conf in ("MEDIUM", "MODERATE"):
                    metrics.medium_confidence_issues += 1
                else:
                    metrics.low_confidence_issues += 1

                raw_v2_deductions += (sev_weight * rule_factor * conf_factor * status_factor)

        slides_count = max(1, total_slides)

        # Baseline V1 Calculation (Preserved for compatibility)
        raw_deductions = (
            (metrics.critical_issues * self.critical_weight) +
            (metrics.error_issues * self.error_weight) +
            (metrics.warning_issues * self.warning_weight) +
            (metrics.info_issues * self.info_weight)
        )
        normalized_deduction = min(100.0, raw_deductions / slides_count)
        metrics.accessibility_score = max(0.0, 100.0 - normalized_deduction)

        # Quality Score V2 Calculation
        normalized_v2 = min(100.0, raw_v2_deductions / slides_count)
        metrics.quality_score_v2 = round(max(0.0, 100.0 - normalized_v2), 1)

        return metrics

    def calculate_metrics_v2(
        self,
        unified_issues: list,
        total_slides: int = 1,
        total_rules: int = 14
    ) -> QualityMetrics:
        """Calculates Quality Score V2 directly from unified correlated issues."""
        metrics = QualityMetrics(total_rules_evaluated=total_rules)
        metrics.correlated_issues_count = len(unified_issues)

        slides_count = max(1, total_slides)
        raw_v2_deductions = 0.0

        for ui in unified_issues:
            metrics.total_issues += 1
            rule_factor = self.rule_importance.get(ui.primary_rule_id, 1.0)

            sev_weight = 0.0
            if ui.severity == IssueSeverity.CRITICAL:
                metrics.critical_issues += 1
                sev_weight = self.critical_weight
            elif ui.severity == IssueSeverity.ERROR:
                metrics.error_issues += 1
                sev_weight = self.error_weight
            elif ui.severity == IssueSeverity.WARNING:
                metrics.warning_issues += 1
                sev_weight = self.warning_weight
            elif ui.severity == IssueSeverity.INFO:
                metrics.info_issues += 1
                sev_weight = self.info_weight

            conf_factor = 1.0 if ui.confidence.value == "HIGH" else (0.75 if ui.confidence.value == "MEDIUM" else 0.5)
            if ui.confidence.value == "HIGH":
                metrics.high_confidence_issues += 1
            elif ui.confidence.value == "MEDIUM":
                metrics.medium_confidence_issues += 1
            else:
                metrics.low_confidence_issues += 1

            status_factor = 1.0
            rec_val = getattr(ui.reconciliation_status, "value", str(ui.reconciliation_status)) if hasattr(ui, "reconciliation_status") else ""
            st_val = getattr(ui.status, "value", str(ui.status)) if hasattr(ui, "status") else ""
            if rec_val in ("REMEDIATED", "RESOLVED") or st_val == "REMEDIATED":
                status_factor = 0.15
                metrics.remediated_issues += 1
            elif rec_val == "MANUAL_REVIEW" or st_val == "MANUAL_REVIEW":
                status_factor = 0.90
                metrics.manual_review_issues += 1
            elif rec_val == "UNSUPPORTED" or st_val == "UNSUPPORTED":
                status_factor = 1.10
                metrics.unsupported_issues += 1
            else:
                metrics.unresolved_issues += 1

            raw_v2_deductions += (sev_weight * rule_factor * conf_factor * status_factor)

        # Baseline V1
        raw_deductions = (
            (metrics.critical_issues * self.critical_weight) +
            (metrics.error_issues * self.error_weight) +
            (metrics.warning_issues * self.warning_weight) +
            (metrics.info_issues * self.info_weight)
        )
        normalized_deduction = min(100.0, raw_deductions / slides_count)
        metrics.accessibility_score = max(0.0, 100.0 - normalized_deduction)

        # V2
        normalized_v2 = min(100.0, raw_v2_deductions / slides_count)
        metrics.quality_score_v2 = round(max(0.0, 100.0 - normalized_v2), 1)

        return metrics
