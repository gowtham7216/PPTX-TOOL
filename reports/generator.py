"""
SlideGuard Report Generator.
Generates comprehensive Markdown and JSON reports detailing accessibility analysis,
reading order sequences, remediation logs, quality scores, and integrity checks.
"""

import json
from pathlib import Path
from typing import Dict, Any
from models.job import JobContext


class ReportGenerator:
    """Generates structured audit and verification reports for SlideGuard jobs."""

    def generate_json_report(self, context: JobContext, output_file: Path) -> Path:
        """Saves machine-readable audit report in JSON format."""
        data = self._build_report_dict(context)
        target = Path(output_file).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        with open(target, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        return target

    def generate_markdown_report(self, context: JobContext, output_file: Path) -> Path:
        """Generates human-readable Markdown accessibility and verification report."""
        m = context.metrics
        c = context.comparison_result
        v = context.verification_result

        stage_val = getattr(context.stage, "value", str(context.stage))

        # Section 1: Executive Summary & Quality Metrics
        pre_q2 = context.pre_remediation_score if context.pre_remediation_score is not None else m.quality_score_v2
        post_q2 = context.post_remediation_score if context.post_remediation_score is not None else (m.quality_score_v2 if (stage_val == "COMPLETED" or context.remediation_actions) else None)
        pre_raw = getattr(context, "pre_accessibility_score", None) if getattr(context, "pre_accessibility_score", None) is not None else m.accessibility_score
        post_raw = getattr(context, "post_accessibility_score", None)
        post_m = getattr(context, "post_metrics", None)
        intel_summary = getattr(context, "intelligence_summary", None) or {}

        pre_q2_str = f"{pre_q2:.1f} / 100" if pre_q2 is not None else "N/A"
        post_q2_str = f"{post_q2:.1f} / 100" if post_q2 is not None else "Pending Remediation"
        pre_raw_str = f"{pre_raw:.1f} / 100" if pre_raw is not None else "N/A"
        post_raw_str = f"{post_raw:.1f} / 100" if post_raw is not None else "Pending Remediation"

        total_unique_pre = m.correlated_issues_count or (len(context.unified_issues) if context.unified_issues else m.total_issues)
        total_unique_post = intel_summary.get("total_post", m.unresolved_issues) if intel_summary else m.unresolved_issues
        unique_remediated = intel_summary.get("remediated_count", m.remediated_issues)
        unique_unresolved = total_unique_post

        raw_pre_count = m.total_issues
        raw_post_count = post_m.total_issues if post_m else (raw_pre_count - m.raw_remediated_issues if m.raw_remediated_issues else raw_pre_count)
        raw_remediated_count = m.raw_remediated_issues or max(0, raw_pre_count - raw_post_count)

        total_post_issues_str = str(unique_unresolved) if stage_val == "COMPLETED" or context.remediation_actions else "Pending"
        remediated_str = str(unique_remediated)

        md_lines = [
            f"# SlideGuard Accessibility & Integrity Report",
            f"",
            f"**Job ID:** `{context.job_id}`  ",
            f"**File Name:** `{context.original_file.name}`  ",
            f"**Status:** `{stage_val}`  ",
            f"**Execution Duration:** `{context.duration_seconds:.2f}s`  ",
            f"",
            f"---",
            f"",
            f"## 1. Executive Summary & Quality Metrics",
            f"",
            f"| Metric | Baseline (Pre-Remediation) | Candidate (Post-Remediation) |",
            f"|---|---|---|",
            f"| **Authoritative Quality Score (Multi-Signal V2)** | **{pre_q2_str}** | **{post_q2_str}** |",
            f"| **Raw Severity Accessibility Score (V1)** | {pre_raw_str} | {post_raw_str} |",
            f"| **Total Slides** | {context.total_slides} | {context.total_slides} |",
            f"| **Slides Analyzed** | {context.slides_analyzed} | {context.slides_analyzed} |",
            f"| **Total Unique Issues** | {total_unique_pre} | {total_unique_post} |",
            f"| **Remediated & Verified Issues** | - | **{remediated_str}** |",
            f"| **Preserved / Manual Review Issues** | - | {total_post_issues_str} |",
            f"| **Raw Rule Check Violations** | {raw_pre_count} | {raw_post_count} |",
            f"| - Remediated Raw Rule Checks | - | **{raw_remediated_count}** |",
            f"| - Critical | {m.critical_issues} | {post_m.critical_issues if post_m else '-'} |",
            f"| - Error | {m.error_issues} | {post_m.error_issues if post_m else '-'} |",
            f"| - Warning | {m.warning_issues} | {post_m.warning_issues if post_m else '-'} |",
            f"| - Info | {m.info_issues} | {post_m.info_issues if post_m else '-'} |",
            f"| **Rules Evaluated** | {m.total_rules_evaluated} | {post_m.total_rules_evaluated if post_m else m.total_rules_evaluated} |",
            f"| - Rules Passed | {m.passed_rules} | {post_m.passed_rules if post_m else m.passed_rules} |",
            f"| - Rules Failed | {m.failed_rules} | {post_m.failed_rules if post_m else m.failed_rules} |",
            f"| - Rules Warning | {m.warning_rules} | {post_m.warning_rules if post_m else m.warning_rules} |",
            f"| **Design Preservation Verified** | {'✅ Yes' if (c and c.integrity_preserved) else '⚠️ Not Verified'} | {'✅ Yes' if (c and c.integrity_preserved) else '⚠️ Not Verified'} |",
            f"",
            f"> **Scoring Methodology & Metric Transparency:**",
            f"> - **Authoritative Unique Issues vs Raw Checks:** SlideGuard distinguishes between *Authoritative Unique Issues* (correlated structural issues mapped to distinct slide/shape targets) and *Raw Rule Check Violations* (individual rule check instances). For example, multiple rules targeting the same slide reading order (`SG-RO-01` and `SG-RO-VAL-01`) or title placeholder (`SG-TTL-01` and `SG-PH-01`) represent 1 unique issue each, not separate issues.",
            f"> - **Authoritative Quality Score (Multi-Signal V2):** Multi-signal weighted metric incorporating rule criticality, WCAG conformance levels, unified issue clustering, and visual signal confidence. Score ranges from 0 to 100.",
            f"> - **Raw Severity Score (V1):** Deterministic penalty formula: `100 - (Critical×15 + Error×8 + Warning×2 + Info×0.5) / Total Slides`.",
            f"> - **Verification Guarantee:** Remediated status is certified only after re-analyzing the candidate presentation OOXML. Items requiring human aesthetic judgment (e.g. brand palette contrast, layout reflow, missing alt-text authoring) are preserved as `MANUAL_REVIEW` without forced alterations.",
            f"",
            f"---",
            f"",
            f"## 2. Reading Order Analysis & Proposal",
            f"",
        ]

        # Extract reading order issues
        ro_issues = []
        for res in context.rule_results:
            if res.rule_id == "SG-RO-01":
                ro_issues.extend(res.issues)

        detailed_acts = getattr(context, "detailed_remediation_actions", None) or context.remediation_actions
        ro_actions_by_slide = {
            a.slide_number: a for a in detailed_acts
            if getattr(a, "rule_id", "") == "SG-RO-01"
        }

        if not ro_issues:
            md_lines.append("All slide reading sequences are logically compliant; no reordering required.\n")
        else:
            if ro_actions_by_slide:
                applied_ro = sum(1 for a in ro_actions_by_slide.values() if getattr(a.status, "value", str(a.status)) == "APPLIED")
                no_changes_ro = sum(1 for a in ro_actions_by_slide.values() if getattr(a.status, "value", str(a.status)) == "NO_CHANGES")
                skipped_ro = sum(1 for a in ro_actions_by_slide.values() if getattr(a.status, "value", str(a.status)) == "SKIPPED")
                failed_ro = sum(1 for a in ro_actions_by_slide.values() if getattr(a.status, "value", str(a.status)) == "FAILED")
                verified_ro = sum(1 for i in ro_issues if getattr(i, "verification_status", "") == "VERIFIED")

                md_lines.extend([
                    f"### Reading Order Remediation Summary",
                    f"",
                    f"| Reading Order Metric | Count |",
                    f"|---|---|",
                    f"| **Total Flagged Slides** | {len(ro_issues)} |",
                    f"| **Applied in Candidate XML** | {applied_ro} |",
                    f"| **No Changes Required** | {no_changes_ro} |",
                    f"| **Skipped** | {skipped_ro} |",
                    f"| **Failed** | {failed_ro} |",
                    f"| **Verified in Candidate Copy** | {verified_ro} |",
                    f"",
                ])

            md_lines.extend([
                f"| Slide # | Confidence | Original Order | Proposed Order | Footers Excluded | Remediation Status | Verification | Reasoning |",
                f"|---|---|---|---|---|---|---|---|",
            ])
            for issue in ro_issues:
                extra = issue.extra_data or {}
                orig = extra.get("original_order", [])
                prop = extra.get("proposed_order", [])
                conf = extra.get("confidence", "MEDIUM")
                footers = extra.get("excluded_footer_objects", [])
                reasons = extra.get("reasoning", [])

                act = ro_actions_by_slide.get(issue.slide_number)
                if act:
                    act_stat = getattr(act.status, "value", str(act.status))
                elif getattr(issue, "status", None) and (getattr(issue.status, "value", str(issue.status)) == "REMEDIATED"):
                    act_stat = "APPLIED"
                else:
                    act_stat = "DETECTED"

                verif = getattr(issue, "verification_status", "PENDING") or "PENDING"
                if verif == "VERIFIED" or (act and getattr(act.status, "value", str(act.status)) == "APPLIED" and verif != "FAILED"):
                    verif_display = "✅ Verified in Candidate XML"
                elif verif == "FAILED":
                    verif_display = "❌ Verification Failed"
                elif act_stat == "NO_CHANGES":
                    verif_display = "✅ Optimal Sequence"
                elif act_stat == "SKIPPED":
                    verif_display = "⚠️ Skipped"
                else:
                    verif_display = "⏳ Pending Verification"

                f_desc = ", ".join(f"Shape {f['shape_id']} ({f.get('text', '')[:15]})" for f in footers) if footers else "None"
                r_desc = "; ".join(reasons[:2]) if reasons else issue.description

                md_lines.append(
                    f"| {issue.slide_number} | **{conf}** | `{orig}` | `{prop}` | {f_desc} | `{act_stat}` | {verif_display} | {r_desc} |"
                )
            md_lines.append("")

        md_lines.extend([
            f"---",
            f"",
            f"## 3. Selective Ungrouping & Hierarchy Analysis",
            f"",
        ])

        # Extract ungrouping issues
        grp_issues = []
        for res in context.rule_results:
            if res.rule_id == "SG-GRP-01":
                grp_issues.extend(res.issues)

        if not grp_issues:
            md_lines.append("No grouped objects requiring ungrouping detected in this presentation.\n")
        else:
            md_lines.extend([
                f"| Slide # | Group ID | Category | Eligible Text/Shapes | Protected Pictures | Safety Status | Proposed Action |",
                f"|---|---|---|---|---|---|---|",
            ])
            for issue in grp_issues:
                extra = issue.extra_data or {}
                grp_id = issue.shape_id or extra.get("group_id", "N/A")
                cat = extra.get("category", "MIXED")
                el_cnt = extra.get("eligible_count", 0)
                pic_cnt = extra.get("picture_count", 0)
                is_safe = extra.get("is_safe_to_remediate", True)
                safety = "✅ Safe to Ungroup" if is_safe else f"⚠️ Preserved ({extra.get('skip_reason', 'Manual')})"
                action = f"Extract {el_cnt} text/shape(s); protect {pic_cnt} picture(s)" if is_safe else "Preserve group container intact"

                md_lines.append(
                    f"| {issue.slide_number} | `{grp_id}` | `{cat}` | {el_cnt} | {pic_cnt} | {safety} | {action} |"
                )
            md_lines.append("")

        md_lines.extend([
            f"---",
            f"",
            f"## 4. Placeholder Conversion & Analysis",
            f"",
        ])

        t4_report = getattr(context, "task_4_placeholder_report", None)
        if t4_report:
            md_lines.extend([
                f"### Task 4 Option C Conversion Audit Summary",
                f"",
                f"| Metric | Value |",
                f"|---|---|",
                f"| **Total Slides Evaluated** | {t4_report.get('total_slides_evaluated', 0)} |",
                f"| **Slides Changed** | {t4_report.get('slides_changed', 0)} |",
                f"| **Slides Unchanged** | {t4_report.get('slides_unchanged', 0)} |",
                f"| **Candidates Detected** | {t4_report.get('candidates_detected', 0)} |",
                f"| **Conversions Proposed** | {t4_report.get('conversions_proposed', 0)} |",
                f"| **Conversions Approved** | {t4_report.get('conversions_approved', 0)} |",
                f"| **Conversions Applied** | {t4_report.get('conversions_applied', 0)} |",
                f"| **Conversions Verified** | {t4_report.get('conversions_verified', 0)} |",
                f"| **Conversions Rolled Back** | {t4_report.get('conversions_rolled_back', 0)} |",
                f"| **Existing Objects Converted** | {t4_report.get('existing_objects_converted', 0)} |",
                f"| **New Placeholders Created** | {t4_report.get('new_placeholders_created', 0)} |",
                f"| **Native Placeholders Preserved** | {t4_report.get('native_placeholders_preserved', 0)} |",
                f"| **Candidates Skipped** | {t4_report.get('candidates_skipped', 0)} |",
                f"| **Candidates Manual Review** | {t4_report.get('candidates_manual_review', 0)} |",
                f"| **Failed Conversions** | {t4_report.get('failed_conversions', 0)} |",
                f"| **Visual Integrity Verified** | {'✅ Yes' if t4_report.get('visual_integrity_verified') else '⚠️ Unverified'} |",
                f"",
            ])

            t4_objects = t4_report.get("objects", [])
            if t4_objects:
                md_lines.extend([
                    f"### Detailed Placeholder & Candidate Records",
                    f"",
                    f"| Slide | Shape ID | Name | Role | Action | Proposed Type | Basis / Reason | Confidence |",
                    f"|---|---|---|---|---|---|---|---|",
                ])
                for obj in t4_objects:
                    s_id = obj.get("shape_id", "N/A")
                    name = obj.get("shape_name", "Shape")
                    role = obj.get("semantic_role", obj.get("category", "OTHER"))
                    act = obj.get("action", "PRESERVE_EXISTING")
                    p_type = obj.get("proposed_placeholder_type", "none")
                    basis = obj.get("conversion_basis", "") or obj.get("reason", "")
                    conf = obj.get("confidence", "HIGH")
                    md_lines.append(
                        f"| {obj.get('slide_number')} | `{s_id}` | {name} | `{role}` | **{act}** | `{p_type}` | {basis[:40]} | {conf} |"
                    )
                md_lines.append("")
        else:
            # Extract placeholder issues from rule results fallback
            ph_issues = []
            for res in context.rule_results:
                if res.rule_id == "SG-PH-01":
                    ph_issues.extend(res.issues)

            if not ph_issues:
                md_lines.append("All slide placeholders are structurally compliant; no conversions required.\n")
            else:
                md_lines.extend([
                    f"| Slide # | Shape ID | Name | Category | Status | Inherited From | Proposed Action / Reason |",
                    f"|---|---|---|---|---|---|---|",
                ])
                for issue in ph_issues:
                    extra = issue.extra_data or {}
                    s_id = issue.shape_id or extra.get("shape_id", "N/A")
                    s_name = issue.shape_name or extra.get("shape_name", "Shape")
                    cat = extra.get("category", "OTHER")
                    st = extra.get("status", "PRESERVED")
                    inh = extra.get("inherited_from", "slide")
                    reason = extra.get("reason", issue.description)
                    rec = extra.get("recommendation", issue.recommendation)
                    status_icon = "✅ Safe to Convert" if st == "SAFE_TO_CONVERT" else ("⚠️ Manual Review" if st == "MANUAL_REVIEW" else "🛡️ Preserved")
                    action_text = rec if st == "SAFE_TO_CONVERT" else reason

                    md_lines.append(
                        f"| {issue.slide_number} | `{s_id}` | {s_name} | `{cat}` | {status_icon} | {inh} | {action_text} |"
                    )
                md_lines.append("")

        md_lines.extend([
            f"---",
            f"",
            f"## 5. List Semantics & Numbering",
            f"",
        ])

        # Extract list issues
        lst_issues = []
        for res in context.rule_results:
            if res.rule_id == "SG-LST-01":
                lst_issues.extend(res.issues)

        if not lst_issues:
            md_lines.append("All slide text structures comply with native list semantics; no list remediation required.\n")
        else:
            md_lines.extend([
                f"| Slide # | Shape ID | Paragraph | Category | Status | Confidence | Proposed Action / Reason |",
                f"|---|---|---|---|---|---|---|",
            ])
            for issue in lst_issues:
                extra = issue.extra_data or {}
                s_id = issue.shape_id or extra.get("shape_id", "N/A")
                p_idx = extra.get("paragraph_index", 0) + 1
                cat = extra.get("category", "NONE")
                st = extra.get("status", "NOT_A_LIST")
                conf = extra.get("confidence", "HIGH")
                act = extra.get("proposed_action", issue.recommendation)
                status_icon = "✅ Safe to Remediate" if st == "SAFE_TO_REMEDIATE" else ("⚠️ Manual Review" if st == "MANUAL_REVIEW" else "🛡️ Valid")

                md_lines.append(
                    f"| {issue.slide_number} | `{s_id}` | P{p_idx} | `{cat}` | {status_icon} | {conf} | {act} |"
                )
            md_lines.append("")

        md_lines.extend([
            f"---",
            f"",
            f"## 6. Accessibility Validation & Remediation",
            f"",
        ])

        # Calculate accessibility validation metrics
        total_rules = m.total_rules_evaluated
        passed_rules = m.passed_rules
        err_count = m.error_issues
        warn_count = m.warning_issues
        info_count = m.info_issues

        all_issues = [issue for res in context.rule_results for issue in res.issues]
        remediated_count = sum(1 for i in all_issues if getattr(i, "status", None) and (getattr(i.status, "value", str(i.status)) == "REMEDIATED"))
        verified_count = sum(1 for i in all_issues if getattr(i, "verification_status", "") == "VERIFIED")
        safe_auto_count = sum(1 for i in all_issues if i.behavior.value == "SAFE_AUTO_REMEDIATE" and getattr(i, "status", None) and (i.status.value == "SAFE_TO_REMEDIATE" or i.status == "SAFE_TO_REMEDIATE"))
        manual_review_count = sum(1 for i in all_issues if getattr(i, "status", None) and (getattr(i.status, "value", str(i.status)) == "MANUAL_REVIEW"))
        detect_only_count = sum(1 for i in all_issues if i.behavior.value == "DETECT_ONLY")
        unsupported_count = sum(1 for i in all_issues if getattr(i, "status", None) and (getattr(i.status, "value", str(i.status)) == "UNSUPPORTED"))
        remediation_failures = sum(1 for a in context.remediation_actions if a.status.value == "FAILED")

        md_lines.extend([
            f"| Metric | Count |",
            f"|---|---|",
            f"| **Total Rules Evaluated** | {total_rules} |",
            f"| **Rules Passed** | {passed_rules} |",
            f"| **Errors** | {err_count} |",
            f"| **Warnings** | {warn_count} |",
            f"| **Info** | {info_count} |",
            f"| **Remediated Issues** | {remediated_count} |",
            f"| **Verified in Candidate** | {verified_count} |",
            f"| **Safe Auto-Remediations** | {safe_auto_count} |",
            f"| **Manual-Review Items** | {manual_review_count} |",
            f"| **Detect-Only Issues** | {detect_only_count} |",
            f"| **Unsupported Cases** | {unsupported_count} |",
            f"| **Remediation Failures** | {remediation_failures} |",
            f"",
        ])

        if not all_issues:
            md_lines.append("No accessibility violations detected in this presentation.\n")
        else:
            md_lines.extend([
                f"| Slide | Shape | Rule ID & Name | Severity | WCAG Criterion | Description | Recommendation | Action / Status | Confidence | Verification |",
                f"|---|---|---|---|---|---|---|---|---|---|",
            ])
            for issue in all_issues:
                st_val = getattr(issue.status, "value", str(issue.status)) if hasattr(issue, "status") else "DETECTED"
                conf_val = getattr(issue, "confidence", "HIGH")
                verif_val = getattr(issue, "verification_status", "PENDING") or "PENDING"
                shape_label = f"`{issue.shape_id}`" if issue.shape_id is not None else "Slide"

                md_lines.append(
                    f"| {issue.slide_number} | {shape_label} | `{issue.rule_id}` {issue.rule_name} | **{issue.severity.value}** | {issue.wcag_mapping or 'N/A'} | {issue.description} | {issue.recommendation} | `{st_val}` | {conf_val} | {verif_val} |"
                )
            md_lines.append("")

        md_lines.extend([
            f"---",
            f"",
            f"## 7. Accessibility Intelligence Summary",
            f"",
        ])

        u_issues = context.unified_issues or []
        intel_summary = context.intelligence_summary or {}
        raw_issues_count = m.total_issues
        correlated_count = len(u_issues)
        dedup_count = max(0, raw_issues_count - correlated_count)
        high_conf = sum(1 for u in u_issues if getattr(u.confidence, "value", str(u.confidence)) == "HIGH")
        med_conf = sum(1 for u in u_issues if getattr(u.confidence, "value", str(u.confidence)) == "MEDIUM")
        low_conf = sum(1 for u in u_issues if getattr(u.confidence, "value", str(u.confidence)) == "LOW")
        safe_candidates = sum(1 for u in u_issues if getattr(u.remediation_eligibility, "value", str(u.remediation_eligibility)) == "ELIGIBLE_AUTO")

        remediated_cnt = intel_summary.get("remediated_count", m.remediated_issues)
        unresolved_cnt = intel_summary.get("total_post", m.unresolved_issues) if intel_summary else m.unresolved_issues
        new_cnt = intel_summary.get("new_count", 0)
        manual_rev_cnt = intel_summary.get("manual_review_count", m.manual_review_issues)
        unsupported_cnt = intel_summary.get("unsupported_count", m.unsupported_issues)
        raw_remediated_cnt = intel_summary.get("raw_issues_remediated", m.raw_remediated_issues)

        score_before = f"{context.pre_remediation_score:.1f} / 100" if context.pre_remediation_score is not None else f"{m.quality_score_v2:.1f} / 100"
        stage_val = getattr(context.stage, "value", str(context.stage))
        if context.post_remediation_score is not None:
            score_after = f"{context.post_remediation_score:.1f} / 100"
        elif stage_val == "COMPLETED" or context.remediation_actions:
            score_after = f"{m.quality_score_v2:.1f} / 100"
        else:
            score_after = "Pending Remediation"

        raw_score_before = f"{context.pre_accessibility_score:.1f} / 100" if getattr(context, "pre_accessibility_score", None) is not None else f"{m.accessibility_score:.1f} / 100"
        if getattr(context, "post_accessibility_score", None) is not None:
            raw_score_after = f"{context.post_accessibility_score:.1f} / 100"
        elif stage_val == "COMPLETED" or context.remediation_actions:
            raw_score_after = f"{m.accessibility_score:.1f} / 100"
        else:
            raw_score_after = "Pending Remediation"

        md_lines.extend([
            f"| Intelligence Metric | Value |",
            f"|---|---|",
            f"| **Total Raw Issues** | {raw_issues_count} |",
            f"| **Correlated Issues** | {correlated_count} |",
            f"| **Deduplicated Issues** | {dedup_count} |",
            f"| **High-Confidence Issues** | {high_conf} |",
            f"| **Medium-Confidence Issues** | {med_conf} |",
            f"| **Low-Confidence Issues** | {low_conf} |",
            f"| **Safe Remediation Candidates** | {safe_candidates} |",
            f"| **Remediated Issues** | {remediated_cnt} |",
            f"| **Unresolved Issues** | {unresolved_cnt} |",
            f"| **New Issues** | {new_cnt} |",
            f"| **Manual-Review Issues** | {manual_rev_cnt} |",
            f"| **Unsupported Issues** | {unsupported_cnt} |",
            f"| **Raw Rule Checks Remediated** | {raw_remediated_cnt} |",
            f"| **Authoritative Quality Score (Before)** | {score_before} |",
            f"| **Authoritative Quality Score (After)** | {score_after} |",
            f"| **Raw Severity Accessibility Score (Before)** | {raw_score_before} |",
            f"| **Raw Severity Accessibility Score (After)** | {raw_score_after} |",
            f"",
        ])

        if not u_issues:
            md_lines.append("No unified accessibility intelligence issues to display.\n")
        else:
            md_lines.extend([
                f"### Correlated Unified Issues",
                f"",
                f"| Slide | Shape | Primary Rule | Related Rules | Severity | Confidence | Explanation | Recommended Action | Final Status | Verification |",
                f"|---|---|---|---|---|---|---|---|---|---|",
            ])
            for u in u_issues:
                shape_str = f"`{u.shape_id}`" if u.shape_id is not None else "Slide"
                related_str = ", ".join(f"`{r}`" for r in u.related_rule_ids) if u.related_rule_ids else "None"
                rec_stat = getattr(u.reconciliation_status, "value", str(u.reconciliation_status))
                conf_str = getattr(u.confidence, "value", str(u.confidence))
                sev_str = getattr(u.severity, "value", str(u.severity))
                verif_str = getattr(u, "verification_status", "PENDING") or "PENDING"
                md_lines.append(
                    f"| {u.slide_number} | {shape_str} | `{u.primary_rule_id}` | {related_str} | **{sev_str}** | **{conf_str}** | {u.explanation} | {u.recommended_action} | `{rec_stat}` | {verif_str} |"
                )
            md_lines.append("")

        md_lines.extend([
            f"---",
            f"",
            f"## 8. Remediation Pipeline Log",
            f"",
        ])

        if not context.remediation_actions:
            md_lines.append("Remediation was not executed for this run (Analysis only).\n")
        else:
            md_lines.extend([
                f"| Step ID | Target Slide | Status | Action Description | Confidence |",
                f"|---|---|---|---|---|",
            ])
            for action in context.remediation_actions:
                md_lines.append(
                    f"| `{action.rule_id}` | {action.slide_number} | `{action.status.value}` | {action.description} | {action.confidence:.0%} |"
                )
            md_lines.append("")

        if v:
            md_lines.extend([
                f"---",
                f"",
                f"## 9. Verification Check Details",
                f"",
                f"**Overall Verification:** {'✅ PASSED' if v.passed else '❌ FAILED'}  ",
                f"Passed: {v.passed_checks} / {v.total_checks} checks  ",
                f"",
                f"| Check Name | Slide | Status | Details |",
                f"|---|---|---|---|",
            ])
            for chk in v.checks:
                md_lines.append(
                    f"| {chk.name} | {chk.slide_number} | {'✅ Pass' if chk.passed else '❌ Fail'} | {chk.details} |"
                )
            md_lines.append("")

        # Phase 9: AI-Assisted Accessibility Authoring & Audit Trail
        ai_recs = getattr(context, "ai_recommendations", [])
        ai_audit = getattr(context, "ai_audit_trail", [])
        ai_summary = getattr(context, "ai_authoring_summary", {}) or {}

        md_lines.extend([
            f"---",
            f"",
            f"## 10. AI-Assisted Accessibility Authoring & Audit Trail",
            f"",
            f"> **Notice:** AI-generated authoring proposals (powered strictly by Gemini 3.8 Flash High) "
            f"require explicit human approval before deterministic OOXML application. AI assistance does not certify WCAG compliance.",
            f"",
            f"### Authoring Metrics Summary",
            f"",
            f"| Metric | Value |",
            f"|---|---|",
            f"| **Recommendations Generated** | {ai_summary.get('generated', len(ai_recs))} |",
            f"| **Recommendations Approved** | {ai_summary.get('approved', sum(1 for r in ai_recs if getattr(r, 'status', None) == 'APPROVED'))} |",
            f"| **Recommendations Edited** | {ai_summary.get('edited', sum(1 for r in ai_recs if getattr(r, 'status', None) == 'EDITED'))} |",
            f"| **Recommendations Rejected** | {ai_summary.get('rejected', sum(1 for r in ai_recs if getattr(r, 'status', None) == 'REJECTED'))} |",
            f"| **Recommendations Applied** | {ai_summary.get('applied', sum(1 for r in ai_recs if getattr(r, 'status', None) == 'APPLIED'))} |",
            f"| **Recommendations Failed** | {ai_summary.get('failed', sum(1 for r in ai_recs if getattr(r, 'status', None) == 'FAILED'))} |",
            f"| **Manual Reviews** | {ai_summary.get('manual_review', sum(1 for r in ai_recs if getattr(r, 'status', None) == 'MANUAL_REVIEW'))} |",
            f"| **AI Fallback Cases** | {ai_summary.get('fallback_cases', 0)} |",
            f"| **Malformed AI Responses** | {ai_summary.get('malformed_responses', 0)} |",
            f"",
        ])

        if not ai_recs:
            md_lines.append("No AI accessibility authoring recommendations generated for this run.\n")
        else:
            md_lines.extend([
                f"### Structured AI Authoring Recommendations",
                f"",
                f"| Rec ID | Slide | Shape | Authoring Type | Original Value | AI Suggestion | Confidence | Grounding | Decision Status |",
                f"|---|---|---|---|---|---|---|---|---|",
            ])
            for r in ai_recs:
                r_id = getattr(r, "recommendation_id", "N/A")
                sl_num = getattr(r, "slide_number", 0)
                sh_id = getattr(r, "shape_id", "Slide")
                a_type = getattr(r, "authoring_type", "")
                if hasattr(a_type, "value"):
                    a_type = a_type.value
                orig = getattr(r, "original_value", "") or getattr(r, "current_value", "")
                sugg = getattr(r, "suggested_value", "")
                conf = getattr(r, "confidence", 0.0)
                ground = getattr(r, "fact_grounding", "SUPPORTED_FACT")
                if hasattr(ground, "value"):
                    ground = ground.value
                stat = getattr(r, "status", "PENDING_REVIEW")
                if hasattr(stat, "value"):
                    stat = stat.value
                md_lines.append(
                    f"| `{r_id}` | {sl_num} | `{sh_id}` | `{a_type}` | {orig[:30]} | {sugg[:45]} | {conf:.0%} | `{ground}` | **{stat}** |"
                )
            md_lines.append("")

        if ai_audit:
            md_lines.extend([
                f"### Authoring Audit Trail (Lineage Log)",
                f"",
                f"| Audit ID | Recommendation ID | Model | Type | Slide | Shape | AI Suggestion | User Decision | Applied Value | Timestamp |",
                f"|---|---|---|---|---|---|---|---|---|---|",
            ])
            for aud in ai_audit:
                md_lines.append(
                    f"| `{aud.get('audit_id')}` | `{aud.get('recommendation_id')}` | `{aud.get('model')}` | `{aud.get('authoring_type')}` | {aud.get('slide_number')} | `{aud.get('shape_id')}` | {aud.get('ai_suggestion')[:30]} | **{aud.get('user_decision')}** | {aud.get('final_applied_value')[:30]} | {aud.get('timestamp')[:19]} |"
                )
            md_lines.append("")

        # Task 3: AI-Assisted Alternative Text & Decorative Image Detection Report
        t3_report = getattr(context, "task_3_report", {}) or {}
        t3_summary = t3_report.get("summary", {})
        t3_images = t3_report.get("changed_images", [])

        if t3_summary or t3_images:
            md_lines.extend([
                f"---",
                f"",
                f"## 11. Task 3: Alternative Text & Decorative Image Report",
                f"",
                f"> **Model Traceability:** SlideGuard runtime semantic AI: `Gemini 3.5` | Fallback: `Gemini 3.1` | Implementation: `Gemini 3.8 Flash High`",
                f"",
                f"### Image Accessibility Metrics",
                f"",
                f"| Metric | Value |",
                f"|---|---|",
                f"| **Total Image Objects** | {t3_summary.get('total_images', len(t3_images))} |",
                f"| **Meaningful Images** | {t3_summary.get('meaningful_images', 0)} |",
                f"| **Decorative Images** | {t3_summary.get('decorative_images', 0)} |",
                f"| **Unknown Images** | {t3_summary.get('unknown_images', 0)} |",
                f"| **Existing Good Alt Text** | {t3_summary.get('existing_good_alt_text', 0)} |",
                f"| **Existing Missing Alt Text** | {t3_summary.get('existing_missing_alt_text', 0)} |",
                f"| **Generated Alt Text** | {t3_summary.get('generated_alt_text', 0)} |",
                f"| **Images Marked Decorative** | {t3_summary.get('images_marked_decorative', 0)} |",
                f"| **Manual Review Items** | {t3_summary.get('manual_review_items', 0)} |",
                f"| **High-Confidence Decisions (>=0.85)** | {t3_summary.get('high_confidence_decisions', 0)} |",
                f"| **Medium-Confidence Decisions (0.70-0.849)** | {t3_summary.get('medium_confidence_decisions', 0)} |",
                f"| **Low-Confidence Decisions (<0.70)** | {t3_summary.get('low_confidence_decisions', 0)} |",
                f"| **Animation-Protected Images** | {t3_summary.get('animation_protected_images', 0)} |",
                f"| **Group-Protected Images** | {t3_summary.get('group_protected_images', 0)} |",
                f"| **Failed Images** | {t3_summary.get('failed_images', 0)} |",
                f"",
            ])

            if t3_images:
                md_lines.extend([
                    f"### Detailed Image Evaluation Records",
                    f"",
                    f"| Slide | Shape ID | Shape Name | Classification | Action | Basis | Quality | Alt Text | Conf | Reason | Status |",
                    f"|---|---|---|---|---|---|---|---|---|---|---|",
                ])
                for img in t3_images:
                    new_txt = img.get('new_alt_text') or img.get('original_alt_text') or "[None]"
                    conf_pct = f"{img.get('confidence', 0):.0%}"
                    basis = img.get('generation_basis', 'STRUCTURAL_CONTEXT')
                    q_class = img.get('quality_class', 'MANUAL_REVIEW')
                    md_lines.append(
                        f"| {img.get('slide_number')} | `{img.get('shape_id')}` | {img.get('shape_name')} | **{img.get('classification')}** | `{img.get('action')}` | `{basis}` | `{q_class}` | {new_txt[:30]} | {conf_pct} | {img.get('reason')} | `{img.get('verification_status')}` |"
                    )
                md_lines.append("")

        target = Path(output_file).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("\n".join(md_lines), encoding="utf-8")
        return target

    def _build_report_dict(self, context: JobContext) -> Dict[str, Any]:
        return {
            "job_id": context.job_id,
            "original_file": str(context.original_file),
            "stage": getattr(context.stage, "value", str(context.stage)),
            "duration_seconds": context.duration_seconds,
            "total_slides": context.total_slides,
            "slides_analyzed": context.slides_analyzed,
            "metrics": context.metrics.to_dict(),
            "rule_results": [res.to_dict() for res in context.rule_results],
            "remediation_actions": [act.to_dict() for act in context.remediation_actions],
            "verification_result": context.verification_result.to_dict() if context.verification_result else None,
            "comparison_result": context.comparison_result.to_dict() if context.comparison_result else None,
            "unified_issues": [u.to_dict() for u in context.unified_issues],
            "intelligence_summary": context.intelligence_summary,
            "pre_remediation_score": context.pre_remediation_score,
            "post_remediation_score": context.post_remediation_score,
            "pre_accessibility_score": getattr(context, "pre_accessibility_score", None),
            "post_accessibility_score": getattr(context, "post_accessibility_score", None),
            "post_metrics": context.post_metrics.to_dict() if getattr(context, "post_metrics", None) else None,
            "ai_authoring_summary": getattr(context, "ai_authoring_summary", {}),
            "ai_recommendations": [r.to_dict() if hasattr(r, "to_dict") else r for r in getattr(context, "ai_recommendations", [])],
            "ai_audit_trail": getattr(context, "ai_audit_trail", []),
            "task_3_report": getattr(context, "task_3_report", {}),
            "image_decisions": [d.to_dict() if hasattr(d, "to_dict") else d for d in getattr(context, "image_decisions", [])],
            "task_4_placeholder_report": getattr(context, "task_4_placeholder_report", None),
        }
