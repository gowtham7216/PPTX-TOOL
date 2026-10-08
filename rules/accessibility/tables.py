"""
SlideGuard Table Accessibility Rule (Phase 7).
Validates table headers, empty cells, and merged cell structures.
Safely marks table header row (firstRow="1") when row 0 contains valid header data.
"""

from typing import List
from lxml import etree

from pptx_engine.ooxml.xml_utils import find_elements_by_xpath
from rules.base import BaseAccessibilityRule
from rules.accessibility.base import get_cnvpr_element, safe_xpath
from models.enums import RuleBehavior, IssueSeverity, ShapeType, AccessibilityStatus, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.remediation import RemediationAction


class TableAccessibilityRule(BaseAccessibilityRule):
    """
    Validates that PowerPoint tables have designated header rows
    and do not contain blank or unidentifiable header structures.
    """

    @property
    def rule_id(self) -> str:
        return "SG-TBL-01"

    @property
    def name(self) -> str:
        return "Table Structural Accessibility"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.SAFE_AUTO_REMEDIATE

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.3.1 Info and Relationships (Level A)"

    @property
    def description(self) -> str:
        return "Ensures data tables include accessible header rows and clear structural markup."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        for shape in slide.shapes:
            if shape.shape_type == ShapeType.TABLE or (shape.raw_element is not None and bool(safe_xpath(shape.raw_element, ".//a:tbl"))):
                if shape.raw_element is None:
                    continue

                tbl_elements = safe_xpath(shape.raw_element, ".//a:tbl")
                if not tbl_elements:
                    continue
                tbl = tbl_elements[0]

                tbl_prs = safe_xpath(tbl, "./a:tblPr")
                has_first_row = False
                if tbl_prs and tbl_prs[0].get("firstRow") in ("1", "true"):
                    has_first_row = True

                # Inspect rows and cells
                rows = safe_xpath(tbl, "./a:tr")
                has_empty_cells = False
                has_merged_cells = False
                row_0_has_text = True

                if rows:
                    row_0_cells = safe_xpath(rows[0], "./a:tc")
                    for tc in row_0_cells:
                        tc_text = "".join(safe_xpath(tc, ".//a:t/text()")).strip()
                        if not tc_text:
                            row_0_has_text = False
                        # Check merged cells
                        if tc.get("gridSpan") or tc.get("rowSpan") or tc.get("hMerge") or tc.get("vMerge"):
                            has_merged_cells = True

                if not has_first_row:
                    if row_0_has_text:
                        issues.append(
                            Issue(
                                issue_id=f"TBL_NO_HEADER_{slide.slide_number}_{shape.shape_id}",
                                rule_id=self.rule_id,
                                rule_name=self.name,
                                slide_number=slide.slide_number,
                                severity=IssueSeverity.WARNING,
                                shape_id=shape.shape_id,
                                shape_name=shape.name,
                                wcag_mapping=self.wcag_mapping,
                                description=f"Table '{shape.name}' does not have a designated header row ('firstRow').",
                                recommendation="Enable header row markup in table design properties.",
                                behavior=self.behavior,
                                status=AccessibilityStatus.SAFE_TO_REMEDIATE,
                                confidence="HIGH",
                                explanation="Row 0 contains non-empty text headings but tblPr lacks firstRow='1' attribute.",
                                extra_data={"table_shape_id": shape.shape_id, "can_set_first_row": True}
                            )
                        )
                    else:
                        issues.append(
                            Issue(
                                issue_id=f"TBL_AMBIGUOUS_HEADER_{slide.slide_number}_{shape.shape_id}",
                                rule_id=self.rule_id,
                                rule_name=self.name,
                                slide_number=slide.slide_number,
                                severity=IssueSeverity.WARNING,
                                shape_id=shape.shape_id,
                                shape_name=shape.name,
                                wcag_mapping=self.wcag_mapping,
                                description=f"Table '{shape.name}' header row is incomplete or ambiguous.",
                                recommendation="Provide descriptive column headers for each column in row 0.",
                                behavior=RuleBehavior.DETECT_ONLY,
                                status=AccessibilityStatus.MANUAL_REVIEW,
                                confidence="HIGH",
                                explanation="One or more cells in the first row contain empty text.",
                                extra_data={"table_shape_id": shape.shape_id, "can_set_first_row": False}
                            )
                        )

                if has_merged_cells:
                    issues.append(
                        Issue(
                            issue_id=f"TBL_MERGED_CELLS_{slide.slide_number}_{shape.shape_id}",
                            rule_id=self.rule_id,
                            rule_name=self.name,
                            slide_number=slide.slide_number,
                            severity=IssueSeverity.INFO,
                            shape_id=shape.shape_id,
                            shape_name=shape.name,
                            wcag_mapping=self.wcag_mapping,
                            description=f"Table '{shape.name}' contains merged cells that may complicate screen reader navigation.",
                            recommendation="Consider simplifying table structure into simple rectangular grid where possible.",
                            behavior=RuleBehavior.DETECT_ONLY,
                            status=AccessibilityStatus.MANUAL_REVIEW,
                            confidence="HIGH",
                            explanation="Table definition utilizes gridSpan or rowSpan attributes.",
                            extra_data={"table_shape_id": shape.shape_id}
                        )
                    )

        return issues

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """Deterministically enables firstRow='1' on <a:tblPr> in-place."""
        extra = issue.extra_data or {}
        target_id = issue.shape_id or extra.get("table_shape_id")

        if not extra.get("can_set_first_row", False):
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                target_shape_id=target_id,
                description="Table header is ambiguous; preserving original structure for manual review.",
                status=RemediationStatus.SKIPPED,
                confidence=1.0,
                is_user_approved=True
            )

        if slide.raw_slide is None or target_id is None:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                target_shape_id=target_id,
                description="Unable to access raw slide element for table remediation.",
                status=RemediationStatus.FAILED,
                confidence=0.0
            )

        try:
            ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
            sp_tree = slide.raw_slide.shapes._spTree
            applied = False

            for child in sp_tree:
                cnvpr = get_cnvpr_element(child)
                if cnvpr is not None and cnvpr.get("id") == str(target_id):
                    tbl_prs = safe_xpath(child, ".//a:tbl/a:tblPr")
                    if tbl_prs:
                        tbl_prs[0].set("firstRow", "1")
                        applied = True
                    else:
                        tbl_nodes = safe_xpath(child, ".//a:tbl")
                        if tbl_nodes:
                            # Prepend a:tblPr
                            tbl_pr = etree.Element(f"{{{ns_a}}}tblPr", attrib={"firstRow": "1"})
                            tbl_nodes[0].insert(0, tbl_pr)
                            applied = True
                    break

            if applied:
                return RemediationAction(
                    action_id=f"rem_{issue.issue_id}",
                    rule_id=self.rule_id,
                    slide_number=slide.slide_number,
                    target_shape_id=target_id,
                    description=f"Enabled header row markup (firstRow='1') on table {target_id} in-place.",
                    status=RemediationStatus.APPLIED,
                    confidence=0.98,
                    is_user_approved=True
                )
            else:
                return RemediationAction(
                    action_id=f"rem_{issue.issue_id}",
                    rule_id=self.rule_id,
                    slide_number=slide.slide_number,
                    target_shape_id=target_id,
                    description=f"Table shape {target_id} not found in slide tree.",
                    status=RemediationStatus.FAILED,
                    confidence=0.0
                )
        except Exception as e:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                target_shape_id=target_id,
                description=f"Table remediation failed: {str(e)}",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message=str(e)
            )
