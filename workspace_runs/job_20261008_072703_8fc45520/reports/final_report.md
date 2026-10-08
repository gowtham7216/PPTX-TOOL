# SlideGuard Accessibility & Integrity Report

**Job ID:** `job_20261008_072703_8fc45520`  
**File Name:** `test_cli.pptx`  
**Status:** `COMPLETED`  
**Execution Duration:** `0.33s`  

---

## 1. Executive Summary & Quality Metrics

| Metric | Value |
|---|---|
| **Accessibility Quality Score** | **93.0 / 100** |
| **Total Slides** | 6 |
| **Slides Analyzed** | 6 |
| **Total Accessibility Issues** | 17 |
| - Critical | 0 |
| - Error | 2 |
| - Warning | 13 |
| - Info | 2 |
| **Rules Evaluated** | 84 |
| - Rules Passed | 69 |
| - Rules Failed | 15 |
| - Rules Warning | 0 |
| **Design Preservation Verified** | ✅ Yes |

---

## 2. Reading Order Analysis & Proposal

All slide reading sequences are logically compliant; no reordering required.

---

## 3. Selective Ungrouping & Hierarchy Analysis

No grouped objects requiring ungrouping detected in this presentation.

---

## 4. Placeholder Conversion & Analysis

| Slide # | Shape ID | Name | Category | Status | Inherited From | Proposed Action / Reason |
|---|---|---|---|---|---|---|
| 2 | `2` | TextBox 1 | `STANDALONE_TITLE_BOX` | 🛡️ Preserved | SLIDE | Slide lacks native title; text box serves as visual title and should be linked for outline navigation. |
| 3 | `2` | TextBox 1 | `STANDALONE_TITLE_BOX` | 🛡️ Preserved | SLIDE | Slide lacks native title; text box serves as visual title and should be linked for outline navigation. |
| 4 | `2` | TextBox 1 | `STANDALONE_TITLE_BOX` | 🛡️ Preserved | SLIDE | Slide lacks native title; text box serves as visual title and should be linked for outline navigation. |
| 6 | `2` | TextBox 1 | `STANDALONE_TITLE_BOX` | 🛡️ Preserved | SLIDE | Slide lacks native title; text box serves as visual title and should be linked for outline navigation. |

---

## 5. List Semantics & Numbering

| Slide # | Shape ID | Paragraph | Category | Status | Confidence | Proposed Action / Reason |
|---|---|---|---|---|---|---|
| 2 | `3` | P2 | `BULLET` | ✅ Safe to Remediate | HIGH | Promote to native list (BULLET) with prefix removed from text. |
| 2 | `3` | P2 | `BULLET` | ✅ Safe to Remediate | HIGH | Promote to native list (BULLET) with prefix removed from text. |

---

## 6. Accessibility Validation & Remediation

| Metric | Count |
|---|---|
| **Total Rules Evaluated** | 84 |
| **Rules Passed** | 69 |
| **Errors** | 2 |
| **Warnings** | 13 |
| **Info** | 2 |
| **Safe Auto-Remediations** | 4 |
| **Manual-Review Items** | 6 |
| **Detect-Only Issues** | 6 |
| **Unsupported Cases** | 0 |
| **Remediation Failures** | 0 |

| Slide | Shape | Rule ID & Name | Severity | WCAG Criterion | Description | Recommendation | Action / Status | Confidence | Verification |
|---|---|---|---|---|---|---|---|---|---|
| 1 | Slide | `SG-LNG-01` Language of Document | **WARNING** | WCAG 2.1 - 3.1.1 Language of Page (Level A) | Presentation document metadata does not declare a default language. | Set default presentation language (e.g. 'en-US') in document properties. | `MANUAL_REVIEW` | HIGH | PENDING |
| 2 | `2` | `SG-PH-01` Native Placeholder Usage | **WARNING** | WCAG 2.1 - 1.3.1 Info and Relationships (Level A) | Slide 2: Standalone text box 'Key Objectives' appears to be a slide title but is not a native slide title placeholder. | Convert standalone title text box into a native slide title placeholder. | `DETECTED` | HIGH | PENDING |
| 2 | `3` | `SG-LST-01` Semantic List Formatting | **WARNING** | WCAG 2.1 - 1.3.1 Info and Relationships (Level A) | Slide 2 Shape 3 ('TextBox 2'): Paragraph 1 contains simulated list / pseudo-list prefix '- ' without native PowerPoint list tags. | Promote to native list (BULLET) with prefix removed from text. | `DETECTED` | HIGH | PENDING |
| 2 | `3` | `SG-LST-01` Semantic List Formatting | **WARNING** | WCAG 2.1 - 1.3.1 Info and Relationships (Level A) | Slide 2 Shape 3 ('TextBox 2'): Paragraph 2 contains simulated list / pseudo-list prefix '- ' without native PowerPoint list tags. | Promote to native list (BULLET) with prefix removed from text. | `DETECTED` | HIGH | PENDING |
| 2 | `2` | `SG-TTL-01` Slide Title Presence | **WARNING** | WCAG 2.1 - 2.4.2 Page Titled (Level A) | Slide 2 has title-like text box 'Key Objectives' lacking native title placeholder semantics. | Promote text box to native title placeholder while preserving layout and geometry. | `SAFE_TO_REMEDIATE` | HIGH | PENDING |
| 3 | `2` | `SG-PH-01` Native Placeholder Usage | **WARNING** | WCAG 2.1 - 1.3.1 Info and Relationships (Level A) | Slide 3: Standalone text box 'Slide with Footer Content' appears to be a slide title but is not a native slide title placeholder. | Convert standalone title text box into a native slide title placeholder. | `DETECTED` | HIGH | PENDING |
| 3 | `2` | `SG-TTL-01` Slide Title Presence | **WARNING** | WCAG 2.1 - 2.4.2 Page Titled (Level A) | Slide 3 has title-like text box 'Slide with Footer Content' lacking native title placeholder semantics. | Promote text box to native title placeholder while preserving layout and geometry. | `SAFE_TO_REMEDIATE` | HIGH | PENDING |
| 4 | `2` | `SG-PH-01` Native Placeholder Usage | **WARNING** | WCAG 2.1 - 1.3.1 Info and Relationships (Level A) | Slide 4: Standalone text box 'Compliance Metrics Table' appears to be a slide title but is not a native slide title placeholder. | Convert standalone title text box into a native slide title placeholder. | `DETECTED` | HIGH | PENDING |
| 4 | `3` | `SG-ALT-01` Missing Alternative Text | **ERROR** | WCAG 2.1 - 1.1.1 Non-text Content (Level A) | Picture 'Table 2' lacks alternative text description. | Provide descriptive alternative text or mark as decorative if purely visual. | `DETECTED` | HIGH | PENDING |
| 4 | `2` | `SG-TTL-01` Slide Title Presence | **WARNING** | WCAG 2.1 - 2.4.2 Page Titled (Level A) | Slide 4 has title-like text box 'Compliance Metrics Table' lacking native title placeholder semantics. | Promote text box to native title placeholder while preserving layout and geometry. | `SAFE_TO_REMEDIATE` | HIGH | PENDING |
| 5 | Slide | `SG-TTL-01` Slide Title Presence | **ERROR** | WCAG 2.1 - 2.4.2 Page Titled (Level A) | Slide 5 has no defined title. | Add a descriptive title to the slide title placeholder. | `MANUAL_REVIEW` | HIGH | PENDING |
| 5 | `3` | `SG-CNT-01` Color Contrast Ratio (Minimum) | **WARNING** | WCAG 2.1 - 1.4.3 Contrast (Minimum) (Level AA) | Text run 'Hard to read low con...' uses low-contrast color #CCCCCC. | Increase color contrast ratio to at least 4.5:1 (calculated: 1.61:1). | `MANUAL_REVIEW` | HIGH | PENDING |
| 5 | `2` | `SG-RDB-01` Typography & Legibility Check | **INFO** | Accessibility Best Practice (Readability) | Empty text box 'TextBox 1' detected on slide 5. | Remove empty shape to prevent confusing screen reader announcement. | `MANUAL_REVIEW` | HIGH | PENDING |
| 5 | `3` | `SG-RDB-01` Typography & Legibility Check | **WARNING** | Accessibility Best Practice (Readability) | Text run 'Hard to read low con...' uses small font size (8.0pt). | Increase font size to at least 14pt (or 12pt minimum) for legibility. | `MANUAL_REVIEW` | HIGH | PENDING |
| 5 | `2` | `SG-TXT-01` Missing or Empty Text Content | **INFO** | WCAG 2.1 - 1.3.1 Info and Relationships (Level A) | Shape 'TextBox 1' contains empty or whitespace-only content. | Populate with meaningful content or remove shape if redundant. | `MANUAL_REVIEW` | HIGH | PENDING |
| 6 | `2` | `SG-PH-01` Native Placeholder Usage | **WARNING** | WCAG 2.1 - 1.3.1 Info and Relationships (Level A) | Slide 6: Standalone text box 'Grouped Elements Slide' appears to be a slide title but is not a native slide title placeholder. | Convert standalone title text box into a native slide title placeholder. | `DETECTED` | HIGH | PENDING |
| 6 | `2` | `SG-TTL-01` Slide Title Presence | **WARNING** | WCAG 2.1 - 2.4.2 Page Titled (Level A) | Slide 6 has title-like text box 'Grouped Elements Slide' lacking native title placeholder semantics. | Promote text box to native title placeholder while preserving layout and geometry. | `SAFE_TO_REMEDIATE` | HIGH | PENDING |

---

## 7. Accessibility Intelligence Summary

| Intelligence Metric | Value |
|---|---|
| **Total Raw Issues** | 17 |
| **Correlated Issues** | 10 |
| **Deduplicated Issues** | 7 |
| **High-Confidence Issues** | 4 |
| **Medium-Confidence Issues** | 4 |
| **Low-Confidence Issues** | 2 |
| **Safe Remediation Candidates** | 4 |
| **Remediated Issues** | 0 |
| **Unresolved Issues** | 6 |
| **New Issues** | 0 |
| **Manual-Review Issues** | 4 |
| **Unsupported Issues** | 0 |
| **Quality Score (Before Remediation)** | 92.4 / 100 |
| **Quality Score (After Remediation)** | 96.1 / 100 |

### Correlated Unified Issues

| Slide | Shape | Primary Rule | Related Rules | Severity | Confidence | Explanation | Recommended Action | Final Status | Verification |
|---|---|---|---|---|---|---|---|---|---|
| 1 | Slide | `SG-LNG-01` | None | **WARNING** | **LOW** | No language code declared in presentation metadata or default text styles. | Set default presentation language (e.g. 'en-US') in document properties. | `MANUAL_REVIEW` | PENDING |
| 2 | `2` | `SG-TTL-01` | `SG-PH-01` | **WARNING** | **HIGH** | Text box is near slide top (top=36.0pt) with concise heading text (14 chars). Corroborated by related checks: Native Placeholder Usage. | Promote text box to native title placeholder while preserving layout and geometry. | `RESOLVED` | PENDING |
| 2 | `3` | `SG-LST-01` | None | **WARNING** | **MEDIUM** | Slide 2 Shape 3 ('TextBox 2'): Paragraph 1 contains simulated list / pseudo-list prefix '- ' without native PowerPoint list tags. | Promote to native list (BULLET) with prefix removed from text. | `CHANGED` | PENDING |
| 3 | `2` | `SG-TTL-01` | `SG-PH-01` | **WARNING** | **HIGH** | Text box is near slide top (top=36.0pt) with concise heading text (25 chars). Corroborated by related checks: Native Placeholder Usage. | Promote text box to native title placeholder while preserving layout and geometry. | `RESOLVED` | PENDING |
| 4 | `2` | `SG-TTL-01` | `SG-PH-01` | **WARNING** | **HIGH** | Text box is near slide top (top=36.0pt) with concise heading text (24 chars). Corroborated by related checks: Native Placeholder Usage. | Promote text box to native title placeholder while preserving layout and geometry. | `RESOLVED` | PENDING |
| 4 | `3` | `SG-ALT-01` | None | **ERROR** | **MEDIUM** | Visual image element has no description or title attribute in OOXML. | Provide descriptive alternative text or mark as decorative if purely visual. | `UNCHANGED` | PENDING |
| 5 | Slide | `SG-TTL-01` | None | **ERROR** | **LOW** | No native title placeholder or prominent heading text box detected on slide. | Add a descriptive title to the slide title placeholder. | `MANUAL_REVIEW` | PENDING |
| 5 | `2` | `SG-RDB-01` | `SG-TXT-01` | **INFO** | **MEDIUM** | Text box contains no text runs or content. Corroborated by related checks: Missing or Empty Text Content. | Remove empty shape to prevent confusing screen reader announcement. | `MANUAL_REVIEW` | PENDING |
| 5 | `3` | `SG-CNT-01` | `SG-RDB-01` | **WARNING** | **MEDIUM** | Foreground #CCCCCC against background #FFFFFF yields ratio 1.61:1 (required: 4.5:1). Corroborated by related checks: Typography & Legibility Check. | Increase color contrast ratio to at least 4.5:1 (calculated: 1.61:1). | `MANUAL_REVIEW` | PENDING |
| 6 | `2` | `SG-TTL-01` | `SG-PH-01` | **WARNING** | **HIGH** | Text box is near slide top (top=36.0pt) with concise heading text (22 chars). Corroborated by related checks: Native Placeholder Usage. | Promote text box to native title placeholder while preserving layout and geometry. | `RESOLVED` | PENDING |

---

## 8. Remediation Pipeline Log

| Step ID | Target Slide | Status | Action Description | Confidence |
|---|---|---|---|---|
| `PIPELINE_STEP_1_READING_ORDER` | 0 | `SKIPPED` | STEP 1: Reading Order Alignment. Evaluated and processed 6 slides in-place. | 95% |
| `PIPELINE_STEP_2_UNGROUPING` | 0 | `SKIPPED` | STEP 2: Selective Ungrouping. Evaluated and processed 6 slides in-place. | 95% |
| `PIPELINE_STEP_3_PLACEHOLDERS` | 0 | `APPLIED` | STEP 3: Native Placeholder Conversion. Evaluated and processed 6 slides in-place. | 95% |
| `PIPELINE_STEP_4_LISTS` | 0 | `APPLIED` | STEP 4: List Semantics & Numbering. Remediated 1 list items in-place. | 95% |
| `PIPELINE_STEP_5_ACCESSIBILITY` | 0 | `APPLIED` | STEP 5: Accessibility Validation & Safe Remediation. Safely remediated 4 issues in-place. | 95% |
| `PIPELINE_STEP_6_RECONCILIATION` | 0 | `APPLIED` | STEP 6: Final Validation & Issue Reconciliation. Reconciled 10 pre-remediation issues: 4 resolved, 0 remediated, 1 unchanged, 0 new. | 100% |

---

## 9. Verification Check Details

**Overall Verification:** ✅ PASSED  
Passed: 4 / 4 checks  

| Check Name | Slide | Status | Details |
|---|---|---|---|
| File Size | 0 | ✅ Pass | File size valid: 33550 bytes |
| OOXML Container Validity | 0 | ✅ Pass | Successfully opened presentation XML package |
| Slide Count Preservation | 0 | ✅ Pass | Slide count verified: 6 slides |
| Slide Shapes Well-Formed | 0 | ✅ Pass | All slide shape trees loaded successfully |

---

## 10. AI-Assisted Accessibility Authoring & Audit Trail

> **Notice:** AI-generated authoring proposals (powered strictly by Gemini 3.8 Flash High) require explicit human approval before deterministic OOXML application. AI assistance does not certify WCAG compliance.

### Authoring Metrics Summary

| Metric | Value |
|---|---|
| **Recommendations Generated** | 6 |
| **Recommendations Approved** | 0 |
| **Recommendations Edited** | 0 |
| **Recommendations Rejected** | 0 |
| **Recommendations Applied** | 0 |
| **Recommendations Failed** | 0 |
| **Manual Reviews** | 0 |
| **AI Fallback Cases** | 6 |
| **Malformed AI Responses** | 0 |

### Structured AI Authoring Recommendations

| Rec ID | Slide | Shape | Authoring Type | Original Value | AI Suggestion | Confidence | Grounding | Decision Status |
|---|---|---|---|---|---|---|---|---|
| `rec_alt_4_3` | 4 | `3` | `ALT_TEXT` | [Missing] | Graphic: Table 2 | 50% | `INFERENCE` | **PENDING_REVIEW** |
| `rec_ttl_2` | 2 | `2` | `SLIDE_TITLE` | [Missing Title] | Key Objectives | 60% | `SUPPORTED_FACT` | **PENDING_REVIEW** |
| `rec_ttl_3` | 3 | `2` | `SLIDE_TITLE` | [Missing Title] | Slide with Footer Content | 60% | `SUPPORTED_FACT` | **PENDING_REVIEW** |
| `rec_ttl_4` | 4 | `2` | `SLIDE_TITLE` | [Missing Title] | Compliance Metrics Table | 60% | `SUPPORTED_FACT` | **PENDING_REVIEW** |
| `rec_ttl_5` | 5 | `None` | `SLIDE_TITLE` | [Missing Title] | Hard to read low contrast text | 60% | `SUPPORTED_FACT` | **PENDING_REVIEW** |
| `rec_ttl_6` | 6 | `2` | `SLIDE_TITLE` | [Missing Title] | Grouped Elements Slide | 60% | `SUPPORTED_FACT` | **PENDING_REVIEW** |
