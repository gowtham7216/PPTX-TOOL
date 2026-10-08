# SlideGuard Accessibility & Integrity Report

**Job ID:** `job_20261007_111737_8053d372`  
**File Name:** `Sec1-5-TrigFunctions (2)original.pptx`  
**Status:** `COMPLETED`  
**Execution Duration:** `46.41s`  

---

## 1. Executive Summary & Quality Metrics

| Metric | Value |
|---|---|
| **Accessibility Quality Score** | **98.3 / 100** |
| **Total Slides** | 19 |
| **Slides Analyzed** | 19 |
| **Total Accessibility Issues** | 13 |
| - Critical | 0 |
| - Error | 1 |
| - Warning | 12 |
| - Info | 0 |
| **Rules Evaluated** | 228 |
| - Rules Passed | 215 |
| - Rules Failed | 13 |
| - Rules Warning | 0 |
| **Design Preservation Verified** | ✅ Yes |

---

## 2. Reading Order Analysis & Proposal

| Slide # | Confidence | Original Order | Proposed Order | Footers Excluded | Reasoning |
|---|---|---|---|---|---|
| 3 | **LOW** | `[2, 3, 4, 5, 6, 7, 8, 9, 10, 11]` | `[2, 3, 6, 8, 4, 5, 9, 10, 7, 11]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Paired caption 'TextBox 7' immediately following image 'Picture 5'. |
| 5 | **HIGH** | `[2, 3, 4, 6, 7, 9, 11, 12]` | `[2, 4, 9, 6, 7, 12, 11, 3]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Reordering proposed to satisfy structural accessibility hierarchy. |
| 6 | **HIGH** | `[2, 3, 5, 7, 9]` | `[2, 5, 3, 9, 7]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Multi-column detected (TWO_COLUMN): reading column-by-column left to right. |
| 7 | **HIGH** | `[2, 4, 5, 7, 6, 8]` | `[2, 7, 8, 4, 5, 6]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Multi-column detected (TWO_COLUMN): reading column-by-column left to right. |
| 8 | **HIGH** | `[2, 4, 7, 6, 8]` | `[2, 7, 8, 4, 6]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Multi-column detected (TWO_COLUMN): reading column-by-column left to right. |
| 9 | **LOW** | `[2, 3, 4, 5, 6, 8, 10, 11, 15, 16, 18, 20]` | `[2, 5, 3, 6, 4, 15, 10, 16, 8, 18, 11, 20]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Paired caption 'TextBox 5' immediately following image 'Picture 1'. |
| 10 | **HIGH** | `[2, 7, 5, 3, 6]` | `[2, 6, 7, 5, 3]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Multi-column detected (TWO_COLUMN): reading column-by-column left to right. |
| 11 | **MEDIUM** | `[2, 3, 5, 7, 9, 14, 15, 17, 19, 6]` | `[2, 15, 7, 9, 14, 5, 17, 6, 3, 19]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Multi-column detected (TWO_COLUMN): reading column-by-column left to right. |
| 12 | **LOW** | `[2, 4, 9, 5, 7, 10, 15, 3, 6]` | `[2, 4, 5, 9, 10, 6, 7, 15, 3]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Reordering proposed to satisfy structural accessibility hierarchy. |
| 13 | **LOW** | `[2, 4, 11, 7, 9, 10, 15, 5, 6]` | `[2, 4, 9, 11, 10, 6, 7, 15, 5]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Reordering proposed to satisfy structural accessibility hierarchy. |
| 15 | **HIGH** | `[2, 4, 7, 6]` | `[2, 6, 4, 7]` | None | Header: placed 1 title/subtitle elements at start of sequence.; Multi-column detected (TWO_COLUMN): reading column-by-column left to right. |

---

## 3. Selective Ungrouping & Hierarchy Analysis

No grouped objects requiring ungrouping detected in this presentation.

---

## 4. Placeholder Conversion & Analysis

All slide placeholders are structurally compliant; no conversions required.

---

## 5. Detected Accessibility Issues

| Slide # | Severity | Rule ID & Name | WCAG Criterion | Description | Recommendation |
|---|---|---|---|---|---|
| 1 | **WARNING** | `SG-LNG-01` Language of Document | WCAG 2.1 - 3.1.1 Language of Page (Level A) | Presentation document metadata does not declare a default language. | Set default presentation language (e.g. 'en-US') in document properties. |
| 3 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 3 reading sequence does not match logical hierarchy. Current: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11] -> Proposed: [2, 3, 6, 8, 4, 5, 9, 10, 7, 11]. | Reorder top-level shapes into proposed logical reading sequence. |
| 4 | **ERROR** | `SG-TTL-01` Slide Title Presence | WCAG 2.1 - 2.4.2 Page Titled (Level A) | Slide 4 has no defined title. | Add a descriptive title to the slide title placeholder. |
| 5 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 5 reading sequence does not match logical hierarchy. Current: [2, 3, 4, 6, 7, 9, 11, 12] -> Proposed: [2, 4, 9, 6, 7, 12, 11, 3]. | Reorder top-level shapes into proposed logical reading sequence. |
| 6 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 6 reading sequence does not match logical hierarchy. Current: [2, 3, 5, 7, 9] -> Proposed: [2, 5, 3, 9, 7]. | Reorder top-level shapes into proposed logical reading sequence. |
| 7 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 7 reading sequence does not match logical hierarchy. Current: [2, 4, 5, 7, 6, 8] -> Proposed: [2, 7, 8, 4, 5, 6]. | Reorder top-level shapes into proposed logical reading sequence. |
| 8 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 8 reading sequence does not match logical hierarchy. Current: [2, 4, 7, 6, 8] -> Proposed: [2, 7, 8, 4, 6]. | Reorder top-level shapes into proposed logical reading sequence. |
| 9 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 9 reading sequence does not match logical hierarchy. Current: [2, 3, 4, 5, 6, 8, 10, 11, 15, 16, 18, 20] -> Proposed: [2, 5, 3, 6, 4, 15, 10, 16, 8, 18, 11, 20]. | Reorder top-level shapes into proposed logical reading sequence. |
| 10 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 10 reading sequence does not match logical hierarchy. Current: [2, 7, 5, 3, 6] -> Proposed: [2, 6, 7, 5, 3]. | Reorder top-level shapes into proposed logical reading sequence. |
| 11 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 11 reading sequence does not match logical hierarchy. Current: [2, 3, 5, 7, 9, 14, 15, 17, 19, 6] -> Proposed: [2, 15, 7, 9, 14, 5, 17, 6, 3, 19]. | Reorder top-level shapes into proposed logical reading sequence. |
| 12 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 12 reading sequence does not match logical hierarchy. Current: [2, 4, 9, 5, 7, 10, 15, 3, 6] -> Proposed: [2, 4, 5, 9, 10, 6, 7, 15, 3]. | Reorder top-level shapes into proposed logical reading sequence. |
| 13 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 13 reading sequence does not match logical hierarchy. Current: [2, 4, 11, 7, 9, 10, 15, 5, 6] -> Proposed: [2, 4, 9, 11, 10, 6, 7, 15, 5]. | Reorder top-level shapes into proposed logical reading sequence. |
| 15 | **WARNING** | `SG-RO-01` Reading Order Logical Sequence | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Slide 15 reading sequence does not match logical hierarchy. Current: [2, 4, 7, 6] -> Proposed: [2, 6, 4, 7]. | Reorder top-level shapes into proposed logical reading sequence. |

---

## 6. Remediation Pipeline Log

| Step ID | Target Slide | Status | Action Description | Confidence |
|---|---|---|---|---|
| `PIPELINE_STEP_1_READING_ORDER` | 0 | `APPLIED` | STEP 1: Reading Order Alignment. Evaluated and processed 19 slides in-place. | 95% |
| `PIPELINE_STEP_2_UNGROUPING` | 0 | `SKIPPED` | STEP 2: Selective Ungrouping. Evaluated and processed 19 slides in-place. | 95% |
| `PIPELINE_STEP_3_PLACEHOLDERS` | 0 | `SKIPPED` | STEP 3: Native Placeholder Conversion. Evaluated and processed 19 slides in-place. | 95% |
| `PIPELINE_STEP_4` | 0 | `NOT_IMPLEMENTED` | List Structure Preservation scheduled for Phase 6. | 100% |
| `PIPELINE_STEP_5` | 0 | `NOT_IMPLEMENTED` | Accessibility Validation scheduled for Phase 7. | 100% |
| `PIPELINE_STEP_6` | 0 | `NOT_IMPLEMENTED` | Integrity Verification scheduled for Phase 8. | 100% |

---

## 7. Verification Check Details

**Overall Verification:** ✅ PASSED  
Passed: 4 / 4 checks  

| Check Name | Slide | Status | Details |
|---|---|---|---|
| File Size | 0 | ✅ Pass | File size valid: 1397049 bytes |
| OOXML Container Validity | 0 | ✅ Pass | Successfully opened presentation XML package |
| Slide Count Preservation | 0 | ✅ Pass | Slide count verified: 19 slides |
| Slide Shapes Well-Formed | 0 | ✅ Pass | All slide shape trees loaded successfully |
