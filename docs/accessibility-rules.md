# SlideGuard Accessibility Rules Specification

## 1. Rule Classification Framework
Every rule in SlideGuard implements `IAccessibilityRule` and explicitly declares its behavior classification:

1. **`DETECT_ONLY`**:
   - Evaluates compliance and flags issues for author awareness.
   - Never modifies presentation XML automatically.
   - Examples: Text Contrast, Typography / Readability checks.

2. **`SAFE_AUTO_REMEDIATE`**:
   - Deterministic remediation with zero risk of visual or layout corruption.
   - Remediated only after user approval of the remediation plan.
   - Examples: Reading order sequence correction, native placeholder alignment, semantic list markup.

3. **`AI_ASSISTED_REVIEW`**:
   - Requires semantic judgment or natural language description.
   - May leverage optional AI models to generate suggestions, but **never** applies changes without explicit author review.
   - Examples: Missing Alternative Text, Decorative element classification.

## 2. Default Rules Matrix

| Rule ID | Rule Name | Behavior | WCAG Standard Mapping | Description |
|---|---|---|---|---|
| `SG-RO-01` | Reading Order Logical Sequence | `SAFE_AUTO_REMEDIATE` | WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A) | Validates logical top-down sequence and excludes footers (e.g. Wiley) from main reading flow. |
| `SG-GRP-01` | Selective Ungrouping for Accessible Text | `SAFE_AUTO_REMEDIATE` | WCAG 2.1 - 1.3.1 Info & Relationships (Level A) | Identifies groups with text/shapes mixed with images, targeting text for exposure while protecting images. |
| `SG-PH-01` | Native Placeholder Usage | `SAFE_AUTO_REMEDIATE` | WCAG 2.1 - 1.3.1 Info & Relationships (Level A) | Identifies standalone text boxes serving as slide headers instead of native layout placeholders. |
| `SG-LST-01` | Semantic List Formatting | `SAFE_AUTO_REMEDIATE` | WCAG 2.1 - 1.3.1 Info & Relationships (Level A) | Detects pseudo-lists typed manually with dashes or numbers lacking native XML list tags. |
| `SG-ALT-01` | Missing Alternative Text | `AI_ASSISTED_REVIEW` | WCAG 2.1 - 1.1.1 Non-text Content (Level A) | Flags non-decorative pictures and graphics lacking alternative text descriptions. |
| `SG-DEC-01` | Decorative Element Tagging | `AI_ASSISTED_REVIEW` | WCAG 2.1 - 1.1.1 Non-text Content (Level A) | Identifies purely visual divider lines or background shapes requiring decorative designation. |
| `SG-TTL-01` | Slide Title Presence | `SAFE_AUTO_REMEDIATE` | WCAG 2.1 - 2.4.2 Page Titled (Level A) | Validates that every slide possesses an accessible title for outline navigation. |
| `SG-TBL-01` | Table Structural Accessibility | `SAFE_AUTO_REMEDIATE` | WCAG 2.1 - 1.3.1 Info & Relationships (Level A) | Ensures data tables include designated header rows (`firstRow`). |
| `SG-CNT-01` | Color Contrast Ratio (Minimum) | `DETECT_ONLY` | WCAG 2.1 - 1.4.3 Contrast (Minimum) (Level AA) | Flags low-contrast foreground text colors against slide backgrounds. |
| `SG-RDB-01` | Typography & Legibility Check | `DETECT_ONLY` | Best Practice (Readability) | Identifies excessively small fonts (<12pt) and empty text boxes. |
| `SG-LNG-01` | Language of Document | `SAFE_AUTO_REMEDIATE` | WCAG 2.1 - 3.1.1 Language of Page (Level A) | Verifies presentation metadata specifies a document language. |
| `SG-META-01` | Document Metadata Information | `SAFE_AUTO_REMEDIATE` | WCAG 2.1 - 2.4.2 Page Titled (Level A) | Checks for document title and author in presentation properties. |

## 3. Auditable Quality Scoring
SlideGuard does not invent arbitrary scores. The scoring formula is transparent and configurable:

$$\text{Weighted Deductions} = (\text{Critical} \times 15) + (\text{Error} \times 8) + (\text{Warning} \times 2) + (\text{Info} \times 0)$$

$$\text{Normalized Deduction} = \min\left(100.0, \frac{\text{Weighted Deductions}}{\max(1, \text{Total Slides})}\right)$$

$$\text{Accessibility Score} = \max(0.0, 100.0 - \text{Normalized Deduction})$$
