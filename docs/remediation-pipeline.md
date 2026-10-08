# SlideGuard Remediation Pipeline Specification

## 1. Pipeline Principles
The SlideGuard remediation engine adheres to three strict requirements:
1. **User Approval First**: The engine never alters a presentation without explicit user review and approval of detected issues.
2. **Deterministic Execution**: Rules modify PowerPoint OpenXML only when the structural transformation is predictable and mathematically lossless.
3. **Strict Step Ordering**: Transformations occur in an invariant, linear pipeline.

## 2. Invariant Step Order

```
[ Input Presentation ]
          │
          ▼
STEP 1: Reading Order Alignment
  - Determines logical z-order / reading hierarchy
  - Explicitly isolates footer content (e.g., Wiley, slide numbers) from reading sequence
          │
          ▼
STEP 2: Selective Ungrouping
  - Traverses nested group hierarchies
  - Unbundles text boxes, rectangles, and auto-shapes for screen reader accessibility
  - Strictly preserves and protects embedded pictures and composite drawings
          │
          ▼
STEP 3: Native Placeholder Conversion
  - Converts loose text headers into native slide title/body placeholders
  - Preserves exact font family, size, color, coordinates, margins, and alignment
          │
          ▼
STEP 4: Semantic List Structure Preservation
  - Detects simulated lists and tags them using native PowerPoint list structures
  - Preserves original numbering styles: bullets, numeric (1, 2, 3), alphabetic (a, b, c)
          │
          ▼
STEP 5: Post-Remediation Accessibility Validation
  - Evaluates all accessibility rules against the remediated file
          │
          ▼
STEP 6: Originality & Integrity Verification
  - Performs semantic comparison between original and candidate presentation
  - Distinguishes intentional accessibility adjustments from unintended visual drift
  - Validates OOXML package integrity
          │
          ▼
[ Final Remediated Presentation & Audit Reports ]
```

## 3. Design Preservation & Originality Guarantee
The remediation pipeline guarantees visual design preservation:
- Fonts, font sizes, and font colors are preserved.
- Bounding box coordinates (`left`, `top`, `width`, `height`) are preserved within a 2pt tolerance.
- Slide background themes, master layouts, and geometries remain untouched.
- Objects are never deleted or flattened unless specifically designated as empty artifacts.
