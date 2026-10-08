# SlideGuard

SlideGuard is an offline, local Windows desktop application for analyzing, remediating, validating, and comparing PowerPoint (`.pptx`) presentations for accessibility compliance while strictly preserving visual design integrity.

---

## 1. Project Purpose
SlideGuard enables organizations to ensure their PowerPoint files adhere to digital accessibility standards (such as WCAG 2.1) without risking visual distortion, layout breakage, or unintended modifications. 

### Core Guarantees:
- **Zero Cloud Dependence**: Operates 100% locally on Windows without cloud APIs or databases.
- **Source Protection**: User source files are never overwritten; operations occur within an isolated job workspace.
- **User Approval Gate**: Issues are presented for audit first. Remediation is only performed after explicit user approval.
- **Design Preservation**: Non-accessibility elements (positions, typography, brand colors, images) are preserved.
- **Deterministic First**: Built on deterministic rules and OpenXML transformations.

---

## 2. Layered Architecture

```
SlideGuard/
│
├── app/                  # Desktop application entry & GUI
│   ├── main.py           # Application entry point
│   ├── controller/       # Application controller (decouples UI from engine)
│   └── ui/               # Tkinter/TTK Desktop User Interface
│
├── engine/               # Processing engines
│   ├── analyzer/         # Presentation analyzer with slide-level failure isolation
│   ├── remediation/      # Invariant 6-step remediation pipeline
│   ├── verification/     # Post-remediation OOXML & package verification
│   ├── comparison/       # Semantic comparison & design preservation
│   ├── accessibility/    # Quality scoring engine
│   └── common/           # Core engine interfaces & contracts
│
├── pptx_engine/          # PowerPoint & OOXML processing
│   ├── reader/           # PresentationReader (python-pptx + lxml)
│   ├── writer/           # PresentationWriter
│   ├── ooxml/            # Low-level XML namespaces & node utilities
│   └── utilities/        # Spatial geometry & footer detection
│
├── rules/                # Modular accessibility rules
│   ├── base.py           # BaseAccessibilityRule abstract contract
│   ├── reading_order/    # Reading order hierarchy & footer exclusion
│   ├── ungrouping/       # Selective ungrouping (protects pictures)
│   ├── placeholders/     # Native placeholder conversion
│   ├── lists/            # List semantics (bullets, numbers, alphabetical)
│   ├── alt_text/         # Meaningful image alt text check
│   ├── decorative/       # Decorative element tagging
│   ├── titles/           # Slide title presence
│   ├── tables/           # Table structural headers
│   ├── contrast/         # Text contrast (DETECT_ONLY)
│   ├── readability/      # Typography & legibility (DETECT_ONLY)
│   ├── language/         # Document language check
│   └── metadata/         # Core properties check
│
├── models/               # Strongly-typed internal data models
├── reports/              # Audit report generator (Markdown & JSON)
├── workspace/            # Sandboxed workspace manager
├── tests/                # Unit, integration, regression test suite
└── docs/                 # Architectural & pipeline specifications
```

---

## 3. Setup & Installation

### Prerequisites
- Windows 10 / 11
- Python 3.10 or higher

### Install Dependencies
```powershell
python -m pip install -r requirements.txt
```

---

## 4. Running the Desktop Application

Launch the desktop UI:
```powershell
python app/main.py
```

To run in automated GUI launch verification mode:
```powershell
python app/main.py --test-mode
```

---

## 5. Running Tests

SlideGuard includes automated unit, integration, and regression tests:

```powershell
# Run full test suite
pytest

# Run tests with verbose output
pytest -v

# Run unit tests only
pytest tests/unit/

# Run integration tests only
pytest tests/integration/

# Run regression tests only
pytest tests/regression/
```

---

## 6. Packaging as Windows .EXE (PyInstaller)

SlideGuard is designed for eventual standalone packaging using PyInstaller without heavy external runtime dependencies:

```powershell
pyinstaller --noconfirm --onedir --windowed --name "SlideGuard" `
  --add-data "rules;rules" `
  --add-data "docs;docs" `
  app/main.py
```
*(Packaging will be finalized in Phase 13 after the remediation engine stages are fully implemented).*

---

## 7. Rule Architecture & Execution Modes

Every rule in SlideGuard implements `IAccessibilityRule` and declares one of three execution behaviors:
1. `DETECT_ONLY`: Flags issues for author awareness without altering content (e.g. contrast, font sizing).
2. `SAFE_AUTO_REMEDIATE`: Deterministic remediation with zero layout risk (applied only after user approval).
3. `AI_ASSISTED_REVIEW`: Generates advisory suggestions (e.g., draft alt-text) requiring explicit user confirmation.

---

## 8. AI Architecture
- AI is strictly optional and isolated behind `IAIProvider`.
- The application functions completely without an API key or internet access.
- AI is never permitted to write directly to PowerPoint files.
- All final file modifications are executed by the deterministic remediation engine.

---

## 9. Current Implementation Status & Roadmap

| Phase | Milestone | Status |
|---|---|---|
| **Phase 1** | **Scaffolding, Models, Workspace, Reader, Rules, UI, Tests** | **COMPLETED** |
| **Phase 2** | **PPTX Loading & OOXML Deep Inspection** | **COMPLETED** |
| **Phase 3** | **Reading Order Alignment & In-Place Reordering** | **COMPLETED** |
| **Phase 4** | **Selective Ungrouping Engine (Text Extraction & Picture Protection)** | **COMPLETED** |
| **Phase 5** | **Native Placeholder Conversion Engine** | **COMPLETED** |
| **Phase 6** | **List Semantics & Numbering Preservation** | **COMPLETED** |
| Phase 7 | Additional Accessibility Rules Expansion | Planned |
| Phase 8 | Remediation Approval Workflow Expansion | Planned |
| Phase 9 | Originality/Integrity Comparison & Visual Diffing | Planned |
| Phase 10 | Reporting & Audit Exports Enhancement | Planned |
| Phase 11 | Comprehensive Regression Suite | Planned |
| Phase 12 | PyInstaller Executable Packaging | Planned |

---

## 10. Current Pipeline Status (Post-Phase 6)
- **Step 1 (Reading Order Alignment)**: Fully implemented and actively remediated in-place.
- **Step 2 (Selective Ungrouping)**: Fully implemented with recursive tree traversal, spatial coordinate promotion, and protected picture preservation.
- **Step 3 (Native Placeholder Conversion)**: Fully implemented with layout/master inheritance traversal, geometry materialization, title box conversion, and strict semantic preservation.
- **Step 4 (List Semantics & Numbering Preservation)**: Fully implemented with OOXML `<a:buChar>`, `<a:buAutoNum>`, startAt preservation, multi-level hierarchy retention, inheritance resolution across layout/master text styles, and safe deterministic remediation.
- **Steps 5–6**: Declared `NOT_IMPLEMENTED` with clear auditing until upcoming milestone phases.
