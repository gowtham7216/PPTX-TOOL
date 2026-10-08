# SlideGuard Architecture Specification

## 1. System Overview
SlideGuard is an offline, local Windows desktop application for analyzing, remediating, validating, and comparing PowerPoint (PPTX) presentations for digital accessibility without compromising visual design integrity.

The architecture strictly decouples the user interface from PowerPoint parsing, rule execution, and file I/O.

```
Desktop UI (Tkinter/TTK)
         │
         ▼
Application Controller (AppController)
         │
 ┌───────┴────────────────────────┐
 │                                │
 ▼                                ▼
Analysis Service          Remediation Service
 │                                │
 └───────┬────────────────────────┘
         │
         ▼
Rule Framework (IAccessibilityRule)
         │
         ▼
PPTX Engine (Reader, Writer, OOXML, Utilities)
         │
         ▼
Isolated Job Workspace (workspace/job_<id>/)
```

## 2. Key Architectural Layers

### Desktop UI (`app/ui/`)
- Minimalist, responsive Windows desktop interface implemented with standard Python `tkinter` / `ttk`.
- Offloads heavy PowerPoint analysis and remediation onto background daemon threads to maintain smooth 60fps UI responsiveness.
- Does not contain any PowerPoint parsing, XML inspection, or file remediation code.

### Application Controller (`app/controller/`)
- Central orchestrator that handles user commands.
- Coordinates workspace lifecycle, calls the presentation reader, triggers analysis, manages remediation approval, and requests report generation.

### PPTX Engine (`pptx_engine/`)
- **Reader (`reader/`)**: Loads presentations into structured internal models (`PresentationModel`, `SlideModel`, `ShapeModel`) using `python-pptx` augmented with direct `lxml` OOXML parsing.
- **Writer (`writer/`)**: Writes presentations to disk with validation.
- **OOXML (`ooxml/`)**: Provides low-level OpenXML element access (`p:cNvPr`, `a:tblPr`, etc.) for features not directly exposed by high-level libraries.
- **Utilities (`utilities/`)**: Coordinate geometry (EMU to Pt), bounding box intersection calculations, and spatial footer region detection.

### Accessibility Rules (`rules/`)
- Modular rule architecture where each rule implements `IAccessibilityRule`.
- Every rule declares its behavioral category (`DETECT_ONLY`, `SAFE_AUTO_REMEDIATE`, `AI_ASSISTED_REVIEW`) and target standard mapping (e.g. WCAG 2.1).
- Provides slide-level failure isolation: an exception raised while analyzing slide N does not abort analysis of remaining slides.

### Workspace Isolation (`workspace/`)
- Every execution run operates in a sandboxed directory under `workspace_runs/job_<timestamp>_<uuid>/`.
- Directory structure:
  - `original/`: Read-only copy of user source PPTX.
  - `working/`: Working copy for analysis.
  - `candidate/`: Remediated candidate presentation.
  - `output/`: Final verified output.
  - `reports/`: Audit JSON and Markdown reports.
  - `logs/`: Dedicated job audit execution log.
- Guarantees the original user file is never modified or overwritten by default.

### AI Isolation & Decision Hierarchy
- SlideGuard is fundamentally **deterministic and rule-based**.
- AI is never a mandatory runtime dependency.
- AI is strictly advisory (e.g., generating alt-text draft suggestions) and never writes directly to PPTX structures.
- User review and explicit approval is required before any AI suggestion can be committed.
