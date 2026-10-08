# SlideGuard Testing & Verification Strategy

## 1. Test Architecture
Testing is structured across four rigorous levels:

```
tests/
├── fixtures/     # Programmatic PPTX generator for all test topologies
├── unit/         # Isolated component tests (models, rules, geometry, reader, workspace)
├── integration/  # End-to-end controller and pipeline execution tests
└── regression/   # Visual design preservation and comparison tests
```

## 2. Test Execution Commands

Run the complete test suite:
```bash
pytest
```

Run with verbose output and test names:
```bash
pytest -v
```

Run only unit tests:
```bash
pytest tests/unit/
```

Run only integration tests:
```bash
pytest tests/integration/
```

Run design preservation regression tests:
```bash
pytest tests/regression/
```

## 3. Preservation & Regression Requirements
Every remediation rule implemented in future phases must include a corresponding preservation test validating that:
1. Coordinates do not shift beyond tolerance.
2. Font styling is preserved.
3. Slide count and element identities match.
4. Validated candidate presentations reopen cleanly in PowerPoint without repair prompts.
