# Requirement-Management-Tool

This repository contains a PyQt-based requirement management utility inspired by
tools such as DOORS and JAMA. The application has been refactored to separate
its user interface, data handling, and traceability logic so that each concern
can be reviewed and verified in isolation as recommended for DO-178C compliant
software development.

## Project layout

```
RequirementTool.py              # Legacy entry point that now delegates to run_app()
requirement_tool/
├── __init__.py                 # Package exports for data manager
├── data_manager.py             # Deterministic Excel loading & numbering logic
└── ui/
    ├── main_window.py          # Main PyQt window coordinating widgets
    └── trace_view.py           # Traceability matrix widget with filtering & export
```

## Running the tool

```bash
python RequirementTool.py
```

## Key improvements

- Deterministic data ingestion through :class:`RequirementDataManager` with
  schema validation and structured error handling.
- Modular widgets that encapsulate traceability behaviour in
  ``TraceMatrixView``; the main window can be unit tested independently from the
  GUI event loop.
- Centralised HTML generation for Word previews to make verification easier and
  to support automated testing without the GUI.

These changes make it easier to apply DO-178C verification activities such as
independent reviews, traceability analysis, and unit-level testing.
