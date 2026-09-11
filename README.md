# Flight Task Automation

Windows desktop automation for assembling flight-task document packages from an airline operations application and Microsoft Excel.

This portfolio project was developed to automate a real operational workflow at **Nordwind Airlines**. The public repository is a **sanitized portfolio snapshot**: company templates, real flight documents, operational data, internal paths, and other non-public materials are intentionally excluded.

## What it automates

- launches document-generation actions in the airline operations desktop application;
- detects and controls the relevant Windows application dialogs;
- attaches to Excel workbooks opened by the source application;
- exports required sheets to PDF;
- assembles several documents into one final PDF package;
- handles domestic and international routing logic;
- resolves airports and IATA codes using `airportsdata` and fuzzy matching;
- supports single-flight and batch processing;
- keeps the launcher open for repeated runs;
- cleans temporary files and closes generated Excel workbooks after processing.

## Why I built it

The original process required repeated manual interaction with the planning system, Excel, document windows and PDF files. The goal was to reduce repetitive work, lower the chance of missed documents, and make package generation more consistent.

## Tech stack

- Python 3
- `pywinauto`
- Windows API / `pywin32`
- Excel COM automation
- `pypdf`
- `airportsdata`
- `rapidfuzz`
- Tkinter

## Repository structure

```text
src/
  flight_task_menu.py      # desktop launcher and batch workflow
  flight_task_worker.py    # document capture, Excel/PDF processing and assembly
requirements.txt
.gitignore
README.md
```

## Notes

This code depends on a specific Windows desktop environment and a proprietary airline operations application. It is not intended to run out of the box outside that environment. The repository is published to demonstrate the automation architecture, Windows UI automation, document processing, error handling and workflow design.

## Privacy / sanitization

The public version does **not** include real passenger or crew data, company documents, internal Excel templates, network shares, credentials, emails, or production logs.

## Author

Dmitry Simutin — Python / AI automation portfolio project.
