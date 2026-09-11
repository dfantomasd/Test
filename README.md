# Flight Task Automation

Production-oriented Windows desktop automation for preparing airline flight-task document packages.

> Portfolio snapshot of a tool developed to support a real operational workflow while working with **Nordwind Airlines**. This repository is personal portfolio material and is **not an official Nordwind Airlines product or repository**. Proprietary templates, operational records, credentials, internal paths and company-specific mappings are intentionally excluded.

## The problem

Preparing flight-task packages required repeated manual work across an airline operations desktop application, Microsoft Excel and PDF documents. The workflow involved opening the correct flight, triggering several document actions, handling Excel workbooks, exporting pages, assembling the final package and repeating the same sequence for multiple flights.

The project was built to make that process more consistent and to reduce repetitive operator actions.

## What the automation does

1. Finds and focuses the airline operations application on Windows.
2. Drives desktop UI actions with `pywinauto` / Win32 automation.
3. Connects to Excel through COM automation.
4. Extracts route and flight metadata.
5. Resolves airport names and IATA/ICAO codes with fuzzy matching.
6. Applies domestic/international routing logic.
7. Exports required Excel sheets to PDF.
8. Merges validated document parts into one final PDF package.
9. Supports single-flight and batch processing.
10. Writes a small result payload and cleans temporary resources.

## Workflow

```text
Operator
   |
   v
Flight Task launcher
   |
   v
Airline desktop application ----> active flight / route
   |
   v
Excel workbooks via COM
   |
   +--> metadata parsing
   +--> airport / IATA resolution
   +--> route-specific validation
   +--> PDF export
   |
   v
PDF assembly
   |
   v
Final flight-task package
```

More detail: [`docs/architecture.md`](docs/architecture.md)

## Tech stack

- Python 3
- Tkinter
- `pywinauto`
- Win32 / `pywin32`
- Microsoft Excel COM automation
- `pypdf`
- `airportsdata`
- `rapidfuzz`
- `Unidecode`

## Repository structure

```text
.
├── src/
│   ├── launcher.py       # desktop launcher and batch orchestration
│   └── worker.py         # sanitized automation / Excel / PDF / route logic
├── docs/
│   └── architecture.md   # architecture and design decisions
├── config.example.json
├── requirements.txt
├── .gitignore
└── README.md
```

## What this project demonstrates

This is not a tutorial project. It demonstrates integration work around a real desktop workflow:

- Windows UI automation where no convenient public API is available;
- orchestration of several external applications;
- Excel automation through COM;
- defensive handling of windows, dialogs and generated files;
- fuzzy matching of airport data;
- branching business logic for different route types;
- PDF generation and package assembly;
- temporary-file cleanup and repeatable batch execution;
- sanitizing a production-oriented project for a public portfolio.

## Public snapshot limitations

The production workflow depends on an authorized Windows environment and a proprietary airline operations application. The public code therefore cannot reproduce the complete production process outside that environment.

The following are deliberately not published:

- company document templates;
- real crew, passenger or flight records;
- internal network paths and shares;
- proprietary menu/document mappings;
- credentials, email addresses and production logs;
- confidential operational rules that are not necessary to demonstrate the engineering approach.

Where production-specific behavior was removed, the public code keeps representative integration points so the architecture remains understandable.

## Running the portfolio snapshot

The source is intended primarily for code review. On Windows, install dependencies with:

```bash
python -m pip install -r requirements.txt
```

Then launch:

```bash
python src/launcher.py
```

Without the authorized airline desktop environment, the worker will stop at the integration boundary rather than attempting to reproduce proprietary behavior.

## Design priorities

- Fail visibly instead of silently producing an incomplete document package.
- Keep temporary outputs isolated until the final PDF is ready.
- Avoid saving changes back into source Excel workbooks.
- Separate launcher/orchestration concerns from document-processing logic.
- Keep the public repository free from operational data and secrets.

## Author

**Dmitry Simutin**  
Python / automation portfolio project
