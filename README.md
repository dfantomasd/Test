# Flight Task Automation

### Intelligent automation for airline operations

A production-oriented Windows automation project that coordinates a proprietary airline desktop system, Excel, route data and PDF generation to turn a repetitive operational process into a repeatable workflow.

> Built from a real workflow used while working with **Nordwind Airlines**. This repository is personal portfolio material and is **not an official Nordwind Airlines product or repository**. Proprietary templates, operational records, credentials, internal paths and company-specific mappings are intentionally excluded.

## Why this project matters

This is more than a script that clicks buttons. The application has to understand the current operational context, collect data from several sources, normalize imperfect route data, make decisions based on route type, validate the generated document set and recover cleanly when an external application behaves unexpectedly.

The current production approach is best described as **intelligent process automation** rather than an AI-powered product. I deliberately do not label deterministic automation as “AI”. The architecture, however, is designed so that an AI layer can be added where probabilistic reasoning actually provides value.

## The problem

Preparing flight-task packages required repeated manual work across an airline operations desktop application, Microsoft Excel and PDF documents. The workflow involved opening the correct flight, triggering several document actions, handling Excel workbooks, exporting pages, assembling the final package and repeating the same sequence for multiple flights.

The project was built to reduce repetitive operator actions, make package generation more consistent and create a foundation for further intelligent automation.

## What the automation does

1. Finds and focuses the airline operations application on Windows.
2. Drives desktop UI actions with `pywinauto` / Win32 automation.
3. Connects to Excel through COM automation.
4. Extracts route and flight metadata.
5. Resolves inconsistent airport names and IATA/ICAO codes with fuzzy matching.
6. Classifies route context and applies domestic/international workflow logic.
7. Exports required Excel sheets to PDF.
8. Validates and merges document parts into one final flight-task package.
9. Supports single-flight and batch processing.
10. Writes a result payload and cleans temporary resources.

## Automation pipeline

```text
                         ┌──────────────────────┐
                         │      Operator        │
                         └──────────┬───────────┘
                                    │
                                    ▼
                         ┌──────────────────────┐
                         │ Flight Task Launcher │
                         └──────────┬───────────┘
                                    │
                   ┌────────────────┼────────────────┐
                   ▼                ▼                ▼
          Airline desktop      Excel / COM      Route data
             application                         + airports
                   │                │                │
                   └────────────┬───┴────────────────┘
                                ▼
                    ┌────────────────────────┐
                    │ Decision / validation  │
                    │ • metadata parsing     │
                    │ • fuzzy matching       │
                    │ • route classification │
                    │ • package rules        │
                    └────────────┬───────────┘
                                 ▼
                    ┌────────────────────────┐
                    │ PDF generation & merge │
                    └────────────┬───────────┘
                                 ▼
                    ┌────────────────────────┐
                    │ Final flight-task PDF  │
                    └────────────────────────┘
```

More detail: [`docs/architecture.md`](docs/architecture.md)

## Intelligent automation layer

The project already contains decision-oriented components rather than a purely linear macro:

- fuzzy entity resolution for airports and route points;
- normalization of inconsistent human-entered text;
- context-dependent branching for route types;
- validation before final document assembly;
- state-aware interaction with external Windows applications;
- batch orchestration with recovery and cleanup boundaries.

These are deterministic today because deterministic rules are safer for operational document generation.

### AI roadmap

The natural next stage is to add AI only where it improves the workflow without replacing hard safety rules:

- **LLM exception assistant** — explain why a package failed and suggest the next operator action;
- **document completeness review** — analyze extracted document metadata and flag unusual/missing combinations;
- **natural-language operations interface** — allow an operator to request a task in plain language while the system converts it into validated actions;
- **semantic route/entity resolution** — use embeddings or an LLM fallback only when deterministic airport matching is uncertain;
- **structured incident summaries** — convert technical automation logs into concise operator-facing explanations.

The rule-based workflow remains the source of truth; AI would sit above it as an assistive reasoning layer.

## Tech stack

**Automation & integration**

- Python 3
- `pywinauto`
- Win32 / `pywin32`
- Microsoft Excel COM automation
- Tkinter

**Data & decision logic**

- `airportsdata`
- `rapidfuzz`
- `Unidecode`
- route classification and validation rules

**Document processing**

- `pypdf`
- atomic temporary builds and cleanup

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

This is not a tutorial project. It demonstrates how I approach a real automation problem when the systems were not designed to integrate with each other:

- Windows UI automation where no convenient public API is available;
- orchestration of several external applications;
- Excel automation through COM;
- defensive handling of windows, dialogs and generated files;
- fuzzy matching and entity resolution;
- branching business logic for different operational contexts;
- PDF generation and package assembly;
- temporary-file cleanup and repeatable batch execution;
- designing a deterministic core that can later support an AI assistant;
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

Without the authorized airline desktop environment, the worker stops at the integration boundary rather than attempting to reproduce proprietary behavior.

## Engineering principles

- Deterministic rules for operationally critical decisions.
- AI only where uncertainty, language or explanation benefits from it.
- Fail visibly instead of silently producing an incomplete package.
- Keep temporary outputs isolated until the final PDF is ready.
- Avoid saving changes back into source Excel workbooks.
- Separate launcher/orchestration concerns from document-processing logic.
- Keep public code free from operational data and secrets.

## Author

**Dmitry Simutin**  
Python · Intelligent Automation · AI Automation
