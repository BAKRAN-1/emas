# EMAS — Invoice Reconciliation Tool

A Python desktop application for extracting invoice data from Excel workbooks, matching bill numbers against SQL Server records, looking up claim document identifiers through an API, and compiling reconciliation reports.

## Features

- Select an Excel workbook and worksheet, then map date, bill-number and amount columns using Excel letters.
- Optionally combine two bill-number columns and remove letters from bill identifiers.
- Match extracted bills against `VW_SOAReconInfo` and retrieve claim document IDs.
- Preview, edit and remove extracted data; compile it into an Excel report.
- Save and reload local reconciliation sessions.
- Insert or update reconciliation records in `tbl_SOARecon`.

## Project structure

```text
main.py                    Application entry point
64-check.py                Compatibility entry point for the original filename
src/emas/
  app.py                   Application controller, sign-in form and main window
  extraction.py            Workbook loading, column mapping and extraction callbacks
  views.py                 Previews, editing, worksheet lists and report export
  sessions.py              Session save/load, reset and close handling
  navigation.py            Widget navigation and appearance callbacks
  helpers.py               Bill formatting, reference IDs and Excel-column conversion
  database.py              SQL queries, record preparation and database operations
  database_actions.py      Reconciliation insert/update workflow
  claims.py                Claims API client
  config.py                Local configuration paths and environment settings
config/db.example.json     Database configuration template
data/sample/               Synthetic example data only
tests/                     Automated regression checks
docs/                      Review notes and operational limitations
```

The controller keeps application state on one instance. Feature classes share that controller through mixins, so modules do not copy mutable global variables or depend on circular imports.

## Setup

Use Python 3.10 or newer with Tkinter (the tested environment used Python 3.12).

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item config/db.example.json config/db.json
```

Fill `config/db.json` with your authorized database connection settings. Install the matching SQL Server ODBC driver on the computer; the example uses the `SQL Server` driver name. A different installed driver can be selected in the configuration.

Set environment variables in the same terminal before launching:

```powershell
$env:EMAS_API_BASE_URL = 'https://your-service.example/api/v1'
$env:EMAS_API_KEY = 'your-local-key'
.\.venv\Scripts\python.exe main.py
```

`EMAS_DB_CONFIG` optionally overrides the database configuration path. `.env.example` documents the settings; `.env` files are not loaded automatically.

Opening the interface does not require a database connection. Extraction and database operations require the original SQL schema and an authorized API endpoint. The sign-in form records the operator's name and reconciliation dates; it is not an authentication system.

To generate a small synthetic workbook:

```powershell
.\.venv\Scripts\python.exe scripts/create_demo_workbook.py
```

This creates `data/sample/invoices.xlsx`; its date, bill-number and amount columns are A, B and C. The sample supports workbook selection and column mapping. Full extraction still requires configured services.

## Validation

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

The regression suite uses synthetic workbooks and mocked SQL/API calls. It covers import behavior, controller wiring, column conversion, reference generation, sheet alignment, session round trips, cancelled saves, parameterized SQL lookup, failed database writes and failed extraction replacement.

Live SQL/API integration and visual desktop behavior require validation in the deployment environment. No real hospital, employee, patient or invoice data is included.

## Local data and credentials

Store real workbooks, exports and sessions under `data/`; these files and `config/db.json` are ignored by Git. Never commit an API key or database password. Only load saved pickle sessions created by this application that you trust, because pickle files can execute code while loading.

The original repository revision contained a hardcoded API key. The refactor removes it from the current source, but does not remove it from Git history or revoke it. Replace that key at the service before using or sharing this integration further.

See [review notes](docs/REVIEW.md) for changes and remaining limitations.
