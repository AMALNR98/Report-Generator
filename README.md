# Monthly, Weekly, Daily Alert Report Generator

A fully local, self-hosted Flask web app that converts a raw Alert Table Excel export into a formatted Alert Report Excel workbook. The app uses local Excel templates for Monthly, Weekly, and Daily report generation.

No Docker, hosted AI APIs, cloud AI services, external embeddings, telemetry, or external user-data transfer are used.

## Features

- Browser UI at `http://localhost:5000`
- Monthly, Weekly, and Daily report type selection
- Local template selection based on report type
- Raw `.xlsx` and `.xls` alert export upload
- Drag and drop upload support
- Alert field mapping and preview table
- Excel preview download
- Generated report download using the selected local template
- Browser File System Access API support for direct save when available
- Normal browser download fallback
- Preview statistics for total records, Critical count, Major count, and SLA Y count
- Color-coded Critical, Major, Minor, SLA Y, and SLA N values
- In-memory upload processing where possible
- Clear errors for missing files, missing columns, missing templates, bad workbooks, and missing `Report` worksheet

## Required Folder Structure

```text
app.py
requirements.txt
README.md
templates_excel/monthly_template.xlsx
templates_excel/weekly_template.xlsx
templates_excel/daily_template.xlsx
```

The `templates_excel/` folder is included with a placeholder file. Add the three template workbooks before generating reports.

## Setup

```bash
pip install -r requirements.txt
python app.py
```

Open:

```text
http://localhost:5000
```

## Usage

1. Place the three local template files in `templates_excel/`.
2. Start the app with `python app.py`.
3. Open `http://localhost:5000`.
4. Select `Monthly`, `Weekly`, or `Daily`.
5. Upload only the raw alert export `.xlsx` or `.xls` file.
6. Click `Process & Preview`.
7. Review the preview table and statistics.
8. Click `Download Excel Preview` to download the mapped preview workbook.
9. Click `Download Generated Report` to download the filled template report.
10. Click `Save Report to Folder` to use the browser folder save option when supported, or normal download fallback when unsupported.

## Expected Raw Alert Columns

The uploaded raw alert export must include:

- `Time`
- `Alert Id`
- `Type`
- `Severity`
- `Confidence`
- `Origin.1` or `Origin`
- `Entity.1` or `Entity`
- `Entity Type`
- `Status`
- `Message`
- `Assignee Email`

Field mapping:

- `Time` -> `Date`
- `Alert Id` -> `Alert Id`
- `Type` -> `Type of Alert`
- `Severity` -> `Alert Level`
- `Confidence` -> `Severity`
- `Origin.1`, fallback `Origin` -> `Origin`
- `Entity.1`, fallback `Entity` -> `Entity`
- `Entity Type` -> `Entity Type`
- `Status` -> `Status`
- `Message` -> `Message`
- `Assignee Email` -> `SLA (Y/N)`

`SLA (Y/N)` is `Y` when `Assignee Email` contains `@`; otherwise it is `N`.

## Template Requirements

Templates must be local files in `templates_excel/`:

- Monthly: `templates_excel/monthly_template.xlsx`
- Weekly: `templates_excel/weekly_template.xlsx`
- Daily: `templates_excel/daily_template.xlsx`

Each template must:

- Be an `.xlsx` workbook
- Contain a worksheet named `Report`
- Have alert body rows beginning at row `15`
- Have enough existing body rows for all processed alert records

The app writes alert data into these columns on the `Report` sheet:

| Field | Column |
| --- | ---: |
| Date | 1 |
| Alert Id | 2 |
| Type of Alert | 5 |
| Alert Level | 8 |
| Severity | 11 |
| Origin | 12 |
| Entity | 14 |
| Entity Type | 16 |
| Status | 17 |
| Message | 18 |
| SLA (Y/N) | 19 |

The source template files are never modified. The app loads the selected template into memory, clears only the configured alert body cell values from row `15` through the template's max row, fills the processed alert rows, and returns a new workbook.

## Local-Only Notes

This project is self-hosted and fully local:

- No Docker is required or included.
- No hosted AI APIs are used.
- No cloud AI services are used.
- No external embeddings are used.
- No telemetry is included.
- Uploaded raw alert data is processed locally in memory and is not intentionally stored permanently.
- User data is not sent outside the local machine by this app.
