# Chipunza House Project Live Dashboard

Live dashboard bundle for the Chipunza house project.

## What is here

- `index.html` renders the public live dashboard.
- `dashboard_data.json` is the published data payload that the page reads.
- `scripts/refresh_dashboard_data.py` converts the control workbook into `dashboard_data.json`.
- `scripts/publish_dashboard_data.py` refreshes the JSON and pushes it when the folder is inside a Git repo with an `origin`.
- `scripts/register_local_autopublish.ps1` creates a Windows scheduled task that republishes the dashboard every minute from this PC.

## Source workbook

The dashboard reads from:

- `../03_Deliverables/Chipunza_Project_Control_Live.xlsx`

## Typical local flow

1. Run `python scripts/refresh_dashboard_data.py`
2. Open `index.html` locally to preview
3. Run `python scripts/publish_dashboard_data.py` to refresh the JSON and push it
4. Publish through GitHub Pages

## Live update model

For true live updates, keep the workbook current and run the local auto-publish task on this PC. That task regenerates `dashboard_data.json` and pushes any changes to GitHub.
