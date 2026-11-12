# BiharFeederOutage

Outage Dashboard Updation: the Bihar 11 kV feeder outage dashboard, and a page that makes the daily report in the browser and sends it to the dashboard.

| Page | Link | What it does |
|---|---|---|
| Dashboard | https://mdsbpdcl-hub.github.io/BiharFeederOutage/ | Overview and Highest Average Outage, with PDFs and call buttons. Anyone with the link can see it. |
| Update page | https://mdsbpdcl-hub.github.io/BiharFeederOutage/update.html | Choose the company CSV files, make the Excel reports, send them to the dashboard. Needs the upload password. |
| Agriculture list | https://mdsbpdcl-hub.github.io/BiharFeederOutage/update.html#list | The shared list of feeders that the companies type wrongly. Needs the upload password to change. |

## How it works

- The update page runs **the same program as the Windows app** (`py/outage_report_app.py`) inside the browser, so the figures and the Excel files are exactly the same. The CSV files never leave the computer; only the finished report is sent.
- The figures stay in the outage Google Sheet. The Apps Script web app (link in `config.js`) checks the password, reads the whole report before writing anything, and locks after 5 wrong passwords.
- The password is asked on the page and kept only while the page is open. It is never saved.
- The Agriculture list is kept in the tab "Agriculture List" of the outage Google Sheet. The update page and the Windows app (version 1.7) both read it before every report.

## Files

| File | What it is |
|---|---|
| `index.html` | The dashboard. Made from `apps_script/Dashboard.html` by `tools/build_index.py`. |
| `gas-shim.js` | Lets the dashboard read its figures from the web app's data link. |
| `update.html`, `worker.js` | The update page and the part that runs the report program in the background. |
| `config.js` | The web app link. Change it here if the web app is deployed again with a new link. |
| `py/outage_report_app.py` | The report program (same file as the Windows app). |
| `pyodide/` | Python for the browser (Pyodide 0.26.4, with sqlite3), kept here so the page does not depend on another site. |

## When something changes

- **Dashboard.html changed:** run `python tools/build_index.py path/to/Dashboard.html` and upload the new `index.html`.
- **The Windows app changed:** copy the new `outage_report_app.py` into `py/`.
- **New web app link:** change `config.js`, and `WEB_APP_URL` in the Windows app.
