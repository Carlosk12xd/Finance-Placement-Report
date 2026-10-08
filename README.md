# Finance Placement Streamlit App

A standalone app for the Finance placement tracker and historical placement statistics. Upload Excel or CSV source exports, then download the two reports in the approved format. No ChatGPT account, Node.js, database connection, or API key is required.

## Start locally

Use Python 3.12. Extract this package and open a terminal in the extracted folder containing `app.py`.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

On Windows, activate with `.venv\Scripts\activate` instead. Open the local URL printed in the terminal.

## Use the app

1. Choose **Three export files** or **One Excel workbook**.
2. Upload the Students, Internship Placements, and Full-time Placements source tables. Each can be `.xlsx` or `.csv`. For one workbook, choose a separate worksheet for each source. The worksheet and header-row controls support exports with title rows above the column headings.
3. Confirm the actual export date. A date in a filename, such as `Students_2026_09_30.xlsx`, is used automatically. Conflicting dates are rejected. For undated filenames, explicitly choose the date the exports were taken; the app does not substitute today's date.
4. Optionally upload the previous tracker in this format. Unchanged Review status and Reviewed on are preserved. Changed issues return to Pending. New students and changed fields appear in a separate changes CSV.
5. Click **Generate Excel reports** and download either workbook or the combined ZIP. Changing an upload, worksheet, header row, or export date clears the old download results until you regenerate.

Choose **Try a fictional example** to test the app immediately. `examples/Finance_Source_Exports_2026_09_30.xlsx` is also a complete three-tab upload example with real Excel date cells and Boolean acceptance values. All examples and templates use fictional students; no real student records are included.

The required source fields are in the app's **Required source columns** section and `REPORT_SPEC.md`. An existing tracker by itself cannot replace the three source exports: it lacks the source IDs and historical offer records needed to build both reports. The previous tracker upload is for comparison and Review fields.

## Host on Streamlit Community Cloud

1. Put the extracted package contents in a GitHub repository. Keep `app.py`, `requirements.txt`, the `finance` folder, `assets`, `examples`, the JSON settings, and `.streamlit/config.toml`. Upload the contents, not only the ZIP.
2. At [Streamlit Community Cloud](https://share.streamlit.io), create an app using that repository and select **app.py** as the entrypoint.
3. In Advanced settings, select **Python 3.12**.
4. In Secrets, configure a staff password using this template:

   ```toml
   app_password = "your-own-long-staff-password"
   ```

5. Deploy. Keep the app's access limited to the staff who use these student reports. The app has a sign-in screen when the password is configured; a private host can instead enforce its own access controls.

For local sign-in, copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and set your password. `APP_PASSWORD` as an environment variable takes precedence. The real secrets file is ignored by Git and Docker. Without either configured password, the app runs without its own sign-in screen, which is convenient for local testing or a host that already controls access.

Official guides: [deployment](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy), [dependencies](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies), and [secrets](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management).

## Host with Docker

From the package root:

```bash
docker build -t finance-placement .
docker run --rm -p 8501:8501 -e APP_PASSWORD finance-placement
```

Set `APP_PASSWORD` in the host environment before running that command, or use your hosting platform's secret settings. Serve the container through your host's HTTPS endpoint. The supplied container binds port 8501 and runs as a non-root user.

## Reports and fixed scope

| Report | Included students | Worksheets |
|---|---|---|
| Placement tracker | BSFin, expected Fall 2026 / Winter 2027 / Spring 2027; recorded actual graduates excluded | Placement Summary, Placement Tracker, Review |
| Historical statistics | BSFin, academic Class Of 2024 / 2025 / 2026 | Placement Summary, Student Detail |

The ten tracker columns, colors, bold identifiers, spacing, widths, sheet order, summary sections, filters and frozen headers remain fixed. Rows and wrapped heights expand with the uploaded data. Graduation cohorts do not advance automatically.

Placed means **Accepted Offer status**. Placement rate uses the whole requested cohort, including Not Seeking and no status. The historical report foregrounds placed students, status counts, rates, year-to-year changes and internship/offer coverage. Internship-employer offer counts remain a small supporting section. No same-company conversion rates or conversion-rate charts are generated.

The source exports do not establish current enrollment. Start Semester is the available Finance admit-term proxy, and Phone is the available number without verified mobile type. Review flags preserve these limitations. Coach notes and manual placement corrections remain in your shared working sheet; this app creates separate report snapshots.

## How the fixed format works

The app fills the two approved XLSX templates in `assets`. They were authored with the original spreadsheet renderer. Python's standard library updates data rows, formulas, their numeric caches, ranges, report dates and source labels. `openpyxl` is used only to read uploads. The hosted app does not need the original proprietary Excel writer.

`format.json`, `scope.json`, `layout_reference.json`, and `format-lock.json` fix the layout and reporting contract. Each run checks code/template hashes, exported headings and styles, table filters, frozen panes, unique scoped students, cached formula values, counts and rates before exposing downloads. Excel formulas remain editable and are set to recalculate on opening. Formula-like uploaded text is written as text, not an executable formula.

Routine refreshes require no code edits. Intentional changes to the report format or cohorts require updating the contract, code, reference checks and version together. Do not remove checks or change hashes just to bypass an unexplained difference.

## Data handling

The app reads uploaded bytes and generates downloads in memory. It does not write student uploads or generated reports to server files, use a shared data cache, send emails, or connect to Zoho. Each browser session has its own result state. **Clear uploads and reports** clears that session's inputs and downloads. The host controls access, transport security and runtime retention.

Each source upload is limited to 25 MB and 100,000 rows. `.xlsx` and `.csv` are supported; save older `.xls` files as `.xlsx`. Dates can be native Excel dates, ISO `YYYY-MM-DD`, or US `MM/DD/YYYY`. Export record IDs and NetIDs as text; the app rejects long numeric identifiers that Excel may already have rounded. Source formulas in required columns must be converted to values in an export copy.

## Tests

```bash
python -m unittest discover -s tests -v
```

The package includes upload, cohort, date, join, placement, template, formula-cache and Streamlit interaction tests. `TEST_RESULTS.md` documents release verification against the real September 30 exports and the original generator. Docker deployment itself requires a Docker host; this package does not publish a live app.

For scheduled or command-line use, `run_reports.py` uses the same generation service:

```bash
python run_reports.py \
  --students exports/Students_2026_09_30.xlsx --students-sheet Students \
  --internships exports/Internship_Placements_2026_09_30.xlsx --internships-sheet Internships \
  --offers exports/Full_Time_Placements_2026_09_30.xlsx --offers-sheet Offers \
  --as-of 2026-09-30 --out-dir reports/september
```

Use the actual worksheet names and export date. For CSV files, omit the worksheet options. Add `--previous-tracker path/to/previous.xlsx` to preserve Review fields and generate the changes CSV. The CLI refuses to overwrite existing outputs.
