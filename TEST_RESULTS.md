# Release verification — Streamlit app 1.0.0

Tested October 8, 2026 with Python 3.12, Streamlit 1.65.0, openpyxl 3.1.5 and pandas 2.2.3. The Excel report format remains version 1.0.0.

- **23 tests passed**, including the three-file upload flow, one-workbook upload flow, fictional demonstration, password gate and removal of stale results after an upload changes.
- The Excel example uses real Excel dates, Boolean acceptance fields and typed class years. It produces the same report data and statistics as the fictional CSV exports.
- Header-row selection, missing columns, duplicate/unmatched IDs, unknown statuses, invalid and mixed dates, completed graduates, future acceptance dates, employer-ID precedence and long numeric identifiers were checked.
- Prior Review status and date survive only when the issue is unchanged. An unchanged prior tracker produces no change rows.
- Empty exports produce valid report tables, zero counts and dashes for unavailable rates.
- Modified format files stop generation. Formula-like source text remains literal text in Excel.
- The real September 30 exports reproduced **310 tracker students, 57 placed, 665 historical students and 519 historical placed students**. Every tracker row, historical row, Review row and aggregate metric matched the original JavaScript generator. Comparison with the previous real tracker showed no changed fields.
- Both generated workbooks passed the approved headers, styles, dimensions, table, filter, freeze-pane, unique-student, population, cached-formula, count and rate checks.
- The real report files were imported into the original spreadsheet engine and recalculated. Counts and placement rates agreed with the independently calculated source totals. All five report sheets were rendered and visually inspected.
- Streamlit started successfully and passed its HTTP health check. UI interaction testing used Streamlit AppTest. A separate browser screenshot run was unavailable because the test environment lacked a Chromium executable.

No real student records or real student-level reports are included in this app package. The Dockerfile and Community Cloud instructions are supplied for deployment; a live host was not provisioned and a Docker image build was not executed here.
