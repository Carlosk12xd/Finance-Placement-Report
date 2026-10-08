"""Run with: python -m streamlit run app.py"""
from hashlib import sha256
from pathlib import Path
import hmac
import json
import os

import pandas as pd
import streamlit as st

from finance.logic import ReportError
from finance.service import ROOT, generate, settings
from finance.uploads import common_export_date, sheet_names

st.set_page_config(page_title="Finance placement reports", page_icon="📊", layout="wide")


def access_gate():
    password = os.environ.get("APP_PASSWORD", "")
    if not password:
        try:
            password = str(st.secrets.get("app_password", ""))
        except FileNotFoundError:
            password = ""
    if not password:
        return
    if st.session_state.get("signed_in"):
        if st.sidebar.button("Sign out"):
            st.session_state.clear()
            st.rerun()
        return
    st.title("Finance placement reports")
    with st.form("sign_in"):
        entered = st.text_input("Staff password", type="password")
        submit = st.form_submit_button("Sign in", type="primary")
    if submit:
        if hmac.compare_digest(entered.encode(), password.encode()):
            st.session_state.signed_in = True
            st.rerun()
        else:
            st.error("The password is incorrect.")
    st.stop()


def clear_session():
    # Change uploader keys, rather than assigning to file-uploader widget state.
    st.session_state.clear()
    st.session_state["upload_generation"] = os.urandom(8).hex()
    st.rerun()


def upload_spec(upload, kind, key):
    if upload is None:
        return None
    data = upload.getvalue()
    names = sheet_names(data, upload.name)
    sheet = None
    row = 1
    if names:
        with st.expander(f"Worksheet and header row: {upload.name}"):
            sheet = st.selectbox("Worksheet", names, key=f"sheet_{key}")
            row = int(st.number_input("Header row", min_value=1, max_value=25, value=1, key=f"header_{key}"))
    return {"data": data, "name": upload.name, "sheet": sheet, "header_row": row}


def input_fingerprint(mode, inputs, report_date, previous):
    digest = sha256(mode.encode())
    digest.update(str(report_date).encode())
    for kind, item in sorted(inputs.items()):
        digest.update(kind.encode())
        digest.update(item["data"])
        digest.update(json.dumps({k: v for k, v in item.items() if k != "data"}, sort_keys=True).encode())
    if previous:
        digest.update(previous["name"].encode())
        digest.update(previous["data"])
    return digest.hexdigest()


access_gate()
try:
    scope, presentation = settings()
except ReportError as error:
    st.error(str(error))
    st.stop()

with st.sidebar:
    st.subheader("Report scope")
    st.markdown("**Finance students only**")
    st.caption("Placement tracker")
    st.write("December 2026 · April 2027 · June 2027")
    st.caption("Historical statistics")
    st.write("Academic classes of 2024, 2025, 2026")
    st.divider()
    st.caption("Fixed Excel format 1.0.0")
    st.button("Clear uploads and reports", on_click=clear_session)

st.title("Finance placement reports")
st.write("Upload your source exports to create the placement tracker and historical placement statistics in the approved Excel format.")
st.caption("Other majors and graduation cohorts are excluded. Placement rate uses every student in the requested cohort.")
mode = st.radio("Upload method", ["Three export files", "One Excel workbook", "Try a fictional example"], horizontal=True, key="mode")
generation = st.session_state.get("upload_generation", "0")
inputs, upload_error = {}, None

st.subheader("1. Source exports")
try:
    if mode == "Three export files":
        for column, kind, label in zip(st.columns(3), ("students", "internships", "offers"), ("Students", "Internship placements", "Full-time placements")):
            with column:
                upload = st.file_uploader(label, type=["xlsx", "csv"], key=f"upload_{kind}_{generation}")
                spec = upload_spec(upload, kind, f"{kind}_{generation}")
                if spec:
                    inputs[kind] = spec
    elif mode == "One Excel workbook":
        upload = st.file_uploader("Workbook containing the three source tables", type=["xlsx"], key=f"combined_{generation}")
        if upload:
            data = upload.getvalue()
            names = sheet_names(data, upload.name)
            if len(names) < 3:
                raise ReportError("This workbook needs separate Students, Internship Placements, and Full-time Placements worksheets.")
            for column, kind, label, hint in zip(st.columns(3), ("students", "internships", "offers"), ("Students worksheet", "Internship worksheet", "Full-time worksheet"), ("student", "intern", "full")):
                with column:
                    default = next((i for i, name in enumerate(names) if hint in name.lower()), 0)
                    sheet = st.selectbox(label, names, index=default, key=f"combined_sheet_{kind}_{generation}")
                    row = int(st.number_input("Header row", min_value=1, max_value=25, value=1, key=f"combined_header_{kind}_{generation}"))
                    inputs[kind] = {"data": data, "name": upload.name, "sheet": sheet, "header_row": row}
            if len({i["sheet"] for i in inputs.values()}) != 3:
                raise ReportError("Select a different worksheet for each source table.")
    else:
        st.info("This example uses fictional students and employers.")
        for kind, name in (("students", "Students_2026_09_30.csv"), ("internships", "Internship_Placements_2026_09_30.csv"), ("offers", "Full_Time_Placements_2026_09_30.csv")):
            inputs[kind] = {"data": (ROOT / "examples" / name).read_bytes(), "name": name, "sheet": None, "header_row": 1}
    known_date = common_export_date([i["name"] for i in inputs.values()])
except ReportError as error:
    upload_error, known_date = str(error), None
    st.error(upload_error)

with st.expander("Required source columns"):
    st.write("Use the original Students and placement exports. An existing placement tracker alone does not contain the source IDs and offer records needed for both reports.")
    from finance.logic import REQUIRED_COLUMNS
    for kind, label in (("students", "Students"), ("internships", "Internships"), ("offers", "Full-time placements")):
        st.markdown(f"**{label}**")
        st.caption(", ".join(REQUIRED_COLUMNS[kind]) + (", Student ID.id (or StudentID.id)" if kind == "offers" else ""))

st.subheader("2. Export date and previous tracker")
date_column, prior_column = st.columns(2)
with date_column:
    identity = sha256("|".join(i["name"] for i in inputs.values()).encode()).hexdigest()[:12]
    report_date = st.date_input("Date of these exports", value=known_date, disabled=bool(known_date), key=f"as_of_{mode}_{identity}")
    st.caption("The report uses the date in the filenames. If no date is present, select the actual export date.")
with prior_column:
    old = st.file_uploader("Previous tracker (optional)", type=["xlsx"], key=f"previous_{generation}")
    previous = {"data": old.getvalue(), "name": old.name} if old else None
    st.caption("Preserves unchanged Review fields and adds a separate list of new students and changed fields.")

fingerprint = input_fingerprint(mode, inputs, report_date, previous)
if st.session_state.get("result_fingerprint") != fingerprint:
    st.session_state.pop("reports", None)
ready = len(inputs) == 3 and report_date is not None and not upload_error
if st.button("Generate Excel reports", type="primary", disabled=not ready, key="generate"):
    st.session_state.pop("reports", None)
    try:
        with st.spinner("Preparing and checking the reports…"):
            reports = generate(inputs, report_date, previous)
        st.session_state.reports = reports
        st.session_state.result_fingerprint = fingerprint
    except ReportError as error:
        st.error(str(error))
    except Exception:
        st.error("The reports could not be generated. Check the uploads and try again. If this persists, ask the app maintainer to check the server logs.")
        import logging
        logging.getLogger(__name__).exception("Report generation failed")

reports = st.session_state.get("reports")
if reports:
    st.divider()
    st.subheader("3. Download reports")
    model = reports["model"]
    tracker_count, placed = len(model["trackerRows"]), sum(s["placed"] for s in model["trackerStats"])
    for column, label, value in zip(st.columns(4), ("Tracker students", "Students placed", "Placement rate", "Records to review"), (f"{tracker_count:,}", f"{placed:,}", f"{placed/tracker_count:.1%}" if tracker_count else "—", f"{len(model['reviewRows']):,}")):
        column.metric(label, value)
    st.success(f"Reports prepared from exports dated {model['asOf']}.")
    excel_files = [(name, data) for name, data in reports["files"].items() if name.endswith(".xlsx")]
    for column, (name, data), label in zip(st.columns(2), excel_files, ("Download placement tracker", "Download historical statistics")):
        column.download_button(label, data, file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore", width="stretch")
    st.download_button("Download all reports", reports["bundle"], file_name=reports["bundle_name"], mime="application/zip", on_click="ignore")
    if reports["changes"] is not None:
        st.caption(f"The comparison includes {len(reports['changes']['rows']):,} new-student or changed-field rows within the current tracker scope.")
    if model["futureAdmitCount"]:
        st.warning(f"Current enrollment is unverified. {model['futureAdmitCount']} students have future admit terms; confirm them on the Review tab.")
    st.caption("Admit semester uses Start Semester. Phone is the available source number; mobile type is unverified. Keep coach updates in the shared working sheet.")
    tracker_tab, stats_tab, review_tab = st.tabs(["Placement tracker preview", "Historical statistics", "Review"])
    with tracker_tab:
        st.dataframe(pd.DataFrame(model["trackerRows"], columns=presentation["headers"]["tracker"]), hide_index=True, width="stretch", height=400)
    with stats_tab:
        rows = [{"Class of": s["year"], "Students": s["students"], "Placed": s["placed"], "Placement rate": f"{s['placed']/s['students']:.1%}" if s["students"] else "—", "Actively Seeking": s["activelySeeking"], "Not Seeking": s["notSeeking"], "no status": s["noStatus"], "Internship recorded": s["internships"]} for s in model["historicalStats"]]
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
    with review_tab:
        st.dataframe(pd.DataFrame(model["reviewRows"], columns=presentation["headers"]["review"]), hide_index=True, width="stretch")
else:
    st.caption("Upload all three source tables and confirm the export date to enable generation.")
