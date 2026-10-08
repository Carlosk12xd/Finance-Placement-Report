"""One checked generation entrypoint shared by the UI, CLI and tests."""
from datetime import date
from io import BytesIO, StringIO
from pathlib import Path
from types import SimpleNamespace
import csv
import hashlib
import json
import zipfile

from .exporter import export_reports
from .logic import ReportError, analyze_sources, tracker_changes
from .uploads import common_export_date, read_previous, read_source

ROOT = Path(__file__).resolve().parents[1]


def settings(root=ROOT):
    lock = json.loads((root / "format-lock.json").read_text())
    for name, expected in lock["sha256"].items():
        if hashlib.sha256((root / name).read_bytes()).hexdigest() != expected:
            raise ReportError(f"Report format check failed for {name}. Restore the released file before generating reports.")
    return json.loads((root / "scope.json").read_text()), json.loads((root / "format.json").read_text())


def changes_csv(rows):
    output = StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(["NetID", "First name", "Last name", "Expected graduation", "Field", "Previous value", "Current value", "Change type"])
    for row in rows:
        writer.writerow([("'" + str(value)) if str(value).startswith(("=", "+", "-", "@")) else value for value in row])
    return output.getvalue().encode("utf-8-sig")


def generate(inputs, as_of, previous=None, root=ROOT):
    """inputs: kind -> {data: bytes, name: str, sheet: str|None, header_row: int}."""
    scope, presentation = settings(root)
    if set(inputs) != {"students", "internships", "offers"}:
        raise ReportError("Upload Students, Internship Placements, and Full-time Placements.")
    named_date = common_export_date([item["name"] for item in inputs.values()])
    if isinstance(as_of, date):
        as_of = as_of.isoformat()
    if named_date and named_date.isoformat() != as_of:
        raise ReportError("The report date must match the export date in the source filenames.")
    if len({id(item["data"]) for item in inputs.values()}) == 1 and len({item.get("sheet") for item in inputs.values()}) != 3:
        raise ReportError("Choose a different worksheet for each of the three exports.")
    sources = {kind: read_source(item["data"], item["name"], kind, item.get("sheet"), item.get("header_row", 1)) for kind, item in inputs.items()}
    old_rows, old_review = read_previous(previous["data"], previous["name"], presentation) if previous else ([], {})
    model = analyze_sources(sources, scope, as_of, old_review)
    source_labels = [item["name"] for item in inputs.values()]
    tracker, history = export_reports(model, presentation, source_labels, root)
    # This verifier only reads XLSX ZIP/XML; it does not author a workbook.
    from verify_reports import verify
    try:
        verify(SimpleNamespace(tracker=BytesIO(tracker), historical=BytesIO(history), model=json.dumps({"tracker": model["trackerStats"], "historical": model["historicalStats"]})))
    except (AssertionError, ValueError, KeyError) as error:
        raise ReportError(f"The generated reports failed verification: {error}") from error
    stamp = as_of.replace("-", "_")
    files = {
        f"Finance_Placement_Tracker_2027_{stamp}.xlsx": tracker,
        f"Finance_Placement_Stats_2024_2026_{stamp}.xlsx": history,
    }
    changes = tracker_changes(model["trackerRows"], old_rows, presentation["headers"]["tracker"]) if previous else None
    if changes is not None:
        files[f"Tracker_Changes_{stamp}.csv"] = changes_csv(changes["rows"])
    manifest = {
        "appVersion": "1.0.0", "formatVersion": presentation["formatVersion"], "snapshotDate": as_of,
        "scope": scope, "sourceFiles": [{"role": kind, "name": item["name"], "worksheet": item.get("sheet"), "headerRow": item.get("header_row", 1), "sha256": hashlib.sha256(item["data"]).hexdigest()} for kind, item in inputs.items()],
        "counts": {"tracker": len(model["trackerRows"]), "historical": len(model["historyRows"]), "review": len(model["reviewRows"]), "futureAdmitTerms": model["futureAdmitCount"]},
        "trackerStats": model["trackerStats"], "historicalStats": model["historicalStats"],
        "changes": {"rows": len(changes["rows"]), "previousStudentsOutsideCurrentScope": changes["previousStudentsOutsideCurrentScope"]} if changes is not None else None,
        "verification": {"fixedLayout": "passed", "scope": "passed", "uniqueStudents": "passed", "cachedFormulas": "passed", "totalsAndRates": "passed"},
    }
    files[f"Report_Manifest_{stamp}.json"] = json.dumps(manifest, indent=2).encode()
    bundle = BytesIO()
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return {"files": files, "bundle": bundle.getvalue(), "bundle_name": f"Finance_Reports_{stamp}.zip", "model": model, "manifest": manifest, "changes": changes}
