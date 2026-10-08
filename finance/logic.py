"""The approved format-1.0.0 data rules, ported from the saved JS generator."""
from collections import defaultdict
from datetime import date
import re
import unicodedata


class ReportError(ValueError):
    """A correctable upload or report validation issue."""


REQUIRED_COLUMNS = {
    "students": ["Record Id", "First Name", "Last Name", "netID", "Major", "Class Of", "Job Search Status", "Actual Graduation Semester", "Expected Graduation Semester", "Start Semester", "Email", "Phone"],
    "internships": ["Record Id", "Student ID.id", "Company.id", "Company", "Start Date"],
    "offers": ["Record Id", "Company.id", "Company", "Offer Date", "Accepted?", "Accepted Date"],
}
STATUSES = ["Actively Seeking", "Not Seeking", "Accepted Offer", "no status"]


def valid_date(value):
    try:
        return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", value or "")) and date.fromisoformat(value).isoformat() == value
    except (ValueError, TypeError):
        return False


def validate_columns(headers, kind):
    if len(set(headers)) != len(headers):
        raise ReportError(f"{kind}: duplicate column headings.")
    missing = [name for name in REQUIRED_COLUMNS[kind] if name not in headers]
    if kind == "offers" and not {"Student ID.id", "StudentID.id"}.intersection(headers):
        missing.append("Student ID.id or StudentID.id")
    if missing:
        raise ReportError(f"{kind}: required columns missing: {', '.join(missing)}. Upload the source export, not a placement report.")


def validate_scope(scope):
    if scope["major"] != "BSFin" or scope["trackerTerms"] != ["Fall 2026", "Winter 2027", "Spring 2027"] or scope["historicalClasses"] != [2024, 2025, 2026] or scope["excludeCompletedGraduatesFromTracker"] is not True:
        raise ReportError("This app uses the approved Finance graduation cohorts. Change the report contract intentionally before changing the scope.")


def placement_status(value):
    value = str(value or "").strip()
    if value in ("Accepted an offer", "Accepted Offer"):
        return "Accepted Offer"
    if value in ("Actively seeking", "Actively Seeking"):
        return "Actively Seeking"
    if value.startswith("Not seeking") or value == "Not Seeking":
        return "Not Seeking"
    if value in ("", "Not Reported", "no status"):
        return "no status"
    raise ReportError(f"Unrecognized placement status: {value!r}. Correct the source or intentionally update the status mapping.")


def offer_student_id(row):
    return row.get("Student ID.id") or row.get("StudentID.id") or ""


def normalized_name(value):
    value = unicodedata.normalize("NFKC", str(value or "")).lower().replace("&", " and ")
    return " ".join("".join(c if c.isalnum() else " " for c in value).split())


def english_key(value):
    value = unicodedata.normalize("NFKD", str(value or "")).casefold()
    return "".join(c for c in value if not unicodedata.combining(c))


def same_employer(internship, offer):
    first, second = internship["Company.id"], offer["Company.id"]
    if first and second:
        return first == second
    return bool(internship["Company"] and offer["Company"] and normalized_name(internship["Company"]) == normalized_name(offer["Company"]))


def company_list(rows):
    return "; ".join(sorted({row["Company"] for row in rows if row["Company"]}, key=english_key))


def reported_companies(rows):
    return (company_list(rows) or "Employer not reported") if rows else ""


def accepted_companies(rows):
    return (company_list(rows) or "Not reported / verify") if rows else ""


def future_admit(student, as_of):
    match = re.search(r"\d{4}$", student["Start Semester"])
    return bool(match and int(match.group()) > int(as_of[:4]))


def class_year(value):
    try:
        year = float(value)
        return int(year) if year.is_integer() else None
    except (ValueError, TypeError):
        return None


def validate_sources(sources):
    ids = set()
    for student in sources["students"]:
        record_id = student["Record Id"]
        if not record_id or record_id in ids:
            raise ReportError("Students: missing or duplicate Record Id.")
        ids.add(record_id)
    for kind in ("internships", "offers"):
        records = set()
        for row in sources[kind]:
            record_id = row["Record Id"]
            if not record_id or record_id in records:
                raise ReportError(f"{kind}: missing or duplicate placement Record Id.")
            records.add(record_id)
            if kind == "offers" and row.get("Student ID.id") and row.get("StudentID.id") and row["Student ID.id"] != row["StudentID.id"]:
                raise ReportError("Full-time student-ID aliases disagree.")
            student_id = offer_student_id(row) if kind == "offers" else row["Student ID.id"]
            if student_id not in ids:
                raise ReportError(f"{kind}: a placement has no matching Student Record Id. Use a complete companion Students export.")
            for field in (["Offer Date", "Accepted Date"] if kind == "offers" else ["Start Date"]):
                if row[field] and not valid_date(row[field]):
                    raise ReportError(f"{kind}: {field} must be a real date; source row {row.get('sourceRow', 'unknown')}.")
            if kind == "offers" and str(row["Accepted?"]).lower() not in ("", "true", "false"):
                raise ReportError("Accepted? must be true, false, or blank.")


def issue_list(detail, as_of, all_offers):
    issues, actions = [], []
    def add(issue, action):
        issues.append(issue)
        actions.append(action)
    if future_admit(detail["s"], as_of):
        add(f"Future admit semester: {detail['s']['Start Semester']}", "Confirm enrollment and Finance admit term.")
    if detail["undated"]:
        add("Internship date missing", "Confirm the latest internship and its start date.")
    if len({r["Company.id"] or r["Company"] for r in detail["latest"] if r["Company.id"] or r["Company"]}) > 1:
        add("Most-recent internship has multiple employers", "Confirm the latest internship employer.")
    if (detail["status"] == "Accepted Offer" and not detail["accepted"]) or (detail["accepted"] and not company_list(detail["accepted"])):
        add("Accepted employer unreported", "Confirm the accepted employer and update the offer record.")
    if any(r["Offer Date"] > as_of or r["Accepted Date"] > as_of for r in all_offers):
        add("Future offer date", "Confirm the offer and acceptance dates.")
    if any(not r["Offer Date"] and not r["Accepted Date"] for r in detail["accepted"]):
        add("Accepted offer dates missing", "Confirm the offer or acceptance date.")
    if detail["internships"] and not company_list(detail["internships"]):
        add("Internship employer unreported", "Confirm the internship employer.")
    if len({r["Company.id"] or r["Company"] for r in detail["accepted"] if r["Company.id"] or r["Company"]}) > 1:
        add("Multiple accepted employers", "Confirm the final employer.")
    if not detail["s"]["Phone"]:
        add("Phone missing", "Collect a current cell phone number.")
    return issues, actions


def analyze_sources(sources, scope, as_of, prior_review=None):
    validate_scope(scope)
    if not valid_date(as_of):
        raise ReportError("Choose the actual export date.")
    validate_sources(sources)
    prior_review = prior_review or {}
    internships_by_id, offers_by_id = defaultdict(list), defaultdict(list)
    for row in sources["internships"]:
        internships_by_id[row["Student ID.id"]].append(row)
    for row in sources["offers"]:
        offers_by_id[offer_student_id(row)].append(row)

    def detail(student):
        internships = [r for r in internships_by_id[student["Record Id"]] if not r["Start Date"] or r["Start Date"] <= as_of]
        offers = [r for r in offers_by_id[student["Record Id"]] if not r["Offer Date"] or r["Offer Date"] <= as_of]
        accepted = [r for r in offers if r["Accepted?"].lower() == "true" and (not r["Accepted Date"] or r["Accepted Date"] <= as_of)]
        dated = [r for r in internships if r["Start Date"]]
        latest_date = max((r["Start Date"] for r in dated), default=None)
        latest = [r for r in dated if r["Start Date"] == latest_date]
        undated = [r for r in internships if not r["Start Date"]]
        matches = [{"i": i, "o": o} for i in internships for o in offers if same_employer(i, o)]
        return dict(s=student, internships=internships, offers=offers, accepted=accepted, latest=latest, undated=undated, matches=matches, status=placement_status(student["Job Search Status"]))

    name_key = lambda d: (english_key(d["s"]["Last Name"]), english_key(d["s"]["First Name"]), english_key(d["s"]["netID"]))
    tracker_details = [detail(s) for s in sources["students"] if s["Major"] == scope["major"] and s["Expected Graduation Semester"] in scope["trackerTerms"] and not s["Actual Graduation Semester"]]
    history_details = [detail(s) for s in sources["students"] if s["Major"] == scope["major"] and class_year(s["Class Of"]) in scope["historicalClasses"]]
    tracker_details.sort(key=lambda d: (scope["trackerTerms"].index(d["s"]["Expected Graduation Semester"]), *name_key(d)))
    history_details.sort(key=lambda d: (class_year(d["s"]["Class Of"]), *name_key(d)))
    for kind, details in (("tracker", tracker_details), ("historical", history_details)):
        netids = [d["s"]["netID"].lower() for d in details]
        if not all(netids) or len(set(netids)) != len(netids):
            raise ReportError(f"{kind}: missing or duplicate NetID in the requested scope.")

    tracker_rows = []
    for d in tracker_details:
        s = d["s"]
        employer = accepted_companies(d["accepted"]) if d["accepted"] else ("Not reported / verify" if d["status"] == "Accepted Offer" else "")
        tracker_rows.append([s["netID"], s["First Name"], s["Last Name"], d["status"], employer, reported_companies(d["latest"] or d["undated"]), s["Expected Graduation Semester"], s["Email"], s["Phone"], s["Start Semester"]])
    history_rows = []
    for d in history_details:
        s, notes = d["s"], []
        if d["status"] == "Accepted Offer" and (not d["accepted"] or not company_list(d["accepted"])):
            notes.append("Accepted employer needs confirmation")
        if d["status"] != "Accepted Offer" and d["accepted"]:
            notes.append("Accepted record differs from placement status")
        if not s["Actual Graduation Semester"]:
            notes.append("Actual graduation unreported")
        if d["matches"] and not any(p["i"]["Start Date"] and p["o"]["Offer Date"] and p["i"]["Start Date"] <= p["o"]["Offer Date"] for p in d["matches"]):
            notes.append("Employer match dates need review")
        if any(not o["Offer Date"] and not o["Accepted Date"] for o in d["accepted"]):
            notes.append("Accepted offer dates missing")
        if d["internships"] and not company_list(d["internships"]):
            notes.append("Internship employer unreported")
        accepted_ids = {o["Record Id"] for o in d["accepted"]}
        history_rows.append([class_year(s["Class Of"]), s["netID"], s["First Name"], s["Last Name"], d["status"], accepted_companies(d["accepted"]), reported_companies(d["internships"]), reported_companies(d["offers"]), "Yes" if d["matches"] else "No", "Yes" if any(p["o"]["Record Id"] in accepted_ids for p in d["matches"]) else "No", s["Actual Graduation Semester"], "; ".join(notes)])

    review_rows = []
    for d in tracker_details:
        s = d["s"]
        issues, actions = issue_list(d, as_of, offers_by_id[s["Record Id"]])
        if not issues:
            continue
        text, old = "; ".join(issues), prior_review.get(s["netID"], {})
        preserve = old.get("issue") == text
        review_rows.append([s["netID"], s["First Name"], s["Last Name"], s["Expected Graduation Semester"], text, " ".join(actions), old.get("status") if preserve else "Pending", old.get("date") if preserve else None])

    def status_summary(details, key):
        return {key: details[0]["s"][key] if details else None, "students": len(details), "placed": sum(d["status"] == "Accepted Offer" for d in details), "activelySeeking": sum(d["status"] == "Actively Seeking" for d in details), "notSeeking": sum(d["status"] == "Not Seeking" for d in details), "noStatus": sum(d["status"] == "no status" for d in details)}
    tracker_stats = []
    for term in scope["trackerTerms"]:
        details = [d for d in tracker_details if d["s"]["Expected Graduation Semester"] == term]
        tracker_stats.append({"term": term, **status_summary(details, "Expected Graduation Semester"), "internships": sum(bool(d["internships"]) for d in details), "acceptedCompany": sum(bool(company_list(d["accepted"])) for d in details), "acceptedEmployerUnreported": sum(d["status"] == "Accepted Offer" and (not d["accepted"] or not company_list(d["accepted"])) for d in details), "emailMissing": sum(not d["s"]["Email"] for d in details), "phoneMissing": sum(not d["s"]["Phone"] for d in details)})
    historical_stats = []
    for year in scope["historicalClasses"]:
        details = [d for d in history_details if class_year(d["s"]["Class Of"]) == year]
        historical_stats.append({"year": year, **status_summary(details, "Class Of"), "internships": sum(bool(d["internships"]) for d in details), "offers": sum(bool(d["offers"]) for d in details), "acceptedRecords": sum(bool(d["accepted"]) for d in details), "acceptedEmployerUnreported": sum(d["status"] == "Accepted Offer" and (not d["accepted"] or not company_list(d["accepted"])) for d in details), "actualGraduationMissing": sum(not d["s"]["Actual Graduation Semester"] for d in details), "internshipEmployerOffer": sum(bool(d["matches"]) for d in details), "internshipEmployerAccepted": sum(any(p["o"]["Record Id"] in {o["Record Id"] for o in d["accepted"]} for p in d["matches"]) for d in details)})
    future = [d for d in tracker_details if future_admit(d["s"], as_of)]
    return dict(asOf=as_of, scope=scope, trackerDetails=tracker_details, historyDetails=history_details, trackerRows=tracker_rows, historyRows=history_rows, reviewRows=review_rows, trackerStats=tracker_stats, historicalStats=historical_stats, futureAdmitCount=len(future), futureAdmitTerms=list(dict.fromkeys(d["s"]["Start Semester"] for d in future)))


def tracker_changes(current, previous, headers):
    previous = {str(r[0]): r for r in previous if r[0]}
    changes = []
    for row in current:
        old = previous.get(row[0])
        if old is None:
            changes.append([row[0], row[1], row[2], row[6], "New student", "", "", "New in scope"])
        else:
            for c in range(1, len(headers)):
                if str(old[c] or "") != str(row[c] or ""):
                    changes.append([row[0], row[1], row[2], row[6], headers[c], old[c] or "", row[c] or "", "Field changed"])
    return {"rows": changes, "previousStudentsOutsideCurrentScope": len(set(previous) - {r[0] for r in current})}
