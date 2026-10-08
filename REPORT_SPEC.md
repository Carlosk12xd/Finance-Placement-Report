# Finance reporting contract — format 1.0.0

## Purpose and precedence

Produce the Finance placement tracker requested by Amy Densley and the historical placement statistics refined in this conversation. The latest reporting choice is to foreground students placed and placement rates. Same-employer offer and acceptance counts remain a small supporting section. Same-company rates and the large employer comparison audit sheets are excluded.

Use this contract and the packaged code for routine refreshes. Do not redesign a workbook because a new chat lacks prior conversation memory. New user instructions can intentionally change this contract; record such changes in a new version.

## Population

The current tracker contains only Students records with `Major = BSFin`, an `Expected Graduation Semester` of **Fall 2026 (December), Winter 2027 (April), or Spring 2027 (June)**, and no recorded Actual Graduation Semester. All other graduation cohorts are excluded from tracker, review, and named change rows.

The historical report contains only `Major = BSFin` and academic `Class Of` **2024, 2025, or 2026**. Use the CRM academic class, including Fall graduates in the preceding calendar year. Missing Actual Graduation Semester remains in the assigned academic class and is flagged.

Neither Class Of nor a missing actual graduation establishes current enrollment. The exports do not supply a verified enrollment field. Retain the enrollment note and flag Start Semester years later than the snapshot year. Phone is the available number; mobile type is unverified. Start Semester is the available Finance admit-semester proxy.

## Source fields

| Export | Required fields |
|---|---|
| Students | Record Id; First Name; Last Name; netID; Major; Class Of; Job Search Status; Actual Graduation Semester; Expected Graduation Semester; Start Semester; Email; Phone |
| Internship Placements | Record Id; Student ID.id; Company.id; Company; Start Date |
| Full-Time Placements | Record Id; Student ID.id or StudentID.id; Company.id; Company; Offer Date; Accepted?; Accepted Date |

Placement dates must be blank or valid ISO `YYYY-MM-DD` values. Accepted? must be true, false, or blank, ignoring capitalization. Source row numbers are used only to explain input errors.

Join `Students.Record Id` to internship `Student ID.id` and full-time `Student ID.id`. Use the full-time `StudentID.id` alias when necessary; populated aliases must agree. Never join on a name. Reject missing/duplicate record IDs, missing/duplicate scoped NetIDs, and unmatched placement student IDs so ambiguous joins do not silently change results.

## Statuses and placement statistics

| Source Job Search Status | Report Placement Status |
|---|---|
| Accepted an offer / Accepted Offer | Accepted Offer |
| Actively seeking / Actively Seeking | Actively Seeking |
| Every Not seeking reason / Not Seeking | Not Seeking |
| Not Reported / blank / no status | no status |

Unfamiliar source statuses require correction or an intentional mapping update.

- **Students placed**: unique students with Accepted Offer status.
- **Placement rate**: students placed divided by every student in the requested graduation term or academic class. Not Seeking and no status remain in the denominator.
- **Change vs. prior class**: current class placement rate minus preceding reported class placement rate, in percentage points.
- **Internship recorded / coverage**: unique students with at least one eligible internship record / all students in the term or class.
- **No internship reported**: all cohort students minus students with an internship record.
- **FT offer recorded**: unique students with an eligible full-time offer record, including declined offers.
- **Accepted FT record**: unique students with an eligible record marked Accepted? true.
- **Accepted employer unreported**: Accepted Offer status without a known accepted employer by the snapshot.
- **Missing email, phone, actual graduation**: blank source fields, counted once per student.
- A zero denominator displays **—**, rather than an invented percentage.

Placement status and accepted offer records are separate source facts. Preserve discrepancies in the review information. Absence of a full-time record does not prove that a student received no offer.

## Ten tracker columns, in order

1. NetID
2. First name
3. Last name
4. Placement Status
5. Full-time placement Company
6. Most-recent internship Company
7. Expected Graduation date, using the semester label
8. Email Address
9. Cell phone #, using the source Phone value
10. Admit semester, using the source Start Semester value

Show accepted full-time employers only. Exclude offers dated after the snapshot and accepted records with future acceptance dates. Undated records remain reported source records and are flagged when both offer and acceptance dates are missing. If Accepted Offer status lacks a reported employer, display `Not reported / verify`. Multiple accepted employers remain visible and are flagged.

For the tracker internship company, choose the greatest known Start Date no later than the snapshot. Show all employers tied on that date and flag the ambiguity. If only undated internships exist, show their employer names and flag that the newest internship cannot be established. With both dated and undated records, show the latest dated internship and flag the missing dates. A recorded internship with no company name displays `Employer not reported` and receives a review flag.

## Supporting internship-employer counts

Compare any eligible recorded internship employer with any eligible recorded full-time offer employer. Prefer Company.id; when either ID is missing, allow only exact normalized company-name equality. If both IDs are present and differ, names do not override them. No fuzzy matching.

Count each student once across multiple internships or offers. Count declined offers in students offered FT by an internship employer. Accepted counts require a matching accepted offer with no known future acceptance date. Missing or reversed internship/offer dates remain in the employer-offer counts and are flagged where the matching date sequence cannot be supported. An employer match does not prove the offer was a direct internship return offer.

## Presentation

Tracker worksheet order: Placement Summary, Placement Tracker, Review. Historical worksheet order: Placement Summary, Student Detail. Do not add dashboards, audit comparison tables, or conversion-rate charts during routine refreshes.

Use Arial, navy headers, white header text, green for accepted placement, amber for seeking/review, gray for Not Seeking, red for no status, and consistent term colors. Keep the approved widths, fonts, fixed header rows, summary positions, filters, validation lists, and freeze panes. Allow body row heights and table ranges to expand with data. ID/contact strings remain text. Percentages display one decimal place.

## Reproducibility

The manifest records the export date, input hashes, scope, aggregate counts, and code/format hashes. Validate workbook formulas against source-derived totals, then validate exported cells and the approved layout. Rendering is the final visual check in ChatGPT Work. Do not relabel a previously generated report with a new date when fresh matching exports are unavailable. Do not send messages, upload records to a database, or modify the shared working tracker as part of this generator.
