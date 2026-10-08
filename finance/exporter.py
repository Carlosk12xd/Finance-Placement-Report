"""Fill approved, Artifact-Tool-authored XLSX templates with standard-library XML.

The template owns fonts, colors, dimensions, tables, validation and sheet layout.
Only data rows, dates, source labels, ranges, formulas and their caches change.
No proprietary spreadsheet runtime is needed by the hosted app.
"""
from copy import deepcopy
from datetime import date
from io import BytesIO
from pathlib import Path
import math
import re
import xml.etree.ElementTree as ET
import zipfile

from .logic import ReportError

NS_URI = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"x": NS_URI}
ET.register_namespace("", NS_URI)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")


def tag(name):
    return f"{{{NS_URI}}}{name}"


def letter(number):
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def column_index(address):
    value = 0
    for c in re.match(r"[A-Z]+", address).group():
        value = value * 26 + ord(c) - 64
    return value


def checked_text(value):
    text = str(value)
    if len(text) > 32767 or any(ord(c) < 32 and c not in "\n\r\t" for c in text):
        raise ReportError("A source value contains unsupported Excel characters or is longer than one cell permits.")
    return text


class Template:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.parts = {name: z.read(name) for name in z.namelist()}
        self.sheets = [ET.fromstring(self.parts[f"xl/worksheets/sheet{i+1}.xml"]) for i in range(3 if "tracker" in Path(path).name else 2)]

    def row(self, sheet, number):
        data = self.sheets[sheet].find("x:sheetData", NS)
        for row in data:
            if int(row.attrib["r"]) == number:
                return row
        row = ET.SubElement(data, tag("row"), {"r": str(number)})
        return row

    def cell(self, sheet, address):
        row = self.row(sheet, int(re.search(r"\d+", address).group()))
        for cell in row:
            if cell.attrib["r"] == address:
                return cell
        return ET.SubElement(row, tag("c"), {"r": address})

    def set(self, sheet, address, value, keep_formula=False):
        cell = self.cell(sheet, address)
        formula = deepcopy(cell.find("x:f", NS)) if keep_formula and cell.find("x:f", NS) is not None else None
        for child in list(cell):
            cell.remove(child)
        cell.attrib.pop("t", None)
        if formula is not None:
            cell.append(formula)
            if isinstance(value, str):
                cell.attrib["t"] = "str"
            ET.SubElement(cell, tag("v")).text = checked_text(value)
        elif value is None or value == "":
            pass
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            ET.SubElement(cell, tag("v")).text = str(value)
        else:
            cell.attrib["t"] = "inlineStr"
            inline = ET.SubElement(cell, tag("is"))
            ET.SubElement(inline, tag("t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = checked_text(value)

    def body(self, sheet, first, rows, widths, check_columns, minimum=28, date_column=None):
        root = self.sheets[sheet]
        data = root.find("x:sheetData", NS)
        base = self.row(sheet, first)
        styles = {column_index(c.attrib["r"]): c.attrib.get("s", "0") for c in base if column_index(c.attrib["r"]) <= len(widths)}
        for row in list(data):
            if int(row.attrib["r"]) >= first:
                for cell in list(row):
                    if column_index(cell.attrib["r"]) <= len(widths):
                        row.remove(cell)
                if not len(row):
                    data.remove(row)
        display = rows or [[None] * len(widths)]
        last = first + len(display) - 1
        for index, values in enumerate(display):
            number = first + index
            row = self.row(sheet, number)
            lines = max([1] + [math.ceil(len(str(values[c] or "")) / max(12, math.floor((widths[c] - 14) / 7.4))) for c in check_columns])
            row.attrib.update(ht=str(max(minimum, lines * 14 + 10)), customHeight="1")
            for c, value in enumerate(values, start=1):
                address = f"{letter(c)}{number}"
                cell = self.cell(sheet, address)
                cell.attrib["s"] = styles.get(c, "0")
                if date_column == c - 1 and value:
                    value = (date.fromisoformat(value) - date(1899, 12, 30)).days
                self.set(sheet, address, value)
        for node in root.findall("x:conditionalFormatting", NS) + root.findall("x:dataValidations/x:dataValidation", NS):
            ref = node.attrib.get("sqref", "")
            node.attrib["sqref"] = re.sub(rf"([A-Z]+){first}:([A-Z]+)\d+", lambda m: f"{m[1]}{first}:{m[2]}{last}", ref)
        return last

    def future_admits(self, terms, last):
        root = self.sheets[1]
        example = None
        destination = None
        for cf in root.findall("x:conditionalFormatting", NS):
            if cf.attrib["sqref"] == f"J9:J{last}":
                destination = cf
                for rule in list(cf):
                    formula = rule.find("x:formula", NS)
                    if formula is not None and (formula.text or "").startswith("$J9="):
                        example = deepcopy(rule)
                        cf.remove(rule)
        if terms and example is None:
            raise ReportError("The template is missing its admit-term style.")
        if destination is not None:
            for index, term in enumerate(terms):
                rule = deepcopy(example)
                rule.attrib["priority"] = str(100 + index)
                escaped = checked_text(term).replace('"', '""')
                rule.find("x:formula", NS).text = f'$J9="{escaped}"'
                destination.append(rule)

    def table_range(self, name, last):
        for part in list(self.parts):
            if part.startswith("xl/tables/") and part.endswith(".xml"):
                table = ET.fromstring(self.parts[part])
                if table.attrib["name"] == name:
                    original = table.attrib["ref"]
                    ref = re.sub(r"\d+$", str(last), original)
                    table.attrib["ref"] = ref
                    table.find("x:autoFilter", NS).attrib["ref"] = ref
                    self.parts[part] = ET.tostring(table, encoding="utf-8", xml_declaration=True)

    def formula_ranges(self, name, first, last):
        pattern = re.compile(rf"('{re.escape(name)}'!\$([A-Z]+)\${first}:\$\2\$)\d+")
        for formula in self.sheets[0].findall(".//x:f", NS):
            formula.text = pattern.sub(lambda m: m[1] + str(last), formula.text or "")

    def finish(self):
        for i, root in enumerate(self.sheets):
            data = root.find("x:sheetData", NS)
            data[:] = sorted(data, key=lambda r: int(r.attrib["r"]))
            for row in data:
                row[:] = sorted(row, key=lambda c: column_index(c.attrib["r"]))
            cells = list(root.findall("x:sheetData/x:row/x:c", NS))
            max_col = max(column_index(c.attrib["r"]) for c in cells)
            max_row = max(int(r.attrib["r"]) for r in data)
            dimension = root.find("x:dimension", NS)
            if dimension is not None:
                dimension.attrib["ref"] = f"A1:{letter(max_col)}{max_row}"
            self.parts[f"xl/worksheets/sheet{i+1}.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        workbook = ET.fromstring(self.parts["xl/workbook.xml"])
        calc = workbook.find("x:calcPr", NS)
        if calc is None:
            calc = ET.SubElement(workbook, tag("calcPr"))
        calc.attrib.update(calcMode="auto", fullCalcOnLoad="1", forceFullCalc="1")
        self.parts["xl/workbook.xml"] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)
        output = BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, contents in self.parts.items():
                archive.writestr(name, contents)
        return output.getvalue()


def ratio(numerator, denominator):
    return numerator / denominator if denominator else "—"


def summary(book, stats, historical=False):
    previous_rate = None
    for number, s in enumerate(stats, start=7):
        rate = ratio(s["placed"], s["students"])
        values = [s["year"] if historical else s["term"], s["students"], s["placed"], rate, s["activelySeeking"], s["notSeeking"], s["noStatus"], ((rate - previous_rate) * 100 if isinstance(rate, (int, float)) and isinstance(previous_rate, (int, float)) else "—") if historical else s["internships"]]
        for c, value in enumerate(values, start=1):
            book.set(0, f"{letter(c)}{number}", value, keep_formula=True)
        previous_rate = rate
    students = sum(s["students"] for s in stats)
    placed = sum(s["placed"] for s in stats)
    total = ["All three classes" if historical else "All requested students", students, placed, ratio(placed, students), sum(s["activelySeeking"] for s in stats), sum(s["notSeeking"] for s in stats), sum(s["noStatus"] for s in stats)]
    if not historical:
        total.append(sum(s["internships"] for s in stats))
    for c, value in enumerate(total, start=1):
        book.set(0, f"{letter(c)}10", value, keep_formula=True)
    coverage = []
    for number, s in enumerate(stats, start=16):
        values = [s["year"] if historical else s["term"], s["internships"], ratio(s["internships"], s["students"]), s["students"] - s["internships"]]
        values += [s["offers"], s["acceptedRecords"], s["acceptedEmployerUnreported"], s["actualGraduationMissing"]] if historical else [s["acceptedCompany"], s["acceptedEmployerUnreported"], s["emailMissing"], s["phoneMissing"]]
        coverage.append(values)
        for c, value in enumerate(values, start=1):
            book.set(0, f"{letter(c)}{number}", value, keep_formula=True)
    if historical:
        for number, s in enumerate(stats, start=26):
            for c, value in enumerate([s["year"], s["internshipEmployerOffer"], s["internshipEmployerAccepted"]], start=1):
                book.set(0, f"{letter(c)}{number}", value, keep_formula=True)
    else:
        total = ["All requested students", sum(s["internships"] for s in stats), ratio(sum(s["internships"] for s in stats), students)]
        total += [sum(row[c] for row in coverage) for c in range(3, 8)]
        for c, value in enumerate(total, start=1):
            book.set(0, f"{letter(c)}19", value, keep_formula=True)


def export_reports(model, settings, source_labels, root):
    as_of = date.fromisoformat(model["asOf"])
    label = f"{as_of:%B} {as_of.day}, {as_of.year}"
    count = model["futureAdmitCount"]
    admit = f"{count} future admit {'term' if count == 1 else 'terms'}"
    tracker = Template(root / "assets/tracker_template.xlsx")
    history = Template(root / "assets/historical_template.xlsx")
    last_t = tracker.body(1, 9, model["trackerRows"], settings["widths"]["tracker"], [1, 2, 4, 5])
    last_r = tracker.body(2, 7, model["reviewRows"], settings["widths"]["review"], [4, 5], minimum=42, date_column=7)
    last_h = history.body(1, 7, model["historyRows"], settings["widths"]["historicalDetail"], [2, 3, 5, 6, 7, 11])
    for name, last in (("FinancePlacementTracker", last_t), ("FinancePlacementReview", last_r)):
        tracker.table_range(name, last)
    history.table_range("FinanceHistoricalStudents", last_h)
    tracker.formula_ranges("Placement Tracker", 9, last_t)
    history.formula_ranges("Student Detail", 7, last_h)
    tracker.future_admits(model["futureAdmitTerms"], last_t)
    tracker.set(0, "A3", f"Requested graduation terms only • As of {label}")
    tracker.set(1, "A3", f"{', '.join(model['scope']['trackerTermLabels'])} • As of {label}")
    tracker.set(1, "A5", f"Current enrollment is not available in the exports; verify {admit} on Review.")
    tracker.set(0, "A22", f"Enrollment is unverified; {admit} {'requires' if count == 1 else 'require'} confirmation on Review.")
    tracker.set(2, "A3", f"{len(model['reviewRows'])} students in the requested graduation terms • Review status and reviewed date are editable")
    classes = ", ".join(map(str, model["scope"]["historicalClasses"]))
    history.set(0, "A3", f"Classes of {classes} only • As of {label}")
    history.set(1, "A3", f"Classes of {classes} only • One row per student • As of {label}")
    for row, source in enumerate(source_labels, start=9):
        tracker.set(1, f"L{row}", source)
    for row, source in enumerate(source_labels, start=7):
        history.set(1, f"N{row}", source)
    summary(tracker, model["trackerStats"])
    summary(history, model["historicalStats"], historical=True)
    return tracker.finish(), history.finish()
