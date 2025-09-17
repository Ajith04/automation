import re
import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
import random

# ---------- CONFIG ----------
TARGET_SHEETS = ["AKUN", "WAMA", "GALAXEA"]
TARGET_RED = "FFC00000"    # ignore instructor if event cell red
HEADERS = ["Event", "Resource", "Configuration", "Date", "Start Time", "End Time"]
HEADER_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
HEADER_FONT = Font(bold=True)
YEAR_FOR_OUTPUT = 2025
NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

# ---------- HELPERS ----------
def safe_str(v):
    return "" if v is None else str(v).strip()

def get_rgb(cell):
    if cell is None:
        return None
    fill = getattr(cell, "fill", None)
    if not fill:
        return None
    color = getattr(fill, "start_color", None)
    if color is None:
        return None
    if hasattr(color, "rgb") and color.rgb:
        return str(color.rgb).upper()
    return None

def is_red(cell):
    return get_rgb(cell) == TARGET_RED

def parse_month_to_num(month_value):
    if month_value is None:
        return None
    s = str(month_value).strip()
    if not s:
        return None
    MONTH_MAP = {
        "january": 1, "jan": 1, "february": 2, "feb": 2,
        "march": 3, "mar": 3, "april": 4, "apr": 4,
        "may": 5, "june": 6, "jun": 6, "july": 7, "jul": 7,
        "august": 8, "aug": 8, "september": 9, "sep": 9, "sept": 9,
        "october": 10, "oct": 10, "november": 11, "nov": 11,
        "december": 12, "dec": 12
    }
    if s.isdigit() and 1 <= int(s) <= 12:
        return int(s)
    key = s.lower()
    if key in MONTH_MAP:
        return MONTH_MAP[key]
    for name, num in MONTH_MAP.items():
        if name in key:
            return num
    m = re.search(r"\b(1[0-2]|0?[1-9])\b", s)
    if m:
        return int(m.group(1))
    return None

def add_headers(ws):
    for col_idx, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL

def clean_instructor_name(name):
    if not name:
        return None
    return re.sub(r"\s*\(.*?\)\s*", "", str(name)).strip()

def get_light_fill():
    colors = [
        "FFFFE5CC", "FFE5FFCC", "FFCCFFE5", "FFCCE5FF",
        "FFFFCCFF", "FFE5CCFF", "FFFFCCCC", "FFCCFFFF"
    ]
    hexcolor = random.choice(colors)
    return PatternFill(start_color=hexcolor, end_color=hexcolor, fill_type="solid")

# ---------- RAW XML HELPERS ----------
def get_data_validations(xlsx_file):
    """Parse all data validation formulas from workbook XML."""
    results = {}
    with zipfile.ZipFile(xlsx_file, "r") as z:
        for sheet_file in [f for f in z.namelist() if f.startswith("xl/worksheets/sheet")]:
            xml_content = z.read(sheet_file)
            root = ET.fromstring(xml_content)
            dvs = root.findall(".//main:dataValidation", NS)
            formulas = []
            for dv in dvs:
                formula = dv.find("main:formula1", NS)
                if formula is not None:
                    formulas.append(formula.text)  # e.g. "A,B,C" or "=Availability!B2:B20"
            results[sheet_file] = formulas
    return results

def read_range_from_sheet(xlsx_file, sheet_name, cell_range):
    """Read values from a given sheet + range (like B2:B20)"""
    values = []
    with zipfile.ZipFile(xlsx_file, "r") as z:
        sheet_file = f"xl/worksheets/{sheet_name}.xml"
        xml_content = z.read(sheet_file)
        root = ET.fromstring(xml_content)
        for c in root.findall(".//main:c", NS):
            ref = c.attrib.get("r")  # e.g., "B2"
            if not ref:
                continue
            match = re.match(r"([A-Z]+)(\d+)", ref)
            if not match:
                continue
            col, row = match.groups()
            row = int(row)
            # crude check: you can expand logic for ranges
            if ref in cell_range or (col + str(row)) in cell_range:
                v = c.find("main:v", NS)
                if v is not None:
                    values.append(v.text)
    return values

def resolve_dropdown_values(xlsx_file, formula):
    """Resolve dropdown values from inline or cross-sheet references."""
    values = []
    if not formula:
        return values
    if formula.startswith('"') and formula.endswith('"'):
        # Inline list
        values = [x.strip() for x in formula.strip('"').split(",")]
    elif formula.startswith("="):
        # Cross-sheet reference like =Availability!B2:B20
        if "!" in formula:
            sheet_name, rng = formula[1:].split("!")
            sheet_name = sheet_name.strip("'")  # handle quotes
            values = read_range_from_sheet(xlsx_file, sheet_name, rng)
    return values

# ---------- STAFF ----------
def preload_staff(staff_file):
    wb = load_workbook(staff_file, data_only=True)
    result = {}
    for sheet in TARGET_SHEETS:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        headers = [safe_str(c.value).lower() for c in ws[1]]
        try:
            pri_idx = headers.index("priority") + 1
        except ValueError:
            pri_idx = 1
        instr_start_col = pri_idx + 1
        sheet_map = {}
        for col in range(instr_start_col, ws.max_column + 1):
            instr_name = clean_instructor_name(safe_str(ws.cell(1, col).value))
            if not instr_name:
                continue
            r = 2
            while r <= ws.max_row:
                val_cell = ws.cell(r, col)
                val = safe_str(val_cell.value)
                if not val:
                    break
                if is_red(val_cell):
                    r += 1
                    continue
                sheet_map.setdefault(val, []).append(instr_name)
                r += 1
        result[sheet] = sheet_map
    return result

# ---------- MAIN ----------
def generate_output(events_file, staff_file, output_file):
    instructors_map = preload_staff(staff_file)
    wb_src = load_workbook(events_file, data_only=True)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    event_color_cache = {}

    # Collect all dropdowns from raw XML
    dropdowns_by_sheet = get_data_validations(events_file)

    for sheet_name in wb_src.sheetnames:
        if sheet_name.upper() not in TARGET_SHEETS:
            continue
        ws_src = wb_src[sheet_name]
        ws_out = wb_out.create_sheet(sheet_name)
        add_headers(ws_out)

        header_map = {safe_str(ws_src.cell(1, c).value).lower(): c
                      for c in range(1, ws_src.max_column + 1)}
        resort_col = header_map.get("resort name")
        activity_col = header_map.get("activity")
        bookable_col = header_map.get("bookable hours")
        month_col = header_map.get("month")
        duration_col = header_map.get("activity duration")

        if not (month_col and activity_col and bookable_col):
            continue

        date_start_col = month_col + 1
        max_check_col = min(ws_src.max_column, date_start_col + 31 - 1)
        out_row = 2

        # detect multi-resort activities
        activity_resorts = {}
        for r in range(2, ws_src.max_row + 1):
            act = safe_str(ws_src.cell(r, activity_col).value)
            res = safe_str(ws_src.cell(r, resort_col).value) if resort_col else ""
            if act:
                activity_resorts.setdefault(act, set()).add(res)

        seen_events = set()
        for r in range(2, ws_src.max_row + 1):
            activity = safe_str(ws_src.cell(r, activity_col).value)
            if not activity:
                continue
            resort = safe_str(ws_src.cell(r, resort_col).value) if resort_col else ""
            duration = safe_str(ws_src.cell(r, duration_col).value) if duration_col else ""
            month_val = ws_src.cell(r, month_col).value

            if sheet_name.upper() == "GALAXEA" and "day" in duration.lower():
                continue

            # first non-empty day
            first_day = None
            for col in range(date_start_col, max_check_col + 1):
                c = ws_src.cell(r, col)
                try:
                    if c.value is not None and float(c.value) > 0:
                        first_day = int(ws_src.cell(1, col).value)
                        break
                except:
                    continue
            if not first_day:
                continue

            month_num = parse_month_to_num(month_val)
            if not month_num:
                continue
            date_str = f"{first_day:02d}/{month_num:02d}/{YEAR_FOR_OUTPUT}"

            resorts_for_activity = activity_resorts.get(activity, set())
            if len(resorts_for_activity) > 1 and resort:
                event_name = f"{activity} - {resort}"
            else:
                event_name = activity

            instrs = instructors_map.get(sheet_name, {}).get(activity, [])

            if event_name not in event_color_cache:
                event_color_cache[event_name] = get_light_fill()
            fill = event_color_cache[event_name]

            # fetch Bookable Hours dropdown
            dropdown_formulas = dropdowns_by_sheet.get(f"xl/worksheets/{sheet_name}.xml", [])
            time_slots = []
            for formula in dropdown_formulas:
                time_slots.extend(resolve_dropdown_values(events_file, formula))

            for slot in time_slots:
                slot = slot.strip()
                if not slot:
                    continue
                if "-" in slot:
                    start_time, end_time = [p.strip() for p in slot.split("-", 1)]
                else:
                    start_time, end_time = slot, ""

                key = (event_name, resort, activity, date_str, start_time, end_time)
                if key in seen_events:
                    continue
                seen_events.add(key)

                # main row
                for col_idx, val in enumerate([event_name, activity, activity, date_str, start_time, end_time], start=1):
                    ws_out.cell(out_row, col_idx, value=val).fill = fill
                out_row += 1

                # instructor rows
                for instr in instrs:
                    for col_idx, val in enumerate([event_name, instr, activity, date_str, start_time, end_time], start=1):
                        ws_out.cell(out_row, col_idx, value=val).fill = fill
                    out_row += 1

    wb_out.save(output_file)
    print(f"✅ Output saved to {output_file}")
