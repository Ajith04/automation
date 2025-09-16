import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import range_boundaries
from datetime import datetime
import random

# ---------- CONFIG ----------
TARGET_SHEETS = ["AKUN", "WAMA", "GALAXEA"]
TARGET_RED = "FFC00000"    # ignore instructor if event cell red
HEADERS = ["Event", "Resource", "Configuration", "Date", "Start Time", "End Time"]
HEADER_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
HEADER_FONT = Font(bold=True)
YEAR_FOR_OUTPUT = 2025

MONTH_MAP = {
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12
}

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
    if s.isdigit() and 1 <= int(s) <= 12:
        return int(s)
    key = s.lower()
    if key in MONTH_MAP:
        return MONTH_MAP[key]
    for name, num in MONTH_MAP.items():
        if name in key:
            return num
    m = re.search(r"\\b(1[0-2]|0?[1-9])\\b", s)
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
    return re.sub(r"\\s*\\(.*?\\)\\s*", "", str(name)).strip()


def get_light_fill():
    colors = [
        "FFFFE5CC", "FFE5FFCC", "FFCCFFE5", "FFCCE5FF",
        "FFFFCCFF", "FFE5CCFF", "FFFFCCCC", "FFCCFFFF"
    ]
    hexcolor = random.choice(colors)
    return PatternFill(start_color=hexcolor, end_color=hexcolor, fill_type="solid")


# ---------- BOOKABLE HOURS EXTRACTION ----------
def extract_bookable_options(ws, bookable_col):
    """Extract dropdown options for the Bookable Hours column by inspecting
    the worksheet's data validations. Returns a list of option strings.

    This handles:
      - Inline lists in formula1 (e.g. "\"08:00 - 10:00,17:00 - 19:00\"")
      - Range references to another sheet (e.g. =BookableHours!$A$1:$A$10)
    """
    options = []

    dv_container = getattr(ws, "data_validations", None)
    if not dv_container:
        return []

    dv_list = getattr(dv_container, "dataValidation", []) or []

    for dv in dv_list:
        # dv.sqref is the cell/range(s) this validation applies to
        if not getattr(dv, "sqref", None):
            continue
        sqref_str = str(dv.sqref)
        # sqref may contain multiple ranges separated by spaces
        for part in sqref_str.split():
            try:
                min_col, min_row, max_col, max_row = range_boundaries(part)
            except Exception:
                # not a range we can parse
                continue
            # if the Bookable Hours column falls into this dv range, process formula
            if bookable_col < min_col or bookable_col > max_col:
                continue

            f = getattr(dv, "formula1", None)
            if not f:
                continue
            fstr = str(f).strip()

            # inline list like "08:00 - 10:00,17:00 - 19:00"
            if fstr.startswith('"') and fstr.endswith('"'):
                raw = fstr.strip('"')
                for opt in [o.strip() for o in raw.split(',') if o.strip()]:
                    options.append(opt)
                continue

            # otherwise it's likely a range reference like =SheetName!$A$1:$A$10
            formula = fstr.lstrip('=')
            if '!' in formula:
                sheet_name, cell_range = formula.split('!', 1)
                sheet_name = sheet_name.strip("'")
                try:
                    rng = ws.parent[sheet_name][cell_range]
                except Exception:
                    # if we cannot resolve the range, skip
                    continue
                for row in rng:
                    for cell in row:
                        if cell.value is not None and str(cell.value).strip():
                            options.append(str(cell.value).strip())
                continue

            # fallback: try to resolve named ranges in the workbook
            try:
                named = ws.parent.defined_names.get(formula)
            except Exception:
                named = None
            if named:
                # defined_names[formula].destinations yields (sheetname, coord)
                try:
                    for title, coord in ws.parent.defined_names[formula].destinations:
                        try:
                            rng = ws.parent[title][coord]
                        except Exception:
                            continue
                        for row in rng:
                            for cell in row:
                                if cell.value is not None and str(cell.value).strip():
                                    options.append(str(cell.value).strip())
                except Exception:
                    pass

    # preserve order and remove duplicates
    seen = set()
    unique = []
    for o in options:
        if o not in seen:
            seen.add(o)
            unique.append(o)
    return unique


# ---------- STAFF LOADER ----------
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
            instr_name = clean_instructor_name(safe_str(ws.cell(row=1, column=col).value))
            if not instr_name:
                continue
            r = 2
            while r <= ws.max_row:
                val_cell = ws.cell(row=r, column=col)
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
    # load workbook WITHOUT data_only so we can inspect data validation formulas
    wb_src = load_workbook(events_file, data_only=False)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    event_color_cache = {}
    seen_events = set()

    for sheet_name in wb_src.sheetnames:
        if sheet_name.upper() not in TARGET_SHEETS:
            continue
        ws_src = wb_src[sheet_name]
        ws_out = wb_out.create_sheet(sheet_name)
        add_headers(ws_out)

        header_map = {}
        for c in range(1, ws_src.max_column + 1):
            val = ws_src.cell(row=1, column=c).value
            if val:
                header_map[safe_str(val).lower()] = c

        resort_col = header_map.get("resort name")
        activity_col = header_map.get("activity")
        duration_col = header_map.get("activity duration")
        month_col = header_map.get("month")
        bookable_col = header_map.get("bookable hours")

        if month_col is None or activity_col is None or bookable_col is None:
            continue

        date_start_col = month_col + 1
        max_check_col = min(ws_src.max_column, date_start_col + 31 - 1)

        out_row = 2

        for r in range(2, ws_src.max_row + 1):
            activity = safe_str(ws_src.cell(row=r, column=activity_col).value)
            if not activity:
                continue
            resort = safe_str(ws_src.cell(row=r, column=resort_col).value) if resort_col else ""
            duration = safe_str(ws_src.cell(row=r, column=duration_col).value) if duration_col else ""
            month_val = ws_src.cell(row=r, column=month_col).value

            if sheet_name.upper() == "GALAXEA" and "day" in duration.lower():
                continue

            first_day = None
            for col in range(date_start_col, max_check_col + 1):
                c = ws_src.cell(row=r, column=col)
                try:
                    if c.value is not None and float(c.value) > 0:
                        hdr = ws_src.cell(row=1, column=col).value
                        first_day = int(hdr)
                        break
                except:
                    continue
            if not first_day:
                continue

            month_num = parse_month_to_num(month_val)
            if not month_num:
                continue
            date_str = f"{first_day:02d}/{month_num:02d}/{YEAR_FOR_OUTPUT}"

            event_name = f"{activity} - {resort}" if resort else activity

            if (sheet_name, event_name) in seen_events:
                continue
            seen_events.add((sheet_name, event_name))

            instrs = instructors_map.get(sheet_name, {}).get(activity, [])

            if event_name not in event_color_cache:
                event_color_cache[event_name] = get_light_fill()
            fill = event_color_cache[event_name]

            # extract all slots from the Bookable Hours dropdown range referenced by the
            # data validation for this column
            slots = extract_bookable_options(ws_src, bookable_col)
            if not slots:
                print(f"⚠️ No bookable slots found for event '{event_name}' (sheet {sheet_name}) row {r}")
                continue

            for slot in slots:
                if "-" in slot:
                    start, end = [p.strip() for p in slot.split("-", 1)]
                else:
                    # defensive: if a cell in the bookable range isn't in start-end form,
                    # skip it
                    print(f"⚠️ Skipping malformed bookable slot for '{event_name}': '{slot}'")
                    continue

                for col_idx, val in enumerate([event_name, activity, activity, date_str, start, end], start=1):
                    ws_out.cell(row=out_row, column=col_idx, value=val).fill = fill
                out_row += 1

                for instr in instrs:
                    for col_idx, val in enumerate([event_name, instr, activity, date_str, start, end], start=1):
                        ws_out.cell(row=out_row, column=col_idx, value=val).fill = fill
                    out_row += 1

    wb_out.save(output_file)
    print(f"✅ Output saved to {output_file}")