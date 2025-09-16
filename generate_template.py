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

# ---------- BOOKABLE HOURS ----------
def get_bookable_slots(ws, row, col):
    """Get timeslots for a specific cell from its dropdown validation."""
    slots = []

    for dv in ws.data_validations.dataValidation:
        ranges = getattr(dv.sqref, "ranges", [str(dv.sqref)])  
        for ref in ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(ref))
            if not (min_col <= col <= max_col and min_row <= row <= max_row):
                continue

            formula = dv.formula1
            if not formula:
                continue
            formula = formula.strip()
            # inline list
            if formula.startswith('"') and formula.endswith('"'):
                items = formula.strip('"').split(",")
                slots.extend([safe_str(it) for it in items if safe_str(it)])
            # range or named range
            else:
                formula = formula.lstrip("=")
                if "!" in formula:
                    sheet_name, rng = formula.split("!")
                    sheet_name = sheet_name.strip("'")
                    if sheet_name in ws.parent.sheetnames:
                        ws_target = ws.parent[sheet_name]
                        minc, minr, maxc, maxr = range_boundaries(rng)
                        for rr in range(minr, maxr + 1):
                            val = safe_str(ws_target.cell(row=rr, column=minc).value)
                            if val:
                                slots.append(val)
                else:
                    if formula in ws.parent.defined_names:
                        dn = ws.parent.defined_names[formula]
                        for dn_dest in dn.destinations:
                            sheet_name, cell_range = dn_dest
                            ws_target = ws.parent[sheet_name]
                            minc, minr, maxc, maxr = range_boundaries(cell_range)
                            for rr in range(minr, maxr + 1):
                                val = safe_str(ws_target.cell(row=rr, column=minc).value)
                                if val:
                                    slots.append(val)
    return slots

# ---------- MAIN ----------
def generate_output(events_file, staff_file, output_file):
    instructors_map = preload_staff(staff_file)
    wb_src = load_workbook(events_file, data_only=False)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    event_color_cache = {}
    seen_events = set()

    for sheet_name in wb_src.sheetnames:
        if sheet_name.strip().upper() not in TARGET_SHEETS:
            continue
        ws_src = wb_src[sheet_name]
        ws_out = wb_out.create_sheet(sheet_name)
        add_headers(ws_out)

        # map headers
        header_map = {safe_str(c.value).lower().strip(): c for c in ws_src[1]}
        resort_col = header_map.get("resort name")
        activity_col = header_map.get("activity")
        duration_col = header_map.get("activity duration")
        bookable_col = header_map.get("bookable hours")
        month_col = header_map.get("month")

        # convert to column numbers
        resort_col = resort_col.col_idx if resort_col else None
        activity_col = activity_col.col_idx if activity_col else None
        duration_col = duration_col.col_idx if duration_col else None
        bookable_col = bookable_col.col_idx if bookable_col else None
        month_col = month_col.col_idx if month_col else None

        if month_col is None or activity_col is None or bookable_col is None:
            print(f"Skipping sheet {sheet_name}: missing required headers")
            continue

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

            month_num = parse_month_to_num(month_val)
            if not month_num:
                print(f"Skipping row {r} in {sheet_name}: month '{month_val}' invalid")
                continue

            date_str = f"01/{month_num:02d}/{YEAR_FOR_OUTPUT}"
            event_name = f"{activity} - {resort}" if resort else activity

            if (sheet_name, event_name) in seen_events:
                continue
            seen_events.add((sheet_name, event_name))

            instrs = instructors_map.get(sheet_name, {}).get(activity, [])

            if event_name not in event_color_cache:
                event_color_cache[event_name] = get_light_fill()
            fill = event_color_cache[event_name]

            # --- Bookable Hours slots ---
            slots = get_bookable_slots(ws_src, r, bookable_col)
            if not slots:
                print(f"No Bookable Hours for row {r} in {sheet_name}")
                continue

            for slot in slots:
                if "-" in slot:
                    parts = [p.strip() for p in slot.split("-", 1)]
                    start, end = parts if len(parts) == 2 else (parts[0], "")
                else:
                    start, end = slot, ""

                # event row
                for col_idx, val in enumerate([event_name, activity, activity, date_str, start, end], start=1):
                    ws_out.cell(row=out_row, column=col_idx, value=val).fill = fill
                out_row += 1

                # instructor rows
                for instr in instrs:
                    for col_idx, val in enumerate([event_name, instr, activity, date_str, start, end], start=1):
                        ws_out.cell(row=out_row, column=col_idx, value=val).fill = fill
                    out_row += 1

    wb_out.save(output_file)
    print(f"✅ Output saved to {output_file}")
