import re
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

# mapping month names/abbreviations to numbers
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

# ---------- MAIN ----------
def generate_output(events_file, staff_file, output_file):
    instructors_map = preload_staff(staff_file)
    wb_src = load_workbook(events_file, data_only=True)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    activity_cache = {}
    event_color_cache = {}

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
        timing_col = header_map.get("timing availability")
        month_col = header_map.get("month")

        if month_col is None or activity_col is None:
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
            timing = safe_str(ws_src.cell(row=r, column=timing_col).value) if timing_col else ""
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

            # Event name is always activity + resort
            event_name = f"{activity} - {resort}" if resort else activity

            if activity in activity_cache:
                instrs = activity_cache[activity]
            else:
                instrs = instructors_map.get(sheet_name, {}).get(activity, [])
                activity_cache[activity] = instrs

            if event_name not in event_color_cache:
                event_color_cache[event_name] = get_light_fill()
            fill = event_color_cache[event_name]

            # Handle multiple time slots separated by |
            timing_slots = [t.strip() for t in timing.split("|") if t.strip()] if timing else [""]

            for slot in timing_slots:
                start_time, end_time = None, None
                if "-" in slot:
                    parts = [p.strip() for p in slot.split("-", 1)]
                    start_time = parts[0]
                    if len(parts) > 1:
                        end_time = parts[1]
                elif slot:
                    start_time = slot

                # main event row
                for col_idx, val in enumerate([event_name, activity, activity, date_str], start=1):
                    ws_out.cell(row=out_row, column=col_idx, value=val).fill = fill
                if start_time:
                    ws_out.cell(row=out_row, column=5, value=start_time).fill = fill
                if end_time:
                    ws_out.cell(row=out_row, column=6, value=end_time).fill = fill
                out_row += 1

                # instructor rows
                for instr in instrs:
                    for col_idx, val in enumerate([event_name, instr, activity, date_str], start=1):
                        ws_out.cell(row=out_row, column=col_idx, value=val).fill = fill
                    if start_time:
                        ws_out.cell(row=out_row, column=5, value=start_time).fill = fill
                    if end_time:
                        ws_out.cell(row=out_row, column=6, value=end_time).fill = fill
                    out_row += 1

    wb_out.save(output_file)
    print(f"✅ Output saved to {output_file}")
