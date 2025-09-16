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
YEAR_FOR_OUTPUT = datetime.now().year  # ✅ current year

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
    """
    Converts a month string or number to its corresponding month number (1-12).
    Handles:
      - Integers or numeric strings ("9" → 9)
      - Full month names ("September" → 9)
      - Abbreviations ("Sep", "Sept." → 9)
    """
    if month_value is None:
        return None

    s = str(month_value).strip().lower()
    if not s:
        return None

    # Remove punctuation (like "Sept.")
    s = re.sub(r"[^\w]", "", s)

    # Direct numeric
    if s.isdigit():
        n = int(s)
        if 1 <= n <= 12:
            return n
        return None

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

    return MONTH_MAP.get(s)

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

    event_color_cache = {}
    seen_events = set()  # track duplicates

    for sheet_name in wb_src.sheetnames:
        if sheet_name.upper() not in TARGET_SHEETS:
            continue
        ws_src = wb_src[sheet_name]
        ws_out = wb_out.create_sheet(sheet_name)
        add_headers(ws_out)

        # map headers
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

            event_name = f"{activity} - {resort}" if resort else activity

            # skip duplicates
            if (sheet_name, event_name) in seen_events:
                continue
            seen_events.add((sheet_name, event_name))

            instrs = instructors_map.get(sheet_name, {}).get(activity, [])

            if event_name not in event_color_cache:
                event_color_cache[event_name] = get_light_fill()
            fill = event_color_cache[event_name]

            # ----------- TIMING LOGIC (regex based) -----------
            if not timing:
                continue

            # extract all "HH:MM - HH:MM" slots
            timing_slots = re.findall(r"\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}", timing)

            if not timing_slots:
                continue

            for slot in timing_slots:
                parts = [p.strip() for p in slot.split("-", 1)]
                if len(parts) != 2:
                    continue

                start, end = parts

                # main event row
                for col_idx, val in enumerate([event_name, activity, activity, date_str, start, end], start=1):
                    ws_out.cell(row=out_row, column=col_idx, value=val).fill = fill
                out_row += 1

                # instructor rows
                for instr in instrs:
                    for col_idx, val in enumerate([event_name, instr, activity, date_str, start, end], start=1):
                        ws_out.cell(row=out_row, column=col_idx, value=val).fill = fill
                    out_row += 1
            # ----------- END TIMING LOGIC -----------

    wb_out.save(output_file)
    print(f"✅ Output saved to {output_file}")
