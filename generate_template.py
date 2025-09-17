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
    """Return the RGB string of a cell's fill color, or None."""
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
    """Convert month string/number to integer 1-12."""
    if month_value is None:
        return None

    s = str(month_value).strip()
    if not s:
        return None

    # Case: numeric string or int
    if s.isdigit() and 1 <= int(s) <= 12:
        return int(s)

    key = s.lower()

    # ---- Step 1: exact dictionary match ----
    if key in MONTH_MAP:
        return MONTH_MAP[key]

    # ---- Step 2: exact equality with known month keys ----
    for name, num in MONTH_MAP.items():
        if key == name:
            return num

    # ---- Step 3: substring match (e.g., "September 2025") ----
    for name, num in MONTH_MAP.items():
        if name in key:
            return num

    # ---- Step 4: numeric extraction from text ----
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
    """Remove brackets content."""
    if not name:
        return None
    return re.sub(r"\s*\(.*?\)\s*", "", str(name)).strip()

# ---------- STAFF LOADER ----------
def preload_staff(staff_file):
    """Load staff workbook once and build {sheet -> {activity -> [instrs]}}"""
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
                    break  # stop at first blank
                if is_red(val_cell):
                    r += 1
                    continue  # skip events with red cell
                sheet_map.setdefault(val, []).append(instr_name)
                r += 1
        result[sheet] = sheet_map
    return result

def build_event_names(sheet_name, resort, activity, product, age_group, guest_price, staff_price):
    """Build one or more event names according to sheet rules."""
    out = []
    act = safe_str(activity)
    res = safe_str(resort)
    prod = safe_str(product)
    ag = safe_str(age_group)
    if not act:
        return out
    sheet_name = sheet_name.upper()
    if sheet_name == "AKUN":
        if ag == "13+":
            if guest_price:
                out.append(f"{act} - {res}")
            if staff_price:
                out.append(f"{act} - {res} - Staff")
        elif re.search(r"8\s*-\s*12", ag) or "8-12" in ag or "years" in ag.lower():
            if guest_price:
                out.append(f"{act} - {res} - Child")
            if staff_price:
                out.append(f"{act} - {res} - RSG Staff - Child")
        else:
            if guest_price:
                out.append(f"{act} - {res}")
            if staff_price:
                out.append(f"{act} - {res} - Staff")
    elif sheet_name in ("WAMA", "GALAXEA"):
        base = f"{act} - {res}" if res else act
        if prod:
            base = f"{base} - {prod}"
        if guest_price:
            out.append(base)
        if staff_price:
            out.append(base + " - Staff")
    return out

def get_light_fill():
    """Return a random light color PatternFill."""
    colors = [
        "FFFFE5CC", "FFE5FFCC", "FFCCFFE5", "FFCCE5FF",
        "FFFFCCFF", "FFE5CCFF", "FFFFCCCC", "FFCCFFFF"
    ]
    hexcolor = random.choice(colors)
    return PatternFill(start_color=hexcolor, end_color=hexcolor, fill_type="solid")

# ---------- MAIN ----------
def generate_output(events_file, staff_file, output_file):
    instructors_map = preload_staff(staff_file)
    wb_src = load_workbook(events_file, data_only=True)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    event_color_cache = {}

    for sheet_name in wb_src.sheetnames:
        if sheet_name.upper() not in TARGET_SHEETS:
            continue
        ws_src = wb_src[sheet_name]
        ws_out = wb_out.create_sheet(sheet_name)
        add_headers(ws_out)

        # map headers
        header_map = {safe_str(ws_src.cell(1, c).value).lower(): c
                      for c in range(1, ws_src.max_column + 1)}
        resort_col = header_map.get("resort name")
        activity_col = header_map.get("activity")
        timing_col = header_map.get("timing availability")
        month_col = header_map.get("month")
        duration_col = header_map.get("activity duration")

        if not (month_col and activity_col):
            continue

        date_start_col = month_col + 1
        max_check_col = min(ws_src.max_column, date_start_col + 31 - 1)
        out_row = 2

        # ---------- Pass 1: detect multi-resort activities ----------
        activity_resorts = {}
        for r in range(2, ws_src.max_row + 1):
            act = safe_str(ws_src.cell(r, activity_col).value)
            res = safe_str(ws_src.cell(r, resort_col).value) if resort_col else ""
            if act:
                activity_resorts.setdefault(act, set()).add(res)

        # ---------- Pass 2: generate output ----------
        seen_events = set()  # prevent duplicates (activity+resort+date+time)
        for r in range(2, ws_src.max_row + 1):
            activity = safe_str(ws_src.cell(r, activity_col).value)
            if not activity:
                continue
            resort = safe_str(ws_src.cell(r, resort_col).value) if resort_col else ""
            duration = safe_str(ws_src.cell(r, duration_col).value) if duration_col else ""
            timing = safe_str(ws_src.cell(r, timing_col).value) if timing_col else ""
            month_val = ws_src.cell(r, month_col).value

            # skip multi-day Galaxea
            if sheet_name.upper() == "GALAXEA" and "day" in duration.lower():
                continue

            # find first non-empty day
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

            # -------- Build Event Name --------
            resorts_for_activity = activity_resorts.get(activity, set())
            if len(resorts_for_activity) > 1 and resort:
                event_name = f"{activity} - {resort}"
            else:
                event_name = activity

            # -------- Instructors --------
            instrs = instructors_map.get(sheet_name, {}).get(activity, [])

            # assign consistent fill color
            if event_name not in event_color_cache:
                event_color_cache[event_name] = get_light_fill()
            fill = event_color_cache[event_name]

            # -------- Handle Timing Availability --------
            time_slots = [timing] if "|" not in timing else timing.split("|")

            for slot in time_slots:
                slot = slot.strip()
                if not slot:
                    continue
                if "-" in slot:
                    parts = [p.strip() for p in slot.split("-", 1)]
                    start_time, end_time = parts if len(parts) == 2 else (parts[0], "")
                else:
                    start_time, end_time = slot, ""

                key = (event_name, resort, activity, date_str, start_time, end_time)
                if key in seen_events:
                    continue  # skip duplicate
                seen_events.add(key)

                # main event row
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
