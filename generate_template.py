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
    """Remove brackets content and trim."""
    if not name:
        return None
    return re.sub(r"\s*\(.*?\)\s*", "", str(name)).strip()


# ---------- STAFF LOADER (unchanged) ----------
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


# ---------- NEW: activity -> resorts detection ----------
def build_activity_resort_map(wb_src):
    """Return dict: activity_lower -> set of resorts where it appears."""
    m = {}
    for sheet_name in wb_src.sheetnames:
        if sheet_name.upper() not in TARGET_SHEETS:
            continue
        ws = wb_src[sheet_name]
        # map headers on this sheet
        headers = {safe_str(c.value).lower(): idx + 1 for idx, c in enumerate(ws[1])}
        activity_col = headers.get("activity")
        resort_col = headers.get("resort name")
        if not activity_col:
            continue
        for r in range(2, ws.max_row + 1):
            activity = safe_str(ws.cell(row=r, column=activity_col).value)
            if not activity:
                continue
            resort = safe_str(ws.cell(row=r, column=resort_col).value) if resort_col else ""
            m.setdefault(activity.lower(), set()).add(resort)
    return m


# ---------- MAIN ----------
def generate_output(events_file, staff_file, output_file):
    instructors_map = preload_staff(staff_file)  # {sheet: {activity: [instrs]}}
    wb_src = load_workbook(events_file, data_only=True)

    # build map to know when an activity appears with multiple resorts
    activity_resort_map = build_activity_resort_map(wb_src)

    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    event_color_cache = {}  # ensure same event always has same color

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
        product_col = header_map.get("product")
        age_col = header_map.get("age group")
        guest_col = header_map.get("guest price")
        staff_col = header_map.get("staff price")
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
            product = safe_str(ws_src.cell(row=r, column=product_col).value) if product_col else ""
            # ignore age/price per request
            duration = safe_str(ws_src.cell(row=r, column=duration_col).value) if duration_col else ""
            timing = safe_str(ws_src.cell(row=r, column=timing_col).value) if timing_col else ""
            month_val = ws_src.cell(row=r, column=month_col).value

            # skip multi-day Galaxea
            if sheet_name.upper() == "GALAXEA" and "day" in duration.lower():
                continue

            # find first cell with number > 0
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

            # determine event name according to new rules:
            # - use raw activity name
            # - if same activity exists on multiple resorts, append resort name to activity
            activity_key = activity.lower()
            resorts_for_activity = activity_resort_map.get(activity_key, set())
            if len(resorts_for_activity) > 1 and resort:
                event_name = f"{activity} - {resort}"
            else:
                event_name = activity

            # Default Resource and Configuration should be the actual activity name (raw)
            default_resource = activity
            configuration = activity

            # get instructors for this activity from staff sheet mapping
            instrs = instructors_map.get(sheet_name, {}).get(activity, [])

            # handle timing: if contains '|' split into multiple slots
            time_slots = [timing]
            if timing and '|' in timing:
                parts = [p.strip() for p in timing.split('|') if p.strip()]
                if parts:
                    time_slots = parts

            for slot in time_slots:
                # parse slot into start/end if '-' present
                start_time = None
                end_time = None
                if slot and '-' in slot:
                    parts = [p.strip() for p in slot.split('-', 1)]
                    start_time = parts[0] if parts[0] else None
                    end_time = parts[1] if len(parts) > 1 and parts[1] else None
                else:
                    start_time = slot if slot else None

                # reuse same color for same event
                if event_name not in event_color_cache:
                    # subtle consistent color by hashing event name
                    event_color_cache[event_name] = get_light_fill()
                fill = event_color_cache[event_name]

                # main event row (Event, Resource, Configuration, Date, Start Time, End Time)
                for col_idx, val in enumerate([event_name, default_resource, configuration, date_str, start_time or "", end_time or ""], start=1):
                    c = ws_out.cell(row=out_row, column=col_idx, value=val)
                    c.fill = fill
                out_row += 1

                # instructor rows (same shading)
                for instr in instrs:
                    for col_idx, val in enumerate([event_name, instr, configuration, date_str, start_time or "", end_time or ""], start=1):
                        c = ws_out.cell(row=out_row, column=col_idx, value=val)
                        c.fill = fill
                    out_row += 1

    wb_out.save(output_file)
    print(f"✅ Output saved to {output_file}")