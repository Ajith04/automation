# generate_template.py
import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
import random

# ---------- CONFIG ----------
EVENTS_FILE = None  # to be passed to function
STAFF_FILE = None   # to be passed to function
OUTPUT_FILE = None  # to be passed to function

TARGET_SHEETS = ["AKUN", "WAMA", "GALAXEA"]
TARGET_GREEN = "FF00B050"  # event availability
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

def is_green(cell):
    return get_rgb(cell) == TARGET_GREEN

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
    return PatternFill(start_color=random.choice(colors),
                       end_color=random.choice(colors),
                       fill_type="solid")

# ---------- MAIN ----------
def generate_output(events_file, staff_file, output_file):
    instructors_map = preload_staff(staff_file)
    wb_src = load_workbook(events_file, data_only=True)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    activity_cache = {}  # raw activity -> instructors

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
            age_group = safe_str(ws_src.cell(row=r, column=age_col).value) if age_col else ""
            guest_price = ws_src.cell(row=r, column=guest_col).value if guest_col else None
            staff_price = ws_src.cell(row=r, column=staff_col).value if staff_col else None
            duration = safe_str(ws_src.cell(row=r, column=duration_col).value) if duration_col else ""
            timing = safe_str(ws_src.cell(row=r, column=timing_col).value) if timing_col else ""
            month_val = ws_src.cell(row=r, column=month_col).value

            # skip multi-day Galaxea
            if sheet_name.upper() == "GALAXEA" and "day" in duration.lower():
                continue

            # find first green date
            first_green_day = None
            for col in range(date_start_col, max_check_col + 1):
                c = ws_src.cell(row=r, column=col)
                if is_green(c):
                    hdr = ws_src.cell(row=1, column=col).value
                    try:
                        first_green_day = int(hdr)
                    except:
                        first_green_day = None
                    break
            if not first_green_day:
                continue

            month_num = parse_month_to_num(month_val)
            if not month_num:
                continue
            date_str = f"{first_green_day:02d}/{month_num:02d}/{YEAR_FOR_OUTPUT}"

            # build event names
            event_names = build_event_names(sheet_name, resort, activity, product,
                                            age_group, guest_price, staff_price)
            if not event_names:
                continue

            # assign random light color for this event
            fill = get_light_fill()

            # get instructors (cache)
            if activity in activity_cache:
                instrs = activity_cache[activity]
            else:
                instrs = instructors_map.get(sheet_name, {}).get(activity, [])
                activity_cache[activity] = instrs

            for event in event_names:
                # main event row
                ws_out.cell(row=out_row, column=1, value=event).fill = fill
                ws_out.cell(row=out_row, column=2, value=event).fill = fill
                ws_out.cell(row=out_row, column=3, value=activity).fill = fill
                ws_out.cell(row=out_row, column=4, value=date_str).fill = fill
                if "-" in timing:
                    parts = [p.strip() for p in timing.split("-", 1)]
                    ws_out.cell(row=out_row, column=5, value=parts[0]).fill = fill
                    if len(parts) > 1:
                        ws_out.cell(row=out_row, column=6, value=parts[1]).fill = fill
                out_row += 1

                # instructor rows
                for instr in instrs:
                    ws_out.cell(row=out_row, column=1, value=event)
                    ws_out.cell(row=out_row, column=2, value=instr)
                    ws_out.cell(row=out_row, column=3, value=activity)
                    ws_out.cell(row=out_row, column=4, value=date_str)
                    if "-" in timing:
                        parts = [p.strip() for p in timing.split("-", 1)]
                        ws_out.cell(row=out_row, column=5, value=parts[0])
                        if len(parts) > 1:
                            ws_out.cell(row=out_row, column=6, value=parts[1])
                    out_row += 1

    wb_out.save(output_file)
    print(f"✅ Output saved to {output_file}")
