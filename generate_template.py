import xlwings as xw
import re
from datetime import datetime
import random

# ---------- CONFIG ----------
TARGET_SHEETS = ["AKUN", "WAMA", "GALAXEA"]
TARGET_RED = "FFC00000"  # ignore instructor if event cell red
HEADERS = ["Event", "Resource", "Configuration", "Date", "Start Time", "End Time"]
YEAR_FOR_OUTPUT = 2025

# ---------- HELPERS ----------
def safe_str(v):
    return "" if v is None else str(v).strip()

def clean_instructor_name(name):
    """Remove brackets content."""
    if not name:
        return None
    return re.sub(r"\s*\(.*?\)\s*", "", str(name)).strip()

def get_light_fill():
    """Return a random light color as RGB string."""
    colors = [
        "FFFFE5CC", "FFE5FFCC", "FFCCFFE5", "FFCCE5FF",
        "FFFFCCFF", "FFE5CCFF", "FFFFCCCC", "FFCCFFFF"
    ]
    return random.choice(colors)

def parse_month_to_num(month_value):
    """Convert month string/number to integer 1-12."""
    if month_value is None:
        return None
    s = str(month_value).strip()
    if not s:
        return None
    month_map = {
        "january":1,"jan":1,"february":2,"feb":2,"march":3,"mar":3,"april":4,"apr":4,
        "may":5,"june":6,"jun":6,"july":7,"jul":7,"august":8,"aug":8,"september":9,"sep":9,"sept":9,
        "october":10,"oct":10,"november":11,"nov":11,"december":12,"dec":12
    }
    key = s.lower()
    if key in month_map:
        return month_map[key]
    for name, num in month_map.items():
        if name in key:
            return num
    m = re.search(r"\b(1[0-2]|0?[1-9])\b", s)
    if m:
        return int(m.group(1))
    return None

# ---------- STAFF LOADER ----------
def preload_staff(staff_file):
    wb = xw.Book(staff_file)
    result = {}
    for sheet_name in TARGET_SHEETS:
        try:
            ws = wb.sheets[sheet_name]
        except:
            continue
        header_row = ws.range("A1").expand("right").value
        headers = [safe_str(h).lower() for h in header_row]
        try:
            pri_idx = headers.index("priority")
        except ValueError:
            pri_idx = 0
        instr_start_col = pri_idx + 1
        sheet_map = {}
        for col_idx in range(instr_start_col, len(headers)):
            instr_name = clean_instructor_name(safe_str(header_row[col_idx]))
            if not instr_name:
                continue
            r = 2
            while True:
                val = safe_str(ws.cells(r, col_idx+1).value)
                if not val:
                    break
                # TODO: skip red cells if needed
                sheet_map.setdefault(val, []).append(instr_name)
                r += 1
        result[sheet_name] = sheet_map
    wb.app.quit()
    return result

# ---------- DROPDOWN TIMESLOTS ----------
def get_timeslots_from_dropdown(cell):
    """Return all dropdown options for a cell."""
    dv_list = cell.api.Validation
    options = []
    if dv_list:
        formula = dv_list.Formula1
        if formula.startswith('"') and formula.endswith('"'):
            # inline list
            options = [i.strip() for i in formula.strip('"').split(',')]
        elif formula.startswith('='):
            # named range or reference
            try:
                rng = cell.sheet.range(formula[1:])
                options = [safe_str(c.value) for c in rng if c.value]
            except:
                options = []
    return options

# ---------- MAIN ----------
def generate_output(events_file, staff_file, output_file):
    instructors_map = preload_staff(staff_file)
    wb_src = xw.Book(events_file)
    wb_out = xw.Book()
    # remove default sheet
    wb_out.sheets[0].delete()
    event_color_cache = {}

    for sheet_name in wb_src.sheets:
        if sheet_name.name.upper() not in TARGET_SHEETS:
            continue
        ws_src = sheet_name
        ws_out = wb_out.sheets.add(sheet_name.name)

        # add headers
        for col_idx, h in enumerate(HEADERS, start=1):
            c = ws_out.cells(1, col_idx)
            c.value = h
            c.color = (255, 255, 0)  # yellow

        # map headers
        header_row = ws_src.range("A1").expand("right").value
        headers = {safe_str(h).lower(): idx+1 for idx, h in enumerate(header_row)}
        resort_col = headers.get("resort name")
        activity_col = headers.get("activity")
        bookable_col = headers.get("bookable hours")
        month_col = headers.get("month")
        duration_col = headers.get("activity duration")

        if not (month_col and activity_col and bookable_col):
            continue

        date_start_col = month_col + 1
        max_check_col = date_start_col + 31 - 1
        out_row = 2

        # detect multi-resort activities
        activity_resorts = {}
        last_row = ws_src.range("A1").expand("down").last_cell.row
        for r in range(2, last_row+1):
            act = safe_str(ws_src.cells(r, activity_col).value)
            res = safe_str(ws_src.cells(r, resort_col).value) if resort_col else ""
            if act:
                activity_resorts.setdefault(act, set()).add(res)

        seen_events = set()
        for r in range(2, last_row+1):
            activity = safe_str(ws_src.cells(r, activity_col).value)
            if not activity:
                continue
            resort = safe_str(ws_src.cells(r, resort_col).value) if resort_col else ""
            duration = safe_str(ws_src.cells(r, duration_col).value) if duration_col else ""
            month_val = ws_src.cells(r, month_col).value

            if ws_src.name.upper() == "GALAXEA" and "day" in duration.lower():
                continue

            # first non-empty day
            first_day = None
            for col in range(date_start_col, max_check_col+1):
                try:
                    val = ws_src.cells(r, col).value
                    if val is not None and float(val) > 0:
                        first_day = int(ws_src.cells(1, col).value)
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

            instrs = instructors_map.get(ws_src.name, {}).get(activity, [])

            if event_name not in event_color_cache:
                event_color_cache[event_name] = get_light_fill()
            fill_color = event_color_cache[event_name]
            fill_rgb = tuple(int(fill_color[i:i+2], 16) for i in (2,4,6))  # convert to RGB

            # get ALL dropdown timeslots
            cell = ws_src.cells(r, bookable_col)
            time_slots = get_timeslots_from_dropdown(cell)

            for slot in time_slots:
                slot = slot.strip()
                if not slot:
                    continue
                if "-" in slot:
                    parts = [p.strip() for p in slot.split("-", 1)]
                    start_time, end_time = parts if len(parts)==2 else (parts[0], "")
                else:
                    start_time, end_time = slot, ""

                key = (event_name, resort, activity, date_str, start_time, end_time)
                if key in seen_events:
                    continue
                seen_events.add(key)

                # main event row
                for col_idx, val in enumerate([event_name, activity, activity, date_str, start_time, end_time], start=1):
                    c = ws_out.cells(out_row, col_idx)
                    c.value = val
                    c.color = fill_rgb
                out_row += 1

                # instructors
                for instr in instrs:
                    for col_idx, val in enumerate([event_name, instr, activity, date_str, start_time, end_time], start=1):
                        c = ws_out.cells(out_row, col_idx)
                        c.value = val
                        c.color = fill_rgb
                    out_row += 1

    wb_out.save(output_file)
    wb_src.app.quit()
    wb_out.app.quit()
    print(f"✅ Output saved to {output_file}")
