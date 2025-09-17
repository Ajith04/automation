import re
import random
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string
import streamlit as st  # For debugging output in Streamlit

# ---------- CONFIG ----------
TARGET_SHEETS = ["AKUN", "WAMA", "GALAXEA"]
TARGET_RED = "FFC00000"    # ignore instructor if event cell red
HEADERS = ["Event", "Resource", "Configuration", "Date", "Start Time", "End Time"]
HEADER_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
HEADER_FONT = Font(bold=True)
YEAR_FOR_OUTPUT = 2025

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
    key = s.lower()
    return MONTH_MAP.get(key)

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

def get_dropdown_values(ws, cell):
    """Return dropdown options including cross-sheet references."""
    dropdowns = []
    for dv in ws.data_validations.dataValidation:
        if cell.coordinate in dv.cells and dv.type == "list" and dv.formula1:
            f = dv.formula1
            st.write(f"Processing dropdown formula: {f}")  # Streamlit debug
            if f.startswith('"') and f.endswith('"'):
                dropdowns.extend([x.strip() for x in f.strip('"').split(",")])
            else:
                ref = f.lstrip("=")
                if "!" in ref:
                    sheet_name, rng = ref.split("!")
                    sheet_name = sheet_name.strip("'")
                    ref_ws = ws.parent[sheet_name]
                else:
                    ref_ws, rng = ws, ref
                st.write(f"Fetching dropdown values from {sheet_name if 'sheet_name' in locals() else ws.title} range {rng}")
                for row in ref_ws[rng]:
                    for c in row:
                        if c.value:
                            dropdowns.append(str(c.value).strip())
    st.write(f"Dropdown values found: {dropdowns}")
    return list(dict.fromkeys(dropdowns))  # remove duplicates

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
    st.write("Loading staff file...")
    instructors_map = preload_staff(staff_file)
    st.write("Staff mapping loaded:", instructors_map)
    
    wb_src = load_workbook(events_file, data_only=True)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    event_color_cache = {}

    for sheet_name in wb_src.sheetnames:
        st.write(f"Processing sheet: {sheet_name}")
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

        st.write(f"Header mapping: {header_map}")

        if not (month_col and activity_col and bookable_col):
            st.write(f"Skipping sheet {sheet_name} because some columns are missing")
            continue

        date_start_col = month_col + 1
        max_check_col = min(ws_src.max_column, date_start_col + 31 - 1)
        out_row = 2

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
            event_name = f"{activity} - {resort}" if len(resorts_for_activity) > 1 and resort else activity
            instrs = instructors_map.get(sheet_name, {}).get(activity, [])

            if event_name not in event_color_cache:
                event_color_cache[event_name] = get_light_fill()
            fill = event_color_cache[event_name]

            bookable_cell = ws_src.cell(r, bookable_col)
            time_slots = get_dropdown_values(ws_src, bookable_cell)

            st.write(f"Activity: {activity}, Time slots: {time_slots}, Instructors: {instrs}")

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

                for col_idx, val in enumerate([event_name, activity, activity, date_str, start_time, end_time], start=1):
                    ws_out.cell(out_row, col_idx, value=val).fill = fill
                out_row += 1

                for instr in instrs:
                    for col_idx, val in enumerate([event_name, instr, activity, date_str, start_time, end_time], start=1):
                        ws_out.cell(out_row, col_idx, value=val).fill = fill
                    out_row += 1

    wb_out.save(output_file)
    st.write(f"✅ Output saved to {output_file}")
