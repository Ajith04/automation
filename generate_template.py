import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
import random

# ---------- CONFIG ----------
TARGET_SHEETS = ["AKUN", "WAMA", "GALAXEA"]
TARGET_RED = "FFC00000"  # ignore instructor if event cell red
HEADERS = ["Event", "Resource", "Configuration", "Date", "Start Time", "End Time"]
HEADER_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
HEADER_FONT = Font(bold=True)
YEAR_FOR_OUTPUT = 2025

# ---------- HELPERS ----------
def safe_str(v):
    return "" if v is None else str(v).strip()

def get_rgb(cell):
    if cell is None or cell.fill is None:
        return None
    color = getattr(cell.fill, "start_color", None)
    if color and hasattr(color, "rgb") and color.rgb:
        return str(color.rgb).upper()
    return None

def is_red(cell):
    return get_rgb(cell) == TARGET_RED

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
    return PatternFill(start_color=random.choice(colors), end_color=random.choice(colors), fill_type="solid")

def get_dropdown_values(ws, cell, wb):
    """
    Get all options from a cell's dropdown, including cross-sheet references.
    """
    dropdown_values = []
    for dv in ws.data_validations.dataValidation:
        if cell.coordinate in dv.cells and dv.type == "list" and dv.formula1:
            formula = dv.formula1
            if formula.startswith('"') and formula.endswith('"'):
                # Inline list
                dropdown_values.extend([x.strip() for x in formula.strip('"').split(",")])
            else:
                # Reference to a range, possibly cross-sheet
                ref = formula.lstrip("=")
                if "!" in ref:
                    sheet_name, rng = ref.split("!")
                    ref_ws = wb[sheet_name]
                else:
                    ref_ws = ws
                    rng = ref
                for row in ref_ws[rng]:
                    for c in row:
                        if c.value:
                            dropdown_values.append(str(c.value).strip())
    # Remove duplicates
    return list(dict.fromkeys(dropdown_values))

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

        out_row = 2
        seen_events = set()

        for r in range(2, ws_src.max_row + 1):
            activity = safe_str(ws_src.cell(r, activity_col).value)
            if not activity:
                continue
            resort = safe_str(ws_src.cell(r, resort_col).value) if resort_col else ""
            duration = safe_str(ws_src.cell(r, duration_col).value) if duration_col else ""
            month_val = ws_src.cell(r, month_col).value

            # find first non-empty day if using date columns
            first_day = 1  # fallback
            month_num = 1
            try:
                month_num = int(month_val)
            except:
                month_num = 1
            date_str = f"{first_day:02d}/{month_num:02d}/{YEAR_FOR_OUTPUT}"

            event_name = f"{activity} - {resort}" if resort else activity
            instrs = instructors_map.get(sheet_name, {}).get(activity, [])

            if event_name not in event_color_cache:
                event_color_cache[event_name] = get_light_fill()
            fill = event_color_cache[event_name]

            # Get all dropdown values (cross-sheet)
            bookable_cell = ws_src.cell(r, bookable_col)
            time_slots = get_dropdown_values(ws_src, bookable_cell, wb_src)

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

                # Event row
                for col_idx, val in enumerate([event_name, activity, activity, date_str, start_time, end_time], start=1):
                    ws_out.cell(out_row, col_idx, value=val).fill = fill
                out_row += 1

                # Instructor rows
                for instr in instrs:
                    for col_idx, val in enumerate([event_name, instr, activity, date_str, start_time, end_time], start=1):
                        ws_out.cell(out_row, col_idx, value=val).fill = fill
                    out_row += 1

    wb_out.save(output_file)
    print(f"✅ Output saved to {output_file}")
