from flask import Flask, render_template, request
import pandas as pd
import os
import glob
from datetime import datetime
import traceback
import re

app = Flask(__name__)

# ───────────────────────────────────────────────────────────────
#  Locate Excel file
# ───────────────────────────────────────────────────────────────
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
excel_files = glob.glob(os.path.join(DATA_DIR, '*.xls*'))

if not excel_files:
    EXCEL_PATH = None
    INSTRUCTORS = ["No Excel file found in /data"]
    print("ERROR: No Excel file (*.xls*, *.xlsx, *.xlsm) found in /data")
else:
    excel_files.sort(key=lambda f: (not f.lower().endswith(('.xlsx', '.xlsm')), f))
    EXCEL_PATH = excel_files[0]
    print(f"Using Excel file: {EXCEL_PATH}")

    try:
        excel_file = pd.ExcelFile(EXCEL_PATH)
        raw_sheets = excel_file.sheet_names

        # Skip first sheet (Combined)
        selectable_sheets = raw_sheets[1:] if len(raw_sheets) > 1 else []
        INSTRUCTORS = sorted(selectable_sheets)

        print(f"Loaded instructors: {INSTRUCTORS}")

    except Exception as e:
        print(f"ERROR loading Excel file: {e}")
        EXCEL_PATH = None
        INSTRUCTORS = ["Excel file found but cannot be read"]


# ───────────────────────────────────────────────────────────────
#  Day ordering
# ───────────────────────────────────────────────────────────────
DAY_ORDER = {
    "M": 0, "MO": 0, "MON": 0, "MONDAY": 0,
    "T": 1, "TU": 1, "TUE": 1, "TUESDAY": 1,
    "W": 2, "WE": 2, "WED": 2, "WEDNESDAY": 2,
    "TH": 3, "R": 3, "THU": 3, "THURSDAY": 3,
    "F": 4, "FRI": 4, "FRIDAY": 4,
    "SA": 5, "SAT": 5, "SATURDAY": 5,
    "SU": 6, "SUN": 6, "SUNDAY": 6,
}


def extract_day_code(class_name):
    """Extract weekday from end of Class Name.
       Supports: M, TU, W, TH, F, SA, SU
       AND: Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday
    """
    if pd.isna(class_name):
        return 99

    text = str(class_name).strip().upper()

    # Check full weekday names
    for full in ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY",
                 "FRIDAY", "SATURDAY", "SUNDAY"]:
        if text.endswith(full):
            return DAY_ORDER[full]

    # Check short codes (1–3 letters)
    match = re.search(r'([A-Z]{1,3})$', text)
    if match:
        code = match.group(1)
        return DAY_ORDER.get(code, 99)

    return 99


# ───────────────────────────────────────────────────────────────
#  Build table HTML
# ───────────────────────────────────────────────────────────────
def get_table_html(instructor: str) -> str:
    if EXCEL_PATH is None:
        return '<p style="color:red;font-weight:bold;">No Excel file found.</p>'

    try:
        if instructor not in INSTRUCTORS:
            return f'<p style="color:red;">Instructor "{instructor}" not found.</p>'

        print(f"[DEBUG] Loading sheet: {instructor}")

        df = pd.read_excel(EXCEL_PATH, sheet_name=instructor, engine='openpyxl')
        df = df.fillna('')

        # Normalize column names
        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        # ─── Sorting ─────────────────────────────────────────────
        if "Class Name" in df.columns:
            df["sort_day"] = df["Class Name"].apply(extract_day_code)
            df = df.sort_values(by=["sort_day", "Class Name"], ascending=[True, True])
            df = df.drop(columns=["sort_day"], errors="ignore")

        # ─── Row coloring ───────────────────────────────────────
        def row_background(row):
            day_num = extract_day_code(row.get("Class Name", ""))
            colors = {
                0: "#e8f0ff",  # Monday
                1: "#fff0e8",  # Tuesday
                2: "#e8fff0",  # Wednesday
                3: "#fff8e8",  # Thursday
                4: "#f0e8ff",  # Friday
                5: "#e8f4ff",  # Saturday
                6: "#f8f8f8",  # Sunday
                99: "#f0f0f0", # Unknown
            }
            return [f"background-color: {colors.get(day_num, '#ffffff')}"] * len(row)

        styled = df.style.apply(row_background, axis=1)
        styled = styled.set_properties(**{"text-align": "left"})
        styled = styled.hide(axis="index")  # REMOVE LEFT NUMBER COLUMN

        table_html = styled.to_html(
            escape=False,
            index=False,
            classes="dataframe table table-striped table-bordered table-hover",
            border=0,
            justify="left"
        )

        return table_html

    except Exception as e:
        traceback.print_exc()
        return f'<p style="color:red;">Error loading data: {e}</p>'


# ───────────────────────────────────────────────────────────────
#  Route
# ───────────────────────────────────────────────────────────────
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def danger_report(path=''):
    instructor_param = request.args.get("instructor", "").strip()

    print(f"[DEBUG] Request instructor={instructor_param}")

    instructor = None
    table_html = None
    message = None

    if instructor_param and instructor_param in INSTRUCTORS:
        instructor = instructor_param
        table_html = get_table_html(instructor)
    else:
        message = (
            '<div style="padding:40px;text-align:center;background:#fafafa;'
            'border-radius:12px;max-width:700px;margin:40px auto;">'
            '<h2>Select an Instructor</h2>'
            '<p>Use the dropdown above to view their report.</p>'
            '</div>'
        )

    from zoneinfo import ZoneInfo
    updated_at = datetime.now(ZoneInfo("America/Chicago")).strftime("%Y-%m-%d %H:%M:%S")

    return render_template(
        "danger_report.html",
        instructor=instructor or "Select an Instructor",
        instructors=INSTRUCTORS,
        table_html=table_html,
        updated_at=updated_at,
        message=message
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)