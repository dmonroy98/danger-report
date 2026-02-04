from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import glob
from datetime import datetime
import traceback
import re
from bs4 import BeautifulSoup
from io import BytesIO

app = Flask(__name__)

# ───────────────────────────────────────────────────────────────
# Locate and load Excel file at startup
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
        # Skip first sheet (Combined) if it exists
        selectable_sheets = raw_sheets[1:] if len(raw_sheets) > 1 else raw_sheets
        INSTRUCTORS = sorted(selectable_sheets)
        print(f"Loaded instructors: {INSTRUCTORS}")
    except Exception as e:
        print(f"ERROR loading Excel file: {e}")
        EXCEL_PATH = None
        INSTRUCTORS = ["Excel file found but cannot be read"]

# ───────────────────────────────────────────────────────────────
# Day ordering for sorting classes
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
    if pd.isna(class_name):
        return 99
    text = str(class_name).strip().upper()
    for full in ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]:
        if text.endswith(full):
            return DAY_ORDER[full]
    match = re.search(r'([A-Z]{1,3})$', text)
    if match:
        code = match.group(1)
        return DAY_ORDER.get(code, 99)
    return 99

# ───────────────────────────────────────────────────────────────
# Generate styled table HTML
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

        # Improved sorting: by Day → Time → Class Name
        if "Class Name" in df.columns:
            def parse_schedule(class_name):
                if pd.isna(class_name):
                    return (99, 9999)  # fallback: very late

                text = str(class_name).strip()
                # Extract day (last 1-3 letters)
                day_match = re.search(r'\b([A-Za-z]{1,3})$', text, re.IGNORECASE)
                day_code = day_match.group(1).upper() if day_match else ""
                day_num = DAY_ORDER.get(day_code, 99)

                # Extract time (e.g. 10:45A, 9:15P, 12:30P)
                time_match = re.search(r'(\d{1,2}):(\d{2})([AP])', text, re.IGNORECASE)
                if not time_match:
                    return (day_num, 9999)

                hour, minute, ampm = time_match.groups()
                h = int(hour)
                m = int(minute)
                if ampm.upper() == 'P' and h != 12:
                    h += 12
                if ampm.upper() == 'A' and h == 12:
                    h = 0

                minutes_since_midnight = h * 60 + m
                return (day_num, minutes_since_midnight)

            # Apply parsing → new columns for sorting
            df[['sort_day', 'sort_time']] = pd.DataFrame(
                df["Class Name"].apply(parse_schedule).tolist(),
                index=df.index
            )

            # Sort: day → time → class name (you can add "Student Name" here too)
            df = df.sort_values(
                by=["sort_day", "sort_time", "Class Name"],
                ascending=[True, True, True]
            )

            # Clean up temporary sort columns
            df = df.drop(columns=["sort_day", "sort_time"], errors="ignore")

        # Day-based row background colors (unchanged)
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

        # Create styled HTML
        styled = df.style \
            .apply(row_background, axis=1) \
            .set_properties(**{
                'text-align': 'left',
                'white-space': 'normal !important',
                'word-break': 'break-word',
                'overflow-wrap': 'break-word',
                'min-width': '0',
                'max-width': '320px'
            }) \
            .hide(axis="index") \
            .set_table_attributes('class="table table-striped table-bordered table-hover"')

        table_html = styled.to_html(
            escape=False,
            index=False,
            border=0,
            justify="left",
            na_rep="-"
        )

        # Clean up unwanted inline widths with BeautifulSoup
        soup = BeautifulSoup(table_html, 'html.parser')
        for tag in soup.find_all(['th', 'td']):
            if tag.has_attr('style'):
                styles = tag['style'].split(';')
                cleaned = [s.strip() for s in styles if s.strip() and not any(k in s.lower() for k in ['width', 'min-width', 'max-width'])]
                if cleaned:
                    tag['style'] = '; '.join(cleaned)
                else:
                    del tag['style']

        return str(soup)

    except Exception as e:
        traceback.print_exc()
        return f'<p style="color:red;">Error loading data: {str(e)}</p>'

# ───────────────────────────────────────────────────────────────
# Server-side Excel download
# ───────────────────────────────────────────────────────────────
@app.route('/download-excel')
def download_excel():
    instructor = request.args.get("instructor", "").strip()
    if not instructor or instructor not in INSTRUCTORS or EXCEL_PATH is None:
        return "Invalid instructor or no data available", 400
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=instructor, engine='openpyxl')
        df = df.fillna('')
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=instructor[:31])  # Excel sheet name limit
        output.seek(0)
        safe_name = re.sub(r'[^a-zA-Z0-9_-]', '_', instructor)
        filename = f"{safe_name}_student_heads_up_report.xlsx"
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        traceback.print_exc()
        return f"Error generating Excel file: {str(e)}", 500

# ───────────────────────────────────────────────────────────────
# Main route
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
            '<h2>HEADS UP REPORT</h2>'
            '<p>Select an instructor to view student information.</p>'
            '</div>'
        )
    from zoneinfo import ZoneInfo
    updated_at = datetime.now(ZoneInfo("America/Chicago")).strftime("%Y-%m-%d %H:%M:%S %Z")
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