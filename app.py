from flask import Flask, render_template, request
import pandas as pd
import os
import glob
from datetime import datetime
import traceback
import re

app = Flask(__name__)

# ─── Configuration ──────────────────────────────────────────────────────────────
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')

# Find any common Excel file (*.xls*, covers xlsx, xlsm, xls, xlsb)
excel_files = glob.glob(os.path.join(DATA_DIR, '*.xls*'))

if not excel_files:
    EXCEL_PATH = None
    INSTRUCTORS = ["No Excel file found in /data"]
    print("ERROR: No Excel file (*.xls*, *.xlsx, *.xlsm, etc.) found in data/ folder")
else:
    excel_files.sort(key=lambda f: (not f.lower().endswith(('.xlsx', '.xlsm')), f))
    EXCEL_PATH = excel_files[0]
    print(f"Using Excel file: {EXCEL_PATH}")
    try:
        excel_file = pd.ExcelFile(EXCEL_PATH)
        INSTRUCTORS = excel_file.sheet_names
        print(f"Successfully loaded {len(INSTRUCTORS)} instructors/sheets: {INSTRUCTORS}")
    except Exception as e:
        print(f"ERROR loading Excel file {EXCEL_PATH}: {e}")
        INSTRUCTORS = ["Excel file found but cannot be read"]
        EXCEL_PATH = None

# ─── Custom day code extraction for sorting ─────────────────────────────────────
DAY_ORDER = {
    'M': 0, 'MO': 0,
    'TU': 1, 'T': 1,
    'W': 2,
    'TH': 3, 'R': 3,
    'F': 4,
    'SA': 5, 'S': 5,
    'SU': 6,
}

def extract_day_code(class_name):
    if pd.isna(class_name):
        return 99
    # Look for 1-2 uppercase letters at the very end (after optional space)
    match = re.search(r'\s*([A-Z]{1,2})$', str(class_name).strip().upper())
    if match:
        code = match.group(1)
        return DAY_ORDER.get(code, 99)
    return 99

# ─── Generate table HTML with safe sorting ──────────────────────────────────────
def get_table_html(instructor):
    if EXCEL_PATH is None:
        return '<p style="color: red; font-weight: bold; padding: 20px;">No valid Excel file available in /data folder.</p>'

    try:
        if instructor not in INSTRUCTORS:
            return f'<p style="color: red;">Sheet for "{instructor}" not found in Excel.</p>'

        print(f"[DEBUG] Loading sheet: '{instructor}'")

        df = pd.read_excel(EXCEL_PATH, sheet_name=instructor, engine='openpyxl')
        df = df.fillna('')

        # Normalize column names (remove extra spaces, make consistent)
        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        # Build sort columns and ascending dynamically
        sort_cols = []
        ascending = []

        if 'Class Name' in df.columns:
            df['sort_day'] = df['Class Name'].apply(extract_day_code)
            sort_cols.append('sort_day')
            ascending.append(True)   # earliest day first

            sort_cols.append('Class Name')
            ascending.append(True)   # alphabetical within day

        if 'Date' in df.columns:
            try:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                sort_cols.append('Date')
                ascending.append(False)  # newest first
            except Exception as date_err:
                print(f"Date parsing failed for {instructor}: {date_err}")

        if sort_cols:
            print(f"[DEBUG] Sorting by: {sort_cols} ascending: {ascending}")
            df = df.sort_values(by=sort_cols, ascending=ascending)
            if 'sort_day' in df.columns:
                df = df.drop(columns=['sort_day'], errors='ignore')

        table_html = df.to_html(
            index=False,
            classes="table table-striped table-bordered table-hover",
            border=0,
            justify="left",
            escape=False
        )
        return table_html

    except Exception as e:
        traceback.print_exc()
        error_msg = f'<p style="color: red; font-weight: bold; padding: 20px;">Error loading data for {instructor}: {str(e)}</p>'
        if 'df' in locals():
            error_msg += f'<p>Available columns: {", ".join(df.columns.tolist())}</p>'
        return error_msg

# ─── Unified route for / and /danger-report ─────────────────────────────────────
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def danger_report(path=''):
    instructor_param = request.args.get('instructor', '').strip()

    print(f"[DEBUG] Requested path: /{path}, instructor param: '{instructor_param}'")

    instructor = instructor_param

    if not instructor or instructor not in INSTRUCTORS:
        if instructor_param:
            print(f"[WARN] Instructor '{instructor_param}' not found in sheets")
        instructor = INSTRUCTORS[0] if INSTRUCTORS and INSTRUCTORS[0] not in [
            "No Excel file found in /data",
            "Excel file found but cannot be read"
        ] else "No Data Available"

    table_html = get_table_html(instructor)
    updated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    return render_template(
        'danger_report.html',
        instructor=instructor,
        instructors=INSTRUCTORS,
        table_html=table_html,
        updated_at=updated_at
    )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)