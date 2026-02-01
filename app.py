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

# ─── Day ordering and class mapping ─────────────────────────────────────────────
DAY_ORDER = {
    'M':   0, 'MO':  0,     # Monday
    'TU':  1, 'T':   1,     # Tuesday
    'W':   2, 'WE':  2,     # Wednesday
    'TH':  3, 'R':   3,     # Thursday
    'F':   4,               # Friday
    'SA':  5, 'S':   5,     # Saturday
    'SU':  6,               # Sunday
}

DAY_CLASSES = {
    0: 'monday',
    1: 'tuesday',
    2: 'wednesday',
    3: 'thursday',
    4: 'friday',
    5: 'saturday',
    6: 'sunday',
    99: 'unknown'
}

def extract_day_code(class_name):
    if pd.isna(class_name):
        return 99
    match = re.search(r'\s*([A-Z]{1,2})$', str(class_name).strip().upper())
    if match:
        code = match.group(1)
        return DAY_ORDER.get(code, 99)
    return 99

def get_day_class(day_num):
    return f"day-row-{DAY_CLASSES.get(day_num, 'unknown')}"

# ─── Generate table HTML ────────────────────────────────────────────────────────
def get_table_html(instructor):
    if EXCEL_PATH is None:
        return '<p style="color: red; font-weight: bold; padding: 20px;">No valid Excel file available in /data folder.</p>'

    try:
        if instructor not in INSTRUCTORS:
            return f'<p style="color: red;">Sheet for "{instructor}" not found in Excel.</p>'

        print(f"[DEBUG] Loading sheet: '{instructor}'")

        df = pd.read_excel(EXCEL_PATH, sheet_name=instructor, engine='openpyxl')
        df = df.fillna('')

        # Normalize column names
        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        # Sorting & row class preparation
        sort_cols = []
        ascending = []

        if 'Class Name' in df.columns:
            df['sort_day'] = df['Class Name'].apply(extract_day_code)
            df['row_class'] = df['sort_day'].apply(get_day_class)

            # Sort: Monday (0) → Sunday (6), then Class Name alphabetical, then Date newest first
            sort_cols.append('sort_day')
            ascending.append(True)   # low number (Monday) first

            sort_cols.append('Class Name')
            ascending.append(True)

            if 'Date' in df.columns:
                try:
                    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                    sort_cols.append('Date')
                    ascending.append(False)  # newest first
                except Exception as date_err:
                    print(f"Date parsing failed for {instructor}: {date_err}")

            print(f"[DEBUG] Sorting by: {sort_cols} ascending: {ascending}")
            df = df.sort_values(by=sort_cols, ascending=ascending)
            df = df.drop(columns=['sort_day'], errors='ignore')

        else:
            df['row_class'] = 'day-row-unknown'
            print(f"Warning: No 'Class Name' column → no day sorting/coloring for {instructor}")

        # Generate HTML table
        table_html = df.to_html(
            index=False,
            classes="table table-striped table-bordered table-hover",
            border=0,
            justify="left",
            escape=False
        )

        # Post-process: move row_class from first <td> to <tr> (hack for pandas to_html limitation)
        table_html = re.sub(
            r'<tr>\s*<td>(day-row-[a-z-]+)</td>',
            r'<tr class="\1">',
            table_html,
            flags=re.IGNORECASE
        )
        # Remove the temp row_class column from header and body
        table_html = re.sub(r'<th>row_class</th>\s*', '', table_html)
        table_html = re.sub(r'<td>day-row-[a-z-]+</td>\s*', '', table_html, flags=re.IGNORECASE)

        return table_html

    except Exception as e:
        traceback.print_exc()
        error_msg = f'<p style="color: red; font-weight: bold; padding: 20px;">Error loading data for {instructor}: {str(e)}</p>'
        if 'df' in locals():
            error_msg += f'<p>Available columns: {", ".join(df.columns.tolist())}</p>'
        return error_msg

# ─── Unified route ──────────────────────────────────────────────────────────────
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def danger_report(path=''):
    instructor_param = request.args.get('instructor', '').strip()

    print(f"[DEBUG] Requested: /{path} ?instructor='{instructor_param}'")

    instructor = instructor_param

    if not instructor or instructor not in INSTRUCTORS:
        if instructor_param:
            print(f"[WARN] '{instructor_param}' not found in sheets")
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