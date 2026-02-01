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

# ─── Day ordering ───────────────────────────────────────────────────────────────
DAY_ORDER = {
    'M':   0, 'MO':  0,
    'TU':  1, 'T':   1,
    'W':   2, 'WE':  2,
    'TH':  3, 'R':   3,
    'F':   4,
    'SA':  5, 'S':   5,
    'SU':  6,
}

def extract_day_code(class_name):
    if pd.isna(class_name):
        return 99
    match = re.search(r'\s*([A-Z]{1,2})$', str(class_name).strip().upper())
    if match:
        code = match.group(1)
        return DAY_ORDER.get(code, 99)
    return 99

# ─── Generate table HTML with updated Apple-inspired colors ─────────────────────
def get_table_html(instructor):
    if EXCEL_PATH is None:
        return '<p style="color: red; font-weight: bold; padding: 20px;">No valid Excel file available in /data folder.</p>'

    try:
        if instructor not in INSTRUCTORS:
            return f'<p style="color: red;">Sheet for "{instructor}" not found in Excel.</p>'

        print(f"[DEBUG] Loading sheet: '{instructor}'")

        df = pd.read_excel(EXCEL_PATH, sheet_name=instructor, engine='openpyxl')
        df = df.fillna('')

        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        sort_cols = []
        ascending = []

        if 'Class Name' in df.columns:
            df['sort_day'] = df['Class Name'].apply(extract_day_code)

            sort_cols.append('sort_day')
            ascending.append(True)

            sort_cols.append('Class Name')
            ascending.append(True)

            if 'Date' in df.columns:
                try:
                    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                    sort_cols.append('Date')
                    ascending.append(False)
                except Exception as date_err:
                    print(f"Date parsing failed for {instructor}: {date_err}")

            print(f"[DEBUG] Sorting by: {sort_cols} ascending: {ascending}")
            df = df.sort_values(by=sort_cols, ascending=ascending)
            df = df.drop(columns=['sort_day'], errors='ignore')

        else:
            print(f"Warning: No 'Class Name' column → no day sorting for {instructor}")

        # ─── Apple-inspired muted palette ───────────────────────────────────────
        def row_background(row):
            day_num = extract_day_code(row.get('Class Name', pd.NA))
            colors = {
                0: '#f0f5ff',   # Monday     very pale blue-gray
                1: '#fff4f0',   # Tuesday    extremely pale warm peach
                2: '#f0fff4',   # Wednesday  very pale mint
                3: '#fffaf0',   # Thursday   pale cream / warm off-white
                4: '#f8f0ff',   # Friday     extremely pale lavender
                5: '#f5f9ff',   # Saturday   pale cool blue-white
                6: '#fdfdfd',   # Sunday     almost pure white with tiny warmth
                99: '#f8f8f8'   # Unknown    very light neutral gray
            }
            bg_color = colors.get(day_num, '#ffffff')
            return [f'background-color: {bg_color}'] * len(row)

        styled = df.style.apply(row_background, axis=1)
        styled = styled.set_properties(**{'text-align': 'left'})

        table_html = styled.to_html(
            escape=False,
            index=False,
            classes="table table-striped table-bordered table-hover",
            border=0,
            justify="left"
        )

        return table_html

    except Exception as e:
        traceback.print_exc()
        error_msg = f'<p style="color: red; font-weight: bold; padding: 20px;">Error loading data for {instructor}: {str(e)}</p>'
        if 'df' in locals():
            error_msg += f'<p>Available columns: {", ".join(df.columns.tolist())}</p>'
        return error_msg

# ─── Unified route – no auto-display of first sheet ─────────────────────────────
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def danger_report(path=''):
    instructor_param = request.args.get('instructor', '').strip()

    print(f"[DEBUG] Requested: /{path} ?instructor='{instructor_param}'")

    instructor = None
    table_html = None
    message = None

    if instructor_param:
        if instructor_param in INSTRUCTORS:
            instructor = instructor_param
            table_html = get_table_html(instructor)
        else:
            print(f"[WARN] Requested instructor '{instructor_param}' not found in sheets")
            message = f'''
            <div style="text-align: center; padding: 32px; background: #fff5f5; border-radius: 12px; margin: 32px auto; max-width: 600px; border: 1px solid #ffcccc;">
                <h3 style="color: #c41e3a; margin-bottom: 16px;">Instructor not found</h3>
                <p style="margin-bottom: 20px;">"{instructor_param}" does not match any sheet in the Excel file.</p>
                <p style="color: #555;">Please select a valid instructor from the dropdown above.</p>
            </div>
            '''
    else:
        message = '''
        <div style="text-align: center; padding: 48px 24px; background: #f9f9f9; border-radius: 18px; margin: 48px auto; max-width: 720px; box-shadow: 0 4px 12px rgba(0,0,0,0.06);">
            <h2 style="margin-bottom: 20px; font-size: 28px;">Welcome to Danger Report</h2>
            <p style="font-size: 17px; color: #444; max-width: 580px; margin: 0 auto 28px;">
                Select an instructor from the dropdown menu above to view their class danger report data.
            </p>
            <p style="color: #777; font-size: 15px;">
                Data is loaded from the Excel file located in the /data folder.
            </p>
        </div>
        '''

    updated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    return render_template(
        'danger_report.html',
        instructor=instructor or "Select an Instructor",
        instructors=INSTRUCTORS,
        table_html=table_html,
        updated_at=updated_at,
        message=message or ""
    )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)