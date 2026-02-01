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
    # Prefer modern Excel formats first
    excel_files.sort(key=lambda f: (not f.lower().endswith(('.xlsx', '.xlsm')), f))
    EXCEL_PATH = excel_files[0]
    print(f"Using Excel file: {EXCEL_PATH}")
    try:
        excel_file = pd.ExcelFile(EXCEL_PATH)
        raw_sheets = excel_file.sheet_names
        # Skip first sheet ("Combined" or summary) completely
        selectable_sheets = raw_sheets[1:] if len(raw_sheets) > 1 else []
        # Sort alphabetically for dropdown
        INSTRUCTORS = sorted(selectable_sheets)
        print(f"Loaded {len(INSTRUCTORS)} selectable instructors (first skipped, sorted): {INSTRUCTORS}")
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
    """Extract a numeric day code from the end of the Class Name string.

    Expected patterns (case-insensitive, trailing):
    - "Ballet 1 M"
    - "Tap 2 TU"
    - "Jazz 3 W"
    - "Hip Hop TH"
    - "Company SA"
    - "Rehearsal SU"

    Returns an integer 0–6 for known days, or 99 as a fallback.
    """
    if pd.isna(class_name):
        return 99
    match = re.search(r'\s*([A-Z]{1,2})$', str(class_name).strip().upper())
    if match:
        code = match.group(1)
        return DAY_ORDER.get(code, 99)
    return 99


# ─── Table generation ───────────────────────────────────────────────────────────

def get_table_html(instructor: str) -> str:
    """Generate styled HTML table for a given instructor sheet.

    - Loads the sheet into a DataFrame
    - Normalizes column names
    - Sorts by day (from Class Name) and then alphabetically within each day
    - Applies row background colors based on day
    - Returns HTML for direct embedding in the template
    """
    if EXCEL_PATH is None:
        return '<p style="color: red; font-weight: bold; padding: 20px;">No valid Excel file available in /data folder.</p>'

    try:
        if instructor not in INSTRUCTORS:
            return f'<p style="color: red;">Instructor/sheet "{instructor}" not found or restricted.</p>'

        print(f"[DEBUG] Loading sheet: '{instructor}'")

        df = pd.read_excel(EXCEL_PATH, sheet_name=instructor, engine='openpyxl')
        df = df.fillna('')

        # Normalize column names: strip spaces, collapse internal whitespace
        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        # ─── Sorting ─────────────────────────────────────────────────────────
        sort_keys = []
        ascending_flags = []

        # Primary: sort by day extracted from Class Name
        if 'Class Name' in df.columns:
            df['sort_day'] = df['Class Name'].apply(extract_day_code)
            sort_keys.append('sort_day')
            ascending_flags.append(True)

            # Secondary: alphabetical by Class Name within each day
            sort_keys.append('Class Name')
            ascending_flags.append(True)

        # Apply sorting if we have keys
        if sort_keys:
            print(f"[DEBUG] Sorting by {sort_keys} asc={ascending_flags}")
            df = df.sort_values(by=sort_keys, ascending=ascending_flags)

        # Clean temp columns
        if 'sort_day' in df.columns:
            df = df.drop(columns=['sort_day'], errors='ignore')

        # ─── Row coloring based on day ──────────────────────────────────────
        def row_background(row):
            day_num = extract_day_code(row.get('Class Name', pd.NA))
            colors = {
                0: '#e8f0ff',  # Monday
                1: '#fff0e8',  # Tuesday
                2: '#e8fff0',  # Wednesday
                3: '#fff8e8',  # Thursday
                4: '#f0e8ff',  # Friday
                5: '#e8f4ff',  # Saturday
                6: '#f8f8f8',  # Sunday
                99: '#f0f0f0', # Unknown
            }
            bg_color = colors.get(day_num, '#ffffff')
            return [f'background-color: {bg_color}'] * len(row)

        styled = df.style.apply(row_background, axis=1)
        styled = styled.set_properties(**{'text-align': 'left'})
	styled = styled.hide(axis="index")

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


# ─── Route ──────────────────────────────────────────────────────────────────────

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def danger_report(path=''):
    instructor_param = request.args.get('instructor', '').strip()

    print(f"[DEBUG] Requested: /{path} ?instructor='{instructor_param}'")

    instructor = None
    table_html = None
    message = None

    if instructor_param and instructor_param in INSTRUCTORS:
        instructor = instructor_param
        table_html = get_table_html(instructor)
    else:
        if instructor_param:
            print(f"[WARN] '{instructor_param}' not found or invalid")
            message = (
                '<div style="text-align:center; padding:32px; background:#fff5f5; '
                'border-radius:12px; margin:32px auto; max-width:600px; border:1px solid #ffcccc;">'
                '<h3 style="color:#c41e3a; margin-bottom:16px;">Not found</h3>'
                f'<p>"{instructor_param}" does not match any available instructor sheet.</p>'
                '<p style="color:#555;">Please select from the dropdown.</p>'
                '</div>'
            )
        else:
            message = (
                '<div style="text-align:center; padding:60px 24px; background:#f9f9f9; '
                'border-radius:18px; margin:48px auto; max-width:720px; '
                'box-shadow:0 4px 12px rgba(0,0,0,0.06);">'
                '<h2 style="margin-bottom:20px; font-size:28px;">Heads Up Report</h2>'
                '<p style="font-size:17px; color:#444; max-width:580px; margin:0 auto 28px;">'
                'Select an instructor from the dropdown menu above to see their list.'
                '</p>'
                '</div>'
            )

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