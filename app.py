from flask import Flask, render_template, request
import pandas as pd
import os
import glob
from datetime import datetime
import traceback

app = Flask(__name__)

# ─── Configuration ──────────────────────────────────────────────────────────────
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')

# Find any .xlsx file in data/ (use the first one found)
xlsx_files = glob.glob(os.path.join(DATA_DIR, '*.xlsx'))

if not xlsx_files:
    EXCEL_PATH = None
    INSTRUCTORS = ["No Excel file found in /data"]
    print("ERROR: No .xlsx file found in data/ folder")
    print(f"Expected path example: {os.path.join(DATA_DIR, 'yourfile.xlsx')}")
else:
    EXCEL_PATH = xlsx_files[0]  # take the first .xlsx file
    print(f"Using Excel file: {EXCEL_PATH}")
    try:
        excel_file = pd.ExcelFile(EXCEL_PATH)
        INSTRUCTORS = excel_file.sheet_names
        print(f"Successfully loaded {len(INSTRUCTORS)} instructors/sheets: {INSTRUCTORS}")
    except Exception as e:
        print(f"ERROR loading Excel file {EXCEL_PATH}: {e}")
        INSTRUCTORS = ["Excel file found but cannot be read"]
        EXCEL_PATH = None

# ─── Helper to generate table HTML ─────────────────────────────────────────────
def get_table_html(instructor):
    if EXCEL_PATH is None:
        return '<p style="color: red; font-weight: bold; padding: 20px;">No valid Excel file available in /data folder.</p>'

    try:
        if instructor not in INSTRUCTORS:
            return f'<p style="color: red;">Sheet for "{instructor}" not found in Excel.</p>'

        # Read the specific sheet
        df = pd.read_excel(EXCEL_PATH, sheet_name=instructor, engine='openpyxl')

        # Clean up
        df = df.fillna('')

        # Sort by Date descending if column exists
        if 'Date' in df.columns:
            try:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                df = df.sort_values('Date', ascending=False)
            except Exception as sort_err:
                print(f"Date sorting failed for {instructor}: {sort_err}")

        # Generate nice Bootstrap table
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
        return f'<p style="color: red; font-weight: bold; padding: 20px;">Error loading data for {instructor}: {str(e)}</p>'

# ─── Routes ────────────────────────────────────────────────────────────────────
@app.route('/danger-report')
def danger_report():
    instructor = request.args.get('instructor', '').strip()

    if not instructor or instructor not in INSTRUCTORS:
        instructor = INSTRUCTORS[0] if INSTRUCTORS else "No Data Available"

    table_html = get_table_html(instructor)
    updated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    return render_template(
        'danger_report.html',
        instructor=instructor,
        instructors=INSTRUCTORS,
        table_html=table_html,
        updated_at=updated_at
    )


@app.route('/')
def index():
    # For root URL, show the danger-report page with first instructor (or welcome message)
    instructor = INSTRUCTORS[0] if INSTRUCTORS and INSTRUCTORS[0] != "No Excel file found in /data" else ""
    table_html = "<p style='padding: 20px;'>Welcome to Danger Report — please select an instructor above.</p>"

    if instructor and instructor not in ["No Excel file found in /data", "Excel file found but cannot be read"]:
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