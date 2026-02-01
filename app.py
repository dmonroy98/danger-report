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
    print("ERROR: No Excel file found")
else:
    excel_files.sort(key=lambda f: (not f.lower().endswith(('.xlsx', '.xlsm')), f))
    EXCEL_PATH = excel_files[0]
    print(f"Using Excel file: {EXCEL_PATH}")
    try:
        excel_file = pd.ExcelFile(EXCEL_PATH)
        INSTRUCTORS = excel_file.sheet_names
        print(f"Loaded {len(INSTRUCTORS)} sheets: {INSTRUCTORS}")
    except Exception as e:
        print(f"ERROR loading Excel: {e}")
        INSTRUCTORS = ["Excel read failed"]
        EXCEL_PATH = None

# ─── Day ordering ───────────────────────────────────────────────────────────────
DAY_ORDER = {
    'M': 0, 'MO': 0,
    'TU': 1, 'T': 1,
    'W': 2, 'WE': 2,
    'TH': 3, 'R': 3,
    'F': 4,
    'SA': 5, 'S': 5,
    'SU': 6,
}

def extract_day_code(class_name):
    if pd.isna(class_name):
        return 99
    match = re.search(r'\s*([A-Z]{1,2})$', str(class_name).strip().upper())
    if match:
        return DAY_ORDER.get(match.group(1), 99)
    return 99

# ─── Table generation ───────────────────────────────────────────────────────────
def get_table_html(instructor):
    if EXCEL_PATH is None:
        return '<p style="color:red; padding:20px;">No Excel file available.</p>'

    try:
        if instructor not in INSTRUCTORS:
            return f'<p style="color:red; padding:20px;">Sheet "{instructor}" not found.</p>'

        print(f"[DEBUG] Loading sheet: '{instructor}'")

        df = pd.read_excel(EXCEL_PATH, sheet_name=instructor, engine='openpyxl')
        df = df.fillna('')

        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        # ─── Sorting ────────────────────────────────────────────────────────────
        sort_keys = []
        ascending_flags = []

        # Day-based sort (primary)
        if 'Class Name' in df.columns:
            df['sort_day'] = df['Class Name'].apply(extract_day_code)
            sort_keys.append('sort_day')
            ascending_flags.append(True)

        # First column numeric sort if it looks like a count/number
        first_col = df.columns[0] if len(df.columns) > 0 else None
        if first_col and first_col.lower() in ['count', '#', 'students', 'enrolled', 'incidents', 'number', 'rank']:
            try:
                df['sort_first'] = pd.to_numeric(df[first_col], errors='coerce')
                sort_keys.append('sort_first')
                ascending_flags.append(True)  # change to False if you want descending
            except:
                pass  # keep as string if conversion fails

        # Class Name alphabetical (secondary)
        if 'Class Name' in df.columns:
            sort_keys.append('Class Name')
            ascending_flags.append(True)

        # Date descending (if present)
        if 'Date' in df.columns:
            try:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                sort_keys.append('Date')
                ascending_flags.append(False)
            except:
                pass

        if sort_keys:
            print(f"[DEBUG] Sorting by {sort_keys} asc={ascending_flags}")
            df = df.sort_values(by=sort_keys, ascending=ascending_flags)

        # Clean temp columns
        for col in ['sort_day', 'sort_first']:
            if col in df.columns:
                df = df.drop(columns=[col])

        # ─── Row colors (muted Apple palette) ──────────────────────────────────
        def row_background(row):
            day_num = extract_day_code(row.get('Class Name', pd.NA))
            colors = {
                0: '#f0f5ff',  # Monday
                1: '#fff4f0',  # Tuesday
                2: '#f0fff4',  # Wednesday
                3: '#fffaf0',  # Thursday
                4: '#f8f0ff',  # Friday
                5: '#f5f9ff',  # Saturday
                6: '#fdfdfd',  # Sunday
                99: '#f8f8f8'  # Unknown
            }
            return [f'background-color: {colors.get(day_num, "#ffffff")}'] * len(row)

        styled = df.style.apply(row_background, axis=1)
        styled = styled.set_properties(**{'text-align': 'left'})

        return styled.to_html(
            escape=False,
            index=False,
            classes="table table-striped table-bordered table-hover",
            border=0,
            justify="left"
        )

    except Exception as e:
        traceback.print_exc()
        msg = f'<p style="color:red; padding:20px;">Error loading "{instructor}": {str(e)}</p>'
        if 'df' in locals():
            msg += f'<p>Columns: {", ".join(df.columns.tolist())}</p