from flask import Flask, render_template, request
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

# Path to your Excel file (inside repo → works on Render)
EXCEL_PATH = os.path.join(os.path.dirname(__file__), 'data', 'instructors.xlsx')

# Load sheet names once at startup (instructor list)
try:
    excel_file = pd.ExcelFile(EXCEL_PATH)
    INSTRUCTORS = excel_file.sheet_names
    print(f"Loaded {len(INSTRUCTORS)} instructors from Excel: {INSTRUCTORS}")
except Exception as e:
    print(f"ERROR loading Excel at startup: {e}")
    INSTRUCTORS = ["Error loading sheets"]

def get_table_html(instructor):
    try:
        if instructor not in INSTRUCTORS:
            return f'<p style="color: red;">Instructor "{instructor}" not found in Excel sheets.</p>'
        
        # Read only the requested sheet
        df = pd.read_excel(EXCEL_PATH, sheet_name=instructor, engine='openpyxl')
        
        # Optional: clean up data if needed (e.g. fill NaN, convert types)
        df = df.fillna('')  # or handle missing values your way
        
        # Optional: your custom sorting or calculations
        # Example: sort by date descending if 'Date' column exists
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.sort_values('Date', ascending=False)
        
        # Convert to HTML with Bootstrap-friendly classes
        table_html = df.to_html(
            index=False,
            classes="table table-striped table-bordered table-hover",
            border=0,
            justify="left",
            escape=False
        )
        return table_html
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return f'<p style="color: red; font-weight: bold;">Error loading data for {instructor}: {str(e)}</p>'

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
    return render_template('danger_report.html',  # or a simple redirect
                           instructor=INSTRUCTORS[0] if INSTRUCTORS else "",
                           instructors=INSTRUCTORS,
                           table_html="<p>Welcome — select an instructor above.</p>",
                           updated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

if __name__ == '__main__':
    app.run(debug=True)