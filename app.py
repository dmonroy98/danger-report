import os
import pandas as pd
from flask import Flask, render_template, jsonify, request, redirect, flash, make_response
from openpyxl import load_workbook

app = Flask(__name__)

# Secret key for flash messages
app.secret_key = os.getenv("FLASK_SECRET_KEY", "dev-secret")

# Password stored securely in Render
UPLOAD_PASSWORD = os.getenv("UPLOAD_PASSWORD")

# Local file path
LOCAL_FILE_PATH = "data/Danger Report Master.xlsm"

# Day mapping for sorting
DAY_ORDER = {
    "M": 1,
    "Tu": 2,
    "W": 3,
    "Th": 4,
    "F": 5,
    "Sa": 6,
    "Su": 7
}

# Row color classes
DAY_COLOR = {
    "M": "day-m",
    "Tu": "day-tu",
    "W": "day-w",
    "Th": "day-th",
    "F": "day-f",
    "Sa": "day-sa",
    "Su": "day-su"
}


def extract_day_code(class_name):
    if not isinstance(class_name, str):
        return ""
    return class_name.split()[-1]


def load_excel_from_local():
    wb = load_workbook(filename=LOCAL_FILE_PATH, data_only=True)
    return wb


@app.route("/")
def home():
    return render_template("main.html")


# ---------------------------------------------------------
# MAIN DANGER REPORT PAGE (with dropdown + back button)
# ---------------------------------------------------------
@app.route("/danger-report", methods=["GET", "POST"])
def danger_report():
    try:
        wb = load_excel_from_local()
    except Exception as e:
        return f"Error loading Excel file: {e}. Please upload a file."

    # 1. Clean sheet names
    sheet_names = [str(s).strip() for s in wb.sheetnames]
    instructors = [s for s in sheet_names if "combined" not in s.lower()]

    # Default instructor (fallback)
    current = instructors[0] if instructors else None

    # 2. Check URL (GET) or Dropdown (POST)
    selected = request.args.get("instructor") or request.form.get("instructor")

    if selected:
        if selected in instructors:
            print(f"DEBUG: Request for '{selected}' - MATCH FOUND.")
            current = selected
        else:
            print(f"DEBUG: Request for '{selected}' - NO MATCH. Available: {instructors[:3]}...")

    # Load data
    table_html = "<p>No data loaded.</p>"
    
    if current:
        ws = wb[current]
        data = list(ws.values)

	if current:
        ws = wb[current]
        data = list(ws.values)
        
        # --- NEW DEBUG LINES ---
        print(f"DEBUG: Sheet '{current}' Total Rows Read: {len(data)}")
        if len(data) > 1:
            print(f"DEBUG: Sample Row 2 Data: {data[1]}")
        else:
            print("DEBUG: Sheet appears to contain ONLY the header row!")
        # -----------------------

   ..	
        
        # --- ROBUST HEADER FINDER ---
        # We look for a row that contains "Class Name" AND "Instructors"
        header_index = -1
        for i, row in enumerate(data[:10]): # Scan first 10 rows
            # Convert row to string, strip whitespace, handle None
            row_clean = [str(cell).strip() for cell in row if cell is not None]
            
            # Check if our key columns exist in this row
            if "Class Name" in row_clean and "Instructors" in row_clean:
                header_index = i
                print(f"DEBUG: Found Headers on Row {i+1}: {row_clean}")
                break
        
        if header_index != -1:
            # Create DataFrame starting from the detected header row
            df = pd.DataFrame(data)
            df.columns = df.iloc[header_index] # Set headers
            df = df[header_index + 1:]         # Keep data after headers
            
            # Clean Column Names (remove spaces like "Class Name ")
            df.columns = [str(c).strip() for c in df.columns]
            
            # --- DATA PROCESSING ---
            # 1. Normalize Instructor Column
            if "Instructors" in df.columns:
                df["Instructors"] = df["Instructors"].astype(str).str.strip()
            
            # 2. Add Day Sorting
            if "Class Name" in df.columns:
                df["__day_code"] = df["Class Name"].apply(extract_day_code)
                df["__day_sort"] = df["__day_code"].map(DAY_ORDER).fillna(999)
                df["__day_color"] = df["__day_code"].map(DAY_COLOR).fillna("")
            else:
                # Fallback if column is missing despite our check
                df["__day_code"] = ""
                df["__day_sort"] = 999
                df["__day_color"] = ""

            # 3. Create Table
            table_html = df.to_html(classes="danger-table", index=False)
        else:
            print(f"DEBUG: CRITICAL ERROR - Could not find headers in sheet '{current}'")
            table_html = f"<p>Error: Could not find headers ('Class Name', 'Instructors') in the first 10 rows of sheet '{current}'.</p>"
    
    # Create response with No-Cache headers to ensure you see changes
    response = make_response(render_template(
        "danger_report.html",
        instructors=instructors,
        current=current,
        table=table_html
    ))
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return response

# ---------------------------------------------------------
# API ENDPOINTS
# ---------------------------------------------------------
@app.route("/api/get-sheets")
def get_sheets():
    try:
        wb = load_excel_from_local()
        sheet_names = [s.strip() for s in wb.sheetnames]
        filtered = [s for s in sheet_names if "combined" not in s.lower()]
        return jsonify({"sheets": filtered})
    except Exception as e:
        print("LOCAL FILE ERROR (/api/get-sheets):", repr(e))
        return jsonify({"error": str(e)}), 500


@app.route("/api/get-sheet-data/<sheet_name>")
def get_sheet_data(sheet_name):
    try:
        wb = load_excel_from_local()
        ws = wb[sheet_name]
        data = ws.values
        df = pd.DataFrame(data)

        # Basic cleanup for API (You might want to apply the Robust Header Finder here too eventually)
        df.columns = df.iloc[0]
        df = df[1:]

        # Normalize instructor column
        if "Instructors" in df.columns:
            df["Instructors"] = df["Instructors"].astype(str).str.strip()

        if "Class Name" in df.columns:
            df["__day_code"] = df["Class Name"].apply(extract_day_code)
            df["__day_sort"] = df["__day_code"].map(DAY_ORDER).fillna(999)
            df["__day_color"] = df["__day_code"].map(DAY_COLOR).fillna("")
        else:
            df["__day_code"] = ""
            df["__day_sort"] = 999
            df["__day_color"] = ""

        result = {
            "columns": list(df.columns),
            "rows": df.fillna("").values.tolist()
        }

        return jsonify(result)

    except Exception as e:
        print("LOCAL FILE ERROR (/api/get-sheet-data):", repr(e))
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------
# PASSWORD-PROTECTED UPLOAD
# ---------------------------------------------------------
@app.route("/upload")
def upload_page():
    return render_template("upload.html")


@app.route("/upload-file", methods=["POST"])
def upload_file():
    password = request.form.get("password")
    file = request.files.get("file")

    # 1. Validate password
    if password != UPLOAD_PASSWORD:
        flash("Invalid password", "error")
        return redirect("/upload")

    # 2. Validate file presence
    if not file:
        flash("No file uploaded", "error")
        return redirect("/upload")

    # 3. Validate file extension
    filename = file.filename.lower()
    if not filename.endswith((".xlsm", ".xlsx")):
        flash("Invalid file type. Must be .xlsm or .xlsx", "error")
        return redirect("/upload")

    # 4. Ensure data directory exists
    data_dir = os.path.dirname(LOCAL_FILE_PATH)
    os.makedirs(data_dir, exist_ok=True)

    # 5. Optional: Backup old file
    if os.path.exists(LOCAL_FILE_PATH):
        backup_path = LOCAL_FILE_PATH + ".backup"
        try:
            os.replace(LOCAL_FILE_PATH, backup_path)
        except Exception as e:
            print("Backup failed:", e)

    # 6. Save new file
    try:
        file.save(LOCAL_FILE_PATH)
    except Exception as e:
        flash(f"Failed to save file: {e}", "error")
        return redirect("/upload")

    # 7. Success message
    flash("File uploaded successfully!", "success")
    return redirect("/danger-report")


if __name__ == "__main__":
    app.run(debug=True)