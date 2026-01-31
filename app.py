import os
import pandas as pd
from flask import Flask, render_template, jsonify, request, redirect, flash
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
    wb = load_excel_from_local()

    # 1. Clean sheet names (strip whitespace)
    sheet_names = [str(s).strip() for s in wb.sheetnames]
    instructors = [s for s in sheet_names if "combined" not in s.lower()]

    # Default instructor
    current = instructors[0] if instructors else None

    # 2. Check URL (GET) or Dropdown (POST)
    selected = request.args.get("instructor") or request.form.get("instructor")

    # 3. DEBUG: Print to Render logs to see why it might fail
    if selected:
        print(f"DEBUG: User requested '{selected}'")
        if selected in instructors:
            print(f"DEBUG: Found match! Switching to '{selected}'")
            current = selected
        else:
            print(f"DEBUG: No exact match found. Available sheets: {instructors}")

    # Load data for the determined 'current' instructor
    if current:
        ws = wb[current]
        data = ws.values
        df = pd.DataFrame(data)

        # Set headers from the first row
        df.columns = df.iloc[0]
        df = df[1:]

        # 4. FIX: Clean headers (remove hidden spaces like "Class Name ")
        df.columns = [str(c).strip() for c in df.columns]

        # Normalize instructor column
        if "Instructors" in df.columns:
            df["Instructors"] = df["Instructors"].astype(str).str.strip()

        # Add day sorting + color
        # This will now work even if your excel header was "Class Name " (with a space)
        if "Class Name" in df.columns:
            df["__day_code"] = df["Class Name"].apply(extract_day_code)
            df["__day_sort"] = df["__day_code"].map(DAY_ORDER).fillna(999)
            df["__day_color"] = df["__day_code"].map(DAY_COLOR).fillna("")
        else:
            # Fallback if "Class Name" is still missing
            df["__day_code"] = ""
            df["__day_sort"] = 999
            df["__day_color"] = ""

        table_html = df.to_html(classes="danger-table", index=False)
    else:
        table_html = "<p>No instructor data found.</p>"

    return render_template(
        "danger_report.html",
        instructors=instructors,
        current=current,
        table=table_html
    )

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