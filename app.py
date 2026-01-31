import os
import pandas as pd
from flask import Flask, render_template, jsonify
from openpyxl import load_workbook
from io import BytesIO
import dropbox

app = Flask(__name__)

# Dropbox settings (stored securely in Render)
DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN")
DROPBOX_APP_KEY = os.getenv("DROPBOX_APP_KEY")
DROPBOX_APP_SECRET = os.getenv("DROPBOX_APP_SECRET")

# Path inside your App Folder (correct format)
DROPBOX_FILE_PATH = "/Danger Report Master.xlsm"


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


def extract_day_code(class_name):
    """Extracts the last token from Class Name (e.g., 'Sa', 'M', 'Tu')."""
    if not isinstance(class_name, str):
        return ""
    return class_name.split()[-1]


def load_excel_from_dropbox():
    """
    Downloads the Excel file from Dropbox using refresh-token authentication
    and returns an openpyxl workbook object.
    """
    try:
        print("DEBUG: Using refresh-token Dropbox auth")
        print("DEBUG: Attempting to download:", DROPBOX_FILE_PATH)

        # Dropbox client using permanent refresh-token flow
        dbx = dropbox.Dropbox(
            oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
            app_key=DROPBOX_APP_KEY,
            app_secret=DROPBOX_APP_SECRET
        )

        metadata, res = dbx.files_download(DROPBOX_FILE_PATH)
        file_bytes = res.content

        wb = load_workbook(filename=BytesIO(file_bytes), data_only=True)
        return wb

    except Exception as e:
        print("DROPBOX ERROR (load_excel_from_dropbox):", repr(e))
        raise RuntimeError(f"Dropbox download failed: {repr(e)}")


@app.route("/")
def home():
    return render_template("main.html")


@app.route("/danger-report")
def danger_report():
    return render_template("danger_report.html")


@app.route("/api/get-sheets")
def get_sheets():
    """
    Returns all sheet names EXCEPT any sheet containing 'combined'
    (case-insensitive).
    """
    try:
        wb = load_excel_from_dropbox()

        sheet_names = [s.strip() for s in wb.sheetnames]
        filtered = [s for s in sheet_names if "combined" not in s.lower()]

        return jsonify({"sheets": filtered})

    except Exception as e:
        print("DROPBOX ERROR (/api/get-sheets):", repr(e))
        return jsonify({"error": str(e)}), 500


@app.route("/api/get-sheet-data/<sheet_name>")
def get_sheet_data(sheet_name):
    """
    Returns the data from a specific sheet as JSON, including:
    - __day_code (M, Tu, W, Th, F, Sa, Su)
    - __day_sort (1–7)
    """
    try:
        wb = load_excel_from_dropbox()
        ws = wb[sheet_name]
        data = ws.values
        df = pd.DataFrame(data)

        # First row becomes header
        df.columns = df.iloc[0]
        df = df[1:]

        # Add day code + numeric sort order
        if "Class Name" in df.columns:
            df["__day_code"] = df["Class Name"].apply(extract_day_code)
            df["__day_sort"] = df["__day_code"].map(DAY_ORDER).fillna(999)
        else:
            df["__day_code"] = ""
            df["__day_sort"] = 999

        result = {
            "columns": list(df.columns),
            "rows": df.fillna("").values.tolist()
        }

        return jsonify(result)

    except Exception as e:
        print("DROPBOX ERROR (/api/get-sheet-data):", repr(e))
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)