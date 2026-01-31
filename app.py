import os
import pandas as pd
from flask import Flask, render_template, jsonify
from openpyxl import load_workbook
from io import BytesIO
import dropbox

app = Flask(__name__)

# Dropbox settings
DROPBOX_TOKEN = os.getenv("DROPBOX_TOKEN")  # Stored securely in Render
DROPBOX_FILE_PATH = "Apps/Danger Report Master.xlsm"  # Path inside Dropbox


def load_excel_from_dropbox():
    """
    Downloads the Excel file from Dropbox using the Dropbox API
    and returns an openpyxl workbook object.
    """
    try:
        # Debug: confirm token exists
        print("DEBUG: DROPBOX_TOKEN exists:", DROPBOX_TOKEN is not None)

        # Debug: show path being requested
        print("DEBUG: Attempting to download:", DROPBOX_FILE_PATH)

        dbx = dropbox.Dropbox(DROPBOX_TOKEN)

        # Attempt download
        metadata, res = dbx.files_download(DROPBOX_FILE_PATH)
        file_bytes = res.content

        wb = load_workbook(filename=BytesIO(file_bytes), data_only=True)
        return wb

    except Exception as e:
        # Full debug output
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
    Returns all sheet names EXCEPT any sheet that contains the word 'combined'
    (case-insensitive, handles trailing spaces, hidden characters, etc.)
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
    Returns the data from a specific sheet as JSON.
    """
    try:
        wb = load_excel_from_dropbox()

        ws = wb[sheet_name]
        data = ws.values
        df = pd.DataFrame(data)

        # First row becomes header
        df.columns = df.iloc[0]
        df = df[1:]

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
