import pandas as pd
from flask import Flask, render_template, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

# Path to your Excel file
EXCEL_PATH = r"C:\Users\DiegoMonroy\Ldesoccer Dropbox\Operations\Drop Danger\Danger Report Master.xlsm"


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
        wb = load_workbook(EXCEL_PATH, data_only=True)

        # Normalize sheet names
        sheet_names = [s.strip() for s in wb.sheetnames]

        # Remove ANY sheet containing "combined"
        filtered = [s for s in sheet_names if "combined" not in s.lower()]

        return jsonify({"sheets": filtered})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/get-sheet-data/<sheet_name>")
def get_sheet_data(sheet_name):
    """
    Returns the data from a specific sheet as JSON.
    """
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)

        data = {
            "columns": list(df.columns),
            "rows": df.fillna("").values.tolist()
        }

        return jsonify(data)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)