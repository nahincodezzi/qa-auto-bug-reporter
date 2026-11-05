import os
import sys
import threading
import time
import webbrowser
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
import openpyxl

# Import core business logic from main.py
from main import (
    upload_image_bytes,
    convert_problem_to_qa_report,
    append_to_bug_report,
    delete_columns_from_excel,
    delete_row_from_excel,
    EXCEL_FILE,
    COLUMNS
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB

def read_all_reports_from_excel(file_path: str = EXCEL_FILE) -> tuple:
    """Reads current headers and all rows dynamically from bug_report.xlsx."""
    p = Path(file_path)
    if not p.exists():
        return COLUMNS, []

    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    if ws.max_row < 1:
        return COLUMNS, []

    headers = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]
    while headers and not headers[-1]:
        headers.pop()
    if not headers:
        headers = COLUMNS

    reports = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if not any(vals):
            continue

        row_dict = {
            "row_num": r,
            "cells": [str(v) if v is not None else "" for v in vals]
        }
        reports.append(row_dict)

    return headers, reports

@app.route("/")
def index():
    headers, reports = read_all_reports_from_excel()
    return render_template("index.html", headers=headers, reports=reports)

@app.route("/api/reports", methods=["GET"])
def api_reports():
    headers, reports = read_all_reports_from_excel()
    return jsonify({"success": True, "headers": headers, "reports": reports})

@app.route("/api/add_bug", methods=["POST"])
def api_add_bug():
    problem_text = request.form.get("problem", "").strip()
    if not problem_text:
        return jsonify({"success": False, "error": "Problem description cannot be empty."}), 400

    # Support multiple screenshot uploads
    image_files = request.files.getlist("images") or request.files.getlist("image")
    cloud_urls = []
    first_img_bytes = None
    first_mime = "image/png"

    for idx, file in enumerate(image_files, 1):
        if file and file.filename != "":
            img_bytes = file.read()
            if img_bytes:
                if first_img_bytes is None:
                    first_img_bytes = img_bytes
                    first_mime = file.mimetype or "image/png"
                url = upload_image_bytes(img_bytes, file.filename or f"screenshot_{idx}.png")
                if url and url != "Upload failed":
                    cloud_urls.append(url)

    if len(cloud_urls) == 1:
        cloud_url = cloud_urls[0]
    elif len(cloud_urls) > 1:
        cloud_url = "\n".join([f"SS {i}: {u}" for i, u in enumerate(cloud_urls, 1)])
    else:
        cloud_url = "-"

    # Structure into 100% professional QA format using AI / Smart Engine
    qa_data = convert_problem_to_qa_report(
        problem_text=problem_text,
        image_bytes=first_img_bytes,
        mime_type=first_mime
    )
    qa_data["screenshort"] = cloud_url

    # Append to Excel with PermissionError protection
    try:
        tid = append_to_bug_report(qa_data, EXCEL_FILE)
    except PermissionError as pe:
        return jsonify({
            "success": False,
            "error": "⚠️ 'bug_report.xlsx' is currently open in Microsoft Excel! Please close the Excel window and click the button again."
        }), 400
    except Exception as e:
        return jsonify({"success": False, "error": f"Failed to save to Excel: {e}"}), 500

    reports = read_all_reports_from_excel()[1]
    headers, reports = read_all_reports_from_excel()
    return jsonify({
        "success": True,
        "tid": tid,
        "data": qa_data,
        "headers": headers,
        "reports": reports
    })

@app.route("/api/delete_columns", methods=["POST"])
def api_delete_columns():
    data = request.json or {}
    indices = data.get("column_indices", [])
    if not indices:
        return jsonify({"success": False, "error": "No column indices specified."}), 400

    try:
        updated_headers = delete_columns_from_excel(indices)
        headers, reports = read_all_reports_from_excel()
        return jsonify({
            "success": True,
            "message": f"Successfully deleted {len(indices)} column(s).",
            "headers": headers,
            "reports": reports
        })
    except PermissionError as pe:
        return jsonify({
            "success": False,
            "error": "⚠️ 'bug_report.xlsx' is currently open in Microsoft Excel! Please close Excel and try again."
        }), 400
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/delete_row", methods=["POST"])
def api_delete_row():
    data = request.json or {}
    row_num = data.get("row_num")
    if not row_num:
        return jsonify({"success": False, "error": "Row number is required."}), 400

    try:
        success = delete_row_from_excel(int(row_num))
        headers, reports = read_all_reports_from_excel()
        return jsonify({
            "success": success,
            "message": f"Row {row_num} deleted.",
            "headers": headers,
            "reports": reports
        })
    except PermissionError as pe:
        return jsonify({
            "success": False,
            "error": "⚠️ 'bug_report.xlsx' is currently open in Microsoft Excel! Please close Excel and try again."
        }), 400
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/set_api_key", methods=["POST"])
def api_set_api_key():
    data = request.json or {}
    key = data.get("api_key", "").strip()
    os.environ["GEMINI_API_KEY"] = key

    # Save to .env file
    try:
        from dotenv import set_key
        env_path = Path(".env")
        if not env_path.exists():
            env_path.write_text(f"GEMINI_API_KEY={key}\n", encoding="utf-8")
        else:
            set_key(str(env_path), "GEMINI_API_KEY", key)
        return jsonify({"success": True, "message": "Gemini API Key saved successfully!"})
    except Exception as e:
        return jsonify({"success": True, "message": f"Key set in session ({e})"})

@app.route("/api/open_excel", methods=["POST"])
def api_open_excel():
    p = Path(EXCEL_FILE).resolve()
    if p.exists():
        try:
            os.startfile(str(p))
            return jsonify({"success": True, "message": "Excel file opened."})
        except Exception as e:
            return jsonify({"success": False, "message": f"Could not launch Excel: {e}"})
    return jsonify({"success": False, "message": "Excel file does not exist yet."})

@app.route("/api/download", methods=["GET"])
def api_download():
    p = Path(EXCEL_FILE).resolve()
    if p.exists():
        return send_file(str(p), as_attachment=True, download_name=EXCEL_FILE)
    return "Excel file not found", 404

def open_browser():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:5000")

if __name__ == "__main__":
    print("=" * 65)
    print("      🚀 QA Bug Reporter Pro (Desktop UI Running)")
    print("=" * 65)
    print("👉 Opening http://127.0.0.1:5000 in your browser...")
    print("👉 Press Ctrl+C in this terminal if you ever want to close.")

    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=5000, debug=False)

