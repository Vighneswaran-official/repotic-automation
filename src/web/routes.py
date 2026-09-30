"""
src/web/routes.py - Flask web application routes and controller endpoints.
"""

import base64
import traceback
from pathlib import Path
from flask import Flask, make_response, request, jsonify, render_template

from src.config import PROJECT_ROOT, get_writable_dir
from src.services.processor import run_automation

# Initialize Flask with explicit templates and static folders relative to project root
app = Flask(
    __name__,
    template_folder=str(PROJECT_ROOT / "templates"),
    static_folder=str(PROJECT_ROOT / "static"),
    static_url_path="/static"
)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


@app.route("/")
def index():
    """Serve the primary HTML application page with no-cache headers."""
    resp = make_response(render_template("index.html"))
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
    resp.headers["Pragma"]        = "no-cache"
    return resp


@app.route("/run", methods=["POST"])
def run():
    """Accept 3 uploaded files, run automation, return JSON result."""
    try:
        repotic_file  = request.files.get("repotic")
        ledger_file   = request.files.get("ledger")
        template_file = request.files.get("template")

        missing = []
        if not repotic_file  or not repotic_file.filename:  missing.append("REPOTIC.xlsx")
        if not ledger_file   or not ledger_file.filename:   missing.append("KK_TN_Sales_Ledger_and_Party_Ledger.xlsx")
        if not template_file or not template_file.filename: missing.append("Template .xlsx")
        if missing:
            return jsonify({"error": f"Missing files: {', '.join(missing)}"}), 400

        output_dir = get_writable_dir()
        fixed_output = output_dir / "Final_Output_Tamilnadu.xlsx"

        # Clear stale output so /download never serves an old file
        if fixed_output.exists():
            try:
                fixed_output.unlink()
            except Exception:
                pass

        # Save uploads to writable locations (handles serverless read-only filesystem)
        repotic_path  = output_dir / "upload_repotic.xlsx"
        ledger_path   = output_dir / "upload_ledger.xlsx"
        template_path = output_dir / "upload_template.xlsx"

        repotic_file.save(repotic_path)
        ledger_file.save(ledger_path)
        template_file.save(template_path)

        result = run_automation(repotic_path, ledger_path, template_path, fixed_output)

        file_base64 = None
        if fixed_output.exists():
            with open(fixed_output, "rb") as fh:
                file_base64 = base64.b64encode(fh.read()).decode("utf-8")

        return jsonify({"ok": True, "file_base64": file_base64, **result})

    except Exception:
        return jsonify({"error": traceback.format_exc()}), 500


@app.route("/download")
def download():
    """
    Serve Final_Output_Tamilnadu.xlsx as a file download.
    Uses make_response(bytes) + explicit Content-Disposition.
    """
    output_dir = get_writable_dir()
    fixed_output = output_dir / "Final_Output_Tamilnadu.xlsx"

    if not fixed_output.exists():
        return jsonify({
            "error": "No output file found. Please run the automation first."
        }), 404

    with open(fixed_output, "rb") as fh:
        data = fh.read()

    resp = make_response(data)
    resp.headers["Content-Type"]        = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    resp.headers["Content-Disposition"] = (
        'attachment; filename="Final_Output_Tamilnadu.xlsx"'
    )
    resp.headers["Content-Length"]      = len(data)
    resp.headers["Cache-Control"]       = "no-store"
    return resp
