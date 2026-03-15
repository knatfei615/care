"""Flask application – PICU pharmacy monitoring web interface."""

from __future__ import annotations

import base64
import binascii
import hmac
from datetime import datetime

from flask import Flask, Response, after_this_request, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from config import (
    BASIC_AUTH_PASS,
    BASIC_AUTH_USER,
    DATA_DIR,
    MAX_UPLOAD_MB,
    OPENAI_API_KEY,
    OPENAI_BASE_URL,
    OPENAI_MODEL,
    PORT,
)
from excel_io import (
    backup_workbook_identifiers,
    ExcelError,
    find_workbook,
    get_slot_status,
    list_patients,
    prepare_download_workbook,
    redact_workbook_identifiers,
    redact_workbooks,
    save_note,
)
from llm import structure_note

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


def _unauthorized_response() -> Response:
    return Response(
        "Unauthorized",
        401,
        {"WWW-Authenticate": 'Basic realm="PICU Pharmacy"'},
    )


def _is_authorized() -> bool:
    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Basic "):
        return False

    token = auth_header.split(" ", 1)[1].strip()
    if not token:
        return False

    try:
        decoded = base64.b64decode(token).decode("utf-8")
    except (ValueError, binascii.Error, UnicodeDecodeError):
        return False

    username, sep, password = decoded.partition(":")
    if not sep:
        return False

    return hmac.compare_digest(username, BASIC_AUTH_USER) and hmac.compare_digest(password, BASIC_AUTH_PASS)


@app.before_request
def require_basic_auth():
    if request.path == "/healthz":
        return None

    if bool(BASIC_AUTH_USER) ^ bool(BASIC_AUTH_PASS):
        return jsonify(error="服务器鉴权配置不完整，请同时设置 BASIC_AUTH_USER 和 BASIC_AUTH_PASS。"), 500

    if not BASIC_AUTH_USER and not BASIC_AUTH_PASS:
        return jsonify(error="服务器未启用访问控制。"), 500

    if _is_authorized():
        return None

    return _unauthorized_response()


def _ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    redact_workbooks(DATA_DIR)


@app.route("/healthz")
def healthz():
    return jsonify(ok=True)


# ── Pages ───────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


# ── Upload / Download ──────────────────────────────────────────────

@app.route("/api/upload", methods=["POST"])
def upload():
    _ensure_data_dir()
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify(error="未选择文件。"), 400
    filename = secure_filename(file.filename)
    if not filename or not filename.lower().endswith(".xlsm"):
        return jsonify(error="仅支持 .xlsm 文件。"), 400

    dest = DATA_DIR / filename
    file.save(str(dest))
    backup_workbook_identifiers(dest)
    redact_workbook_identifiers(dest)
    return jsonify(ok=True, filename=filename)


@app.route("/api/download")
def download():
    _ensure_data_dir()
    wb = find_workbook(DATA_DIR)
    if not wb:
        return jsonify(error="云端没有工作簿，请先上传。"), 404
    try:
        export_wb = prepare_download_workbook(wb)
    except ExcelError as exc:
        return jsonify(error=str(exc)), 400

    @after_this_request
    def cleanup_export(response):
        try:
            export_wb.unlink(missing_ok=True)
        except OSError:
            pass
        return response

    return send_file(str(export_wb), as_attachment=True, download_name=wb.name)


# ── Patients ────────────────────────────────────────────────────────

@app.route("/api/patients")
def patients():
    _ensure_data_dir()
    wb = find_workbook(DATA_DIR)
    if not wb:
        return jsonify(error="云端没有工作簿，请先上传。"), 404
    try:
        rows = list_patients(wb)
    except ExcelError as exc:
        return jsonify(error=str(exc)), 400
    return jsonify(patients=rows)


# ── Generate (LLM) ─────────────────────────────────────────────────

@app.route("/api/generate", methods=["POST"])
def generate():
    if not OPENAI_API_KEY:
        return jsonify(error="服务器未配置 OPENAI_API_KEY。"), 500

    data = request.get_json(silent=True) or {}
    raw_text = (data.get("raw_text") or "").strip()
    if not raw_text:
        return jsonify(error="请输入查房记录。"), 400

    patient_info = (data.get("patient_info") or "").strip()

    result = structure_note(OPENAI_API_KEY, OPENAI_MODEL, patient_info, raw_text, OPENAI_BASE_URL)
    return jsonify(result)


# ── Save to Excel ──────────────────────────────────────────────────

@app.route("/api/save", methods=["POST"])
def save():
    _ensure_data_dir()
    wb = find_workbook(DATA_DIR)
    if not wb:
        return jsonify(error="云端没有工作簿，请先上传。"), 404

    data = request.get_json(silent=True) or {}
    row_idx = data.get("row_idx")
    note_text = (data.get("note_text") or "").strip()
    level = data.get("level", "")
    note_type = data.get("note_type", "")
    date_str = data.get("date", "")
    overwrite_slot = data.get("overwrite_slot")

    if not row_idx or not note_text:
        return jsonify(error="缺少必填项（row_idx, note_text）。"), 400

    try:
        record_date = datetime.strptime(date_str, "%Y-%m-%d").date() if date_str else datetime.today().date()
    except ValueError:
        return jsonify(error="日期格式错误，应为 YYYY-MM-DD。"), 400

    try:
        result = save_note(wb, int(row_idx), record_date, level, note_type, note_text, overwrite_slot)
    except ExcelError as exc:
        return jsonify(error=str(exc)), 400

    return jsonify(ok=True, **result)


# ── Slot status ─────────────────────────────────────────────────────

@app.route("/api/slots/<int:row_idx>")
def slots(row_idx):
    _ensure_data_dir()
    wb = find_workbook(DATA_DIR)
    if not wb:
        return jsonify(error="云端没有工作簿，请先上传。"), 404
    try:
        status = get_slot_status(wb, row_idx)
    except ExcelError as exc:
        return jsonify(error=str(exc)), 400
    return jsonify(slots=status)


# ── Main ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _ensure_data_dir()
    app.run(host="0.0.0.0", port=PORT, debug=True)
