"""Flask application – PICU pharmacy monitoring web interface."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from flask import Flask, after_this_request, jsonify, render_template, request, send_file
from flask_login import current_user, login_required
from flask_wtf.csrf import CSRFProtect, generate_csrf
from werkzeug.utils import secure_filename

from config import (
    DATA_DIR,
    MAX_UPLOAD_MB,
    OPENAI_API_KEY,
    OPENAI_BASE_URL,
    OPENAI_MODEL,
    PORT,
    SECRET_KEY,
    SQLALCHEMY_DATABASE_URI,
    ADMIN_USERNAME,
    ADMIN_PASSWORD,
)
from excel_io import (
    ExcelError,
    backup_workbook_identifiers,
    find_workbook,
    get_slot_detail,
    get_slot_status,
    list_patients,
    prepare_download_workbook,
    redact_workbook_identifiers,
    redact_workbooks,
    save_note,
)
from llm import structure_note
from models import db, init_db, login_manager

app = Flask(__name__)
csrf = CSRFProtect()
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["SECRET_KEY"] = SECRET_KEY
app.config["SQLALCHEMY_DATABASE_URI"] = SQLALCHEMY_DATABASE_URI

db.init_app(app)
login_manager.init_app(app)
csrf.init_app(app)

from admin import admin_bp  # noqa: E402
from auth import auth_bp  # noqa: E402

app.register_blueprint(auth_bp)
app.register_blueprint(admin_bp)


@app.context_processor
def inject_csrf_token():
    """Always expose csrf_token() in templates."""
    return {"csrf_token": generate_csrf}


def _user_data_dir() -> Path:
    """Return the data directory scoped to the current logged-in user."""
    return DATA_DIR / "users" / str(current_user.id)


def _ensure_user_dir() -> Path:
    d = _user_data_dir()
    d.mkdir(parents=True, exist_ok=True)
    redact_workbooks(d)
    return d


@app.route("/healthz")
def healthz():
    return jsonify(ok=True)


# ── Pages ───────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    return render_template("index.html")


# ── Upload / Download ──────────────────────────────────────────────

@app.route("/api/upload", methods=["POST"])
@login_required
def upload():
    user_dir = _ensure_user_dir()
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify(error="未选择文件。"), 400
    filename = secure_filename(file.filename)
    if not filename or not filename.lower().endswith(".xlsm"):
        return jsonify(error="仅支持 .xlsm 文件。"), 400

    dest = user_dir / filename
    file.save(str(dest))
    backup_workbook_identifiers(dest)
    redact_workbook_identifiers(dest)
    return jsonify(ok=True, filename=filename)


@app.route("/api/download")
@login_required
def download():
    user_dir = _ensure_user_dir()
    wb = find_workbook(user_dir)
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
@login_required
def patients():
    user_dir = _ensure_user_dir()
    wb = find_workbook(user_dir)
    if not wb:
        return jsonify(error="云端没有工作簿，请先上传。"), 404
    try:
        rows = list_patients(wb)
    except ExcelError as exc:
        return jsonify(error=str(exc)), 400
    return jsonify(patients=rows)


# ── Generate (LLM) ─────────────────────────────────────────────────

@app.route("/api/generate", methods=["POST"])
@login_required
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
@login_required
def save():
    user_dir = _ensure_user_dir()
    wb = find_workbook(user_dir)
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
@login_required
def slots(row_idx):
    user_dir = _ensure_user_dir()
    wb = find_workbook(user_dir)
    if not wb:
        return jsonify(error="云端没有工作簿，请先上传。"), 404
    try:
        status = get_slot_status(wb, row_idx)
    except ExcelError as exc:
        return jsonify(error=str(exc)), 400
    return jsonify(slots=status)


@app.route("/api/slots/<int:row_idx>/<int:slot_index>")
@login_required
def slot_detail(row_idx, slot_index):
    user_dir = _ensure_user_dir()
    wb = find_workbook(user_dir)
    if not wb:
        return jsonify(error="云端没有工作簿，请先上传。"), 404
    try:
        detail = get_slot_detail(wb, row_idx, slot_index)
    except ExcelError as exc:
        return jsonify(error=str(exc)), 400
    return jsonify(slot=detail)


# ── User info API (for frontend) ───────────────────────────────────

@app.route("/api/me")
@login_required
def me():
    return jsonify(
        username=current_user.username,
        display_name=current_user.display_name,
        role=current_user.role,
    )


# ── Main ────────────────────────────────────────────────────────────

DATA_DIR.mkdir(parents=True, exist_ok=True)
init_db(app)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=True)
