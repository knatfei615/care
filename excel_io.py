"""Excel I/O wrapper for the web layer.

Re-uses functions from the existing CLI scripts while converting
``SystemExit`` (raised by ``fail()``) into a catchable ``ExcelError``.
"""

from __future__ import annotations

import json
import shutil
import threading
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any
from uuid import uuid4

import openpyxl

from import_case_list import RECORD_KEYWORD, TEMP_PREFIX
from write_picu_note_to_excel import (
    NOTE_SLOTS,
    LEVEL_OPTIONS,
    MAX_ROW,
    START_ROW,
    TYPE_OPTIONS,
    PatientRow,
    format_cell,
    load_patient_rows,
    find_previous_defaults,
    resolve_slot,
    write_note_to_sheet,
)

_lock = threading.Lock()
ID_COLUMN = "D"
NAME_COLUMN = "F"
BASE_COLUMNS = ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K")
ANON_ID_PREFIX = "ANON-"
TEMP_EXPORT_PREFIX = "~export-"
ANON_NAME_PREFIX = "匿名患者"
STATUS_TODAY_UPDATED = "today_updated"
STATUS_PENDING = "pending"
STATUS_STALE_48H = "stale_48h"
STATUS_NO_RECORD = "no_record"

DIAGNOSIS_TAG_RULES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("抗菌药物", ("感染", "抗感染", "抗菌", "肺炎", "败血症", "sepsis", "abx", "antibiotic")),
    ("TDM", ("tdm", "血药浓度", "万古霉素", "他克莫司", "丙戊酸", "地高辛")),
    ("CRRT", ("crrt", "透析", "血液净化", "连续肾脏替代", "ecmo")),
    ("肾功能异常", ("肾功能", "肾损伤", "少尿", "无尿", "肌酐", "bun", "aki", "ckd")),
    ("ADR", ("不良反应", "过敏", "皮疹", "肝损伤", "adr", "药物不良", "药疹")),
)

class ExcelError(Exception):
    """Raised when an Excel operation fails."""


def _catch(fn, *args, **kwargs):
    """Call *fn* and convert ``SystemExit`` into ``ExcelError``."""
    try:
        return fn(*args, **kwargs)
    except SystemExit as exc:
        raise ExcelError(str(exc)) from exc


def _identifier_backup_path(wb_path: Path) -> Path:
    return wb_path.with_name(f"{wb_path.stem}.identifiers.json")


def _parse_excel_date(value: Any) -> datetime | None:
    """Parse excel/date-like values into datetime."""
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())

    text = format_cell(value).strip()
    if not text:
        return None

    text = text.replace("年", "-").replace("月", "-").replace("日", "")
    text = text.replace("/", "-").strip()
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return None


def _extract_tags_from_diagnosis(diagnosis: str) -> list[str]:
    normalized = (diagnosis or "").lower()
    tags: list[str] = []
    for tag, keywords in DIAGNOSIS_TAG_RULES:
        if any(keyword.lower() in normalized for keyword in keywords):
            tags.append(tag)
    return tags


def _extract_major_issue(note_text: str) -> str:
    if not note_text:
        return ""

    match = re.search(r"主观资料：\s*(.*?)\s*(?:客观资料：|$)", note_text, flags=re.S)
    if not match:
        match = re.search(r"问题：\s*(.*?)\s*(?:分析：|$)", note_text, flags=re.S)
    if not match:
        return ""

    issue = re.sub(r"\s+", " ", match.group(1)).strip("；;。 ")
    return issue[:100]


def _collect_slot_records(sheet, row_idx: int) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Collect raw slot status list and effective note records."""
    slots: list[dict[str, Any]] = []
    records: list[dict[str, Any]] = []

    for slot in NOTE_SLOTS:
        note_val = format_cell(sheet[f"{slot['note']}{row_idx}"].value)
        date_raw = sheet[f"{slot['date']}{row_idx}"].value
        date_val = format_cell(date_raw)
        level_val = format_cell(sheet[f"{slot['level']}{row_idx}"].value)
        type_val = format_cell(sheet[f"{slot['type']}{row_idx}"].value)
        parsed_date = _parse_excel_date(date_raw or date_val)

        slots.append({
            "index": slot["index"],
            "has_note": bool(note_val),
            "date": date_val,
            "preview": note_val[:60] if note_val else "",
            "level": level_val,
            "note_type": type_val,
        })

        if note_val:
            records.append({
                "slot_index": slot["index"],
                "date": date_val,
                "parsed_date": parsed_date,
                "level": level_val,
                "note_type": type_val,
                "note": note_val,
            })

    records.sort(
        key=lambda item: (
            item["parsed_date"] or datetime.min,
            item["slot_index"],
        ),
        reverse=True,
    )
    return slots, records


def _compute_status(last_note_dt: datetime | None) -> str:
    if last_note_dt is None:
        return STATUS_NO_RECORD

    now = datetime.now()
    if last_note_dt.date() == now.date():
        return STATUS_TODAY_UPDATED

    elapsed = now - last_note_dt
    if elapsed.total_seconds() > 48 * 3600:
        return STATUS_STALE_48H
    return STATUS_PENDING


def _build_tracking_summary(records: list[dict[str, Any]]) -> dict[str, Any]:
    latest = records[0] if records else None
    latest_dt = latest["parsed_date"] if latest else None
    latest_date = ""
    if latest:
        latest_date = latest["date"] or (latest_dt.strftime("%Y-%m-%d") if latest_dt else "")

    recent_records = []
    for item in records[:3]:
        preview = re.sub(r"\s+", " ", item["note"]).strip()[:50]
        recent_records.append({
            "slot_index": item["slot_index"],
            "date": item["date"] or (item["parsed_date"].strftime("%Y-%m-%d") if item["parsed_date"] else ""),
            "level": item["level"],
            "note_type": item["note_type"],
            "preview": preview,
        })

    return {
        "latest_note_date": latest_date,
        "major_issue": _extract_major_issue(latest["note"]) if latest else "",
        "recent_records": recent_records,
    }


def backup_workbook_identifiers(wb_path: Path) -> Path:
    """Persist original patient identifiers so downloads can restore them."""
    with _lock:
        book = openpyxl.load_workbook(wb_path, keep_vba=True, data_only=False)
        sheet = book[book.sheetnames[0]]
        rows: dict[str, dict[str, str]] = {}

        for row_idx in range(START_ROW, MAX_ROW + 1):
            raw_id = format_cell(sheet[f"{ID_COLUMN}{row_idx}"].value)
            raw_name = format_cell(sheet[f"{NAME_COLUMN}{row_idx}"].value)
            if not raw_id and not raw_name:
                continue
            rows[str(row_idx)] = {
                "inpatient_no": raw_id,
                "name": raw_name,
            }

        book.close()

    backup_path = _identifier_backup_path(wb_path)
    backup_path.write_text(
        json.dumps({"rows": rows}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return backup_path


def upsert_backup_identifier_row(wb_path: Path, row_idx: int, inpatient_no: str, name: str) -> Path:
    """Update one row in identifier backup without rewriting existing rows."""
    backup_path = _identifier_backup_path(wb_path)
    rows: dict[str, dict[str, str]] = {}
    if backup_path.exists():
        try:
            backup_data = json.loads(backup_path.read_text(encoding="utf-8"))
            rows = backup_data.get("rows", {})
        except (json.JSONDecodeError, OSError):
            rows = {}

    rows[str(row_idx)] = {
        "inpatient_no": inpatient_no,
        "name": name,
    }
    backup_path.write_text(
        json.dumps({"rows": rows}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return backup_path


def prepare_download_workbook(wb_path: Path) -> Path:
    """Create a temporary workbook copy with original identifiers restored."""
    backup_path = _identifier_backup_path(wb_path)
    if not backup_path.exists():
        raise ExcelError("未找到原始姓名/住院号备份，当前文件无法恢复下载。请重新上传原始工作簿。")

    backup_data = json.loads(backup_path.read_text(encoding="utf-8"))
    rows = backup_data.get("rows", {})
    export_path = wb_path.with_name(f"{TEMP_EXPORT_PREFIX}{uuid4().hex}-{wb_path.name}")
    shutil.copy2(wb_path, export_path)

    with _lock:
        book = openpyxl.load_workbook(export_path, keep_vba=True)
        sheet = book[book.sheetnames[0]]

        for row_idx_text, values in rows.items():
            row_idx = int(row_idx_text)
            sheet[f"{ID_COLUMN}{row_idx}"] = values.get("inpatient_no", "")
            sheet[f"{NAME_COLUMN}{row_idx}"] = values.get("name", "")

        book.save(export_path)
        book.close()
    return export_path


# ── workbook discovery ──────────────────────────────────────────────

def find_workbook(data_dir: Path) -> Path | None:
    """Return the first ``.xlsm`` workbook found in *data_dir*, or ``None``."""
    if not data_dir.is_dir():
        return None
    candidates = [
        p for p in data_dir.iterdir()
        if p.is_file()
        and p.suffix.lower() == ".xlsm"
        and not p.name.startswith(TEMP_PREFIX)
        and not p.name.startswith(TEMP_EXPORT_PREFIX)
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def redact_workbook_identifiers(wb_path: Path) -> bool:
    """Replace patient name / inpatient number with row-based aliases."""
    changed = False
    with _lock:
        book = openpyxl.load_workbook(wb_path, keep_vba=True)
        sheet = book[book.sheetnames[0]]

        for row_idx in range(START_ROW, MAX_ROW + 1):
            raw_id = format_cell(sheet[f"{ID_COLUMN}{row_idx}"].value)
            raw_name = format_cell(sheet[f"{NAME_COLUMN}{row_idx}"].value)
            if not raw_id and not raw_name:
                continue

            alias_num = row_idx - START_ROW + 1
            masked_id = f"{ANON_ID_PREFIX}{alias_num:04d}"
            masked_name = f"{ANON_NAME_PREFIX}{alias_num:03d}"

            if raw_id != masked_id:
                sheet[f"{ID_COLUMN}{row_idx}"] = masked_id
                changed = True
            if raw_name != masked_name:
                sheet[f"{NAME_COLUMN}{row_idx}"] = masked_name
                changed = True

        if changed:
            book.save(wb_path)
        book.close()

    return changed


def redact_workbooks(data_dir: Path) -> int:
    """Redact every workbook in the data directory. Returns changed file count."""
    if not data_dir.is_dir():
        return 0

    changed = 0
    for wb_path in data_dir.iterdir():
        if (
            wb_path.is_file()
            and wb_path.suffix.lower() == ".xlsm"
            and not wb_path.name.startswith(TEMP_PREFIX)
            and not wb_path.name.startswith(TEMP_EXPORT_PREFIX)
        ):
            if redact_workbook_identifiers(wb_path):
                changed += 1
    return changed


# ── patient list ────────────────────────────────────────────────────

def list_patients(wb_path: Path) -> list[dict]:
    """Return a JSON-friendly list of patients with previous defaults."""
    with _lock:
        book = openpyxl.load_workbook(wb_path, keep_vba=True, data_only=True)
        sheet = book[book.sheetnames[0]]
        patients = load_patient_rows(sheet)
        result = []
        for p in patients:
            prev_level, prev_type = find_previous_defaults(sheet, p)
            _, records = _collect_slot_records(sheet, p.row_idx)
            latest = records[0] if records else None
            latest_dt = latest["parsed_date"] if latest else None
            last_note_date = ""
            if latest:
                last_note_date = latest["date"] or (latest_dt.strftime("%Y-%m-%d") if latest_dt else "")
            status = _compute_status(latest_dt)
            result.append({
                "row_idx": p.row_idx,
                "inpatient_no": p.inpatient_no,
                "bed_no": p.bed_no,
                "name": p.name,
                "age": p.age,
                "sex": p.sex,
                "weight": p.weight,
                "admission_date": p.admission_date,
                "diagnosis": p.diagnosis,
                "prev_level": prev_level,
                "prev_type": prev_type,
                "last_note_date": last_note_date,
                "status": status,
                "tags": _extract_tags_from_diagnosis(p.diagnosis),
            })
        book.close()
    return result


def _clean_patient_payload(payload: dict[str, Any]) -> dict[str, str]:
    """Normalize manual patient payload from request JSON."""
    return {
        "A": format_cell(payload.get("department")),
        "B": format_cell(payload.get("pharmacist")),
        "C": format_cell(payload.get("employee_id")),
        "D": format_cell(payload.get("inpatient_no")),
        "E": format_cell(payload.get("bed_no")),
        "F": format_cell(payload.get("name")),
        "G": format_cell(payload.get("age")),
        "H": format_cell(payload.get("sex")),
        "I": format_cell(payload.get("weight")),
        "J": format_cell(payload.get("admission_date")),
        "K": format_cell(payload.get("diagnosis")),
    }


def add_patient(wb_path: Path, payload: dict[str, Any]) -> dict[str, Any]:
    """Append one patient base-info row into A-K columns."""
    row_data = _clean_patient_payload(payload)
    if not row_data["D"] and not row_data["F"]:
        raise ExcelError("住院号和姓名不能同时为空。")

    target_row: int | None = None
    with _lock:
        book = openpyxl.load_workbook(wb_path, keep_vba=True)
        sheet = book[book.sheetnames[0]]

        for row_idx in range(START_ROW, MAX_ROW + 1):
            inpatient_no = format_cell(sheet[f"D{row_idx}"].value)
            name = format_cell(sheet[f"F{row_idx}"].value)
            if not inpatient_no and not name:
                target_row = row_idx
                break

        if target_row is None:
            book.close()
            raise ExcelError(f"患者区域已满（{START_ROW}-{MAX_ROW} 行）。")

        for column in BASE_COLUMNS:
            sheet[f"{column}{target_row}"] = row_data[column]

        book.save(wb_path)
        book.close()

    upsert_backup_identifier_row(
        wb_path=wb_path,
        row_idx=target_row,
        inpatient_no=row_data["D"],
        name=row_data["F"],
    )
    redact_workbook_identifiers(wb_path)

    return {
        "row_idx": target_row,
        "department": row_data["A"],
        "pharmacist": row_data["B"],
        "employee_id": row_data["C"],
        "inpatient_no": row_data["D"],
        "bed_no": row_data["E"],
        "name": row_data["F"],
        "age": row_data["G"],
        "sex": row_data["H"],
        "weight": row_data["I"],
        "admission_date": row_data["J"],
        "diagnosis": row_data["K"],
    }


# ── slot status ─────────────────────────────────────────────────────

def _load_patient_row(sheet, row_idx: int) -> PatientRow:
    """Load and validate patient row data from worksheet."""
    if row_idx < START_ROW or row_idx > MAX_ROW:
        raise ExcelError(f"行号超出范围（{START_ROW}-{MAX_ROW}）。")

    patient = PatientRow(
        row_idx=row_idx,
        inpatient_no=format_cell(sheet[f"D{row_idx}"].value),
        bed_no=format_cell(sheet[f"E{row_idx}"].value),
        name=format_cell(sheet[f"F{row_idx}"].value),
        age=format_cell(sheet[f"G{row_idx}"].value),
        sex=format_cell(sheet[f"H{row_idx}"].value),
        weight=format_cell(sheet[f"I{row_idx}"].value),
        admission_date=format_cell(sheet[f"J{row_idx}"].value),
        diagnosis=format_cell(sheet[f"K{row_idx}"].value),
    )

    if not patient.inpatient_no and not patient.name:
        raise ExcelError(f"行 {row_idx} 没有患者数据。")

    return patient


def get_slot_status(wb_path: Path, row_idx: int) -> list[dict]:
    """Return the occupancy status of all 6 note slots for a given row."""
    with _lock:
        book = openpyxl.load_workbook(wb_path, keep_vba=True, data_only=True)
        sheet = book[book.sheetnames[0]]
        _load_patient_row(sheet, row_idx)
        result, _ = _collect_slot_records(sheet, row_idx)
        book.close()
    return result


def get_slot_status_with_summary(wb_path: Path, row_idx: int) -> dict[str, Any]:
    """Return slot status with patient tracking summary for a row."""
    with _lock:
        book = openpyxl.load_workbook(wb_path, keep_vba=True, data_only=True)
        sheet = book[book.sheetnames[0]]
        _load_patient_row(sheet, row_idx)
        slots, records = _collect_slot_records(sheet, row_idx)
        summary = _build_tracking_summary(records)
        summary["status"] = _compute_status(records[0]["parsed_date"] if records else None)
        book.close()
    return {"slots": slots, "summary": summary}


def get_slot_detail(wb_path: Path, row_idx: int, slot_index: int) -> dict:
    """Return the full content of one note slot for a given patient row."""
    slot = next((s for s in NOTE_SLOTS if s["index"] == slot_index), None)
    if not slot:
        raise ExcelError("记录槽位必须为 1-6。")

    with _lock:
        book = openpyxl.load_workbook(wb_path, keep_vba=True, data_only=True)
        sheet = book[book.sheetnames[0]]
        _load_patient_row(sheet, row_idx)

        note_val = format_cell(sheet[f"{slot['note']}{row_idx}"].value)
        date_val = format_cell(sheet[f"{slot['date']}{row_idx}"].value)
        level_val = format_cell(sheet[f"{slot['level']}{row_idx}"].value)
        type_val = format_cell(sheet[f"{slot['type']}{row_idx}"].value)
        book.close()

    return {
        "index": slot["index"],
        "has_note": bool(note_val),
        "date": date_val,
        "level": level_val,
        "note_type": type_val,
        "note": note_val,
    }


# ── save note ───────────────────────────────────────────────────────

def save_note(
    wb_path: Path,
    row_idx: int,
    record_date: date,
    level: str,
    note_type: str,
    note_text: str,
    overwrite_slot: int | None = None,
) -> dict:
    """Write a note into the workbook and return a summary dict."""
    with _lock:
        book = openpyxl.load_workbook(wb_path, keep_vba=True)
        sheet = book[book.sheetnames[0]]
        patient = _load_patient_row(sheet, row_idx)

        slot = _catch(resolve_slot, sheet, patient, overwrite_slot)

        write_note_to_sheet(
            sheet=sheet,
            patient=patient,
            slot=slot,
            record_date=record_date,
            level=level,
            note_type=note_type,
            note=note_text,
        )
        book.save(wb_path)
        book.close()
    return {
        "row_idx": row_idx,
        "name": patient.name,
        "slot_index": slot["index"],
        "date": record_date.isoformat(),
        "level": level,
        "note_type": note_type,
    }
