from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl

from generate_picu_note import (
    collect_values,
    load_templates,
    parse_set_args,
    render_note,
    select_template,
)
from import_case_list import RECORD_KEYWORD, TEMP_PREFIX, fail


START_ROW = 3
MAX_ROW = 962
NOTE_SLOTS = [
    {"index": 1, "date": "L", "level": "M", "type": "N", "note": "O"},
    {"index": 2, "date": "P", "level": "Q", "type": "R", "note": "S"},
    {"index": 3, "date": "T", "level": "U", "type": "V", "note": "W"},
    {"index": 4, "date": "X", "level": "Y", "type": "Z", "note": "AA"},
    {"index": 5, "date": "AB", "level": "AC", "type": "AD", "note": "AE"},
    {"index": 6, "date": "AF", "level": "AG", "type": "AH", "note": "AI"},
]
LEVEL_OPTIONS = ["一级监护", "二级监护", "三级监护"]
TYPE_OPTIONS = ["药学查房", "药物重整", "药学监护", "用药咨询", "用药教育"]


@dataclass
class PatientRow:
    row_idx: int
    inpatient_no: str
    bed_no: str
    name: str
    diagnosis: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a PICU monitoring note and write it directly into the Excel workbook."
    )
    parser.add_argument(
        "--record-book",
        type=Path,
        help="Path to the monitoring workbook. Defaults to the latest *住院药学诊查记录*.xlsm file.",
    )
    parser.add_argument(
        "--list-patients",
        action="store_true",
        help="List all patients currently in the workbook and exit.",
    )
    parser.add_argument("--row", type=int, help="Excel row number for the patient.")
    parser.add_argument("--inpatient-no", help="住院号/登记号，用于定位患者。")
    parser.add_argument("--name", help="姓名，用于定位患者。")
    parser.add_argument("--bed", help="床号，用于定位患者。")
    parser.add_argument("--template", help="PICU note template ID.")
    parser.add_argument(
        "--set",
        action="append",
        default=[],
        metavar="KEY=VALUE",
        help="Set a template field directly. Can be used multiple times.",
    )
    parser.add_argument(
        "--date",
        help="Record date in YYYY-MM-DD format. Defaults to today.",
    )
    parser.add_argument("--level", choices=LEVEL_OPTIONS, help="监护级别。")
    parser.add_argument("--note-type", choices=TYPE_OPTIONS, help="监护类型。")
    parser.add_argument(
        "--overwrite-slot",
        type=int,
        choices=[1, 2, 3, 4, 5, 6],
        help="Write into a specific note slot instead of the next empty one.",
    )
    parser.add_argument(
        "--multiline",
        action="store_true",
        help="Write note as multiple lines inside the cell.",
    )
    return parser.parse_args()


def newest_record_book(directory: Path) -> Path:
    candidates = [
        path
        for path in directory.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".xlsm"
        and not path.name.startswith(TEMP_PREFIX)
        and RECORD_KEYWORD in path.stem
    ]
    if not candidates:
        fail(f"Could not find a monitoring workbook in {directory}")

    preferred = [path for path in candidates if "已导入" in path.stem]
    target_pool = preferred or candidates
    return max(target_pool, key=lambda path: path.stat().st_mtime)


def format_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def load_patient_rows(sheet: openpyxl.worksheet.worksheet.Worksheet) -> list[PatientRow]:
    rows: list[PatientRow] = []
    for row_idx in range(START_ROW, MAX_ROW + 1):
        inpatient_no = format_cell(sheet[f"D{row_idx}"].value)
        name = format_cell(sheet[f"F{row_idx}"].value)
        if not inpatient_no and not name:
            continue
        rows.append(
            PatientRow(
                row_idx=row_idx,
                inpatient_no=inpatient_no,
                bed_no=format_cell(sheet[f"E{row_idx}"].value),
                name=name,
                diagnosis=format_cell(sheet[f"K{row_idx}"].value),
            )
        )
    return rows


def print_patients(patients: list[PatientRow]) -> None:
    print("当前工作簿患者列表：")
    for patient in patients:
        print(
            f"行 {patient.row_idx}: 住院号={patient.inpatient_no} "
            f"床号={patient.bed_no or '-'} 姓名={patient.name} "
            f"诊断={patient.diagnosis[:40]}"
        )


def choose_from_matches(matches: list[PatientRow]) -> PatientRow:
    print_patients(matches)
    while True:
        raw = input("\n请输入要写入的行号：").strip()
        if raw.isdigit():
            row_idx = int(raw)
            for patient in matches:
                if patient.row_idx == row_idx:
                    return patient
        print("未识别该行号，请重新输入。")


def resolve_patient(args: argparse.Namespace, patients: list[PatientRow]) -> PatientRow:
    if args.row is not None:
        for patient in patients:
            if patient.row_idx == args.row:
                return patient
        fail(f"Could not find patient on row {args.row}")

    matches = patients
    if args.inpatient_no:
        keyword = args.inpatient_no.strip()
        matches = [patient for patient in matches if patient.inpatient_no == keyword]
    if args.name:
        keyword = args.name.strip()
        matches = [patient for patient in matches if keyword in patient.name]
    if args.bed:
        keyword = args.bed.strip()
        matches = [patient for patient in matches if keyword in patient.bed_no]

    if not matches:
        fail("No matching patient was found.")
    if len(matches) == 1:
        return matches[0]

    print("找到多个匹配患者，请进一步选择。")
    return choose_from_matches(matches)


def find_previous_defaults(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    patient: PatientRow,
) -> tuple[str, str]:
    previous_level = ""
    previous_type = ""
    for slot in NOTE_SLOTS:
        note_value = format_cell(sheet[f"{slot['note']}{patient.row_idx}"].value)
        if note_value:
            level = format_cell(sheet[f"{slot['level']}{patient.row_idx}"].value)
            note_type = format_cell(sheet[f"{slot['type']}{patient.row_idx}"].value)
            if level:
                previous_level = level
            if note_type:
                previous_type = note_type
    return previous_level, previous_type


def resolve_slot(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    patient: PatientRow,
    overwrite_slot: int | None,
) -> dict[str, str]:
    if overwrite_slot is not None:
        return NOTE_SLOTS[overwrite_slot - 1]

    for slot in NOTE_SLOTS:
        note_value = format_cell(sheet[f"{slot['note']}{patient.row_idx}"].value)
        if not note_value:
            return slot
    fail(f"Patient row {patient.row_idx} already has 6 note entries. Use --overwrite-slot to replace one.")


def prompt_choice(label: str, options: list[str], default: str) -> str:
    option_text = "/".join(options)
    while True:
        raw = input(f"{label} [{default or option_text}]：").strip()
        if not raw and default:
            return default
        if raw in options:
            return raw
        if not raw and not default:
            print(f"请输入以下选项之一：{option_text}")
            continue
        print(f"无效输入，请输入以下选项之一：{option_text}")


def parse_record_date(raw_date: str | None) -> date:
    if not raw_date:
        return date.today()
    try:
        return datetime.strptime(raw_date, "%Y-%m-%d").date()
    except ValueError as exc:
        fail(f"Invalid --date value: {raw_date}. Expected YYYY-MM-DD. ({exc})")


def write_note_to_sheet(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    patient: PatientRow,
    slot: dict[str, str],
    record_date: date,
    level: str,
    note_type: str,
    note: str,
) -> None:
    row_idx = patient.row_idx
    sheet[f"{slot['date']}{row_idx}"] = record_date
    sheet[f"{slot['level']}{row_idx}"] = level
    sheet[f"{slot['type']}{row_idx}"] = note_type
    sheet[f"{slot['note']}{row_idx}"] = note


def main() -> None:
    args = parse_args()
    workbook_path = (args.record_book or newest_record_book(Path.cwd())).resolve()

    book = openpyxl.load_workbook(workbook_path, keep_vba=True)
    sheet = book[book.sheetnames[0]]
    patients = load_patient_rows(sheet)
    if not patients:
        fail("No patient rows were found in the workbook. Please import the case list first.")

    if args.list_patients:
        print_patients(patients)
        return

    patient = resolve_patient(args, patients)
    slot = resolve_slot(sheet, patient, args.overwrite_slot)

    templates = load_templates()
    template = select_template(templates, args.template)
    values = collect_values(template, parse_set_args(args.set))
    note = render_note(template, values, args.multiline)

    previous_level, previous_type = find_previous_defaults(sheet, patient)
    default_level = args.level or previous_level or "二级监护"
    default_type = args.note_type or previous_type or template.default_note_type or "药学监护"
    level = args.level or prompt_choice("监护级别", LEVEL_OPTIONS, default_level)
    note_type = args.note_type or prompt_choice("监护类型", TYPE_OPTIONS, default_type)
    record_date = parse_record_date(args.date)

    write_note_to_sheet(
        sheet=sheet,
        patient=patient,
        slot=slot,
        record_date=record_date,
        level=level,
        note_type=note_type,
        note=note,
    )
    book.save(workbook_path)

    print("\n[OK] 已写入 Excel。")
    print(f"工作簿: {workbook_path}")
    print(f"患者: 行 {patient.row_idx} | 住院号={patient.inpatient_no} | 姓名={patient.name}")
    print(f"写入位置: 诊查记录{slot['index']} ({slot['date']}/{slot['level']}/{slot['type']}/{slot['note']})")
    print(f"日期: {record_date.isoformat()} | 级别: {level} | 类型: {note_type}")


if __name__ == "__main__":
    main()
