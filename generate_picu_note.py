from __future__ import annotations

import argparse
import json
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


TEMPLATE_FILE = Path(__file__).with_name("picu_note_templates.json")


@dataclass
class FieldSpec:
    key: str
    label: str
    default: str


@dataclass
class TemplateSpec:
    template_id: str
    name: str
    description: str
    default_note_type: str
    fields: list[FieldSpec]
    template: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate PICU pharmacy monitoring note text for Excel."
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List available PICU note templates.",
    )
    parser.add_argument(
        "--template",
        help="Template ID to use. If omitted, the script starts in interactive mode.",
    )
    parser.add_argument(
        "--set",
        action="append",
        default=[],
        metavar="KEY=VALUE",
        help="Set a field value directly. Can be used multiple times.",
    )
    parser.add_argument(
        "--multiline",
        action="store_true",
        help="Output note in multiple lines instead of a single line.",
    )
    parser.add_argument(
        "--copy",
        action="store_true",
        help="Copy the generated note to the Windows clipboard.",
    )
    return parser.parse_args()


def fail(message: str) -> None:
    print(f"[ERROR] {message}", file=sys.stderr)
    raise SystemExit(1)


def load_templates() -> list[TemplateSpec]:
    if not TEMPLATE_FILE.exists():
        fail(f"Template file not found: {TEMPLATE_FILE}")

    with TEMPLATE_FILE.open("r", encoding="utf-8") as handle:
        raw = json.load(handle)

    templates: list[TemplateSpec] = []
    for item in raw.get("templates", []):
        fields = [
            FieldSpec(
                key=field["key"],
                label=field["label"],
                default=field.get("default", ""),
            )
            for field in item.get("fields", [])
        ]
        templates.append(
            TemplateSpec(
                template_id=item["id"],
                name=item["name"],
                description=item.get("description", ""),
                default_note_type=item.get("default_note_type", "药学监护"),
                fields=fields,
                template=item["template"],
            )
        )
    if not templates:
        fail("No templates were found in the template file.")
    return templates


def parse_set_args(items: list[str]) -> dict[str, str]:
    values: dict[str, str] = {}
    for item in items:
        if "=" not in item:
            fail(f"Invalid --set value: {item}. Expected KEY=VALUE.")
        key, value = item.split("=", 1)
        key = key.strip()
        if not key:
            fail(f"Invalid --set value: {item}. Field name cannot be empty.")
        values[key] = value.strip()
    return values


def print_template_list(templates: list[TemplateSpec]) -> None:
    print("可用 PICU 模板：")
    for index, template in enumerate(templates, start=1):
        print(f"{index}. {template.template_id} | {template.name}")
        print(f"   {template.description}")


def choose_template_interactively(templates: list[TemplateSpec]) -> TemplateSpec:
    print_template_list(templates)
    while True:
        choice = input("\n请输入模板编号或 template_id：").strip()
        if not choice:
            continue
        if choice.isdigit():
            index = int(choice) - 1
            if 0 <= index < len(templates):
                return templates[index]
        else:
            for template in templates:
                if template.template_id == choice:
                    return template
        print("未识别该模板，请重新输入。")


def select_template(templates: list[TemplateSpec], template_id: str | None) -> TemplateSpec:
    if template_id is None:
        return choose_template_interactively(templates)
    for template in templates:
        if template.template_id == template_id:
            return template
    fail(f"Unknown template_id: {template_id}")


def prompt_field(field: FieldSpec, preset: str | None) -> str:
    if preset is not None:
        return preset

    prompt = f"{field.label}"
    if field.default:
        prompt += f" [{field.default}]"
    prompt += "："
    value = input(prompt).strip()
    return value or field.default


def collect_values(template: TemplateSpec, preset_values: dict[str, str]) -> dict[str, str]:
    values: dict[str, str] = {}
    print(f"\n已选择模板：{template.name}")
    print(f"说明：{template.description}\n")
    for field in template.fields:
        values[field.key] = prompt_field(field, preset_values.get(field.key))
    return values


def render_note(template: TemplateSpec, values: dict[str, str], multiline: bool) -> str:
    note = template.template.format(**values).strip()
    if not multiline:
        return " ".join(note.split())

    replacements = [
        ("问题：", "问题："),
        ("分析：", "\n分析："),
        ("处理：", "\n处理："),
        ("结果/计划：", "\n结果/计划："),
    ]
    for source, target in replacements[1:]:
        note = note.replace(source, target)
    return note


def copy_to_clipboard(text: str) -> None:
    command = ["powershell", "-NoProfile", "-Command", "Set-Clipboard -Value @'\n" + text + "\n'@"]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        fail(f"Failed to copy note to clipboard: {result.stderr.strip()}")


def main() -> None:
    args = parse_args()
    templates = load_templates()

    if args.list:
        print_template_list(templates)
        return

    preset_values = parse_set_args(args.set)
    template = select_template(templates, args.template)
    values = collect_values(template, preset_values)
    note = render_note(template, values, args.multiline)

    print("\n生成的诊查记录：\n")
    print(note)

    if args.copy:
        copy_to_clipboard(note)
        print("\n[OK] 已复制到剪贴板。")


if __name__ == "__main__":
    main()
