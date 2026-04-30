# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Project Overview

PICU clinical pharmacy monitoring system for Sichuan University West China Second Hospital. A Flask web app + three CLI tools that import patient data from case list Excel files, generate standardized pharmacy monitoring notes (via LLM or templates), and write them into `.xlsm` workbooks with VBA macros preserved.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run web app (development)
python app.py                  # http://localhost:5000

# Run web app (production)
gunicorn app:app --bind 0.0.0.0:$PORT

# CLI: generate a note interactively
python generate_picu_note.py
python generate_picu_note.py --list                    # list templates
python generate_picu_note.py --template abx_review --copy  # generate + clipboard

# CLI: import case list into monitoring workbook
python import_case_list.py --pharmacist "姓名" --employee-id "工号"

# CLI: write note directly into Excel workbook
python write_picu_note_to_excel.py --list-patients
python write_picu_note_to_excel.py --inpatient-no "12345" --template tdm --set "药物=万古霉素"
```

No test suite, linter, or build system.

## Architecture

Two interfaces (web + CLI) share the same Excel I/O and template logic:

```
Case List Excel (病例清单.xlsx)
        │
        ▼
  import_case_list.py ──► Monitoring Workbook (.xlsm, cols A-K)
                                    │
                    ┌───────────────┴───────────────┐
                    ▼                               ▼
              Flask Web App                    CLI Scripts
            (app.py + index.html)      (generate_picu_note.py,
                    │                   write_picu_note_to_excel.py)
                    ▼                               │
               excel_io.py ◄────────────────────────┘
            (thread-safe wrapper)
                    │
                    ▼
              openpyxl (keep_vba=True)
```

### Web layer (`app.py` → `excel_io.py` → `llm.py`)

- **app.py** — Flask routes: upload/download `.xlsm`, list patients, LLM-generate notes, save notes to Excel. Single-page app served from `templates/index.html`.
- **excel_io.py** — Thread-safe (`threading.Lock`) wrapper around CLI functions. Converts `SystemExit` from CLI `fail()` calls into catchable `ExcelError`. Imports shared functions from `import_case_list.py` and `write_picu_note_to_excel.py`.
- **llm.py** — Calls OpenAI-compatible API (configured for OpenRouter) to structure free-form text into 问题→分析→处理→結果/計劃 format. Validates output markers, retries once with correction hint.
- **config.py** — Loads `.env` via python-dotenv. Key vars: `OPENAI_API_KEY`, `OPENAI_BASE_URL` (default: OpenRouter), `OPENAI_MODEL`, `DATA_DIR`, `PORT`.

### CLI layer (standalone scripts)

- **import_case_list.py** — Reads `*病例清单*.xlsx`, deduplicates by 住院号, writes patient demographics into columns A–K (rows 3–962).
- **generate_picu_note.py** — Template-based note generation from `picu_note_templates.json`. No Excel dependency. Windows clipboard via PowerShell.
- **write_picu_note_to_excel.py** — Combines template selection + Excel writing. Patient lookup by 住院号/姓名/床号. 6 note slots per row with auto-detection of previous level/type defaults.

## Excel Workbook Layout

Monitoring workbook (`*.xlsm`): rows 3–962 are data rows, rows 1–2 are headers.

| Columns | Content |
|---------|---------|
| A–K | Patient base info (科室, 药师, 工号, 住院号, 床号, 姓名, 年龄, 性别, 体重, 入院日期, 入院诊断) |
| L–O | Note slot 1 (日期, 分级, 类型, 记录) |
| P–S | Note slot 2 |
| T–W | Note slot 3 |
| X–AA | Note slot 4 |
| AB–AE | Note slot 5 |
| AF–AI | Note slot 6 |

**Important:** Always open `.xlsm` files with `keep_vba=True` in openpyxl.

## Key Conventions

- File auto-discovery: scripts find the newest matching file by glob, skipping `~$*` temp files
- Care levels: 一级监护, 二级监护, 三级监护
- Note types: 药学查房, 药物重整, 药学监护, 用药咨询, 用药教育
- CLI scripts use `fail()` → `SystemExit`; web layer catches these as `ExcelError`
- LLM notes must contain 4 markers: `问题：`, `分析：`, `处理：`, `结果/计划：`
- 8 templates in `picu_note_templates.json`: `abx_review`, `renal_crrt`, `tdm`, `adr`, `interaction`, `sedation_analgesia`, `nutrition`, `infusion_compatibility`

## Deployment

Configured for Render (`render.yaml`). Environment variables: `OPENAI_API_KEY`, `OPENAI_BASE_URL`, `OPENAI_MODEL`, `DATA_DIR`. Workbook storage on persistent disk at `/data`.
