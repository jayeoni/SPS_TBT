# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the App

```bat
# Launch (from sps_tool/)
start.bat          # kills port 5000, starts Flask at http://localhost:5000

# Repair broken Python environment
setup.bat          # re-downloads python313.zip and reinstalls packages

# Direct Python (for quick tests)
python\python-3.13.13-embed-amd64\python.exe app.py
```

All code runs inside `sps_tool/` using the **bundled Python 3.13.13 embeddable** at `sps_tool/python/python-3.13.13-embed-amd64/`. Never assume a system Python.

There are no automated tests or lint scripts. Verification is done manually by running the app and uploading sample `.docx` files (`GSPSNCHL881.docx`, `GSPSNCRI349A1.docx`) from the repo root.

---

## Architecture

Single-user Flask web app (MAFRA government analyst). The user uploads WTO SPS notification `.docx` files; the tool fills a monthly Excel workbook and creates bilingual `*_번역.docx` files.

### Processing Pipeline (`app.py: process_single_file`)

```
.docx upload
  → parser.py          # extract structured fields from WTO form tables
  → dept_lookup.py     # translate regions → Korean; set is_korea_targeted / is_all_partners flags
  → llm.py             # single LLM call: translate + classify (returns JSON with ~18 fields)
  → [post-LLM]         # compute 국내수출품목 based on region flags + export_lookup
  → date_engine.py     # resolve "60 days from 배포일" style formulas
  → excel_writer.py    # find pre-existing row by 문서번호; write ~13 fields
  → word_writer.py     # copy source .docx; append Korean paragraphs to each form cell
```

### Module Responsibilities

| Module | Role |
|---|---|
| `parser.py` | Extracts raw fields from WTO SPS form tables. Handles Layout A (label embedded in content cell) and Layout B (label | content cells). Special handling for addendum docs (all content in single-column table). |
| `llm.py` | Builds a single large prompt and calls Ollama (default: `qwen2.5:7b`) or Anthropic Claude. Returns JSON with fields: 제목, 내용, 해당품목, 목적, 해당국가, 통보국_kr, 담당기관_kr, 주간보고, 구분, 중요도, 관련부서, 통보내용, 통보_세부, 품목, flags, source_language. |
| `dept_lookup.py` | Two responsibilities: (1) `translate_regions()` converts English region names to Korean; (2) `lookup_dept()` maps `(구분, 통보내용, 통보_세부)` → department list using `DEPT_TABLE` that mirrors the `★매뉴얼` sheet. |
| `export_lookup.py` | Loads `2025_연간 전체실적_전체.xlsx` (203K rows, sheet `연간실적`). `lookup(country, product, is_all_partners)` infers HS chapters from product keywords and queries Korean exports. |
| `excel_writer.py` | Finds the pre-populated row by `문서번호` in the monthly sheet. `_detect_col_map()` reads header row 1 to get actual column positions (falls back to hardcoded `COL`). Skips non-empty cells. |
| `word_writer.py` | `ROW_PATTERNS` → `ROW_BUILDERS` dispatch table for all 14 form row types (+ 5 addendum-specific). `_translate_doc_titles()` scans both paragraphs and table cells (addendum docs have no top-level paragraphs). |
| `date_engine.py` | Resolves raw date strings ("60 days", specific dates) relative to `배포일` from Excel. |

### Key Data Flows

**해당국가**: Read from `.docx` 4. regions field → `dept_lookup.translate_regions()` → written directly. Never from LLM.

**국내수출품목**:
- `is_korea_targeted` (regions contains 한국, not all-partners) → use `해당품목` from LLM directly
- `is_all_partners` → query `export_lookup` by notifying country + product
- Third-country restriction → `'-'`

**관련부서**: `dept_lookup.lookup_dept(구분, 통보내용, 통보_세부)` from `DEPT_TABLE`; falls back to LLM output only if no table match, and flags the cell yellow.

**Addendum documents** (`/Add.N` in doc number or "Addendum" in header): Single-column table layout. `word_writer._translate_addendum_reg_title()` uses positional detection (cell after `___` separator) since the regulation title cell has no label.

### Config & Persistence

- `config.json` — excel path, export data path, target month (e.g. `26.4월`), LLM backend/model. API key never saved to JSON; goes to `.env`.
- `terminology.json` — ~100 domain terms passed to LLM for consistent translation (first 50 entries only).
- Column positions detected dynamically from Excel header row; hardcoded `COL` dict in `excel_writer.py` is the fallback.

### LLM Backends

Configured via `Settings → LLM Backend`:
- **ollama** (default): Local Ollama server at `localhost:11434`, model `qwen2.5:7b`. No API key needed. `num_predict: 4096`.
- **anthropic**: Cloud API, model `claude-sonnet-4-6`, requires `ANTHROPIC_API_KEY` in `.env`.

The `올라마 상태` indicator on the index page checks `/ollama-status` before processing.
