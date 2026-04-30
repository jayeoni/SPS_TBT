# SPS Notification Processing Tool — Design Document

*Updated 2026-04-30 to reflect detailed implementation interview*

---

## Overview

This document defines the design and implementation specifications for a local automation tool that processes WTO SPS (Sanitary and Phytosanitary) notification Word files. The tool is used by a single analyst at a Korean government agency (MAFRA) to reduce the manual workload of translating, classifying, and recording SPS notifications in the official monthly Excel workbook.

The tool is a local Flask web application running at `http://localhost:5000`. It is launched by double-clicking `start.bat`, processes `.docx` files via drag-and-drop in the browser, writes results directly into the pre-populated Excel workbook, and generates bilingual Word output files alongside the source documents.

---

## Scope and Context

**User:** Single analyst. The same person processes and self-reviews all outputs — no separate supervisor approval step.

**Volume:** 5–15 notifications per week, approximately 15–20 minutes each manually. The tool targets near-full automation of field extraction, translation, classification, and Excel writing, reducing the analyst's role to reviewing flagged cells and confirming uncertain outputs.

**Source materials:**
- WTO SPS notification Word files (`.docx`), received via ePing and placed in monthly subfolders
- Source languages: English, Spanish, Portuguese (Spanish and Portuguese notifications are translated directly for short documents; complex/long Portuguese documents get a lime-green flag pending English version confirmation)
- Addendum documents (A1, A2 suffix) are processed as **new rows**, never as updates to the original row

**Output:**
- Updated cells in the monthly Excel sheet (pre-populated rows matched by 문서번호)
- Bilingual Word files (`*_번역.docx`) with Korean translation inserted as a second paragraph in each content cell

---

## File Layout

All files live in `C:\Users\mafra\SPS통보문\`:

| File | Role |
|---|---|
| `260428 SPS 통보문 내역(4월 4주).xlsx` | Master Excel workbook (one sheet per month) |
| `2025_연간 전체실적_전체.xlsx` | Korea annual export performance (203K rows, queried for 국내수출품목) |
| `26.N월 SPS 통보문\*.docx` | Source WTO SPS notification files (one per notification or one joint file for multi-country notifications) |
| `26.N월 SPS 통보문\*_번역.docx` | Bilingual output files produced by the tool |
| `sps_tool\` | The tool's Python application folder |

---

## Excel Structure

**Workbook:** One file per year. Sheets: `★매뉴얼('26.4.)`, `26.4월`, `26.3월`, `26.2월`, `26.1월 `.

The active month sheet is auto-detected by current date (overrideable in Settings).

**Row pre-population:** Rows are pre-populated by an upstream process (ePing) with basic identification fields before the analyst begins processing. The tool finds the correct row by matching `문서번호` (column 7) and fills in the remaining fields.

**Column mapping (1-based):**

| Col | Field | Pre-filled? | Tool Action |
|---|---|---|---|
| 1 | 담당자 | ✅ Yes | Read only |
| 2 | 순번 | ✅ Yes | Read only |
| 3 | 중요도 | ❌ Empty | LLM recommends; yellow fill if uncertain |
| 4 | 통보유형 | ❌ Empty | Script fills from document header (일반/긴급) |
| 5 | 통보국 | ✅ Yes | LLM provides Korean country name |
| 6 | 배포일 | ✅ Yes | Used as Day 0 for all date formula calculations |
| 7 | 문서번호 | ✅ Yes | **Primary key for row matching** |
| 8 | 제목 | ❌ Empty | LLM translates to Korean |
| 9 | 내용 | ❌ Empty | LLM translates/summarizes in 개조식 style |
| 10 | 해당품목 | ❌ Empty | LLM translates with 학명 in parentheses |
| 11 | 목적 | ❌ Empty | LLM normalizes to standard Korean institutional phrases |
| 12 | 해당국가 | ❌ Empty | Script extracts; LLM fallback; '모든 교역국' if all partners |
| 13 | 의견마감일 | Partially filled | Script calculates if absent (see date logic) |
| 14 | 발효일 | Partially filled | Script calculates if absent |
| 15 | 국내수출품목 | ❌ Empty | Script looks up in export file; `-` if not triggered |
| 16 | 관련부서 | ❌ Empty | LLM maps from manual; yellow if multiple plausible depts |
| 17 | 주간보고 | ❌ Empty | LLM drafts one-line 개조식 summary |
| 18 | 구분 | ❌ Empty | LLM classifies (동물/식물/식품) |
| 19 | 품목 | ❌ Empty | LLM extracts short product label |
| 20 | 검토메모 | ❌ Empty | Tool writes reviewer notes for flagged fields |
| 21–22 | (Internal) | ❌ Empty | Reserved for approval date, status |

**Row matching rules:**
- Match on column 7 (`문서번호`) — normalize whitespace and uppercase before comparing
- For joint notifications (multiple document numbers like `G/SPS/N/BDI/149, G/SPS/N/KEN/358, ...`): match on the first ID in the cell string
- If no match found: warn in UI, skip all Excel writes for that file (do not modify the workbook)
- For date fields (의견마감일, 발효일): do not overwrite cells that already have a non-empty value

---

## Word Document Processing

### Source Document Structure

WTO SPS notification files are bilingual tables (English/Korean form labels; source-language content). Key structural facts:
- The document number (`G/SPS/N/XXX/YYY`) appears in the document header paragraphs or the first table rows
- Notification type (regular/emergency) is identified from the document header
- Addendum documents have an "Addendum" header and different content structure
- Objective checkboxes appear as `[X]` markers in the objectives section
- Comment deadlines and entry-into-force dates appear in specific numbered sections

### Field Extraction (Script)

The parser (`parser.py`) extracts these fields using label-matching across table cells:

- `doc_number` — WTO document symbol (G/SPS/N/XXX/YYY)
- `is_emergency` / `is_addendum` — detected from header text and document symbol
- `notifying_member` — country name from "Notifying Member" row
- `products` — products/items from "Products covered" row
- `regions` — "모든 교역국" if `[X] All trading partners` is checked; otherwise specific country/region name
- `title` — from "Title" row
- `description` — from "Description of content" row (main translatable text)
- `objectives_korean` — list of standard Korean phrases for checked objectives
- `comment_deadline_raw` — raw text from "Final date for comments" row
- `entry_force_raw` — raw text from "Proposed date of entry into force" row
- `source_language` — detected from character distribution (`en`, `es`, `pt`)

Filename fallback for document number: `GSPSNBRA2474` → `G/SPS/N/BRA/2474`, `GSPSNCRI347A1` → `G/SPS/N/CRI/347/Add.1`

### Bilingual Word Output

- Output filename: `{source_stem}_번역.docx` in the same folder as the source
- Korean text is inserted as a **second paragraph within each content cell**, below the original text
- Font: 맑은 고딕, matching the source document's font size
- Scientific names formatted as: 국문명(학명)
- If source language is Spanish or Portuguese:
  - Apply **lime/green fill** (`#CCFF99`) to 제목 and 내용 cells in both the Excel and Word output
  - This signals that the English version has not yet been confirmed
  - The analyst manually removes the lime fill after verifying against the English source

---

## Date Calculation Logic

**Base date:** `배포일` from Excel column 6 (= WTO circulation date). The same date is also embedded in the Word document body; the Excel value takes precedence.

| Raw text pattern | Output |
|---|---|
| Explicit calendar date (e.g., "16 March 2026", "24 September 2025") | Normalize to `DD/MM/YYYY` |
| Formula: "N days from circulation/publication" | `배포일 + N days` |
| Formula: "N months after publication" | `배포일 + N months` |
| "To be determined" / "Por determinar" / empty | Write `추후결정` |
| Emergency notification + no deadline stated | Write `-` |
| Cell already has a value | **Do not overwrite** |

The source expression (e.g., "sixty days from circulation") is stored in the processing log alongside the resolved date, for auditability.

---

## 중요도 Classification Logic

Rules from `★매뉴얼('26.4.)` sheet:

**검토** (Review required):
- 해당국가 is Korea (한국) AND Korea has matching export products to the notifying country
- 해당국가 is 모든 교역국 AND Korea has matching export products
- MRL is stricter than the Korean domestic standard (regardless of export volume)
- Sensitive domestic issues: electronic phytosanitary certificate introduction, GMO/LMO regulations, BSE, beef processing facility registration, customs tightening measures

**참고** (For reference):
- Notification is from one of the 24 export quarantine agreement countries targeting all trading partners
- MRL is same as, weaker than, or absent compared to Korean domestic standard
- Minor domestic relevance — useful for quarantine practitioners but not requiring action

**- (dash)** (Not applicable):
- Under jurisdiction of another ministry (해수부, 식약처, 기후부/환경부)
- Third-country bilateral restriction (Korea not involved as importer or exporter)
- No domestic export products, no export quarantine negotiations in progress

The 24 export quarantine agreement countries: USA, Japan, EU, China, Australia, Canada, New Zealand, Philippines, Vietnam, Taiwan, Thailand, Singapore, Indonesia, Malaysia, Hong Kong, UAE, Russia, Kazakhstan, Mexico, Chile, Peru, Colombia, India, Saudi Arabia.

Yellow fill is applied to the 중요도 cell when the LLM's confidence is low or when two rules give conflicting signals.

---

## 구분 (Category) Classification

| 구분 | Covered content |
|---|---|
| 식물 | Plant quarantine measures, plant pest/disease regulations, seeds and planting materials, wood packaging materials, oilseed crops (excl. processed), mushrooms/ginseng/medicinal plants, organic/eco-friendly produce, insects/sericulture, plant fertilizers, plant-related GMO/LMO, invasive species (plant) |
| 동물 | Animal quarantine measures, veterinary drug MRL (livestock products), livestock feed and feed additives, pet animals, wildlife/hunting trophies, antibiotic regulations, import suspension for animal diseases (HPAI, ASF, FMD, BSE), animal GMO/LMO |
| 식품 | Pesticide MRL (agricultural products), processed food standards, food additives, Codex Alimentarius standards, new food materials (배양육 등), aquatic/fisheries products, tobacco |

---

## 관련부서 Mapping

Departments are mapped from 구분 + product sub-type using the manual table. Multiple departments are written on separate lines within the cell. Key rules:

- **Plant quarantine** (seeds, wood, pests): 수출지원과(검본) ± 위험관리과(검본)
- **Plant oilseed crops**: 원예산업과
- **Mushrooms/ginseng/medicinal**: 농식품수출진흥과, 원예산업과
- **Organic/eco-friendly**: 친환경농업과, 인증관리과(농관원)
- **Insects/sericulture**: 그린바이오산업팀
- **Plant fertilizer**: 농산업수출진흥과
- **Plant GMO/LMO**: 수출지원과(검본), 연구개발과(농진청), 생물안전성과(농과원) + notify 검역정책과 식물계
- **Seeds**: 종자산업지원과, 연구개발과(농진청)
- **Animal quarantine (livestock)**: 동물검역과(검본), 위험평가과(검본)
- **Livestock products/meat**: 위험평가과(검본), 축산물수출위생팀(검본)
- **Veterinary drug MRL**: 동물약품평가과(검본), 축산물수출위생팀(검본)
- **Antibiotics**: 조류인플루엔자방역과, 동물약품평가과(검본), 축산물수출위생팀(검본)
- **Pet animals**: 반려산업동물의료과, 축산환경자원과, 축산물수출위생팀(검본)
- **HPAI**: 동물검역과(검본), 위험평가과(검본), 축산물수출위생팀(검본), 조류인플루엔자방역과
- **ASF/FMD**: 동물검역과(검본), 위험평가과(검본), 축산물수출위생팀(검본)
- **Feed/feed additives**: 축산환경자원과, 축산물수출위생팀(검본)
- **Wildlife/invasive species**: 기후부
- **Pesticide MRL (agricultural)**: 농식품수출진흥과, 잔류화학평가과(농과원); if simultaneous vet drug MRL: add 동물약품평가과(검본)
- **Agricultural quality**: 품질조사과(농관원), 식약처
- **Heavy metals/mycotoxins**: 안전성분석과(농과원), 식약처
- **Enoki mushroom (Listeria)**: 농식품수출진흥과, 소비안전과(농관원) — notify same day
- **Processed food / food additives / Codex**: 식약처
- **New food materials**: 연구개발과(농진청), 식약처
- **Food GMO/LMO**: 연구개발과(농진청), 생물안전성과(농과원), 식약처
- **Aquatic/fisheries products**: 해수부, 식약처
- **Feed standards**: 축산환경자원과
- **Tobacco**: 식약처, 보건복지부

Yellow fill applied if the LLM identifies more than one plausible mapping.

---

## 국내수출품목 Lookup

**Trigger rule:**
- If 해당국가 = 한국 OR 모든 교역국 → perform lookup in export file
- If 해당국가 = a specific third country (not Korea) → write `-`
- If under another ministry's jurisdiction → write `-`

**Export file:** `2025_연간 전체실적_전체.xlsx`, sheet `연간실적`
- Columns: 기준년월, 국가코드, 국가명, 구분(E=export), AGCODE, HSCODE, 품목명, 당월중량, 누계중량, 당월금액, ...
- 203,178 rows; loaded into pandas at app startup with caching

**Lookup method:**
1. Filter for rows where 국가명 = notifying country AND 구분 = 'E' AND 누계중량 > 0
2. Narrow by HS chapter inferred from the product description keywords
3. Return comma-separated `품목명` values (up to 5), e.g., `사과(신선, 건조), 배(신선), 감귤`
4. If no match: write `-`
5. If HS chapter mapping is uncertain: write best guess + yellow fill

---

## LLM Processing

**Model:** `claude-sonnet-4-6` via Anthropic API

**Single combined prompt** (not separate per-field prompts) receives:
- All extracted fields from the parser
- Export lookup result
- Terminology dictionary (terminology.json)
- 중요도, 구분, 관련부서 classification rules from the manual

**Returns JSON with:**
- `제목`, `내용`, `해당품목`, `목적`, `해당국가`, `통보국_kr` — Korean translations
- `주간보고` — one-line 개조식 summary
- `구분` — 동물/식물/식품 with reasoning
- `중요도` — 검토/참고/- with reasoning
- `관련부서` — newline-separated department list
- `품목` — short product label
- `flags` — list of field names that are uncertain
- `source_language` — detected source language

**Korean style requirements:**
- 목적: use only approved standard phrases (식품안전; 동물위생; 식물보호; 동식물 병해충 또는 질병으로부터 사람 보호; 해충으로 인한 피해로부터 영토 보호)
- 내용, 주간보고: 개조식 (noun-phrase, concise bullet style)
- 학명: included in parentheses as 국문명(학명)
- Terminology dictionary (`terminology.json`): ~100 standard term translations from the ★매뉴얼 sheet, used as few-shot examples in the prompt

---

## Uncertainty Flagging

| Condition | Flag |
|---|---|
| LLM confidence below threshold or multiple plausible values for any field | Yellow cell fill on that specific column |
| Source language is Spanish or Portuguese (제목, 내용) | Lime/green cell fill (#CCFF99) |
| Export lookup HS chapter mapping was uncertain | Yellow fill on 국내수출품목 |
| Document number not found in Excel | Skip + warning in UI (no Excel writes) |
| Multiple plausible 관련부서 options | Yellow fill on 관련부서 |
| Reviewer notes written | Column 20 (검토메모) populated |

Analyst clears yellow/lime fills manually after reviewing and confirming the field.

---

## Processing Modes

### Single-File Mode (daily use)
1. Analyst drops one `.docx` file onto the web UI
2. Parser extracts all fields
3. Export lookup runs (if triggered)
4. LLM processes in one API call (~15–30 seconds)
5. Date calculations applied
6. Excel row matched by 문서번호, all fields written
7. `*_번역.docx` generated in the source folder
8. UI shows result card: status, flags, Excel row number, Word filename

### Batch Mode (multiple files at once)
1. Analyst selects multiple `.docx` files
2. Each file processed sequentially with progress indicator
3. All Excel writes committed after processing completes
4. Summary table shows all results, flags, and errors

---

## Weekly Report

**Not automated.** The analyst writes the weekly email/PDF report manually using the Excel as reference. The tool's contribution is limited to:
- Populating the 주간보고 (column 17) one-line summary for each row
- Format: `Header + brief narrative on any 긴급 notifications + numbered list of all notifications, grouped by 담당자`

---

## Application Architecture

```
sps_tool/
  app.py            — Flask server, processing pipeline, routes
  parser.py         — Word document field extractor
  llm.py            — Claude API (translation + classification)
  date_engine.py    — Date formula resolver
  export_lookup.py  — Export performance file querying (pandas)
  excel_writer.py   — Row matching and cell writing (openpyxl)
  word_writer.py    — Bilingual _번역.docx generator
  terminology.json  — ~100 standard Korean term translations
  start.bat         — Launch script (installs deps, opens browser)
  .env              — API key + file paths (not committed)
  templates/
    index.html      — Main drag-and-drop UI
    settings.html   — API key and file path settings
```

**Stack:** Python 3.12+, Flask, python-docx, openpyxl, pandas, anthropic SDK, python-dotenv, python-dateutil

**Launch:** `start.bat` double-click → installs dependencies on first run → opens `http://localhost:5000`

---

## Design Principles (Revised)

The original design principle of separating deterministic, LLM, and human-validation layers is preserved:

| Task | Owner | Method |
|---|---|---|
| Field extraction from Word | Script (parser.py) | Label-matching on WTO form table cells |
| Date arithmetic | Script (date_engine.py) | Regex + dateutil formula resolution |
| Export item lookup | Script (export_lookup.py) | Pandas query on 203K-row export file |
| Excel row matching + writing | Script (excel_writer.py) | openpyxl + 문서번호 primary key |
| Translation + normalization | LLM (llm.py) | Single combined Claude API prompt |
| Classification (중요도, 구분, 관련부서) | LLM | Rules from ★매뉴얼 embedded in prompt |
| 주간보고 drafting | LLM | Included in combined prompt |
| Uncertainty review | Human | Yellow/lime cell flags in Excel |
| Weekly report writing | Human (manual) | Uses completed Excel as reference |

The LLM has no responsibility for arithmetic, spreadsheet formatting, or file I/O. Scripts have no responsibility for translation style or policy-sensitive classification judgment.

---

## Exception Handling (Concrete Triggers)

The system routes outputs for human review when:
- `문서번호` cannot be found in the Excel — entire file is skipped with a UI warning
- LLM returns a field in its `flags` list — yellow cell applied to that field
- Export lookup returns an uncertain HS chapter match — yellow cell on 국내수출품목
- Source language is Spanish or Portuguese — lime cells on 제목 and 내용 pending English version
- More than 2 관련부서 lines are generated — yellow cell on 관련부서
- The Word document contains no recognizable table structure — error shown in UI, no outputs written

---

## Phased Implementation

### Phase 1 — Core (implemented)
- Flask app shell with drag-and-drop file upload
- Word parser: all standard WTO SPS form fields
- LLM combined prompt: translation + classification + 주간보고
- Excel row matching by 문서번호 + cell writing
- Bilingual `*_번역.docx` generation with lime flagging
- Yellow cell uncertainty flags

### Phase 2 — Refinement
- Date engine: improved formula parsing for edge cases (e.g., "approximately 110 days")
- Batch mode progress: per-file status updates in real time (SSE)
- Terminology auto-update: confirm a translation → add to terminology.json

### Phase 3 — Export lookup tuning
- HS chapter inference improvement using product description embeddings
- Fuzzy country name matching for the export file

### Phase 4 — Audit trail
- `processing_log.jsonl`: one entry per file with source expressions, LLM reasoning, flag history
- Log viewer in the UI settings page

### Phase 5 — Historical precedent
- Store finalized (yellow-cleared) outputs as precedent examples
- Feed nearest similar examples into the LLM prompt for consistency
