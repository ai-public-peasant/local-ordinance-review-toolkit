# Technical Appendix for AI Assistants

This document is intended for another AI coding assistant that needs to understand or modify the repository.

## Runtime

- OS tested: Windows 11 / PowerShell
- Python: 3.11+
- External site: ELIS, `https://www.elis.go.kr`
- Python dependencies: see `requirements.txt`
- Law.go.kr Open API `OC`: not required by this repository's default scripts.

## Authentication / OC Key

The default pipeline does not call the Law.go.kr Open API. It scrapes public ELIS web pages and downloads public attachment files through ELIS endpoints. Therefore there is no `OC` parameter, no API-key environment variable, and no login step.

If this repository is combined with Korean Law MCP or direct Law.go.kr Open API calls, that separate tool may require an `OC` key. Keep that distinction clear in user-facing documentation: ELIS collection in this repo does not need OC; external Law.go.kr API tools may need OC.

## Pipeline

1. `scripts/verify_elis_targets.py`
   - Input: CSV with `권역,지자체명,ctpvCd,sggCd`.
   - Optional count columns: `예상건수` or any column starting with `ELIS_목록건수`.
   - Output: verification CSV/Markdown under `data/`.

2. `scripts/collect_elis_fulltext.py`
   - Input: target CSV.
   - Output root: `02_원천수집/ELIS_자치법규_원천자료_YYYYMMDD/{entity}`.
   - Collects list JSON/XLSX, detail HTML, body text, attachment originals, extracted attachment text, and integrated search text.
   - Date can be fixed with environment variable `ELIS_COLLECT_DATE=YYYYMMDD`.
   - Attachment parsing:
     - `.hwp` / OLE CFB files: custom HWP5 text extraction using `olefile`, reading compressed/uncompressed `BodyText/Section*` streams.
     - `.hwpx`, `.docx`, `.xlsx`, `.zip`: ZIP/XML text extraction using Python `zipfile`.
     - `.pdf`: page text extraction with `pypdf`.
   - HWP parsing is best-effort text extraction, not a layout-preserving conversion.

3. `scripts/run_body_delivery_all.py`
   - Searches law body text / detail article blocks.
   - Produces per-entity Excel and DOCX files under `04_제출서식/전체_전달용_본문기준_YYYYMMDD`.
   - Use `--only` and `--exclude` for scope control.
   - `--match-reference-names` enables heavier direct matching against regional ordinance names.

4. `scripts/scan_attachment_delivery_all.py`
   - Reads attachment manifest and extracted attachment text.
   - `--source-date` selects the raw collection folder suffix.
   - Produces attachment review workbooks under `04_제출서식/전체_별지정밀검토_YYYYMMDD`.

5. `scripts/merge_body_attachment_delivery.py`
   - Merges body and attachment review files.
   - `--body-date` and `--attachment-date` select input result folders.

## Important Design Choices

- The scripts intentionally produce candidate review material, not final legal conclusions.
- Attachment text extraction is intentionally simple and dependency-light. It is suitable for keyword discovery but not for faithful reconstruction of tables or official forms.
- Generated folders and documents are ignored by Git because they can contain large downloaded files and internal work products.
- `data/targets_sample.csv` is the only included target dataset.
- The current pattern logic is tailored to a Jeonnam/Gwangju integration example but can be edited in `patterns_for(entity)` inside body and attachment scanning scripts.
- The `권역` column currently drives pattern selection. `전남` and `광주` have separate search patterns.

## Common Modification Points

- Add a new region type:
  - Edit `patterns_for(entity)` in `run_body_delivery_all.py`.
  - Edit `patterns_for(entity)` in `scan_attachment_delivery_all.py`.
  - Add target rows to the CSV with the new `권역` value.

- Change output columns:
  - Edit `save_xlsx`-style functions in body/attachment scripts.
  - Keep the first worksheet named `제출서식` if you want merge script compatibility.

- Fix ELIS HTML parsing:
  - List parsing: `fetch_law_list`.
  - Detail parsing: `fetch_detail_articles` or equivalent detail collection logic.
  - Attachment extraction: functions in `collect_elis_fulltext.py`.

## Validation Checklist

Run syntax checks:

```powershell
python -m py_compile .\scripts\verify_elis_targets.py .\scripts\collect_elis_fulltext.py .\scripts\run_body_delivery_all.py .\scripts\scan_attachment_delivery_all.py .\scripts\merge_body_attachment_delivery.py
```

Run a small live test before full collection:

```powershell
python .\scripts\verify_elis_targets.py --targets .\data\targets_sample.csv
python .\scripts\collect_elis_fulltext.py --entities-csv .\data\targets_sample.csv --only "여수시"
python .\scripts\run_body_delivery_all.py --targets .\data\targets_sample.csv --only "여수시"
python .\scripts\scan_attachment_delivery_all.py --targets .\data\targets_sample.csv --source-date YYYYMMDD --only "여수시"
```
