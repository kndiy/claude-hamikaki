# HOP DONG GMPC

## Purpose and goals
Draft GMPc's customer contracts and related paperwork (contracts, đề nghị tạm ứng / thanh toán, biên bản nghiệm thu / thanh lý, …) from GMPc templates. This is the main contract workstream. OKS-issued contracts live in the sibling project `HOP DONG OKS`, which uses different templates.

## Tech stack and key dependencies
- Python 3.13 (stdlib; `defusedxml`, `lxml` for the docx skill scripts).
- This machine has no pandoc, LibreOffice, or `zip`. Read .docx with `scripts/dump_docx.py` and rezip with Python `zipfile`. Check the output in Word.

## Structure and key files
- `templates/`: GMPc templates (**none yet**).
- `scripts/dump_docx.py`: prints a .docx's text.
- `scripts/xlsx_reprice.py`: reprices a GMPc M&E estimate workbook to a target total by editing the original XML.
  - Recomputes and writes every formula's cached value. This is needed because the estimates use manual calc mode, and an openpyxl save leaves the formula cells blank.
  - Countable equipment (listed in `keep_units`) keeps its original price, or gets the fixed factor given as the optional 7th argument `keep_k`. Quantities are never changed.
  - Set `PYTHONIOENCODING=utf-8` when running it, otherwise the final print fails on the Windows console.
  - Applies header text edits from a JSON file and writes a summary JSON and a CSV of the changed rows.
  - Options (2026-10-30): `--qty-k/--qty-units` (uplift quantities of measured units, rounded down), `--vt-nc VT,NC` (separate VT/NC targets so a 10%/8% VAT stays put), `--exact` (rounding residual into one 'lô' row so contract figures hit the target to the đồng).
- `scripts/dump_highlight.py`: prints a .docx's text with highlighted runs wrapped as `[[HL:...]]`. Convert a .doc to .docx with Word COM first.
- `scripts/vn_number.py`: `dots()` and `words()` (đọc số tiền bằng chữ), with half-up rounding as in Excel.
- `khach-hang/<năm>-<số NAS> <KHÁCH>/`: one folder per customer contract, with `NOTES.md` (contract number, value, payment schedule, status, open items), `raw/` (copies from the NAS, with the source path recorded in NOTES.md), and `output/`.

## Commands
- Read: `python scripts/dump_docx.py <file.docx>`
- Fill a template: unzip, run `merge_runs.py`, string-replace in `word/document.xml`, rezip with Python, then `validate.py --original`.

## Conventions, constraints, decisions
- The GMPc group rules apply in full: font (Calibri Light for every document until the new brand launches on 25/12/2026), the exact GMPc company header details (post-2025 addresses), confidential customer data, and read-only NAS.
- Use dots for numbers in Vietnamese text and commas in English text.
- Contract number = signing date: `DDMM/HDTV/<Khach>-GMPc/YYYY` (e.g. `3010/HDTV/HungThai-GMPc/2026` signed 30/10/2026). User rule, 2026-10-30.
- Advance invoices (user rule, 2026-10-30): the advance (tạm ứng) is requested with the Đề nghị tạm ứng only. GMPc issues the VAT invoice right **after** the money arrives, never as part of the advance dossier, so no invoice has to be cancelled when a customer signs and then backs out. Apply this to every new contract; the ECO CNC template still lists the invoice in the advance dossier.

## Open questions and next steps
- Where are the GMPc templates on the NAS (contract, advance/payment request, etc.)? Copy them into `templates/` and record their sources here.

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| defusedxml | `python -c "import defusedxml, lxml"` | recommended | `pip install defusedxml lxml` | installed |
<!-- launcher:tools:windows:end -->

## Change log
- 2026-10-08: Project created, split from OKS contracts because the templates differ.
- 2026-10-28: `khach-hang/2026-1 HSC/`: M&E construction contract and appendix 01 at 70 bn (scaled from 54.3 bn by unit price only). The contract is based on the TIDO 2024 draft. Open items are listed in its NOTES.md.
- Decision (2026-10-28): estimate values are raised through unit prices, never by padding quantities beyond the design take-off.
- 2026-10-28 (v3): appendix rebuilt with `xlsx_reprice.py` (cached values verified in Excel), target ≈ 71,64 bn, not a round figure (user). Lesson: never save these workbooks with openpyxl.
- 2026-10-28 (v4): alternative HSC version at 85.700.065.124 đ, alongside v3. Paused here (user: "lưu giúp mình"). v3 or v4 has not been chosen yet. Later the contract price table was simplified to 5 columns with one VAT rate of 8% (user); the appendix keeps the VT/NC split. Equipment ×1,30, other rows ×1,836. `xlsx_reprice.py` gained the `keep_k` argument for this (the default of 1 keeps v3 reproducible).
- 2026-10-28 (v5): HSC re-targeted to 85 bn **after** 8% VAT (84.997.620.410), built on the user's own edited contract from the NAS. `khach-hang/2026-1 HSC/update_price.py` swaps only the money figures. Rule: once the user edits a contract on the NAS, that file becomes the base. Don't rebuild it from the template. Paused here at v5 (user: "lưu lại giúp mình"). Open review points are listed in that folder's NOTES.md.
- 2026-10-30: `khach-hang/2026-HUNG THAI/`: HS GMP training contract (25.920.000 đ incl. 8% VAT, 60/40) filled from the ECO CNC CGMP training contract (`N:\...\Năm 2026\14. ECO CNC - Đào tạo CGMP\`) by `make_contract.py`. That ECO CNC file is the de-facto template for GMPc training contracts. Word COM (read-only open + PDF export) works on this machine for page counts. Number 3010/HDTV/HungThai-GMPc/2026, signed 30/10/2026. Surcharge clause removed, 20-30 trainees kept (user). Draft complete, 8 pages.
- 2026-10-30: `khach-hang/2026-1 MERAP TK/`: payment dossier #5 for the Merap design contract (FPF appraisal, 10% = 294.317.334). `make_ls5.ps1` edits copies with Word COM, keeps the .doc format, rewrites the English to match the highlighted Vietnamese using the contract's wording, and drops the highlights. Convention: yellow highlight in a user's file marks Vietnamese that changed, and the English has to follow it. Open items are in its NOTES.md (a second PL 03 draft that may lower the value; the invoice bank account differs). Paused here (user: "lưu lại nhé"). Output has not been copied to the NAS yet; that is the user's step.
- 2026-10-30: `khach-hang/2026-1 HSC/compare_versions.py` builds the comparison workbook Gốc / V3 / V5 (`output/261030_HSC_So sanh gia Goc-V3-V5.xlsx`). The user corrected the comparison from V2 to V3. Found an error in the original estimate, carried into every version: BMS row 46 (Tủ DDC, 25 tr) has no J formula, so it is excluded from the totals. Not fixed yet; the user decides. Lesson: some qty cells in these estimates are stored as text ('1'), and some rows have H but no J, so check that sheet totals reconcile.
- 2026-10-30: HSC V3-KL / V5-KL. The user overrode the 2026-10-28 "no quantity padding" decision: measured-unit quantities (m, mét, m², kg) may rise by up to +10% (hao hụt / dự phòng, noted in the appendix). The rest of the uplift goes into unit prices, and countable equipment keeps its original qty and price. Claude flagged the lump-sum-by-design risk. `update_rows.py` swaps the Điều 2.1 per-system figures. Fixed an `xlsx_reprice.py` bug: text formulas such as `=B16` were cached as 0. Details are in the HSC NOTES.md. Later the same day V5-KL changed: equipment +15% (V5 only, user), with the rest in non-countable prices ×1,5964.
- 2026-11-04: Merap ĐNTT "lần 4" (it actually requests installment 5, FPF; installment 4 is on hold). `make_ls4.ps1` adds English for the highlighted Vietnamese and removes an empty signature-table row so the file stays at 2 pages. Title convention (user): the ĐNTT "lần N" counts the requests sent, not the contract installment, so the English title is "FOURTH PAYMENT REQUEST". Note: .ps1 files with Vietnamese text must be saved as UTF-8 **with BOM** for Windows PowerShell 5.1.
- 2026-10-28 (v2): HSC uplift moved onto non-countable rows only (m, m², kg, lô, gói…; ×1,557); countable equipment is unchanged. Payment schedule changed to 30/30/30/10 to match the work done at each milestone.
