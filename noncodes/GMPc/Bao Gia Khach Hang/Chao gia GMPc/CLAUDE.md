# Chao gia GMPc

Created 2026-10-13 by the launcher.

## Purpose and goals
A reusable tool for drafting GMPc quotations (bản chào giá). For each customer request, the user sends the details (customer, sector, which parts of the work, prices, timing, payment terms, language). Claude fills in a request file (phiếu yêu cầu) and generates the quote .xlsx from the company's 2025 templates.

## Tech stack and key dependencies
- Python 3.13 + openpyxl 3.1.5 + Pillow. Without Pillow, openpyxl silently drops the logo and seal, so `make_quote.py` imports PIL to fail loudly instead. `tomllib` (stdlib) reads the request files.
- Excel (COM) is used only to export PDFs when checking the layout.

## Structure and key files
- **Template source (read-only, never edit):** `..\2025 Mau chao gia\`, a local folder in the parent group, not the NAS. It has 8 sectors: 1 Dược phẩm, 2 Mỹ phẩm, 3 TPBVSK, 4 Thú y, 5 Kho GSP/GDP, 6 ĐĐKSX MP, 7 Tái đánh giá, 8 Lập dự án đầu tư. Each sector has a step file for each phase: 1 Kế hoạch tổng thể, 2 Thiết kế XD/M&E, 3 Giám sát thi công, 4 Đào tạo + hồ sơ. A "4 bước" file combines the four steps. `Song ngu.xlsx` (a real customer example) and the `3. Bảng tính…` files are skipped.
- `scripts/build_catalog.py` reads the templates and writes:
  - `catalog/catalog.json`: the sections and items of 43 templates.
  - `catalog/DANH_MUC.md`: a human-readable menu of template ids and section codes.
  - `catalog/en_goi_y.json`: Vietnamese strings that still have no English (currently empty).
- `catalog/song_ngu.json`: the **bilingual library**, one English text for each Vietnamese string, keyed by the whitespace-normalized VN text. Each entry has `trang_thai` set to `cho_duyet` or `da_duyet`, and some have a `ghi_chu`.
- `scripts/_seed_song_ngu.py`: the one-off seed of the library (2026-10-13), kept for history. Don't rerun it, because it would overwrite review work.
- `scripts/song_ngu_xlsx.py xuat|nhap`: round-trips the library through `catalog/Duyet song ngu.xlsx` so the user can review it in Excel.
- `mau_phieu_yeu_cau.toml`: the request template, with comments explaining every field.
- `scripts/make_quote.py <phieu.toml>`: builds the quote next to the request file.
- `output/<YYMMDD Ten khach>/`: one folder per quote, holding `phieu_yeu_cau.toml` and the generated xlsx (versioned ` vN`, never overwritten).

## Commands (build / run / test / lint)
- `python scripts/build_catalog.py`: rerun after the templates change.
- `python scripts/make_quote.py "output/<YYMMDD Ten khach>/phieu_yeu_cau.toml"`
- `python scripts/song_ngu_xlsx.py xuat`, then the user edits the xlsx, then `python scripts/song_ngu_xlsx.py nhap`.
- Visual check: open the xlsx in Excel via PowerShell COM and use `ExportAsFixedFormat(0, pdf)`, then read the PDF.

## Workflow per customer
1. The user sends the customer info and needs. Claude picks the template and sections from `DANH_MUC.md` and **asks for anything missing**, especially prices, which Claude never invents.
2. Claude creates `output/<YYMMDD Ten khach>/phieu_yeu_cau.toml` from `mau_phieu_yeu_cau.toml` and runs `make_quote.py`.
3. Claude reports the review items. Yellow cells and the `Ghi chu soan thao` sheet show what still needs input. Before sending, remove the yellow fills and delete the review sheet.

## Conventions, constraints, decisions
- 2026-10-13 decisions (user):
  - The user provides prices for each quote. Missing or 0 prices are highlighted yellow.
  - The 2025 template set is the master. The Cafune 22/09/2026 quote is not used as a base.
  - Language is chosen per quote: `vn`, `vn+en` (two sheets) or `song_ngu` (VN with EN on the next line in the same cell, and labels as "VN / EN").
  - Quotes are stored in this project under `output/`.
  - **Build a bilingual library.** The old ENG sheets are not line-by-line translations of the VN sheets: they have extra lines, different numbering, pharma wording inside cosmetics templates, and a wrong VAT formula. So the EN output always comes from `song_ngu.json`, and VN and EN always cover the same scope.
- Rendering: the VN sheet of the chosen template is the base (it keeps the header, logo, column widths and print setup). The body from the intro row down is rebuilt using the template's own row styles. Font is Calibri Light everywhere. The company header is rewritten with the GMPc group strings. The date is fixed (no `=NOW()`). Items are renumbered after lines are dropped or added. Totals are `SUM`, VAT is `ROUND(total*rate)`.
- **Legal-risk wording is blocked** (regex in both scripts: `ngoại giao | thúc đẩy hồ sơ | (chi phí|bao gồm)…tiếp đoàn`):
  - GDP item 12 "Chi phí ngoại giao, tiếp đoàn đánh giá" and the GDP footnote are dropped and listed in the review sheet.
  - Section III of `kho_gsp/chao_gia_gsp_thiet_ke_phan_mem` ("Chi phí thúc đẩy hồ sơ và tiếp đoàn đánh giá…", 200M in the template) makes the script exit with an explanation.
  - Legitimate "hỗ trợ… tiếp Đoàn đánh giá" (preparing the client for the inspection) is allowed.
- The seal and signature image is off by default (`dong_dau = true` to include it). The seal still reads "H. Hoài Đức". Only send a stamped file once the content is final, preferably as a PDF.
- VAT 8% applies until 31/12/2026 (Decree 174/2025). For a quote dated after that, the script highlights VAT yellow and requires `vat` plus `ghi_chu.vat`.
- English terminology: DAV (Cục QLD), VFA (Cục ATTP), DAH (Cục Thú y), Provincial Department of Health (Sở Y tế), Investor (CĐT), MEP, pre-audit (mock inspection) for tiền đánh giá. "Self-inspection" is used only for tự thanh tra.

## Quotes made
- **261013 Haki Medica** (GDP, 160 m2 warehouse for dental anaesthetics, HCMC; 70,000,000 VND before VAT):
  - Built from the user's CVI quote, copied from the NAS `Q:\2026\45. CVI _ GDP (Ha Noi)\2. Bao gia _ Hop dong\1. Chao gia\260630 Bao gia GDP CVI.xlsx` to `output/261013 Haki Medica/data_raw/`.
  - VN only, stamped (`dong_dau = true`). Section I is "Hỗ trợ" (free), section II has the CVI's 16 lines.
  - The exclusion note uses neutral wording: "lệ phí thẩm định… các chi phí phát sinh khác (nếu có)", replacing the user's phrasing "chi phí phát sinh trong quá trình tiếp đoàn".
  - **Sent to the customer by the user on 2026-10-13** as `261013 Bao gia GDP Haki Medica v2.xlsx` (75,600,000 VND incl. 8% VAT). The user will report updates. Next revision → ` v3`. The first file (without ` v2`) is an unstamped draft that was never sent.
  - 2026-10-16: The customer asked GMPc to introduce staff and estimate investment costs. Sales approach (user): sign the contract first and give details after. The user's reply only says that, once the quote is accepted, GMPc will assign its project consultants to advise. No layout, rack list or cost ranges are given before signing. GMPc has **no candidate pool** for the customer's chuyên môn staff, so it never promises to introduce candidates. For the contract: the timeline should also depend on the customer having its người phụ trách chuyên môn, and should say "Cơ quan tiếp nhận" instead of "Sở Y tế".

## Regulatory references (copied 2026-10-13, read-only source)
- Full write-up with verification status: `docs/GDP - pham vi va yeu cau kho (thuoc te nha khoa).md`.
- Copied from `K:\0_UP_DRIVE\TLCN_VBPL\1_Duoc_Pham\VBPL\` to `data/raw/vbpl/`:
  - VBHN 10/VBHN-BYT (2025) GDP
  - Luật 44/2024/QH15
  - TT 18/2026/TT-BYT (KSĐB)
  - TT 06/2017/TT-BYT (danh mục thuốc độc)
  - NĐ 163/2025 is a scanned PDF and has not been text-extracted.
- Findings used for the Haki GDP advice:
  - GDP minimum storage area is 30 m2 / 100 m3.
  - Temperature mapping is required, with at least one self-recording logger (1-2 readings per hour).
  - A distribution management software connected to the national pharma data system is required.
  - Since 01/07/2025 the dossier goes to the "Cơ quan tiếp nhận" (the provincial agency), per TT 11/2025.
  - **Mepivacain, Bupivacain, Levobupivacain and Procain injections are on the thuốc độc list (TT 06/2017). Lidocain and Articain are not.** Toxic drugs need a separate storage area and an XNT log per TT 18/2026 Điều 11 + Phụ lục XVI.

## Open questions and next steps
- Haki Medica (quote sent, waiting for customer feedback): ask which anaesthetics they will trade. If Mepivacain is included, the scope and the quote need the thuốc độc items. Dental supplies are medical devices (TTBYT), outside the GDP scope.
- Haki also wants **ISO 13485** (it mainly distributes medical devices). GMPc only consults; an independent third-party body certifies. Web research 2026-10-16: the original NĐ 98/2021 Điều 40 (conditions for trading class B/C/D devices) does not require ISO 13485. Amendments NĐ 07/2023, 96/2023, 04/2025 and a 2026 BYT circular have not been checked. The K:\ library has no medical-device folder. No ISO quote template exists yet. Waiting for the user: pricing, the usual certification body, and a combined or separate contract. Full note: `docs/ISO 13485 - quy trinh chung nhan (phan phoi TTBYT).md`. **On hold until Haki replies** (user, 2026-10-16).
- **User review of `catalog/Duyet song ngu.xlsx`** (246 rows, 0 approved). Yellow rows have notes:
  - Chức danh "General Director" or "General Manager"
  - Tái đánh giá WHO: VN says 5 years, old ENG said 3 years
  - The GSP template says "kiến thức cơ bản về GMP" (probably should be GSP)
  - GSP dossier fee "21 triệu" needs confirming
  - One VN note cell has two sentences merged
- Confirm replacement wording, if wanted, for the GDP footnote and the GSP "thúc đẩy hồ sơ" section.
- Investment-report templates (sector 8) refer to attachments: a table-of-contents .doc and a cost-norm sheet. These are not attached automatically, and the cost-norm sheet uses 10% VAT and has a junk Sheet3.
- Template defects to fix at the source (user's call): sheet "Mỹ phẩm" inside the ĐĐKSX step-4 file; the ENG step-2 VAT formula `*10%`.

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| openpyxl | `python -c "import openpyxl"` | required | `python -m pip install --user openpyxl` | installed |
| Pillow | `python -c "import PIL"` | required | `python -m pip install --user pillow` | installed |
<!-- launcher:tools:windows:end -->

## Change log
- 2026-10-13: Project created.
- 2026-10-13: Built the quote tool. Added the catalog from the 2025 templates, a 246-string bilingual library (pending review), the request template, the generator with yellow flags and a review sheet, and blocking of legal-risk wording. Tested on VN, VN+EN with seal, and bilingual GDP dated 2027 (VAT flag); the layouts were checked via PDF export.
- 2026-10-13: Added the `[gioi_thieu]` field (customer-specific opening paragraph) and text prices (e.g. "Hỗ trợ"). Fixes: no EN warnings for VN-only quotes; one timeline is printed without a "Mục" prefix; extra notes are merged into the single "* Ghi chú" cell. First real quote: Haki Medica.
- 2026-10-13: Row-height rule (see the group CLAUDE.md), implemented in `make_quote.py` (`apply_row_rule`, `ROW_HEIGHTS`, `CHARS_PER_WIDTH = 1.15`). The line count uses a greedy word wrap over the merged-cell width and was checked against the Haki PDF. The Haki quote was regenerated.
- 2026-10-13: Bold text wraps earlier on screen than in the PDF. Haki row 23 (the bold section title) shows 3 lines in Excel, but the count gave 2. Added `CHARS_PER_WIDTH_BOLD = 1.0` (regular text stays 1.15), so row 23 is now 50. If the user reports more mismatches, recalibrate against what Excel shows on screen, not the PDF.
- 2026-10-13: Company QR (`assets/gmpc_qr.jpg`, copied from `Desktop\GMPc-QR.jpg`; its content was not decoded or verified) sits bottom-left of the signature block on every quote (`qr = false` to omit).
  - The seal is scaled to 300 px and right-aligned under "Tổng giám đốc".
  - `fit_pages()` estimates the print height. If the last page would hold only a sliver (< 0.35 page, after a 1.1 page-break waste factor), the sheet is shrunk to one page fewer, so the seal never splits across pages.
- 2026-10-13: Researched the GDP scope and warehouse requirements for Haki (dental anaesthetics). Saved the note in `docs/` and the regulations in `data/raw/vbpl/`, and added a pointer and key facts to the GMPc group CLAUDE.md.
  - Haki was switched to `dong_dau = true` at the user's request and is now 2 pages. Because the first file was open in Excel, the stamped version was saved as ` v2`.
- 2026-10-13: The user sent Haki ` v2` to the customer. Status recorded under "Quotes made".
- 2026-10-16: Haki follow-up. Drafted the user's short reply to the customer (sign first, details after; no promise to introduce candidates). Researched ISO 13485 certification online and saved it to `docs/ISO 13485 - quy trinh chung nhan (phan phoi TTBYT).md`. Everything is on hold until Haki replies.
