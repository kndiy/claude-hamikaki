# BAO GIA SHICONS

Created 2026-11-04 by the launcher.

## Purpose and goals
Prepare quotations (báo giá / bảng chào giá) for SHICONS E&C services, mainly warehouse temperature/humidity mapping (thẩm định kho GSP/GDP). Deliverables: .xlsx plus a .pdf for sending to the customer.

## Tech stack and key dependencies
- Microsoft Excel (COM automation from PowerShell) to edit and export PDF; keeps logo/format of the template.
- Python 3.13 + openpyxl (only for reading/inspecting workbooks). pywin32 is NOT installed.

## Structure and key files
- Source material (read-only, outside this root): `C:\Users\Khiem Nguyen\Documents\02 WORK MR LUAN\`
  - Template: `260626 BAO GIÁ SHICOS - VINPHA.xlsx`, sheet `TỔNG HỢP` (print area A1:F26).
  - Reference competitor quote: `Thanh an Báo giá thẩm định kho.pdf` (27.6M before VAT, 32 sensors, 8 days, raw data only).
  - Past mapping projects/reports: `THAM DINH KHO\`, `2026 Vinpha\`.
- Output: `YYMMDD BAO GIÁ SHICONS - <KHÁCH HÀNG>.xlsx` / `.pdf` in this folder.

## Commands
- Build: a PowerShell script that copies the template, deletes all sheets except `TỔNG HỢP`, sets cells (A9 khách hàng, A10 dự án, A11 địa điểm, B13/E13 hạng mục + giá, B14 sensors/ngày, A19 bằng chữ, D22 ngày), saves, then `ExportAsFixedFormat(0, pdf)`.
- PS 5.1 needs the .ps1 saved as UTF-8 **with BOM** for Vietnamese text.
- After editing B13, check row height (row 13 set to 190 pt) so text is not clipped.

## Conventions, constraints, decisions
- Totals: E16 `=ROUND(SUM(E13:E14),-3)`, VAT 8% in E17, E18 after tax. "Bằng chữ" (A19) is typed by hand: update it whenever the price changes.
- Payment terms from template: 50% on signing, 50% within 5 days after handover of documents.
- Kho thường conditions (confirmed by user): 15–30 °C, < 75 %RH, logging every 15 min.
- VAT 8% (confirmed). Free 1-day pre-test (item 1.2) is standard: kept to check the warehouse meets requirements before the official run.
- Don't add extra terms (e.g. retest fee) unless asked.

## Open questions and next steps
- None for the Ms. Nhung quote (company name intentionally left blank, location TP HCM).

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| Microsoft Excel | `powershell -NoProfile -Command "New-Object -ComObject Excel.Application \| Out-Null"` | required | `manual: install Microsoft 365 Excel` | installed |
<!-- launcher:tools:windows:end -->

## Change log

Older entries: `zClaudeDocs/CLAUDE_CHANGELOG.md`.

- 2026-11-10: Drafted contract `261110 HỢP ĐỒNG SHICONS - JK` (.docx/.pdf) with `build_hd_jk.ps1`, from the 2023 template `02 WORK MR LUAN\Bao gia\Hợp đồng bảo trì hệ thống HVAC và thẩm đinh kho JK.doc`. Addresses updated per user: JK = 6 đường số 8, KDC Ven Sông, P. Tân Hưng, TP HCM; SHICONS = 78H, Đường số 9, KP5, P. Hiệp Bình, TP HCM (also English header). Table mirrors the FN quote. Contract no., ngày ký and thời gian thực hiện left as '.....'. Removed the irrelevant QĐ 5522/QĐ-UBND legal basis; fixed typos (thẩ định, 04 (năm) bản, Điện Thoai).
- Word COM notes: iterating `StoryRanges/NextStoryRange` hangs on this doc (use Content + section headers/footers); `SaveAs2` straight to the Vietnamese-named path in this folder hangs, so save to `%TEMP%` then copy. The table has vertical merges, so insert rows with `Selection.InsertRowsBelow`, not `Rows.Add`.
- 2026-11-10: User finalized the JK contract by hand in Word (changes not reviewed by Claude). Do NOT rerun `build_hd_jk.ps1`: it would overwrite the hand edits. The open items below may already be handled in those edits.

## Open items (contract JK)
- Confirm still valid from 2023: JK phone/fax, bank account ACB 213786009, representative Bà Nguyễn Kim Cương (Tổng Giám đốc); SHICONS account ACB 25107986 and representative Bà Dương Thị Kiều Huyên.
- Clauses carried over from a calibration template that don't fit mapping (left unchanged, awaiting user decision): Điều 3.2a 'Giấy chứng nhận kiểm định…', Điều 4.4 'kết quả kiểm định phương tiện đo của Bên A không đạt', Điều 5.6 'Giấy chứng nhận kiểm'. Footer code `SHICONS-VIWA-CT1-03`.
