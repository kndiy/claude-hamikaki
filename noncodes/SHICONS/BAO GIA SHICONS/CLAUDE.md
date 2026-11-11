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
- 2026-11-04: Project created.
- 2026-11-04: Created quotation `261104 BAO GIÁ SHICONS - MS NHUNG` (thẩm định kho GDP/GSP nguyên liệu làm thuốc, 76.8 m² + 27.8 m², cao 2.6 m, 24 LogTag, 7 ngày, 10,000,000 VNĐ trước VAT / 10,800,000 sau VAT 8%) from the Vinpha template.
- 2026-11-04: Ms. Nhung quote finalized: A9 company left blank, A11 = TP Hồ Chí Minh; other content unchanged.
- 2026-11-09: Revised Ms. Nhung quote → `261109 BAO GIÁ SHICONS - MS NHUNG` (old 261104 file kept). Added rows 16–17: (2) thẩm định xe vận chuyển 1 xe 2,000,000; (3) thẩm định khu lấy mẫu 4,000,000 (wording from old JK/VIWA contract 2023, where it was 3,500,000; user set 4M). Total 16,000,000 / 17,280,000 after VAT. Totals now E18:E20, bằng chữ A21, date D24, print area A1:F28, page setup fit to 1 page. A10 shortened so it fits the width.
- 2026-11-09: Item 2 (xe) details confirmed by user: 5 sensors, 15–30 °C, 1.5 days; row 16 height 132 pt. Price/totals unchanged.
- 2026-11-09: User asked to redo the Ms. Nhung quote with the content of `02 WORK MR LUAN\Bao gia\Q.AS-PM.PA3.xlsx` (VIWA 2023 quote, attn chị Nhung) "y nguyên", in the SHICONS form. Output `261109 BAO GIÁ SHICONS - MS NHUNG V2` (.xlsx/.pdf); the earlier 261109 version (kho/xe/khu lấy mẫu = 16M) is kept. V2 layout: new column E "ĐƠN GIÁ" (Thành tiền = F, Ghi chú = G), items rows 13–22 with group rows bold, totals F23:F25 = 21,900,000 / 1,752,000 / 23,652,000, bằng chữ A26, terms B27:B29 from PA3 (10 ngày, thanh toán 100% sau hoàn thành, bảo hành 12 tháng), date D31, print area A1:G35, fit to 1 page, logo moved left to fit. Vehicle item and free pre-test dropped (not in PA3).
- 2026-11-09: User edited V2 by hand: customer now does only 1 kho (nguyên liệu, tầng trệt; tầng trên kho thành phẩm dropped) + vệ sinh 2 FCU tầng trệt. V2 items now: (1) kho 76.8+27.8 m², 24 sensors, 7 ngày, 10M; (2) xe 1 xe, 5 sensors, 2M; (3) AHU-LM khu lấy mẫu 4M; (4) vệ sinh máy lạnh FCU 2 bộ × 1M. Total 18,000,000 / VAT 1,440,000 / 19,440,000; A9 = CÔNG TY CỔ PHẦN DƯỢC PHẨM JK - Ms. Nhung. Claude fixed duplicate STT (A20=3, A24=4) and re-exported the PDF (1 page). Totals F26:F28, bằng chữ A29, date D34, print area A1:G38.
- 2026-11-09: User found row spacing too large in the PDF. Rows 13–25 now AutoFit + 6 pt padding (minimum 20 pt; detail rows were 24), totals rows 26–28 set to 22 pt (were 27.6), rows 29–34 set to 18 pt (were 20.25). PDF re-exported.
- 2026-11-10: Final quote sent to customer = `02 WORK MR LUAN\Bao gia\261109 BG Tham dinh Kho GSP - Ms. Nhung FN.pdf` (same 4 items, 19,440,000 after VAT), so xe + AHU-LM are in scope.
- 2026-11-10: Drafted contract `261110 HỢP ĐỒNG SHICONS - JK` (.docx/.pdf) with `build_hd_jk.ps1`, from the 2023 template `02 WORK MR LUAN\Bao gia\Hợp đồng bảo trì hệ thống HVAC và thẩm đinh kho JK.doc`. Addresses updated per user: JK = 6 đường số 8, KDC Ven Sông, P. Tân Hưng, TP HCM; SHICONS = 78H, Đường số 9, KP5, P. Hiệp Bình, TP HCM (also English header). Table mirrors the FN quote. Contract no., ngày ký and thời gian thực hiện left as '.....'. Removed the irrelevant QĐ 5522/QĐ-UBND legal basis; fixed typos (thẩ định, 04 (năm) bản, Điện Thoai).
- Word COM notes: iterating `StoryRanges/NextStoryRange` hangs on this doc (use Content + section headers/footers); `SaveAs2` straight to the Vietnamese-named path in this folder hangs, so save to `%TEMP%` then copy. The table has vertical merges, so insert rows with `Selection.InsertRowsBelow`, not `Rows.Add`.

- 2026-11-10: User finalized the JK contract by hand in Word (changes not reviewed by Claude). Do NOT rerun `build_hd_jk.ps1`: it would overwrite the hand edits. The open items below may already be handled in those edits.

## Open items (contract JK)
- Confirm still valid from 2023: JK phone/fax, bank account ACB 213786009, representative Bà Nguyễn Kim Cương (Tổng Giám đốc); SHICONS account ACB 25107986 and representative Bà Dương Thị Kiều Huyên.
- Clauses carried over from a calibration template that don't fit mapping (left unchanged, awaiting user decision): Điều 3.2a 'Giấy chứng nhận kiểm định…', Điều 4.4 'kết quả kiểm định phương tiện đo của Bên A không đạt', Điều 5.6 'Giấy chứng nhận kiểm'. Footer code `SHICONS-VIWA-CT1-03`.
