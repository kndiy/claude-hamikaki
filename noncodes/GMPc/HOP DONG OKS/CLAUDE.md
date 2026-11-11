# HOP DONG OKS

## Purpose and goals
Draft contracts and contract paperwork (đề nghị tạm ứng / thanh toán, etc.) issued by **Công ty cổ phần OKS** (OKS Joint Stock Company), filled from the signed contract and OKS templates. GMPc's own contracts live in the sibling project `HOP DONG GMPC`, which uses different templates.

## OKS company details (from contract MERAP-OKS, 2026)
- Name: CÔNG TY CỔ PHẦN OKS / OKS JOINT STOCK COMPANY. Tax code: 0106129081.
- Address: Tầng 6, tòa nhà MD Complex, số 68, phố Nguyễn Cơ Thạch, phường Từ Liêm, TP. Hà Nội.
- Representative: Bà Phạm Hương Thảo, Giám đốc. Tel: 0989 686 580.
- Bank: 1903.601.657.1011, Ngân hàng TMCP Kỹ Thương Việt Nam (Techcombank), CN Đông Đô.

## Tech stack and key dependencies
- Python 3.13 (stdlib; `defusedxml`, `lxml` for the docx skill scripts).
- This machine has no pandoc, LibreOffice, or `zip`. Read .docx with `scripts/dump_docx.py` and rezip with Python `zipfile`. Output can't be rendered to PDF here, so check it in Word.

## Structure and key files
- `templates/`: OKS templates. `De nghi tam ung - mau OKS.docx` is the advance request (it still holds the old 2024 Trung Sơn Alpha content).
- `scripts/dump_docx.py`: prints a .docx's text.
- `khach-hang/<năm>-<số NAS> <KHÁCH>/`: one folder per customer contract, with `NOTES.md`, `raw/` (copies from the NAS), and `output/`.
- `nhan-su/<năm> <Tên NLĐ>/`: OKS employment contracts (HĐLĐ), same layout. Source: `O:\02. Quản lý nhân sự\6. Hợp đồng lao động\`. The NAS HĐLĐ templates have GMPc as employer; swap in the OKS details above when the employee signs with OKS.

## Commands
- Read: `python scripts/dump_docx.py <file.docx>`
- Fill a template: unzip, run `merge_runs.py`, string-replace in `word/document.xml`, rezip with Python, then `validate.py --original`.

## Conventions, constraints, decisions
- OKS documents keep the OKS template formatting (Times New Roman). The GMPc group rules for Calibri Light font and GMPc header details **do not apply** to OKS-issued documents. The confidentiality and NAS read-only rules still apply.
- Use dots for numbers in Vietnamese text and commas in English text.
- Title documents to match the contract wording (for example "ĐỀ NGHỊ TẠM ỨNG" for an advance, not "THANH TOÁN").

## Customers
- `2026-20 MERAP`: consultancy contract (project management and construction supervision) for Meraplion. See its `NOTES.md`.
- `2026-11 MD CARE`: construction contract 1704/HDTC/CELL-DERM-OKS/2026 with CELL-DERM Laboratory Việt Nam. Re-signed version with new prices drafted 2026-11-10. See its `NOTES.md`.

## Employees
- `nhan-su/2026 Vo Duy Cuong`: HĐLĐ chính thức + HĐLĐ-BHXH, 0311/2026/OKS-HĐLĐ (2026-11-05). See its `NOTES.md`.

## Open questions and next steps
- Make blank OKS templates (advance request, payment request, contract) instead of reusing old filled-in files.

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| defusedxml | `python -c "import defusedxml, lxml"` | recommended | `pip install defusedxml lxml` | installed |
<!-- launcher:tools:windows:end -->

## Change log
- 2026-11-10: Added `khach-hang/2026-11 MD CARE`; drafted the re-signed CELL-DERM contract (new price, payments = amounts received). VAT recalculated to the exact 8% at the user's request; the NAS annex spreadsheet still has the old hard-coded VAT. `dump_docx.py` needs `PYTHONIOENCODING=utf-8` in this shell.
- 2026-11-05: Added `nhan-su/` for OKS employment contracts; drafted Võ Duy Cường's two HĐLĐ.
- 2026-10-08: Created from the old `SOAN HOP DONG DNTT` project and split away from GMPc contracts because the templates differ. Restructured into templates/, scripts/, and per-customer folders.
