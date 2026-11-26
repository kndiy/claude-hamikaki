# 20261001 Bao Gia Minh Viet

Created 2026-10-01 by the launcher.

## Purpose and goals
Draft a consulting quotation (bản chào giá) from GMPc to **Công Ty Cổ Phần Quốc Tế Minh Việt** (contact: chị Lan) covering CGMP-ASEAN layout design, a renovation cost estimate and a renovation schedule (section I, 45,000,000 VND before VAT).

## Tech stack and key dependencies
- Excel (.xlsx) deliverable, edited with Python 3 + openpyxl 3.1.5.

## Structure and key files
- `data/raw/` holds source copies. Never edit them. They were copied on 2026-10-01 from `C:\Users\Khiem Nguyen\Desktop\Chao gia\` (local Desktop, not the NAS):
  - `261001 Bao gia Minh Viet.xlsx`: the user's draft and the basis for the output.
  - `Chao gia tu van CGMP-MP Cafune.xlsx`: the newer template (22/09/2026), a reference for addresses, the VAT note and exclusion wording.
- `data/assets/`: images extracted from the Minh Viet raw file. `gmpc_logo.png` is the logo. `gmpc_stamp_signature_TGD.png` is the company seal plus the signature of Tổng giám đốc Trần Thị Ngọc Hà. The seal still reads "H. Hoài Đức", the physical seal's old address.
- `scripts/draft_v3.py` builds the **current** draft: the Minh Viet file is the template, and sections II and III come from the Cafune "Mỹ phẩm" sheet. It also places the logo and the stamp/signature. (The v1 and v2 scripts and outputs are kept for history. Their outputs have **no images**, because Pillow was missing at the time.)
- `output/261001 Bao gia Minh Viet - draft v3.xlsx` is the current draft. Yellow cells need confirmation, and the `Ghi chu soan thao` sheet lists the changes and open items.

- **Contract (2026-10-14):** `data/raw/hopdong/` holds copies from `N:\02. Dự án\1. Hợp đồng khách hàng\Năm 2026\Minh Viet\`: the template `261014 Draft Hợp đồng tư vấn Minh Viet - GMPc.docx` (still the Cafune contract text, Vietnamese only, Calibri Light) and `261001 Bao gia Minh Viet.pdf` (the quote as sent). `scripts/hopdong_v1.py` built draft v1. The user then edited v1 by hand in Word, and **`output/261014 Hợp đồng tư vấn Minh Viet - GMPc - draft v2.docx` is now the current contract**: the user's v1 plus one fix by Claude. From here on, edit the .docx directly. Don't regenerate it from the script, which now refuses to overwrite v1.

## Commands (build / run / test / lint)
- Contract page count: open it in Word (COM `ComputeStatistics(2)`) and keep the "bao gồm NN trang" sentence in Điều 11 in sync.
- `python scripts/draft_v3.py` (run from the project root) regenerates the draft.
- Pitfall: without Pillow, openpyxl **silently drops images** on save. The v3 script imports PIL so that it fails loudly instead. It also clears the template's own images and re-anchors them, because image anchors don't move when rows are inserted.

## Conventions, constraints, decisions
- The Minh Viet file is the template. Content from section II onward comes from the Cafune quote (user's decision, 2026-10-01).
- Bilingual VN/EN on every line. English uses the terminology of the Cafune ENG sheet (Investor, MEP, Guide to…, self-inspection, budget estimating).
- Execution time and payment terms follow Cafune's multi-stage structure (Part I keeps Minh Viet's "3 weeks"; payments 30/20/20/20/remainder). These are pending confirmation.
- Font and GMPc addresses follow the GMPc group CLAUDE.md.
- The VAT 8% note cites Decree 174/2025/NĐ-CP, valid until 31/12/2026. Recheck it for any quote issued after that date.
- The price is kept as in the user's draft. Claude doesn't set prices.
- Customer: Công Ty Cổ Phần Quốc Tế Minh Việt / INTERNATIONAL MINH VIET JOINT STOCK COMPANY, Khu Công nghiệp Tân Phú Trung, ấp Trạm Bơm, Xã Củ Chi, TP Hồ Chí Minh, Việt Nam (confirmed by the user on 2026-10-01). Phone and email are skipped, and rows 12-13 are hidden.
- Before sending: remove the yellow fills and delete the review sheet. The file carries the seal and signature, so send it only after prices and terms are final, preferably as a PDF.

## Contract decisions (user, 2026-10-14)
- Scope is Part I only: 45,000,000 + 8% VAT 3,600,000 = 48,600,000 VND. Vietnamese only, same format as the template.
- Bên A: MST 0300721002, represented by (Ông) Yansen Nagakarsa Adi, Tổng giám đốc.
- Timing is 21 working days from signing and receipt of the advance.
- The user's own edits to v1: payment is 60% advance and 40% within 05 working days of handing over all 2.2 deliverables, then acceptance and liquidation, each with a list of payment documents. The VND amounts per installment were removed. The price column and the totals were removed from the 2.1 table (the fee appears only in 3.1). The Part I liability clause in 5.2 and the document-ownership part of Điều 9 were removed. The term is 12 months from signing. The yellow highlights were cleared. The document is 8 pages.
- v2 fix (Claude): installment 2 now reads "05 ngày làm việc", matching installment 1.

## Open questions and next steps
**Contract v2, still open (raised with the user 2026-10-14):** the Part I title still says "Quy hoạch tổng thể" with no matching line item. Clause 7.4 still has uncapped damages for Bên B (from Cafune), and since the 5.2 cap was removed, the contract has no overall liability cap. The signature block has no names or titles. Contract number and date: the user kept 1410/... and 14/10/2026.

**Resume here (session paused 2026-10-01):** draft v3 is built. The user still has to answer the items below, which match the yellow cells in v3. Also unanswered: whether to move the Pillow tool row up to the GMPc group CLAUDE.md.
- Prices for sections II (150M) and III (260M) are copied unchanged from Cafune (a ~1,100 m2 wet-cosmetics plant). They need repricing for Minh Viet.
- Payment terms: confirm the 5 installments from Cafune (the original Minh Viet draft had 2 installments, with a VN 50% / EN 30% conflict).
- Section I.2 (renovation cost estimate) overlaps with II.3 (renovation budget estimate). Define each clearly.
- Section I scope: new layout design or feedback on the renovation? The header also mentions master planning, which no line item covers.
- Optional: add an "information provided" condition to the execution time, and exclusions to the notes.

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| openpyxl | `python -c "import openpyxl"` | required | `python -m pip install --user openpyxl` | installed |
| Pillow | `python -c "import PIL"` | required | `python -m pip install --user pillow` | installed |
<!-- launcher:tools:windows:end -->

## Change log

Older entries: `zClaudeDocs/CLAUDE_CHANGELOG.md`.

- 2026-10-01: Draft v3: customer address and English name, phone/email rows hidden, logo and stamp/signature restored (they were lost in v1/v2 because Pillow was missing), Pillow added to tools.
- 2026-10-14: Contract draft v1 (Part I only) built from the Cafune-based contract template on N:.
- 2026-10-14: The user hand-edited v1. Claude reviewed it and saved v2 with the working-days fix.
