# CLAUDE.md change log archive

Older change-log entries moved out of `../CLAUDE.md` on 2026-11-26, verbatim and in their original order. The newest entries stay in `CLAUDE.md`; when an entry leaves it, append it here.

- 2026-10-13: Project created.
- 2026-10-13: Built the quote tool. Added the catalog from the 2025 templates, a 246-string bilingual library (pending review), the request template, the generator with yellow flags and a review sheet, and blocking of legal-risk wording. Tested on VN, VN+EN with seal, and bilingual GDP dated 2027 (VAT flag); the layouts were checked via PDF export.
- 2026-10-13: Added the `[gioi_thieu]` field (customer-specific opening paragraph) and text prices (e.g. "Hỗ trợ"). Fixes: no EN warnings for VN-only quotes; one timeline is printed without a "Mục" prefix; extra notes are merged into the single "* Ghi chú" cell. First real quote: Haki Medica.
- 2026-10-13: Row-height rule (see the group CLAUDE.md), implemented in `make_quote.py` (`apply_row_rule`, `ROW_HEIGHTS`, `CHARS_PER_WIDTH = 1.15`). The line count uses a greedy word wrap over the merged-cell width and was checked against the Haki PDF. The Haki quote was regenerated.
- 2026-10-13: Bold text wraps earlier on screen than in the PDF. Haki row 23 (the bold section title) shows 3 lines in Excel, but the count gave 2. Added `CHARS_PER_WIDTH_BOLD = 1.0` (regular text stays 1.15), so row 23 is now 50. If the user reports more mismatches, recalibrate against what Excel shows on screen, not the PDF.
- 2026-10-13: Company QR (`assets/gmpc_qr.jpg`, copied from `Desktop\GMPc-QR.jpg`; its content was not decoded or verified) sits bottom-left of the signature block on every quote (`qr = false` to omit).
  - The seal is scaled to 300 px and right-aligned under "Tổng giám đốc".
  - `fit_pages()` estimates the print height. If the last page would hold only a sliver (< 0.35 page, after a 1.1 page-break waste factor), the sheet is shrunk to one page fewer, so the seal never splits across pages.
