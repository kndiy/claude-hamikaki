# Thuc tap tot nghiep

Created 2026-11-10 by the launcher.

## Purpose and goals
Support paperwork for the pharmacy graduation internship (Thực tế ngành + Thực tập tốt nghiệp, class K19D33.1, Trường CĐ Công nghệ Y-Dược Việt Nam) — mainly reviewing training contracts with internship sites.

## Tech stack and key dependencies
Word documents only. Text extraction uses a throwaway Python script (zipfile + xml, no extra packages).

## Structure and key files
- `output\` — edited copies of the contracts (originals stay untouched).
Source documents live outside this root (user's personal folder, read-only for Claude unless asked):
- `C:\Users\Khiem Nguyen\Documents\05 DUOC\2026 DUOC\00 THUC TAP TOT NGHIEP\Nhà thuốc - Doanh nghiệp\`
  - `HĐ Nguyen tac - Bidopharma .docx` — framework contract (Bên A = Bidopharma USA, Bên B = school)
  - `HĐ Chi tiet - Bidopharma.docx` — detailed contract + student list appendix (Bên A = school, Bên B = Bidopharma)

## Commands (build / run / test / lint)
None.

## Conventions, constraints, decisions
- Contract reviews check logic/consistency only. Dates and contract numbers are left blank and filled in by hand by the teacher.
- NEVER edit the source .docx files. Write every edited version as a new file in `output\` inside this project (user rule, 2026-11-10).

## Open questions and next steps
- Which school phone number is correct: 0236.6295333 (detailed contract) or 0236 6395333 (framework contract)?
- Waiting on user: apply the review fixes to the .docx files or not.
- Detailed contract: header says "Năm học 2025 - 2026" but Điều 2.3 / Điều 3 say 2026-2027 and dates are Nov 2026–Jan 2027 → fix one way.
- The edited file has not been opened in Word yet (no Word/LibreOffice on this machine to render it) → user to check the layout.

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
<!-- launcher:tools:windows:end -->

## Change log

Older entries: `zClaudeDocs/CLAUDE_CHANGELOG.md`.

- 2026-11-10: User found Phụ lục 02 too detailed. Rebuilt as v2: PL02 = only a lecturer table (ThS. Đào Thị Kim Thắng, 0702.511.225, giảng viên phụ trách). No plan table, site info or signature block. Reference sentence before Điều 5 shortened to match. Saved as `output\HĐ Chi tiet - Bidopharma (bo sung PL02 v2).docx`.
- 2026-11-10: PL02 in v2 restyled to match PL01: uppercase column headers, size 13pt, a closing line "Danh sách có 01 giảng viên ./.". No signature (user: the contract already covers everything else). Overwrote v2. It had a 21:14 timestamp, so the user may have saved it in Word → asked them to check.
- 2026-11-10: Deleted the old detailed version `output\HĐ Chi tiet - Bidopharma (bo sung PL02).docx` (user said ok). v2 is the only output.
