# ON THI

Created 2026-10-26 by the launcher.

## Purpose and goals
Help the user (Dược student, K19) review for exams. Current focus: two make-up ("Bù") courses:
- **03. Kỹ năng giao tiếp – TT-GDSK** (KNGT)
- **07. Giải phẫu sinh lý** (GPSL)

Deliverables are study aids (summaries, quizzes, flashcards), not software.

## Tech stack and key dependencies
- Python 3.13 + `pymupdf` (PDF text). `.docx`/`.pptx` read with stdlib zip/XML. Old `.doc` read via Word COM (PowerShell); `antiword` returned nothing for these files.

## Structure and key files
- Source material (read-only, do not edit): `C:\Users\Khiem Nguyen\Documents\05 DUOC\2026 DUOC\` → `03. KY NANG GIAO TIEP (Bù)`, `07. GIAI PHAU SINH LY (Bù)`.
- `extracted/03_KNGT/`, `extracted/07_GPSL/`: plain-text dumps of every doc/pdf/pptx (attendance lists and the big reference textbooks in `Tham khao/` skipped).
  - `extracted/03_KNGT/DAP_AN_200_cau_KNGT_(tu_highlight).txt`: answer key for the 200-question KNGT bank, taken from the highlighted options in the source docx.
  - `extracted/07_GPSL/On thi__200_cau_trac_nghiem_Giai_phau_Sinh_ly.docx.txt`: 200 MCQ; the correct option is the one followed by an explanation sentence.
- `scripts/extract.py`: `python -I scripts/extract.py <src_dir> <out_dir>`.
- `scripts/merge_decuong.py` → `extracted/03_KNGT/decuong_parsed.json` (parses the 2 sets in the KNGT đề cương); `scripts/build_decuong_docx.py` → `output/KNGT_De_cuong_tong_hop_khong_trung.docx` (dedup by normalized question text, fixes obvious typos, table STT | Câu hỏi | Đáp án; user wants questions and answers only, no source tags).
- `scripts/parse_kngt200.py` → `extracted/03_KNGT/kngt200_parsed.json` (200-câu bank: bài, question, options, answer letter from highlight, full answer text); `scripts/build_kngt200_docx.py` → `output/KNGT_200_cau_theo_bai.docx`. Layout per user request (2026-10-26): one continuous table with columns Câu hỏi | Đáp án, no STT column, no page breaks; each bài is a shaded full-width row inside the table. 9 pages.
- `notes/KNGT_on_tap.md`: source of the KNGT study guide (edit this, then rebuild). `scripts/md_to_docx.py <md> <docx>` converts it (supports #/## headings, nested bullets / numbered lists, `> ` note boxes, `|` tables, `**bold**`) → `output/KNGT_Tai_lieu_on_thi.docx`.
- `notes/GPSL_on_tap.md` → `output/GPSL_Tai_lieu_on_thi.docx` (13 pages). Build with `python -I scripts/md_to_docx.py notes/GPSL_on_tap.md output/GPSL_Tai_lieu_on_thi.docx --continuous`; `--continuous` means no page breaks, which the user prefers for printing.
- `scripts/parse_gpsl200.py` → `extracted/07_GPSL/gpsl200_parsed.json`. The correct option is the one followed by an explanation sentence; the answer text excludes that explanation. `scripts/build_kngt200_docx.py <json> <docx> [title] [subtitle]` builds a 2-column table (Câu hỏi | Đáp án); a bài header row is added only when items have `bai`. Output: `output/GPSL_200_cau_trac_nghiem.docx`. A copy was saved at the user's request to the local course folder `C:\Users\Khiem Nguyen\Documents\05 DUOC/2026 DUOC/07. GIAI PHAU SINH LY (Bù)/On thi/GPSL_200_cau_hoi_dap_an.docx` (new file; nothing overwritten).
- `output/`: finished study files for the user.
- **Do not regenerate `output/KNGT_200_cau_theo_bai.docx` without asking.** The user has re-saved it in Word (the XML is now Word's), so they may have edited it.

## Commands (build / run / test / lint)
- Re-extract: `python -I scripts/extract.py "<course folder>" extracted/<name>`

## Conventions, constraints, decisions
- Answer in Vietnamese.
- **Source authority (user decision 2026-10-26):** giáo trình + GV slides are the authority. The KNGT 200-câu bank was written by students and contains errors, so use it only for reference and list where it disagrees. When the đề cương or slides ask for steps or characteristics, list every item with a short explanation, not just the count.
- Flag (don't silently fix) content that looks wrong. When something is explained from general knowledge rather than the course materials (e.g. ABCD, AIDMA), label it as such.
- Known issues spotted in GPSL lecture notes (`Giai phau 1/*Tab*.docx`, look like auto-transcripts): "cầu não nối hai bán cầu đại não" (wrong), đám rối thắt lưng "L2–L4" (usually L1–L4), đám rối cùng "L4–S5" (usually L4–S4), "dây TK đùi chạy ra ở nếp lằn mông" (should be TK tọa). `7-5` and `7-7` Tab notes are identical.

## Open questions and next steps
- Exam dates and format (trắc nghiệm only, or also tự luận — the KNGT "CÂU HỎI ÔN KT CUỐI MÔN" list is short-answer style)?
- GPSL đề cương: resolved, the user gave the topic list (see change log 2026-10-26).
- Ask the lecturer: which KNGT steps the essay expects (thay đổi hành vi 4 vs 5; lập kế hoạch 6 vs 8); whether the "thông điệp tốt" answer is 7 or 8; whether ABCD/AIDMA can be asked beyond A/D.
- **Status at end of session 2026-10-26 (user paused, "lưu lại"):** KNGT is done (đề cương 75 câu, 200 câu table, study guide). GPSL has the study guide and the 200-câu table, with answers checked.
- Next, pending the user: offered to check KNGT 200-câu answers question by question against the giáo trình (the user hasn't answered yet). Other ideas: quiz the user on either bank; flashcards.

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| PyMuPDF | `python -c "import pymupdf"` | recommended | `python -m pip install --user pymupdf` | - |
<!-- launcher:tools:windows:end -->

## Change log

Older entries: `zClaudeDocs/CLAUDE_CHANGELOG.md`.

- Conflicts found between sources (flagged in the guide, not resolved): thay đổi hành vi has 4 giai đoạn in the bank vs 5 bước in the textbook; lập kế hoạch has 6 bước in the bank vs 8 bước in the textbook and slides; the "đặc điểm thông điệp tốt" question says 7 but Sơ đồ 2.4 lists 8. Open: ask the lecturer which version the essay expects.
- 2026-10-26: Built the GPSL 200-câu Q&A table (one continuous table, no STT, 9 pages). All 200 questions parsed with exactly one answer each, and none are duplicates. Answers were reviewed against the slides and study guide with no conflicts. The phân thùy question explicitly uses the "gộp S1+S2, S7+S8 → 8" convention, which is consistent with the guide's note.
- 2026-10-26: At the user's request, did a full answer check of the GPSL 200-câu bank, reviewing every question with all 4 options. No wrong keys found, and no question has a second defensible option. Minor notes only: Q3 answer 8 relies on the stated convention (slide table gives 9 for phổi trái); Q151 "L2–L4, nổi bật L4" is close to the usual L3–L4; Q73 "2–3 cm" vs slide "3 cm"; Q111 "L3–L4 hoặc L4–L5" vs slide L4–L5. No file changes were needed.
