# dang-vien

Created 2026-11-18 (by hand, same structure as the launcher: `.claude-project` marker + this file).

## Purpose and goals
Material for the user's Party-member (đảng viên) online study contests and exams: the source documents, the screenshots of the
exam questions, and the answer sheets worked out from them. Deliverables are answer files (.md / .docx), not software.

## Tech stack and key dependencies
- None to install. `pdftotext` (poppler) for the PDFs that have a text layer; scanned PDFs have to be read page by page.
- `python-docx` (system Python) builds the .docx answer files.

## Structure and key files
- `261118_1300 TW4 Nghi quyet Hoi nghi lan 4 BCH TW khoa XIV/` (was `TW 4`): Nghị quyết Hội nghị lần thứ tư, BCH TW Đảng khóa XIV. The PDFs are the source documents (NQ 28, 29, 30, 31, KL 105, 106, QĐ 224; all scanned). `Quizz/` holds the exam screenshots, `Dap an theo Cau.md` / `.docx` (answers sorted by the printed "Câu" number, 26 questions, no Câu 26 and 27) and `Cau 28 - essay.md` (the essay, about 100 words, a draft in the user's name).
- `261118_1305 Thi tim hieu chuyen doi so quoc gia/` (was `Tim Hieu So`): Cuộc thi trực tuyến tìm hiểu về chuyển đổi số quốc gia. PDFs (NQ 57, Luật KH,CN&ĐMST 93/2025, QĐ 757 skills framework, contest rules QĐ 187; most have a text layer, `Tailieu1.pdf` is a scan of NQ 57). `Quizz/` holds the screenshots and the answer files (`Dap an theo Cau.md` = the 11 questions of the first set; `Dap an.md` = the same with sources).

## Commands (build / run / test / lint)
- Text of a PDF: `pdftotext -layout file.pdf out.txt`.
- Answer .docx is generated with python-docx from `Dap an theo Cau.md` (no script is kept yet).

## Conventions, constraints and decisions
- Folder names: `YYMMDD_HHmm <short content>`, the date and local time (UTC+7) of the exam attempt, ASCII only.
- Answers are sorted by the printed "Câu" number; the contest app shuffles the order on screen. The Thi tim hieu chuyen doi so set has no printed "Câu", so its own 1 to 11 is used.
- The documents may not contain every answer (decrees, laws and decisions outside the folder were looked up on the web). Mark answers as firm (in the documents) or less sure (web, elimination, memory).
- The exam app presents different question sets per attempt (a second Tim Hieu So set came with questions about NQ 19, semiconductor strategy, AI Law scope, a UN speech); its answers are not saved in files.

## Open questions and next steps
- The second Thi tim hieu chuyen doi so set was answered in chat only (questions 1 to 6 and 10, 11); save if wanted.
- The prediction question (number of people answering everything correctly) was guessed at 1,400,000 by the user.
