# Quy Trinh Chung GMPc

## Purpose
Draft the company-wide rules and procedures of Công ty Cổ phần GMPc Việt Nam (not tied to one department). The first deliverable is the **Nội quy chung** (general company rules), built from content the user supplies and laid out on the company document format. Deliverables are Word documents.

## Stack
- Final deliverables are `.docx`, in Calibri (GMPc group rule for internal documents).
- Python (root Tools table) for helper scripts. Reuse the builders from the sibling project `..\20260927 Quy Trinh MDC\scripts\` (`form_builder.py` for the company form frame, `sop_builder.py` for the BM01 SOP template, `extract_office_text.py` for reading Office files) instead of rewriting them.

## Structure
- `source/`: input files placed by the user (content drafts, old versions, templates). Keep them unchanged. Copies from the NAS also go here, with the exact source path recorded below.
- `docs/`: Markdown extracts and notes made from `source/`.
- `output/`: generated `.docx` deliverables, each with a `.sha256` build stamp next to it.
- `templates/`: copy of the BM01 SOP template. `assets/`: images taken from source files (`HD-01_So_do_khoi.png`).
- `scripts/qd01.py`: QĐ-01. It is self-contained and adds the Quyết định page and the TOC.
- `scripts/sop_builder.py` plus `hd01.py` / `hd02.py` / `hd03.py`: the HD documents in BM01 layout (one module per document, holding its content as data). The builder is adapted from the MDC project's builder. Differences: the subtitle comes from `KIND`, the footer is in Calibri, the Sơ đồ khối can be an existing PNG (`FLOW_IMAGE`), inline markup `**bold**` / `__italic__` is supported, and the overwrite guard uses the SHA stamp.

## Commands
```
PYTHONIOENCODING=utf-8 python scripts/qd01.py      # builds output/GMPc-QĐ-01 Noi quy lao dong.docx (needs Word for the TOC)
PYTHONIOENCODING=utf-8 python scripts/hd01.py      # or hd02.py / hd03.py
```
`qd01.py` holds the content as data and builds it on `templates/GMPc-QT-01.BM01_Biểu mẫu trắng SOP.docx`, which is a byte-identical copy of the BM01 template in `..\20260927 Quy Trinh MDC\MDC\SOP\`. It writes `<output>.sha256` after each build. It refuses to overwrite the output if the file has changed since then (for example, edited in Word). Copy those edits into the script first, then use `--force`. Close the file in Word before building.

## Conventions and constraints
- Internal company documents. Don't publish or share them outside the company.
- Document codes follow `GMPc-<type>-<nn>`. The code for Nội quy chung is not decided yet.

## Source files (placed by the user, 2026-10-16)
- `source/Nội Quy lao động.docx`: draft of the labour rules, catalogue code **GMPc-QĐ-01** (QĐ = quy định). It has a decision page, a table of contents and 5 chapters. The header has no code. There's no sign-off or history table.
- `source/GMPc-HD-01 …GWS….docx`, `source/GMPc-HD-02 …Base….docx`: guides already laid out on the BM01 SOP frame (HD = hướng dẫn). HD-02's header has no code, and both headers say "QUY TRÌNH THAO TÁC CHUẨN (SOP)".
- `source/Cac QT chung.png`: excerpt of the company-wide catalogue: QĐ-01 Nội quy lao động, QĐ-02 Chính sách lương thưởng & phúc lợi, HD-01…04 (GWS, Base, Claude AI & NotebookLM, Quy ước hỏi đáp & Tri thức).
- All 3 docx use the same header logo as the BM01 template (visually identical).
- Review of the Nội quy draft: `docs/2026-10-16_Nhan-xet_Noi-quy-lao-dong.md`. It covers format fixes and labour-law risks (no annual leave, cash fines, loss of entitlements on resignation, missing mandatory chapters, registration). These were not checked against the law text.

## Open questions and next steps
- **Status 2026-10-16: paused.** The user's role here is only to compile the documents and make their format consistent. Content decisions belong to each document's drafter (HCNS for QĐ-01, IT for HD-01…03). Don't change content unless asked. The user will come back when needed.
- Proposed but not confirmed: move the `.sha256` build stamps from `output/` to a separate folder, so that `output/` holds only the Word files.
- The legal issues in the review need HR's decision before the content is finalised. They are kept verbatim in `qd01.py` at the user's request.
- HD-01…03: rebuilt with option 1 (layout only). The content points in the review are still open and wait for the drafter (Lê Quang Huy, IT). HD-02 and HD-03 keep an empty Sơ đồ khối section, as in the source. Not uploaded yet: QĐ-02 and HD-04 (HD-04 may overlap HD-03 and Base Square).
- QĐ-01: the Quyết định number and the date are blank. The signatory "BAN GIÁM ĐỐC" is kept from the draft.
- HD-01/HD-02 could get the same treatment (HD-02 header code, "SOP" subtitle → "HƯỚNG DẪN"). Not asked for yet.

## Change log
- 2026-10-16: The user chose **option 1** for HD-01…03: rebuild on the QĐ-01 frame without changing content. Outputs are in `output/GMPc-HD-0x …docx` (4 pages each, checked as PDF). The source files were not touched.
  - Layout: header on every page, subtitle "HƯỚNG DẪN", codes HD-01/02/03, effective date and history date blank (`__/__/____`) until signed, steps in title-row plus content-row style, Calibri footer. HD-01 keeps its own flowchart image.
  - Wording fixes only: "Quy trình này" → "Hướng dẫn này"; IT & GWS / Nhân viên IT / IT Admin / nhóm IT → Bộ phận IT; Hành chính nhân sự / Bộ phận HCNS → Phòng HCNS; phom mẫu → biểu mẫu; Wework → WeWork; stray "/ /"; "có thế mạnh mạnh mẽ" → "có thế mạnh"; "trí tuệ nhân tạo lớn" → "ngôn ngữ lớn"; "xử lý ngôn ngữ tự nhiên logic" → "…tự nhiên, lập luận logic"; click/reset/copy/link → bấm/đặt lại mật khẩu/sao chép/đường liên kết. The empty definition rows were dropped.
  - Not changed (review points): account deletion, SLA figure, "Xin nghỉ phép", the HD-03 security items, and "Tổng giám đốc" vs QĐ-01's "CEO".
- 2026-10-16: Reviewed HD-01 (re-saved by the user at 10:47), HD-02 and the new upload HD-03 (read-only; each is 4 pages). Notes are in `docs/2026-10-16_Nhan-xet_HD-01-03.md`. Main points:
  - The header shows on page 1 only.
  - The subtitle still says "SOP".
  - HD-02 and HD-03 have no code.
  - The effective dates are filled in although the documents aren't signed.
  - HD-01: the account deletion step doesn't transfer data first. (The gmpc.vn flag was withdrawn: the user confirmed on 2026-10-16 that **gmpc.vn** is the official email domain.)
  - HD-02: SLA has no figure, and the "Xin nghỉ phép" example contradicts QĐ-01 Điều 3.
  - HD-03: customer data is missing from the list of things banned from AI; turning off data training is optional; it uses shared licence accounts.
  - QĐ-01: the user had filled the sign-off in Word (drafter Chuyên viên HCNS / Trần Thị Thanh Thảo, reviewer Trưởng phòng CS / Ngô Thị Hồng Thắm, approver CEO / Đào Xuân Hưởng). That is now synced into `qd01.py` (`SIGN`). Rebuilt with `--force` (a backup of the user's file is in the session scratchpad). The text is identical to the user's file, and the footer is now Calibri.
- 2026-10-16: Built **GMPc-QĐ-01 Nội quy lao động** (8 pages, checked as PDF). User decisions: Calibri for internal documents (Nunito for customers; recorded in the group CLAUDE.md), the BM01 cross logo is the new logo, and the BM01 frame is used for QĐ documents.
  - Layout: header with subtitle "QUY ĐỊNH", code GMPc-QĐ-01 and title "NỘI QUY LAO ĐỘNG"; the sign-off table; the Quyết định ban hành on page 1, with a standard two-column heading and a blank "Số: ……/QĐ-GMPc"; an automatic TOC; Chương as Heading1 and Điều as Heading2; then Lịch sử ban hành and Lưu tài liệu. The SOP-only sections were removed. The footer font was changed from Times New Roman to Calibri.
  - Text changes are limited to spelling, case and consistency. Examples: Neu/họp/trường họp/máy baytrong/mocrofiche/Calender/kí; GMPC → GMPc; CBCNV → CBNV; PHC/P.HC/Phòng HC/Phòng nhân sự → Phòng HCNS; "Bên B" and NSDLĐ → NLĐ / Công ty; "Luật Lao Động" → Bộ luật Lao động; "Chương 4, mục 4" → Điều 10, Chương III; -/+/=> → bullets, sub-items and a bold line. No rule was changed.
- 2026-10-16: Read the 4 uploaded source files. Text dumps are in `docs/_extract/`. Wrote the Nội quy review (see Source files). Nothing has been built yet.
- 2026-10-16: Created the project (`source/`, `docs/`, `output/`, `.claude-project`) at the user's request, as a separate project from `20260927 Quy Trinh MDC` so the MD&C SOPs (on hold for the manager's comments) don't mix with company-wide documents.
