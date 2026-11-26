# CLAUDE.md change log archive

Older change-log entries moved out of `../CLAUDE.md` on 2026-11-26, verbatim and in their original order. The newest entries stay in `CLAUDE.md`; when an entry leaves it, append it here.

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
