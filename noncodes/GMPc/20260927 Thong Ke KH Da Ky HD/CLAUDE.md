# Thống kê khách hàng đã ký hợp đồng

## Mục tiêu (đã xác nhận 2026-09-27)
Thống kê KH đã ký HĐ (2011–2026) và **đối chiếu / bổ sung dữ liệu từ các hợp đồng scan** để check lại và điền thông tin còn thiếu.
- Nguồn scan: `L:\04. Tổng hợp Data khách hàng\01. Danh sách KH đã kí HĐ\Scan hợp đồng tư vấn\` (thư mục 2021–2026 + OKS, 135 PDF).
- Mẫu: theo `260625 Danh sách KH đã kí HĐ 2011-2026.xlsx`, tổng hợp tất cả các năm; đã đề xuất mẫu mới (xem Output).
- Output: **file Excel** trong `output/`.
- Thông tin còn thiếu (kể cả dòng Công ty trống): **để trống**, user bổ sung sau. Không tự suy đoán/forward-fill.

## Quy tắc dữ liệu (bắt buộc)
- `L:\` là ổ map tới NAS công ty → **KHÔNG edit trực tiếp**. Chỉ **đọc** từ NAS (không copy scan PDF về); kết quả ghi ra local.
- Không sửa file trong `data/raw/`. Dữ liệu trung gian ở `data/processed/`, kết quả ở `output/`.

## Cấu trúc
Vị trí: `C:\Users\Khiem Nguyen\Documents\projects\noncodes\GMPc\20260927 Thong Ke KH Da Ky HD` (local, KHÔNG dùng OneDrive; thuộc group `GMPc`).
```
data/raw/                     bản copy nguyên gốc file Excel mẫu từ NAS (read-only)
data/processed/scan_index.csv danh mục 135 file scan (id S001–S135, thư mục, số trang, trang đã render)
data/processed/scan_text/     lớp chữ OCR có sẵn trong PDF (lỗi dấu tiếng Việt; chỉ dùng dò trang/số tiền)
data/processed/scan_extract.jsonl  dữ liệu trích từ từng HĐ scan (đọc thủ công từ ảnh trang) - NGUỒN CHÍNH
scripts/01_prepare_scans.py   lập danh mục scan, trích lớp chữ, render ảnh trang (arg1 = thư mục ảnh tạm)
scripts/02_build_report.py    đối chiếu Excel ↔ scan, xuất báo cáo Excel
output/260927 Thong ke KH da ky HD 2011-2026.xlsx  báo cáo
requirements.txt
```
Chạy lại: `python -m pip install -r requirements.txt` → `python -X utf8 scripts/02_build_report.py`
(chỉ chạy lại 01 nếu thư mục scan trên NAS có file mới; id S### sẽ đánh lại theo thứ tự thư mục/tên file).

## File mẫu (Excel gốc)
`data/raw/260625 Danh sách KH đã kí HĐ 2011-2026.xlsx` (copy 2026-09-27, hash khớp NAS).
Sheet `Tổng hợp` A1:L419: STT, Công ty, Thời điểm kí, Số HĐ, Nội dung, Giá trị (không VAT), Tình trạng, Lĩnh vực, Đầu mối, SĐT, Email, Đầu mối GMPc.
- Khối theo năm 2026→2011; dòng tiêu đề năm có công thức SUM ở cột F (vd 38,27 tỷ = tổng 2025) → loại khi xử lý.
- Một số giá trị là công thức `=x/1.08` (tính ngược từ tổng gồm VAT).
- Dòng Công ty trống = HĐ tiếp theo của KH dòng trên (giữ trống theo yêu cầu, script bổ sung tên từ scan nếu có).

## Output - mẫu mới (đề xuất)
Sheets: Hướng dẫn | **Danh sách HĐ** (1 dòng/HĐ, tất cả các năm, 27 cột) | TK theo năm | TK theo lĩnh vực | TK theo tình trạng | Khách hàng | Đối chiếu scan | Cần kiểm tra | Danh mục.
- Màu: xanh lá = bổ sung từ scan; vàng = Excel ≠ scan (giữ Excel, ghi giá trị scan ở Ghi chú); xanh dương = chỉ có trong scan; xám = không tính thống kê.
- KH định danh theo MST (nếu có) hoặc tên đã chuẩn hoá (bỏ dấu, bỏ "Công ty/CP/TNHH…", bỏ phần trong ngoặc).
- Lĩnh vực/Tình trạng chuẩn hoá theo bảng `LINH_VUC_MAP` / `TINH_TRANG_MAP` trong script (đề xuất, chờ user duyệt).

## Kết quả lần chạy 2026-09-27
- Excel 401 dòng HĐ; scan 135 file → 126 khớp (117 theo số HĐ, 7 theo tên+thời gian, 2 thủ công), 6 chỉ có trong scan, 3 loại.
- Chỉ có trong scan: S130 Minh Đan 2026 (GMPc), S131–S135 (5 HĐ do **Công ty CP OKS** ký).
- Loại: S047 (GMPc là bên thuê 5SPRO), S107 (trùng S106 Vincos), S002 (HĐ Dabaco–Anh Khang, GMPc không ký).
- Khớp thủ công: S046 Medtec ↔ dòng 103; S054 Vinmec ↔ dòng 107.
- 29 dòng lệch Excel↔scan (chủ yếu tháng ký, số HĐ, vài giá trị) → sheet "Cần kiểm tra".
- Giá trị HĐ: đầy đủ 2023–2026, 2022 65%, 2021 chỉ 1 HĐ; 2011–2020 không có.

## Công cụ
- Python 3.13.15 (winget, user scope) `%LOCALAPPDATA%\Programs\Python\Python313\python.exe`, có trong User PATH.
- Thư viện: openpyxl, pymupdf (render PDF), pypdf; pandas đã cài nhưng chưa dùng.
- Excel COM (PowerShell) dùng để xuất PDF kiểm tra hiển thị.
- Máy chưa có poppler (`pdftoppm`) → công cụ Read của Claude không đọc PDF trực tiếp; đang render qua pymupdf.
- Có sẵn: Excel, Git 2.45.1, winget 1.29.

### Công cụ tự kiểm tra (launcher)
Launcher (`_launcher`) kiểm tra bảng này khi mở project và đề nghị cài phần còn thiếu (chạy lệnh Install trong thư mục project).

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| Python | `python --version` | required | `winget install --id Python.Python.3.13 -e --scope user --source winget --accept-source-agreements --accept-package-agreements` | - |
| Python packages | `python -c "import openpyxl, pymupdf, pypdf"` | required | `python -m pip install -r requirements.txt` | - |
<!-- launcher:tools:windows:end -->

### Cài đặt trên máy mới
1. Tắt alias python.exe / python3.exe: Settings → Apps → Advanced app settings → App execution aliases.
2. `winget install -e --id Python.Python.3.13 --scope user --source winget --accept-source-agreements --accept-package-agreements`
   (bắt buộc có `--source winget` + `--accept-*`, nếu không winget hỏi điều khoản msstore và lỗi 0x8a150042 trong Claude Code).
3. Mở lại terminal → `python -m pip install -r requirements.txt` (trong thư mục project).
4. Tuỳ chọn (chưa cài, chưa xác minh ID gói winget): Poppler (để Claude đọc PDF trực tiếp), Tesseract OCR + tiếng Việt (OCR máy cho scan mới). Tra ID bằng `winget search poppler` / `winget search tesseract`.

## Open questions (chờ user)
1. Duyệt bảng chuẩn hoá Lĩnh vực (sheet Danh mục), đặc biệt nhóm "Sinh phẩm/Vaccine/Tế bào" và "Khác".
2. HĐ OKS có tính vào thống kê chung với GMPc không? (hiện tách cột riêng "Số HĐ OKS").
3. Với dòng lệch: lấy giá trị theo HĐ scan làm chuẩn hay giữ Excel? (hiện giữ Excel, chỉ đánh dấu).
4. HĐ Kim Thành (S040): HĐ ghi năm 2020 nhưng ở thư mục 2023 và Excel ghi 06/2023.
5. MST Chí Linh (S129) trùng MST Kella (S128) trong bản HĐ - nghi nhầm khi soạn HĐ.
6. S089 JCV: tên file ghi "Draft Contract" - đã ký chưa?
7. Có dùng thêm file `2600813 Danh sách KH đã kí HĐ 2023-2024-2025.xlsx` và file OKS trên NAS để đối chiếu không?

## Changelog

Older entries: `zClaudeDocs/CLAUDE_CHANGELOG.md`.

- 2026-09-27: Thêm `requirements.txt` và hướng dẫn cài đặt công cụ vào CLAUDE.md.
- 2026-09-27: Chuyển project vào `projects\noncodes\` (root đổi tên thành chữ thường; cấu trúc root > subRoot > project mới), thêm file đánh dấu `.claude-project`. Script dùng đường dẫn tương đối nên không cần sửa.
- 2026-09-27: Chuyển project từ `C:\Users\Khiem Nguyen\Documents\projects\noncodes\20260927 Thong Ke Khach Hang Da Ky Hop Dong` vào group `noncodes\GMPc\20260927 Thong Ke KH Da Ky HD` (robocopy, bỏ `__pycache__`). Sửa dòng "Vị trí"; check công cụ đổi `fitz` → `pymupdf` (tránh cảnh báo deprecated). Giữ nguyên `SCAN_ROOT` trên `L:\` (chỉ đọc). Kiểm tra: 142 file khớp hash; chạy lại `02_build_report.py` cho kết quả giống hệt bản cũ (401 dòng, 126 khớp, 6 chỉ có trong scan, 29 lệch; nội dung các sheet trùng khớp).
