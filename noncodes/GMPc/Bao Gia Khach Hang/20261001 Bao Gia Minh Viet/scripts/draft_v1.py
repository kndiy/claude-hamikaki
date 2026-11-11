"""Build draft v1 of the Minh Viet quotation from data/raw (raw file is never modified)."""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

SRC = "data/raw/261001 Bao gia Minh Viet.xlsx"
DST = "output/261001 Bao gia Minh Viet - draft v1.xlsx"
YELLOW = PatternFill("solid", fgColor="FFFF00")

wb = openpyxl.load_workbook(SRC)
ws = wb["Báo giá"]

def flag(*cells):
    for c in cells:
        ws[c].fill = YELLOW

# 1. GMPc addresses -> post-2025 administrative names (as in the Cafune template, 22/09/2026)
ws["B2"] = "Trụ sở: 18 Măng Cầm 1,  An Lạc Green Symphony, Sơn Đồng, Hà Nội"
ws["B3"] = "VPGD: 32 Vĩ Cầm 5, An Lạc Green Symphony, Sơn Đồng, Hà Nội"
ws["B4"] = "VPĐD: Lầu 2, 156/1/1 Cộng Hòa, Phường Bảy Hiền, TP Hồ Chí Minh"

# 2. Customer details still missing / to verify
ws["B12"] = ":  [Bổ sung số điện thoại]"
ws["B13"] = ":  [Bổ sung email]"
flag("B10", "B12", "B13")

# 3. Quote date: unambiguous dd/mm/yyyy (was "10/01/2026" = 10 Jan in VN format)
ws["B14"] = ":  01/10/2026"

# 4. Section I header: add English line like the sub-items
ws["B22"] = ("Tư vấn và thiết kế mặt bằng công nghệ các hạng mục theo CGMP-ASEAN; "
             "Quy hoạch tổng thể, Khái toán chi phí đầu tư\n"
             "Consulting and technological layout design in accordance with CGMP-ASEAN; "
             "master planning and preliminary investment cost estimate.")
ws.row_dimensions[22].height = 75
# Item 1: VN says "design layout", EN says "feedback on renovation layout" -> needs decision
flag("B22", "B23")

# 5. VAT note: Decree 72/2024 expired 31/12/2024 -> Decree 174/2025 (valid to 31/12/2026)
ws["B30"] = (":  Thuế suất thuế GTGT 8% áp dụng theo nghị định số 174/2025/NĐ-CP ban hành ngày "
             "30/06/2025 có hiệu lực thi hành từ ngày 01/07/2025 đến hết ngày 31/12/2026.\n"
             "    VAT rate of 8% is applied in accordance with Decree No. 174/2025/NĐ-CP, issued on "
             "June 30, 2025 and effective from July 1, 2025 to December 31, 2026.")

# 6. Payment: VN said 50% advance, EN said 30% -> aligned EN to VN (50/50), flagged for confirmation
ws["B32"] = (":  - Lần 1: Tạm ứng 50% phí dịch vụ trong vòng 05 ngày kể từ ngày ký kết hợp đồng.\n"
             "       1st payment: Advance 50% of the contract amount within 05 days after the signing of the Contract.\n"
             "    - Lần 2: Thanh toán 50% còn lại trong vòng 05 ngày sau khi nghiệm thu, thanh lý hợp đồng.\n"
             "       2nd payment: The remaining 50% of the contract amount within 05 days after project acceptance and contract liquidation.")
ws.row_dimensions[32].height = 80
flag("B32")

# Review sheet (delete before sending to customer)
rv = wb.create_sheet("Ghi chu soan thao")
rows = [
    ("#", "Loại", "Ô", "Nội dung", "Trạng thái"),
    (1, "Đã sửa", "B2:B4", "Địa chỉ GMPc cập nhật theo địa giới hành chính mới (Sơn Đồng; P. Bảy Hiền), giống mẫu Cafune 22/09/2026.", "Xong"),
    (2, "Đã sửa", "B14", "Ngày báo giá '10/01/2026' (= 10 tháng 1 theo định dạng VN) -> 01/10/2026.", "Xong"),
    (3, "Đã sửa", "B30", "Ghi chú VAT dẫn NĐ 72/2024 (hết hiệu lực 31/12/2024) -> NĐ 174/2025, hiệu lực đến 31/12/2026. Đã bỏ lỗi lặp 'lực lực'.", "Xong"),
    (4, "Đã sửa", "B22", "Thêm dòng tiếng Anh cho tiêu đề mục I.", "Kiểm tra câu chữ"),
    (5, "Cần xác nhận", "B32", "Bản gốc: VN tạm ứng 50%, EN tạm ứng 30% (EN cộng lại chỉ 80%). Bản nháp tạm theo VN: 50% / 50%. Xác nhận tỷ lệ.", "Chờ"),
    (6, "Cần xác nhận", "B22/B23", "Tiêu đề mục I có 'Quy hoạch tổng thể, Khái toán chi phí đầu tư' nhưng các hạng mục con là cải tạo (mặt bằng, khái toán cải tạo, tiến độ). Mục 1: VN 'Thiết kế mặt bằng' vs EN 'Providing feedback on the renovation' - là thiết kế mới hay góp ý cải tạo?", "Chờ"),
    (7, "Cần xác nhận", "B10", "Địa chỉ KH dùng 'Củ Chi' (cấp huyện đã bỏ từ 01/07/2025) và 'KCN' chưa có tên. Cần địa chỉ theo ĐKKD hiện hành.", "Chờ"),
    (8, "Cần bổ sung", "B12/B13", "SĐT và email chị Lan.", "Chờ"),
    (9, "Cần xác nhận", "B17", "Tên tiếng Anh 'Minh Viet International Joint Stock Company' - đối chiếu ĐKKD.", "Chờ"),
    (10, "Đề xuất", "B31", "Nên thêm điều kiện 'và Chủ đầu tư cung cấp đầy đủ thông tin/bản vẽ hiện trạng' vào thời gian thực hiện (như mẫu Cafune).", "Tùy chọn"),
    (11, "Đề xuất", "B30", "Nên thêm câu loại trừ: chưa bao gồm thiết kế chi tiết M&E, chi phí khảo sát/đo vẽ, phí nhà nước, chi phí phát sinh khác.", "Tùy chọn"),
    (12, "Lưu ý", "-", "Bỏ tô vàng và xoá sheet này trước khi gửi khách hàng. Giá 45.000.000 VND giữ nguyên theo bản gốc.", "-"),
]
for r in rows:
    rv.append(r)
for c in rv[1]:
    c.font = Font(bold=True)
for col, w in zip("ABCDE", (5, 14, 10, 100, 16)):
    rv.column_dimensions[col].width = w
for row in rv.iter_rows(min_row=2):
    for c in row:
        c.alignment = Alignment(wrap_text=True, vertical="top")

wb.save(DST)
print("saved", DST)
