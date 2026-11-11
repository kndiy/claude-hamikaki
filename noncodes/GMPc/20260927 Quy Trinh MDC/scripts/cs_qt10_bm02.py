"""CS-QT-10.BM02 — Sổ quản lý chăm sóc khách hàng hiện hữu (one workbook, one data source for CS-QT-10).

Styled like MDC-QT-03.BM01: navy title row 1, teal header row 2, grey example row 4, data from row 5,
sheet "Hướng dẫn" and dropdown lists in "Danh mục". Deadlines use the CS-QT-10 proposals
(recertification notice 6 months ahead, acknowledgement within 1 working day, survey within 30 days).
Works in Excel 2010 and later (AGGREGATE instead of MAXIFS, which this company's Excel lacks). Build: python scripts/cs_qt10_bm02.py
"""
import os
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

import cs_qt10
from cs_qt10_forms import CRITERIA

OUT = os.path.join(cs_qt10.OUT_BASE, cs_qt10.OUT[0], "CS-QT-10.BM02 So quan ly cham soc KH hien huu.xlsx")
CODE = "CS-QT-10.BM02"
FIRST, LAST = 5, 504  # data rows (row 4 = example)

NAVY, TEAL, GREY = "1B1464", "00B5A5", "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---------------------------------------------------------------- dropdown lists
LISTS = {
    "TrangThaiKH": ["Đang chăm sóc", "Tạm ngừng", "Đã đóng"],
    "CoKhong": ["Có", "Không"],
    "CoKhongChuaRo": ["Có", "Không", "Chưa rõ"],
    "ChungNhan": ["WHO-GMP", "EU-GMP", "PIC/S-GMP", "HS-GMP", "CGMP-ASEAN", "GSP", "GDP", "GLP", "ISO 13485",
                  "ISO 22000 / HACCP", "Khác"],
    "LoaiMoc": ["Tái đánh giá", "Đánh giá duy trì", "Gia hạn chứng nhận"],
    "HoatDong": ["Khảo sát sau dự án", "Liên hệ định kỳ", "Khảo sát hằng năm", "Tin tức ngành / quy định mới",
                 "Tri ân / sự kiện", "Nhắc tái đánh giá", "Khác"],
    "Kenh": ["Điện thoại", "Email", "Zalo", "Gặp trực tiếp", "Hội thảo / sự kiện", "Khác"],
    "TrangThaiKH_KeHoach": ["Kế hoạch", "Đã thực hiện", "Hoãn", "Hủy"],
    "PhatSinh": ["Không", "Yêu cầu / khiếu nại → 5_YeuCau", "Cơ hội mới → 6_CoHoi"],
    "LoaiYeuCau": ["Yêu cầu thông tin", "Hỗ trợ chuyên môn", "Hóa đơn / thanh toán", "Khiếu nại", "Góp ý", "Khác"],
    "MucDo": ["Thường", "Nghiêm trọng"],
    "BoPhan": ["CS (tự xử lý)", "RTA", "Kế toán (CS)", "MD&C", "Trưởng phòng CS", "COO / CEO"],
    "TrangThaiYC": ["Mới", "Đang xử lý", "Hoàn tất", "Đã đóng"],
    "NguonCoHoi": ["Liên hệ định kỳ", "Khảo sát", "Tái đánh giá", "KH chủ động liên hệ", "RTA chuyển", "Khác"],
    "KetQuaCoHoi": ["Đã chuyển MD&C", "MD&C đang xử lý", "Đã báo giá", "Đã ký HĐ", "Không tiếp tục"],
    "DotKhaoSat": ["Sau dự án", "Hằng năm"],
    "GioiThieu": ["Chắc chắn có", "Có thể", "Không"],
}

EX = "Công ty CP Dược phẩm Ví Dụ"
MA = "KH-26-001"


def rng(sheet, col):
    return f"'{sheet}'!${col}${FIRST}:${col}${LAST}"


def name_of(sheet_code_col="B"):
    """Customer name looked up from 1_KhachHang by Mã KH in column sheet_code_col."""
    return (lambda r: f'=IF(${sheet_code_col}{r}="","",IFERROR(INDEX(\'1_KhachHang\'!$C:$C,'
                      f'MATCH(${sheet_code_col}{r},\'1_KhachHang\'!$B:$B,0)),"(chưa có trong 1_KhachHang)"))')


# Each sheet: (name, title, [(header, width, kind)], example row, notes)
# kind: None = input; "list:<key>" = dropdown; "date"; "int"; "score"; callable(r) -> formula
SHEETS = [
    ("1_KhachHang", "DANH SÁCH KHÁCH HÀNG HIỆN HỮU", [
        ("STT", 5, "int"), ("Mã KH", 12, None), ("Tên khách hàng", 30, None), ("Địa chỉ", 28, None),
        ("Lĩnh vực / tiêu chuẩn đã đạt", 22, None), ("Người liên hệ", 18, None), ("Chức vụ", 14, None),
        ("Điện thoại", 14, None), ("Email", 22, None), ("Số / ngày hợp đồng", 18, None),
        ("Ngày hoàn tất HĐ", 12, "date"), ("Ngày tiếp nhận bàn giao (BM01)", 13, "date"),
        ("Người phụ trách CS", 16, None), ("Tần suất liên hệ (tháng)", 10, "int"),
        ("Ngày liên hệ gần nhất (tự động)", 13, lambda r: f'=IF(B{r}="","",IFERROR(_xlfn.AGGREGATE(14,6,'
                                                          f"'4_NhatKy'!$B$4:$B${LAST}/('4_NhatKy'!$C$4:$C${LAST}=B{r}),1),\"\"))"),
        ("Ngày liên hệ tiếp theo (tự động)", 13, lambda r: f'=IF(OR(B{r}="",N{r}=""),"",EDATE(IF(O{r}="",L{r},O{r}),N{r}))'),
        ("Cảnh báo (tự động)", 16, lambda r: f'=IF(OR(P{r}="",R{r}<>"Đang chăm sóc"),"",IF(P{r}<TODAY(),"Quá hạn liên hệ",'
                                             f'IF(P{r}-TODAY()<=14,"Sắp đến hạn","")))'),
        ("Trạng thái", 14, "list:TrangThaiKH"), ("Đồng ý nhận tin tức", 10, "list:CoKhong"), ("Ghi chú", 28, None),
    ], [1, MA, EX, "KCN Ví Dụ, tỉnh Ví Dụ", "WHO-GMP", "Nguyễn Văn A", "Giám đốc nhà máy", "0901 234 567",
        "vana@viidu.com.vn", "HĐ 12/2025/GMPc – 10/03/2025", "2026-06-30", "2026-07-03", "Trần Thị C (CS)", 3,
        None, None, None, "Đang chăm sóc", "Có", "Ví dụ minh họa"]),

    ("2_ChungNhan", "THEO DÕI CHỨNG NHẬN & MỐC TÁI ĐÁNH GIÁ", [
        ("STT", 5, "int"), ("Mã KH", 12, None), ("Tên khách hàng (tự động)", 28, name_of()),
        ("Chứng nhận / tiêu chuẩn", 16, "list:ChungNhan"), ("Số chứng nhận", 16, None), ("Cơ quan cấp", 16, None),
        ("Ngày cấp", 12, "date"), ("Hiệu lực đến", 12, "date"), ("Loại mốc", 16, "list:LoaiMoc"),
        ("Ngày mốc", 12, "date"),
        ("Hạn báo RTA & KH (mốc − 6 tháng, tự động)", 14, lambda r: f'=IF(J{r}="","",EDATE(J{r},-6))'),
        ("Số ngày còn lại đến mốc (tự động)", 11, lambda r: f'=IF(J{r}="","",J{r}-TODAY())'),
        ("Ngày đã báo RTA", 12, "date"), ("Ngày đã nhắc KH", 12, "date"),
        ("Cảnh báo (tự động)", 18, lambda r: f'=IF(J{r}="","",IF(AND(TODAY()>=K{r},OR(M{r}="",N{r}="")),"PHẢI BÁO NGAY",'
                                             f'IF(AND(OR(M{r}="",N{r}=""),K{r}-TODAY()<=30),"Sắp đến hạn báo","")))'),
        ("KH có nhu cầu hỗ trợ", 12, "list:CoKhongChuaRo"), ("Mã cơ hội (6_CoHoi)", 12, None), ("Ghi chú", 28, None),
    ], [1, MA, None, "WHO-GMP", "123/GCN-QLD", "Cục Quản lý Dược", "2024-06-15", "2027-06-15", "Tái đánh giá",
        "2027-06-15", None, None, None, None, None, "Chưa rõ", None, "Ví dụ minh họa"]),

    ("3_KeHoach", "KẾ HOẠCH CHĂM SÓC KHÁCH HÀNG", [
        ("STT", 5, "int"), ("Mã KH", 12, None), ("Tên khách hàng (tự động)", 28, name_of()), ("Năm", 8, "int"),
        ("Hoạt động", 22, "list:HoatDong"), ("Kênh", 14, "list:Kenh"), ("Thời điểm dự kiến", 12, "date"),
        ("Người thực hiện", 16, None), ("Trạng thái", 13, "list:TrangThaiKH_KeHoach"), ("Ngày thực hiện", 12, "date"),
        ("Ngày TP CS phê duyệt kế hoạch", 13, "date"), ("Ghi chú", 30, None),
    ], [1, MA, None, 2026, "Khảo sát sau dự án", "Email", "2026-08-01", "Trần Thị C (CS)", "Kế hoạch", None,
        "2026-07-05", "Ví dụ minh họa"]),

    ("4_NhatKy", "NHẬT KÝ CHĂM SÓC KHÁCH HÀNG", [
        ("STT", 5, "int"), ("Ngày liên hệ", 12, "date"), ("Mã KH", 12, None),
        ("Tên khách hàng (tự động)", 28, name_of("C")), ("Người thực hiện", 16, None), ("Hình thức", 13, "list:Kenh"),
        ("Loại hoạt động", 20, "list:HoatDong"), ("Nội dung trao đổi", 36, None), ("Kết quả / phản hồi KH", 32, None),
        ("Phát sinh", 24, "list:PhatSinh"), ("Mã phiếu YC / cơ hội", 12, None), ("Hành động tiếp theo", 26, None),
        ("Ngày theo dõi tiếp theo", 12, "date"),
    ], [1, "2026-07-20", MA, None, "Trần Thị C (CS)", "Điện thoại", "Liên hệ định kỳ",
        "Hỏi thăm tình hình vận hành sau khi đạt WHO-GMP", "KH hài lòng, quan tâm nâng cấp lên EU-GMP năm 2027",
        "Cơ hội mới → 6_CoHoi", "CH-26-001", "Chuyển MD&C", "2026-08-20"]),

    ("5_YeuCau", "THEO DÕI YÊU CẦU – PHẢN HỒI – KHIẾU NẠI", [
        ("Mã phiếu", 11, None), ("Ngày tiếp nhận", 12, "date"), ("Mã KH", 12, None),
        ("Tên khách hàng (tự động)", 26, name_of("C")), ("Kênh tiếp nhận", 13, "list:Kenh"),
        ("Người phản ánh (phía KH)", 18, None), ("Loại", 18, "list:LoaiYeuCau"), ("Mức độ", 12, "list:MucDo"),
        ("Nội dung", 36, None),
        ("Hạn xác nhận tiếp nhận (+1 ngày LV, tự động)", 13, lambda r: f'=IF(B{r}="","",WORKDAY(B{r},1))'),
        ("Ngày xác nhận với KH", 12, "date"), ("Bộ phận xử lý", 16, "list:BoPhan"), ("Người xử lý", 16, None),
        ("Hạn xử lý", 12, "date"), ("Kết quả xử lý", 30, None), ("Ngày hoàn tất", 12, "date"),
        ("Ngày phản hồi KH", 12, "date"), ("Trạng thái", 12, "list:TrangThaiYC"),
        ("Cảnh báo (tự động)", 18, lambda r: f'=IF(B{r}="","",IF(AND(K{r}="",TODAY()>J{r}),"Quá hạn xác nhận",'
                                             f'IF(AND(R{r}<>"Hoàn tất",R{r}<>"Đã đóng",N{r}<>"",TODAY()>N{r}),"Quá hạn xử lý","")))'),
        ("Ghi chú", 24, None),
    ], ["YC-26-001", "2026-07-22", MA, None, "Email", "Nguyễn Văn A", "Hỗ trợ chuyên môn", "Thường",
        "Hỏi cách cập nhật SOP vệ sinh theo quy định mới", None, "2026-07-22", "RTA", "Lê Văn D (RTA)",
        "2026-07-29", "RTA gửi hướng dẫn cập nhật SOP", "2026-07-27", "2026-07-27", "Hoàn tất", None,
        "Ví dụ minh họa"]),

    ("6_CoHoi", "CƠ HỘI KINH DOANH MỚI CHUYỂN MD&C", [
        ("Mã cơ hội", 11, None), ("Ngày ghi nhận", 12, "date"), ("Mã KH", 12, None),
        ("Tên khách hàng (tự động)", 26, name_of("C")), ("Nguồn", 18, "list:NguonCoHoi"),
        ("Nhu cầu / dịch vụ", 32, None), ("Tiêu chuẩn / lĩnh vực", 16, None), ("Thời điểm dự kiến", 14, None),
        ("Người liên hệ phía KH", 18, None), ("Ngày chuyển MD&C", 12, "date"), ("Người MD&C tiếp nhận", 16, None),
        ("Mã yêu cầu do MD&C cấp", 14, None), ("Kết quả", 16, "list:KetQuaCoHoi"), ("Ngày cập nhật", 12, "date"),
        ("Ghi chú", 26, None),
    ], ["CH-26-001", "2026-07-20", MA, None, "Liên hệ định kỳ", "Tư vấn nâng cấp nhà máy lên EU-GMP", "EU-GMP",
        "Năm 2027", "Nguyễn Văn A", "2026-07-21", "Nguyễn Thị B (MD&C)", None, "Đã chuyển MD&C", "2026-07-21",
        "Ví dụ minh họa"]),

    ("7_KhaoSat", "KẾT QUẢ KHẢO SÁT MỨC ĐỘ HÀI LÒNG (BM03)", [
        ("STT", 5, "int"), ("Mã KH", 12, None), ("Tên khách hàng (tự động)", 26, name_of()),
        ("Đợt", 12, "list:DotKhaoSat"),
        ("Hạn gửi (sau dự án: bàn giao + 30 ngày, tự động)", 14,
         lambda r: f'=IF(OR(B{r}="",D{r}<>"Sau dự án"),"",IFERROR(INDEX(\'1_KhachHang\'!$L:$L,'
                   f'MATCH(B{r},\'1_KhachHang\'!$B:$B,0))+30,""))'),
        ("Ngày gửi", 12, "date"), ("Ngày nhận", 12, "date"),
    ] + [(f"TC{k + 1}. {c}", 11, "score") for k, c in enumerate(CRITERIA)] + [
        ("Điểm TB (tự động)", 9, lambda r: f'=IF(COUNT(H{r}:O{r})=0,"",ROUND(AVERAGE(H{r}:O{r}),2))'),
        ("Có tiêu chí ≤ 2 (tự động)", 13, lambda r: f'=IF(COUNT(H{r}:O{r})=0,"",IF(MIN(H{r}:O{r})<=2,"Có → Bước 7",""))'),
        ("Sẵn sàng giới thiệu", 13, "list:GioiThieu"), ("Nhu cầu 12 tháng tới", 26, None),
        ("Mã phiếu YC / cơ hội", 12, None), ("Góp ý", 30, None),
    ], [1, MA, None, "Sau dự án", None, "2026-07-25", "2026-08-02", 5, 4, 4, 5, 5, 4, 4, 5, None, None,
        "Chắc chắn có", "Nâng cấp EU-GMP", "CH-26-001", "Ví dụ minh họa"]),
]

GUIDE = [
    ("Mục đích", f"{CODE} là nguồn dữ liệu duy nhất của quy trình CS-QT-10 — Chăm sóc khách hàng sau hợp đồng. "
                 "Mọi khách hàng, chứng nhận, kế hoạch, lần liên hệ, yêu cầu – khiếu nại, cơ hội và kết quả khảo sát "
                 "đều ghi trong file này (không lập file riêng)."),
    ("Các sheet", "1_KhachHang (Bước 2) · 2_ChungNhan (Bước 6) · 3_KeHoach (Bước 3) · 4_NhatKy (Bước 5) · "
                  "5_YeuCau (Bước 7) · 6_CoHoi (Bước 8) · 7_KhaoSat (Bước 4, 5) · 8_BaoCao (Bước 9)."),
    ("Mã KH", "Nhập Mã KH ở sheet 1_KhachHang trước; các sheet khác chỉ nhập Mã KH, tên khách hàng tự hiển thị. "
              "Quy tắc mã khách hàng chung toàn công ty chưa chốt — tạm dùng KH-YY-NNN."),
    ("Mã phiếu", "Yêu cầu / khiếu nại: YC-YY-NNN. Cơ hội: CH-YY-NNN. Mã yêu cầu chính thức do MD&C cấp khi tiếp nhận."),
    ("Cột tự động", "Cột có chữ (tự động) chứa công thức — không nhập tay. Cảnh báo tô đỏ = quá hạn / phải báo ngay; "
                    "tô cam = sắp đến hạn."),
    ("Mốc thời gian", "Báo RTA & nhắc KH tối thiểu 6 tháng trước mốc tái đánh giá (GMPc-CN-01). Xác nhận tiếp nhận "
                      "yêu cầu trong 1 ngày làm việc; khảo sát sau dự án trong 30 ngày sau bàn giao; liên hệ theo tần "
                      "suất ở cột N sheet 1_KhachHang (đề xuất 3 tháng). Các mốc này là ĐỀ XUẤT, chờ phê duyệt."),
    ("Dòng ví dụ", "Dòng 4 ở mỗi sheet là ví dụ minh họa — không tính vào báo cáo; có thể ẩn khi dùng thật."),
    ("Báo cáo", "Sheet 8_BaoCao: nhập Năm và Quý, các chỉ tiêu tự tính cho Bước 9."),
    ("Yêu cầu", "Excel 2010 trở lên."),
]


def style_title(ws, text, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(1, 1, text)
    c.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 26


def add_lists(wb):
    ws = wb.create_sheet("Danh mục")
    refs = {}
    for k, (key, items) in enumerate(LISTS.items(), start=1):
        col = L(k)
        h = ws.cell(1, k, key)
        h.font = Font(bold=True, color=NAVY)
        h.fill = PatternFill("solid", fgColor=TEAL)
        for n, v in enumerate(items, start=2):
            ws.cell(n, k, v)
        ws.column_dimensions[col].width = max(14, max(len(x) for x in items) + 2)
        refs[key] = f"'Danh mục'!${col}$2:${col}${len(items) + 1}"
    return refs


def build_sheet(wb, refs, name, title, cols, example):
    ws = wb.create_sheet(name)
    style_title(ws, f"{CODE} — {title}", len(cols))
    ws.row_dimensions[2].height = 45
    ws.row_dimensions[3].height = 4
    for k, (hdr, width, kind) in enumerate(cols, start=1):
        col = L(k)
        c = ws.cell(2, k, hdr)
        c.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        c.fill = PatternFill("solid", fgColor=TEAL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[col].width = width
        for r in range(4, LAST + 1):
            cell = ws.cell(r, k)
            cell.border = BORDER
            cell.font = Font(name="Calibri", size=10, italic=(r == 4))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if r == 4:
                cell.fill = PatternFill("solid", fgColor=GREY)
            if callable(kind):
                cell.value = kind(r)
                if r > 4:
                    cell.fill = PatternFill("solid", fgColor="F7F7FB")
            if kind == "date" or (callable(kind) and ("EDATE" in kind(r) or "WORKDAY" in kind(r)
                                                      or "AGGREGATE" in kind(r) or "+30" in kind(r))):
                cell.number_format = "dd/mm/yyyy"
        if isinstance(kind, str) and kind.startswith("list:"):
            dv = DataValidation(type="list", formula1=refs[kind[5:]], allow_blank=True)
            dv.add(f"{col}4:{col}{LAST}")
            ws.add_data_validation(dv)
        elif kind == "score":
            dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True,
                                error="Nhập điểm từ 1 đến 5", showErrorMessage=True)
            dv.add(f"{col}4:{col}{LAST}")
            ws.add_data_validation(dv)
            for r in range(4, LAST + 1):
                ws.cell(r, k).alignment = Alignment(horizontal="center", vertical="top")
        if hdr.startswith("Cảnh báo") or hdr.startswith("Có tiêu chí"):
            area = f"{col}4:{col}{LAST}"
            ws.conditional_formatting.add(area, FormulaRule(
                formula=[f'OR(ISNUMBER(SEARCH("Quá hạn",{col}4)),ISNUMBER(SEARCH("PHẢI",{col}4)),'
                         f'ISNUMBER(SEARCH("Bước 7",{col}4)))'],
                fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(bold=True, color="9C0006")))
            ws.conditional_formatting.add(area, FormulaRule(
                formula=[f'ISNUMBER(SEARCH("Sắp",{col}4))'], fill=PatternFill("solid", fgColor="FFE699")))
    for k, v in enumerate(example, start=1):
        if v is None or callable(cols[k - 1][2]):
            continue
        cell = ws.cell(4, k)
        if isinstance(v, str) and len(v) == 10 and v[4] == "-" and v[7] == "-":
            from datetime import date
            cell.value = date(int(v[:4]), int(v[5:7]), int(v[8:]))
        else:
            cell.value = v
    ws.freeze_panes = "D4" if name != "4_NhatKy" else "E4"
    ws.auto_filter.ref = f"A2:{L(len(cols))}{LAST}"
    ws.sheet_view.zoomScale = 90
    return ws


def build_report(wb):
    ws = wb.create_sheet("8_BaoCao")
    style_title(ws, f"{CODE} — BÁO CÁO CHĂM SÓC KHÁCH HÀNG THEO QUÝ (Bước 9)", 3)
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 50
    lab = Font(name="Calibri", size=11, bold=True, color=NAVY)
    inp = PatternFill("solid", fgColor="FFF2CC")
    rows = [
        ("Năm", 2026, "Nhập năm"),
        ("Quý (1–4)", 3, "Nhập quý"),
        ("Từ ngày", "=DATE(B3,(B4-1)*3+1,1)", ""),
        ("Đến ngày", "=EDATE(B5,3)-1", ""),
    ]
    K, Y, CN, NK, KS, CH = "'1_KhachHang'", "'5_YeuCau'", "'2_ChungNhan'", "'4_NhatKy'", "'7_KhaoSat'", "'6_CoHoi'"
    f, t = "\">=\"&$B$5", "\"<=\"&$B$6"

    def between(sheet, col):
        return f"{sheet}!${col}${FIRST}:${col}${LAST},{f},{sheet}!${col}${FIRST}:${col}${LAST},{t}"

    ind = [
        ("KHÁCH HÀNG", None, None),
        ("Số khách hàng đang chăm sóc (hiện tại)", f'=COUNTIF({K}!$R${FIRST}:$R${LAST},"Đang chăm sóc")', "1_KhachHang"),
        ("Khách hàng tiếp nhận mới trong kỳ", f"=COUNTIFS({between(K, 'L')})", "Theo ngày tiếp nhận bàn giao"),
        ("Khách hàng quá hạn liên hệ (hiện tại)", f'=COUNTIF({K}!$Q${FIRST}:$Q${LAST},"Quá hạn liên hệ")', ""),
        ("Số lần liên hệ trong kỳ", f"=COUNTIFS({between(NK, 'B')})", "4_NhatKy"),
        ("MỨC ĐỘ HÀI LÒNG", None, None),
        ("Số phiếu khảo sát nhận được trong kỳ", f"=COUNTIFS({between(KS, 'G')})", "7_KhaoSat, theo ngày nhận"),
        ("Điểm hài lòng trung bình trong kỳ (1–5)", f'=IFERROR(ROUND(AVERAGEIFS({KS}!$P${FIRST}:$P${LAST},{between(KS, "G")}),2),"")', ""),
        ("Số phiếu có tiêu chí ≤ 2 điểm", f'=COUNTIFS({between(KS, "G")},{KS}!$Q${FIRST}:$Q${LAST},"Có*")', "Phải xử lý theo Bước 7"),
        ("YÊU CẦU – KHIẾU NẠI", None, None),
        ("Yêu cầu / phản hồi tiếp nhận trong kỳ", f"=COUNTIFS({between(Y, 'B')})", "5_YeuCau"),
        ("   trong đó: khiếu nại", f'=COUNTIFS({between(Y, "B")},{Y}!$G${FIRST}:$G${LAST},"Khiếu nại")', ""),
        ("   đã hoàn tất / đóng", f'=COUNTIFS({between(Y, "B")},{Y}!$R${FIRST}:$R${LAST},"Hoàn tất")'
                                 f'+COUNTIFS({between(Y, "B")},{Y}!$R${FIRST}:$R${LAST},"Đã đóng")', ""),
        ("Phiếu quá hạn xác nhận / xử lý (hiện tại)", f'=COUNTIF({Y}!$S${FIRST}:$S${LAST},"Quá hạn*")', ""),
        ("TÁI ĐÁNH GIÁ / CHỨNG NHẬN", None, None),
        ("Mốc tái đánh giá trong 6 tháng tới", f'=COUNTIFS({CN}!$J${FIRST}:$J${LAST},">="&TODAY(),{CN}!$J${FIRST}:$J${LAST},"<="&EDATE(TODAY(),6))', "2_ChungNhan"),
        ("Chứng nhận PHẢI BÁO NGAY (hiện tại)", f'=COUNTIF({CN}!$O${FIRST}:$O${LAST},"PHẢI BÁO NGAY")', ""),
        ("CƠ HỘI KINH DOANH", None, None),
        ("Cơ hội chuyển MD&C trong kỳ", f"=COUNTIFS({between(CH, 'J')})", "6_CoHoi, theo ngày chuyển"),
        ("   trong đó: đã ký HĐ", f'=COUNTIFS({between(CH, "J")},{CH}!$M${FIRST}:$M${LAST},"Đã ký HĐ")', ""),
    ]
    r = 3
    for label, val, note in rows:
        ws.cell(r, 1, label).font = lab
        c = ws.cell(r, 2, val)
        c.border = BORDER
        if r <= 4:
            c.fill = inp
        else:
            c.number_format = "dd/mm/yyyy"
        ws.cell(r, 3, note).font = Font(italic=True, color="808080")
        r += 1
    r += 1
    for label, val, note in ind:
        if val is None:
            c = ws.cell(r, 1, label)
            c.font = Font(bold=True, color=NAVY)
            for k in (1, 2, 3):
                ws.cell(r, k).fill = PatternFill("solid", fgColor=TEAL)
        else:
            ws.cell(r, 1, label).border = BORDER
            v = ws.cell(r, 2, val)
            v.border = BORDER
            v.alignment = Alignment(horizontal="center")
            ws.cell(r, 3, note).font = Font(italic=True, color="808080")
        r += 1
    ws.cell(r + 1, 1, "Dòng ví dụ (dòng 4) ở các sheet không được tính.").font = Font(italic=True, color="808080")


def main():
    if os.path.exists(OUT) and "--force" not in sys.argv:
        props = load_workbook(OUT, read_only=True).properties
        if props.lastModifiedBy:  # openpyxl leaves it empty; Excel sets it on save
            sys.exit(f"Refusing to overwrite {OUT}: it has been edited and saved in Excel. "
                     "Copy those edits into the script first, then re-run with --force.")
    wb = Workbook()
    guide = wb.active
    guide.title = "Hướng dẫn"
    guide.column_dimensions["A"].width = 2
    guide.column_dimensions["B"].width = 18
    guide.column_dimensions["C"].width = 110
    t = guide.cell(2, 2, f"{CODE} — SỔ QUẢN LÝ CHĂM SÓC KHÁCH HÀNG HIỆN HỮU (quy trình CS-QT-10)")
    t.font = Font(size=14, bold=True, color=NAVY)
    for n, (k, v) in enumerate(GUIDE, start=4):
        guide.cell(n, 2, k).font = Font(bold=True, color=NAVY)
        c = guide.cell(n, 3, v)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        guide.cell(n, 2).alignment = Alignment(vertical="top")
    guide.sheet_view.showGridLines = False
    refs = add_lists(wb)
    for name, title, cols, example in SHEETS:
        build_sheet(wb, refs, name, title, cols, example)
    build_report(wb)
    wb.move_sheet("Danh mục", offset=len(wb.sheetnames))
    wb.properties.creator = "CS-QT-10 build script"
    wb.properties.lastModifiedBy = None
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
