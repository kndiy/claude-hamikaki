"""MDC-QT-07 — Sản xuất & tùy biến Hồ sơ năng lực (HSNL).

Source: draft 'MDC-QT-07 - Quy trình Sản xuất tùy biến Hồ sơ năng lực HSNL .docx' (same folder).
Build: python scripts/qt07.py
"""
import sys

from flowchart import Chart
import sop_builder

CODE = "MDC-QT-07"
TITLE = "SẢN XUẤT & TÙY BIẾN HỒ SƠ NĂNG LỰC"
OUT = ("MDC-QT-07 QT Sản xuất & tùy biến Hồ sơ năng lực (HSNL)", "MDC-QT-07 QT San xuat va tuy bien HSNL.docx")

PTTT = "Bộ phận Phát triển thị trường"
TH = "Bộ phận Thương hiệu & Truyền thông"

# ---------------------------------------------------------------- content
SIGN = [  # filled in by the user in Word, 2026-09-30
    ("Chuyên viên IT", "Lê Quang Huy", None),
    ("Trưởng phòng MD&C", "Đào Xuân Hưởng", None),
    ("Tổng giám đốc", "Đào Xuân Hưởng", None),
]

MUC_DICH = [
    "Quy trình này quy định các bước chuẩn hóa từ việc trích xuất dữ liệu gốc, tùy biến nội dung theo từng "
    "nhóm khách hàng mục tiêu, kiểm duyệt nhận diện thương hiệu cho đến khâu xuất bản HSNL. Mục tiêu là gia "
    "tăng tỷ lệ chuyển đổi khách hàng, tối ưu hóa tốc độ làm việc của bộ phận Thương hiệu & Truyền thông / "
    "Phát triển thị trường, và đảm bảo tính nhất quán của thương hiệu GMPc Việt Nam.",
]

PHAM_VI = [
    "Quy trình này áp dụng cho bộ phận Thương hiệu & Truyền thông, bộ phận Phát triển thị trường và các bộ "
    "phận kỹ thuật phối hợp cung cấp số liệu dự án.",
]

TRACH_NHIEM = [
    # The draft had two rows both named "Bộ phận Thương hiệu & Truyền thông"; merged into one.
    (TH, [
        "Quản lý và cập nhật tài liệu gốc Tuyên ngôn TN-001.",
        "Chịu trách nhiệm biên tập, tùy biến nội dung văn bản phù hợp với từng ngành.",
        "Thiết kế khung layout, đồ họa, tinh chỉnh hình ảnh dự án.",
        "Đảm bảo tính nhất quán về nhận diện thương hiệu (Màu sắc, Font chữ, Logo).",
    ]),
    (PTTT, [
        "Cung cấp thông tin chi tiết về chân dung khách hàng/đối tác cần gửi HSNL.",
        "Phối hợp nghiệm thu và trực tiếp gửi HSNL cho đối tác.",
    ]),
    ("Trưởng phòng MD&C", [
        "Phê duyệt nội dung chiến lược và các phiên bản tùy biến đặc biệt.",
    ]),
]

DINH_NGHIA = [
    ("MD&C", "Phòng Phát triển Thị trường & Truyền thông (Market Development & Communications)."),
    ("PTTT", "Bộ phận Phát triển thị trường (thuộc Phòng MD&C)."),
    ("HSNL", "Hồ sơ năng lực của doanh nghiệp."),
    ("Tùy biến", "Hành động lọc bỏ các thông tin không liên quan và nhấn mạnh/bổ sung các thông tin có giá trị "
                 "cao nhất đối với nhóm khách hàng đích."),
    ("Brand Guidelines", "Bộ quy chuẩn nhận diện thương hiệu của công ty."),
]

BUOC = [  # (số, tên bước, [nội dung], người thực hiện / ghi chú)
    ("1.", "Tiếp nhận yêu cầu và phân loại đối tượng khách hàng", [
        f"{PTTT} gửi yêu cầu làm HSNL về {TH} qua hệ thống quản lý công việc.",
        "Xác định rõ khách hàng thuộc nhóm nào trong 6 ngành chính sau:",
        "~1. Ngành Dược phẩm (Tập trung tiêu chuẩn GMP-WHO, EU-GMP).",
        "~2. Ngành Mỹ phẩm (Tập trung CGMP-ASEAN).",
        "~3. Ngành Thực phẩm bảo vệ sức khỏe / TP chức năng (GMP HS).",
        "~4. Ngành Thiết bị y tế (ISO 13485).",
        "~5. Ngành Thuốc thú y / Thủy sản (GMP-WHO cho thú y).",
        "~6. Ngành Sản phẩm sinh học / Phòng sạch công nghiệp (ISO/HACCP).",
    ], f"{PTTT} (yêu cầu).\n{TH} (tiếp nhận)."),
    ("2.", "Trích xuất dữ liệu gốc từ Tuyên ngôn TN-001", [
        "Mở tài liệu Master TN-001 (không được tự ý bịa số liệu bên ngoài).",
        "Lấy các thông tin nền tảng cố định: Giới thiệu chung, Sơ đồ tổ chức, Pháp lý doanh nghiệp.",
    ], f"{TH}."),
    ("3.", "Tùy biến nội dung", [
        "Thực hiện nguyên tắc: Khách hàng ngành nào, show dự án tiêu biểu và chuyên gia của ngành đó.",
        "Ví dụ: Nếu khách hàng ngành Dược, lọc từ TN-001 ra danh sách các nhà máy Dược đã tư vấn thành công, "
        "các hình ảnh phòng sạch, chứng nhận GMP tương ứng. Lược bỏ bớt thông tin ngành Mỹ phẩm/Thú y để tránh "
        "loãng hồ sơ.",
    ], f"{TH}.\nPhối hợp với Kỹ thuật để xin thêm ảnh dự án mới (nếu có)."),
    ("4.", "Thiết kế và duyệt nhận diện thương hiệu", [
        "Nhân viên thiết kế đưa nội dung đã tùy biến vào khung Template chuẩn của GMPc.",
        "Kiểm tra bộ nhận diện: Font chữ quy định, màu sắc chủ đạo (xanh/trắng), tỷ lệ logo, chất lượng hình "
        "ảnh nhà máy thiết kế không bị vỡ/mờ.",
        "Trưởng phòng Marketing duyệt chất lượng thiết kế hình ảnh.",
    ], f"{TH}."),
    ("5.", "Phê duyệt và xuất bản", [
        "Gửi bản thảo hoàn thiện cho Ban Giám đốc duyệt (đối với gói thầu/dự án đặc biệt bám sát chiến lược lớn).",
        "Xuất bản định dạng tùy theo yêu cầu: file PDF nén tối ưu dung lượng để gửi Email/Zalo; hoặc xuất file "
        "chất lượng cao gửi nhà in làm bản cứng.",
    ], "Trưởng bộ phận Thương hiệu & Truyền thông.\nBan Giám đốc."),
    ("6.", f"Bàn giao cho {PTTT} và lưu trữ", [
        f"Bàn giao sản phẩm cuối cùng cho {PTTT} để gửi khách hàng đúng hẹn.",
        "Lưu trữ file gốc và file xuất bản (PDF) lên Google Drive chung của phòng MD&C theo quy chuẩn đặt tên: "
        "[Năm]_[Tên Khách Hàng]_[Tên Ngành]_[Mã HSNL].",
    ], f"{TH}.\nNhân viên Phát triển thị trường tiếp nhận. Kết thúc QT-07."),
]

TAI_LIEU = []  # none in the draft

LICH_SU = [
    ("__/__/____", "01", "Ban hành lần đầu",
     "Ban hành quy trình MDC-QT-07 — Sản xuất & tùy biến Hồ sơ năng lực."),
]

LUU = ["", "", ""]  # none in the draft


# ---------------------------------------------------------------- sơ đồ khối
def flowchart():
    c = Chart(1000, 1180)
    cx = 440
    t = c.box(cx, 20, 520, 70, [CODE, TITLE], "title", size=18)
    b1 = c.box(cx, 135, 400, 78, ["B1. Tiếp nhận yêu cầu & phân loại KH",
                                  "Yêu cầu làm HSNL từ Bộ phận PTTT",
                                  "Xác định 1 trong 6 ngành chính"])
    n1 = c.box(830, 118, 300, 112, ["6 ngành chính:", "Dược phẩm · Mỹ phẩm · TPBVSK",
                                    "Thiết bị y tế · Thú y / Thủy sản",
                                    "Sinh học / Phòng sạch CN"], "note", bold_first=True, size=14)
    b2 = c.box(cx, 263, 400, 78, ["B2. Trích xuất dữ liệu gốc", "Master TN-001: Giới thiệu chung,",
                                  "Sơ đồ tổ chức, Pháp lý doanh nghiệp"])
    n2 = c.box(830, 280, 300, 44, ["Không tự ý bịa số liệu ngoài TN-001"], "note", size=14)
    b3 = c.box(cx, 391, 400, 78, ["B3. Tùy biến nội dung theo ngành", "Dự án tiêu biểu & chuyên gia",
                                  "đúng ngành của khách hàng"])
    n3 = c.box(830, 408, 300, 44, ["Kỹ thuật: ảnh dự án mới (nếu có)"], "note", size=14)
    b4 = c.box(cx, 519, 400, 78, ["B4. Thiết kế & duyệt nhận diện", "Template chuẩn GMPc",
                                  "Font, màu, logo, chất lượng ảnh"])
    d = c.diamond(cx, 645, 300, 110, ["Gói thầu / dự án", "đặc biệt?"])
    bgd = c.box(830, 672, 240, 56, ["Ban Giám đốc duyệt"], "note", size=15)
    b5 = c.box(cx, 805, 400, 78, ["B5. Xuất bản", "PDF nén (Email/Zalo) hoặc",
                                  "file chất lượng cao gửi nhà in"])
    b6 = c.box(cx, 933, 400, 78, ["B6. Bàn giao & lưu trữ", "Bàn giao Bộ phận PTTT",
                                  "Lưu Google Drive MD&C"])
    end = c.box(cx, 1061, 400, 60, ["Bộ phận PTTT gửi HSNL", "cho khách hàng"], "io", size=16)

    for a, b in [(t, b1), (b1, b2), (b2, b3), (b3, b4), (b4, d), (b5, b6), (b6, end)]:
        c.down(a, b)
    c.arrow([(cx, d["bot"]), (cx, b5["top"])])
    c.label(cx + 12, d["bot"] + 30, "Không", anchor="start")
    c.arrow([(d["r"], d["mid"]), (bgd["l"], bgd["mid"])])
    c.label((d["r"] + bgd["l"]) / 2, d["mid"] - 8, "Có")
    c.arrow([(bgd["cx"], bgd["bot"]), (bgd["cx"], b5["mid"]), (b5["r"], b5["mid"])])
    for b, n in [(b1, n1), (b2, n2), (b3, n3)]:
        c.link([(b["r"], b["mid"]), (n["l"], n["mid"])])
    return c


if __name__ == "__main__":
    sop_builder.build(sys.modules[__name__])
