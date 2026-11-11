"""MDC-QT-08 — Tìm kiếm & quản lý đối tác.

Source: draft 'MDC_QT_08_Doi tac QT_290726.docx' (same folder). Step titles come from the draft's own
flowchart (Word shapes); step 12 exists only in that flowchart, not in its steps table.
Build: python scripts/qt08.py
"""
import sys

from flowchart import Chart
import sop_builder

CODE = "MDC-QT-08"
TITLE = "TÌM KIẾM & QUẢN LÝ ĐỐI TÁC"
OUT = ("MDC-QT-08 - Tìm kiếm & quản lý đối tác", "MDC-QT-08 QT Tim kiem va quan ly doi tac.docx")

# ---------------------------------------------------------------- content
SIGN = [  # filled in by the user in Word, 2026-09-30
    ("Chuyên viên Phát triển đối tác chiến lược", "Nguyễn Kim Nhũ", None),
    ("Trưởng phòng MD&C", "Đào Xuân Hưởng", None),
    ("Tổng giám đốc", "Đào Xuân Hưởng", None),
]

MUC_DICH = [
    "Quy định trình tự, tiêu chí tìm kiếm, đánh giá và lựa chọn đối tác tư vấn GMP tiên tiến (EU-GMP, PIC/S, "
    "US FDA...) nước ngoài, nhằm đảm bảo Công ty lựa chọn được đối tác có đủ năng lực, kinh nghiệm và uy tín "
    "để phối hợp triển khai các dự án tư vấn nâng cấp nhà máy đạt tiêu chuẩn GMP tiên tiến cho khách hàng.",
]

PHAM_VI = [
    "Áp dụng cho toàn bộ hoạt động tìm kiếm, thu thập thông tin, đánh giá và lựa chọn các đơn vị/đối tác tư vấn "
    "GMP tiên tiến ở nước ngoài (ví dụ: Pharmexpert, Pharmazone, DAC, Kipa, Doshi và các đối tác khác) để hợp "
    "tác triển khai dịch vụ tư vấn GMP tiên tiến (EU-GMP, PIC/S, US FDA...) cho khách hàng của Công ty tại "
    "Việt Nam. Quy trình không áp dụng cho việc tìm kiếm, đánh giá đối tác/chuyên gia tư vấn GMP WHO hoặc các "
    "tiêu chuẩn GMP cơ bản trong nước.",
]

TRACH_NHIEM = [
    ("Trợ lý CEO / MD&C", [
        "Chủ trì tìm kiếm, thu thập thông tin, tổ chức trao đổi và lập hồ sơ đánh giá đối tác tư vấn GMP tiên tiến;",
        "Tổng hợp kết quả, đề xuất trình phê duyệt.",
    ]),
    ("RTA", [  # the draft says "Bộ phận TRA"; every step uses RTA
        "Tham gia đánh giá năng lực chuyên môn, phạm vi dịch vụ và đội ngũ chuyên gia của đối tác;",
        "Đóng góp ý kiến kỹ thuật trong quá trình phỏng vấn, trao đổi.",
    ]),
    ("Bộ phận Pháp lý/Hợp đồng", [
        "Rà soát tư cách pháp nhân của đối tác;",
        "Soạn thảo, thẩm định các thỏa thuận hợp tác (NDA, MOU, hợp đồng khung...).",
    ]),
    ("CEO", [
        "Phê duyệt danh sách đối tác được đánh giá, phê duyệt kết quả đánh giá và quyết định hợp tác chính thức.",
    ]),
]

DINH_NGHIA = [
    ("MD&C", "Phòng Phát triển Thị trường & Truyền thông (Market Development & Communications)."),
    ("RTA", "Phòng Quy chế & Kỹ thuật."),
    ("CEO", "Tổng Giám đốc."),
    ("GMP", "Good Manufacturing Practice – Thực hành sản xuất tốt."),
    ("EU-GMP", "GMP theo tiêu chuẩn của Liên minh Châu Âu (European Union GMP)."),
    ("PIC/S", "Pharmaceutical Inspection Co-operation Scheme – Hệ thống hợp tác thanh tra dược phẩm quốc tế."),
    ("US FDA", "Cục Quản lý Thực phẩm và Dược phẩm Hoa Kỳ (United States Food and Drug Administration)."),
    ("QMS", "Quality Management System – Hệ thống quản lý chất lượng."),
    ("Kích hoạt thanh tra", "Hoạt động chuẩn bị, nộp hồ sơ và làm việc với cơ quan quản lý để đăng ký/kích hoạt "
                            "đợt thanh tra GMP chính thức (EU-GMP/PIC/S/US FDA...)."),
    ("MA", "Marketing Authorization – Giấy phép lưu hành (đăng ký lưu hành thuốc) do cơ quan quản lý dược cấp."),
]

BUOC = [  # (số, tên bước, [nội dung], người thực hiện / ghi chú)
    ("1.", "Xác định nhu cầu tư vấn & tiêu chuẩn mục tiêu", [
        "Xác định nhu cầu tư vấn GMP tiên tiến của khách hàng/dự án và tiêu chuẩn mục tiêu cần đáp ứng "
        "(EU-GMP, PIC/S, US FDA...).",
    ], "Trợ lý CEO"),
    ("2.", "Lập danh sách đối tác tiềm năng", [
        "Lập danh sách sơ bộ các đơn vị tư vấn GMP tiên tiến tiềm năng (tham khảo các đối tác đã hợp tác như "
        "Pharmexpert, Pharmazone, DAC, Kipa, Doshi...; tìm kiếm bổ sung qua hiệp hội ngành, hội thảo, giới thiệu "
        "từ khách hàng/đối tác).",
    ], "Trợ lý CEO"),
    ("3.", "Kiểm tra hồ sơ năng lực / website", [
        "Kiểm tra thông tin công ty: hồ sơ năng lực (profile), website chính thức, tư cách pháp nhân, "
        "chứng chỉ/giấy phép hoạt động.",
    ], "Trợ lý CEO"),
    ("4.", "Đánh giá kinh nghiệm & số dự án", [
        "Đánh giá số năm kinh nghiệm và số lượng dự án tư vấn GMP tiên tiến đã/đang triển khai; ưu tiên các dự "
        "án đã thành công hoặc đang triển khai tại Việt Nam; yêu cầu cung cấp case study, danh sách khách hàng "
        "tham chiếu.",
    ], "Trợ lý CEO"),
    ("5.", "Đánh giá phạm vi dịch vụ cung cấp", [
        "Đánh giá phạm vi công việc trong gói tư vấn mà đối tác có thể cung cấp: thiết kế nhà máy/hệ thống; xây "
        "dựng hệ thống QMS; thẩm định (qualification/validation); đánh giá sẵn sàng thanh tra; hỗ trợ kích hoạt "
        "thanh tra.",
    ], "Trợ lý CEO, RTA"),
    ("6.", "Kiểm tra hồ sơ MA phục vụ kích hoạt thanh tra", [
        "Kiểm tra năng lực và kinh nghiệm của đối tác trong việc chuẩn bị, nộp và theo đuổi hồ sơ Giấy phép lưu "
        "hành (MA – Marketing Authorization) tại các thị trường mục tiêu; xác minh đối tác đã từng hỗ trợ khách "
        "hàng hoàn thiện hồ sơ MA làm căn cứ để cơ quan quản lý kích hoạt thanh tra GMP chính thức.",
    ], "RTA"),
    ("7.", "Trao đổi / phỏng vấn xác minh thực tế", [
        "Trao đổi/phỏng vấn trực tiếp (online hoặc offline) với đối tác để xác minh đối tác đã thực sự triển khai "
        "các phạm vi công việc nêu trên; đề nghị cung cấp bằng chứng như hợp đồng mẫu (đã che thông tin bảo mật), "
        "chứng nhận GMP mà khách hàng cũ đạt được, thư giới thiệu/thư tham chiếu.",
    ], "Trợ lý CEO, RTA"),
    ("8.", "Rà soát pháp lý & thỏa thuận hợp tác", [
        "Rà soát tư cách pháp nhân và soạn thảo/thẩm định các thỏa thuận hợp tác cần thiết (NDA, MOU, hợp đồng "
        "khung...) trước khi hợp tác chính thức.",
    ], "Bộ phận Pháp lý / RTA"),
    ("9.", "Tổng hợp, chấm điểm & lập báo cáo đề xuất", [
        "Tổng hợp toàn bộ kết quả đánh giá vào Bảng đánh giá đối tác tư vấn GMP tiên tiến, đề xuất hợp tác.",
    ], "Trợ lý CEO"),
    ("10.", "CEO phê duyệt", [
        "Trình Tổng giám đốc xem xét, phê duyệt danh sách đối tác đạt yêu cầu và quyết định ưu tiên hợp tác.",
    ], "Tổng giám đốc"),
    ("11.", "Cập nhật danh sách đối tác đã phê duyệt & lưu hồ sơ", [
        "Cập nhật Danh sách đối tác tư vấn GMP tiên tiến đã được phê duyệt và lưu trữ toàn bộ hồ sơ đánh giá.",
    ], "Trợ lý CEO"),
    ("12.", "Định kỳ rà soát, cập nhật đánh giá", [
        "Định kỳ rà soát, cập nhật đánh giá đối tác (≥ 1 lần/năm).",
    ], ""),  # from the draft's flowchart; performer not stated
]

TAI_LIEU = []  # none in the draft

LICH_SU = [
    ("__/__/____", "01", "Ban hành lần đầu", "Ban hành quy trình MDC-QT-08 — Tìm kiếm & quản lý đối tác."),
]

LUU = ["Phòng Hành chính / Trợ lý CEO", "", ""]


# ---------------------------------------------------------------- sơ đồ khối
def flowchart():
    """Two columns like the draft: B1–B6 down the left, B7–B12 up the right."""
    import textwrap
    c = Chart(1000, 760)
    lx, rx, w, h, gap, y0 = 245, 755, 450, 70, 100, 120
    c.box(500, 20, 620, 64, [CODE, TITLE], "title", size=18)

    def lines(i):
        return textwrap.wrap(f"B{i + 1}. {BUOC[i][1]}", 44)
    left = [c.box(lx, y0 + i * gap, w, h, lines(i), bold_first=False, size=17) for i in range(6)]
    right = [c.box(rx, y0 + (5 - i) * gap, w, h, lines(i + 6), "title" if i == 3 else "step",
                   bold_first=False, size=17) for i in range(6)]
    c.arrow([(500, 84), (500, 100), (lx, 100), (lx, left[0]["top"])])
    for a, b in zip(left, left[1:]):
        c.down(a, b)
    c.arrow([(left[5]["r"], left[5]["mid"]), (right[0]["l"], right[0]["mid"])])
    for a, b in zip(right, right[1:]):
        c.arrow([(a["cx"], a["top"]), (b["cx"], b["bot"])])
    return c


if __name__ == "__main__":
    sop_builder.build(sys.modules[__name__])
