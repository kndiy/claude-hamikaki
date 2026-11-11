"""CS-QT-10 — Chăm sóc khách hàng sau hợp đồng (Phòng Corporate Service).

New SOP proposed by Claude, not from a draft. Sources:
- GMPc-CN-01 (docs/260623_…): RTA→CS boundary "Hoàn tất HĐ", CS duties 1–5, rights 1–2, matrix "Giai đoạn 3".
- JD 'Corporate Service (CS)/260702 JD_ChuyenVien_HC_VPMN_CSKH.docx' part A (duties 1–11).
- 'Corporate Service (CS)/So_Do_To_Chuc_Phong_CS.pdf' (roles in CS).
- Code CS-QT-10 comes from the MD&C SOP catalogue v0.5 and MDC-QT-03 ("chăm sóc quan hệ sau bán: CS-QT-10").
Numbers marked "ĐỀ XUẤT" are Claude's proposals, not company rules yet.

Build: python scripts/cs_qt10.py
"""
import os
import sys

from flowchart import Chart
import sop_builder

CODE = "CS-QT-10"
TITLE = "CHĂM SÓC KHÁCH HÀNG SAU HỢP ĐỒNG"
OUT_BASE = os.path.join(sop_builder.ROOT, "Corporate Service (CS)")
OUT = ("CS-QT-10 QT Cham soc khach hang sau hop dong", "CS-QT-10 QT Cham soc khach hang sau hop dong.docx")

CV = "Chuyên viên CSKH"
TP = "Trưởng phòng CS"

# ---------------------------------------------------------------- content
SIGN = [  # filled in by the user in Word, 2026-09-30
    ("Chuyên viên Tư vấn", "Cao Thị Thu Hà", None),
    ("Trưởng phòng CS", "Ngô Thị Hồng Thắm", None),
    ("Tổng giám đốc", "Đào Xuân Hưởng", None),
]

MUC_DICH = [
    "Quy trình này quy định trình tự tiếp nhận khách hàng từ Phòng RTA sau khi hoàn tất hợp đồng, lập kế hoạch "
    "và thực hiện chăm sóc khách hàng dài hạn, theo dõi mức độ hài lòng, thời hạn tái đánh giá / duy trì chứng "
    "nhận, tiếp nhận phản hồi – khiếu nại và chuyển giao cơ hội kinh doanh mới cho Phòng MD&C.",
    "Mục tiêu: duy trì quan hệ lâu dài với khách hàng, để khách hàng cảm nhận được sự đồng hành của GMPc sau dự "
    "án, nâng cao chất lượng dịch vụ và phát triển doanh thu từ khách hàng hiện hữu.",
]

PHAM_VI = [
    "Áp dụng cho Phòng Corporate Service (CS), trong đó Chuyên viên Chăm sóc khách hàng là người thực hiện chính, "
    "và các phòng phối hợp (RTA, MD&C) đối với toàn bộ khách hàng của GMPc đã hoàn tất hợp đồng.",
    "Điểm bắt đầu: hợp đồng hoàn tất (hoàn tất nghĩa vụ hợp đồng hoặc ký nghiệm thu cuối cùng) và RTA bàn giao "
    "hồ sơ khách hàng cho CS.",
    "Điểm kết thúc: khách hàng ngừng hoạt động hoặc đề nghị ngừng liên hệ; hồ sơ khách hàng được đóng và lưu trữ.",
    "Quy trình không bao gồm:",
    "* Thanh toán, công nợ, hóa đơn (thực hiện theo quy định tài chính – kế toán của Phòng CS);",
    "* Tư vấn, xử lý chuyên môn kỹ thuật (do RTA thực hiện);",
    "* Tư vấn, chào giá, ký hợp đồng dịch vụ mới (do MD&C thực hiện theo các quy trình MDC-QT).",
]

TRACH_NHIEM = [
    (f"{CV}\n(Phòng CS)", [
        "Chủ trì toàn bộ quy trình: tiếp nhận hồ sơ bàn giao, lập và thực hiện kế hoạch chăm sóc;",
        "Khảo sát mức độ hài lòng sau dự án và hàng năm;",
        "Theo dõi thời hạn tái đánh giá / duy trì chứng nhận của khách hàng;",
        "Tiếp nhận, điều phối và theo dõi yêu cầu, phản hồi, khiếu nại đến khi hoàn tất;",
        "Ghi nhận và chuyển cơ hội kinh doanh mới cho MD&C;",
        "Cập nhật dữ liệu khách hàng và lập báo cáo định kỳ.",
    ]),
    (TP, [
        "Phê duyệt kế hoạch chăm sóc khách hàng;",
        "Giám sát việc thực hiện, xem xét báo cáo định kỳ;",
        "Xử lý khiếu nại vượt thẩm quyền của Chuyên viên CSKH; báo cáo COO/CEO khi khiếu nại nghiêm trọng.",
    ]),
    ("RTA", [
        "Bàn giao đầy đủ hồ sơ dự án, hồ sơ nghiệm thu cho CS khi hoàn tất hợp đồng; bổ sung khi CS yêu cầu;",
        "Phối hợp xử lý các yêu cầu, khiếu nại có nội dung chuyên môn;",
        "Phối hợp chuẩn bị kế hoạch hỗ trợ khách hàng trước thời hạn tái đánh giá;",
        "Khi khách hàng liên hệ trực tiếp RTA sau hợp đồng: xử lý và cập nhật cho CS.",
    ]),
    ("MD&C", [
        "Tiếp nhận cơ hội kinh doanh mới do CS chuyển giao và kích hoạt lại quy trình từ bước tư vấn ban đầu "
        "(MDC-QT-02);",
        "Phản hồi cho CS kết quả xử lý cơ hội.",
    ]),
    ("Kế toán (Phòng CS)", [
        "Phối hợp xử lý các yêu cầu, khiếu nại liên quan đến hóa đơn, thanh toán, công nợ.",
    ]),
    ("COO / CEO", [
        "Xem xét báo cáo chăm sóc khách hàng và mức độ hài lòng;",
        "Chỉ đạo xử lý khiếu nại nghiêm trọng.",
    ]),
]

DINH_NGHIA = [
    ("CS", "Phòng Dịch vụ Doanh nghiệp (Corporate Service)."),
    ("CSKH", "Chăm sóc khách hàng."),
    ("RTA", "Phòng Quy chế & Kỹ thuật (Regulatory & Technical Affairs)."),
    ("MD&C", "Phòng Phát triển Thị trường & Truyền thông (Market Development & Communications)."),
    ("KH", "Khách hàng."),
    ("Hoàn tất HĐ", "Thời điểm hoàn tất nghĩa vụ hợp đồng hoặc ký nghiệm thu cuối cùng; là ranh giới chuyển giao "
                    "trách nhiệm từ RTA sang CS (theo GMPc-CN-01)."),
    ("KH hiện hữu", "Khách hàng đã hoàn tất hợp đồng với GMPc và đang được CS chăm sóc."),
    ("Tái đánh giá / duy trì chứng nhận", "Đợt đánh giá lại định kỳ để khách hàng duy trì hoặc gia hạn chứng nhận "
                                          "đã đạt (GMP, GDP, GSP, ISO…)."),
    ("Khiếu nại", "Phản hồi của khách hàng về việc không hài lòng với sản phẩm, dịch vụ hoặc cách phục vụ của GMPc, "
                  "có yêu cầu GMPc xem xét, giải quyết."),
]

NGUYEN_TAC = [
    "Sau khi hoàn tất hợp đồng, CS là đầu mối chủ động duy nhất liên hệ khách hàng (theo GMPc-CN-01).",
    "Trách nhiệm chỉ chuyển từ RTA sang CS khi CS xác nhận hồ sơ bàn giao đầy đủ; CS có quyền từ chối tiếp nhận "
    "nếu danh mục hồ sơ chưa đầy đủ.",
    "CS không tư vấn chuyên môn và không báo giá: nội dung chuyên môn chuyển RTA; nhu cầu dịch vụ mới chuyển MD&C.",
    "Mọi lần liên hệ, phản hồi, khiếu nại và cơ hội đều được ghi nhận vào CS-QT-10.BM02 (một nguồn dữ liệu duy nhất).",
    "Thông tin khách hàng được bảo mật; chỉ sử dụng cho mục đích chăm sóc và phục vụ khách hàng.",
]

BUOC = [  # (số, tên bước, [nội dung], người thực hiện / ghi chú)
    ("1.", "Tiếp nhận bàn giao hồ sơ khách hàng từ RTA", [
        "Khi hoàn tất hợp đồng, RTA bàn giao hồ sơ khách hàng cho CS theo CS-QT-10.BM01, tối thiểu gồm: hợp đồng, "
        "biên bản nghiệm thu / thanh lý; chứng nhận, giấy phép khách hàng đã đạt (số, ngày cấp, thời hạn hiệu lực); "
        "đầu mối liên hệ phía khách hàng; lịch tái đánh giá / duy trì chứng nhận; các cam kết sau hợp đồng (nếu có); "
        "ghi chú của RTA về khách hàng.",
        "CS kiểm tra danh mục hồ sơ trong 03 ngày làm việc.",  # ĐỀ XUẤT
        "Hồ sơ chưa đầy đủ: CS yêu cầu RTA bổ sung hoặc từ chối tiếp nhận.",
        "Hồ sơ đầy đủ: hai bên ký xác nhận bàn giao; trách nhiệm chăm sóc chuyển sang CS.",
    ], f"{CV}.\nRTA bàn giao, bổ sung khi được yêu cầu."),
    ("2.", "Lập hồ sơ khách hàng hiện hữu", [
        "Cập nhật thông tin khách hàng, dự án, chứng nhận và lịch sử giao dịch vào CS-QT-10.BM02.",
        "Lưu hồ sơ bàn giao trên Google Drive của Phòng CS theo quy định lưu trữ của Công ty.",
    ], f"{CV}."),
    ("3.", "Lập kế hoạch chăm sóc khách hàng", [
        "Lập kế hoạch chăm sóc dài hạn cho từng khách hàng, theo loại dịch vụ đã thực hiện: kênh liên hệ, tần suất, "
        "các mốc (khảo sát sau dự án, liên hệ định kỳ, tri ân, mốc tái đánh giá).",
        "Trình Trưởng phòng CS phê duyệt.",
        "Rà soát, cập nhật kế hoạch tối thiểu hằng năm và khi có thay đổi.",
    ], f"{CV} lập.\n{TP} phê duyệt."),
    ("4.", "Khảo sát mức độ hài lòng sau dự án", [
        "Gửi Phiếu khảo sát CS-QT-10.BM03 cho khách hàng trong vòng 30 ngày sau khi tiếp nhận bàn giao.",  # ĐỀ XUẤT
        "Tổng hợp kết quả vào BM02.",
        "Điểm không hài lòng hoặc góp ý cụ thể được xử lý như một phản hồi theo Bước 7.",
    ], f"{CV}.\nKết quả thông báo cho MD&C và RTA."),
    ("5.", "Chăm sóc khách hàng định kỳ", [
        "Liên hệ khách hàng theo kế hoạch (điện thoại, email hoặc hình thức phù hợp), tối thiểu 01 lần/quý.",  # ĐỀ XUẤT
        "Cập nhật cho khách hàng tin tức ngành, quy định mới, sự kiện của GMPc.",
        "Phối hợp tổ chức hội thảo, chương trình tri ân khách hàng.",
        "Ghi nhận nội dung mỗi lần liên hệ vào nhật ký chăm sóc (BM02).",
        "Khảo sát mức độ hài lòng định kỳ hằng năm (BM03).",
    ], f"{CV}."),
    ("6.", "Theo dõi tái đánh giá / duy trì chứng nhận", [
        "Theo dõi thời hạn chứng nhận của từng khách hàng trên BM02.",
        "Tối thiểu 06 tháng trước thời hạn: thông báo RTA để chuẩn bị và nhắc khách hàng.",
        "Phối hợp RTA chuẩn bị kế hoạch hỗ trợ khách hàng trước thời hạn.",
        "Khách hàng có nhu cầu sử dụng dịch vụ hỗ trợ tái đánh giá: xử lý như cơ hội mới theo Bước 8.",
    ], f"{CV}.\nRTA phối hợp."),
    ("7.", "Tiếp nhận và xử lý yêu cầu, phản hồi, khiếu nại", [
        "Ghi nhận mọi yêu cầu, phản hồi, khiếu nại vào BM02; xác nhận đã tiếp nhận với khách hàng trong 01 ngày "
        "làm việc.",  # ĐỀ XUẤT
        "Điều phối đến bộ phận liên quan: RTA (chuyên môn), Kế toán (hóa đơn, thanh toán), MD&C (dịch vụ mới).",
        "Theo dõi việc xử lý đến khi hoàn tất; phản hồi kết quả cho khách hàng.",
        "Khiếu nại vượt thẩm quyền: báo Trưởng phòng CS; khiếu nại nghiêm trọng: báo COO/CEO.",
    ], f"{CV} điều phối.\nBộ phận liên quan xử lý."),
    ("8.", "Chuyển giao cơ hội kinh doanh mới cho MD&C", [
        "Ghi nhận nhu cầu đầu tư hoặc dịch vụ mới của khách hàng hiện hữu.",
        "Chuyển đầy đủ thông tin cho MD&C (BM02, sheet Cơ hội); MD&C kích hoạt lại quy trình từ bước tư vấn ban đầu "
        "(MDC-QT-02).",
        "Theo dõi và cập nhật kết quả xử lý cơ hội do MD&C phản hồi.",
    ], f"{CV}.\nMD&C tiếp nhận."),
    ("9.", "Báo cáo định kỳ và cải tiến", [
        "Lập báo cáo hằng quý: tình hình chăm sóc, mức độ hài lòng, yêu cầu – khiếu nại, cơ hội đã chuyển MD&C.",  # ĐỀ XUẤT
        "Báo cáo Trưởng phòng CS; tổng hợp báo cáo COO/CEO.",
        "Đề xuất giải pháp cải tiến chất lượng dịch vụ và quy trình chăm sóc khách hàng; cập nhật kế hoạch chăm sóc "
        "(quay lại Bước 3).",
    ], f"{CV} lập.\n{TP} xem xét. Kết thúc chu kỳ."),
]

TAI_LIEU = [
    ("1", "Biên bản bàn giao hồ sơ khách hàng (RTA → CS)", "CS-QT-10.BM01"),
    ("2", "Sổ quản lý chăm sóc khách hàng hiện hữu (danh sách KH & chứng nhận, kế hoạch & nhật ký chăm sóc, "
          "yêu cầu – khiếu nại, cơ hội chuyển MD&C)", "CS-QT-10.BM02"),
    ("3", "Phiếu khảo sát mức độ hài lòng khách hàng", "CS-QT-10.BM03"),
]

LICH_SU = [
    ("__/__/____", "01", "Ban hành lần đầu", "Ban hành quy trình CS-QT-10 — Chăm sóc khách hàng sau hợp đồng."),
]

LUU = [
    "Phòng Hành chính nhân sự",
    "Hệ thống quản lý tài liệu (Google Drive) — thư mục CS / Quy trình",
    "Hệ thống quản lý tài liệu (Google Drive) — thư mục CS / Quy trình",
]


# ---------------------------------------------------------------- sơ đồ khối
def flowchart():
    c = Chart(1000, 1230)
    cx, rx = 390, 800
    t = c.box(cx, 20, 520, 64, [CODE, TITLE], "title", size=18)
    src = c.box(cx, 112, 420, 56, ["Từ RTA: hoàn tất HĐ / nghiệm thu cuối cùng"], "io", size=15)
    b1 = c.box(cx, 200, 400, 70, ["B1. Tiếp nhận bàn giao hồ sơ", "Kiểm tra theo BM01"])
    d = c.diamond(cx, 305, 260, 100, ["Hồ sơ", "đầy đủ?"])
    fix = c.box(rx, 327, 300, 56, ["RTA bổ sung hồ sơ"], "note", size=15)
    b2 = c.box(cx, 445, 400, 70, ["B2. Lập hồ sơ KH hiện hữu", "Cập nhật BM02"])
    b3 = c.box(cx, 555, 400, 70, ["B3. Lập kế hoạch chăm sóc", "Trưởng phòng CS phê duyệt"])
    b4 = c.box(cx, 665, 400, 70, ["B4. Khảo sát hài lòng sau dự án", "Phiếu khảo sát BM03"])
    b5 = c.box(cx, 775, 400, 78, ["B5. Chăm sóc định kỳ", "Liên hệ, tin tức ngành, tri ân,", "khảo sát hằng năm"])
    b6 = c.box(rx, 760, 330, 70, ["B6. Theo dõi tái đánh giá", "Báo RTA & KH ≥ 6 tháng trước hạn"], size=15)
    b7 = c.box(rx, 870, 330, 70, ["B7. Yêu cầu / khiếu nại", "Điều phối đến khi hoàn tất"], size=15)
    b8 = c.box(rx, 980, 330, 70, ["B8. Cơ hội dịch vụ mới", "Chuyển MD&C"], size=15)
    mdc = c.box(rx, 1110, 330, 56, ["MD&C kích hoạt lại MDC-QT-02"], "io", size=15)
    b9 = c.box(cx, 1100, 400, 76, ["B9. Báo cáo định kỳ & cải tiến", "Trưởng phòng CS, COO/CEO"])

    for a, b in [(t, src), (src, b1), (b1, d), (b2, b3), (b3, b4), (b4, b5), (b5, b9), (b8, mdc)]:
        c.down(a, b)
    c.arrow([(cx, d["bot"]), (cx, b2["top"])])
    c.label(cx + 12, d["bot"] + 26, "Có", anchor="start")
    c.arrow([(d["r"], d["mid"]), (fix["l"], fix["mid"])])
    c.label((d["r"] + fix["l"]) / 2, d["mid"] - 8, "Không")
    c.arrow([(fix["cx"], fix["top"]), (fix["cx"], b1["mid"]), (b1["r"], b1["mid"])], dashed=True)
    # events that arise while caring for the customer
    bus = b5["r"] + 18
    for b in (b6, b7, b8):
        c.arrow([(b5["r"], b5["mid"]), (bus, b5["mid"]), (bus, b["mid"]), (b["l"], b["mid"])])
    # yearly loop back to the plan
    x = 60
    c.arrow([(b9["l"], b9["mid"]), (x, b9["mid"]), (x, b3["mid"]), (b3["l"], b3["mid"])], dashed=True)
    c.label(x + 24, (b9["mid"] + b3["mid"]) / 2, "Cập nhật kế hoạch hằng năm", rotate=-90)
    return c


if __name__ == "__main__":
    sop_builder.build(sys.modules[__name__])
