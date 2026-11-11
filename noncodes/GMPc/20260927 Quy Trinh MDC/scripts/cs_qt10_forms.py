"""Word forms of CS-QT-10: BM01 (biên bản bàn giao RTA → CS) and BM03 (phiếu khảo sát mức độ hài lòng).

Content follows CS-QT-10 steps 1 and 4 (scripts/cs_qt10.py). The satisfaction criteria and the 1–5 scale
are Claude's proposal. Build: python scripts/cs_qt10_forms.py
"""
import os

from form_builder import (BOX, DOTS, TEXT_W, b, build_form, check, h1, i, p, signatures, table)
import cs_qt10

DIR = os.path.join(cs_qt10.OUT_BASE, cs_qt10.OUT[0])
FULL = "…" * 54  # one full line of dots


def line(label, dots=DOTS):
    return p(label + " " + dots)


# ---------------------------------------------------------------- BM01
def bm01():
    w_no, w_doc, w_st, w_loc = 600, 3757, 2700, TEXT_W - 600 - 3757 - 2700
    status = f"{BOX} Có   {BOX} Không   {BOX} KAD"
    docs = [
        "Hợp đồng và các phụ lục hợp đồng",
        "Biên bản nghiệm thu cuối cùng / biên bản thanh lý hợp đồng",
        "Chứng nhận, giấy phép khách hàng đã đạt (bản sao)",
        "Thông tin đầu mối liên hệ phía khách hàng (mục III)",
        "Lịch tái đánh giá / duy trì chứng nhận (mục II)",
        "Các cam kết sau hợp đồng: bảo hành, hỗ trợ, hạng mục còn theo dõi (nếu có)",
        "Ghi chú của RTA về khách hàng: đặc điểm, lưu ý khi làm việc",
        "Hồ sơ khác: " + "…" * 20,
    ]
    return [
        p(i("Căn cứ GMPc-CN-01 và quy trình CS-QT-10, hôm nay, ngày ……/……/20……, tại "), i("…" * 25),
          i(", Phòng RTA bàn giao hồ sơ khách hàng sau khi hoàn tất hợp đồng cho Phòng CS như sau:")),
        h1("I. THÔNG TIN CHUNG"),
        table([3300, TEXT_W - 3300], [
            ["Tên khách hàng", ""],
            ["Mã khách hàng", ""],
            ["Số / ngày hợp đồng", ""],
            ["Dịch vụ / phạm vi đã thực hiện", ""],
            ["Ngày hoàn tất HĐ / nghiệm thu cuối cùng", ""],
            ["Người bàn giao (RTA)", ""],
            ["Người tiếp nhận (CS)", ""],
        ], header=False),
        h1("II. CHỨNG NHẬN / GIẤY PHÉP KHÁCH HÀNG ĐÃ ĐẠT"),
        table([2400, 1700, 1500, 1500, TEXT_W - 7100], [
            ["Chứng nhận / tiêu chuẩn", "Số chứng nhận", "Ngày cấp", "Hiệu lực đến", "Mốc tái đánh giá / duy trì gần nhất"],
            ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""],
        ]),
        h1("III. ĐẦU MỐI LIÊN HỆ PHÍA KHÁCH HÀNG"),
        table([2700, 2000, 1800, TEXT_W - 6500], [
            ["Họ và tên", "Chức vụ", "Điện thoại", "Email"],
            ["", "", "", ""], ["", "", "", ""],
        ]),
        h1("IV. DANH MỤC HỒ SƠ BÀN GIAO"),
        table([w_no, w_doc, w_st, w_loc],
              [["STT", "Hồ sơ", "Tình trạng", "Nơi lưu / đường dẫn"]]
              + [[str(n + 1), d, status, ""] for n, d in enumerate(docs)],
              center_cols=(0,)),
        p(i("KAD: không áp dụng cho dự án này."), size=20),
        h1("V. KẾT QUẢ KIỂM TRA CỦA CS"),
        p(i("(CS kiểm tra danh mục hồ sơ trong 03 ngày làm việc kể từ ngày nhận)"), size=20),
        check("Hồ sơ đầy đủ — CS tiếp nhận; trách nhiệm chăm sóc khách hàng chuyển sang Phòng CS."),
        check("Hồ sơ chưa đầy đủ — đề nghị RTA bổ sung các hồ sơ sau:"),
        p(FULL[:-4], indent=567), p(FULL[:-4], indent=567),
        p("Thời hạn bổ sung: ……/……/20……", indent=567),
        p(""),
        signatures(("ĐẠI DIỆN PHÒNG RTA", "(Bên bàn giao)"), ("ĐẠI DIỆN PHÒNG CS", "(Bên tiếp nhận)")),
    ]


# ---------------------------------------------------------------- BM03
CRITERIA = [
    "Năng lực chuyên môn của đội ngũ tư vấn",
    "Chất lượng tài liệu, hồ sơ, sản phẩm bàn giao",
    "Tiến độ thực hiện so với cam kết",
    "Khả năng phản hồi và giải quyết vấn đề",
    "Thái độ phục vụ, giao tiếp",
    "Sự phối hợp giữa các bộ phận của GMPc",
    "Hiệu quả dịch vụ so với chi phí",
    "Mức độ hài lòng chung",
]


def bm03():
    w_no, w_sc = 600, 700
    w_c = TEXT_W - w_no - 5 * w_sc
    return [
        p("Kính gửi Quý Khách hàng,"),
        p("Trân trọng cảm ơn Quý Khách hàng đã tin tưởng và đồng hành cùng GMPc Việt Nam. Để không ngừng nâng cao "
          "chất lượng dịch vụ, chúng tôi rất mong nhận được đánh giá của Quý Khách hàng qua phiếu khảo sát dưới đây. "
          "Mọi thông tin sẽ được bảo mật và chỉ dùng để cải tiến dịch vụ."),
        h1("I. THÔNG TIN CHUNG"),
        table([3300, TEXT_W - 3300], [
            ["Tên khách hàng", ""],
            ["Người trả lời / chức vụ", ""],
            ["Dự án / dịch vụ đã sử dụng", ""],
            ["Đợt khảo sát", f"{BOX} Sau dự án      {BOX} Định kỳ hằng năm"],
            ["Ngày trả lời", "……/……/20……"],
        ], header=False),
        h1("II. ĐÁNH GIÁ MỨC ĐỘ HÀI LÒNG"),
        p(i("Vui lòng đánh dấu (X) vào ô phù hợp: 1 = Rất không hài lòng · 2 = Không hài lòng · 3 = Bình thường · "
            "4 = Hài lòng · 5 = Rất hài lòng"), size=21),
        table([w_no, w_c] + [w_sc] * 5,
              [["STT", "Tiêu chí", "1", "2", "3", "4", "5"]]
              + [[str(n + 1), c, "", "", "", "", ""] for n, c in enumerate(CRITERIA)],
              center_cols=(0, 2, 3, 4, 5, 6)),
        h1("III. KHẢ NĂNG GIỚI THIỆU"),
        p("Quý Khách hàng có sẵn sàng giới thiệu GMPc cho đối tác, doanh nghiệp khác không?"),
        check("Chắc chắn có", "Có thể", "Không"),
        h1("IV. NHU CẦU TRONG 12 THÁNG TỚI"),
        p(i("(có thể chọn nhiều)"), size=21),
        check("Tái đánh giá / duy trì chứng nhận"),
        check("Nâng cấp tiêu chuẩn (EU-GMP, PIC/S, US FDA…)"),
        check("Mở rộng / xây dựng nhà máy, kho mới"),
        check("Đào tạo nhân sự"),
        check("Khác: " + "…" * 30),
        check("Chưa có nhu cầu"),
        h1("V. GÓP Ý KHÁC"),
        p(FULL), p(FULL), p(FULL),
        check("Tôi đồng ý nhận thông tin tin tức ngành, quy định mới và sự kiện của GMPc.", indent=0),
        p(b("Trân trọng cảm ơn Quý Khách hàng!"), align="center", before=120),
        h1("PHẦN DÀNH CHO GMPc (Phòng CS ghi)"),
        table([3300, TEXT_W - 3300], [
            ["Ngày nhận phiếu / người nhận", ""],
            ["Điểm trung bình (8 tiêu chí)", ""],
            ["Có tiêu chí ≤ 2 điểm → xử lý theo Bước 7", f"{BOX} Có   {BOX} Không   Mã phiếu yêu cầu: ……………"],
            ["Có nhu cầu mới → chuyển MD&C theo Bước 8", f"{BOX} Có   {BOX} Không   Mã cơ hội: ……………"],
            ["Đã cập nhật vào CS-QT-10.BM02", f"{BOX} Rồi   Ngày: ……/……/20……"],
        ], header=False),
    ]


if __name__ == "__main__":
    build_form(os.path.join(DIR, "CS-QT-10.BM01 Bien ban ban giao ho so KH RTA-CS.docx"),
               ["BIÊN BẢN BÀN GIAO HỒ SƠ KHÁCH HÀNG", "(RTA → CS, SAU KHI HOÀN TẤT HỢP ĐỒNG)"],
               "CS-QT-10.BM01", bm01())
    build_form(os.path.join(DIR, "CS-QT-10.BM03 Phieu khao sat muc do hai long KH.docx"),
               ["PHIẾU KHẢO SÁT MỨC ĐỘ HÀI LÒNG KHÁCH HÀNG"], "CS-QT-10.BM03", bm03())
