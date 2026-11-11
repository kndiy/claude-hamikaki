"""MDC-QT-06 — Phát triển nội dung & truyền thông thương hiệu.

Source: draft 'GMPc-QT-01.BM01_Quy trình phát triển nội dung & truyền thông (1).docx' (same folder).
Build: python scripts/qt06.py
"""
import sys

from flowchart import Chart
import sop_builder

CODE = "MDC-QT-06"
TITLE = "PHÁT TRIỂN NỘI DUNG & TRUYỀN THÔNG THƯƠNG HIỆU"
OUT = ("MDC-QT-06 QT phát triển nội dung & truyền thông", "MDC-QT-06 QT Phat trien noi dung va truyen thong.docx")

# ---------------------------------------------------------------- content
SIGN = [  # (chức vụ, họ tên, ngày)
    ("Chuyên viên Truyền thông & Thương hiệu", "Ngô Thị Kim Dung", "Ngày: 14/08/2026"),
    ("Trưởng phòng MD&C", "Đào Xuân Hưởng", None),  # changed by the user in Word, 2026-09-30
    ("Tổng giám đốc", "Đào Xuân Hưởng", None),
]

MUC_DICH = [
    "Quy trình này quy định các bước chuẩn hóa hoạt động sản xuất nội dung và truyền thông thương hiệu "
    "trên các kênh sở hữu và lan truyền của GMPc. Quy trình nhằm bảo đảm hoạt động truyền thông chạy "
    "thường xuyên, giữ vững thứ hạng SEO, duy trì tính nhất quán tuyệt đối với bộ nhận diện và tư tưởng "
    "cốt lõi của Bộ Tuyên ngôn V1.0, đồng thời hỗ trợ hiệu quả trên toàn bộ phễu khách hàng.",
]

PHAM_VI = [
    "Quy trình áp dụng cho Phòng MD&C, đặc biệt là Chuyên viên Truyền thông & Thương hiệu / Marketing, "
    "cùng các cá nhân, bộ phận liên quan trực tiếp đến việc cung cấp nội dung chuyên môn "
    "(8 AI Leader, Phòng RTA) tại GMPc Việt Nam.",
]

TRACH_NHIEM = [
    ("MD&C\n(Chuyên viên Truyền thông & Thương hiệu / Marketing)", [
        "Chịu trách nhiệm chính;",
        "Lập kế hoạch, trực tiếp sáng tạo nội dung;",
        "Quản lý thiết kế ấn phẩm;",
        "Quản trị các nền tảng (Website, LinkedIn);",
        "Tối ưu SEO và đo lường hiệu quả truyền thông.",
    ]),
    ("RTA", [
        "Cung cấp dữ liệu thực chiến, case study dự án;",
        "Kiểm duyệt tính chính xác về mặt kỹ thuật, quy chế GMP cho các nội dung chuyên ngành.",
    ]),
    ("COO", [
        "Kiểm tra, giám sát sự tuân thủ quy trình vận hành truyền thông, đảm bảo đồng bộ với hoạt động thực tế.",
    ]),
    ("CEO", [
        "Phê duyệt định hướng nội dung chiến lược, định vị thương hiệu và các thông điệp lớn trên truyền thông.",
    ]),
]

DINH_NGHIA = [
    ("MD&C", "Phòng Phát triển Thị trường & Truyền thông (Market Development & Communications)."),
    ("RTA", "Phòng Quy chế & Kỹ thuật."),
    ("COO", "Giám đốc Vận hành."),
    ("CEO", "Tổng Giám đốc."),
    ("Bộ Tuyên ngôn", "Tài liệu nền tảng L1, mã tài liệu GMPc-L1-TN-001, kim chỉ nam tư tưởng của GMPc 2026-2030."),
    ("Tuyên ngôn LEAD", "Khung định vị chiến lược 2026-2030 gồm 4 trụ cột: Leadership, Expertise, Alliance, Dedication."),
    ("Khẩu hiệu chính", "\"Good Consultant - Expert on Practices\", dùng cho truyền thông quốc tế, hồ sơ năng lực chính thức."),
    ("Khẩu hiệu đối thoại", "\"GMPc - Give More Profit for Customer\", dùng tiếp cận khách hàng nội địa, truyền thông thân mật."),
    ("AI (Gemini / Claude / NotebookLM)", "Hệ sinh thái công nghệ thông minh hỗ trợ truy xuất kho tri thức 15 năm và gia tăng tốc độ, chất lượng sản xuất nội dung."),
]

BUOC = [  # (số, tên bước, [nội dung], người thực hiện / ghi chú)
    ("1.", "Lập kế hoạch nội dung", [
        "Xây dựng ma trận nội dung theo 3 tầng phễu KH: Đầu phễu (Nhận diện), Giữa phễu (Cân nhắc), Cuối phễu (Chuyển đổi).",
        "Trình duyệt Content Calendar hàng tháng.",
    ], "MD&C (Chuyên viên Truyền thông & Thương hiệu).\nGhi chú: Kế hoạch phải chạy thường xuyên."),
    ("2.", "Sản xuất & Biên tập nội dung", [
        "Truy xuất kho tri thức 15 năm qua NotebookLM; sử dụng Gemini/Claude AI hỗ trợ lên dàn ý, viết dự thảo.",
        "Lựa chọn Khẩu hiệu phù hợp với đối tượng mục tiêu.",
        "Mọi đề xuất tư vấn trong nội dung phải có minh chứng từ dự án thực tế.",
        "Thiết kế ấn phẩm theo chuẩn Brand Guidelines.",
    ], "MD&C (Chuyên viên Truyền thông & Thương hiệu).\nHỗ trợ: AI Leader/RTA."),
    ("3.", "Thẩm định Nhất quán & Chuyên môn", [
        "Gửi nội dung chuyên ngành cho RTA kiểm duyệt tính chính xác về pháp lý, tiêu chuẩn GMP.",
        "Đảm bảo nội dung, hình ảnh nhất quán với định hướng của Bộ Tuyên ngôn V1.0.",
    ], "MD&C (Chuyên viên Truyền thông & Thương hiệu) và RTA."),
    ("4.", "Phân phối & Quản trị kênh", [
        "Quản trị 2 Website: gmp.com.vn (bảo vệ SEO) và gmpc.vn (định hướng mới).",
        "Mạng xã hội (LinkedIn): Lan tỏa bài viết chuyên gia, tương tác cộng đồng B2B.",
        "Ấn phẩm chuyên ngành: Gửi bài PR lên các kênh uy tín.",
    ], "MD&C (Chuyên viên Truyền thông & Thương hiệu)."),
    ("5.", "Đo lường & Tối ưu", [
        "Theo dõi KPI: Traffic Website, thứ hạng SEO, tỷ lệ tương tác LinkedIn, số lượng Lead.",
        "Báo cáo định kỳ trên Base Workflow / Google Workspace.",
        "Tinh chỉnh kế hoạch nội dung tiếp theo dựa trên hiệu quả thực tế.",
    ], "MD&C (Chuyên viên Truyền thông & Thương hiệu).\nBáo cáo cho: COO & CEO. Kết thúc QT-06."),
]

TAI_LIEU = [
    ("1", "Bộ Tuyên ngôn Bản đầy đủ V1.0 - Kim chỉ nam tư tưởng của GMPc 2026 – 2030", "GMPc-L1-TN-001"),
    ("2", "Kế hoạch Content Marketing & SEO Tháng/Quý", "GMPc-L3-MKT-001"),
    ("3", "Hướng dẫn ứng dụng hệ sinh thái AI trong sản xuất nội dung", "GMPc-L3-AI-002"),
    ("4", "Bảng Tiêu chuẩn Nhận diện Thương hiệu trên nền tảng số", "GMPc-L3-MD&C-003"),
]

LICH_SU = [
    ("__/__/____", "01", "Ban hành lần đầu",
     "Ban hành quy trình MDC-QT-06 — Phát triển nội dung & truyền thông thương hiệu."),
]

LUU = [
    "Phòng Hành chính Nhân sự / Tủ lưu trữ quy trình công ty.",
    "Thư mục Google Drive lưu trữ hệ thống L2 của Phòng MD&C (Google Workspace).",
    "Luồng quản lý quy trình Base Workflow của công ty.",
]


# ---------------------------------------------------------------- sơ đồ khối
def flowchart():
    c = Chart(1000, 1010)
    cx = 470
    t = c.box(cx, 20, 560, 70, [CODE, TITLE], "title", size=18)
    b1 = c.box(cx, 135, 400, 78, ["B1. Lập kế hoạch nội dung", "Ma trận 3 tầng phễu KH",
                                  "Trình duyệt Content Calendar hàng tháng"])
    b2 = c.box(cx, 263, 400, 78, ["B2. Sản xuất & biên tập nội dung", "NotebookLM · Gemini/Claude AI",
                                  "Thiết kế theo Brand Guidelines"])
    n2 = c.box(835, 280, 250, 44, ["Hỗ trợ: AI Leader / RTA"], "note", size=15)
    b3 = c.box(cx, 391, 400, 78, ["B3. Thẩm định nhất quán & chuyên môn",
                                  "RTA kiểm duyệt pháp lý, tiêu chuẩn GMP", "Đối chiếu Bộ Tuyên ngôn V1.0"])
    b4 = c.box(cx, 519, 400, 56, ["B4. Phân phối & quản trị kênh"])
    ch = [c.box(x, 625, 250, 66, lines, "io", size=15) for x, lines in [
        (190, ["Website", "gmp.com.vn · gmpc.vn"]),
        (cx, ["LinkedIn", "Bài viết chuyên gia, B2B"]),
        (750, ["Ấn phẩm chuyên ngành", "Bài PR trên kênh uy tín"])]]
    b5 = c.box(cx, 741, 400, 78, ["B5. Đo lường & tối ưu", "KPI: Traffic, thứ hạng SEO,",
                                  "tương tác LinkedIn, số lượng Lead"])
    n5 = c.box(835, 752, 250, 56, ["Báo cáo COO & CEO", "(Base Workflow / Google Workspace)"], "note", size=14)
    end = c.box(cx, 900, 400, 60, ["Kết thúc chu kỳ", "Tinh chỉnh kế hoạch nội dung kỳ tiếp theo"],
                "title", bold_first=True, size=16)

    for a, b in [(t, b1), (b1, b2), (b2, b3), (b3, b4), (b5, end)]:
        c.down(a, b)
    for x in ch:  # B4 fans out to the 3 channels, which merge into B5
        if x["cx"] == cx:
            c.arrow([(cx, b4["bot"]), (cx, x["top"])])
            c.arrow([(cx, x["bot"]), (cx, b5["top"])])
        else:
            c.arrow([(cx, b4["bot"]), (cx, b4["bot"] + 22), (x["cx"], b4["bot"] + 22), (x["cx"], x["top"])])
            c.arrow([(x["cx"], x["bot"]), (x["cx"], x["bot"] + 25), (cx, x["bot"] + 25), (cx, b5["top"])])
    for b, n in [(b2, n2), (b5, n5)]:
        c.link([(b["r"], b["mid"]), (n["l"], n["mid"])])
    x = 40  # end of cycle -> next plan (B1)
    c.arrow([(end["l"], end["mid"]), (x, end["mid"]), (x, b1["mid"]), (b1["l"], b1["mid"])], dashed=True)
    ym = (end["mid"] + b1["mid"]) / 2
    c.label(x + 24, ym, "Kỳ tiếp theo: quay lại B1", rotate=-90)
    return c


if __name__ == "__main__":
    sop_builder.build(sys.modules[__name__])
