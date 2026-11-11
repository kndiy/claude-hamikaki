"""GMPc-HD-01 Hướng dẫn sử dụng GWS (Gmail, Drive, Meet).

Source: source/GMPc-HD-01 … .docx (user's version of 2026-10-16 10:47). Content kept; only wording made
consistent ("Quy trình này" -> "Hướng dẫn này", IT & GWS -> Bộ phận IT, Hành chính nhân sự -> Phòng HCNS)
and the effective date cleared until approval. Points in docs/2026-10-16_Nhan-xet_HD-01-03.md are NOT applied.
"""
import os
import sop_builder

CODE = "GMPc-HD-01"
TITLE = "HƯỚNG DẪN SỬ DỤNG GWS (GMAIL, DRIVE, MEET)"
KIND = "HƯỚNG DẪN"
OUT = "GMPc-HD-01 Huong dan su dung GWS Gmail Drive Meet.docx"

SIGN = [("Nhân viên IT", "Lê Quang Huy", "8/7/2026"),
        ("Giám đốc điều hành", "Nguyễn Thế Giang", None),
        ("Tổng giám đốc", "Đào Xuân Hưởng", None)]

MUC_DICH = [
    "Hướng dẫn này quy định các bước, nguyên tắc và hướng dẫn chi tiết nhằm đảm bảo toàn bộ cán bộ nhân viên "
    "(CBNV) trong Công ty sử dụng hệ sinh thái Google Workspace (đặc biệt là Gmail, Google Drive và Google Meet) "
    "một cách thống nhất, an toàn, bảo mật và đạt hiệu quả cao nhất trong công việc.",
]
PHAM_VI = [
    "Hướng dẫn áp dụng cho tất cả các phòng ban, cá nhân đang làm việc tại Công ty được cấp phát tài khoản "
    "email theo tên miền của Công ty.",
]

TRACH_NHIEM = [
    ("Bộ phận IT", [
        "Quản lý, cấp phát, thu hồi tài khoản Google Workspace.",
        "Hướng dẫn, hỗ trợ kỹ thuật và giải quyết sự cố cho người dùng.",
        "Giám sát việc tuân thủ bảo mật dữ liệu.",
    ]),
    ("Phòng HCNS", [
        "Thông báo cho Bộ phận IT về nhân sự mới để cấp tài khoản và nhân sự nghỉ việc để thu hồi tài khoản.",
    ]),
    ("Nhân viên Công ty", [
        "Bảo mật tài khoản cá nhân, sử dụng đúng mục đích công việc, tuân thủ các hướng dẫn và quy định về lưu "
        "trữ, chia sẻ tài liệu trong hướng dẫn này.",
    ]),
]

DINH_NGHIA = [
    ("Google Workspace", "Bộ công cụ làm việc và cộng tác trực tuyến toàn diện do Google phát triển dành cho "
                         "doanh nghiệp."),
    ("Gmail", "Dịch vụ thư điện tử của Google, dùng để trao đổi thông tin trong và ngoài Công ty."),
    ("Google Drive", "Dịch vụ lưu trữ, đồng bộ hóa và chia sẻ tập tin dựa trên nền tảng đám mây."),
    ("Google Meet", "Dịch vụ liên lạc qua video trực tuyến để tổ chức các cuộc họp, hội thảo từ xa."),
]

FLOW_IMAGE = os.path.join(sop_builder.ROOT, "assets", "HD-01_So_do_khoi.png")  # from the source docx

BUOC = [
    ("1", "Cấp phát và đăng nhập tài khoản", [
        "Bộ phận IT nhận yêu cầu từ Phòng HCNS, tạo tài khoản thư điện tử (VD: anh.nguyen@gmpc.vn) và cung cấp "
        "mật khẩu mặc định.",
        "Người dùng đăng nhập lần đầu tại __gmail.com__, bắt buộc thay đổi mật khẩu và thiết lập bảo mật 2 lớp.",
    ], ["Bộ phận IT", "Người dùng mới"]),
    ("2", "Hướng dẫn sử dụng Gmail", [
        "**Chữ ký email:** Phải cài đặt chữ ký theo đúng tiêu chuẩn nhận diện thương hiệu của Công ty.",
        "**Tiêu đề (Subject):** Cần rõ ràng, tóm tắt được nội dung email.",
        "**Phân loại & gắn thẻ (Labels):** Khuyến khích tạo các nhãn (Labels) để phân loại email theo đối tác, "
        "dự án, mức độ ưu tiên.",
        "**Bảo mật:** Không bấm vào các liên kết lạ, không tải tệp đính kèm đáng ngờ.",
    ], ["Bộ phận IT", "Nhân viên Công ty"]),
    ("3", "Hướng dẫn sử dụng Google Drive", [
        "**Drive của tôi (My Drive):** Lưu trữ tài liệu cá nhân, dữ liệu công việc đang xử lý.",
        "**Bộ nhớ dùng chung (Shared Drives):** Nơi lưu trữ dữ liệu chung của phòng ban/dự án. Trưởng phòng có "
        "quyền quản lý thư mục này.",
        "**Quy tắc chia sẻ:** Chỉ cấp quyền tương ứng với nhu cầu (Người xem/Người nhận xét/Người chỉnh sửa). "
        "Hạn chế tối đa việc bật chia sẻ liên kết \"Bất kỳ ai có đường liên kết\" đối với các tài liệu nội bộ, "
        "bảo mật.",
    ], ["Bộ phận IT", "Nhân viên Công ty"]),
    ("4", "Hướng dẫn sử dụng Google Meet", [
        "**Tạo lịch họp:** Dùng Google Calendar tạo sự kiện, thêm khách mời và hệ thống sẽ tự sinh đường liên "
        "kết Google Meet.",
        "**Tham gia họp:** Tham gia đúng giờ. Tắt micro (Mute) khi không phát biểu để tránh tạp âm.",
        "**Trình bày (Presenting):** Sử dụng tính năng \"Trình bày ngay\" (Share screen) để chia sẻ tài liệu, "
        "trang tính trong quá trình thảo luận. Bật camera (nếu được yêu cầu).",
    ], ["Bộ phận IT", "Nhân viên Công ty"]),
    ("5", "Xử lý sự cố và thu hồi tài khoản", [
        "Khi quên mật khẩu hoặc gặp lỗi kỹ thuật, liên hệ ngay Bộ phận IT để được hỗ trợ đặt lại mật khẩu.",
        "Khi nhân sự nghỉ việc, Phòng HCNS thông báo, Bộ phận IT tiến hành đổi mật khẩu, sao lưu dữ liệu (nếu "
        "cần) và khóa hoặc xóa tài khoản theo quy định Công ty.",
    ], ["Bộ phận IT", "Phòng HCNS", "Nhân viên Công ty"]),
]

TAI_LIEU = [("1", "Cẩm nang Quản trị Google Workspace cho người mới bắt đầu", ""),
            ("2", "NotebookLM \"Sổ tay công cụ mới GMPc 2026\"", "")]
LICH_SU = [("__/__/____", "01", "Ban hành lần đầu", ""), ("", "", "", ""), ("", "", "", "")]
LUU = ["", "", ""]

if __name__ == "__main__":
    sop_builder.build(__import__(__name__))
