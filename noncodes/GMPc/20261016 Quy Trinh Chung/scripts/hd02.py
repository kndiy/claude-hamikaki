"""GMPc-HD-02 Hướng dẫn sử dụng Base (WeWork / Request / Square).

Source: source/GMPc-HD-02 … .docx. Content kept; fixed: "phom mẫu" -> "biểu mẫu", Wework -> WeWork,
stray "Ban Giám đốc / / Trưởng phòng", IT/HCNS names made consistent, "Quy trình này" -> "Hướng dẫn này",
empty definition rows dropped, effective date cleared until approval.
Points in docs/2026-10-16_Nhan-xet_HD-01-03.md (SLA figure, "Xin nghỉ phép" vs QĐ-01) are NOT applied.
"""
import sop_builder

CODE = "GMPc-HD-02"
TITLE = "HƯỚNG DẪN SỬ DỤNG BASE (WEWORK/REQUEST/SQUARE)"
KIND = "HƯỚNG DẪN"
OUT = "GMPc-HD-02 Huong dan su dung Base WeWork Request Square.docx"

SIGN = [("Nhân viên IT", "Lê Quang Huy", None),
        ("Giám đốc điều hành", "Nguyễn Thế Giang", None),
        ("Tổng giám đốc", "Đào Xuân Hưởng", None)]

MUC_DICH = [
    "Hướng dẫn này quy định các bước chuẩn hóa và nguyên tắc sử dụng nền tảng quản trị Base nội bộ, nhằm tối "
    "ưu hóa việc quản lý tiến độ công việc (WeWork), xét duyệt đề xuất nhanh chóng (Request) và kết nối hỏi đáp "
    "nội bộ hiệu quả (Square) trong toàn Công ty.",
]
PHAM_VI = ["Hướng dẫn này áp dụng bắt buộc đối với toàn bộ nhân viên."]

TRACH_NHIEM = [
    ("Bộ phận IT", [
        "Cấp phát, phân quyền và khóa tài khoản theo luân chuyển nhân sự.",
        "Thiết lập, cấu hình biểu mẫu đề xuất và luồng phê duyệt.",
        "Hỗ trợ kỹ thuật, xử lý lỗi hệ thống cho người dùng.",
    ]),
    ("Ban Giám đốc / Trưởng phòng", [
        "Giám sát, đánh giá hiệu suất công việc của nhân viên thông qua WeWork.",
        "Tiếp nhận, xử lý và ra quyết định phê duyệt các đề xuất trên Request đúng hạn.",
    ]),
    ("Nhân viên Công ty", [
        "Cập nhật tiến độ công việc hằng ngày, báo cáo đúng hạn trên WeWork.",
        "Sử dụng Request để tạo đề xuất đúng biểu mẫu khi có nhu cầu.",
        "Theo dõi, tương tác với câu hỏi, thảo luận trên Square.",
    ]),
]

DINH_NGHIA = [
    ("Base", "Nền tảng phần mềm quản trị doanh nghiệp toàn diện bao gồm nhiều ứng dụng chuyên biệt."),
    ("Base WeWork", "Ứng dụng quản lý công việc, dự án, phân công và giám sát tiến độ công việc nội bộ."),
    ("Base Request", "Ứng dụng xử lý luồng đề xuất, phê duyệt, xin ý kiến trực tuyến thay thế văn bản giấy."),
    ("Base Square", "Mạng xã hội nội bộ và kho tri thức số dành cho doanh nghiệp."),
]

BUOC = [
    ("1", "Kích hoạt tài khoản và cài đặt hồ sơ (profile)", [
        "Người dùng nhận email kích hoạt tài khoản hệ thống từ Bộ phận IT.",
        "Thực hiện đăng nhập, đổi mật khẩu bảo mật cá nhân.",
        "Cập nhật đầy đủ ảnh chân dung, họ tên, phòng ban và số điện thoại liên hệ để nhận diện trên hệ thống.",
    ], ["Người dùng mới"]),
    ("2", "Sử dụng Base Square (hỏi đáp nội bộ)", [
        "CBNV có trách nhiệm truy cập ứng dụng hằng ngày để cập nhật các thông tin, câu hỏi, dữ liệu từ nhân sự "
        "Công ty.",
        "Thực hiện tương tác (like, comment, share) văn minh, lịch sự, đúng văn hóa doanh nghiệp.",
    ], ["Toàn thể CBNV"]),
    ("3", "Sử dụng Base WeWork (quản lý công việc và dự án)", [
        "**Tạo việc:** Trưởng phòng/Quản lý dự án tạo công việc, mô tả chi tiết yêu cầu, gán người thực hiện và "
        "đặt thời hạn.",
        "**Thực hiện và báo cáo:** Người nhận việc cập nhật trạng thái công việc, đính kèm kết quả công việc cụ "
        "thể.",
    ], ["Quản lý bộ phận", "Nhân viên thực hiện"]),
    ("4", "Sử dụng Base Request (tạo và phê duyệt đề xuất)", [
        "**Tạo yêu cầu:** Nhân viên chọn đúng nhóm đề xuất cần gửi (ví dụ: __Đề xuất mua sắm__, __Xin nghỉ "
        "phép__, __Tạm ứng tiền__...), điền đầy đủ thông tin theo biểu mẫu bắt buộc và ấn __Gửi__.",
        "**Duyệt yêu cầu:** Người có thẩm quyền phê duyệt kiểm tra thông tin, thực hiện chọn __Chấp thuận__, "
        "__Từ chối__ hoặc __Chuyển tiếp__ trong khung giờ cam kết (SLA). Trường hợp từ chối bắt buộc phải ghi "
        "rõ lý do.",
    ], ["Người tạo yêu cầu", "Cấp quản lý phê duyệt"]),
    ("5", "Hỗ trợ vận hành và giải quyết sự cố", [
        "Trong quá trình vận hành, nếu gặp lỗi không đăng nhập được, lỗi tính năng hoặc lỗi luồng duyệt sai biểu "
        "mẫu quy định, người dùng gửi yêu cầu hỗ trợ tới Bộ phận IT.",
        "Bộ phận IT tiếp nhận và xử lý sự cố trong vòng tối đa 2 giờ làm việc.",
    ], ["Toàn thể CBNV", "Bộ phận IT"]),
    ("6", "Thu hồi tài khoản (khi nhân sự nghỉ việc)", [
        "Phòng HCNS gửi danh sách nhân sự nghỉ việc cho Bộ phận IT.",
        "Bộ phận IT thực hiện khóa tài khoản, chuyển giao toàn bộ dữ liệu dự án trên WeWork và lịch sử Request "
        "của nhân sự đó sang cho quản lý trực tiếp quản lý.",
    ], ["Phòng HCNS", "Bộ phận IT"]),
]

TAI_LIEU = [("1", "Tài liệu hướng dẫn sử dụng Base WeWork", ""),
            ("2", "Tài liệu hướng dẫn sử dụng Base Request", ""),
            ("3", "Tài liệu hướng dẫn sử dụng Base Square", "")]
LICH_SU = [("__/__/____", "01", "Ban hành lần đầu", ""), ("", "", "", ""), ("", "", "", "")]
LUU = ["", "", ""]

if __name__ == "__main__":
    sop_builder.build(__import__(__name__))
