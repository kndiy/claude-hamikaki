"""GMPc-HD-03 Hướng dẫn sử dụng Claude AI & NotebookLM.

Source: source/GMPc-HD-03 … .docx. Content kept; fixed: "có thế mạnh mạnh mẽ" -> "có thế mạnh",
"Mô hình trí tuệ nhân tạo lớn (LLM)" -> "Mô hình ngôn ngữ lớn (LLM)", Nhân viên IT -> Bộ phận IT,
"Quy trình này" -> "Hướng dẫn này", effective date cleared until approval.
The data-security points in docs/2026-10-16_Nhan-xet_HD-01-03.md are NOT applied (waiting for the drafter).
"""
import sop_builder

CODE = "GMPc-HD-03"
TITLE = "HƯỚNG DẪN SỬ DỤNG CLAUDE AI & NOTEBOOKLM"
KIND = "HƯỚNG DẪN"
OUT = "GMPc-HD-03 Huong dan su dung Claude AI NotebookLM.docx"

SIGN = [("Nhân viên IT", "Lê Quang Huy", None),
        ("Giám đốc điều hành", "Nguyễn Thế Giang", None),
        ("Tổng giám đốc", "Đào Xuân Hưởng", None)]

MUC_DICH = [
    "Hướng dẫn này quy định các nguyên tắc, phương pháp thao tác và giới hạn bảo mật dữ liệu khi ứng dụng hai "
    "trợ lý ảo AI (Claude AI và NotebookLM) vào công việc, nhằm nâng cao hiệu suất viết báo cáo, nghiên cứu tài "
    "liệu kỹ thuật, đồng thời ngăn chặn triệt để nguy cơ rò rỉ thông tin mật của doanh nghiệp và khách hàng.",
]
PHAM_VI = ["Hướng dẫn áp dụng cho toàn bộ nhân viên có nhu cầu sử dụng AI để xử lý thông tin."]

TRACH_NHIEM = [
    ("Bộ phận IT", [
        "Cập nhật xu hướng công nghệ, tổ chức các buổi đào tạo kỹ năng viết câu lệnh (Prompt Engineering).",
        "Cấp phát tài khoản dùng chung bản quyền (nếu có).",
    ]),
    ("Trưởng phòng / Bộ phận", [
        "Kiểm soát đầu ra sản phẩm của nhân viên khi có ứng dụng AI (đảm bảo tính chính xác, không sao chép "
        "nguyên bản).",
    ]),
    ("Nhân viên Công ty", [
        "Nghiêm túc thực hiện các quy định về bảo mật dữ liệu khi tải tài liệu lên AI.",
        "Phải tự kiểm tra chéo thông tin trước khi áp dụng vào thực tế, chịu trách nhiệm hoàn toàn về nội dung "
        "do mình xuất bản.",
    ]),
]

DINH_NGHIA = [
    ("Claude AI", "Mô hình ngôn ngữ lớn (LLM) của Anthropic, có thế mạnh về phân tích dữ liệu, xử lý ngôn ngữ "
                  "tự nhiên, lập luận logic, viết code và tóm tắt văn bản dài."),
    ("NotebookLM", "Trợ lý nghiên cứu ảo dựa trên AI của Google, hoạt động như một cuốn sổ tay thông minh dựa "
                   "trên chính các tài liệu nguồn (PDF, Doc, Link) do người dùng cung cấp."),
    ("Prompt (Câu lệnh)", "Các đoạn văn bản yêu cầu, câu hỏi hoặc ngữ cảnh mà người dùng nhập vào ô chat để "
                          "điều khiển AI làm việc theo ý muốn."),
    ("Hallucination (Ảo giác AI)", "Hiện tượng AI tự bịa đặt ra thông tin trông có vẻ hợp lý nhưng thực tế "
                                   "hoàn toàn sai lệch hoặc không có thật."),
]

BUOC = [
    ("1", "Đăng nhập và thiết lập bảo mật dữ liệu", [
        "Người dùng truy cập __claude.ai__ và __notebooklm.google.com__ bằng email Công ty được cấp.",
        "**Quy tắc bảo mật tối cao:** Tuyệt đối không dán, tải lên (upload) các thông tin thô mang tính nhạy "
        "cảm như: __báo cáo tài chính chưa công bố, danh sách số điện thoại khách hàng, bí mật công nghệ độc "
        "quyền, mật khẩu hệ thống__.",
        "Tắt tính năng \"Cho phép AI học từ dữ liệu của bạn\" (Data Training) trong phần cài đặt nếu hệ thống "
        "yêu cầu.",
    ], ["Tất cả nhân viên"]),
    ("2", "Phân loại tác vụ để chọn công cụ phù hợp", [
        "Chọn **Claude AI**: khi cần viết báo cáo, lập luận logic, dịch thuật kỹ thuật, sáng tạo nội dung "
        "truyền thông, tối ưu hóa bảng tính Excel hoặc viết mã code tự động.",
        "Chọn **NotebookLM**: khi cần nghiên cứu một tập hợp nhiều tài liệu lớn (quy chuẩn xây dựng, sách "
        "hướng dẫn vận hành, nhiều tài liệu PDF nội bộ) và chỉ muốn AI trả lời dựa trên những nguồn tài liệu "
        "đó (không lấy kiến thức bên ngoài).",
    ], ["Tất cả nhân viên"]),
    ("3", "Thao tác ứng dụng trên Claude AI", [
        "Áp dụng cấu trúc Prompt chuẩn: [Bối cảnh/Vai trò] + [Dữ liệu đầu vào] + [Nhiệm vụ cụ thể] + [Tiêu "
        "chí/Định dạng đầu ra].",
        "Tận dụng tính năng __Artifacts__ của Claude để hiển thị các đoạn mã hoặc văn bản dài tách biệt với ô "
        "chat chính, giúp dễ chỉnh sửa.",
        "Nếu câu trả lời chưa đạt yêu cầu, tiếp tục ra lệnh tinh chỉnh từng phần.",
    ], ["Người sử dụng trực tiếp"]),
    ("4", "Thao tác ứng dụng trên NotebookLM", [
        "Tạo một \"Notebook\" mới đặt tên theo dự án.",
        "Tải các nguồn tài liệu lên (hỗ trợ Google Docs, Slides, file PDF hoặc văn bản sao chép).",
        "Sử dụng tính năng câu hỏi gợi ý tự động hoặc chat để tìm liên kết giữa các tài liệu.",
        "Có thể sử dụng tính năng __Audio Overview__ (nếu cần nghe tóm tắt tài liệu dạng podcast bằng tiếng "
        "Anh để nắm ý chính).",
    ], ["Người sử dụng trực tiếp"]),
    ("5", "Kiểm tra chéo thông tin (Fact-checking)", [
        "**Bắt buộc:** Do AI có thể gặp hiện tượng \"ảo giác\", toàn bộ số liệu kỹ thuật, điều khoản luật, công "
        "thức toán học do Claude cung cấp đều phải được kiểm tra lại bằng nguồn chính thống.",
        "Đối với NotebookLM, nhấp vào các số trích dẫn để đối chiếu chính xác thông tin nằm ở trang nào của tài "
        "liệu gốc.",
    ], ["Người sử dụng trực tiếp", "Cấp trên kiểm tra lại"]),
    ("6", "Nghiệm thu văn bản và lưu trữ nội bộ", [
        "Biên tập lại văn bản từ AI theo đúng văn phong, ngôn từ chuẩn mực của GMPc Việt Nam.",
        "Không sao chép 100% văn bản một cách máy móc. Lưu trữ file kết quả cuối cùng lên Google Drive Công ty "
        "theo quy định.",
    ], ["Người sử dụng trực tiếp"]),
]

TAI_LIEU = []  # none in the source; template blank rows kept
LICH_SU = [("__/__/____", "01", "Ban hành lần đầu", ""), ("", "", "", ""), ("", "", "", "")]
LUU = ["", "", ""]

if __name__ == "__main__":
    sop_builder.build(__import__(__name__))
