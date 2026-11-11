"""GMPc-QĐ-01 Nội quy lao động, laid out on the BM01 frame (header, sign-off table, history, storage).

Content comes from source/Nội Quy lao động.docx with spelling and layout fixed only; points flagged in
docs/2026-10-16_Nhan-xet_Noi-quy-lao-dong.md are kept as written.

Body markup (lists of strings):
  "* text"  bullet          "~ text"  sub-item (dash, indented)
  "# text"  bold line       other     justified paragraph
Run:  PYTHONIOENCODING=utf-8 python scripts/qd01.py [--force]
Needs Microsoft Word (COM) to fill the table of contents. Refuses to overwrite an output edited by hand.
"""
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, "templates", "GMPc-QT-01.BM01_Biểu mẫu trắng SOP.docx")
OUT = os.path.join(ROOT, "output", "GMPc-QĐ-01 Noi quy lao dong.docx")
STAMP = OUT + ".sha256"  # hash of the last build; a different hash means someone edited the file

CODE = "GMPc-QĐ-01"
TITLE = "NỘI QUY LAO ĐỘNG"
KIND = "QUY ĐỊNH"  # replaces "QUY TRÌNH THAO TÁC CHUẨN (SOP)" in the header
# (chức vụ, họ tên): soạn thảo, kiểm tra, phê duyệt. Entered by the user in Word on 2026-10-16.
SIGN = [("Chuyên viên HCNS", "Trần Thị Thanh Thảo"),
        ("Trưởng phòng CS", "Ngô Thị Hồng Thắm"),
        ("CEO", "Đào Xuân Hưởng")]

# ---------------------------------------------------------------- content
QD_CAN_CU = [
    "Căn cứ Bộ luật Lao động số 45/2019/QH14 ngày 20/11/2019;",
    "Căn cứ Điều lệ Công ty Cổ phần GMPc Việt Nam;",
]
QD_DIEU = [
    ("Điều 1.", "Ban hành kèm theo Quyết định này Nội quy lao động của Công ty Cổ phần GMPc Việt Nam."),
    ("Điều 2.", "Quyết định này có hiệu lực kể từ ngày ký."),
    ("Điều 3.", "Trưởng các phòng/ban và toàn thể người lao động Công ty Cổ phần GMPc Việt Nam chịu trách "
                "nhiệm thi hành Quyết định này./."),
]

# (chương, [(điều, [body lines])])
CHUONG = [
    ("CHƯƠNG I. QUY ĐỊNH CHUNG", [
        ("Điều 1. Mục đích và phạm vi áp dụng", [
            "1. Nội quy lao động này quy định các quy tắc về kỷ luật lao động, thời giờ làm việc, thời giờ nghỉ "
            "ngơi, trật tự tại nơi làm việc, an toàn, vệ sinh lao động, phòng chống quấy rối tình dục, bảo vệ "
            "tài sản, bí mật kinh doanh, bí mật công nghệ và trách nhiệm vật chất.",
            "2. Nội quy áp dụng đối với toàn thể người lao động (NLĐ) làm việc tại Công ty Cổ phần GMPc Việt Nam "
            "theo mọi hình thức hợp đồng lao động (HĐLĐ).",
        ]),
        ("Điều 2. Nguyên tắc áp dụng", [
            "1. Mọi NLĐ đều bình đẳng về quyền và nghĩa vụ trong việc chấp hành Nội quy lao động.",
            "2. Việc xử lý vi phạm kỷ luật lao động phải bảo đảm tính công khai, minh bạch, đúng trình tự, thủ "
            "tục quy định của Bộ luật Lao động.",
        ]),
    ]),
    ("CHƯƠNG II. QUY ĐỊNH VỀ THỜI GIAN LÀM VIỆC", [
        ("Điều 3. Thời gian làm việc", [
            "* CBNV tuân thủ giờ giấc làm việc chung theo quy định của Công ty như sau:",
            "* Thời gian làm việc theo quy định của Công ty từ thứ 2 đến thứ 6 hằng tuần và 2 ngày thứ 7 trong "
            "tháng. Công ty đã quy đổi ngày nghỉ phép sang nghỉ thứ 7, vì vậy thời gian làm việc là từ thứ 2 đến "
            "thứ 6 và không còn ngày nghỉ phép năm cho tất cả mọi trường hợp.",
            "~ Sáng: 8h00 đến 12h00",
            "~ Chiều: 13h00 đến 17h00",
            "* Mọi trường hợp vắng mặt hơn 30 phút trong giờ làm việc đều phải báo cáo với Trưởng bộ phận và báo "
            "lên nhóm Zalo “11. GMPc Báo nghỉ - Công tác”.",
            "* Nhân viên đi làm muộn, về sớm, ra ngoài việc cá nhân quá 30 phút sẽ bị trừ lương.",
            "~ Cách tính trừ lương: Mức lương / số ngày công trong tháng / 8 × số thời gian.",
        ]),
        ("Điều 4. Thời gian làm thêm", [
            "* Làm việc ngoài giờ được tính là thời gian làm thêm của NLĐ ngoài thời gian quy định của Công ty.",
            "* CBNV có nghĩa vụ thực hiện yêu cầu làm thêm giờ khi được phân công.",
            "* Trưởng các bộ phận có trách nhiệm bố trí lao động hợp lý và phải bảo đảm hiệu quả công việc khi "
            "làm thêm giờ.",
            "* Việc tổ chức làm thêm phải có sự xác nhận của NLĐ, Trưởng bộ phận và phụ trách Phòng HCNS.",
            "* Số giờ làm thêm không quá 4h/ngày (theo Điều 60, Chương VII, Nghị định 145/2020/NĐ-CP), ngoại trừ "
            "trường hợp theo yêu cầu của cấp trên.",
            "# Các trường hợp không được công nhận làm việc ngoài giờ:",
            "* Thời gian di chuyển từ nhà đến Công ty, từ Công ty đến khách hàng và ngược lại.",
            "* Thời gian di chuyển trên tàu xe, máy bay trong các chuyến điều động đi công tác.",
            "* Số giờ làm thêm dưới 1h và vượt quá 4h/ngày làm việc bình thường (trừ trường hợp theo yêu cầu của "
            "cấp trên).",
            "* Không áp dụng làm thêm giờ với các trường hợp NLĐ làm theo hợp đồng khoán việc.",
            "* Không áp dụng tính làm thêm giờ vào giờ hành chính (8h–12h và 13h–17h); đi làm vào ngày thứ 7 thì "
            "hưởng lương như ngày thường.",
        ]),
        ("Điều 5. Quy định làm việc tại nhà", [
            "* Cá nhân xin làm việc tại nhà cần lập kế hoạch công việc hằng ngày, gửi Trưởng bộ phận và Trưởng "
            "phòng HCNS để phê duyệt và làm căn cứ chấm công.",
            "* Trưởng bộ phận và Trưởng phòng HCNS căn cứ kế hoạch công việc và tính chất công việc để quyết định "
            "HỆ SỐ LƯƠNG của ngày làm việc tại nhà.",
            "* Trường hợp cá nhân xin làm việc tại nhà: số ngày công thực tế = HỆ SỐ LƯƠNG của ngày làm việc tại "
            "nhà × số ngày công làm việc tại nhà.",
            "* HỆ SỐ LƯƠNG thông thường = 0,7. Các trường hợp bất thường do Trưởng bộ phận, Trưởng phòng HCNS "
            "quyết định.",
            "* Các trường hợp không thực hiện đúng quy định: ngày làm việc tại nhà không được chấm công, tính "
            "nghỉ không lương.",
            "* Công ty chỉ tạo điều kiện làm việc tại nhà trong các trường hợp bất khả kháng không thể đến văn "
            "phòng làm việc.",
            "* Trường hợp Công ty yêu cầu làm việc tại nhà được tính lương, phụ cấp như làm việc tại văn phòng.",
        ]),
        ("Điều 6. Quy định về trang phục làm việc", [
            "* Mặc đồng phục áo trắng phối hợp quần/váy tối màu vào thứ 2 và thứ 6 hằng tuần.",
            "* Khi đi gặp khách hàng bắt buộc phải mặc đồng phục Công ty.",
            "* Trường hợp nhân viên nữ trong thời kỳ mang thai áp dụng mặc màu trắng vào thứ 2, thứ 6 hằng tuần "
            "và khi gặp khách hàng.",
            "* Chu kỳ hỗ trợ đồng phục: 12 tháng/lần. Nếu trong chu kỳ xảy ra hư hỏng, mất mát thì NLĐ phải chịu "
            "hoàn toàn chi phí để mua lại.",
            "* Công ty yêu cầu mọi cá nhân chấp hành theo đúng quy định trên. Phòng HCNS có trách nhiệm theo dõi "
            "việc thực thi quy định. Trường hợp vi phạm, Phòng HCNS áp dụng mức phạt 50.000 đ/lần trừ vào lương "
            "(tiền phạt thu vào quỹ công đoàn) đối với các trường hợp mặc không đúng quy định.",
            "* Làm việc tại nhà máy/công trình khi tham gia thi công, sản xuất phải mặc trang phục bảo hộ lao động "
            "theo quy định.",
        ]),
        ("Điều 7. Quy định về báo nghỉ, công tác", [
            "* Nhân sự có việc cá nhân cần nghỉ có kế hoạch trước phải làm Request xin nghỉ ít nhất trước 1 ngày. "
            "Nội dung Request bao gồm: thời gian nghỉ; lý do nghỉ; công việc cấp bách cần giải quyết trước khi "
            "nghỉ và trong thời gian nghỉ. Sau đó nhập lên “Lịch chung của Công ty trên Google Calendar”.",
            "* Request được phê duyệt bởi người quản lý trực tiếp. Cấu trúc mặc định của Request xin nghỉ được gắn "
            "tất cả các thành viên của GMPc để nắm bắt thông tin, phối hợp công việc.",
            "* Trường hợp nghỉ đột xuất mà không thể làm Request, NLĐ báo lên nhóm Zalo chung “11. GMPc Báo nghỉ - "
            "Công tác” để Phòng HCNS theo dõi chấm công và để các thành viên liên quan giải quyết công việc.",
            "* Các trường hợp đi công tác có kế hoạch, nhân sự nhập lên “Lịch chung của Công ty trên Google "
            "Calendar” để các bộ phận khác nắm bắt, phối hợp thực hiện.",
            "* Đối với trường hợp xin nghỉ (chấm dứt hợp đồng) cần báo trước ít nhất 30 ngày (nội dung và quy "
            "trình xử lý nêu chi tiết tại Điều 10, Chương III).",
        ]),
    ]),
    ("CHƯƠNG III. QUY ĐỊNH VỀ HỢP ĐỒNG LAO ĐỘNG", [
        ("Điều 8. Hợp đồng thử việc (HĐTV)", [
            "* Sau 03 ngày làm việc đầu tiên, hai bên tiến hành ký kết HĐTV.",
            "* Tối đa 15 ngày kể từ ngày nhận việc, nhân sự hoàn tất nộp hồ sơ.",
            "* Đối với NLĐ làm việc thời vụ thì không cần thử việc.",
            "* Trường hợp NLĐ nghỉ trước thời hạn thì thông báo tới Công ty trước ít nhất 3 ngày.",
            "* Trường hợp Công ty cho nghỉ trước thời hạn thì làm Thông báo chấm dứt HĐTV trước thời hạn, thông báo "
            "tới NLĐ trước ít nhất 3 ngày kể từ ngày thông báo có hiệu lực.",
            "* Công ty thanh toán mọi chế độ theo quy định trong HĐTV trong vòng 30 ngày kể từ ngày NLĐ nghỉ nếu "
            "NLĐ thực hiện theo đúng quy định của Công ty.",
        ]),
        ("Điều 9. Hợp đồng lao động (HĐLĐ)", [
            "* Trước thời điểm kết thúc thời gian thử việc 07 ngày, NLĐ làm báo cáo thử việc nộp trực tiếp cho "
            "Trưởng bộ phận để có đánh giá, nhận xét quá trình thử việc của nhân sự.",
            "* Thông qua báo cáo của NLĐ cũng như ý kiến của Trưởng bộ phận, Ban Giám đốc sẽ có quyết định về việc "
            "ký kết hợp đồng, tiếp nhận chính thức.",
            "* HĐLĐ được triển khai ký kết theo đúng quy định của Công ty và pháp luật hiện hành.",
        ]),
        ("Điều 10. Thanh lý hợp đồng", [
            "Nhân sự nghỉ việc (chấm dứt HĐLĐ) phải làm đơn trình lên Ban Giám đốc phê duyệt trước 30 ngày. Trường "
            "hợp nhân viên tự ý nghỉ việc mà chưa được sự đồng ý của Ban lãnh đạo thì nhân viên đó không được "
            "hưởng bất cứ chế độ, quyền lợi nào từ Công ty và có trách nhiệm bồi thường thiệt hại cho Công ty nếu "
            "gây thiệt hại nghiêm trọng.",
            "# Quy trình:",
            "* NLĐ phải Request trình đơn xin chấm dứt/thôi việc lên Trưởng phòng HCNS và Trưởng bộ phận trực "
            "tiếp. Trong đơn xin chấm dứt/thôi việc phải nêu rõ lý do phù hợp quy định của Bộ luật Lao động và "
            "ghi ngày đề nghị chính thức thôi việc.",
            "* Trưởng bộ phận trực tiếp sẽ trao đổi với NLĐ để xác định lý do xin thôi việc và xem xét khả năng NLĐ "
            "có thể tiếp tục công việc tại Công ty.",
            "* Trưởng bộ phận trực tiếp xác định và ghi ngày tháng lên đơn xin chấm dứt/thôi việc trước khi chuyển "
            "cho Phòng HCNS để Phòng HCNS trình Ban Giám đốc.",
            "* Thỏa thuận chấm dứt HĐLĐ sẽ được Phòng HCNS soạn thảo và Ban Giám đốc ký hoặc người được ủy quyền "
            "ký. NLĐ giữ 01 bản và 01 bản được Công ty lưu trong hồ sơ cá nhân của NLĐ.",
            "* Trong 2 tuần cuối cùng, NLĐ có trách nhiệm hợp tác với Trưởng bộ phận trực tiếp và Phòng HCNS để "
            "hoàn thành quy trình bàn giao thôi việc, bao gồm: toàn bộ các công việc đang làm, các thông tin liên "
            "hệ công việc, các CCDC làm việc và các vật dụng khác nếu có. Tất cả các bên liên quan phải ký văn bản "
            "bàn giao.",
            "* Phòng HCNS có trách nhiệm hoàn tất các thủ tục chốt sổ với cơ quan BHXH. NLĐ có nghĩa vụ bàn giao "
            "sổ BHXH cho Phòng HCNS để thực hiện thủ tục. Sổ BHXH kèm tờ rời chốt quá trình đóng sẽ được Phòng "
            "HCNS trả lại cho NLĐ ngay sau khi cơ quan BHXH hoàn tất thủ tục và gửi kết quả về Công ty.",
            "* Phòng HCNS phối hợp với bộ phận Kế toán thanh toán đầy đủ lương và các chế độ thôi việc cho NLĐ theo "
            "đúng quy định của Bộ luật Lao động và HĐLĐ đã ký giữa hai bên.",
        ]),
    ]),
    ("CHƯƠNG IV. BẢO MẬT THÔNG TIN TÀI LIỆU CỦA CÔNG TY", [
        ("Điều 11. Thông tin bảo mật", [
            "* Thông tin được bảo mật của Công ty bao gồm: tất cả các bí mật thương mại, bất cứ thông tin kỹ thuật, "
            "kinh tế, tài chính, marketing hay các thông tin khác như số liệu tài chính, thống kê kế toán, thông "
            "tin về khách hàng, nhà cung cấp, các hoạt động kinh doanh bí mật, hợp đồng và giấy phép mua bán, kế "
            "toán, hệ thống kinh doanh và chương trình vi tính; bất cứ thông tin nào nêu trên liên quan đến hoạt "
            "động của Công ty, đến việc kiện tụng mà liên quan, ảnh hưởng đến Công ty.",
            "* Các thông tin mật như vậy tồn tại dưới bất cứ hình thức nào, kể cả giấy tờ, bản in, thẻ, microfilm "
            "hoặc microfiche, băng từ, đĩa mềm, thông tin trong các file máy tính, qua lời nói và những vật dụng "
            "mang tin khác.",
        ]),
        ("Điều 12. Nghĩa vụ bảo mật thông tin", [
            "* Nhân viên không được thảo luận bất kỳ thông tin hoặc tài liệu nào có chứa thông tin mật với bất cứ "
            "người nào. Quy định vẫn áp dụng ngay cả đối với nhân viên sau khi thôi việc ở Công ty vì bất cứ lý do "
            "gì.",
            "* Nhân viên không được mua, bán, sử dụng, chuyển giao hoặc theo bất kỳ cách thức nào tiết lộ thông tin "
            "mật mà mình biết được trong quá trình làm việc tại Công ty cho bất kỳ bên thứ ba nào, đặc biệt là các "
            "đối thủ cạnh tranh của Công ty, để làm lợi cho bên thứ ba hoặc đối thủ cạnh tranh hoặc gây thiệt hại "
            "cho Công ty.",
            "* Trong thời gian làm việc tại Công ty, nhân viên không được phép cung cấp các dịch vụ tư vấn mà Công "
            "ty đang thực hiện với danh nghĩa cá nhân; nhân viên không được phép lấy thông tin khách hàng của Công "
            "ty để phục vụ cho mục đích riêng của cá nhân. Nhân viên nào vi phạm sẽ không được hưởng các chính "
            "sách về lương, thưởng, nghỉ phép, phụ cấp, đãi ngộ khác của Công ty và Công ty có quyền cho nhân viên "
            "đó thôi việc ngay, kể cả trường hợp chưa hết thời hạn HĐLĐ đã ký.",
            "* Trước khi thôi làm việc cho Công ty, ngoài việc thanh quyết toán các nghĩa vụ tài chính với Công ty, "
            "nhân viên phải bàn giao lại toàn bộ các hồ sơ giấy tờ có chứa các thông tin mật mà nhân viên đó đã "
            "được giao hoặc thuộc mình quản lý, kể cả các file trong máy tính. Nghiêm cấm việc khóa hoặc xóa mã các "
            "file máy tính có chứa thông tin mật trước khi thôi làm việc cho Công ty nếu không được phép của Ban "
            "Giám đốc hoặc người được ủy quyền.",
        ]),
        ("Điều 13. Những trường hợp được phép", [
            "* Nhân viên chỉ được tiết lộ thông tin mật trong các trường hợp sau:",
            "~ Việc tiết lộ là nghĩa vụ bắt buộc theo quy định của pháp luật;",
            "~ Cấp trên trực tiếp được tiết lộ thông tin mật cho nhân viên thuộc quyền quản lý của mình liên quan "
            "trực tiếp đến công việc được giao cho nhân viên đó;",
            "~ Thông tin mật được nhân viên tiết lộ cho nhân viên khác của Công ty hoặc cho một người khác ngoài "
            "Công ty theo yêu cầu hoặc sự phân công của Ban lãnh đạo Công ty.",
        ]),
        ("Điều 14. Vi phạm quy định", [
            "* Nhân viên vi phạm quy định bảo mật thông tin tài liệu của Công ty sẽ không được hưởng các chính sách "
            "về lương, thưởng, nghỉ phép, phụ cấp, đãi ngộ khác của Công ty và Công ty có quyền cho nhân viên đó "
            "thôi việc ngay, kể cả trường hợp chưa hết thời hạn HĐLĐ đã ký.",
            "* Trường hợp nhân viên vi phạm quy định bảo mật thông tin làm Công ty thiệt hại về kinh tế thì phải "
            "bồi thường thiệt hại cho Công ty.",
        ]),
    ]),
    ("CHƯƠNG V. ĐIỀU KHOẢN THI HÀNH", [
        (None, [
            "* Nội quy này được áp dụng với mọi NLĐ trong doanh nghiệp. Mọi trường hợp phát sinh không quy định "
            "trong Nội quy này sẽ được thực hiện theo quy định của Bộ luật Lao động.",
            "* Những quy định trong Nội quy này trái với Bộ luật Lao động thì không có hiệu lực thi hành.",
            "* Căn cứ theo từng thời điểm cụ thể, những quy định bổ sung cho Nội quy sẽ được ban hành dưới hình "
            "thức phụ lục đính kèm và là một phần không thể tách rời của Nội quy này, để phù hợp với tình hình thực "
            "tế và quy định của Nhà nước.",
        ]),
    ]),
]

# ---------------------------------------------------------------- xml helpers
FONT = '<w:rFonts w:ascii="Calibri" w:cs="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/>'
NAVY = '<w:color w:val="002060"/>'
BULLET_NUM_ID = "90"


def run(text, bold=False, italic=False, size=24, color=""):
    b = '<w:b w:val="1"/><w:bCs w:val="1"/>' if bold else ""
    i = '<w:i w:val="1"/><w:iCs w:val="1"/>' if italic else ""
    return (f'<w:r><w:rPr>{FONT}{b}{i}{color}<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'
            f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r>')


def p(runs, jc="both", before=0, after=80, ind="", style="", extra=""):
    st = f'<w:pStyle w:val="{style}"/>' if style else ""
    return (f'<w:p><w:pPr>{st}{extra}<w:spacing w:before="{before}" w:after="{after}"/>{ind}'
            f'<w:jc w:val="{jc}"/><w:rPr>{FONT}</w:rPr></w:pPr>{runs}</w:p>')


def text(t, **kw):
    return p(run(t), **kw)


def bullet(t):
    return p(run(t), extra=f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{BULLET_NUM_ID}"/></w:numPr>',
             ind='<w:ind w:left="284" w:hanging="284"/>', after=60)


def sub(t):
    return p(run("–\t" + t), ind='<w:ind w:left="567" w:hanging="283"/>', after=40,
             extra='<w:tabs><w:tab w:val="left" w:pos="567"/></w:tabs>')


def body(lines):
    out = []
    for ln in lines:
        if ln.startswith("* "):
            out.append(bullet(ln[2:]))
        elif ln.startswith("~ "):
            out.append(sub(ln[2:]))
        elif ln.startswith("# "):
            out.append(p(run(ln[2:], bold=True), before=60, after=60))
        else:
            out.append(text(ln))
    return "".join(out)


def heading(t, level):
    """Heading1 = chương (centred, navy), Heading2 = điều. No list numbering; feeds the TOC."""
    if level == 1:
        return p(run(t, bold=True, size=26, color=NAVY), jc="center", before=240, after=120, style="Heading1",
                 ind='<w:ind w:left="0" w:firstLine="0"/>', extra="<w:keepNext/>")
    return p(run(t, bold=True, color=NAVY), jc="left", before=160, after=80, style="Heading2",
             ind='<w:ind w:left="0" w:firstLine="0"/>', extra="<w:keepNext/>")


PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
NO_BORDERS = ('<w:tblBorders>' + "".join(f'<w:{s} w:val="nil"/>' for s in
              ("top", "left", "bottom", "right", "insideH", "insideV")) + '</w:tblBorders>')


def plain_table(widths, cells_xml):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tcs = "".join(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{c}</w:tc>'
                  for w, c in zip(widths, cells_xml))
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/>{NO_BORDERS}'
            f'<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr>{tcs}</w:tr></w:tbl>')


def decision():
    c = dict(jc="center", after=0)
    left = (p(run("CÔNG TY CỔ PHẦN", bold=True), **c) + p(run("GMPc VIỆT NAM", bold=True), **c)
            + p(run("Số: ........./QĐ-GMPc"), jc="center", before=120, after=0))
    right = (p(run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold=True), **c)
             + p(run("Độc lập - Tự do - Hạnh phúc", bold=True), **c) + p(run("—————————"), **c))
    out = [plain_table([3600, 5754], [left, right]),
           p(run("Hà Nội, ngày ...... tháng ...... năm 2026", italic=True), jc="right", before=200, after=240),
           p(run("QUYẾT ĐỊNH", bold=True, size=28), jc="center", after=0),
           p(run("Về việc ban hành Nội quy lao động", bold=True), jc="center", after=240),
           p(run("BAN GIÁM ĐỐC CÔNG TY CỔ PHẦN GMPc VIỆT NAM", bold=True), jc="center", after=160)]
    out += [p(run(x, italic=True), ind='<w:ind w:firstLine="567"/>', after=60) for x in QD_CAN_CU]
    out.append(p(run("QUYẾT ĐỊNH:", bold=True), jc="center", before=160, after=160))
    out += [p(run(a + " ", bold=True) + run(b), ind='<w:ind w:firstLine="567"/>', after=100) for a, b in QD_DIEU]
    left = (p(run("Nơi nhận:", bold=True, italic=True, size=22), jc="left", after=0)
            + p(run("- Như Điều 3;", size=22), jc="left", after=0) + p(run("- Lưu HCNS.", size=22), jc="left", after=0))
    right = p(run("BAN GIÁM ĐỐC", bold=True), jc="center", after=0)
    out.append(p("", after=120))
    out.append(plain_table([4200, 5154], [left, right]))
    return "".join(out)


def toc():
    fld = ('<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9344"/></w:tabs></w:pPr>'
           '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
           '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>'
           '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
           f'{run("(Mục lục được cập nhật khi dựng file)")}'
           '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    return p(run("MỤC LỤC", bold=True, size=26, color=NAVY), jc="center", after=200) + fld


def content():
    out = [decision(), PAGE_BREAK, toc(), PAGE_BREAK]
    for ch, dieu in CHUONG:
        out.append(heading(ch, 1))
        for title, lines in dieu:
            if title:
                out.append(heading(title, 2))
            out.append(body(lines))
    return "".join(out)


# ---------------------------------------------------------------- build
def section_heading(t):
    """Unnumbered navy heading for the template's trailing sections (history, storage)."""
    return p(run(t, bold=True, size=24, color=NAVY), jc="left", before=240, after=120, style="Heading1",
             ind='<w:ind w:left="0" w:firstLine="0"/>', extra="<w:keepNext/>")


def para_span(doc, needle):
    i = doc.index(needle)
    return doc.rindex("<w:p>", 0, i), doc.index("</w:p>", i) + len("</w:p>")


def fill_signatures(doc):
    """Row 'Họ và tên:' of the sign-off table: one cell per signer, filled left to right."""
    s = doc.index("Họ và tên:")
    r_s = doc.rindex("<w:tr>", 0, s)
    r_e = doc.index("</w:tr>", s)
    row = doc[r_s:r_e]
    tcs = re.findall(r"<w:tc>.*?</w:tc>", row, re.S)
    assert len(tcs) == len(SIGN), len(tcs)
    for tc, (cv, ten) in zip(tcs, SIGN):
        new = re.sub(r'<w:t xml:space="preserve">Chức vụ:[^<]*</w:t>', t_("Chức vụ: " + escape(cv)), tc, count=1)
        new = new.replace(t_("." * 40), t_(escape(ten)), 1)
        assert new.count(escape(ten)) == 1, ten
        row = row.replace(tc, new, 1)
    return doc[:r_s] + row + doc[r_e:]


def build_document(doc):
    doc = re.sub(r' w:rsid\w*="\w+"| w14:(?:paraId|textId)="\w+"', "", doc)
    doc = re.sub(r"<w:proofErr [^>]*/>", "", doc)
    doc = fill_signatures(doc)
    # keep the sign-off table, drop the SOP sections up to (not including) the history heading
    start = doc.index("</w:tbl>") + len("</w:tbl>")
    hist_s, hist_e = para_span(doc, "LỊCH SỬ BAN HÀNH")
    doc = doc[:start] + content() + section_heading("LỊCH SỬ BAN HÀNH VÀ SỬA ĐỔI") + doc[hist_e:]
    # the empty numbered Heading1 + "LƯU TÀI LIỆU" heading
    s, e = para_span(doc, "LƯU TÀI LIỆU")
    prev = doc.rindex("<w:p>", 0, s)
    if '<w:numId w:val="1"/>' in doc[prev:s] and "<w:t" not in doc[prev:s].replace('<w:t xml:space="preserve"></w:t>', ""):
        s = prev
    return doc[:s] + section_heading("LƯU TÀI LIỆU") + doc[e:]


def t_(x):
    return f'<w:t xml:space="preserve">{x}</w:t>'


def build_header(xml):
    for old, new in [("Mã số tài liệu:", f"Mã số tài liệu: {CODE}"), ("BIỂU MẪU TRẮNG SOP", TITLE),
                     ("QUY TRÌNH THAO TÁC CHUẨN (SOP)", KIND)]:
        if t_(old) in xml:
            xml = xml.replace(t_(old), t_(escape(new)))
        else:
            n = xml.count(f">{old}<")
            assert n == 1, (old, n)
            xml = xml.replace(f">{old}<", f">{escape(new)}<")
    return xml


def build_numbering(xml):
    absn = ('<w:abstractNum w:abstractNumId="90"><w:multiLevelType w:val="singleLevel"/>'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>'
            '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="284" w:hanging="284"/></w:pPr>'
            f'<w:rPr>{FONT}</w:rPr></w:lvl></w:abstractNum>')
    i = xml.index("<w:num ")
    xml = xml[:i] + absn + xml[i:]
    return xml.replace("</w:numbering>", f'<w:num w:numId="{BULLET_NUM_ID}"><w:abstractNumId w:val="90"/></w:num></w:numbering>')


def edit_part(tmp, name, fn):
    path = os.path.join(tmp, name)
    with open(path, encoding="utf-8") as f:
        xml = f.read()
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(fn(xml))


def sha(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def update_fields_in_word(path):
    """Fill the TOC and page fields with Word, then save in place."""
    ps = ("$w = New-Object -ComObject Word.Application; $w.Visible = $false; "
          f"$d = $w.Documents.Open('{path}'); $d.TablesOfContents(1).Update(); $d.Fields.Update() | Out-Null; "
          "$n = $d.ComputeStatistics(2); $d.Save(); $d.Close(); $w.Quit(); Write-Output \"pages=$n\"")
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
    print(r.stdout.strip() or r.stderr.strip())


def main():
    if os.path.exists(OUT) and "--force" not in sys.argv:
        stamp = open(STAMP).read().strip() if os.path.exists(STAMP) else None
        if stamp != sha(OUT):
            sys.exit(f"Refusing to overwrite {OUT}: it was changed after the last build. "
                     "Copy those edits into scripts/qd01.py first, then re-run with --force.")
    tmp = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(TEMPLATE) as z:
            names = z.namelist()
            z.extractall(tmp)
        edit_part(tmp, "word/document.xml", build_document)
        edit_part(tmp, "word/header2.xml", build_header)
        edit_part(tmp, "word/header3.xml", build_header)
        edit_part(tmp, "word/numbering.xml", build_numbering)
        for f in ("footer1", "footer2", "footer3"):  # template footer line is in Times New Roman
            edit_part(tmp, f"word/{f}.xml", lambda x: x.replace('"Times New Roman"', '"Calibri"'))
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
            for name in names:
                z.write(os.path.join(tmp, name), name)
    finally:
        shutil.rmtree(tmp)
    update_fields_in_word(OUT)
    with open(STAMP, "w") as f:
        f.write(sha(OUT))
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
