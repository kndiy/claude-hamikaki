"""Draft v2: Minh Viet file as template + sections II, III from the Cafune quote (bilingual VN/EN).
English wording follows the terminology of the Cafune ENG sheet. Raw files in data/raw are never modified.
Run from the project root: python scripts/draft_v2.py"""
import math
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

SRC = "data/raw/261001 Bao gia Minh Viet.xlsx"
DST = "output/261001 Bao gia Minh Viet - draft v2.xlsx"
FONT = "Calibri Light"
YELLOW = PatternFill("solid", fgColor="FFFF00")

wb = openpyxl.load_workbook(SRC)
ws = wb["Báo giá"]

# ---- header / customer fixes ----------------------------------------------
# GMPc office addresses: new administrative names, as in the Cafune quote
ws["B2"] = "Trụ sở: 18 Măng Cầm 1,  An Lạc Green Symphony, Sơn Đồng, Hà Nội"
ws["B3"] = "VPGD: 32 Vĩ Cầm 5, An Lạc Green Symphony, Sơn Đồng, Hà Nội"
ws["B4"] = "VPĐD: Lầu 2, 156/1/1 Cộng Hòa, Phường Bảy Hiền, TP Hồ Chí Minh"
ws["B12"] = ":  [Bổ sung số điện thoại]"
ws["B13"] = ":  [Bổ sung email]"
ws["B14"] = ":  01/10/2026"
ws["B22"] = ("Tư vấn và thiết kế mặt bằng công nghệ các hạng mục theo CGMP-ASEAN; "
             "Quy hoạch tổng thể, Khái toán chi phí đầu tư\n"
             "Consulting and site layout design in accordance with CGMP-ASEAN; "
             "Site master planning, Preliminary Estimate of Investment Project")
ws.row_dimensions[22].height = 75
for c in ("B10", "B12", "B13", "B22", "B23"):
    ws[c].fill = YELLOW

# ---- sections II, III (Cafune sheet "Mỹ phẩm" rows 25-68) -------------------
# (no, level, VN, EN, price)  level: S = section, 1 = bold item, 2 = sub-item
ITEMS = [
    ("II", "S", "Góp ý cải tạo xây dựng, thiết kế M&E và Lập dự toán",
     "Construction renovation consulting, MEP design and Budget estimating", 150000000),
    (1, "1", "Góp ý cải tạo xây dựng", "Construction renovation consulting", None),
    ("1.1", "2", "Khảo sát, lập bản vẽ thực trạng công trình",
     "Field survey, Existing building drawing", None),
    ("1.2", "2", "Liệt kê các hạng mục xây dựng cần cải tạo: nhà xưởng, nhà văn phòng, các hạng mục phụ trợ và hệ thống hạ tầng kỹ thuật",
     "Listing construction items to be renovated: premises, office buildings, auxiliary items and technical infrastructure systems", None),
    (2, "1", "Thiết kế M&E", "MEP Design", None),
    ("2.1", "2", "Thiết kế hệ thống HVAC", "Design HVAC system", None),
    ("2.2", "2", "Thiết kế điện động lực và chiếu sáng", "Design the electrical and lighting system", None),
    ("2.3", "2", "Thiết kế hệ thống cấp và thoát nước", "Design the water supply and drainage system", None),
    ("2.4", "2", "Thiết kế hệ thống cấp nước RO, nước sinh hoạt, nước cất (nếu có), …",
     "Design the RO water supply system, domestic water, distilled water (if any), …", None),
    ("2.5", "2", "Thiết kế hệ thống khí nén", "Design the compressed air system", None),
    ("2.6", "2", "Thiết kế hệ thống vách trần panel, sơn epoxy", "Design ceiling & wall panel system, epoxy paint", None),
    ("2.7", "2", "Các hạng mục khác (Thông tin liên lạc, Internet, Camera...)",
     "Other items (Communication, Internet, CCTV...)", None),
    (3, "1", "Lập dự toán các hạng mục cải tạo xây dựng và M&E",
     "Detail budget estimating for all construction renovation items, MEP categories", None),
    ("III", "S", "Hướng dẫn chuẩn bị hồ sơ xin chứng nhận Đủ điều kiện sản xuất Mỹ phẩm và chứng nhận CGMP ASEAN",
     "GMP Training - Consulting on drafting and finalizing the registration dossier for the Certificate of Eligibility for Cosmetics Manufacturing and CGMP-ASEAN certification", 260000000),
    (1, "1", "Yêu cầu nhân sự", "Recruitment support", None),
    ("1.1", "2", "Tư vấn, lập Sơ đồ, cơ cấu tổ chức nhân sự đáp ứng yêu cầu GMP",
     "Consulting, drafting of company personnel structure in compliance with GMP requirement", None),
    ("1.2", "2", "Lập mô tả công việc và yêu cầu về nhân sự cho các vị trí chủ chốt (Trưởng bộ phận, các vị trí yêu cầu chuyên môn, …) để CĐT tuyển dụng nhân sự",
     "Drafting job description and requirement for Investors to recruit suitable key personnel (Heads of department, specialist positions, …)", None),
    ("1.3", "2", "Góp ý về chuyên môn của ứng viên khi chủ đầu tư tuyển dụng",
     "Providing professional feedback on applicants during the Investor's recruitment", None),
    (2, "1", "Đào tạo cơ bản", "Basic training", None),
    ("2.1", "2", "Lập kế hoạch đào tạo", "Develop a training plan", None),
    ("2.2", "2", "Đào tạo kiến thức cơ bản về GMP và các quy định pháp luật liên quan cho nhân sự chủ chốt (10-15 nhân sự)",
     "Training basic knowledge about GMP and related legal regulations for key personnel (10-15 persons)", None),
    ("2.3", "2", "Hướng dẫn xây dựng kế hoạch, lập hồ sơ đào tạo nội bộ cho các phòng ban liên quan trong nhà máy.",
     "Guide to develop internal training plans and records for related departments in the factory", None),
    ("2.4", "2", "Liệt kê các nội dung đào tạo nhà máy tự thực hiện với cơ quan có chức năng: PCCC, ATLĐ, vận hành thiết bị áp lực, vi sinh, …",
     "Listing training contents the factory arranges with competent authorities: fire protection, occupational safety, pressure equipment operation, microbiology, …", None),
    (3, "1", "Hướng dẫn soạn thảo hồ sơ-tài liệu", "Guide to draft documents - records", None),
    ("3.1", "2", "Lập danh mục hồ sơ tài liệu và kế hoạch chuẩn bị.",
     "Create document catalogs, records and preparation plans.", None),
    ("3.2", "2", "Cung cấp bộ SOP mẫu cho các phòng ban liên quan trong nhà máy (QA, QC, sản xuất, cơ điện, kế hoạch-kho, nhân sự) và hướng dẫn nhân sự nhà máy đọc hiểu, hiệu chỉnh nội dung phù hợp.",
     "Provide template SOPs for related departments (QA, QC, production, engineering, planning-warehouse, HR) and guide factory personnel to understand and adjust them accordingly.", None),
    ("3.3", "2", "Hướng dẫn soạn thảo hồ sơ tài liệu liên quan tổ chức, nhân sự cho các phòng ban và cho từng nhân sự làm việc trong nhà máy",
     "Guide to complete the organizational charts of departments, personnel documents for each employee in the factory", None),
    ("3.4", "2", "Hướng dẫn soạn thảo hồ sơ tài liệu cần thiết cho các hệ thống phụ trợ (HVAC, hệ thống nước, khí nén, …)",
     "Guide to draft documents for auxiliary systems (HVAC system, water system, compressed air system, …)", None),
    ("3.5", "2", "Hướng dẫn soạn thảo hồ sơ tài liệu của trang thiết bị sản xuất, kiểm nghiệm, bảo quản (Vận hành, vệ sinh, bảo trì, IQ, OQ, PQ)",
     "Guide to complete the operation procedure, cleaning and maintenance of equipment; IQ, OQ, PQ for production equipment, testing equipment, storage equipment", None),
    ("3.6", "2", "Hướng dẫn soạn thảo hồ sơ tài liệu cho từng dạng sản phẩm (Tiêu chuẩn nguyên vật liệu, tiêu chuẩn sản phẩm, quy trình sản xuất, hồ sơ lô, thẩm định quy trình sản xuất, nghiên cứu độ ổn định, …)",
     "Guide to develop documents for each product type (standards for raw materials, finished products, master manufacturing process, batch records, manufacturing process validation, stability study, …)", None),
    ("3.7", "2", "Các hồ sơ tài liệu khác liên quan", "Other documents", None),
    (4, "1", "Hướng dẫn thực hành và hoàn thiện hồ sơ", "Practice Guidance and finalizing records", None),
    ("4.1", "2", "Xây dựng kế hoạch vận hành, thẩm định nhà xưởng",
     "Develop a plan for facility operation and validation", None),
    ("4.2", "2", "Hướng dẫn trang bị các phương tiện, vật dụng cần thiết cho xưởng sản xuất, phòng kiểm nghiệm, kho, ... để chuẩn bị vận hành",
     "Guide to prepare necessary tools and supplies for production area, quality control area, storage warehouse, ... for operation", None),
    ("4.3", "2", "Hướng dẫn tổ chức, hoạt động các phòng ban chức năng như sản xuất theo các SOP đã viết",
     "Guide to organize the activities in functional areas such as production, ... according to the written SOPs", None),
    ("4.4", "2", "Hướng dẫn thực hiện công tác đào tạo, đánh giá nhà cung cấp, tự thanh tra, thẩm định",
     "Guide to carry out training, supplier evaluation, self-inspection and validation", None),
    ("4.5", "2", "Hướng dẫn chuẩn bị nguyên vật liệu để sản xuất các lô thử nghiệm/ thẩm định phục vụ cho việc đánh giá GMP",
     "Guide to prepare raw materials for producing trial/ validation batches for the GMP assessment", None),
    ("4.6", "2", "Hướng dẫn ghi chép, hoàn thiện hồ sơ các nội dung đã thực hiện",
     "Guide to record and finalize the records of completed activities", None),
    ("4.7", "2", "Các nội dung khác liên quan", "Other instructions", None),
    (5, "1", "Hướng dẫn và trực tiếp tham gia tiến hành tiền đánh giá",
     "Guide and directly participate in the pre-assessment (self-inspection)", None),
    (6, "1", "Hướng dẫn hoàn thiện, khắc phục các tồn tại sau tiền đánh giá",
     "Guidance for completing and overcoming the shortcomings after the pre-assessment", None),
    (7, "1", "Hướng dẫn chuẩn bị, tổng hợp toàn bộ hồ sơ xin đánh giá Đủ điều kiện sản xuất mỹ phẩm",
     "Guidance on preparing and summarizing all dossiers of application for assessment of Eligibility for Cosmetics Manufacturing", None),
    (8, "1", "Hướng dẫn chuẩn bị, tổng hợp toàn bộ hồ sơ xin đánh giá CGMP",
     "Guidance on preparing and summarizing all dossiers of application for CGMP assessment", None),
    (9, "1", "Hỗ trợ Chủ đầu tư chuẩn bị tiếp Đoàn đánh giá Đủ điều kiện sản xuất Mỹ phẩm, CGMP và khắc phục sau đánh giá đến khi nhận được giấy chứng nhận Đủ điều kiện sản xuất mỹ phẩm, CGMP",
     "Support the Investor to welcome the audit team for Eligibility for Cosmetics Manufacturing and CGMP, and remedy existing problems after the audit until the date of receiving the Certificate of Eligibility for Cosmetics Manufacturing and the CGMP certificate", None),
]
N = len(ITEMS)
FOOT_FIRST, FOOT_LAST = 26, 39


def shift_ref(ref, n):
    col = "".join(ch for ch in ref if ch.isalpha())
    row = int("".join(ch for ch in ref if ch.isdigit()))
    return f"{col}{row + n}"


# move footer (totals, terms, signature) down by N rows
foot_merges = [str(m) for m in ws.merged_cells.ranges if m.min_row >= FOOT_FIRST]
heights = {r: ws.row_dimensions[r].height for r in range(FOOT_FIRST, FOOT_LAST + 1)}
for m in foot_merges:
    ws.unmerge_cells(m)
ws.move_range(f"A{FOOT_FIRST}:D{FOOT_LAST}", rows=N)
for m in foot_merges:
    a, b = m.split(":")
    ws.merge_cells(f"{shift_ref(a, N)}:{shift_ref(b, N)}")
for r in range(FOOT_FIRST, FOOT_LAST + 1):
    ws.row_dimensions[r].height = None
for r, h in heights.items():
    ws.row_dimensions[r + N].height = h


def style_from(dst, src, bold=None):
    f = copy(src.font)
    if bold is not None:
        f = Font(name=f.name, sz=f.sz, b=bold, i=f.i, color=f.color)
    dst.font = f
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format


def est_height(text, chars=62, pt=15):
    lines = sum(math.ceil(max(len(s), 1) / chars) for s in text.split("\n"))
    return lines * pt + 6


for i, (no, lvl, vn, en, price) in enumerate(ITEMS):
    r = FOOT_FIRST + i
    tpl = 22 if lvl == "S" else 23
    for col in "ABCD":
        sub = lvl == "2" and col == "B"
        style_from(ws[f"{col}{r}"], ws[f"{col}{tpl}"], bold=False if sub else None)
    ws[f"A{r}"] = no
    ws[f"B{r}"] = f"{vn}\n{en}"
    if lvl == "S":
        ws[f"C{r}"] = 1
        ws[f"D{r}"] = price
        ws[f"D{r}"].fill = YELLOW  # price copied from the Cafune quote
    ws.row_dimensions[r].height = est_height(ws[f"B{r}"].value)

T = FOOT_FIRST + N  # total row
ws[f"D{T}"] = f"=SUM(D22:D{T - 1})"
ws[f"D{T + 1}"] = f"=D{T}*8%"
ws[f"D{T + 2}"] = f"=D{T}+D{T + 1}"

# ---- terms (originally rows 30/31/32) -------------------------------------
NOTE, TIME, PAY = 30 + N, 31 + N, 32 + N
ws[f"B{NOTE}"] = (
    ":  Thuế suất thuế GTGT 8% áp dụng theo nghị định số 174/2025/NĐ-CP ban hành ngày 30/06/2025 có hiệu lực thi hành từ ngày 01/07/2025 đến hết ngày 31/12/2026.\n"
    "    VAT rate of 8% is applied in accordance with Decree No. 174/2025/NĐ-CP, issued on June 30, 2025 and effective from July 1, 2025 to December 31, 2026.\n"
    "    Lưu ý: Báo giá trên chưa bao gồm chi phí nộp hồ sơ xin chứng nhận Đủ điều kiện sản xuất/ CGMP và các chi phí phát sinh khác. Không bao gồm dịch vụ tư vấn chuyển giao công thức sản phẩm và công nghệ thiết bị, chủ đầu tư cần cung cấp thông tin về sản phẩm và thiết bị sản xuất theo nhu cầu của nhà máy để GMPc thực hiện thiết kế.\n"
    "    Note: The above quotation excludes the fees for submitting the application for the Certificate of Eligibility for Manufacturing/ CGMP and other incidental costs. It does not include consulting on product formulation and equipment technology transfer; the Investor shall provide information on products and manufacturing equipment as required by the factory for GMPc to carry out the design.")
ws.row_dimensions[NOTE].height = 165
ws[f"B{TIME}"] = (
    ":  - Mục I: Trong vòng 3 tuần làm việc kể từ ngày ký kết hợp đồng và nhận được tiền tạm ứng.\n"
    "      Part I: Within 3 working weeks from the date of contract signing and receipt of the advance payment.\n"
    "    - Mục II: Trong vòng 30 ngày kể từ ngày hai bên thống nhất chuyển sang giai đoạn thiết kế chi tiết.\n"
    "      Part II: Within 30 days from both parties' agreement on the initiation of the detail-design step.\n"
    "    - Mục III: Trong vòng 3-4 tháng kể từ ngày Chủ đầu tư tuyển dụng đủ nhân sự chủ chốt để triển khai hồ sơ tài liệu, GMPc sẽ đề xuất thời gian bắt đầu triển khai phù hợp với tiến độ hoàn thành thi công nhà xưởng và lắp đặt trang thiết bị phục vụ sản xuất, kiểm nghiệm (thông thường trước khi nhà xưởng hoàn thiện khoảng 2 tháng).\n"
    "      Part III: Within 3-4 months from the date the Investor has recruited sufficient key personnel to develop the documentation. GMPc will propose a suitable start date in line with the completion of factory construction and installation of production and testing equipment (usually about 2 months before the factory is completed).")
ws.row_dimensions[TIME].height = 180
ws[f"B{PAY}"] = (
    ":  - Lần 1: Tạm ứng 30% phí dịch vụ trong vòng 05 ngày kể từ ngày ký kết hợp đồng.\n"
    "       1st payment: Advance 30% of the contract amount within 05 days after the signing of the Contract.\n"
    "    - Lần 2: Thanh toán 20% phí dịch vụ trong vòng 05 ngày sau khi hoàn thành mục II.\n"
    "       2nd payment: Payment 20% of the contract amount within 05 days after completion of Part II.\n"
    "    - Lần 3: Thanh toán 20% phí dịch vụ trong vòng 05 ngày sau khi CĐT được cấp chứng nhận Đủ điều kiện sản xuất.\n"
    "       3rd payment: Payment 20% of the contract amount within 05 days after the Investor is granted the Certificate of Eligibility for Manufacturing.\n"
    "    - Lần 4: Thanh toán 20% phí dịch vụ trong vòng 05 ngày sau khi CĐT nộp hồ sơ xin chứng nhận CGMP.\n"
    "       4th payment: Payment 20% of the contract amount within 05 days after the Investor submits the CGMP application dossier.\n"
    "    - Lần 5: Thanh toán phần còn lại trong vòng 05 ngày sau khi nghiệm thu, thanh lý hợp đồng.\n"
    "       5th payment: The remaining of the contract amount shall be paid within 05 days after project acceptance and contract liquidation.")
ws.row_dimensions[PAY].height = 215
for c in (f"B{NOTE}", f"B{TIME}", f"B{PAY}"):
    ws[c].fill = YELLOW

ws.print_area = f"A1:D{46 + N}"

# ---- font: Calibri Light everywhere (keep size / bold / italic / colour) ----
for row in ws.iter_rows():
    for c in row:
        if c.font.name != FONT:
            f = copy(c.font)
            f.name = FONT
            c.font = f

# ---- review sheet (delete before sending) ---------------------------------
P2, P3 = FOOT_FIRST, FOOT_FIRST + 13
rv = wb.create_sheet("Ghi chu soan thao")
rows = [
    ("#", "Loại", "Ô", "Nội dung", "Trạng thái"),
    (1, "Đã sửa", "B2:B4", "Địa chỉ văn phòng GMPc theo mẫu Cafune (Sơn Đồng; P. Bảy Hiền).", "Xong"),
    (2, "Đã sửa", "B14", "Ngày báo giá -> 01/10/2026.", "Xong"),
    (3, "Đã sửa", "Toàn sheet", "Font thống nhất Calibri Light.", "Xong"),
    (4, "Đã thêm", f"A{P2}:D{T - 1}", "Mục II, III chép từ báo giá Cafune (sheet 'Mỹ phẩm'). Tiếng Anh theo thuật ngữ sheet ENG của Cafune; dòng không có trong ENG do Claude dịch. Sửa lỗi gốc: 'CGMP ASAN' -> 'CGMP ASEAN'.", "Soát câu chữ EN"),
    (5, "Cần xác nhận", f"D{P2}, D{P3}", "Giá mục II (150tr) và III (260tr) lấy nguyên từ Cafune (xưởng ~1.100 m2, mỹ phẩm dạng ướt). Cần định giá theo quy mô Minh Việt.", "Chờ"),
    (6, "Cần xác nhận", f"B{NOTE}", "Ghi chú VAT -> NĐ 174/2025; thêm câu loại trừ chi phí từ Cafune (song ngữ).", "Chờ"),
    (7, "Cần xác nhận", f"B{TIME}", "Thời gian thực hiện tách Mục I/II/III: Mục I giữ '3 tuần' của Minh Việt, Mục II/III theo Cafune.", "Chờ"),
    (8, "Cần xác nhận", f"B{PAY}", "Thanh toán thay bằng 5 đợt của Cafune (30/20/20/20/còn lại 10%). Bản gốc Minh Việt có 2 đợt và mâu thuẫn VN 50% / EN 30%.", "Chờ"),
    (9, "Cần xác nhận", "B22/B23", "Mục I: tiêu đề có 'Quy hoạch tổng thể' nhưng không có hạng mục con; mục 1 VN 'Thiết kế mặt bằng' vs EN 'feedback on renovation'. Mục I.2 (khái toán cải tạo) trùng một phần với II.3 (dự toán cải tạo).", "Chờ"),
    (10, "Cần xác nhận", "B10", "Địa chỉ KH: 'Củ Chi' không còn từ 01/07/2025, KCN chưa có tên.", "Chờ"),
    (11, "Cần bổ sung", "B12/B13/B17", "SĐT, email chị Lan; đối chiếu tên tiếng Anh với ĐKKD.", "Chờ"),
    (12, "Lưu ý", "-", "Bỏ tô vàng và xoá sheet này trước khi gửi khách hàng.", "-"),
]
for r in rows:
    rv.append(r)
for row in rv.iter_rows():
    for c in row:
        c.font = Font(name=FONT, sz=11, b=(c.row == 1))
        c.alignment = Alignment(wrap_text=True, vertical="top")
for col, w in zip("ABCDE", (5, 14, 14, 100, 16)):
    rv.column_dimensions[col].width = w

wb.save(DST)
print("saved", DST, "| items:", N, "| total row:", T)
