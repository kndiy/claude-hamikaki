"""Build the Minh Viet consulting contract (scope: Part I of the 01/10/2026 quote).

Template: data/raw/hopdong/261014 Draft Hợp đồng tư vấn Minh Viet - GMPc.docx
(the Cafune contract, copied from N:\\02. Dự án\\1. Hợp đồng khách hàng\\Năm 2026\\Minh Viet\\).
Edits whole paragraphs in word/document.xml, keeping each paragraph's pPr and its run
formatting. Text that still needs confirmation is highlighted yellow.

Run from the project root:  python scripts/hopdong_v1.py
"""
import copy
import sys
import zipfile
from pathlib import Path

from lxml import etree

SRC = Path("data/raw/hopdong/261014 Draft Hợp đồng tư vấn Minh Viet - GMPc.docx")
OUT = Path("output/261014 Hợp đồng tư vấn Minh Viet - GMPc - draft v1.docx")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

# Contract number / date / page count: pending confirmation (highlighted).
SO_HD = "1410/HDTV/MinhViet-GMPc/2026"
NGAY, THANG, NAM = "14", "10", "2026"
SO_TRANG = sys.argv[1] if len(sys.argv) > 1 else "05 (năm)"

# rPr child order (subset of the schema sequence) so inserted elements stay valid.
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
             "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
             "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
             "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
             "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]


def set_rpr(rpr, name, on, val=None):
    for e in rpr.findall(W + name):
        rpr.remove(e)
    if not on:
        return
    el = etree.Element(W + name)
    if val:
        el.set(W + "val", val)
    pos = RPR_ORDER.index(name)
    for i, child in enumerate(rpr):
        tag = etree.QName(child).localname
        if tag in RPR_ORDER and RPR_ORDER.index(tag) > pos:
            rpr.insert(i, el)
            return
    rpr.append(el)


def ptext(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def set_para(p, segments, base="last"):
    """Replace the runs of paragraph p. segments: str or list of (text, opts) where
    opts may hold b (bold) and hl (yellow highlight). Formatting comes from an existing run."""
    if isinstance(segments, str):
        segments = [(segments, {})]
    runs = [r for r in p.iter(W + "r") if r.find(W + "t") is not None]
    src = runs[-1] if base == "last" else runs[0]
    base_rpr = src.find(W + "rPr")
    base_bold = base_rpr is not None and base_rpr.find(W + "b") is not None
    for child in list(p):
        if etree.QName(child).localname not in ("pPr",):
            p.remove(child)
    for text, opts in segments:
        r = etree.SubElement(p, W + "r")
        rpr = copy.deepcopy(base_rpr) if base_rpr is not None else etree.Element(W + "rPr")
        bold = opts.get("b", base_bold)
        set_rpr(rpr, "b", bold)
        set_rpr(rpr, "bCs", bold)
        set_rpr(rpr, "highlight", opts.get("hl", False), "yellow")
        r.append(rpr)
        t = etree.SubElement(r, W + "t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def expect(el, prefix):
    got = ptext(el)
    assert got.strip().startswith(prefix), f"template changed? expected {prefix!r}, got {got[:60]!r}"


def cell_par(tr, col):
    return tr.findall(W + "tc")[col].find(W + "p")


HL = {"hl": True}
B = {"b": True}
BHL = {"b": True, "hl": True}
N = {"b": False}


def edit_document(xml):
    root = etree.fromstring(xml)
    body = root.find(W + "body")
    el = list(body)  # indices refer to the untouched template

    expect(el[4], "(Số 2809")
    set_para(el[4], [("(Số ", {}), (SO_HD, HL), (")", {})])

    expect(el[7], "Công Ty Cổ Phần")
    set_para(el[7], "Công Ty Cổ Phần Quốc Tế Minh Việt có nhu cầu thiết kế mặt bằng công nghệ "
                    "và cải tạo nhà xưởng sản xuất đạt tiêu chuẩn CGMP-ASEAN;")
    expect(el[8], "Công ty cổ phần GMPc")
    set_para(el[8], "Công ty cổ phần GMPc Việt Nam là một Công ty tư vấn thành lập và hoạt động "
                    "theo pháp luật Việt Nam có khả năng cung cấp dịch vụ tư vấn theo yêu cầu của "
                    "Công Ty Cổ Phần Quốc Tế Minh Việt.")
    expect(el[9], "Hợp Đồng Dịch Vụ Tư Vấn này")
    set_para(el[9], [("Hợp Đồng Dịch Vụ Tư Vấn này (sau đây gọi tắt là “Hợp Đồng”) được ký kết vào ", {}),
                     (f"ngày {NGAY} tháng {THANG} năm {NAM}", BHL),
                     (", tại Công Ty Cổ Phần Quốc Tế Minh Việt giữa và bởi: ", {})])
    expect(el[10], "BÊN A:")
    set_para(el[10], "BÊN A: CÔNG TY CỔ PHẦN QUỐC TẾ MINH VIỆT")

    # Bên A details
    rows = el[11].findall(W + "tr")
    for tr, val in zip(rows, ["0300721002",
                              "Khu Công nghiệp Tân Phú Trung, ấp Trạm Bơm, Xã Củ Chi, TP Hồ Chí Minh, Việt Nam",
                              "(Ông) Yansen Nagakarsa Adi",
                              "Tổng giám đốc"]):
        set_para(cell_par(tr, 2), val)

    # 2.1 scope table: keep Part I only
    rows = el[20].findall(W + "tr")
    expect(cell_par(rows[1], 1), "Thiết kế mặt bằng")
    set_para(cell_par(rows[1], 1), "Tư vấn và thiết kế mặt bằng công nghệ các hạng mục theo CGMP-ASEAN; "
                                   "Quy hoạch tổng thể, Khái toán chi phí đầu tư")
    set_para(cell_par(rows[1], 2), "45.000.000")
    items = ["Thiết kế mặt bằng xưởng sản xuất, khu kiểm nghiệm theo tiêu chuẩn CGMP-ASEAN",
             "Lập khái toán cải tạo phần xây dựng và M&E",
             "Lập tiến độ cải tạo dự án"]
    for tr, txt in zip(rows[2:5], items):
        set_para(cell_par(tr, 1), txt)
    for tr in rows[5:51]:
        el[20].remove(tr)
    expect(cell_par(rows[51], 1), "TỔNG CỘNG TRƯỚC THUẾ")
    set_para(cell_par(rows[51], 2), "45.000.000")
    set_para(cell_par(rows[52], 2), "3.600.000")
    set_para(cell_par(rows[53], 2), "48.600.000")

    # 2.2 deliverables (Part I)
    expect(el[24], "03 (ba) bản vẽ")
    set_para(el[24], "03 (ba) bộ bản vẽ mặt bằng công nghệ xưởng sản xuất, khu kiểm nghiệm;")
    for txt in reversed(["Bảng khái toán chi phí cải tạo phần xây dựng và M&E;",
                         "Tiến độ cải tạo dự án;"]):
        new = copy.deepcopy(el[24])
        set_para(new, [(txt, HL)])
        el[24].addnext(new)
    expect(el[26], "Mục II")
    expect(el[31], "Các file mềm")
    for e in el[26:32]:
        body.remove(e)

    # 3.1 fee
    expect(el[34], "Phí dịch vụ trước thuế")
    set_para(el[34], [("Phí dịch vụ trước thuế     : \t", N), ("45.000.000 VNĐ", B)])
    set_para(el[35], [("Thuế GTGT 8%     :\t", N), ("  3.600.000 VNĐ", B)])
    set_para(el[36], [("Phí dịch vụ bao gồm 8% thuế GTGT  : ", N), ("48.600.000 VNĐ", B)])
    for p in (el[34], el[35]):  # restore tab characters written as \t
        for t in p.iter(W + "t"):
            if "\t" in t.text:
                before, after = t.text.split("\t")
                t.text = before
                r = t.getparent()
                r.append(etree.Element(W + "tab"))
                if after:
                    t2 = etree.SubElement(r, W + "t")
                    t2.text = after
    set_para(el[37], " (Bằng chữ: Bốn mươi tám triệu sáu trăm nghìn đồng /.)")
    expect(el[39], "Phí dịch vụ chưa bao gồm")
    set_para(el[39], [("Phí dịch vụ không bao gồm các nội dung góp ý cải tạo xây dựng, thiết kế M&E, "
                       "lập dự toán chi tiết và hướng dẫn chuẩn bị hồ sơ xin chứng nhận CGMP-ASEAN "
                       "(Mục II, III của Bản chào giá ngày 01/10/2026). Nếu Bên A có nhu cầu, các nội dung "
                       "này sẽ được Các Bên thỏa thuận bằng phụ lục hoặc hợp đồng riêng.", HL)])

    # 3.3 payment: 50% advance, 50% after liquidation
    expect(el[50], "Lần 1:")
    set_para(el[50], "Lần 1: Tạm ứng 50% phí dịch vụ (tương đương 24.300.000 VNĐ) trong vòng 05 ngày "
                     "làm việc kể từ ngày Hợp Đồng này được ký kết. Bên B cung cấp hóa đơn GTGT tương "
                     "ứng 50% phí dịch vụ.")
    set_para(el[51], [("Lần 2: Thanh toán 50% phí dịch vụ còn lại (tương đương 24.300.000 VNĐ) ", {}),
                      ("trong vòng 05 ngày làm việc", HL),
                      (" kể từ ngày Các Bên ký Biên bản nghiệm thu và thanh lý Hợp Đồng. Bên B cung cấp "
                       "hóa đơn GTGT tương ứng 50% phí dịch vụ.", {})])
    expect(el[54], "Lần 5:")
    for e in el[52:55]:
        body.remove(e)

    # Điều 4: 21 working days
    expect(el[57], "Mục I:")
    set_para(el[57], [("Mục I:", B),
                      (" Trong vòng 21 (hai mươi mốt) ngày làm việc kể từ ngày Hợp Đồng được ký kết và "
                       "Bên B nhận được tiền tạm ứng. Thời gian thực hiện không bao gồm thời gian chờ CĐT "
                       "cung cấp thông tin đầu bài thiết kế và chấp thuận, phê duyệt kết quả thiết kế.", N)])
    expect(el[60], "Thời gian thực hiện không bao gồm thời gian chờ Đoàn")
    for e in el[58:61]:
        body.remove(e)

    # 5.2: liability clause re-scoped to Part I
    expect(el[71], "Tr")
    set_para(el[71], [("Trường hợp hồ sơ thiết kế mặt bằng công nghệ (Mục I) do Bên B lập không đáp ứng "
                       "các nguyên tắc, tiêu chuẩn CGMP-ASEAN do lỗi của Bên B, Bên B có trách nhiệm chỉnh "
                       "sửa, bổ sung miễn phí. Tổng trách nhiệm bồi thường của Bên B theo Hợp Đồng không "
                       "vượt quá giá trị Hợp Đồng. (Trừ trường hợp các lỗi này Bên B đã chỉ ra trong quá "
                       "trình tư vấn và Bên A chấp nhận chưa đầu tư hoặc sẽ thực hiện bổ sung sau nhằm giảm "
                       "chi phí giai đoạn đầu, và trừ các lỗi do đơn vị thi công triển khai không đúng "
                       "thiết kế.)", HL)])

    # 9: ownership
    expect(el[122], "Toàn bộ bản vẽ")
    set_para(el[122], [("Toàn bộ bản vẽ, khái toán, hồ sơ và tài liệu do Bên B lập riêng cho Bên A theo "
                        "Hợp Đồng thuộc quyền sở hữu của Bên A kể từ khi Bên A ", {}),
                       ("thanh toán đầy đủ phí dịch vụ", HL),
                       (". Bên B không sử dụng hồ sơ có chứa thông tin của Bên A cho bất kỳ bên thứ ba nào.", {})])

    # 11: term, page count
    expect(el[127], "Hợp Đồng có hiệu lực")
    set_para(el[127], [("Hợp Đồng có hiệu lực kể từ ngày ký kết cho đến khi Các Bên hoàn thành các nghĩa "
                        "vụ theo Hợp Đồng và ký Biên bản thanh lý Hợp Đồng.", HL)])
    expect(el[130], "Hợp đồng này bao gồm")
    set_para(el[130], [("Hợp đồng này bao gồm ", {}), (SO_TRANG, HL),
                       (" trang, được lập thành 04 (bốn) bản bằng Tiếng Việt, có giá trị pháp lý như nhau, "
                        "mỗi bên giữ 02 (hai) bản./.", {})])

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def edit_header(xml):
    root = etree.fromstring(xml)
    for p in root.iter(W + "p"):
        ts = list(p.iter(W + "t"))
        if "HDTV" in "".join(t.text or "" for t in ts):
            ts[0].text = SO_HD
            for t in ts[1:]:
                t.text = ""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main():
    # v1 was later edited by hand in Word (2026-10-14); never overwrite it.
    if OUT.exists():
        sys.exit(f"{OUT} exists (hand-edited by the user); refusing to overwrite")
    OUT.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = edit_document(data)
            elif item.filename == "word/header1.xml":
                data = edit_header(data)
            zout.writestr(item, data)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
