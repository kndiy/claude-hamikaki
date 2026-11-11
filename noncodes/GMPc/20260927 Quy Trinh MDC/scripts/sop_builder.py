"""Fill an MD&C SOP into a clean copy of the blank template GMPc-QT-01.BM01.

Each SOP keeps its content in its own module (scripts/qt06.py, qt07.py, ...) and calls build().
Layout follows MDC-QT-05: bullets in cells, a title row + content row per step, a Sơ đồ khối image.

Content conventions for a SOP module:
- SIGN: 3 tuples (chức vụ, họ tên or None, ngày or None); None leaves the template's dots.
- MUC_DICH / PHAM_VI: list of paragraphs.
- TRACH_NHIEM: [(bộ phận, [bullets])]; DINH_NGHIA: [(thuật ngữ, giải thích)].
- BUOC: [(số, tên bước, [items], người thực hiện / ghi chú)]. An item starting with "~" is an
  indented sub-line (no bullet), e.g. a numbered sub-list.
- Optional NGUYEN_TAC: [bullets] -> "6.1. Nguyên tắc thực hiện" before the steps table (as in MDC-QT-05),
  the table then gets the sub-heading "6.2. Các bước thực hiện".
- In MUC_DICH / PHAM_VI a paragraph starting with "* " becomes a bullet.
- Optional OUT_BASE: folder that OUT is relative to (default MDC/SOP).
- TAI_LIEU: [(stt, tên, mã)] (empty list leaves the template's blank rows);
  LICH_SU: [(ngày, phiên bản, lý do, nội dung)]; LUU: 3 strings ("" leaves the cell blank).
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOP_DIR = os.path.join(ROOT, "MDC", "SOP")
TEMPLATE = os.path.join(SOP_DIR, "GMPc-QT-01.BM01_Biểu mẫu trắng SOP.docx")

BULLET_NUM_ID = "90"
FLOW_RID = "rId100"
FLOW_MEDIA = "word/media/flowchart.png"
FLOW_WIDTH_EMU = 5900000  # text width is ~6.5in (A4, margins 1418/1134)

FONT = '<w:rFonts w:ascii="Calibri" w:cs="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/>'
SZ = '<w:sz w:val="24"/><w:szCs w:val="24"/>'
DOTS = "........................................"
NGAY_DOTS = "Ngày: ......./......./ ........"


# ---------------------------------------------------------------- xml helpers
def run(text, bold=False):
    b = '<w:b w:val="1"/><w:bCs w:val="1"/>' if bold else ""
    return f'<w:r><w:rPr>{FONT}{b}{SZ}</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(text, center=False, bold=False, after=None, indent=None):
    """One paragraph; '\n' in text becomes separate paragraphs."""
    out = []
    for line in text.split("\n"):
        jc = '<w:jc w:val="center"/>' if center else ""
        sp = f'<w:spacing w:after="{after}" w:lineRule="auto"/>' if after is not None else ""
        ind = f'<w:ind w:left="{indent}"/>' if indent else ""
        out.append(f'<w:p><w:pPr>{sp}{ind}{jc}<w:rPr>{FONT}{SZ}</w:rPr></w:pPr>{run(line, bold)}</w:p>')
    return "".join(out)


def bullet(text):
    return (f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/>'
            f'<w:numId w:val="{BULLET_NUM_ID}"/></w:numPr><w:spacing w:after="40"/>'
            f'<w:ind w:left="284" w:hanging="284"/><w:rPr>{FONT}{SZ}</w:rPr></w:pPr>{run(text)}</w:p>')


def items_xml(items):
    return "".join(para(i[1:], after=40, indent=567) if i.startswith("~") else bullet(i) for i in items)


def body_xml(paras):
    return "".join(bullet(x[2:]) if x.startswith("* ") else para(x, after=100) for x in paras)


def cells(row):
    return re.findall(r"<w:tc>.*?</w:tc>", row, re.S)


def fill_row(row, contents):
    """Replace the paragraphs of each cell in a template row with the given XML."""
    out, pos = [], 0
    for tc, body in zip(cells(row), contents):
        i = row.index(tc, pos)
        out.append(row[pos:i])
        tcpr = re.search(r"<w:tcPr>.*?</w:tcPr>", tc, re.S).group(0)
        out.append(f"<w:tc>{tcpr}{body}</w:tc>")
        pos = i + len(tc)
    out.append(row[pos:])
    return "".join(out)


def cell_body(tc):
    return tc[tc.index("</w:tcPr>") + len("</w:tcPr>"):-len("</w:tc>")]


def table_at(doc, marker):
    """(start, end) of the table that contains the marker text."""
    i = doc.index(marker)
    s = doc.rindex("<w:tbl>", 0, i)
    return s, doc.index("</w:tbl>", s) + len("</w:tbl>")


def rows(tbl):
    return re.findall(r"<w:tr>.*?</w:tr>", tbl, re.S)


def rebuild_table(doc, marker, new_rows_fn, header_rows=1):
    s, e = table_at(doc, marker)
    tbl = doc[s:e]
    rs = rows(tbl)
    head, proto = rs[:header_rows], rs[header_rows]
    first = tbl.index(rs[0])
    last = tbl.rindex("</w:tr>") + len("</w:tr>")
    new = tbl[:first] + "".join(head) + "".join(new_rows_fn(proto)) + tbl[last:]
    return doc[:s] + new + doc[e:]


def replace_body_para(doc, placeholder, xml):
    i = doc.index(placeholder)
    s = doc.rindex("<w:p>", 0, i)
    e = doc.index("</w:p>", i) + len("</w:p>")
    return doc[:s] + xml + doc[e:]


def t(text):
    return f'<w:t xml:space="preserve">{text}</w:t>'


def set_text(doc, old, new):
    assert t(old) in doc, old
    return doc.replace(t(old), t(escape(new)))


# ---------------------------------------------------------------- build
def fill_signatures(doc, sign):
    """Signature table: row 'Chức vụ' and row 'Ngày', one cell per signer."""
    s, e = table_at(doc, "Họ và tên:")
    tbl = doc[s:e]
    for r in rows(tbl):
        if "Họ và tên:" in r:
            new = r
            for tc, (cv, ten, _) in zip(cells(r), sign):
                body = cell_body(tc)
                if cv:
                    body = re.sub(r"<w:t xml:space=\"preserve\">Chức vụ:[^<]*</w:t>", t("Chức vụ: " + escape(cv)), body)
                if ten:
                    body = body.replace(t(DOTS), t(escape(ten)))
                new = new.replace(tc, tc.replace(cell_body(tc), body), 1)
            tbl = tbl.replace(r, new, 1)
        elif NGAY_DOTS in r:
            new = r
            for tc, (_, _, ngay) in zip(cells(r), sign):
                if ngay:
                    new = new.replace(tc, tc.replace(t(NGAY_DOTS), t(escape(ngay))), 1)
            tbl = tbl.replace(r, new, 1)
    return doc[:s] + tbl + doc[e:]


def build_document(doc, sop):
    doc = re.sub(r' w:rsid\w*="\w+"| w14:(?:paraId|textId)="\w+"', "", doc)
    doc = re.sub(r"<w:proofErr [^>]*/>", "", doc)

    doc = fill_signatures(doc, sop.SIGN)
    doc = replace_body_para(doc, "Quy trình này quy định….", body_xml(sop.MUC_DICH))
    doc = replace_body_para(doc, "Quy trình áp dụng cho …..", body_xml(sop.PHAM_VI))

    doc = rebuild_table(doc, "Bộ phận / Vị trí", lambda p: [
        fill_row(p, [para(who, center=True, bold=True), items_xml(items)]) for who, items in sop.TRACH_NHIEM])

    doc = rebuild_table(doc, "Từ viết tắt / Thuật ngữ", lambda p: [
        fill_row(p, [para(k, center=True, bold=True), para(v)]) for k, v in sop.DINH_NGHIA])

    def steps(p):
        out = []
        for so, ten, items, ghi_chu in sop.BUOC:
            out.append(fill_row(p, [para(so, center=True, bold=True), para(ten, bold=True), para("")]))
            out.append(fill_row(p, [para(""), items_xml(items), para(ghi_chu)]))
        return out
    doc = rebuild_table(doc, "Người thực hiện / Ghi chú", steps)
    if getattr(sop, "NGUYEN_TAC", None):
        s, _ = table_at(doc, "Người thực hiện / Ghi chú")
        xml = (para("6.1. Nguyên tắc thực hiện", bold=True, after=60) + "".join(bullet(b) for b in sop.NGUYEN_TAC)
               + para("6.2. Các bước thực hiện", bold=True, after=60))
        doc = doc[:s] + xml + doc[s:]

    if sop.TAI_LIEU:
        doc = rebuild_table(doc, "Tên tài liệu", lambda p: [
            fill_row(p, [para(a, center=True), para(b), para(c, center=True)]) for a, b, c in sop.TAI_LIEU])

    doc = rebuild_table(doc, "Lý do sửa đổi", lambda p: [
        fill_row(p, [para(a, center=True), para(b, center=True), para(c), para(d)]) for a, b, c, d in sop.LICH_SU])

    # Lưu tài liệu: fill column 2 of the 3 existing rows
    s, e = table_at(doc, "Tài liệu gốc")
    tbl = doc[s:e]
    for r, txt in zip(rows(tbl)[1:], sop.LUU):
        if txt:
            c = cells(r)
            tbl = tbl.replace(r, fill_row(r, [cell_body(c[0]), para(txt)]), 1)
    return doc[:s] + tbl + doc[e:]


def build_header(xml, sop):
    xml = set_text(xml, "Mã số tài liệu:", f"Mã số tài liệu: {sop.CODE}")
    return set_text(xml, "BIỂU MẪU TRẮNG SOP", sop.TITLE)


def build_numbering(xml):
    absn = (
        '<w:abstractNum w:abstractNumId="90"><w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>'
        '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="284" w:hanging="284"/></w:pPr>'
        f'<w:rPr>{FONT}</w:rPr></w:lvl></w:abstractNum>')
    num = f'<w:num w:numId="{BULLET_NUM_ID}"><w:abstractNumId w:val="90"/></w:num>'
    assert 'w:abstractNumId="90"' not in xml
    i = xml.index("<w:num ")
    xml = xml[:i] + absn + xml[i:]
    return xml.replace("</w:numbering>", num + "</w:numbering>")


def flowchart_para(w, h, name):
    cx, cy = FLOW_WIDTH_EMU, FLOW_WIDTH_EMU * h // w
    return (
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="100" name="So do khoi {name}"/><wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="flowchart.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{FLOW_RID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        '</wp:inline></w:drawing></w:r></w:p>')


def edit_part(tmp, name, fn):
    p = os.path.join(tmp, name)
    with open(p, encoding="utf-8") as f:
        xml = f.read()
    with open(p, "w", encoding="utf-8", newline="") as f:
        f.write(fn(xml))


def add_flowchart(tmp, names, sop):
    """Render sop.flowchart() and place it under 'SƠ ĐỒ KHỐI QUY TRÌNH'."""
    png = os.path.join(tmp, "_flow", "flowchart.png")
    os.makedirs(os.path.dirname(png))
    w, h = sop.flowchart().render(png)
    shutil.copy(png, os.path.join(tmp, FLOW_MEDIA))
    names.append(FLOW_MEDIA)

    edit_part(tmp, "word/_rels/document.xml.rels", lambda x: x.replace("</Relationships>",
              f'<Relationship Id="{FLOW_RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
              f'relationships/image" Target="media/flowchart.png"/></Relationships>'))
    edit_part(tmp, "[Content_Types].xml", lambda x: x if 'Extension="png"' in x else x.replace(
              "<Default ", '<Default ContentType="image/png" Extension="png"/><Default ', 1))

    def place(doc):
        i = doc.index("SƠ ĐỒ KHỐI QUY TRÌNH")
        s = doc.rindex("<w:p>", 0, i)
        e = doc.index("</w:p>", i) + len("</w:p>")
        heading = doc[s:e].replace('<w:pStyle w:val="Heading1"/>', '<w:pStyle w:val="Heading1"/><w:keepNext/>', 1)
        return doc[:s] + heading + flowchart_para(w, h, sop.CODE) + doc[e:]
    edit_part(tmp, "word/document.xml", place)


def saved_by_word(path):
    """Our builds have no docProps/; a file that has them was opened and saved in Word (hand edits)."""
    if not os.path.exists(path):
        return False
    with zipfile.ZipFile(path) as z:
        return "docProps/core.xml" in z.namelist()


def build(sop, force=False):
    out = os.path.join(getattr(sop, "OUT_BASE", SOP_DIR), *sop.OUT)
    if saved_by_word(out) and not force and "--force" not in sys.argv:
        sys.exit(f"Refusing to overwrite {out}: it has been edited and saved in Word. "
                 "Copy those edits into the SOP module first, then re-run with --force.")
    tmp = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(TEMPLATE) as z:
            names = z.namelist()
            z.extractall(tmp)
        edit_part(tmp, "word/document.xml", lambda x: build_document(x, sop))
        edit_part(tmp, "word/header2.xml", lambda x: build_header(x, sop))
        edit_part(tmp, "word/header3.xml", lambda x: build_header(x, sop))
        edit_part(tmp, "word/numbering.xml", build_numbering)
        if getattr(sop, "flowchart", None):
            add_flowchart(tmp, names, sop)
        os.makedirs(os.path.dirname(out), exist_ok=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for name in names:  # keep original order ([Content_Types].xml first)
                z.write(os.path.join(tmp, name), name)
    finally:
        shutil.rmtree(tmp)
    print("Wrote", out)
