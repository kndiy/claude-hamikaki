"""Fill a company-wide document (HD = hướng dẫn, QT = quy trình) into a clean copy of the BM01 template.

Adapted from ../20260927 Quy Trinh MDC/scripts/sop_builder.py for this project:
- header subtitle comes from KIND (e.g. "HƯỚNG DẪN") instead of "QUY TRÌNH THAO TÁC CHUẨN (SOP)";
- footer font set to Calibri;
- Sơ đồ khối can be an existing image (FLOW_IMAGE) instead of a rendered chart;
- inline markup in any text: **bold**, __italic__;
- overwrite guard is a SHA-256 stamp written next to the output (any later change, e.g. a Word save, blocks).

Module fields: CODE, TITLE, KIND, OUT (file name in output/), SIGN [(chức vụ, họ tên, ngày|None)] x3,
MUC_DICH / PHAM_VI [paragraphs], TRACH_NHIEM [(bộ phận, [bullets])], DINH_NGHIA [(thuật ngữ, giải thích)],
BUOC [(số, tên bước, [items], [người thực hiện])], optional FLOW_IMAGE (path), TAI_LIEU [(stt, tên, mã)],
LICH_SU [(ngày, phiên bản, lý do, nội dung)], LUU [3 strings].
"""
import hashlib
import os
import re
import shutil
import struct
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, "templates", "GMPc-QT-01.BM01_Biểu mẫu trắng SOP.docx")
OUT_DIR = os.path.join(ROOT, "output")

BULLET_NUM_ID = "90"
FLOW_RID = "rId100"
FLOW_WIDTH_EMU = 5900000  # text width is ~6.5in (A4, margins 1418/1134)
FLOW_MAX_HEIGHT_EMU = 5486400  # 6in

FONT = '<w:rFonts w:ascii="Calibri" w:cs="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/>'
SZ = '<w:sz w:val="24"/><w:szCs w:val="24"/>'
DOTS = "........................................"
NGAY_DOTS = "Ngày: ......./......./ ........"


# ---------------------------------------------------------------- xml helpers
def runs(text, bold=False):
    """Text with **bold** / __italic__ markup -> runs."""
    out = []
    # italic needs a letter inside, so date blanks like "__/__/____" stay literal
    parts = re.split(r"(\*\*.+?\*\*|__(?=[^_]*[^\W_])[^_]+?__)", text)
    for n, part in enumerate(parts):
        if not part:
            continue
        b, i = bold, False
        if n % 2:  # odd items are the captured markup
            b, i = (True, False) if part.startswith("**") else (bold, True)
            part = part[2:-2]
        bx = '<w:b w:val="1"/><w:bCs w:val="1"/>' if b else ""
        ix = '<w:i w:val="1"/><w:iCs w:val="1"/>' if i else ""
        out.append(f'<w:r><w:rPr>{FONT}{bx}{ix}{SZ}</w:rPr><w:t xml:space="preserve">{escape(part)}</w:t></w:r>')
    return "".join(out)


def para(text, center=False, bold=False, after=None, indent=None):
    out = []
    for line in text.split("\n"):
        jc = '<w:jc w:val="center"/>' if center else '<w:jc w:val="both"/>'
        sp = f'<w:spacing w:after="{after}" w:lineRule="auto"/>' if after is not None else ""
        ind = f'<w:ind w:left="{indent}"/>' if indent else ""
        out.append(f'<w:p><w:pPr>{sp}{ind}{jc}<w:rPr>{FONT}{SZ}</w:rPr></w:pPr>{runs(line, bold)}</w:p>')
    return "".join(out)


def bullet(text):
    return (f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/>'
            f'<w:numId w:val="{BULLET_NUM_ID}"/></w:numPr><w:spacing w:after="40"/>'
            f'<w:ind w:left="284" w:hanging="284"/><w:rPr>{FONT}{SZ}</w:rPr></w:pPr>{runs(text)}</w:p>')


def items_xml(items):
    return "".join(bullet(i) for i in items)


def who_xml(who):
    return para(who[0], center=True) if len(who) == 1 else items_xml(who)


def cells(row):
    return re.findall(r"<w:tc>.*?</w:tc>", row, re.S)


def fill_row(row, contents):
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


def replace_text(xml, old, new):
    """Replace one text node's content, whether or not it carries xml:space."""
    if t(old) in xml:
        return xml.replace(t(old), t(escape(new)))
    assert xml.count(f">{old}<") >= 1, old
    return xml.replace(f">{old}<", f">{escape(new)}<")


# ---------------------------------------------------------------- build
def fill_signatures(doc, sign):
    s, e = table_at(doc, "Họ và tên:")
    tbl = doc[s:e]
    for r in rows(tbl):
        if "Họ và tên:" in r:
            new = r
            for tc, (cv, ten, _) in zip(cells(r), sign):
                body = cell_body(tc)
                body = re.sub(r"<w:t xml:space=\"preserve\">Chức vụ:[^<]*</w:t>", t("Chức vụ: " + escape(cv)), body)
                body = body.replace(t(DOTS), t(escape(ten)))
                new = new.replace(tc, tc.replace(cell_body(tc), body), 1)
            tbl = tbl.replace(r, new, 1)
        elif NGAY_DOTS in r:
            new = r
            for tc, (_, _, ngay) in zip(cells(r), sign):
                if ngay:
                    new = new.replace(tc, tc.replace(t(NGAY_DOTS), t(escape("Ngày: " + ngay))), 1)
            tbl = tbl.replace(r, new, 1)
    return doc[:s] + tbl + doc[e:]


def build_document(doc, m):
    doc = re.sub(r' w:rsid\w*="\w+"| w14:(?:paraId|textId)="\w+"', "", doc)
    doc = re.sub(r"<w:proofErr [^>]*/>", "", doc)

    doc = fill_signatures(doc, m.SIGN)
    doc = replace_body_para(doc, "Quy trình này quy định….", "".join(para(x, after=100) for x in m.MUC_DICH))
    doc = replace_body_para(doc, "Quy trình áp dụng cho …..", "".join(para(x, after=100) for x in m.PHAM_VI))

    doc = rebuild_table(doc, "Bộ phận / Vị trí", lambda p: [
        fill_row(p, [para(who, center=True, bold=True), items_xml(items)]) for who, items in m.TRACH_NHIEM])
    doc = rebuild_table(doc, "Từ viết tắt / Thuật ngữ", lambda p: [
        fill_row(p, [para(k, center=True, bold=True), para(v)]) for k, v in m.DINH_NGHIA])

    def steps(p):
        out = []
        for so, ten, items, who in m.BUOC:
            out.append(fill_row(p, [para(so, center=True, bold=True), para(ten, bold=True), para("")]))
            out.append(fill_row(p, [para(""), items_xml(items), who_xml(who)]))
        return out
    doc = rebuild_table(doc, "Người thực hiện / Ghi chú", steps)

    if m.TAI_LIEU:
        doc = rebuild_table(doc, "Tên tài liệu", lambda p: [
            fill_row(p, [para(a, center=True), para(b), para(c, center=True)]) for a, b, c in m.TAI_LIEU])
    doc = rebuild_table(doc, "Lý do sửa đổi", lambda p: [
        fill_row(p, [para(a, center=True), para(b, center=True), para(c), para(d)]) for a, b, c, d in m.LICH_SU])

    s, e = table_at(doc, "Tài liệu gốc")
    tbl = doc[s:e]
    for r, txt in zip(rows(tbl)[1:], m.LUU):
        if txt:
            c = cells(r)
            tbl = tbl.replace(r, fill_row(r, [cell_body(c[0]), para(txt)]), 1)
    return doc[:s] + tbl + doc[e:]


def build_header(xml, m):
    xml = replace_text(xml, "Mã số tài liệu:", f"Mã số tài liệu: {m.CODE}")
    xml = replace_text(xml, "QUY TRÌNH THAO TÁC CHUẨN (SOP)", m.KIND)
    return replace_text(xml, "BIỂU MẪU TRẮNG SOP", m.TITLE)


def build_numbering(xml):
    absn = ('<w:abstractNum w:abstractNumId="90"><w:multiLevelType w:val="singleLevel"/>'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>'
            '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="284" w:hanging="284"/></w:pPr>'
            f'<w:rPr>{FONT}</w:rPr></w:lvl></w:abstractNum>')
    i = xml.index("<w:num ")
    xml = xml[:i] + absn + xml[i:]
    return xml.replace("</w:numbering>", f'<w:num w:numId="{BULLET_NUM_ID}"><w:abstractNumId w:val="90"/></w:num></w:numbering>')


def png_size(path):
    with open(path, "rb") as f:
        head = f.read(24)
    assert head[:8] == b"\x89PNG\r\n\x1a\n", path
    return struct.unpack(">II", head[16:24])


def image_para(w, h, name):
    cx, cy = FLOW_WIDTH_EMU, FLOW_WIDTH_EMU * h // w
    if cy > FLOW_MAX_HEIGHT_EMU:  # tall images: fit the height so the chart stays on one page
        cx, cy = FLOW_MAX_HEIGHT_EMU * w // h, FLOW_MAX_HEIGHT_EMU
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


def add_image(tmp, names, m):
    media = "word/media/flowchart.png"
    shutil.copy(m.FLOW_IMAGE, os.path.join(tmp, media))
    names.append(media)
    w, h = png_size(m.FLOW_IMAGE)
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
        return doc[:s] + heading + image_para(w, h, m.CODE) + doc[e:]
    edit_part(tmp, "word/document.xml", place)


def sha(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def build(m):
    out = os.path.join(OUT_DIR, m.OUT)
    stamp = out + ".sha256"
    if os.path.exists(out) and "--force" not in sys.argv:
        last = open(stamp).read().strip() if os.path.exists(stamp) else None
        if last != sha(out):
            sys.exit(f"Refusing to overwrite {out}: it was changed after the last build. "
                     "Copy those edits into the module first, then re-run with --force.")
    tmp = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(TEMPLATE) as z:
            names = z.namelist()
            z.extractall(tmp)
        edit_part(tmp, "word/document.xml", lambda x: build_document(x, m))
        edit_part(tmp, "word/header2.xml", lambda x: build_header(x, m))
        edit_part(tmp, "word/header3.xml", lambda x: build_header(x, m))
        edit_part(tmp, "word/numbering.xml", build_numbering)
        for f in ("footer1", "footer2", "footer3"):  # template footer line is in Times New Roman
            edit_part(tmp, f"word/{f}.xml", lambda x: x.replace('"Times New Roman"', '"Calibri"'))
        if getattr(m, "FLOW_IMAGE", None):
            add_image(tmp, names, m)
        os.makedirs(OUT_DIR, exist_ok=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for name in names:
                z.write(os.path.join(tmp, name), name)
    finally:
        shutil.rmtree(tmp)
    with open(stamp, "w") as f:
        f.write(sha(out))
    print("Wrote", out)
