"""Build a Word form (biểu mẫu) on the company form frame used by MDC-QT-02.BM03A:
letterhead + logo in the header, form title and "Mã số" under it, then our own body.

The body is a list of blocks made with the helpers below (h1, p, check, table, ...).
Like sop_builder, it refuses to overwrite a file that has been saved in Word (use --force after syncing).
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

from sop_builder import ROOT, SOP_DIR, saved_by_word

FRAME = os.path.join(SOP_DIR, "MDC-QT-02 QT Tiếp nhận và điều phối khách hàng", "MDC-QT-02.BM03A PKS WHO GMP Tan duoc.docx")
TEXT_W = 11909 - 1418 - 1134  # page width minus margins (dxa)

NAVY, TEAL = "1B1464", "00A2CA"
FONT = '<w:rFonts w:ascii="Calibri" w:cs="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/>'
BOX = "☐"
DOTS = "…" * 40


def run(text, bold=False, italic=False, color=None, size=24):
    rpr = FONT + ('<w:b/><w:bCs/>' if bold else "") + ('<w:i/><w:iCs/>' if italic else "")
    rpr += f'<w:color w:val="{color}"/>' if color else ""
    rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def p(*parts, align=None, before=0, after=80, indent=None, keep=False, size=24):
    """Paragraph from strings or (text, dict(style)) tuples."""
    runs = "".join(run(x, size=size) if isinstance(x, str) else run(x[0], size=size, **x[1]) for x in parts)
    ppr = '<w:keepNext/>' if keep else ""
    ppr += f'<w:spacing w:before="{before}" w:after="{after}" w:line="264" w:lineRule="auto"/>'
    ppr += f'<w:ind w:left="{indent}"/>' if indent else ""
    ppr += f'<w:jc w:val="{align}"/>' if align else ""
    return f'<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>'


def h1(text):
    return p((text, {"bold": True, "color": NAVY}), before=160, after=80, keep=True)


def b(text):
    return (text, {"bold": True})


def i(text):
    return (text, {"italic": True})


def check(*options, indent=284):
    """One line of check boxes: ☐ A    ☐ B."""
    return p("     ".join(f"{BOX} {o}" for o in options), indent=indent)


def cell(content, w, fill=None, bold=False, color=None, align=None, span=1):
    paras = content if isinstance(content, list) else [content]
    body = ""
    for x in paras:
        if isinstance(x, str) and x.startswith("<w:p>"):
            body += x
        else:
            txt = x if isinstance(x, str) else x[0]
            body += (f'<w:p><w:pPr><w:spacing w:before="20" w:after="20"/>'
                     f'{f"<w:jc w:val=\"{align}\"/>" if align else ""}</w:pPr>'
                     f'{run(txt, bold=bold, color=color, size=22)}</w:p>')
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    gs = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{gs}{shd}<w:vAlign w:val="center"/></w:tcPr>'
            f'{body}</w:tc>')


def table(widths, rows, header=True, borders=True, center_cols=()):
    """rows: list of lists of cell contents; a row given as ('span', text) is one full-width shaded row."""
    total = sum(widths)
    bd = "single" if borders else "nil"
    tblpr = (f'<w:tblPr><w:tblW w:w="{total}" w:type="dxa"/><w:tblBorders>'
             + "".join(f'<w:{s} w:val="{bd}" w:sz="4" w:space="0" w:color="808080"/>'
                       for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
             + '</w:tblBorders><w:tblLayout w:type="fixed"/>'
               '<w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>')
    grid = "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths) + "</w:tblGrid>"
    out = []
    for n, r in enumerate(rows):
        hdr = header and n == 0
        trpr = '<w:trPr><w:cantSplit/>' + ('<w:tblHeader/>' if hdr else "") + '</w:trPr>'
        if isinstance(r, tuple) and r[0] == "span":
            out.append(f"<w:tr>{trpr}{cell(r[1], total, fill='E8E8F4', bold=True, color=NAVY, span=len(widths))}</w:tr>")
            continue
        cs = "".join(cell(c, w, fill=NAVY if hdr else None, bold=hdr, color="FFFFFF" if hdr else None,
                          align="center" if hdr or k in center_cols else None)
                     for k, (c, w) in enumerate(zip(r, widths)))
        out.append(f"<w:tr>{trpr}{cs}</w:tr>")
    return f"<w:tbl>{tblpr}{grid}{''.join(out)}</w:tbl>" + p("", after=0)


def signatures(*titles):
    """Signature row without borders. Each title is a string or (title, subtitle)."""
    w = TEXT_W // len(titles)
    row = []
    for t in titles:
        t, sub = (t, None) if isinstance(t, str) else t
        row.append([p(b(t), align="center", after=0)] + ([p(i(sub), align="center", after=0)] if sub else [])
                   + [p(i("(Ký, ghi rõ họ tên)"), align="center", after=1400)])
    return table([w] * len(titles), [row], header=False, borders=False)


# ---------------------------------------------------------------- build
def _fix_header(xml, title_lines, code):
    xml = re.sub(r' w:rsid\w*="\w+"| w14:(?:paraId|textId)="\w+"', "", xml)
    xml = re.sub(r"<w:proofErr [^>]*/>", "", xml)
    t1 = '<w:t xml:space="preserve">PHIẾU KHẢO SÁT THÔNG TIN KHÁCH HÀNG </w:t>'
    t2 = "<w:t>PHẠM VI: WHO-GMP TÂN DƯỢC</w:t>"
    assert t1 in xml and t2 in xml, "frame header changed"
    xml = xml.replace(t1, f'<w:t xml:space="preserve">{escape(title_lines[0])}</w:t>')
    if len(title_lines) > 1:
        xml = xml.replace(t2, f'<w:t xml:space="preserve">{escape(title_lines[1])}</w:t>')
    else:  # drop the second title paragraph
        e = xml.index(t2)
        s = xml.rindex("<w:p>", 0, e)
        xml = xml[:s] + xml[xml.index("</w:p>", e) + 6:]
    code_runs = re.search(r"<w:r><w:rPr>(?:(?!</w:r>).)*<w:t>Mã</w:t></w:r>.*?<w:t>: MDC-QT-02\.BM03A</w:t></w:r>", xml, re.S)
    assert code_runs, "frame code line changed"
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", code_runs.group(0), re.S).group(0)
    return xml.replace(code_runs.group(0), f'<w:r>{rpr}<w:t xml:space="preserve">Mã số: {escape(code)}</w:t></w:r>')


def _page_footer(xml):
    """The frame's footer only has a NUMPAGES field, so every page shows the total. Use 'Trang X / Y'."""
    fld = lambda instr: (f'<w:fldSimple w:instr=" {instr} \\* MERGEFORMAT ">{run("1", size=20)}</w:fldSimple>')
    para = (f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>{run("Trang ", size=20)}{fld("PAGE")}'
            f'{run(" / ", size=20)}{fld("NUMPAGES")}</w:p>')
    head = xml[:xml.index(">", xml.index("<w:ftr")) + 1]
    return head + para + "</w:ftr>"


def _drop_docprops(tmp, names):
    """Our builds carry no docProps/, so saved_by_word() can tell them apart from Word-saved files."""
    for n in [n for n in names if n.startswith("docProps/")]:
        os.remove(os.path.join(tmp, n))
        names.remove(n)
    for part, pat in [("[Content_Types].xml", r'<Override PartName="/docProps/[^"]+"[^>]*/>'),
                      ("_rels/.rels", r'<Relationship [^>]*Target="docProps/[^"]+"[^>]*/>')]:
        fp = os.path.join(tmp, part)
        with open(fp, encoding="utf-8") as f:
            x = f.read()
        with open(fp, "w", encoding="utf-8", newline="") as f:
            f.write(re.sub(pat, "", x))


def build_form(out, title_lines, code, blocks, force=False):
    if saved_by_word(out) and not force and "--force" not in sys.argv:
        sys.exit(f"Refusing to overwrite {out}: it has been edited and saved in Word. "
                 "Copy those edits into the script first, then re-run with --force.")
    tmp = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(FRAME) as z:
            names = z.namelist()
            z.extractall(tmp)
        for name, fn in [("word/header2.xml", lambda x: _fix_header(x, title_lines, code)),
                         ("word/document.xml", lambda x: x[:x.index("<w:body>") + 8] + "".join(blocks)
                          + x[x.rindex("<w:sectPr"):]),
                         ("word/footer1.xml", _page_footer), ("word/footer2.xml", _page_footer)]:
            fp = os.path.join(tmp, name)
            with open(fp, encoding="utf-8") as f:
                x = f.read()
            with open(fp, "w", encoding="utf-8", newline="") as f:
                f.write(fn(x))
        _drop_docprops(tmp, names)
        os.makedirs(os.path.dirname(out), exist_ok=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for n in names:
                z.write(os.path.join(tmp, n), n)
    finally:
        shutil.rmtree(tmp)
    print("Wrote", out)
