"""Print a .docx's text in body order (paragraphs and table cells),
wrapping highlighted / shaded runs as [[HL:...]]. Usage: dump_highlight.py <file.docx>"""
import sys, zipfile
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def run_hl(r):
    rpr = r.find(W + "rPr")
    if rpr is None:
        return None
    h = rpr.find(W + "highlight")
    if h is not None and h.get(W + "val") not in (None, "none"):
        return h.get(W + "val")
    s = rpr.find(W + "shd")
    if s is not None and s.get(W + "fill") not in (None, "auto", "FFFFFF"):
        return "shd" + s.get(W + "fill")
    return None


def para_text(p):
    out, cur = [], None
    for r in p.iter(W + "r"):
        t = "".join(x.text or "" for x in r.iter(W + "t", W + "tab"))
        t = t or ("\t" if r.find(W + "tab") is not None else "")
        if not t:
            continue
        hl = run_hl(r)
        if hl != cur:
            if cur:
                out.append("]]")
            if hl:
                out.append("[[HL:")
            cur = hl
        out.append(t)
    if cur:
        out.append("]]")
    return "".join(out)


def walk(el, depth=0):
    for ch in el:
        if ch.tag == W + "p":
            t = para_text(ch)
            if t.strip():
                print("  " * depth + t)
        elif ch.tag == W + "tbl":
            for i, tr in enumerate(ch.iter(W + "tr")):
                cells = []
                for tc in tr.findall(W + "tc"):
                    cells.append(" / ".join(para_text(p) for p in tc.iter(W + "p") if para_text(p).strip()))
                print("  " * depth + f"|R{i}| " + " || ".join(cells))
        elif ch.tag in (W + "sdt", W + "sdtContent", W + "customXml"):
            walk(ch, depth)


with zipfile.ZipFile(sys.argv[1]) as z:
    root = etree.fromstring(z.read("word/document.xml"))
walk(root.find(W + "body"))
