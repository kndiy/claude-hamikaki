"""Print the paragraph text of a .docx (no pandoc needed).
Usage: python scripts/dump_docx.py file.docx"""
import re, sys, zipfile, html

x = zipfile.ZipFile(sys.argv[1]).read('word/document.xml').decode('utf8')
for p in re.findall(r'<w:p[ >].*?</w:p>', x, re.S):
    t = ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', p))
    if t.strip():
        print(html.unescape(t))
