"""Swap only the per-system figures of Điều 2.1 in a contract .docx after its appendix was rebuilt.
Totals, VAT, words and the 5% guarantee must be unchanged (asserted), so nothing else in the contract moves.
Usage: python update_rows.py <in.docx> <out.docx> <old_summary.json> <new_summary.json> <split|total>
  split: the table has VT | NC | Tổng per row (v3 style); total: only Tổng per row (v5 style)."""
import json, os, re, sys, zipfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'scripts'))
from vn_number import dots, rnd

src, dst, old, new, mode = sys.argv[1:6]
rows = lambda p: [(rnd(vt), rnd(nc)) for vt, nc in json.load(open(p, encoding='utf8'))['sheets'].values()]
o, n = rows(old), rows(new)
assert sum(a + b for a, b in o) == sum(a + b for a, b in n), 'pre-VAT total changed'
if mode == 'split':
    assert sum(a for a, _ in o) == sum(a for a, _ in n) and sum(b for _, b in o) == sum(b for _, b in n), 'VT/NC split changed'

z = zipfile.ZipFile(src)
x = z.read('word/document.xml').decode('utf8')
pairs = []
for (ov, on), (nv, nn) in zip(o, n):
    pairs += [(ov + on, nv + nn)] + ([(ov, nv), (on, nn)] if mode == 'split' else [])
tokens = {}
for i, (a, b) in enumerate(pairs):              # two passes via placeholders: an old figure may equal a new one
    old_s = f'>{dots(a)}<'
    assert x.count(old_s) == 1, (dots(a), x.count(old_s))
    x = x.replace(old_s, f'>@@{i}@@<'); tokens[f'>@@{i}@@<'] = f'>{dots(b)}<'
for t, v in tokens.items():
    x = x.replace(t, v)
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
    for info in z.infolist():
        zo.writestr(info, x.encode('utf8') if info.filename == 'word/document.xml' else z.read(info.filename))
for (a, b) in pairs:
    print(f'{dots(a):>18} -> {dots(b):>18}')
