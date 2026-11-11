"""Re-price the user's edited contract (raw/261028_HSC_Hop dong thi cong_85ty.docx, unpacked) in place.
Usage: python update_price.py <unpacked_dir> <old_summary.json> <new_summary.json>
Only the money figures in Điều 2.1 / Điều 3 and the words line change; every other edit of the user is kept.
VAT is one 8% rate on the total (user, 2026-10-28)."""
import json, os, re, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'scripts'))
from vn_number import dots, rnd, words

UNP, OLD, NEW = sys.argv[1:4]
VAT_RATE = 0.08

def figures(path):
    s = json.load(open(path, encoding='utf8'))
    rows = [rnd(vt) + rnd(nc) for vt, nc in s['sheets'].values()]
    total = sum(rows)
    assert total == rnd(s['new_total']), (total, s['new_total'])
    vat = rnd(total * VAT_RATE)
    return rows, total, vat, total + vat, rnd(total * 0.05)

o_rows, o_total, o_vat, o_grand, o_bl = figures(OLD)
n_rows, n_total, n_vat, n_grand, n_bl = figures(NEW)

p = f'{UNP}/word/document.xml'
x = open(p, encoding='utf8').read()

def sub(old, new):
    global x
    assert x.count(old) == 1, (old, x.count(old))
    x = x.replace(old, new)

for a, b in zip(o_rows + [o_total, o_vat, o_grand], n_rows + [n_total, n_vat, n_grand]):
    sub(f'>{dots(a)}<', f'>{dots(b)}<')
sub(f'{dots(o_bl)} VNĐ', f'{dots(n_bl)} VNĐ')
sub(f'Bằng chữ: {words(o_grand)} đồng./.)', f'Bằng chữ: {words(n_grand)} đồng./.)')
open(p, 'w', encoding='utf8').write(x)

print('TRƯỚC VAT', dots(n_total), '| VAT', dots(n_vat), '| SAU VAT', dots(n_grand), '| BẢO LÃNH 5%', dots(n_bl))
print(words(n_grand))
