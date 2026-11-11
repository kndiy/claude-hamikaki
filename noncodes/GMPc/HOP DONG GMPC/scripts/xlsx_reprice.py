"""Reprice a GMPc M&E estimate workbook by editing its XML in place, so the copy keeps
everything of the original (styles, print setup, manual calc mode) and every formula
cell carries its correctly computed cached value (the workbooks use calcMode="manual").

Unit prices F (vật tư) and G (nhân công) of item rows are multiplied by k and rounded to
1.000 đ; rows whose unit (column D) is in keep_units get the fixed factor keep_k instead
(default 1 = prices kept). Quantities (E) are unchanged unless --qty-k is given. k is found so the grand total
(TH DU TOAN!F20) lands on the target; no residual is forced, so the total stays a natural,
non-round figure.

Usage: python scripts/xlsx_reprice.py <in.xlsx> <out.xlsx> <target_total> <keep_units> [edits.json] [summary.json] [keep_k]
  edits.json: {"cells": {"<sheet>": {"A2": "text", ...}}, "merge": {"<sheet>": ["B9:H9"]},
               "words_cell": ["TH DU TOAN", "B21", "Bằng chữ: {words} đồng./."]}
  summary.json gets VT/NC per sheet and the total; <summary>_rows.csv lists changed rows.

Options (after the positional arguments):
  --qty-k 1.10 --qty-units m,mét,m2,kg
      quantities (E, numeric constants only) of non-kept rows whose unit is in qty-units are multiplied by
      qty-k and rounded DOWN (integers stay integers, decimals to 0,1), so a row never exceeds qty-k.
  --vt-nc <VT>,<NC>
      separate factors for F and G so the contract's VT and NC totals (sum of per-sheet rounded values) hit
      these figures; target_total is then ignored. Keeps a VT 10% / NC 8% VAT unchanged.
  --exact
      the residual left by 1.000-đ rounding goes into the largest 'lô' row with qty 1, so the contract
      figures (per-sheet rounded VT + NC) equal the target exactly.
"""
import csv, json, math, re, sys, unicodedata, zipfile
from html import escape
import openpyxl
from vn_number import words

argv, opts = [], {}
it = iter(sys.argv[1:])
for a in it:
    if a == '--exact': opts['exact'] = True
    elif a.startswith('--'): opts[a[2:]] = next(it)
    else: argv.append(a)
src, dst, target, keep_arg = argv[0], argv[1], float(argv[2]), argv[3]
edits = json.load(open(argv[4], encoding='utf8')) if len(argv) > 4 else {}
summary_path = argv[5] if len(argv) > 5 else None
keep_k = float(argv[6]) if len(argv) > 6 else 1.0
qty_k = float(opts.get('qty-k', 1))
vt_nc = [float(v) for v in opts['vt-nc'].split(',')] if 'vt-nc' in opts else None
norm = lambda u: unicodedata.normalize('NFC', str(u).strip().lower())
keep_units = {norm(u) for u in keep_arg.split(',')}
qty_units = {norm(u) for u in opts.get('qty-units', '').split(',') if u}

wb = openpyxl.load_workbook(src)          # formulas (shared formulas come translated)
cells = {ws.title: {c.coordinate: c.value for row in ws.iter_rows() for c in row
                    if c.value is not None} for ws in wb.worksheets}

# ---- tiny evaluator for the formula shapes these workbooks use
REF = r"(?:'([^']+)'!)?\$?([A-Z]{1,3})\$?(\d+)"
def col2n(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n
def n2col(n):
    s = ''
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s
def num(v):
    if isinstance(v, (int, float)): return v
    try: return float(v)
    except (TypeError, ValueError): return 0
memo = {}
def val(sheet, ref):
    key = (sheet, ref)
    if key not in memo:
        v = cells[sheet].get(ref)
        memo[key] = evaluate(sheet, v[1:]) if isinstance(v, str) and v.startswith('=') else v
    return memo[key]
def rng(sheet, a, b, skip_subtotal=False):
    (ca, ra), (cb, rb) = [re.match(r'([A-Z]+)(\d+)', x.replace('$', '')).groups() for x in (a, b)]
    out = []
    for c in range(col2n(ca), col2n(cb) + 1):
        for r in range(int(ra), int(rb) + 1):
            ref = f'{n2col(c)}{r}'
            if skip_subtotal and 'SUBTOTAL' in str(cells[sheet].get(ref, '')):
                continue
            v = val(sheet, ref)
            if isinstance(v, (int, float)): out.append(v)   # SUM/SUBTOTAL skip text
    return out
def evaluate(sheet, f):
    f = f.lstrip('+')
    m = re.fullmatch(REF, f)
    if m:                                 # plain reference (=B16): keep text as text, not 0
        return val(m.group(1) or sheet, m.group(2) + m.group(3))
    m = re.fullmatch(r'SUBTOTAL\(9,([A-Z$\d]+):([A-Z$\d]+)\)', f)
    if m: return sum(rng(sheet, *m.groups(), skip_subtotal=True))
    expr = re.sub(r'SUM\(([A-Z$\d]+):([A-Z$\d]+)\)',
                  lambda m: repr(sum(rng(sheet, *m.groups()))), f)
    expr = re.sub(REF, lambda m: repr(num(val(m.group(1) or sheet, m.group(2) + m.group(3)))), expr)
    assert re.fullmatch(r'[\d.eE+\-*/() ]+', expr), (sheet, f, expr)
    return eval(expr)

# ---- item rows: between row 7 and the SUBTOTAL row of each item sheet, with a J formula
items = []
for name in wb.sheetnames[2:]:
    ws = wb[name]
    end = next(r for r in range(7, ws.max_row + 1) if 'SUBTOTAL' in str(ws.cell(r, 10).value))
    for r in range(7, end):
        if str(ws.cell(r, 10).value).startswith('='):
            items.append((name, r, norm(ws.cell(r, 4).value) in keep_units))
orig_total = val('TH DU TOAN', 'F20')

# the evaluator must reproduce every cached value of the original before we trust it
cached = openpyxl.load_workbook(src, data_only=True)
for s in cells:
    for ref, v in cells[s].items():
        if isinstance(v, str) and v.startswith('='):
            c = cached[s][ref].value
            e = val(s, ref)
            ok = (num(c) == 0 and num(e) == 0) if not isinstance(c, (int, float)) else abs(num(e) - c) < 0.01
            assert ok, (s, ref, v, c, e)
base = {(s, r): (cells[s].get(f'F{r}'), cells[s].get(f'G{r}')) for s, r, _ in items}

base_q = {(s, r): cells[s].get(f'E{r}') for s, r, _ in items}
qty_changed = set()
if qty_k != 1:
    for s, r, kept in items:
        q = base_q[(s, r)]
        if kept or norm(cells[s].get(f'D{r}')) not in qty_units or not isinstance(q, (int, float)):
            continue
        nq = math.floor(q * qty_k + 1e-9) if float(q).is_integer() else math.floor(q * qty_k * 10 + 1e-9) / 10
        if nq != q:
            cells[s][f'E{r}'] = nq
            qty_changed.add((s, r))

r1k = lambda v: int(round(v / 1000.0)) * 1000
adj = {}                                  # (sheet, row, col) -> đ added by --exact
def apply(kv, kn=None):
    kn = kv if kn is None else kn
    memo.clear()
    for s, r, kept in items:
        for col, v, k in zip('FG', base[(s, r)], (kv, kn)):
            if isinstance(v, (int, float)):
                f = keep_k if kept else k
                cells[s][f'{col}{r}'] = (v if f == 1 else r1k(v * f)) + adj.get((s, r, col), 0)
    return val('TH DU TOAN', 'F20')
def contract():                           # per-sheet rounded VT / NC, as the contract table prints them
    vt = nc = 0
    for s in dict.fromkeys(s for s, _, _ in items):
        rows = [(num(val(s, f'E{r}')), num(cells[s].get(f'F{r}')), num(cells[s].get(f'G{r}')))
                for s2, r, _ in items if s2 == s]
        vt += math.floor(sum(q * f for q, f, _ in rows) + 0.5)
        nc += math.floor(sum(q * g for q, _, g in rows) + 0.5)
    return vt, nc
def bisect(f):
    lo, hi = 0.5, 5.0
    for _ in range(60):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if f(mid) < 0 else (lo, mid)
    return hi

if vt_nc:
    kv = bisect(lambda k: (apply(k, 1.0), contract()[0])[1] - vt_nc[0])
    kn = bisect(lambda k: (apply(kv, k), contract()[1])[1] - vt_nc[1])
    k = [kv, kn]
else:
    k = kv = kn = bisect(lambda k: apply(k) - target)
total = apply(kv, kn)
if opts.get('exact'):
    _, ls, lr = max((num(base[(s, r)][0]), s, r) for s, r, kept in items
                    if not kept and norm(cells[s].get(f'D{r}')) == 'lô' and num(val(s, f'E{r}')) == 1)
    for _ in range(5):
        vt, nc = contract()
        dv, dn = (vt_nc[0] - vt, vt_nc[1] - nc) if vt_nc else (target - vt - nc, 0)
        if not dv and not dn:
            break
        for col, d in (('F', dv), ('G', dn)):
            if d:
                adj[(ls, lr, col)] = adj.get((ls, lr, col), 0) + int(d)
        total = apply(kv, kn)
    assert not dv and not dn, ('--exact did not converge', dv, dn)

# ---- text edits
for s, d in edits.get('cells', {}).items():
    cells[s].update(d)
if 'words_cell' in edits:
    s, ref, tmpl = edits['words_cell']
    cells[s][ref] = tmpl.format(words=words(total))
memo.clear()

# ---- patch the XML
z = zipfile.ZipFile(src)
wbx = z.read('xl/workbook.xml').decode('utf8')
rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
rid2target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels)) | \
             {a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)}
sheet_xml = {}
for nm, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
    nm = nm.replace('&amp;', '&')
    sheet_xml[nm] = 'xl/' + rid2target[rid].lstrip('/').removeprefix('xl/')

def fmt(v):
    return repr(int(v)) if float(v).is_integer() else repr(float(v))

out_parts = {}
for s, path in sheet_xml.items():
    x = z.read(path).decode('utf8')
    text_edits = dict(edits.get('cells', {}).get(s, {}))
    if 'words_cell' in edits and edits['words_cell'][0] == s:
        text_edits[edits['words_cell'][1]] = cells[s][edits['words_cell'][1]]
    def fix(m):
        attrs, inner, ref = m.group(1), m.group(2) or '', re.search(r'r="([A-Z]+\d+)"', m.group(1)).group(1)
        if ref in text_edits:
            a = re.sub(r'\s*t="[^"]*"', '', attrs)
            return f'<c{a} t="inlineStr"><is><t xml:space="preserve">{escape(text_edits.pop(ref), quote=False)}</t></is></c>'
        if '<f' in inner:                                   # formula: refresh cached value
            v = val(s, ref)
            if not isinstance(v, (int, float)):
                return m.group(0)
            a = re.sub(r'\s*t="[^"]*"', '', attrs)
            inner = re.sub(r'<v>[^<]*</v>|<v/>', '', inner) + f'<v>{fmt(v)}</v>'
            return f'<c{a}>{inner}</c>'
        if re.match(r'[EFG]\d+$', ref) and isinstance(cells[s].get(ref), (int, float)) and 't="' not in attrs:
            return f'<c{attrs}>' + re.sub(r'<v>[^<]*</v>', f'<v>{fmt(cells[s][ref])}</v>', inner) + '</c>'
        return m.group(0)
    x = re.sub(r'<c( r="[A-Z]+\d+"[^>]*?)(?:/>|>(.*?)</c>)', fix, x, flags=re.S)
    for ref, t in text_edits.items():                       # cell absent from the XML: insert it
        col, rn = re.match(r'([A-Z]+)(\d+)', ref).groups()
        sty = edits.get('style_from', {}).get(s, {}).get(ref)
        sidx = re.search(rf'<c r="{sty}"[^>]*?s="(\d+)"', x).group(1) if sty else None
        cell = (f'<c r="{ref}"' + (f' s="{sidx}"' if sidx else '') +
                f' t="inlineStr"><is><t xml:space="preserve">{escape(t, quote=False)}</t></is></c>')
        m = re.search(rf'<row r="{rn}"([^>]*?)(/>|>(.*?)</row>)', x, re.S)
        if m:
            body = m.group(3) or ''
            refs = re.findall(r'<c r="([A-Z]+)\d+"', body)
            pos = next((i for i, c in enumerate(refs) if col2n(c) > col2n(col)), None)
            if pos is None:
                body = body + cell
            else:
                i = [mm.start() for mm in re.finditer(r'<c r="', body)][pos]
                body = body[:i] + cell + body[i:]
            attrs = re.sub(r'\s*spans="[^"]*"', '', m.group(1))
            x = x[:m.start()] + f'<row r="{rn}"{attrs}>{body}</row>' + x[m.end():]
        else:
            nxt = next(mm for mm in re.finditer(r'<row r="(\d+)"', x) if int(mm.group(1)) > int(rn))
            x = x[:nxt.start()] + f'<row r="{rn}">{cell}</row>' + x[nxt.start():]
    for mref in edits.get('merge', {}).get(s, []):
        cnt = int(re.search(r'<mergeCells count="(\d+)"', x).group(1))
        x = re.sub(r'<mergeCells count="\d+">', f'<mergeCells count="{cnt + 1}"><mergeCell ref="{mref}"/>', x)
    out_parts[path] = x

with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
    for info in z.infolist():
        data = out_parts[info.filename].encode('utf8') if info.filename in out_parts else z.read(info.filename)
        zo.writestr(info, data)

# ---- summary
sheets = {}
for s, r, _ in items:
    q = num(val(s, f'E{r}'))
    vt, nc = sheets.setdefault(s, [0, 0])
    sheets[s] = [vt + q * num(cells[s].get(f'F{r}')), nc + q * num(cells[s].get(f'G{r}'))]
out = {'k': k, 'keep_k': keep_k, 'qty_k': qty_k, 'qty_rows_changed': len(qty_changed),
       'exact_adjust': {f'{s}!{c}{r}': v for (s, r, c), v in adj.items()},
       'orig_total': orig_total, 'new_total': total, 'sheets': sheets}
print(json.dumps(out, ensure_ascii=False, indent=1))
if summary_path:
    json.dump(out, open(summary_path, 'w', encoding='utf8'), ensure_ascii=False, indent=1)
    with open(summary_path.rsplit('.', 1)[0] + '_rows.csv', 'w', encoding='utf-8-sig', newline='') as fh:
        w = csv.writer(fh)
        w.writerow(['sheet', 'row', 'item', 'unit', 'qty', 'VT_old', 'NC_old', 'VT_new', 'NC_new', 'qty_new'])
        for s, r, kept in items:
            if (not kept or keep_k != 1) and any(isinstance(v, (int, float)) and v for v in base[(s, r)]):
                w.writerow([s, r, str(cells[s].get(f'B{r}')).splitlines()[0][:80], cells[s].get(f'D{r}'),
                            num(base_q[(s, r)]), *base[(s, r)], cells[s].get(f'F{r}'), cells[s].get(f'G{r}'),
                            num(val(s, f'E{r}'))])
