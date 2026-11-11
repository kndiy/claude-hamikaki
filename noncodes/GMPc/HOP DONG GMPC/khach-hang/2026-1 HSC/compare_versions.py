"""Bảng so sánh giá: dự toán gốc (54,33 tỷ) vs V3-KL (71,64 tỷ trước VAT) vs V5-KL (85 tỷ sau VAT).

V3-KL / V5-KL (2026-10-30): phần tăng chia vào khối lượng (dòng m, mét, m², kg: tối đa +10%, làm tròn xuống)
và đơn giá vật tư/không đếm được; thiết bị đếm được giữ nguyên số lượng, đơn giá gốc ở V3 và +15% ở V5.
Gốc: raw/260521_HSC_DU TOAN THIET KE HE THONG M&E.xlsx (giá trị cache).
V3 / V5: output/261030_HSC_PL01_DU TOAN HE THONG M&E_v3-KL.xlsx / _v5-KL.xlsx (giá trị cache, đã kiểm bằng Excel).
Mục A lấy đúng số trong HĐ (Điều 2.1) = VT, NC làm tròn theo từng hệ trong _v3kl/_v5kl_summary.json.

Usage (từ thư mục dự án): python "khach-hang/2026-1 HSC/compare_versions.py"
Sau đó mở bằng Excel (COM) và CalculateFull + Save để có giá trị cache.
"""
import json, os, sys, unicodedata
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', '..', 'scripts'))
from vn_number import rnd

OUTDIR = os.path.join(HERE, 'output')
RAW = os.path.join(HERE, 'raw', '260521_HSC_DU TOAN THIET KE HE THONG M&E.xlsx')
V3 = os.path.join(OUTDIR, '261030_HSC_PL01_DU TOAN HE THONG M&E_v3-KL.xlsx')
V5 = os.path.join(OUTDIR, '261030_HSC_PL01_DU TOAN HE THONG M&E_v5-KL.xlsx')
OUT = os.path.join(OUTDIR, '261030_HSC_So sanh gia Goc-V3-V5.xlsx')
KEEP = 'cái,bộ,hệ,set,chiếc,tủ,bình,chai,mẫu,None'
norm = lambda u: unicodedata.normalize('NFC', str(u).strip().lower())
keep_units = {norm(u) for u in KEEP.split(',')}

SHEETS = ['1.ĐHKK_NX', '2.ĐIỆN ĐL-CS ', '3.ĐIỆN NHẸ ', '4.PANEL', '5.EPOXY', '6. CTN-NƯỚC LÀM MỀM',
          '7. RO ', '8. BMS', '9. KHÍ Y TẾ']
NAMES = ['Hệ thống điều hòa không khí xưởng sản xuất', 'Hệ thống điện động lực - chiếu sáng',
         'Hệ thống điện nhẹ', 'Hệ thống panel', 'Hệ thống nền vinyl và epoxy',
         'Hệ thống xử lý nước thải - thoát nước thải - cấp nước sinh hoạt', 'Hệ thống nước RO',
         'Hệ thống BMS', 'Hệ thống khí y tế']
hd = lambda p: [rnd(vt) + rnd(nc) for vt, nc in json.load(open(os.path.join(OUTDIR, p), encoding='utf8'))['sheets'].values()]
HD_V3, HD_V5 = hd('_v3kl_summary.json'), hd('_v5kl_summary.json')
V3_VAT, V3_GRAND, V5_GRAND = (5735725360, 1142471680), 78516346643, 84997620410   # HĐ v3-KL / v5-KL

goc, w3, w5 = (openpyxl.load_workbook(p, data_only=True) for p in (RAW, V3, V5))
goc_f = openpyxl.load_workbook(RAW)   # công thức: xem dòng nào có cột J (Tổng thành tiền)
num = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)
def fnum(v):
    try: return float(v) if isinstance(v, str) else v   # vài ô SL lưu dạng text ('1')
    except ValueError: return v
rows = []
for s in SHEETS:
    g, a, b, gf = goc[s], w3[s], w5[s], goc_f[s]
    for r in range(1, g.max_row + 1):
        qty = fnum(g.cell(r, 5).value)
        if not num(qty) or not (num(g.cell(r, 6).value) or num(g.cell(r, 7).value)):
            continue
        unit = g.cell(r, 4).value
        name = str(g.cell(r, 2).value or '').strip().split('\n')[0][:150]
        p = lambda ws, c: fnum(ws.cell(r, c).value) if num(fnum(ws.cell(r, c).value)) else 0
        rows.append([s, r, name, unit,
                     'Thiết bị đếm được' if norm(unit) in keep_units else 'Vật tư/không đếm được',
                     qty, p(a, 5), p(b, 5),
                     p(g, 6), p(g, 7), p(a, 6), p(a, 7), p(b, 6), p(b, 7),
                     'Có' if gf.cell(r, 10).value is not None else 'Không'])

# ---------------------------------------------------------------- styles
F = 'Calibri'
font = lambda **k: Font(name=F, size=k.pop('size', 11), **k)
thin = Side(style='thin', color='999999')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
head_fill = PatternFill('solid', fgColor='1F3864')
tot_fill = PatternFill('solid', fgColor='FFF2CC')
chg_fill = PatternFill('solid', fgColor='E2EFDA')
MONEY, PCT, QTY = '#,##0;[Red]-#,##0;"-"', '0.0%;[Red]-0.0%;"-"', '#,##0.##'
BLUE = '0000FF'
L = openpyxl.utils.get_column_letter

def header(ws, row, labels, widths=None):
    for i, t in enumerate(labels, 1):
        c = ws.cell(row, i, t)
        c.font = font(bold=True, color='FFFFFF'); c.fill = head_fill; c.border = box
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[row].height = 45
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[L(i)].width = w

def style_row(ws, row, ncol, fmts, fill=None, bold=False, blue=()):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c)
        cell.border = box
        cell.font = font(bold=bold, color=BLUE if c in blue else None)
        if fill: cell.fill = fill
        if fmts.get(c): cell.number_format = fmts[c]
        cell.alignment = Alignment(vertical='center', wrap_text=c <= 2)

wb = openpyxl.Workbook()
ws = wb.active; ws.title = 'Tổng hợp'
dt = wb.create_sheet('Chi tiết')
nt = wb.create_sheet('Ghi chú')

# ---------------------------------------------------------------- Chi tiết
DH = ['Hệ thống (sheet)', 'Dòng', 'Hạng mục', 'ĐVT', 'Nhóm', 'SL gốc', 'SL V3', 'SL V5',
      'ĐG VT gốc', 'ĐG NC gốc', 'ĐG VT V3', 'ĐG NC V3', 'ĐG VT V5', 'ĐG NC V5',
      'TT gốc', 'TT V3', 'TT V5', 'V3 - gốc', '% V3', 'V5 - gốc', '% V5',
      'V3: tăng do KL', 'V3: tăng do đơn giá', 'V5: tăng do KL', 'V5: tăng do đơn giá', 'Có trong cột Tổng (J)?']
header(dt, 1, DH, [20, 6, 50, 8, 18, 9, 9, 9] + [13] * 6 + [16] * 3 + [15, 8, 15, 8] + [15] * 4 + [11])
dt.freeze_panes = 'F2'
fm = {c: MONEY for c in list(range(9, 19)) + [20] + list(range(22, 26))} | {6: QTY, 7: QTY, 8: QTY, 19: PCT, 21: PCT}
for i, d in enumerate(rows, 2):
    for c, v in enumerate(d[:14], 1):
        dt.cell(i, c, v)
    dt.cell(i, 26, d[14])
    on = f'(Z{i}="Có")'   # dự toán chỉ cộng dòng có công thức ở cột J
    dt.cell(i, 15, f'={on}*F{i}*(I{i}+J{i})')
    dt.cell(i, 16, f'={on}*G{i}*(K{i}+L{i})')
    dt.cell(i, 17, f'={on}*H{i}*(M{i}+N{i})')
    dt.cell(i, 18, f'=P{i}-O{i}'); dt.cell(i, 19, f'=IF(O{i}=0,0,R{i}/O{i})')
    dt.cell(i, 20, f'=Q{i}-O{i}'); dt.cell(i, 21, f'=IF(O{i}=0,0,T{i}/O{i})')
    dt.cell(i, 22, f'={on}*(G{i}-F{i})*(I{i}+J{i})'); dt.cell(i, 23, f'=R{i}-V{i}')
    dt.cell(i, 24, f'={on}*(H{i}-F{i})*(I{i}+J{i})'); dt.cell(i, 25, f'=T{i}-X{i}')
    style_row(dt, i, 26, fm, blue=range(6, 15), fill=tot_fill if d[14] == 'Không' else None)
    for c in (7, 8):
        if d[c - 1] != d[5]:
            dt.cell(i, c).fill = chg_fill      # khối lượng đã đổi
last = len(rows) + 1
t = last + 1
dt.cell(t, 3, 'TỔNG (theo phụ lục)')
for c in [15, 16, 17, 18, 20, 22, 23, 24, 25]:
    dt.cell(t, c, f'=SUM({L(c)}2:{L(c)}{last})')
dt.cell(t, 19, f'=R{t}/O{t}'); dt.cell(t, 21, f'=T{t}/O{t}')
style_row(dt, t, 26, fm, fill=tot_fill, bold=True)
dt.auto_filter.ref = f'A1:Z{last}'

R = lambda col: f"'Chi tiết'!${col}$2:${col}${last}"

# ---------------------------------------------------------------- Tổng hợp
ws['A1'] = 'HSC - SO SÁNH GIÁ HỢP ĐỒNG THI CÔNG M&E: V3 VÀ V5 SO VỚI DỰ TOÁN GỐC'
ws['A1'].font = font(bold=True, size=14)
ws['A2'] = ('Gốc = dự toán thiết kế 21/05/2026 (chưa VAT). V3 = HĐ 71,64 tỷ trước VAT (VAT VT 10%, NC 8%). '
            'V5 = HĐ 85 tỷ sau VAT (VAT 8% trên tổng). Phần tăng chia vào khối lượng (ống, dây, m², kg: tối đa +10%) '
            'và đơn giá vật tư; thiết bị đếm được giữ nguyên số lượng, đơn giá gốc ở V3 và +15% ở V5. Đơn vị: đồng. Số màu xanh = từ file nguồn.')
ws['A2'].font = font(italic=True, size=10)
ws.merge_cells('A2:M2'); ws['A2'].alignment = Alignment(wrap_text=True, vertical='top'); ws.row_dimensions[2].height = 45

ws['A4'] = 'A. THEO GIÁ TRỊ TRONG HỢP ĐỒNG (Điều 2.1)'; ws['A4'].font = font(bold=True, size=12)
H = ['STT', 'Hệ thống', 'Gốc', 'V3', 'V3 - gốc', '% V3', 'V5', 'V5 - gốc', '% V5',
     'V3: tăng do KL', 'V3: tăng do đơn giá', 'V5: tăng do KL', 'V5: tăng do đơn giá']
header(ws, 5, H, [6, 46, 17, 17, 16, 8, 17, 16, 8, 15, 16, 15, 16])
fmA = {c: MONEY for c in (3, 4, 5, 7, 8, 10, 11, 12, 13)} | {6: PCT, 9: PCT}
def deltas(r):
    ws.cell(r, 5, f'=D{r}-C{r}'); ws.cell(r, 6, f'=IF(C{r}=0,0,E{r}/C{r})')
    ws.cell(r, 8, f'=G{r}-C{r}'); ws.cell(r, 9, f'=IF(C{r}=0,0,H{r}/C{r})')
for i, (s, nm) in enumerate(zip(SHEETS, NAMES)):
    r = 6 + i
    ws.cell(r, 1, i + 1); ws.cell(r, 2, nm)
    ws.cell(r, 3, f'=SUMIFS({R("O")},{R("A")},"{s}")')
    ws.cell(r, 4, HD_V3[i]); ws.cell(r, 7, HD_V5[i])
    for c, q in zip((10, 11, 12, 13), 'VWXY'):
        ws.cell(r, c, f'=SUMIFS({R(q)},{R("A")},"{s}")')
    deltas(r); style_row(ws, r, 13, fmA, blue=(4, 7))
TP, VAT, SV = 15, 16, 17
ws.cell(TP, 2, 'Tổng cộng trước VAT')
for c in 'CDGJKLM': ws[f'{c}{TP}'] = f'=SUM({c}6:{c}14)'
deltas(TP); style_row(ws, TP, 13, fmA, fill=tot_fill, bold=True)
ws.cell(VAT, 2, 'Thuế GTGT')
ws[f'C{VAT}'] = f'=ROUND(C{TP}*0.08,0)'
ws[f'C{VAT}'].comment = Comment('Dự toán gốc chưa có VAT. Giả định 8% trên tổng như V5 để so sau thuế.', 'Claude')
ws[f'D{VAT}'] = f'={V3_VAT[0]}+{V3_VAT[1]}'
ws[f'D{VAT}'].comment = Comment('HĐ V3: VAT vật tư 10% (5.735.725.360) + nhân công 8% (1.142.471.680).', 'Claude')
ws[f'G{VAT}'] = f'=ROUND(G{TP}*0.08,0)'
deltas(VAT); style_row(ws, VAT, 9, fmA)
ws.cell(SV, 2, 'Tổng cộng sau VAT')
for c in 'CDG': ws[f'{c}{SV}'] = f'={c}{TP}+{c}{VAT}'
deltas(SV); style_row(ws, SV, 9, fmA, fill=tot_fill, bold=True)
ws.cell(18, 2, 'Kiểm: HĐ ghi sau VAT'); ws['D18'] = V3_GRAND; ws['G18'] = V5_GRAND; ws['C18'] = 'chưa VAT'
style_row(ws, 18, 9, fmA, blue=(4, 7))
ws['D19'] = '=D17-D18'; ws['G19'] = '=G17-G18'; ws.cell(19, 2, 'Lệch (phải = 0)')
style_row(ws, 19, 9, fmA)

# B. VT / NC
b0 = 22
ws.cell(b0 - 1, 1, 'B. TÁCH VẬT TƯ / NHÂN CÔNG (theo phụ lục dự toán)').font = font(bold=True, size=12)
header(ws, b0, ['STT', 'Hệ thống', 'VT gốc', 'VT V3', '% V3', 'VT V5', '% V5', 'NC gốc', 'NC V3', '% V3', 'NC V5', '% V5'])
fmB = {c: MONEY for c in (3, 4, 6, 8, 9, 11)} | {5: PCT, 7: PCT, 10: PCT, 12: PCT}
def sp(qcol, pcol, s): return f'SUMPRODUCT(({R("A")}="{s}")*{R(qcol)}*{R(pcol)})'
def bdel(r):
    ws.cell(r, 5, f'=IF(C{r}=0,0,D{r}/C{r}-1)'); ws.cell(r, 7, f'=IF(C{r}=0,0,F{r}/C{r}-1)')
    ws.cell(r, 10, f'=IF(H{r}=0,0,I{r}/H{r}-1)'); ws.cell(r, 12, f'=IF(H{r}=0,0,K{r}/H{r}-1)')
for i, (s, nm) in enumerate(zip(SHEETS, NAMES)):
    r = b0 + 1 + i
    ws.cell(r, 1, i + 1); ws.cell(r, 2, nm)
    for c, (q, pc) in zip((3, 4, 6, 8, 9, 11), ('FI', 'GK', 'HM', 'FJ', 'GL', 'HN')):
        ws.cell(r, c, '=' + sp(q, pc, s))
    bdel(r); style_row(ws, r, 12, fmB)
bt = b0 + 10
ws.cell(bt, 2, 'Tổng cộng')
for c in (3, 4, 6, 8, 9, 11):
    ws.cell(bt, c, f'=SUM({L(c)}{b0+1}:{L(c)}{bt-1})')
bdel(bt); style_row(ws, bt, 12, fmB, fill=tot_fill, bold=True)
ws.cell(bt + 1, 2, 'VT hệ BMS gồm cả Tủ DDC 25.000.000 đ không có trong tổng (xem Ghi chú).').font = font(italic=True, size=10)

# C. Nhóm
c0 = bt + 4
ws.cell(c0 - 1, 1, 'C. THEO NHÓM HẠNG MỤC').font = font(bold=True, size=12)
header(ws, c0, ['', 'Nhóm', 'Gốc', 'V3', 'V3 - gốc', '% V3', 'V5', 'V5 - gốc', '% V5',
                'V3: tăng do KL', 'V3: tăng do đơn giá', 'V5: tăng do KL', 'V5: tăng do đơn giá'])
for j, g in enumerate(['Thiết bị đếm được', 'Vật tư/không đếm được']):
    r = c0 + 1 + j
    ws.cell(r, 2, g)
    for c, q in zip((3, 4, 7, 10, 11, 12, 13), 'OPQVWXY'):
        ws.cell(r, c, f'=SUMIFS({R(q)},{R("E")},"{g}")')
    deltas(r); style_row(ws, r, 13, fmA)
r = c0 + 3; ws.cell(r, 2, 'Tổng cộng (phụ lục)')
for c in 'CDGJKLM': ws[f'{c}{r}'] = f'=SUM({c}{c0+1}:{c}{c0+2})'
deltas(r); style_row(ws, r, 13, fmA, fill=tot_fill, bold=True)
ws.cell(r + 1, 2, 'Lệch so với tổng HĐ (mục A, làm tròn)')
for c in 'CDG': ws[f'{c}{r+1}'] = f'=ROUND({c}{r}-{c}{TP},0)'
style_row(ws, r + 1, 9, fmA)
ws.cell(r + 2, 2, ('Khối lượng: dòng m, mét, m², kg ×1,10 làm tròn xuống (211 dòng). Đơn giá vật tư/không đếm được: '
                   'V3 VT ×1,4830, NC ×1,5225; V5 ×1,5964 (làm tròn 1.000 đ/đơn giá). Thiết bị: V3 giữ nguyên, V5 ×1,15.'))
ws.cell(r + 2, 2).font = font(italic=True, size=10)
ws.freeze_panes = 'C6'

# ---------------------------------------------------------------- Ghi chú
notes = [
    ('Nguồn', ''),
    ('Gốc', r'raw/260521_HSC_DU TOAN THIET KE HE THONG M&E.xlsx (NAS: N:\02. Dự án\1. Hợp đồng khách hàng\Năm 2026\1.Merap\3. Hồ sơ thầu thi công\), 54.333.626.667 đ chưa VAT.'),
    ('V3', 'output/261030_HSC_PL01_DU TOAN HE THONG M&E_v3-KL.xlsx và 261030_HSC_Hop dong thi cong_v3-KL.docx: 71.638.149.603 trước VAT (VT 57.357.253.597 + NC 14.280.896.006, giữ như V3 cũ nên VAT không đổi), 78.516.346.643 sau VAT.'),
    ('V5', 'output/261030_HSC_PL01_DU TOAN HE THONG M&E_v5-KL.xlsx và 261030_HSC_Hop dong thi cong_v5-KL.docx: 78.701.500.380 trước VAT, 84.997.620.410 sau VAT.'),
    ('', ''),
    ('Cách phân bổ', ''),
    ('Thiết bị', 'ĐVT thuộc {cái, bộ, hệ, set, chiếc, tủ, bình, chai, mẫu, trống} (điều hòa, AHU, chiller, tủ điện, bơm…): giữ nguyên số lượng. Đơn giá: V3 giữ giá gốc; V5 ×1,15 (user, 2026-10-30), làm tròn 1.000 đ.'),
    ('Khối lượng', 'Dòng ĐVT m, mét, m², kg: SL ×1,10 rồi làm tròn xuống (số nguyên giữ nguyên, số lẻ đến 0,1), nên không dòng nào vượt +10%. Dòng SL < 10 có thể không đổi. Phụ lục ghi chú: khối lượng ống, dây cáp, vật tư phụ đã bao gồm hao hụt và dự phòng thi công. Ô SL đã đổi tô xanh lá ở sheet Chi tiết.'),
    ('Đơn giá', 'Toàn bộ dòng vật tư/không đếm được (cả lô, gói, công tác…): đơn giá ×k, làm tròn 1.000 đ. Phần lẻ do làm tròn dồn vào 1 dòng "lô" SL 1 (ĐHKK dòng 193) để tổng đúng từng đồng.'),
    ('Tăng do KL / do đơn giá', 'Tăng do KL = (SL mới − SL gốc) × đơn giá gốc. Tăng do đơn giá = phần còn lại = SL mới × (đơn giá mới − đơn giá gốc).'),
    ('', ''),
    ('Lưu ý khi đọc', ''),
    ('VAT', 'V3 tính VAT vật tư 10% + nhân công 8%; V5 tính 8% trên tổng. So sánh công bằng nhất là dòng "Tổng cộng trước VAT".'),
    ('VAT gốc', 'Dự toán gốc không có VAT; dòng VAT của cột Gốc là giả định 8% như V5.'),
    ('BMS dòng 46', 'LỖI TRONG DỰ TOÁN GỐC: Tủ DDC 25.000.000 đ có ở cột Thành tiền vật tư (H46) nhưng ô Tổng thành tiền J46 để trống, nên KHÔNG được cộng vào tổng hệ BMS và tổng dự toán; lỗi này nằm trong cả V3 và V5. Ở sheet Chi tiết, dòng tô vàng, cột Z = "Không", TT tính = 0 để khớp tổng HĐ; mục B vẫn tính dòng này như cột VT của phụ lục.'),
    ('Dựng lại', 'python "khach-hang/2026-1 HSC/compare_versions.py", sau đó mở bằng Excel, tính lại toàn bộ và lưu.'),
]
nt.column_dimensions['A'].width = 22; nt.column_dimensions['B'].width = 120
for i, (a, b) in enumerate(notes, 1):
    nt.cell(i, 1, a).font = font(bold=True); nt.cell(i, 2, b).font = font()
    nt.cell(i, 2).alignment = Alignment(wrap_text=True, vertical='top')
    nt.cell(i, 1).alignment = Alignment(vertical='top')

for w in (ws, dt):
    w.sheet_view.zoomScale = 90
ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(OUT)
print(OUT, len(rows), 'dòng')
