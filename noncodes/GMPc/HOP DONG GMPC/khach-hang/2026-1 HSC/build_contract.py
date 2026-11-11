"""Build the HSC construction contract from the TIDO 2024 draft (unpacked + merge_runs'd).
Usage: python build_contract.py <unpacked_dir> <appendix.xlsx> <summary.json>
Edits <unpacked_dir>/word/document.xml and header1.xml in place; rezip separately.
VT/NC per system come from the summary written by scripts/xlsx_reprice.py."""
import json, os, re, sys
import openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'scripts'))
from vn_number import dots, rnd, words

UNP, XLSX, SUMMARY = sys.argv[1:4]
CONTRACT_NO = '......../HĐTC/HSC-GMPc/2026'
APPENDIX_NO = '01/PLHD/HSC-GMPc/2026'
BLANK = '..............'

summ = json.load(open(SUMMARY, encoding='utf8'))
th = openpyxl.load_workbook(XLSX)['TH DU TOAN']
rows = []
for r in range(11, 20):
    sheet = re.search(r"'(.+)'!", th.cell(r, 6).value).group(1)
    vt, nc = (rnd(v) for v in summ['sheets'][sheet])
    rows.append((str(r - 10), th.cell(r, 2).value.strip(), th.cell(r, 4).value,
                 str(th.cell(r, 5).value), vt, nc))
# totals = sums of the rows shown; must equal the appendix grand total rounded to the đồng
VT = sum(r[4] for r in rows); NC = sum(r[5] for r in rows); TOTAL = VT + NC
assert TOTAL == rnd(summ['new_total']), (TOTAL, summ['new_total'])
# one VAT rate on the total (user, 2026-10-28); the table shows no VT/NC split
VAT_RATE = 0.08
VAT = rnd(TOTAL * VAT_RATE)
GRAND = TOTAL + VAT

p = f'{UNP}/word/document.xml'
x = open(p, encoding='utf8').read()

def sub(old, new, count=1):
    global x
    assert x.count(old) == count, (old, x.count(old))
    x = x.replace(old, new)

# ---- header, parties
sub('(Số 0606/HDTC/TIDO-GMPc/2024)', f'(Số {CONTRACT_NO})')
sub('ký kết vào ngày 06 tháng 06 năm 2024, tại văn phòng Công ty TNHH Sản xuất Xuất nhập khẩu Dược Mỹ phẩm TIDO giữa và bởi:',
    f'ký kết vào ngày ...... tháng ...... năm 2026, tại {BLANK} giữa và bởi:')
sub('<w:t>CÔNG TY TNHH SẢN XUẤT XUẤT NHẬP KHẨU DƯỢC MỸ PHẨM TIDO</w:t>',
    '<w:t>CÔNG TY CỔ PHẦN CÔNG NGHỆ TẾ BÀO GỐC HÒA LẠC</w:t>')
sub('Lô C29-C32 CCN Hải Sơn Đức Hòa Đông, xã Đức Hòa Đông, Huyện Đức Hòa, tỉnh Long An',
    'Nhà số 6, ngõ 38, ngách 26 Xuân La, Phường Xuân Đỉnh, TP Hà Nội')
sub('<w:t>0922438899</w:t>', f'<w:t>{BLANK}</w:t>')
sub('<w:t>Trần Minh Bình</w:t>', '<w:t>Phan Tú Quân</w:t>')
sub('<w:t>Giám đốc</w:t>', '<w:t>Giám đốc điều hành</w:t>')
sub('<w:t>1102055458</w:t>', '<w:t>0107997450</w:t>')
sub('Số nhà 18, đường Măng Cầm 1, Khu Đô Thị mới An Lạc Green Symphony, xã Vân Canh, huyện Hoài Đức, Thành phố Hà Nội, Việt Nam',
    '18 Măng Cầm 1, An Lạc Green Symphony, Sơn Đồng, Hà Nội')
sub('Số nhà 32, đường Vĩ Cầm 5, Khu Đô Thị mới An Lạc Green Symphony, xã Vân Canh, huyện Hoài Đức, Thành phố Hà Nội, Việt Nam',
    '32 Vĩ Cầm 5, An Lạc Green Symphony, Sơn Đồng, Hà Nội')
sub('<w:t>024.3994.5757</w:t>', '<w:t>(+84) 24 3787 2242</w:t>')

# ---- Điều 1
sub('>ệ thống thiết bị cơ điện xưởng sản xuất mỹ phẩm đạt tiêu chuẩn CGMP ASEAN </w:t>',
    '>ệ thống cơ điện (M&amp;E), phòng sạch và các hệ thống phụ trợ của Tổ hợp Công nghệ Tế bào gốc HSC đáp ứng tiêu chuẩn cGMP FDA </w:t>')
sub('Lô C29 - C32, CCN Hải Sơn Đức Hòa Đông, Xã Đức Hòa Đông, Huyện Đức Hòa, Tỉnh Long An.',
    'KCN Hòa Lạc, Thạch Thất, Hà Nội.')
sub('<w:t>01/PLHD/TIDO-GMPc/2024</w:t>', f'<w:t>{APPENDIX_NO}</w:t>')

# ---- Điều 2 price table: rebuild the item rows from the row of item 2
trs = list(re.finditer(r'<w:tr[ >].*?</w:tr>', x, re.S))
items = [m for m in trs if re.search(r'<w:t>(1|2|3|4)</w:t>', m.group()) and 'Hệ' in m.group()]
assert len(items) == 4
proto = items[1].group()
new_rows = []
for i, (no, name, unit, q, vt, nc) in enumerate(rows):
    t = proto
    for old, new in [('<w:t>2</w:t>', '<w:t>@@NO@@</w:t>'),   # placeholders first: no/qty can collide
                     ('<w:t>1</w:t>', '<w:t>@@QTY@@</w:t>'),
                     ('@@NO@@', no), ('@@QTY@@', q),
                     ('HỆ THỐNG ĐIỆN ĐỘNG LỰC', name.replace('&', '&amp;')),
                     ('<w:t>Hệ</w:t>', f'<w:t>{unit}</w:t>'),
                     ('1,007,069,300', dots(vt)), ('225,038,900', dots(nc)),
                     ('1,232,108,200', dots(vt + nc))]:
        assert t.count(old) == 1, old
        t = t.replace(old, new)
    t = re.sub(r'w14:paraId="([0-9A-F]{6})[0-9A-F]{2}"',
               lambda m: f'w14:paraId="{m.group(1)}{0xA0 + i:02X}"', t)
    new_rows.append(t)
x = x[:items[0].start()] + ''.join(new_rows) + x[items[3].end():]

sub('6,654,370,000', dots(TOTAL)); sub('645,308,000', dots(VAT)); sub('7,299,678,000', dots(GRAND))
sub('<w:t>THUẾ VAT VẬT TƯ 10%, NHÂN CÔNG VAT 8%</w:t>', f'<w:t>THUẾ VAT {round(VAT_RATE * 100)}%</w:t>')

# columns: STT | DIỄN GIẢI | ĐƠN VỊ | SỐ LƯỢNG | TỔNG THÀNH TIỀN (drop VẬT TƯ / NHÂN CÔNG)
W = [535, 4625, 794, 1121, 2910]                     # same 9985 total width as the template
i = x.find('DIỄN GIẢI'); ts = x.rfind('<w:tbl>', 0, i); te = x.find('</w:tbl>', i) + len('</w:tbl>')
tbl = x[ts:te]
tbl = re.sub(r'<w:tblGrid>.*?</w:tblGrid>',
             '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{w}"/>' for w in W) + '</w:tblGrid>', tbl, flags=re.S)
def width(tc, w):
    return re.sub(r'<w:tcW w:w="\d+"', f'<w:tcW w:w="{w}"', tc, count=1)
def rebuild(tr):
    tcs = list(re.finditer(r'<w:tc>.*?</w:tc>', tr, re.S))
    c = [m.group() for m in tcs]
    if 'VẬT TƯ' in tr and 'DIỄN GIẢI' not in tr:      # 2nd header row: VẬT TƯ / NHÂN CÔNG
        return ''
    if 'DIỄN GIẢI' in tr:                             # 1st header row: drop the spanning cell, no vMerge
        assert len(c) == 6
        c = [width(t.replace('<w:vMerge w:val="restart"/>', ''), w) for t, w in zip(c[:4] + c[5:], W)]
    elif len(c) == 7:                                 # item row
        c = [width(t, w) for t, w in zip(c[:4] + c[6:], W)]
    elif len(c) == 5:                                 # total rows: label spans the first 4 columns
        c = [width(c[0].replace('<w:gridSpan w:val="3"/>', '<w:gridSpan w:val="4"/>'), sum(W[:4])),
             width(c[4], W[4])]
    else:
        raise AssertionError(len(c))
    return tr[:tcs[0].start()] + ''.join(c) + tr[tcs[-1].end():]
tbl = re.sub(r'<w:tr[ >].*?</w:tr>', lambda m: rebuild(m.group()), tbl, flags=re.S)
x = x[:ts] + tbl + x[te:]
m = re.search(r'Bằng chữ\s*:\s*[^<]*', x)
x = x[:m.start()] + f'Bằng chữ: {words(GRAND)} đồng./.)' + x[m.end():]
m = re.search(r'Thuế suất thuế GTGT áp dụng theo nghị định số 94/2023[^<]*', x)
x = x[:m.start()] + 'Thuế suất thuế GTGT áp dụng theo quy định của pháp luật tại thời điểm xuất hóa đơn.' + x[m.end():]

# ---- Điều 3: 4 đợt sized to the work done at each milestone (see NOTES.md)
sub('Bên A sẽ tạm ứng tiếp cho Bên B 40% tổng giá trị Hợp Đồng',
    'Bên A sẽ tạm ứng tiếp cho Bên B 30% tổng giá trị Hợp Đồng')
sub('Vật tư, thiết bị chính bao gồm: Hệ thống điều hòa không khí (Không bao gồm AHU), Tủ điện, ống gió.',
    'Vật tư, thiết bị chính bao gồm: Chiller, tháp giải nhiệt, bơm nước lạnh và nước giải nhiệt, AHU, FCU, '
    'tủ điện, panel phòng sạch, thiết bị hệ thống nước RO, thiết bị hệ thống xử lý nước thải và thiết bị hệ thống BMS.')
sub('Bên A sẽ tạm ứng tiếp cho Bên B 20% tổng giá trị Hợp đồng',
    'Bên A sẽ tạm ứng tiếp cho Bên B 30% tổng giá trị Hợp đồng')
sub('lắp đặt hoàn thiện tất cả các hệ thống (Bao gồm cả lắp đặt AHU) và tiến hành chạy thử',
    'lắp đặt hoàn thiện tất cả các hệ thống theo Phụ Lục Hợp Đồng và tiến hành chạy thử')
sub('>Thanh toán phần còn lại của giá trị Hợp Đồng và',
    '>Thanh toán phần còn lại (10% tổng giá trị Hợp Đồng) và')

# ---- Điều 3, 4, 7, 12
sub('332.718.500 VNĐ', f'{dots(TOTAL * 0.05)} VNĐ')
sub('Thời hạn thi công là 65 ngày', 'Thời hạn thi công là ...... ngày')
sub('thiết bị cơ điện xưởng sản xuất mỹ phẩm đạt tiêu chuẩn CGMP ASEAN trong 01 tháng',
    'cơ điện (M&amp;E), phòng sạch và các hệ thống phụ trợ của Tổ hợp Công nghệ Tế bào gốc HSC trong 01 tháng')
sub('phục vụ cho việc xin giấy phép cGMP và vận hành', 'phục vụ cho việc đánh giá đáp ứng tiêu chuẩn cGMP FDA và vận hành')
sub('<w:t>CÔNG TY TNHH SẢN XUẤT XUẤT NHẬP KHẨU DƯỢC MỸ PHẨM TIDO.</w:t>',
    '<w:t>CÔNG TY CỔ PHẦN CÔNG NGHỆ TẾ BÀO GỐC HÒA LẠC.</w:t>')
sub('gồm 10 (mười) trang và 01 Phụ Lục Hợp Đồng (28 trang)',
    'gồm ...... (..........) trang và 01 Phụ Lục Hợp Đồng (...... trang)')

assert 'TIDO' not in x and 'Long An' not in x and 'Vân Canh' not in x, 'leftover old text'
open(p, 'w', encoding='utf8').write(x)

h = f'{UNP}/word/header1.xml'
hx = open(h, encoding='utf8').read()
# the header number is split over 10 runs: put it all in the first run, empty the others
ts = list(re.finditer(r'(<w:t[^>]*>)([^<]*)(</w:t>)', hx))
assert ''.join(m.group(2) for m in ts) == '0606/HDTC/TIDO-GMPc/2024'
for k, m in reversed(list(enumerate(ts))):
    hx = hx[:m.start()] + m.group(1) + (CONTRACT_NO if k == 0 else '') + m.group(3) + hx[m.end():]
open(h, 'w', encoding='utf8').write(hx)

print('VT', dots(VT), 'NC', dots(NC), 'TOTAL', dots(TOTAL))
print('VAT', dots(VAT), 'GRAND', dots(GRAND))
print(words(GRAND))
