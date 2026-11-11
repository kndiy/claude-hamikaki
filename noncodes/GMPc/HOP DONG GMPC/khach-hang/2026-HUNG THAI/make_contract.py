"""Fill the HS GMP training contract for Hung Thai from the ECO CNC CGMP template.

Usage: python make_contract.py <unpacked_dir>   (unpacked + merge_runs.py already applied)
Edits word/document.xml and word/header1.xml in place; every edit must hit exactly once.
"""
import re
import sys

D = sys.argv[1]
T_RE = re.compile(r'(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)')


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def text_of(frag):
    return ''.join(m.group(2) for m in T_RE.finditer(frag))


def span_replace(frag, old, new):
    """Replace `old` (may cross runs) inside frag; new text goes into the first run touched."""
    ts = list(T_RE.finditer(frag))
    full = ''.join(m.group(2) for m in ts)
    i = full.find(old)
    assert i >= 0 and full.count(old) == 1, (old, full)
    j = i + len(old)
    out, pos, last = [], 0, 0
    inserted = False
    for m in ts:
        s, e = pos, pos + len(m.group(2))
        pos = e
        txt = m.group(2)
        if e <= i or s >= j:
            continue
        a, b = max(i, s) - s, min(j, e) - s
        nt = txt[:a] + ('' if inserted else esc(new)) + txt[b:]
        inserted = True
        out.append(frag[last:m.start()])
        out.append('<w:t xml:space="preserve">' + nt + m.group(3))
        last = m.end()
    out.append(frag[last:])
    return ''.join(out)


def set_cell_text(tc, new):
    tc = tc.replace('<w:br/>', '')
    return span_replace(tc, text_of(tc), new)


def para_edit(x, key, old, new):
    paras = [m for m in re.finditer(r'<w:p[ >].*?</w:p>', x, re.S) if key in text_of(m.group(0))]
    assert len(paras) == 1, (key, len(paras))
    m = paras[0]
    return x[:m.start()] + span_replace(m.group(0), old, new) + x[m.end():]


def row_set(row, num, desc):
    row = re.sub(r'<w:bookmark(?:Start|End) [^>]*/>', '', row)  # cloned rows would duplicate ids
    cells = list(re.finditer(r'<w:tc>.*?</w:tc>', row, re.S))
    assert len(cells) == 2, len(cells)
    c0, c1 = cells
    return (row[:c0.start()] + set_cell_text(c0.group(0), num) + row[c0.end():c1.start()]
            + set_cell_text(c1.group(0), desc) + row[c1.end():])


A_NAME = 'Công ty cổ phần tập đoàn dược Hưng Thái'
NO = '3010/HDTV/HungThai-GMPc/2026'  # number = signing date DDMM

x = open(D + '/word/document.xml', encoding='utf-8').read()

E = [  # (paragraph key, old, new)
    ('(Số 1507', '1507/HDTV/ECOCNC-GMPc/2026', NO),
    ('có nhu cầu cần cung cấp', 'Công ty cổ phần dược phẩm ECO CNC', A_NAME),
    ('có nhu cầu cần cung cấp', 'CGMP - ASEAN', 'HS GMP'),
    ('khả năng cung cấp dịch vụ', 'Công ty cổ phần dược phẩm ECO CNC', A_NAME),
    ('được ký kết vào ngày', 'ngày 15 tháng 07 năm 2026', 'ngày 30 tháng 10 năm 2026'),
    ('được ký kết vào ngày', 'Công ty cổ phần dược phẩm ECO CNC', A_NAME),
    ('BÊN A: CÔNG TY', 'CÔNG TY CỔ PHẦN DƯỢC PHẨM ECO CNC', 'CÔNG TY CỔ PHẦN TẬP ĐOÀN DƯỢC HƯNG THÁI'),
    ('0109067309', '0109067309', '0901241554'),
    ('Thôn 1, Xã Tây', 'Thôn 1, Xã Tây Phương, TP Hà Nội, Việt Nam',
     'Thôn Thọ Bình, Xã Triệu Việt Vương, Tỉnh Hưng Yên, Việt Nam'),
    ('Vũ Thị Nga', 'Vũ Thị Nga', 'Đỗ Thị Cúc'),
    ('Lưu ý:', 'nhà máy bố trí nhân sự tham gia và phòng đào tạo có bàn ghế, máy tính để thuận tiện trao đổi thông tin.',
     'Bên A chuẩn bị phòng đào tạo có máy chiếu, bố trí nhân sự tham gia và trao đổi, thảo luận các nội dung liên quan trong khóa đào tạo.'),
    ('Chứng chỉ đào tạo GMP.', 'Chứng chỉ đào tạo GMP.', 'Chứng chỉ đào tạo HS GMP.'),
    ('Phí dịch vụ trước thuế', '20.000.000', '24.000.000'),
    ('Thuế GTGT 8%', '1.600.000', '1.920.000'),
    ('Phí dịch vụ bao gồm 8%', '21.600.000', '25.920.000'),
    ('Bằng chữ', 'Hai mươi mốt triệu sáu trăm nghìn', 'Hai mươi lăm triệu chín trăm hai mươi nghìn'),
    ('Lần 1: Tạm ứng', '70 %', '60%'),
    ('Lần 2: Thanh toán', '30%', '40%'),
    ('lịch làm việc của cơ sở', 'của cơ sở', 'của nhà máy'),
]
for k, o, n in E:
    x = para_edit(x, k, o, n)

# No surcharge for extra trainees (user, 2026-10-30): drop that paragraph
sur = [m for m in re.finditer(r'<w:p[ >].*?</w:p>', x, re.S) if 'phụ thu' in text_of(m.group(0))]
assert len(sur) == 1
x = x[:sur[0].start()] + x[sur[0].end():]

# Advance invoice only after the money arrives (user, 2026-10-30): avoids cancelling
# invoices when a customer signs and then backs out
x = para_edit(x, 'Lần 1: Tạm ứng', 'hồ sơ tạm ứng của bên B bao gồm:', 'Đề nghị tạm ứng của Bên B.')
adv = [m for m in re.finditer(r'<w:p[ >].*?</w:p>', x, re.S) if text_of(m.group(0)).strip() == 'Đề nghị tạm ứng']
assert len(adv) == 1
x = x[:adv[0].start()] + x[adv[0].end():]
x = para_edit(x, 'giá trị tạm ứng', text_of(
    [m.group(0) for m in re.finditer(r'<w:p[ >].*?</w:p>', x, re.S) if 'giá trị tạm ứng' in text_of(m.group(0))][0]),
    'Hóa đơn GTGT tương ứng giá trị tạm ứng: Bên B xuất và gửi cho Bên A ngay sau khi nhận được khoản tạm ứng.')

# Bên A "Chức vụ": the first "Giám đốc" cell (Bên A table comes before Bên B's)
t0 = x.index('<w:tbl>')
t0e = x.index('</w:tbl>', t0)
x = x[:t0] + span_replace(x[t0:t0e], 'Giám đốc', 'Tổng giám đốc') + x[t0e:]

# Scope table (3rd table): rebuild rows from the quote dated 29/10/2026
tbls = [m.start() for m in re.finditer('<w:tbl>', x)]
s = tbls[2]
e = x.index('</w:tbl>', s)
tbl = x[s:e]
rows = list(re.finditer(r'<w:tr[ >].*?</w:tr>', tbl, re.S))
assert len(rows) == 13
R = [r.group(0) for r in rows]
head, sec, item, note = R[0], R[1], R[5], R[12]
new_rows = [head] + [row_set(sec if n.isalpha() else item, n, d) for n, d in [
    ('I', 'Đào tạo HS GMP cơ bản'),
    ('1', 'Giới thiệu chung về HS GMP cho TPBVSK'),
    ('2', 'Cơ sở vật chất và trang thiết bị'),
    ('3', 'Nguyên vật liệu'),
    ('4', 'Nhân sự và đào tạo'),
    ('II', 'Đào tạo HS GMP nâng cao'),
    ('1', 'Hồ sơ và tài liệu'),
    ('2', 'Thẩm định và đánh giá'),
    ('III', 'Trao đổi, thảo luận các vấn đề liên quan'),
    ('IV', 'Phương thức đào tạo'),
    ('1', 'Thời lượng: 2 buổi (4 giờ/buổi) đào tạo trực tiếp'),
    ('2', 'Thời gian thực hiện dự kiến: do 2 bên sắp xếp thuận lợi cho lịch làm việc của nhà máy'),
    ('3', 'Địa điểm đào tạo: trực tiếp tại nhà máy của ' + A_NAME),
    ('4', 'Ngôn ngữ đào tạo: Tiếng Việt'),
    ('5', 'Số nhân sự tham gia: 20 - 30 người'),
    ('6', 'Số nhân sự được cấp chứng nhận: 20 - 30 người'),
]] + [note]
tbl = tbl[:rows[0].start()] + ''.join(new_rows) + tbl[rows[-1].end():]
x = x[:s] + tbl + x[e:]

assert 'ECO' not in text_of(x) and 'CGMP' not in text_of(x), 'template leftovers'
open(D + '/word/document.xml', 'w', encoding='utf-8').write(x)

h = open(D + '/word/header1.xml', encoding='utf-8').read()
h = span_replace(h, '1507/HDTV/ECOCNC-GMPc/2026', NO)
open(D + '/word/header1.xml', 'w', encoding='utf-8').write(h)
print('ok')
