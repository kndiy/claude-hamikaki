"""Vietnamese number helpers: dots() -> '1.234.567', words() -> 'Một triệu hai trăm ...'.
Rounding is half-up like Excel ROUND (Python round() is half-even)."""
import math

def rnd(n):
    return int(math.floor(n + 0.5))

def dots(n):
    return f'{rnd(n):,}'.replace(',', '.')

DIGITS = 'không một hai ba bốn năm sáu bảy tám chín'.split()

def _read3(n, full):
    h, t, u = n // 100, n // 10 % 10, n % 10
    out = []
    if full or h:
        out += [DIGITS[h], 'trăm']
    if t == 0 and u and (full or h):
        out.append('linh')
    elif t == 1:
        out.append('mười')
    elif t > 1:
        out += [DIGITS[t], 'mươi']
    if u:
        out.append('mốt' if u == 1 and t > 1 else 'lăm' if u == 5 and t else DIGITS[u])
    return out

def words(n):
    n = rnd(n)
    units = ['', 'nghìn', 'triệu', 'tỷ']
    groups = []
    while n:
        groups.append(n % 1000); n //= 1000
    out = []
    for i in range(len(groups) - 1, -1, -1):
        if groups[i]:
            out += _read3(groups[i], i < len(groups) - 1) + ([units[i]] if units[i] else [])
    s = ' '.join(out)
    return s[0].upper() + s[1:]
