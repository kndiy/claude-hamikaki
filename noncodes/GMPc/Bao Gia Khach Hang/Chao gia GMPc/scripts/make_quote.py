"""Soạn bản chào giá từ một phiếu yêu cầu (.toml).

    python scripts/make_quote.py "output/<YYMMDD Ten khach>/phieu_yeu_cau.toml"

Lấy file mẫu tương ứng trong "2025 Mau chao gia" làm nền (giữ header, logo, định dạng),
dựng lại phần thân theo phiếu, xuất .xlsx vào cùng thư mục với phiếu.
Ô thiếu thông tin tô vàng; sheet "Ghi chu soan thao" liệt kê việc cần rà trước khi gửi.
"""
import copy
import datetime as dt
import io
import json
import math
import re
import sys
import tomllib
from pathlib import Path

import openpyxl
import PIL  # noqa: F401  openpyxl silently drops images on save without Pillow
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = ROOT.parent / "2025 Mau chao gia"
CATALOG = json.loads((ROOT / "catalog" / "catalog.json").read_text(encoding="utf-8"))
SONG_NGU = json.loads((ROOT / "catalog" / "song_ngu.json").read_text(encoding="utf-8"))

QR_PATH = ROOT / "assets" / "gmpc_qr.jpg"  # copied from Desktop\GMPc-QR.jpg, 2026-10-13
SEAL_WIDTH_PX = 300  # seal + signature image, scaled from the template's 661 px
QR_PX = 120
PAGE_BREAK_WASTE = 1.1  # Excel never splits a row, so each page loses some space
SPILL_SHRINK = 0.35  # last-page spill (fraction of a page) small enough to shrink away
EMU = 9525  # EMU per pixel


def px(width):
    """Excel column width -> pixels (Calibri 11)."""
    return int((width or 8.43) * 7 + 5)


FONT = "Calibri Light"
YELLOW = PatternFill("solid", fgColor="FFFF00")
PLACEHOLDER = re.compile(r"…|\.{3,}")
BLOCKED = re.compile(r"ngoại giao|thúc đẩy hồ sơ|(chi phí|bao gồm)[^.]*tiếp đoàn", re.I)
ROMAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]
VAT_8_UNTIL = dt.date(2026, 12, 31)  # Decree 174/2025/NĐ-CP

# GMPc header, per the GMPc group CLAUDE.md (post-2025 administrative boundaries).
COMPANY = {
    "Công ty": ("Công ty cổ phần GMPc Việt Nam", "GMPc Vietnam JSC."),
    "Trụ sở": ("Trụ sở: 18 Măng Cầm 1, An Lạc Green Symphony, Sơn Đồng, Hà Nội",
               "HO: 18 Mang Cam 1 street, An Lạc Green Symphony, Son Dong Commune, Ha Noi"),
    "VPGD": ("VPGD: 32 Vĩ Cầm 5, An Lạc Green Symphony, Sơn Đồng, Hà Nội",
             "RO: 32 Vi Cam 5 street, An Lạc Green Symphony, Son Dong Commune, Ha Noi"),
    "VPĐD": ("VPĐD: Lầu 2, 156/1/1 Cộng Hòa, Phường Bảy Hiền, TP Hồ Chí Minh",
             "RO: 2nd floor, 156/1/1 Cong Hoa, Bay Hien Ward, HCMC"),
    "Tel": ("Tel: (+84) 24 3787 2242 / (+84) 28 3811 7383",) * 2,
    "Web": ("Web: gmp.com.vn / Email: contact@gmp.com.vn",) * 2,
}
LABELS = {  # key: (vn, en)
    "title": ("BẢN CHÀO GIÁ", "QUOTATION"),
    "info": ("Khách hàng & Thông tin báo giá", "Customer & Quote Details"),
    "ten": ("Khách hàng", "Customer"),
    "dia_chi": ("Địa chỉ", "Address"),
    "nguoi_lien_he": ("Người liên hệ", "Contact Person"),
    "dien_thoai": ("Số điện thoại", "Phone"),
    "email": ("Email liên hệ", "Email"),
    "ngay": ("Ngày báo giá", "Issued date"),
    "kinh_gui": ("Kính gửi", "Dear"),
    "stt": ("STT", "Item"),
    "pham_vi": ("Phạm vi công tác tư vấn", "Scope of Work"),
    "so_luong": ("Số lượng", "Quantity"),
    "chi_phi": ("Chi phí (VNĐ)", "Price (VND)"),
    "tong": ("Tổng cộng", "Total"),
    "vat": ("Thuế GTGT {r}%", "VAT ({r}%)"),
    "tong_vat": ("Tổng cộng bao gồm thuế GTGT {r}%", "Total with {r}% VAT"),
    "n_ghi_chu": ("* Ghi chú", "* Note"),
    "n_thoi_gian": ("* Thời gian\nthực hiện", "* Execution time"),
    "n_thanh_toan": ("* Tiến độ\nthanh toán", "* Payment terms"),
    "muc": ("Mục", "Part"),
}


def norm(s):
    return re.sub(r"\s+", " ", s).strip() if s else s


def renumber(items):
    """After lines are dropped/added: top level 1, 2, ...; sub level 1.1, 1.2 under each parent."""
    top, sub, out = 0, 0, []
    for it in items:
        c = it["code"] or ""
        if re.match(r"^\d+$", c):
            top, sub = top + 1, 0
            c = str(top)
        elif re.match(r"^\d+\.\d+$", c):
            sub += 1
            c = f"{top}.{sub}"
        out.append({**it, "code": c or None})
    return out


class Quote:
    def __init__(self, req_path):
        self.req_path = Path(req_path)
        self.req = tomllib.loads(self.req_path.read_text(encoding="utf-8"))
        self.review = []  # (where, issue)
        bg = self.req["bao_gia"]
        self.tid = bg["mau"]
        if self.tid not in CATALOG:
            sys.exit(f"Không có mẫu '{self.tid}'. Xem catalog/DANH_MUC.md")
        self.tpl = CATALOG[self.tid]
        self.lang = bg.get("ngon_ngu", "vn")
        assert self.lang in ("vn", "vn+en", "song_ngu"), "ngon_ngu phải là vn | vn+en | song_ngu"
        self.date = bg.get("ngay") or dt.date.today()
        self.stamp = bg.get("dong_dau", False)
        self.qr = bg.get("qr", True)
        self.vat = bg.get("vat")
        self.uses_en = self.lang != "vn"

    # ---------- text helpers ----------
    def en(self, vn, where):
        if vn is None or not self.uses_en:
            return None
        hit = SONG_NGU.get(norm(vn))
        if not hit:
            self.review.append((where, f"Chưa có bản EN cho: \"{vn[:70]}\""))
            return None
        if hit.get("trang_thai") != "da_duyet":
            self.flag_en = True
        return hit["en"]

    def text(self, vn, en, mode):
        """mode: 'vn' | 'en' | 'song_ngu'. Returns (text, needs_yellow)."""
        if mode == "vn":
            return vn, False
        if mode == "en":
            return (en, False) if en else (f"[EN?] {vn}", True)
        if not en:
            return f"{vn}\n[EN?]", True
        return f"{vn}\n{en}", False

    def pair(self, item, where):
        """Request entries are either a VN string or {vn=..., en=...}."""
        if isinstance(item, dict):
            return item.get("vn"), item.get("en") or self.en(item.get("vn"), where)
        return item, self.en(item, where)

    # ---------- content assembly ----------
    def sections(self):
        tpl_secs = {s["code"]: s for s in self.tpl["sections"]}
        out = []
        for k, m in enumerate(self.req.get("muc", [])):
            code = ROMAN[k]
            if "ma" in m:
                src = tpl_secs.get(m["ma"])
                if not src:
                    sys.exit(f"Mẫu {self.tid} không có mục '{m['ma']}' (có: {list(tpl_secs)})")
                if src.get("blocked") or BLOCKED.search(src["title"] or ""):
                    sys.exit(f"Mục {m['ma']} của mẫu ({src['title']}) có thể bị hiểu là chi phí cho cơ quan "
                             "cấp phép, nên không đưa vào báo giá. Nếu có phạm vi hợp lệ (theo dõi hồ sơ, hỗ trợ "
                             "làm việc với đoàn kiểm tra), hãy ghi thành mục tự do với câu chữ rõ ràng.")
                title = (src["title"], self.en(src["title"], f"Mục {code}"))
                items = []
                drop = [str(x) for x in m.get("bo", [])]
                edits = {str(c): v for c, v in m.get("sua", {}).items()}
                for it in src["items"]:
                    c = it["code"]
                    cs = None if c is None else (str(int(c)) if isinstance(c, float) and c == int(c) else str(c))
                    if cs and any(cs == d or cs.startswith(d + ".") for d in drop):
                        continue
                    if it.get("blocked") or BLOCKED.search(it["text"]):
                        self.review.append((f"Mục {code}", f"Đã loại dòng có rủi ro pháp lý: \"{it['text']}\". Nếu cần, thêm câu thay thế qua 'them'."))
                        continue
                    vn, en = self.pair(edits[cs], f"Mục {code}") if cs in edits else (it["text"], self.en(it["text"], f"Mục {code}"))
                    items.append({"code": cs, "vn": vn, "en": en})
            else:
                title = self.pair(m["tieu_de"], f"Mục {code}")
                items = []
            for t in [title[0]] + [a["vn"] if isinstance(a, dict) else a for a in m.get("them", [])]:
                if t and BLOCKED.search(t):
                    sys.exit(f"Câu \"{t}\" có thể bị hiểu là chi phí cho cơ quan cấp phép. Sửa lại câu chữ trong phiếu.")
            if "tieu_de" in m and "ma" in m:
                title = self.pair(m["tieu_de"], f"Mục {code}")
            for add in m.get("them", []):
                vn, en = self.pair(add, f"Mục {code}")
                new = {"code": str(add.get("ma", "")) if isinstance(add, dict) else "", "vn": vn, "en": en}
                after = str(add.get("sau", "")) if isinstance(add, dict) else ""
                pos = next((i + 1 for i, x in enumerate(items) if x["code"] == after), len(items))
                items.insert(pos, new)
            if m.get("danh_lai_so", True):
                items = renumber(items)
            out.append({
                "code": code,
                "src": m.get("ma"),
                "title": title,
                "items": items,
                "price": m.get("gia"),
                "time": self.pair(m["thoi_gian"], f"Mục {code}") if "thoi_gian" in m else None,
            })
        if not out:
            sys.exit("Phiếu chưa có [[muc]] nào.")
        return out

    def notes(self, secs):
        n, t = self.req.get("ghi_chu", {}), self.tpl["notes"]
        # VAT
        rate = self.vat if self.vat is not None else 8
        vat_vn = n.get("vat") or t.get("vat")
        vat_flag = False
        if self.date > VAT_8_UNTIL and self.vat is None:
            vat_flag = True
            self.review.append(("Thuế GTGT", f"Ngày báo giá {self.date:%d/%m/%Y} sau 31/12/2026: NĐ 174/2025 hết hiệu lực. Xác nhận thuế suất ('vat' trong phiếu) và câu ghi chú VAT."))
        if rate != 8 and not n.get("vat"):
            vat_flag = True
            self.review.append(("Thuế GTGT", f"Thuế suất {rate}% nhưng câu ghi chú vẫn là của mức 8%. Thêm ghi_chu.vat vào phiếu."))
        vat = self.pair(vat_vn, "Ghi chú VAT") if isinstance(vat_vn, str) else self.pair(vat_vn, "Ghi chú VAT")
        # Execution time: per-section values from the request, else the template text.
        timed = [s for s in secs if s["time"]]
        if timed:
            if len(timed) == 1:  # one timeline covers the whole quote
                time = timed[0]["time"]
            else:
                vn, en = [], []
                for s in timed:
                    tv, te = s["time"]
                    vn.append(f"- {LABELS['muc'][0]} {s['code']}: {tv}")
                    en.append(f"- {LABELS['muc'][1]} {s['code']}: {te}" if te else None)
                time = ("\n".join(vn), None if None in en else "\n".join(en))
        else:
            time = self.pair(t.get("time"), "Thời gian thực hiện")
            if len(secs) < len(self.tpl["sections"]) and t.get("time") and "Mục" in t["time"]:
                self.review.append(("Thời gian thực hiện", "Câu mẫu liệt kê theo Mục I-IV của mẫu đầy đủ nhưng báo giá chỉ chọn một phần. Cần ghi 'thoi_gian' cho từng mục."))
        # Payment
        if "thanh_toan" in self.req:
            dots = [self.pair(d, "Thanh toán") for d in self.req["thanh_toan"]["dot"]]
            pay = ("\n".join(d[0] for d in dots), None if any(d[1] is None for d in dots) else "\n".join(d[1] for d in dots))
        else:
            pay = self.pair(t.get("payment"), "Thanh toán")
        extra = [self.pair(x, "Ghi chú thêm") for x in n.get("them", [])]
        foot = []
        for f in (n["cuoi"] if "cuoi" in n else self.tpl["footnotes"]):
            vn = f["vn"] if isinstance(f, dict) else f
            if BLOCKED.search(vn):
                self.review.append(("Ghi chú cuối", f"Đã bỏ câu có rủi ro pháp lý: \"{vn}\". Thêm câu thay thế qua ghi_chu.cuoi nếu cần."))
                continue
            foot.append(self.pair(f, "Ghi chú cuối"))
        return {"rate": rate, "vat": vat, "vat_flag": vat_flag, "time": time, "pay": pay, "extra": extra, "foot": foot}


# ---------- worksheet rendering ----------
def find_layout(ws):
    """Locate prototype rows in a template sheet (for copying styles)."""
    rows = {}
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        a = str(a).strip() if a is not None else ""
        if "head" not in rows and re.match(r"^(STT|Item)", a):
            rows["head"] = r
            rows["intro"] = r - 1
        elif "head" in rows and "total" not in rows:
            if a in ROMAN and "section" not in rows:
                rows["section"] = r
            elif re.match(r"^\d+$", a) and "item" not in rows:
                rows["item"] = r
            elif re.match(r"^\d+\.\d+$", a) and "sub" not in rows:
                rows["sub"] = r
            elif re.match(r"^(Tổng|Total)", a):
                rows["total"] = r
        elif "total" in rows and a.startswith("*") and "note" not in rows:
            rows["note"] = r
        elif "note" in rows and a and not a.startswith("*") and "closing" not in rows:
            rows["closing"] = r
    rows.setdefault("section", rows["head"] + 1)
    rows.setdefault("item", rows["section"] + 1)
    rows.setdefault("sub", rows["item"])
    return rows


def snapshot(ws, r, cols=5):
    return [copy.copy(ws.cell(r, c)._style) for c in range(1, cols + 1)], ws.row_dimensions[r].height


# Row-height convention for the content block, rows 15 → "Tổng cộng bao gồm thuế" (user, 2026-10-13):
# 1 line = 30, 2 lines = 35, 3 lines = 50. Beyond 3 lines: +15 per line (assumption, pending confirmation).
CONTENT_FIRST_ROW = 15
ROW_HEIGHTS = {1: 30, 2: 35, 3: 50}
CHARS_PER_WIDTH = 1.15  # Calibri Light 11: characters per Excel column-width unit, calibrated on PDF output
CHARS_PER_WIDTH_BOLD = 1.0  # bold wraps earlier on screen (Haki row 23: 133 chars = 3 lines in Excel)


def rule_height(lines):
    return ROW_HEIGHTS.get(lines, ROW_HEIGHTS[3] + 15 * (lines - 3))


def wrapped_lines(text, chars):
    """Greedy word wrap, like Excel: count the lines a text takes in a cell `chars` wide."""
    total = 0
    for para in str(text).split("\n"):
        n, cur = 1, 0
        for word in para.split():
            need = len(word) if cur == 0 else cur + 1 + len(word)
            if need > chars and cur > 0:
                n, cur = n + 1, len(word)
            else:
                cur = need
        total += n
    return total


def est_height(text, chars_per_line, base=15.6):
    if not text:
        return base + 4
    lines = sum(max(1, math.ceil(len(part) / chars_per_line)) for part in str(text).split("\n"))
    return max(base + 4, lines * base + 6)


class Sheet:
    def __init__(self, q, ws, mode, protos, images):
        self.q, self.ws, self.mode, self.p, self.images = q, ws, mode, protos, images

    def t(self, vn, en):
        return self.q.text(vn, en, self.mode)

    def label(self, key, stacked=False, **kw):
        """stacked: bilingual labels in narrow column A go on two lines."""
        vn, en = (s.format(**kw) for s in LABELS[key])
        both = f"{vn.replace(chr(10), ' ')}\n{en}" if stacked else f"{vn} / {en}"
        return {"vn": vn, "en": en, "song_ngu": both}[self.mode]

    def put(self, r, c, value, style=None, yellow=False):
        cell = self.ws.cell(r, c)
        if style is not None:
            cell._style = copy.copy(style)
        cell.value = value
        cell.font = Font(**{**{k: getattr(cell.font, k) for k in ("sz", "b", "i", "u", "color")}, "name": FONT})
        if isinstance(value, str) and "\n" in value and not cell.alignment.wrap_text:
            al = copy.copy(cell.alignment)
            al.wrap_text = True
            cell.alignment = al
        if yellow:
            cell.fill = YELLOW
        return cell

    def apply_row_rule(self, first, last):
        """GMPc convention for the content block (row 15 → grand-total row): height by wrapped line count."""
        ws = self.ws
        spans = {}
        for m in ws.merged_cells.ranges:
            if m.min_row == m.max_row:
                spans[(m.min_row, m.min_col)] = m.max_col
        for r in range(first, last + 1):
            if ws.row_dimensions[r].hidden:
                continue
            lines = 1
            for c in range(1, 6):
                v = ws.cell(r, c).value
                if not isinstance(v, str) or v.startswith("="):
                    continue
                width = sum(ws.column_dimensions[get_column_letter(k)].width or 8.43
                            for k in range(c, spans.get((r, c), c) + 1))
                factor = CHARS_PER_WIDTH_BOLD if ws.cell(r, c).font.b else CHARS_PER_WIDTH
                lines = max(lines, wrapped_lines(v, width * factor))
            ws.row_dimensions[r].height = rule_height(lines)

    def render(self, secs, notes):
        ws, q, p = self.ws, self.q, self.p
        cust = q.req.get("khach_hang", {})
        # header block
        for r in range(1, 7):
            for c in range(1, 6):
                v = ws.cell(r, c).value
                if isinstance(v, str):
                    for k, (vn, en) in COMPANY.items():
                        if v.strip().startswith(k) or (k == "Công ty" and "GMPc" in v and r == 1):
                            ws.cell(r, c).value = en if self.mode != "vn" else vn
        ws.cell(7, 1).value = self.label("title")
        ws.cell(8, 1).value = self.label("info")
        fields = [("ten", 9), ("dia_chi", 10), ("nguoi_lien_he", 11), ("dien_thoai", 12), ("email", 13)]
        for key, r in fields:
            self.put(r, 1, self.label(key, stacked=True))
            vn = cust.get(key)
            en = cust.get(key + "_en") or vn
            val = {"vn": vn, "en": en, "song_ngu": vn if en == vn else f"{vn}\n{en}"}[self.mode] if vn else None
            self.put(r, 3, val, yellow=key in ("ten", "dia_chi") and not vn)
            if not vn and key not in ("ten", "dia_chi"):
                ws.row_dimensions[r].hidden = True
            elif self.mode == "song_ngu":
                ws.row_dimensions[r].height = 31
        self.put(14, 1, self.label("ngay", stacked=True))
        d = self.put(14, 3, q.date)
        d.number_format = "dd/mm/yyyy"
        d.alignment = Alignment(horizontal="left", vertical=d.alignment.vertical)
        self.put(16, 1, self.label("kinh_gui", stacked=True))
        kg = self.put(16, 3, ws.cell(9, 3).value)
        kg.alignment = Alignment(wrap_text=True, vertical="center")
        if self.mode == "song_ngu":
            ws.row_dimensions[14].height = ws.row_dimensions[16].height = 31

        # body
        r = p["rows"]["intro"]
        if "gioi_thieu" in q.req:  # customer-specific opening (current state / needs)
            intro = self.t(*q.pair(q.req["gioi_thieu"], "Câu mở đầu"))
        else:
            intro = self.t(q.tpl["intro"], q.en(q.tpl["intro"], "Câu mở đầu"))
        self.put(r, 1, intro[0], p["intro"][0][0], intro[1])
        ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=5)
        ws.row_dimensions[r].height = est_height(intro[0], 100)
        r += 1
        hs, hh = p["head"]
        for c, key in ((1, "stt"), (3, "pham_vi"), (4, "so_luong"), (5, "chi_phi")):
            self.put(r, c, self.label(key), hs[c - 1])
        self.put(r, 2, None, hs[1])
        ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=2)
        ws.row_dimensions[r].height = hh if self.mode != "song_ngu" else 31
        first = r + 1
        for s in secs:
            r += 1
            ss, _ = p["section"]
            title = self.t(*s["title"])
            for c in range(1, 6):
                self.put(r, c, None, ss[c - 1])
            self.put(r, 1, s["code"], ss[0])
            self.put(r, 3, title[0], ss[2], title[1])
            self.put(r, 4, 1, ss[3])
            price = self.put(r, 5, s["price"], ss[4], not s["price"])
            price.number_format = "#,##0"
            ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=2)
            ws.row_dimensions[r].height = est_height(title[0], 60)
            for it in s["items"]:
                r += 1
                st, _ = p["sub"] if it["code"] and "." in it["code"] else p["item"]
                txt = self.t(it["vn"], it["en"])
                for c in range(1, 6):
                    self.put(r, c, None, st[c - 1])
                code = it["code"]
                self.put(r, 1, int(code) if code and code.isdigit() else (code or None), st[0])
                self.put(r, 3, txt[0], st[2], txt[1])
                ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=2)
                ws.row_dimensions[r].height = est_height(txt[0], 66)
        last = r
        ts = p["total"]
        rate = notes["rate"]
        for k, (key, formula) in enumerate((
            ("tong", f"=SUM(E{first}:E{last})"),
            ("vat", f"=ROUND(E{last + 1}*{rate}%,0)"),
            ("tong_vat", f"=E{last + 1}+E{last + 2}"),
        )):
            r += 1
            st, h = ts[k]
            for c in range(1, 6):
                self.put(r, c, None, st[c - 1])
            self.put(r, 1, self.label(key, r=rate), st[0], key != "tong" and notes["vat_flag"])
            cell = self.put(r, 5, formula, st[4])
            cell.number_format = "#,##0"
            ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=3)
        self.apply_row_rule(CONTENT_FIRST_ROW, r)
        for vn, en in notes["foot"]:
            r += 1
            txt = self.t(vn, en)
            self.put(r, 1, txt[0], p["intro"][0][0], txt[1]).font = Font(name=FONT, i=True)
            ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="center")
            ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=5)
            ws.row_dimensions[r].height = est_height(txt[0], 100)
        ns, _ = p["note"]
        note = notes["vat"]
        if notes["extra"]:  # one "* Ghi chú" cell, each note as a "- " line
            dash = lambda s: s if s is None or s.lstrip().startswith("-") else f"- {s}"
            parts = [note] + list(notes["extra"])
            note = ("\n".join(dash(p[0]) for p in parts),
                    None if any(p[1] is None for p in parts) else "\n".join(dash(p[1]) for p in parts))
        blocks = [("n_ghi_chu", note, notes["vat_flag"])]
        blocks += [("n_thoi_gian", notes["time"], False), ("n_thanh_toan", notes["pay"], False)]
        for key, (vn, en), flag in blocks:
            r += 1
            txt = self.t(vn, en)
            yellow = flag or txt[1] or bool(vn and PLACEHOLDER.search(vn))
            self.put(r, 1, self.label(key, stacked=True), ns[0])
            self.put(r, 2, ":", ns[1])
            self.put(r, 3, txt[0], ns[2], yellow)
            ws.merge_cells(start_row=r, end_row=r, start_column=3, end_column=5)
            ws.row_dimensions[r].height = est_height(txt[0], 95)
            if yellow and vn and PLACEHOLDER.search(vn):
                self.q.review.append((self.label(key).replace("\n", " "), "Còn chỗ trống '…' cần điền."))
        # closing: thank-you paragraph, regards, company, title
        cs = p["closing"]
        r += 1
        closing = q.tpl["closing"]
        for k, vn in enumerate(closing):
            st, h = cs[min(k, len(cs) - 1)]
            txt = self.t(vn.strip(), q.en(vn.strip(), "Lời kết"))
            self.put(r, 1, txt[0], st[0], txt[1])
            ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=5)
            ws.row_dimensions[r].height = est_height(txt[0], 100) if (k == 0 or self.mode == "song_ngu") else h
            if k == 2:
                company_row = r
            r += 1
        # images: logo always, seal/signature only on request
        logo, seal = self.images
        if logo:
            ws.add_image(logo())
        if seal and q.stamp:  # scaled down and right-aligned under the title so the block stays on one page
            x = seal()
            w, h = SEAL_WIDTH_PX, round(SEAL_WIDTH_PX * x.height / x.width)
            right = sum(px(ws.column_dimensions[k].width) for k in "CDE")
            x.anchor = OneCellAnchor(
                _from=AnchorMarker(col=2, colOff=(right - w) * EMU, row=company_row + 1, rowOff=0),
                ext=XDRPositiveSize2D(w * EMU, h * EMU))
            ws.add_image(x)
        if q.qr and QR_PATH.exists():  # company QR, bottom-left of the signature block
            qr = Image(str(QR_PATH))
            qr.width = qr.height = QR_PX
            ws.add_image(qr, f"A{company_row}")
        image_pt = max(h * 0.75 if q.stamp and seal else 0, QR_PX * 0.75 if q.qr else 0)
        end = max(r, company_row + 1 + math.ceil(image_pt / 15))
        for k in range(r, end + 1):
            ws.row_dimensions[k].height = ws.row_dimensions[k].height or 15
        ws.print_area = f"A1:E{end}"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, self.fit_pages(end)
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    def fit_pages(self, end):
        """If only a sliver spills onto the last page (e.g. half the seal), shrink to one page fewer."""
        ws, pm = self.ws, self.ws.page_margins
        landscape = ws.page_setup.orientation == "landscape"
        page_w, page_h = (842, 595) if landscape else (595, 842)  # A4, points
        usable_w = page_w - (pm.left + pm.right) * 72
        usable_h = page_h - (pm.top + pm.bottom) * 72
        content_w = sum(px(ws.column_dimensions[c].width) for c in "ABCDE") * 0.75
        scale = min(1, usable_w / content_w)
        content_h = sum(ws.row_dimensions[k].height or 15.6 for k in range(1, end + 1)
                        if not ws.row_dimensions[k].hidden)
        pages = content_h * scale / usable_h * PAGE_BREAK_WASTE
        whole = math.floor(pages)
        return whole if whole >= 1 and pages - whole < SPILL_SHRINK else 0


def build(req_path):
    q = Quote(req_path)
    q.flag_en = False
    secs = q.sections()
    notes = q.notes(secs)
    wb = openpyxl.load_workbook(TEMPLATE_DIR / q.tpl["source"])
    base = wb[q.tpl["sheet_vn"]]
    rows = find_layout(base)
    protos = {
        "rows": rows,
        "intro": snapshot(base, rows["intro"]),
        "head": snapshot(base, rows["head"]),
        "section": snapshot(base, rows["section"]),
        "item": snapshot(base, rows["item"]),
        "sub": snapshot(base, rows["sub"]),
        "total": [snapshot(base, rows["total"] + k) for k in range(3)],
        "note": snapshot(base, rows["note"]),
        "closing": [snapshot(base, rows["closing"] + k) for k in range(4)],
    }
    # Keep image bytes, then rebuild each image fresh (one Image object per sheet).
    imgs = sorted(base._images, key=lambda im: im.anchor._from.row)
    def factory(im):
        # Keep the template's two-cell anchor so the image keeps its on-page size.
        data, anchor = im._data(), copy.deepcopy(im.anchor)
        def make(shift=0):
            x = Image(io.BytesIO(data))
            x.anchor = copy.deepcopy(anchor)
            x.anchor._from.row += shift
            if getattr(x.anchor, "to", None) is not None:
                x.anchor.to.row += shift
            return x
        return make
    images = (factory(imgs[0]) if imgs else None, factory(imgs[1]) if len(imgs) > 1 else None)

    def reset(ws):
        ws._images = []
        for m in list(ws.merged_cells.ranges):
            if m.max_row >= rows["intro"] or m.min_col >= 9:  # body, and side notes in I:P
                ws.unmerge_cells(str(m))
        ws.delete_rows(rows["intro"], ws.max_row)
        for c in ("I", "J", "K", "L", "M", "N", "O", "P"):
            for rr in range(1, 20):
                ws[f"{c}{rr}"].value = None

    for ws in [w for w in wb.worksheets if w is not base]:
        wb.remove(ws)
    reset(base)
    sheets = []
    if q.lang == "vn+en":
        en_ws = wb.copy_worksheet(base)
        en_ws.title = "ENG"
        sheets = [(base, "vn"), (en_ws, "en")]
    else:
        sheets = [(base, "vn" if q.lang == "vn" else "song_ngu")]
    for ws, mode in sheets:
        Sheet(q, ws, mode, protos, images).render(secs, notes)

    # review sheet
    if q.uses_en and q.flag_en:
        q.review.append(("Bản EN", "Thư viện song ngữ còn câu 'cho_duyet' (chưa duyệt). Đọc lại phần tiếng Anh trước khi gửi."))
    if q.stamp:
        q.review.append(("Con dấu", "File có dấu + chữ ký TGĐ (dấu còn in 'H. Hoài Đức'). Chỉ gửi khi giá và điều khoản đã chốt, nên gửi PDF."))
    if any(not s["price"] for s in secs):
        q.review.append(("Chi phí", "Có mục chưa có giá (ô vàng)."))
    if q.tpl["sector"] == "du_an_dau_tu":
        q.review.append(("Đính kèm", "Mẫu này nhắc tới mục lục và bảng tính định mức đính kèm. Chưa tự kèm; bảng tính mẫu dùng VAT 10% và có Sheet3 rác, cần làm sạch trước khi gửi."))
    rv = wb.create_sheet("Ghi chu soan thao")
    rv.append(["Mẫu", q.tpl["source"]])
    rv.append(["Phiếu", str(q.req_path.name)])
    rv.append([])
    rv.append(["Chỗ", "Cần rà trước khi gửi (xóa sheet này và bỏ tô vàng khi xong)"])
    for where, issue in dict.fromkeys(q.review):
        rv.append([where, issue])
    rv.column_dimensions["A"].width, rv.column_dimensions["B"].width = 24, 120
    for row in rv.iter_rows():
        for c in row:
            c.font = Font(name=FONT, b=c.row == 4)
            c.alignment = Alignment(wrap_text=True, vertical="top")

    folder = q.req_path.parent
    name = q.req["bao_gia"].get("ten_file") or f"{folder.name} - draft"
    out = folder / f"{name}.xlsx"
    v = 2
    while out.exists():
        out = folder / f"{name} v{v}.xlsx"
        v += 1
    wb.save(out)
    print(f"-> {out}")
    for where, issue in dict.fromkeys(q.review):
        print(f"  [{where}] {issue}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    build(sys.argv[1])
