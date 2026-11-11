"""Bước 2: Đối chiếu sheet 'Tổng hợp' (Excel gốc) với dữ liệu đọc từ HĐ scan và xuất báo cáo Excel.

Input:
    data/raw/260625 Danh sách KH đã kí HĐ 2011-2026.xlsx   (sheet 'Tổng hợp')
    data/processed/scan_extract.jsonl                      (1 dòng / file PDF scan, đọc thủ công từ ảnh)
    data/processed/scan_index.csv                          (danh mục file scan)
Output:
    output/<YYMMDD> Thong ke KH da ky HD 2011-2026.xlsx

Nguyên tắc:
- Không sửa file gốc. Thiếu thông tin thì để trống, không suy đoán.
- Excel có giá trị -> giữ nguyên; Excel trống mà scan có -> bổ sung (tô xanh);
  hai nguồn khác nhau -> giữ Excel, tô vàng và ghi rõ giá trị trong scan.
"""
import csv
import datetime as dt
import difflib
import json
import os
import re
import unicodedata
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

PROJECT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(PROJECT, "data", "raw", "260625 Danh sách KH đã kí HĐ 2011-2026.xlsx")
SCAN_JSONL = os.path.join(PROJECT, "data", "processed", "scan_extract.jsonl")
SCAN_INDEX = os.path.join(PROJECT, "data", "processed", "scan_index.csv")
SCAN_ROOT = r"L:\04. Tổng hợp Data khách hàng\01. Danh sách KH đã kí HĐ\Scan hợp đồng tư vấn"
OUT_DIR = os.path.join(PROJECT, "output")

# Scan không tính là HĐ bán của GMPc / bản trùng
EXCLUDE = {
    "S047": "GMPc là bên THUÊ dịch vụ (5SPRO cung cấp) - không phải KH",
    "S107": "Bản scan trùng số HĐ với S106 (1010/HDTV/Vincos-GMPc/2025)",
    "S002": "HĐ giữa Dabaco (Bên A) và Anh Khang (Bên B), GMPc không ký - liên quan dòng Excel 112 "
            "(HĐ 2601/HDTV/AK-GMPc/2022 giữa Anh Khang và GMPc, 355tr) nhưng chưa có scan HĐ đó",
}
# Khớp thủ công (đã kiểm tra): số HĐ / pháp nhân ghi khác nhau giữa Excel và scan
MANUAL_MATCH = {
    "S046": (103, "Thủ công: Excel ghi số 20230911/HDTV/MEDTEC, scan ghi 1109/HDTV/MEDTEC"),
    "S054": (107, "Thủ công: cùng HĐ nguyên tắc 2512/HDNT/VINMEC; Excel ghi 'BV ĐK quốc tế Vinmec', scan ghi 'Viện NC Tế bào gốc và Công nghệ Gen Vinmec'"),
}

LINH_VUC_MAP = {
    "dược phẩm": "Dược phẩm", "who gmp": "Dược phẩm", "dược phẩm who gmp": "Dược phẩm",
    "đông dược who gmp": "Dược phẩm", "đông dược": "Dược phẩm", "thuốc phóng xạ": "Dược phẩm",
    "dược phẩm/ tp bvsk": "Dược phẩm",
    "mỹ phẩm": "Mỹ phẩm", "mỹ phẩm cgmp": "Mỹ phẩm", "mỹ phẩm - thiết bị y tế": "Mỹ phẩm",
    "tpbvsk": "TPBVSK", "tp bvsk": "TPBVSK", "hs gmp": "TPBVSK", "hs - gmp": "TPBVSK",
    "thực phẩm bvsk - hs gmp": "TPBVSK",
    "thú y thuỷ sản": "Thú y", "thú y thủy sản": "Thú y", "thú y": "Thú y", "thuốc thú y": "Thú y",
    "kho gsp": "Kho GSP/GDP", "kho gsp/ gdp": "Kho GSP/GDP", "kho gdp": "Kho GSP/GDP", "gdp": "Kho GSP/GDP",
    "kho gdp, điểm tiêm vaccine": "Kho GSP/GDP",
    "bao bì dược phẩm": "Bao bì dược phẩm", "bao bì dp": "Bao bì dược phẩm",
    "thiết bị y tế": "Thiết bị y tế", "thiết bị y tế - iso": "Thiết bị y tế",
    "vacxin": "Sinh phẩm/Vaccine/Tế bào", "tế bào gốc": "Sinh phẩm/Vaccine/Tế bào",
    "cgmp": "Sinh phẩm/Vaccine/Tế bào",
    "lập báo cáo nghiên cứu khả thi": "Khác", "dv công bố sp": "Khác", "phòng lab": "Khác",
    "sữa chua": "Khác",
}
TINH_TRANG_MAP = {
    "đã hoàn thành": "Đã hoàn thành", "đã thanh lý": "Đã thanh lý",
    "đang thực hiện": "Đang thực hiện", "đang triển khai": "Đang thực hiện",
}


# ---------------------------------------------------------------- helpers
def strip_accents(s):
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def clean(v):
    if v is None:
        return ""
    if isinstance(v, str):
        return re.sub(r"\s+", " ", v).strip()
    return v


def norm_contract(s):
    s = strip_accents(str(s or "")).upper()
    s = re.sub(r"^\s*SO\s*", "", s)
    return re.sub(r"[^A-Z0-9]", "", s)


COMPANY_STOP = ("cong ty", "co phan", "tnhh", "mtv", "mot thanh vien", "tap doan", "chi nhanh",
                "viet nam", "vietnam", "cp", "jsc", "ltd", "group", "san xuat", "thuong mai",
                "dau tu", "quoc te", "duoc pham", "duoc my pham", "my pham", "duoc", "phat trien")


def norm_company(s):
    s = strip_accents(str(s or "")).lower()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    for w in COMPANY_STOP:
        s = re.sub(rf"\b{w}\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def company_sim(a, b):
    a, b = norm_company(a), norm_company(b)
    if not a or not b:
        return 0.0
    if a in b or b in a:
        return 1.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def parse_excel_date(v):
    """Trả về (tháng, năm) từ ô 'Thời điểm kí' (text MM/YYYY hoặc số serial Excel)."""
    if v in (None, ""):
        return None, None
    if isinstance(v, (int, float)):
        d = dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
        return d.month, d.year
    if isinstance(v, dt.datetime):
        return v.month, v.year
    m = re.search(r"(\d{1,2})\s*/\s*(\d{4})", str(v))
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r"(\d{4})", str(v))
    return (None, int(m.group(1))) if m else (None, None)


def parse_scan_date(s):
    m = re.match(r"(\d{2}|__)/(\d{2}|__)/(\d{4})", s or "")
    if not m:
        return None, None, None
    d = None if m.group(1) == "__" else int(m.group(1))
    mo = None if m.group(2) == "__" else int(m.group(2))
    return d, mo, int(m.group(3))


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    return None


# ---------------------------------------------------------------- load
def load_excel_rows():
    wb = openpyxl.load_workbook(RAW, data_only=True, read_only=True)
    ws = wb["Tổng hợp"]
    rows, year = [], None
    for i, r in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=12, values_only=True), start=2):
        r = [clean(x) for x in r]
        stt, cty, tdk, sohd, nd, gt, tt, lv, dm, sdt, email, dmg = r
        if isinstance(cty, (int, float)) and 2000 < cty < 2100 and not sohd and not nd:
            year = int(cty)          # dòng tiêu đề năm (cột F có thể là tổng SUM của năm -> bỏ)
            continue
        if not any([stt, cty, tdk, sohd, nd, gt, tt, lv, dm, sdt, email, dmg]):
            continue
        month, y2 = parse_excel_date(tdk)
        rows.append({
            "excel_row": i, "block_year": year, "stt": stt, "cong_ty": cty or "", "thoi_diem": tdk,
            "thang": month, "nam_ky": y2, "so_hd": str(sohd or ""), "noi_dung": nd or "",
            "gia_tri": num(gt), "tinh_trang": tt or "", "linh_vuc": lv or "", "dau_moi": dm or "",
            "sdt": str(sdt or ""), "email": email or "", "dau_moi_gmpc": dmg or "",
        })
    return rows


def load_scans():
    idx = {r["id"]: r for r in csv.DictReader(open(SCAN_INDEX, encoding="utf-8-sig"))}
    scans = []
    for line in open(SCAN_JSONL, encoding="utf-8"):
        s = json.loads(line)
        s.update({"folder": idx[s["id"]]["folder"], "file": idx[s["id"]]["file"]})
        d, mo, y = parse_scan_date(s["ngay_ky"])
        s.update({"ngay": d, "thang": mo, "nam": y})
        loai = s.get("loai", "")
        s["ben_ky"] = "OKS" if "OKS" in loai else ("Liên danh GMPc - INOTEK" if "liên danh" in loai.lower() else "GMPc")
        scans.append(s)
    return scans


# ---------------------------------------------------------------- match
def match(excel_rows, scans):
    by_no = defaultdict(list)
    for r in excel_rows:
        if r["so_hd"]:
            by_no[norm_contract(r["so_hd"])].append(r)
    used, result = set(), {}
    row_by_no = {r["excel_row"]: r for r in excel_rows}
    for sid, (row, how) in MANUAL_MATCH.items():
        result[sid] = (row_by_no[row], how)
        used.add(row)
    for s in scans:
        if s["id"] in EXCLUDE or s["id"] in result:
            continue
        key = norm_contract(s["so_hd"])
        cands = [r for r in by_no.get(key, []) if r["excel_row"] not in used] if key else []
        how = "Số HĐ"
        if not cands:
            # fallback: cùng năm, tháng lệch <=1, tên công ty giống
            best, score = None, 0
            for r in excel_rows:
                if r["excel_row"] in used:
                    continue
                y = r["nam_ky"] or r["block_year"]
                if s["nam"] and y and abs(y - s["nam"]) > 0:
                    continue
                if s["thang"] and r["thang"] and abs(r["thang"] - s["thang"]) > 1:
                    continue
                sc = company_sim(s["cong_ty"], r["cong_ty"])
                lead_s = re.match(r"\d{2,4}", s["so_hd"] or "")
                lead_r = re.match(r"\s*(\d{2,4})", r["so_hd"] or "")
                if lead_s and lead_r and lead_s.group(0) == lead_r.group(1):
                    sc += 0.3
                if sc > score:
                    best, score = r, sc
            if best and score >= 0.75:
                cands, how = [best], f"Tên công ty + thời gian (điểm {score:.2f})"
        if cands:
            r = cands[0]
            used.add(r["excel_row"])
            result[s["id"]] = (r, how)
    return result


# ---------------------------------------------------------------- build master
COLS = [
    ("STT", 6), ("Năm", 7), ("Tháng ký", 9), ("Ngày ký (theo HĐ)", 12), ("Số hợp đồng", 30),
    ("Loại HĐ", 16), ("Khách hàng", 40), ("MST", 14), ("Lĩnh vực (gốc)", 16), ("Lĩnh vực (chuẩn hoá)", 18),
    ("Nội dung", 60), ("Giá trị trước VAT (VNĐ)", 18), ("Tổng gồm VAT (VNĐ)", 18), ("Tình trạng (gốc)", 15),
    ("Tình trạng (chuẩn hoá)", 15), ("Đại diện KH (theo HĐ)", 30), ("Đầu mối liên hệ", 18), ("SĐT", 15),
    ("Email", 24), ("Đầu mối GMPc", 14), ("Bên ký", 14), ("File scan", 40), ("Nguồn dữ liệu", 18),
    ("Đối chiếu", 20), ("Tính vào thống kê", 12), ("Ghi chú / điểm cần kiểm tra", 60), ("Dòng Excel gốc", 10),
]
C = {name: i for i, (name, _) in enumerate(COLS)}

FILL_ADD = PatternFill("solid", fgColor="E2EFDA")    # bổ sung từ scan
FILL_DIFF = PatternFill("solid", fgColor="FFF2CC")   # lệch giữa 2 nguồn
FILL_NEW = PatternFill("solid", fgColor="DDEBF7")    # dòng chỉ có trong scan
FILL_OUT = PatternFill("solid", fgColor="EDEDED")    # không tính thống kê
FILL_HEAD = PatternFill("solid", fgColor="1F4E78")


def loai_from_so(so):
    s = strip_accents(so).upper()
    for k, v in (("HDTVGS", "HĐ tư vấn giám sát"), ("HDTVTK", "HĐ tư vấn thiết kế"), ("HDTK", "HĐ tư vấn thiết kế"),
                 ("HDTC", "HĐ thi công"), ("HDTP", "HĐ thầu phụ"), ("HDDV", "HĐ dịch vụ"), ("HDTV", "HĐ tư vấn"),
                 ("HDGS", "HĐ tư vấn giám sát")):
        if k in s.replace("/", "").replace(" ", "") or k in s:
            return v
    return ""


def build_master(excel_rows, scans, matched):
    scan_by_row = {r["excel_row"]: (sid, how) for sid, (r, how) in matched.items()}
    scan_by_id = {s["id"]: s for s in scans}
    out = []  # list of (values, cell_fills dict, row_fill)

    def base_excel(r):
        v = [""] * len(COLS)
        v[C["Năm"]] = r["block_year"] or r["nam_ky"] or ""
        v[C["Tháng ký"]] = f"{r['thang']:02d}/{r['nam_ky']}" if r["thang"] and r["nam_ky"] else (r["thoi_diem"] if isinstance(r["thoi_diem"], str) else "")
        v[C["Số hợp đồng"]] = r["so_hd"]
        v[C["Loại HĐ"]] = loai_from_so(r["so_hd"])
        v[C["Khách hàng"]] = r["cong_ty"]
        v[C["Lĩnh vực (gốc)"]] = r["linh_vuc"]
        v[C["Lĩnh vực (chuẩn hoá)"]] = LINH_VUC_MAP.get(r["linh_vuc"].lower(), "")
        v[C["Nội dung"]] = r["noi_dung"]
        v[C["Giá trị trước VAT (VNĐ)"]] = round(r["gia_tri"]) if r["gia_tri"] else ""
        v[C["Tình trạng (gốc)"]] = r["tinh_trang"]
        v[C["Tình trạng (chuẩn hoá)"]] = TINH_TRANG_MAP.get(r["tinh_trang"].lower(), "")
        v[C["Đầu mối liên hệ"]] = r["dau_moi"]
        v[C["SĐT"]] = r["sdt"]
        v[C["Email"]] = r["email"]
        v[C["Đầu mối GMPc"]] = r["dau_moi_gmpc"]
        v[C["Bên ký"]] = "GMPc"
        v[C["Nguồn dữ liệu"]] = "Excel"
        v[C["Tính vào thống kê"]] = "Có"
        v[C["Dòng Excel gốc"]] = r["excel_row"]
        return v

    notes_global = []
    for r in excel_rows:
        v, fills, notes = base_excel(r), {}, []
        if r["linh_vuc"] and not v[C["Lĩnh vực (chuẩn hoá)"]]:
            notes.append(f"Lĩnh vực '{r['linh_vuc']}' không thuộc danh mục chuẩn hoá")
        if r["excel_row"] in scan_by_row:
            sid, how = scan_by_row[r["excel_row"]]
            s = scan_by_id[sid]
            v[C["Nguồn dữ liệu"]] = "Excel + Scan"
            v[C["File scan"]] = f"{s['folder']}\\{s['file']}"
            v[C["Bên ký"]] = s["ben_ky"]
            v[C["Đối chiếu"]] = "Khớp" if how == "Số HĐ" else ("Khớp (thủ công)" if how.startswith("Thủ công") else "Khớp (theo tên)")
            if how.startswith("Thủ công"):
                notes.append(how)

            def fill_or_diff(col, scan_val, fmt=str, compare=None):
                cur = v[C[col]]
                if scan_val in (None, ""):
                    return
                if cur in (None, ""):
                    v[C[col]] = scan_val
                    fills[C[col]] = FILL_ADD
                elif compare and not compare(cur, scan_val):
                    fills[C[col]] = FILL_DIFF
                    notes.append(f"{col}: scan = {fmt(scan_val)}")

            fill_or_diff("Số hợp đồng", s["so_hd"], compare=lambda a, b: norm_contract(a) == norm_contract(b))
            fill_or_diff("Khách hàng", s["cong_ty"], compare=lambda a, b: company_sim(a, b) >= 0.6)
            if s["ngay"] and s["thang"]:
                v[C["Ngày ký (theo HĐ)"]] = f"{s['ngay']:02d}/{s['thang']:02d}/{s['nam']}"
                fills[C["Ngày ký (theo HĐ)"]] = FILL_ADD
            if s["thang"]:
                fill_or_diff("Tháng ký", f"{s['thang']:02d}/{s['nam']}",
                             compare=lambda a, b: str(a).strip() == str(b))
            gt = s.get("gia_tri")
            fill_or_diff("Giá trị trước VAT (VNĐ)", round(gt) if gt else None, fmt=lambda x: f"{x:,.0f}",
                         compare=lambda a, b: abs(float(a) - float(b)) <= 1000)
            if s.get("tong_gom_vat"):
                v[C["Tổng gồm VAT (VNĐ)"]] = round(s["tong_gom_vat"])
                fills[C["Tổng gồm VAT (VNĐ)"]] = FILL_ADD
            for col, key in (("MST", "mst"), ("Đại diện KH (theo HĐ)", "dai_dien"), ("Loại HĐ", "loai")):
                if s.get(key):
                    if col == "Loại HĐ" and v[C[col]]:
                        continue
                    v[C[col]] = s[key]
                    fills[C[col]] = FILL_ADD
            if not v[C["SĐT"]] and s.get("sdt_cty"):
                v[C["SĐT"]] = s["sdt_cty"]
                fills[C["SĐT"]] = FILL_ADD
            if s.get("gia_tri_note") and ("tính ra" in s["gia_tri_note"] or "không tách" in s["gia_tri_note"]):
                notes.append(f"Giá trị HĐ: {s['gia_tri_note']}")
            if s.get("ghi_chu"):
                notes.append(f"Scan: {s['ghi_chu']}")
            if any(f is FILL_DIFF for f in fills.values()):
                v[C["Đối chiếu"]] = "Lệch - cần kiểm tra"
        else:
            y = v[C["Năm"]]
            v[C["Đối chiếu"]] = "Chưa có scan" if (isinstance(y, int) and y >= 2021) else "Không có thư mục scan"
        if r["gia_tri"] and v[C["Nguồn dữ liệu"]] == "Excel" and isinstance(r["gia_tri"], float) and r["gia_tri"] % 1:
            notes.append("Giá trị trong Excel là công thức chia (tổng/1,08) - số lẻ")
        v[C["Ghi chú / điểm cần kiểm tra"]] = "; ".join(notes)
        out.append((v, fills, None))

    # Scan không khớp -> dòng mới
    for s in scans:
        if s["id"] in matched:
            continue
        v, notes = [""] * len(COLS), []
        v[C["Năm"]] = s["nam"] or ""
        v[C["Tháng ký"]] = f"{s['thang']:02d}/{s['nam']}" if s["thang"] else ""
        v[C["Ngày ký (theo HĐ)"]] = f"{s['ngay']:02d}/{s['thang']:02d}/{s['nam']}" if s["ngay"] and s["thang"] else ""
        v[C["Số hợp đồng"]] = s["so_hd"]
        v[C["Loại HĐ"]] = s["loai"]
        v[C["Khách hàng"]] = s["cong_ty"]
        v[C["MST"]] = s["mst"]
        v[C["Nội dung"]] = s["noi_dung"]
        v[C["Giá trị trước VAT (VNĐ)"]] = round(s["gia_tri"]) if s.get("gia_tri") else ""
        v[C["Tổng gồm VAT (VNĐ)"]] = round(s["tong_gom_vat"]) if s.get("tong_gom_vat") else ""
        v[C["Đại diện KH (theo HĐ)"]] = s["dai_dien"]
        v[C["SĐT"]] = s.get("sdt_cty", "")
        v[C["Bên ký"]] = s["ben_ky"]
        v[C["File scan"]] = f"{s['folder']}\\{s['file']}"
        v[C["Nguồn dữ liệu"]] = "Scan"
        if s["id"] in EXCLUDE:
            v[C["Đối chiếu"]] = "Loại"
            v[C["Tính vào thống kê"]] = "Không"
            notes.append(EXCLUDE[s["id"]])
            row_fill = FILL_OUT
        else:
            v[C["Đối chiếu"]] = "Chỉ có trong scan"
            v[C["Tính vào thống kê"]] = "Có"
            row_fill = FILL_NEW
        if s.get("gia_tri_note"):
            notes.append(f"Giá trị HĐ: {s['gia_tri_note']}")
        if s.get("ghi_chu"):
            notes.append(f"Scan: {s['ghi_chu']}")
        v[C["Ghi chú / điểm cần kiểm tra"]] = "; ".join(notes)
        out.append((v, {}, row_fill))

    # Sắp xếp: năm giảm dần, rồi tháng
    def sort_key(item):
        v = item[0]
        y = v[C["Năm"]] if isinstance(v[C["Năm"]], int) else 0
        m = re.match(r"(\d{2})/", str(v[C["Tháng ký"]]))
        return (-y, int(m.group(1)) if m else 99, v[C["Dòng Excel gốc"]] or 9999)

    out.sort(key=sort_key)
    for i, item in enumerate(out, 1):
        item[0][C["STT"]] = i
    return out


# ---------------------------------------------------------------- stats
def customer_key(v):
    if v[C["MST"]]:
        return "MST:" + re.sub(r"\D", "", str(v[C["MST"]]))[:10]
    n = norm_company(v[C["Khách hàng"]])
    return "TEN:" + n if n else None


def build_stats(master):
    rows = [v for v, _, _ in master if v[C["Tính vào thống kê"]] == "Có"]
    # Gộp khách hàng: tên -> MST nếu cùng tên đã có MST ở dòng khác
    name_to_mst = {}
    for v in rows:
        if v[C["MST"]] and v[C["Khách hàng"]]:
            name_to_mst[norm_company(v[C["Khách hàng"]])] = customer_key(v)
    keys = []
    for v in rows:
        k = customer_key(v)
        if k and k.startswith("TEN:"):
            k = name_to_mst.get(k[4:], k)
        keys.append(k)

    years = sorted({v[C["Năm"]] for v in rows if isinstance(v[C["Năm"]], int)}, reverse=True)
    by_year = []
    for y in years:
        idx = [i for i, v in enumerate(rows) if v[C["Năm"]] == y]
        gm = [i for i in idx if rows[i][C["Bên ký"]] != "OKS"]
        oks = [i for i in idx if rows[i][C["Bên ký"]] == "OKS"]
        vals = [rows[i][C["Giá trị trước VAT (VNĐ)"]] for i in gm if isinstance(rows[i][C["Giá trị trước VAT (VNĐ)"]], (int, float))]
        kh = {keys[i] for i in gm if keys[i]}
        no_id = sum(1 for i in gm if not keys[i])
        scan = sum(1 for i in gm if "Scan" in rows[i][C["Nguồn dữ liệu"]])
        by_year.append([y, len(gm), len(kh), no_id, len(vals), sum(vals) if vals else 0,
                        round(len(vals) / len(gm), 3) if gm else 0, scan, len(oks)])

    lv_counter = Counter((v[C["Năm"]], v[C["Lĩnh vực (chuẩn hoá)"]] or "(Chưa xác định)") for v in rows if v[C["Bên ký"]] != "OKS")
    lv_names = sorted({k[1] for k in lv_counter}, key=lambda x: (x == "(Chưa xác định)", x))
    tt_counter = Counter((v[C["Năm"]], v[C["Tình trạng (chuẩn hoá)"]] or "(Trống)") for v in rows if v[C["Bên ký"]] != "OKS")
    tt_names = sorted({k[1] for k in tt_counter}, key=lambda x: (x == "(Trống)", x))

    # Khách hàng
    cust = defaultdict(lambda: {"names": Counter(), "mst": "", "years": set(), "n": 0, "val": 0.0, "lv": Counter()})
    for v, k in zip(rows, keys):
        if not k or v[C["Bên ký"]] == "OKS":
            continue
        c = cust[k]
        c["names"][v[C["Khách hàng"]]] += 1
        if v[C["MST"]]:
            c["mst"] = v[C["MST"]]
        if isinstance(v[C["Năm"]], int):
            c["years"].add(v[C["Năm"]])
        c["n"] += 1
        if isinstance(v[C["Giá trị trước VAT (VNĐ)"]], (int, float)):
            c["val"] += v[C["Giá trị trước VAT (VNĐ)"]]
        if v[C["Lĩnh vực (chuẩn hoá)"]]:
            c["lv"][v[C["Lĩnh vực (chuẩn hoá)"]]] += 1
    cust_rows = []
    for k, c in cust.items():
        name = max(c["names"], key=lambda n: (len(n), c["names"][n])) if c["names"] else ""
        ys = sorted(c["years"])
        cust_rows.append([name, c["mst"], len(ys), min(ys) if ys else "", max(ys) if ys else "",
                          ", ".join(map(str, ys)), c["n"], c["val"] or "",
                          ", ".join(n for n, _ in c["lv"].most_common()), "Có" if len(ys) > 1 or c["n"] > 1 else ""])
    cust_rows.sort(key=lambda r: (-r[6], -(r[7] or 0)))
    return by_year, (lv_counter, lv_names, years), (tt_counter, tt_names), cust_rows


# ---------------------------------------------------------------- write
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = FILL_HEAD
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER


def write_table(ws, header, rows, start_row=1, widths=None, money_cols=(), pct_cols=(), name=None):
    for j, h in enumerate(header, 1):
        ws.cell(row=start_row, column=j, value=h)
    style_header(ws, start_row, len(header))
    for i, r in enumerate(rows, start_row + 1):
        for j, val in enumerate(r, 1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=isinstance(val, str) and len(val) > 30)
            if j - 1 in money_cols:
                cell.number_format = "#,##0"
            if j - 1 in pct_cols:
                cell.number_format = "0%"
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    if name and rows:
        ref = f"A{start_row}:{get_column_letter(len(header))}{start_row + len(rows)}"
        t = Table(displayName=name, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
        ws.add_table(t)


def main():
    excel_rows = load_excel_rows()
    scans = load_scans()
    matched = match(excel_rows, scans)
    master = build_master(excel_rows, scans, matched)
    by_year, lv, tt, cust_rows = build_stats(master)

    wb = openpyxl.Workbook()

    # 1. Hướng dẫn
    ws = wb.active
    ws.title = "Hướng dẫn"
    today = dt.date.today().strftime("%d/%m/%Y")
    n_new = sum(1 for v, _, _ in master if v[C["Đối chiếu"]] == "Chỉ có trong scan")
    n_diff = sum(1 for v, _, _ in master if v[C["Đối chiếu"]].startswith("Lệch"))
    lines = [
        ("THỐNG KÊ KHÁCH HÀNG ĐÃ KÝ HỢP ĐỒNG 2011–2026", True),
        (f"Ngày lập: {today}", False),
        ("", False),
        ("Nguồn dữ liệu", True),
        ("1. Excel gốc: '260625 Danh sách KH đã kí HĐ 2011-2026.xlsx' - sheet 'Tổng hợp' (bản copy từ NAS, không sửa).", False),
        (f"2. Hợp đồng scan: {SCAN_ROOT} - {len(scans)} file PDF (2021–2026 và OKS), đọc trực tiếp từ ảnh trang HĐ.", False),
        ("", False),
        ("Kết quả đối chiếu", True),
        (f"- {len(matched)} file scan khớp với dòng trong Excel; {n_new} HĐ chỉ có trong scan (đã thêm dòng mới); {n_diff} dòng có thông tin lệch giữa 2 nguồn.", False),
        ("- Năm 2011–2020 không có thư mục scan nên chỉ có dữ liệu Excel (thiếu Số HĐ, Giá trị).", False),
        ("", False),
        ("Màu trong sheet 'Danh sách HĐ'", True),
        ("Xanh lá  = ô trống trong Excel, đã BỔ SUNG từ hợp đồng scan", False),
        ("Vàng     = Excel và scan KHÁC NHAU - giữ giá trị Excel, xem cột 'Ghi chú' để biết giá trị trong scan", False),
        ("Xanh dương = HĐ CHỈ CÓ trong scan, chưa có trong Excel (dòng mới)", False),
        ("Xám      = Không tính vào thống kê (xem lý do ở cột Ghi chú)", False),
        ("", False),
        ("Quy ước", True),
        ("- Thông tin còn thiếu để TRỐNG, chưa suy đoán. Bổ sung dần trực tiếp vào sheet 'Danh sách HĐ'.", False),
        ("- 'Giá trị trước VAT' là giá trị chưa thuế. HĐ chỉ ghi tổng đã gồm VAT thì giá trước thuế được tính ngược (ghi chú rõ).", False),
        ("- HĐ liên danh Cửu Long: chỉ tính phần VNĐ của GMPc; phần USD của INOTEK không cộng vào.", False),
        ("- 'Lĩnh vực (chuẩn hoá)' và 'Tình trạng (chuẩn hoá)' là đề xuất gom nhóm - bảng quy đổi ở sheet 'Danh mục'.", False),
        ("- Dòng tiêu đề năm trong Excel gốc có công thức SUM (vd 38,27 tỷ năm 2025) - đã loại, không tính là HĐ.", False),
        ("- HĐ do Công ty CP OKS ký (thư mục OKS) được ghi 'Bên ký = OKS' và thống kê tách riêng.", False),
    ]
    for i, (t, bold) in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=t)
        c.font = Font(bold=bold, size=14 if i == 1 else 11)
    for i, f in zip(range(13, 17), (FILL_ADD, FILL_DIFF, FILL_NEW, FILL_OUT)):
        ws.cell(row=i, column=1).fill = f
    ws.column_dimensions["A"].width = 130

    # 2. Danh sách HĐ
    ws = wb.create_sheet("Danh sách HĐ")
    header = [n for n, _ in COLS]
    for j, h in enumerate(header, 1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, len(header))
    money = {C["Giá trị trước VAT (VNĐ)"], C["Tổng gồm VAT (VNĐ)"]}
    for i, (v, fills, row_fill) in enumerate(master, 2):
        for j, val in enumerate(v):
            cell = ws.cell(row=i, column=j + 1, value=val if val != "" else None)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=j in (C["Khách hàng"], C["Nội dung"], C["Ghi chú / điểm cần kiểm tra"], C["Số hợp đồng"]))
            if j in money:
                cell.number_format = "#,##0"
            if row_fill:
                cell.fill = row_fill
            elif j in fills:
                cell.fill = fills[j]
    for j, (_, w) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "H2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{len(master) + 1}"
    ws.row_dimensions[1].height = 32

    # 3. Thống kê theo năm
    ws = wb.create_sheet("TK theo năm")
    hdr = ["Năm", "Số HĐ (GMPc)", "Số KH (xác định được)", "HĐ không xác định được KH", "Số HĐ có giá trị",
           "Tổng giá trị trước VAT (VNĐ)", "Tỷ lệ HĐ có giá trị", "Số HĐ có scan", "Số HĐ OKS"]
    tot = ["Tổng", sum(r[1] for r in by_year), "", sum(r[3] for r in by_year), sum(r[4] for r in by_year),
           sum(r[5] for r in by_year), "", sum(r[7] for r in by_year), sum(r[8] for r in by_year)]
    write_table(ws, hdr, by_year + [tot], widths=[8, 12, 14, 16, 14, 22, 12, 12, 10], money_cols={5}, pct_cols={6})
    last = len(by_year) + 2
    for j in range(1, len(hdr) + 1):
        ws.cell(row=last, column=j).font = Font(bold=True)
    ws.cell(row=last + 2, column=1, value="Lưu ý: Tổng giá trị chỉ cộng các HĐ đã có giá trị; năm có tỷ lệ thấp thì tổng chưa phản ánh đủ.")
    ws.cell(row=last + 3, column=1, value="Số KH đếm theo MST (nếu có) hoặc tên công ty; 'Tổng' số KH toàn giai đoạn xem sheet 'Khách hàng'.")
    ws.freeze_panes = "A2"

    # 4. Theo lĩnh vực
    ws = wb.create_sheet("TK theo lĩnh vực")
    counter, names, years = lv
    rows = [[y] + [counter.get((y, n), 0) for n in names] + [sum(counter.get((y, n), 0) for n in names)] for y in years]
    rows.append(["Tổng"] + [sum(counter.get((y, n), 0) for y in years) for n in names] + [sum(counter.values())])
    write_table(ws, ["Năm"] + names + ["Tổng"], rows, widths=[8] + [14] * (len(names) + 1))
    ws.freeze_panes = "B2"

    # 5. Theo tình trạng
    ws = wb.create_sheet("TK theo tình trạng")
    counter, names = tt
    rows = [[y] + [counter.get((y, n), 0) for n in names] + [sum(counter.get((y, n), 0) for n in names)] for y in years]
    write_table(ws, ["Năm"] + names + ["Tổng"], rows, widths=[8] + [16] * (len(names) + 1))

    # 6. Khách hàng
    ws = wb.create_sheet("Khách hàng")
    write_table(ws, ["Khách hàng", "MST", "Số năm có HĐ", "Năm đầu", "Năm gần nhất", "Các năm", "Số HĐ",
                     "Tổng giá trị trước VAT (VNĐ)", "Lĩnh vực", "KH quay lại"],
                cust_rows, widths=[45, 14, 10, 9, 10, 30, 8, 20, 25, 10], money_cols={7}, name="KhachHang")
    ws.freeze_panes = "B2"

    # 7. Đối chiếu scan
    ws = wb.create_sheet("Đối chiếu scan")
    rows = []
    for s in scans:
        if s["id"] in EXCLUDE:
            st, how, row = "Loại", EXCLUDE[s["id"]], ""
        elif s["id"] in matched:
            r, how = matched[s["id"]]
            st, row = "Khớp Excel", r["excel_row"]
        else:
            st, how, row = "Chỉ có trong scan", "", ""
        rows.append([s["id"], s["folder"], s["file"], s["so_hd"], s["ngay_ky"], s["cong_ty"], s["mst"],
                     s.get("gia_tri") or "", s.get("tong_gom_vat") or "", s["ben_ky"], st, how, row, s.get("ghi_chu", "")])
    write_table(ws, ["Mã", "Thư mục", "Tên file", "Số HĐ (theo scan)", "Ngày ký", "Bên A (KH)", "MST",
                     "Giá trị trước VAT", "Tổng gồm VAT", "Bên ký", "Kết quả", "Cách khớp / lý do", "Dòng Excel", "Ghi chú khi đọc"],
                rows, widths=[6, 7, 40, 30, 11, 40, 13, 16, 16, 12, 16, 26, 9, 50], money_cols={7, 8}, name="DoiChieu")
    ws.freeze_panes = "D2"

    # 8. Cần kiểm tra
    ws = wb.create_sheet("Cần kiểm tra")
    rows = [[v[C["STT"]], v[C["Năm"]], v[C["Số hợp đồng"]], v[C["Khách hàng"]], v[C["Đối chiếu"]], v[C["Ghi chú / điểm cần kiểm tra"]]]
            for v, _, _ in master if v[C["Ghi chú / điểm cần kiểm tra"]] or v[C["Đối chiếu"]].startswith("Lệch")]
    write_table(ws, ["STT (Danh sách HĐ)", "Năm", "Số HĐ", "Khách hàng", "Đối chiếu", "Nội dung cần kiểm tra"],
                rows, widths=[10, 7, 30, 40, 18, 90], name="CanKiemTra")
    ws.freeze_panes = "A2"

    # 9. Danh mục
    ws = wb.create_sheet("Danh mục")
    write_table(ws, ["Lĩnh vực (gốc, viết thường)", "Lĩnh vực (chuẩn hoá)"], sorted(LINH_VUC_MAP.items()), widths=[35, 25])
    start = len(LINH_VUC_MAP) + 4
    write_table(ws, ["Tình trạng (gốc)", "Tình trạng (chuẩn hoá)"], sorted(TINH_TRANG_MAP.items()), start_row=start)

    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, f"{dt.date.today():%y%m%d} Thong ke KH da ky HD 2011-2026.xlsx")
    wb.save(out)
    print("Saved:", out)
    print(f"Excel rows: {len(excel_rows)} | scans: {len(scans)} | matched: {len(matched)} | "
          f"scan-only: {n_new} | diff rows: {n_diff} | master rows: {len(master)}")
    for s in scans:
        if s["id"] in matched and matched[s["id"]][1] != "Số HĐ":
            r = matched[s["id"]][0]
            print(f"  fuzzy {s['id']} {s['cong_ty'][:40]!r} -> R{r['excel_row']} {r['cong_ty'][:40]!r} | {matched[s['id']][1]}")
    print("Scan-only:", [s["id"] for s in scans if s["id"] not in matched and s["id"] not in EXCLUDE])


if __name__ == "__main__":
    main()
