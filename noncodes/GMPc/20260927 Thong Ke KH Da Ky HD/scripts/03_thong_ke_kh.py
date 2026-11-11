"""Bước 3: Thống kê KH đã ký HĐ theo năm / theo khách hàng / KH quay lại + kiểm tra lỗi dữ liệu.

Nguồn chính (theo yêu cầu user 2026-10-04): sheet 'Tổng hợp' của file Excel gốc trên NAS
(bản copy trong data/raw/). Scan HĐ chỉ dùng để đánh dấu "đã có scan" và đối chiếu.

Input:
    data/raw/261005 Danh sách KH đã kí HĐ 2011-2026.xlsx   (sheet 'Tổng hợp')
    data/processed/scan_index.csv + scan_extract.jsonl      (dữ liệu đọc từ 135 HĐ scan 2021-2026, lần 2026-09-27)
    SCAN_ROOT trên NAS (chỉ liệt kê tên file, không mở/sửa)
Output:
    output/<YYMMDD> Thong ke KH va CSKH 2011-2026.xlsx

Quy ước:
- Không sửa file gốc, không suy đoán: thiếu thì để trống.
- Ô gộp (merge) trong Excel gốc được hiểu là cùng giá trị cho các dòng trong vùng gộp (vd tên KH của HĐ ký thêm).
- "KH quay lại": HĐ mới cách HĐ trước đó của cùng KH >= GAP_QUAY_LAI_THANG tháng
  (HĐ cách nhau ít hơn được coi là cùng dự án / ký thêm gói).
"""
import csv
import datetime as dt
import json
import os
import re
import unicodedata
from collections import Counter, defaultdict

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PROJECT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_NAME = "261005 Danh sách KH đã kí HĐ 2011-2026.xlsx"
RAW = os.path.join(PROJECT, "data", "raw", RAW_NAME)
NAS_SRC = r"L:\04. Tổng hợp Data khách hàng\01. Danh sách KH đã kí HĐ" + "\\" + RAW_NAME
SCAN_ROOT = r"L:\04. Tổng hợp Data khách hàng\01. Danh sách KH đã kí HĐ\Scan hợp đồng tư vấn"
SCAN_JSONL = os.path.join(PROJECT, "data", "processed", "scan_extract.jsonl")
SCAN_INDEX = os.path.join(PROJECT, "data", "processed", "scan_index.csv")
OUT_DIR = os.path.join(PROJECT, "output")
TODAY = dt.date.today()

GAP_QUAY_LAI_THANG = 12      # >= 12 tháng kể từ HĐ trước -> tính là KH quay lại (user: "thường sau 1-3 năm")
DATA_TO = (2026, 9)          # dữ liệu 2026 chưa đủ năm

FONT_NAME = "Calibri Light"  # font chuẩn tài liệu GMPc

# ------------------------------------------------------------------ danh mục
LINH_VUC_MAP = {
    "dược phẩm": "Dược phẩm", "who gmp": "Dược phẩm", "dược phẩm who gmp": "Dược phẩm",
    "đông dược who gmp": "Dược phẩm", "đông dược": "Dược phẩm", "thuốc phóng xạ": "Dược phẩm",
    "phóng xạ, hạt nhân": "Dược phẩm", "dược phẩm/tp bvsk": "Dược phẩm", "dược phẩm/dược liệu": "Dược phẩm",
    "dược phẩm/đông dược": "Dược phẩm", "dược phẩm lập báo cáo đầu tư": "Dược phẩm",
    "dược phẩm/lập cáo cáo đầu tư": "Dược phẩm", "dược phẩm/kho gsp": "Dược phẩm", "dược phẩm kho gsp": "Dược phẩm",
    "mỹ phẩm": "Mỹ phẩm", "mỹ phẩm cgmp": "Mỹ phẩm", "mỹ phẩm - thiết bị y tế": "Mỹ phẩm",
    "tpbvsk": "TPBVSK", "tp bvsk": "TPBVSK", "hs gmp": "TPBVSK", "hs - gmp": "TPBVSK",
    "thực phẩm bvsk - hs gmp": "TPBVSK", "thực phẩm": "Thực phẩm", "sữa chua": "Thực phẩm",
    "thú y thủy sản": "Thú y", "thú y": "Thú y", "thuốc thú y": "Thú y",
    "kho gsp": "Kho GSP/GDP", "kho gsp/gdp": "Kho GSP/GDP", "kho gdp": "Kho GSP/GDP", "gdp": "Kho GSP/GDP",
    "kho gdp, điểm tiêm vaccine": "Kho GSP/GDP",
    "bao bì dược phẩm": "Bao bì dược phẩm", "bao bì dp": "Bao bì dược phẩm",
    "thiết bị y tế": "Thiết bị y tế", "thiết bị y tế - iso": "Thiết bị y tế",
    "vacxin": "Sinh phẩm/Vaccine/Tế bào", "tế bào gốc": "Sinh phẩm/Vaccine/Tế bào", "cgmp": "Sinh phẩm/Vaccine/Tế bào",
    "lập báo cáo nghiên cứu khả thi": "Khác", "dv công bố sp": "Khác", "phòng lab": "Khác",
}
TINH_TRANG_MAP = {"đã hoàn thành": "Đã hoàn thành", "đã thanh lý": "Đã thanh lý",
                  "đang thực hiện": "Đang thực hiện", "đang triển khai": "Đang thực hiện"}

# Gộp KH có nhiều cách ghi tên (key sau chuẩn hoá -> key nhóm). Đã rà thủ công 2026-10-04.
ALIAS = {
    "greeblab vn": "greenlab thuoc thu y gmp who",
    "logistic hang khong": "logistics hang khong",
    "truong cd hai duong": "truong cao dang hai duong",
    "vien huyet hoc truyen mau tw": "vien huyet hoc truyen mau",
    "thiet bi y te ha noi hapharco": "thiet bi y te ha noi",
    "ky thuat co dien lanh qui long": "co dien lanh qui long",
    "a chau nha may thu y gmp who": "thuoc thu y a chau",
    "bao bi in nong nghiep app": "bao bi in nong nghiep",
    "world essence usa": "happy secret",
    "benh vien da khoa vinmec": "vien nghien cuu te bao goc gen vinmec thai son",
    "xay dung m e hoang sa ebc": "ebc",      # user xác nhận gộp 2026-10-04
    "e b c giang dien": "ebc",
    "tu van delta": "delta",
}
# KH khác pháp nhân nhưng cùng tập đoàn: vẫn tính là KH riêng (công ty con liên hệ riêng), ghi chú nhóm
GROUP_NOTE = {"fpt": "Tập đoàn FPT", "fpt long chau": "Tập đoàn FPT"}  # DPharma (Thế Giới Số) KHÔNG thuộc FPT (user 2026-10-04)
ALIAS_NOTE = {
    "happy secret": "Gộp Happy Secret với World Essence USA (tên ghi '(Happy Secret)') - khác pháp nhân, cùng nhóm KH",
    "vien nghien cuu te bao goc gen vinmec thai son": "Gộp BV ĐKQT Vinmec (HĐ 2023) với Viện NC Tế bào gốc Vinmec - scan HĐ 2023 do Viện ký",
    "thuoc thu y a chau": "Gộp 'Á Châu - Nhà máy dược thú y' (2011) với 'Thuốc thú y Á Châu' (2013), cùng Cái Răng - Cần Thơ",
    "bao bi in nong nghiep": "Gộp 'Bao bì và In Nông nghiệp' (2021) với '... APP' (2013)",
    "ebc": "Gộp EBC Group, Hoàng Sa - EBC, E.B.C mỹ phẩm Giang Điền (user xác nhận 2026-10-04)",
    "delta": "Gộp Quốc tế Delta (2021) và Tư vấn Delta Việt Nam (2024) (user xác nhận 2026-10-04)",
}
# Cặp tên giống nhau nhưng KHÔNG gộp (khác KH)
NOT_MERGED = [
    ("co so dong cuu long", "cuu long", "Đông dược Cửu Long (Cần Thơ, 2014) và CP Dược phẩm Cửu Long (2025) - user xác nhận khác KH"),
    ("fpt", "fpt long chau", "Dược phẩm FPT (2023) và FPT Long Châu (2024) - cùng tập đoàn, công ty con liên hệ riêng → KH riêng (user 2026-10-04)"),
    ("co dien phong sach anh khang", "dabaco", "Anh Khang (Dabaco) 2022 là nhà thầu cho dự án Dabaco; Dabaco ký trực tiếp 2023"),
    ("vincos", "vncos", "Vincos (Hà Nội) và VNCOS (HCM) - khác KH (chỉ giống tên)"),
]

# Dòng không tính vào thống kê (dòng Excel gốc -> lý do). Vẫn hiện trong 'Danh sách HĐ' (xám) và 'Lỗi cần sửa'.
# (2026-10-04: user đã xoá dòng Kim Thành trùng trong file gốc -> hiện không loại dòng nào. Số dòng thay đổi khi file gốc sửa!)
EXCLUDE_ROWS = {}

SERVICE_RULES = [  # (nhóm dịch vụ, regex trên nội dung đã bỏ dấu, viết thường)
    ("Lập dự án / BC đầu tư", r"kha thi|de xuat dau tu|du an dau tu|co h[oo]i dau tu|ke hoach tong the|khai toan|tong muc dau tu|bao cao dau tu"),
    ("Quy hoạch & thiết kế", r"thiet ke|mat bang|quy hoach|layout|m&e|ban ve|du toan"),
    ("Thi công / cung cấp", r"thi cong|lap dat|cung cap|cai tao he thong|san vinyl"),
    ("Giám sát / QLDA", r"giam sat|quan ly du an|qlda|dieu hanh"),
    ("Đào tạo & hồ sơ GMP", r"dao tao|huan luyen|tap huan|soan thao|ho so|phan mem|pif|cong bo"),
    ("Đánh giá / tái đánh giá", r"tai danh gia|tu thanh tra|hau kiem|danh gia|tham dinh|ra soat|khao sat hien trang|phan bien"),
    ("Bảo trì / sửa chữa", r"bao tri|bao duong|sua chua"),
]

# Màu theo mẫu biểu GMPc mới (MDC-QT-03.BM01): tiêu đề navy, hàng tiêu đề cột teal
NAVY, TEAL, GREY = "1B1464", "00B5A5", "F2F2F2"
FILL_NAVY = PatternFill("solid", fgColor=NAVY)
FILL_TEAL = PatternFill("solid", fgColor=TEAL)
FILL_ERR = PatternFill("solid", fgColor="F8CBAD")    # lỗi cần sửa
FILL_WARN = PatternFill("solid", fgColor="FFF2CC")   # cần kiểm tra
FILL_INFO = PatternFill("solid", fgColor="DDEBF7")   # thông tin
FILL_BACK = PatternFill("solid", fgColor="E2EFDA")   # quay lại
FILL_OUT = PatternFill("solid", fgColor="D9D9D9")    # không tính thống kê
FILL_TOTAL = PatternFill("solid", fgColor="D9F2EF")   # dòng tổng (teal nhạt)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ------------------------------------------------------------------ helpers
def sa(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def clean(v):
    if isinstance(v, str):
        v = re.sub(r"\s+", " ", v).strip()
        return v
    return "" if v is None else v


STOP = ("cong ty", "co phan", "tnhh", "mtv", "mot thanh vien", "tap doan", "chi nhanh", "viet nam", "vietnam",
        "cp", "jsc", "ltd", "group", "san xuat", "thuong mai", "dau tu", "quoc te", "duoc pham", "duoc my pham",
        "my pham", "duoc", "phat trien", "cty", "ct", "xuat nhap khau", "dich vu", "tm", "sx", "dv", "xnk", "va",
        "cong nghe", "trung uong")


def name_key(s):
    s = sa(s).lower()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    for w in STOP:
        s = re.sub(rf"\b{w}\b", " ", s)
    k = re.sub(r"\s+", " ", s).strip()
    return ALIAS.get(k, k)


def norm_contract(s):
    s = sa(s).upper()
    s = re.sub(r"^\s*SO\s*", "", s)
    return re.sub(r"[^A-Z0-9]", "", s)


def lv_key(s):
    s = clean(s).lower().replace("thuỷ", "thủy")
    s = re.sub(r"\s*/\s*", "/", s)
    return s


PROVINCES = [  # (mẫu tìm trong tên đã bỏ dấu, tên tỉnh/thành theo cách ghi trong dữ liệu - trước sáp nhập 2025)
    (r"ho chi minh|hcm|tphcm|sai gon|q\.\s*1|quan 7|hoc mon|binh tan|phu nhuan", "TP HCM"),
    (r"ha noi|gia lam|soc son|ba dinh|thanh tri|thanh xuan", "Hà Nội"),
    (r"hai phong", "Hải Phòng"), (r"da nang", "Đà Nẵng"), (r"can tho|ninh kieu|cai rang|binh thuy", "Cần Thơ"),
    (r"hung yen|yen my", "Hưng Yên"), (r"tay ninh", "Tây Ninh"), (r"long an", "Long An"),
    (r"dong nai|trang bom", "Đồng Nai"), (r"binh duong", "Bình Dương"), (r"ha nam", "Hà Nam"),
    (r"hai duong", "Hải Dương"), (r"bac ninh", "Bắc Ninh"), (r"vinh phuc", "Vĩnh Phúc"), (r"phu tho", "Phú Thọ"),
    (r"lang son", "Lạng Sơn"), (r"quang ninh", "Quảng Ninh"), (r"lao cai", "Lào Cai"), (r"yen bai", "Yên Bái"),
    (r"ha giang", "Hà Giang"), (r"thai binh", "Thái Bình"), (r"ninh binh", "Ninh Bình"), (r"hoa binh", "Hòa Bình"),
    (r"quang tri", "Quảng Trị"), (r"lam dong|da lat", "Lâm Đồng"), (r"khanh hoa|nha trang", "Khánh Hòa"),
    (r"dak lak|dac lac", "Đắk Lắk"), (r"ben tre", "Bến Tre"), (r"dong thap", "Đồng Tháp"),
    (r"tien giang", "Tiền Giang"), (r"vinh long|binh minh", "Vĩnh Long"), (r"tra vinh", "Trà Vinh"),
    (r"bac lieu", "Bạc Liêu"), (r"thai nguyen|pho yen", "Thái Nguyên"), (r"duc hoa", "Long An"),
]


def province(name):
    t = sa(name).lower()
    parts = re.findall(r"\((.*?)\)", t) + [t]
    for p in parts:
        for pat, prov in PROVINCES:
            if re.search(rf"\b(?:{pat})\b", p):
                return prov
    return ""


def parse_date(v):
    """-> (ngày|None, tháng|None, năm|None). Ô có thể là text MM/YYYY, DD/MM/YYYY, M/YYYY hoặc datetime."""
    if v in (None, ""):
        return None, None, None
    if isinstance(v, dt.datetime):
        return (v.day if v.day != 1 else None), v.month, v.year
    s = str(v).strip()
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    m = re.fullmatch(r"(\d{1,2})/(\d{4})", s)
    if m:
        return None, int(m.group(1)), int(m.group(2))
    m = re.search(r"(\d{4})", s)
    return (None, None, int(m.group(1))) if m else (None, None, None)


def services(text, so_hd):
    t = sa(text).lower()
    code = sa(so_hd).upper()
    t2 = re.sub(r"giam sat thi cong|ban ve (ky thuat )?thi cong|thiet ke (ban ve )?thi cong|ho so thiet ke", " ", t)
    out = []
    for name, pat in SERVICE_RULES:
        src = t2 if name == "Thi công / cung cấp" else t
        if re.search(pat, src):
            out.append(name)
    if ("HDTC" in code or "HDKT" in code) and "Thi công / cung cấp" not in out and re.search(r"thi cong|cung cap|lap dat", t):
        out.append("Thi công / cung cấp")
    return out or ["Khác"]


def fmt_ym(y, m):
    return f"{m:02d}/{y}" if (y and m) else (str(y) if y else "")


# ------------------------------------------------------------------ load
def load_rows():
    wb_f = openpyxl.load_workbook(RAW)                      # công thức
    wb_v = openpyxl.load_workbook(RAW, data_only=True)      # giá trị
    ws_f, ws_v = wb_f["Tổng hợp"], wb_v["Tổng hợp"]
    merged = {}
    for mr in ws_f.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != (mr.min_row, mr.min_col):
                    merged[(r, c)] = (mr.min_row, mr.min_col)

    def val(r, c, ws=ws_v):
        rr, cc = merged.get((r, c), (r, c))
        return ws.cell(rr, cc).value

    rows, year, year_heads = [], None, []
    for r in range(2, ws_v.max_row + 1):
        b = ws_v.cell(r, 2).value
        if isinstance(b, int) and 2000 < b < 2100:
            year = b
            year_heads.append((r, b, ws_f.cell(r, 6).value))
            continue
        raw = [clean(val(r, c)) for c in range(1, 13)]
        if not any(raw[1:6]):
            continue
        stt, cty, tdk, sohd, nd, gt, tt, lv, dm, sdt, email, gmpc = raw
        f_formula = ws_f.cell(*merged.get((r, 6), (r, 6))).value
        rec = {
            "row": r, "block": year, "stt": stt, "cong_ty": cty, "tdk_raw": val(r, 3), "so_hd": str(sohd),
            "noi_dung": nd, "gia_tri": float(gt) if isinstance(gt, (int, float)) else None,
            "gia_tri_formula": f_formula if isinstance(f_formula, str) and f_formula.startswith("=") else "",
            "tinh_trang": tt, "linh_vuc": lv, "dau_moi": dm, "sdt": str(sdt), "email": str(email), "gmpc": gmpc,
            "quyet_toan": "", "gcn": "", "ghi_chu_goc": "", "notes": [],
            "company_merged": (r, 2) in merged, "row_merged": any((r, c) in merged for c in (2, 3)),
        }
        # --- cột liên hệ bị lệch ở các khối năm cũ
        if "@" in rec["sdt"] and not rec["email"]:
            s = rec["sdt"]
            if ":" in s:
                rec["dau_moi"] = rec["dau_moi"] or s.split(":")[0].strip()
                s = s.split(":", 1)[1]
            rec["email"], rec["sdt"] = s.strip().lstrip("'"), ""
        if lv_key(rec["sdt"]) in LINH_VUC_MAP and not LINH_VUC_MAP.get(lv_key(rec["linh_vuc"])):
            rec["ghi_chu_goc"] = rec["linh_vuc"]          # khối 2015: cột H là ghi chú, lĩnh vực nằm ở cột J
            rec["linh_vuc"], rec["sdt"] = rec["sdt"], ""
        if year and year <= 2013:
            if rec["email"] and "@" not in rec["email"]:
                rec["quyet_toan"], rec["email"] = rec["email"], ""
            if rec["gmpc"]:
                rec["gcn"], rec["gmpc"] = str(rec["gmpc"]), ""
        for k in ("quyet_toan", "gcn"):
            if isinstance(rec[k], dt.datetime):
                rec[k] = rec[k].strftime("%d/%m/%Y")
        if isinstance(rec["gmpc"], (int, float)) or re.fullmatch(r"\d+", str(rec["gmpc"])):
            rec["gmpc"] = ""
        if rec["sdt"].lower() in ("không có số",):
            rec["sdt"] = ""
        rec["ngay"], rec["thang"], rec["nam"] = parse_date(rec["tdk_raw"])
        rows.append(rec)
    return rows, year_heads


def load_scan_extract():
    if not os.path.exists(SCAN_JSONL):
        return {}
    idx = {r["id"]: r for r in csv.DictReader(open(SCAN_INDEX, encoding="utf-8-sig"))}
    out = {}
    for line in open(SCAN_JSONL, encoding="utf-8"):
        s = json.loads(line)
        out[(idx[s["id"]]["folder"], idx[s["id"]]["file"])] = s
    return out


def list_scans():
    files = []
    if not os.path.isdir(SCAN_ROOT):
        return files
    for folder in sorted(os.listdir(SCAN_ROOT)):
        p = os.path.join(SCAN_ROOT, folder)
        if os.path.isdir(p):
            for f in sorted(os.listdir(p)):
                if f.lower().endswith(".pdf"):
                    files.append((folder, f))
    return files


# ------------------------------------------------------------------ customers
def assign_customers(rows):
    for r in rows:
        r["kh_key"] = name_key(r["cong_ty"]) if r["cong_ty"] else ""
    groups = defaultdict(list)
    for r in rows:
        if r["kh_key"]:
            groups[r["kh_key"]].append(r)
    ids = {}
    # mã KH theo năm HĐ đầu tiên rồi tên
    order = sorted(groups, key=lambda k: (min((x["nam"] or x["block"] or 9999) for x in groups[k]), k))
    for i, k in enumerate(order, 1):
        ids[k] = f"KH{i:03d}"
    for r in rows:
        r["kh_id"] = ids.get(r["kh_key"], "")
    return groups, ids


def ym_index(r):
    """Tháng tuyệt đối để tính khoảng cách; thiếu tháng -> giữa năm (ước tính)."""
    y = r["nam"] or r["block"]
    if not y:
        return None, True
    return y * 12 + ((r["thang"] or 7) - 1), r["thang"] is None


def episodes(groups):
    """Chia HĐ của mỗi KH thành các 'đợt' (dự án). Đợt mới khi cách HĐ trước >= GAP_QUAY_LAI_THANG."""
    events = []
    for k, rs in groups.items():
        rs.sort(key=lambda r: (ym_index(r)[0] or 0, r["row"]))
        ep, prev = 0, None
        for r in rs:
            idx, est = ym_index(r)
            if prev is None:
                ep = 1
                r["loai_hd"] = "HĐ đầu tiên"
                r["gap"] = ""
            else:
                gap = idx - prev[0] if idx is not None and prev[0] is not None else None
                r["gap"] = gap if gap is not None else ""
                r["gap_est"] = est or prev[1]
                if gap is not None and gap >= GAP_QUAY_LAI_THANG:
                    ep += 1
                    r["loai_hd"] = "KH quay lại"
                    events.append((k, prev[2], r, gap, est or prev[1]))
                else:
                    r["loai_hd"] = "Ký thêm (cùng dự án)"
            r["dot"] = ep
            prev = (idx, est, r)
    return events


# ------------------------------------------------------------------ checks
def checks(rows, year_heads, groups):
    issues = []  # (mức, dòng, KH, loại, mô tả, đề xuất)

    def add(level, r, kind, desc, fix=""):
        issues.append([level, r["row"] if r else "", (r["cong_ty"][:70] if r else ""), kind, desc, fix])

    # tổng năm (công thức SUM ở dòng tiêu đề năm)
    for i, (hr, y, f) in enumerate(year_heads):
        last = (year_heads[i + 1][0] - 1) if i + 1 < len(year_heads) else max(r["row"] for r in rows)
        block_rows = [r["row"] for r in rows if r["block"] == y]
        if isinstance(f, str) and f.startswith("=SUM"):
            m = re.match(r"=SUM\(F(\d+):F(\d+)\)", f.replace("$", ""))
            if m and block_rows and (int(m.group(1)) > min(block_rows) or int(m.group(2)) < max(block_rows)):
                miss = [x for x in block_rows if x < int(m.group(1)) or x > int(m.group(2))]
                issues.append(["Lỗi", hr, f"Tiêu đề năm {y}", "Tổng năm sai vùng",
                               f"Công thức {f} không bao gồm dòng {', '.join(map(str, miss))}",
                               f"Sửa thành =SUM(F{min(block_rows)}:F{max(block_rows)})"])

    seen_no = defaultdict(list)
    for r in rows:
        y_blk = r["block"]
        if r["so_hd"]:
            seen_no[norm_contract(r["so_hd"])].append(r)
        # thiếu thông tin
        if not r["tdk_raw"]:
            add("Cần bổ sung", r, "Thiếu thời điểm ký", "Cột 'Thời điểm kí' trống", "Bổ sung tháng/năm ký")
        if not r["so_hd"] and (y_blk or 0) >= 2016:
            add("Cần bổ sung", r, "Thiếu số HĐ", "Cột 'Số hợp đồng' trống")
        elif r["so_hd"] and not re.search(r"\d", r["so_hd"][:6]) and len(r["so_hd"]) < 25:
            add("Cần kiểm tra", r, "Số HĐ không đầy đủ", f"Số HĐ '{r['so_hd']}' thiếu phần số")
        if r["gia_tri"] is None and (y_blk or 0) >= 2016:
            add("Cần bổ sung", r, "Thiếu giá trị", "Cột 'Giá trị' trống")
        if not r["linh_vuc"]:
            add("Cần bổ sung", r, "Thiếu lĩnh vực", "Cột 'Lĩnh vực' trống")
        elif not LINH_VUC_MAP.get(lv_key(r["linh_vuc"])):
            add("Cần kiểm tra", r, "Lĩnh vực lạ", f"'{r['linh_vuc']}' chưa có trong danh mục chuẩn hoá")
        if r["tinh_trang"] and not TINH_TRANG_MAP.get(r["tinh_trang"].lower()):
            add("Cần kiểm tra", r, "Tình trạng lạ", f"'{r['tinh_trang']}'")
        y_ = r["nam"] or y_blk
        if TINH_TRANG_MAP.get(r["tinh_trang"].lower()) == "Đang thực hiện" and y_ and y_ <= TODAY.year - 3:
            add("Cần kiểm tra", r, "Tình trạng chưa cập nhật", f"HĐ năm {y_} vẫn ghi '{r['tinh_trang']}'",
                "Cập nhật Đã hoàn thành / Đã thanh lý")
        # năm
        if r["nam"] and y_blk and r["nam"] != y_blk:
            add("Lỗi", r, "Sai khối năm", f"Thời điểm ký {r['tdk_raw'] if not isinstance(r['tdk_raw'], dt.datetime) else r['tdk_raw']:%m/%Y} nhưng nằm trong khối năm {y_blk}"
                if isinstance(r["tdk_raw"], dt.datetime) else f"Thời điểm ký '{r['tdk_raw']}' nhưng nằm trong khối năm {y_blk}",
                "Chuyển dòng sang đúng năm hoặc sửa ngày")
        m = re.search(r"(?<=[/\-_ ])(20[12]\d)(?!\d)", r["so_hd"])   # năm sau dấu '/', bỏ qua 'DDMM/' đầu số HĐ
        y_sign = r["nam"] or y_blk
        if m and y_sign and abs(int(m.group(1)) - y_sign) >= 1:
            add("Cần kiểm tra", r, "Năm trong số HĐ khác năm ký",
                f"Số HĐ '{r['so_hd']}' ghi năm {m.group(1)}, thời điểm ký năm {y_sign}")
        # tháng trong số HĐ (quy ước DDMM/...)
        m = re.match(r"\s*(?:Số\s*)?(\d{2})(\d{2})[/_]", r["so_hd"])
        if m and r["thang"] and 1 <= int(m.group(1)) <= 31 and 1 <= int(m.group(2)) <= 12 and int(m.group(2)) != r["thang"]:
            add("Cần kiểm tra", r, "Tháng ký ≠ số HĐ",
                f"Thời điểm ký tháng {r['thang']:02d}/{r['nam']}, số HĐ '{r['so_hd'].strip()}' gợi ý ngày {m.group(1)}/{m.group(2)}")
        # giá trị
        if r["gia_tri_formula"]:
            f = r["gia_tri_formula"]
            if re.search(r"/\s*1\.0?8|/\s*1\.1\b", f):
                add("Thông tin", r, "Giá trị tính ngược từ VAT", f"Công thức {f} (tổng gồm VAT ÷ hệ số) → số lẻ",
                    "Nên nhập giá trước VAT theo HĐ")
            elif "+" in f:
                add("Cần kiểm tra", r, "Gộp nhiều HĐ trong 1 dòng", f"Giá trị {f} - có thể là 2 HĐ", "Tách thành 2 dòng nếu là 2 HĐ")
        if r["gia_tri"] and r["gia_tri"] >= 10e9:
            add("Thông tin", r, "Giá trị rất lớn", f"{r['gia_tri']:,.0f} đ", "Xác nhận đúng phần giá trị của GMPc")
        if re.match(r"^\(", r["cong_ty"] or ""):
            add("Cần kiểm tra", r, "Tên KH lẫn tên người", f"Tên bắt đầu bằng '{r['cong_ty'][:25]}…'", "Đưa tên người liên hệ sang cột Đầu mối")
        if "cá nhân" in (r["cong_ty"] or "").lower():
            add("Thông tin", r, "Ghi chú trong tên KH", "Tên KH có chữ '- cá nhân'")

    # số HĐ trùng
    for k, rs in seen_no.items():
        if k and len(rs) > 1:
            for r in rs:
                others = ", ".join(str(x["row"]) for x in rs if x is not r)
                add("Lỗi" if len({x["kh_key"] for x in rs}) == 1 else "Cần kiểm tra", r, "Trùng số HĐ",
                    f"Số HĐ '{r['so_hd']}' trùng với dòng {others}", "Kiểm tra có phải nhập trùng HĐ không")
    # nghi trùng: cùng KH, cùng giá trị
    for k, rs in groups.items():
        for i, a in enumerate(rs):
            for b in rs[i + 1:]:
                same_val = a["gia_tri"] and b["gia_tri"] and abs(a["gia_tri"] - b["gia_tri"]) < 1
                same_txt = sa(a["noi_dung"]).lower()[:40] == sa(b["noi_dung"]).lower()[:40] and a["noi_dung"]
                both_numbered = all(re.match(r"\s*\d", x["so_hd"]) for x in (a, b))   # 2 số HĐ khác nhau -> 2 HĐ thật
                if (same_val or same_txt) and not both_numbered and norm_contract(a["so_hd"]) != norm_contract(b["so_hd"] or "x"):
                    add("Cần kiểm tra", b, "Nghi trùng HĐ",
                        f"Cùng KH với dòng {a['row']} và cùng {'giá trị' if same_val else 'nội dung'} "
                        f"({fmt_ym(a['nam'] or a['block'], a['thang'])} ↔ {fmt_ym(b['nam'] or b['block'], b['thang'])})",
                        "Xác nhận là 2 HĐ khác nhau hay nhập trùng")
    # SĐT / email giống nhau giữa các KH khác nhau
    for field, label in (("sdt", "SĐT"), ("email", "Email")):
        by = defaultdict(set)
        for r in rows:
            v = re.sub(r"\D", "", r[field]) if field == "sdt" else r[field].lower()
            if v and len(v) >= 8:
                by[v].add(r["kh_key"])
        for v, ks in by.items():
            if len(ks) > 1:
                for r in rows:
                    rv = re.sub(r"\D", "", r[field]) if field == "sdt" else r[field].lower()
                    if rv == v:
                        add("Lỗi", r, f"{label} dùng cho nhiều KH", f"{label} '{r[field]}' xuất hiện ở {len(ks)} KH khác nhau",
                            "Có thể copy nhầm từ dòng khác")
    # STT
    by_block = defaultdict(list)
    for r in rows:
        if isinstance(r["stt"], (int, float)) and not r["row_merged"]:
            by_block[r["block"]].append(r)
    for y, rs in by_block.items():
        c = Counter(int(r["stt"]) for r in rs)
        for r in rs:
            if c[int(r["stt"])] > 1:
                add("Thông tin", r, "STT trùng", f"STT {int(r['stt'])} lặp lại trong khối năm {y}")
    # cột lệch
    for y in sorted({r["block"] for r in rows if r["ghi_chu_goc"]}):
        issues.append(["Thông tin", "", f"Khối năm {y}", "Lệch cột",
                       "Cột 'Lĩnh vực' chứa ghi chú, lĩnh vực nằm ở cột 'Số điện thoại', email ở cột 'Email' bị lệch",
                       "Báo cáo đã tự đọc đúng; nên sửa lại cột trong file gốc"])
    old = sorted({r["block"] for r in rows if r["block"] and r["block"] <= 2020 and r["email"]})
    if old:
        issues.append(["Thông tin", "", f"Khối năm {old[0]}–{old[-1]}", "Lệch cột",
                       "Email/đầu mối nằm ở cột 'Số điện thoại'; 2011–2013: cột Email/Đầu mối GMPc chứa 'ngày quyết toán' và 'ngày đạt GCN'",
                       "Báo cáo đã tự tách; nên chuẩn hoá cột trong file gốc"])
    return issues


# ------------------------------------------------------------------ scans
STOPTOK = {"hd", "hdtv", "hdkt", "hdtc", "gmpc", "thiet", "ke", "dao", "tao", "phan", "mem", "thi", "cong", "giam", "sat",
           "tu", "van", "hop", "dong", "pdf", "oks", "goi", "va", "the", "hs", "gmp", "ban", "my", "pham", "duoc",
           "vien", "lieu", "truong", "cd", "mr", "ms", "tv", "qlda", "gop", "y", "mb", "dvtv", "kho", "hn", "so", "lan",
           "nha", "may", "cty", "cp", "tnhh", "viet", "nam", "quoc", "te", "danh", "gia", "khao", "sat", "thuc", "hanh",
           "can", "tho", "tay", "nguyen", "dn", "ca", "nhan", "cn", "hstv", "dntt", "hd01", "tvgs", "hdtvgs", "hdtvtk", "ql", "da", "hd1", "hd2", "02", "01"}


def scan_tokens(fname):
    base = os.path.splitext(fname)[0]
    base = re.sub(r"^\d{4,8}\s*", "", base)
    toks = re.findall(r"[a-z0-9]+", sa(base).lower())
    return [t for t in toks if len(t) >= 3 and t not in STOPTOK and not t.isdigit()]


def scan_date(fname, folder):
    m = re.match(r"(\d{4})(\d{2})?(\d{2})?", fname)
    if not m:
        return (int(folder) if folder.isdigit() else None), None, None
    y = int(m.group(1))
    mo = int(m.group(2)) if m.group(2) and 1 <= int(m.group(2)) <= 12 else None
    d = int(m.group(3)) if m.group(3) and 1 <= int(m.group(3)) <= 31 else None
    return y, mo, d


# Tên viết tắt / viết khác trong tên file scan (regex trên tên file bỏ dấu) -> cụm có trong tên KH / số HĐ
SCAN_ALIAS = [
    (r"hapharco|harphaco", "thiet bi y te ha noi"), (r"viet tho", "donavet"), (r"vinpesco", "thuoc sat trung"),
    (r"vien duoc lieu", "vien duoc lieu"), (r"armephaco", "amephaco"), (r"nchn", "nghien cuu hat nhan"),
    (r"happysecret", "happy secret"), (r"pham gia", "pham gia"), (r"vien hh", "huyet hoc"),
    (r"gia dinh", "familiar"), (r"\bwe\b", "world essence"), (r"\bjk\b|\bik\b", "jk"), (r"dapharco", "da nang"),
    (r"beaufist", "beaufirst"), (r"\bcp\b", "c.p"), (r"nam duoc", "nam duoc"),
]
# File scan không khớp dòng nào - lý do đã biết
SCAN_NOT_GMPC = {
    ("2022", "20220126 DBC-AK .pdf"): "HĐ giữa Dabaco và Anh Khang - GMPc không ký",
    ("2025", "20251010 HĐTV Vincos.pdf"): "Bản scan thứ 2 của HĐ 1010/HDTV/Vincos-GMPc/2025 (trùng)",
    ("2020", "20200617 HDTV KIMTHANH.pdf"): "Cùng HĐ 1706/HDTV/KIMTHANH-GMPc/2020 với file 2023\\20230617 (đã khớp dòng Kim Thành). "
                                            "HĐ ký 17/06/2020 nhưng Excel đang ghi 06/2023",
}


def match_scans(rows, scans, extract):
    by_no = defaultdict(list)
    for r in rows:
        if r["so_hd"]:
            by_no[norm_contract(r["so_hd"])].append(r)
    cand = []
    for si, (folder, f) in enumerate(scans):
        if (folder, f) in SCAN_NOT_GMPC or folder == "OKS":
            continue
        ex = extract.get((folder, f))
        if ex and ex.get("so_hd"):
            for r in by_no.get(norm_contract(ex["so_hd"]), []):
                cand.append((10.0, si, r["row"], "Số HĐ (đọc từ scan)"))
        y, mo, d = scan_date(f, folder)
        toks = scan_tokens(f)
        fname = sa(f).lower()
        aliases = [target for pat, target in SCAN_ALIAS if re.search(pat, fname)]
        fy = int(folder) if folder.isdigit() else (y if y and y > 2000 else None)
        for r in rows:
            ry = r["nam"] or r["block"]
            if not ry or (fy and abs(ry - fy) > 1):   # thư mục không theo năm (vd "HĐ cá nhân") -> không lọc năm
                continue
            hay = sa(r["cong_ty"] + " " + r["so_hd"]).lower()
            hit = [t for t in toks if re.search(rf"\b{re.escape(t)}\b", hay)] + [a for a in aliases if a in hay]
            m = re.match(r"\s*(\d{2})(\d{2})", r["so_hd"])
            ddmm = bool(m and d and mo and int(m.group(1)) == d and int(m.group(2)) == mo and ry == y)
            same_day = bool(d and r["ngay"] == d and r["thang"] == mo and r["nam"] == y)
            if not fy and len(hit) < 2:   # thư mục không theo năm (HĐ cá nhân): cần khớp >= 2 từ trong tên
                continue
            if not hit and (toks or aliases or not (ddmm or same_day)):
                continue   # tên file có tên KH mà không khớp -> không dùng ngày để khớp
            sc = (2.0 if hit else 0.5) + 0.5 * len(hit)
            if y and mo and r["nam"] and r["thang"]:
                sc -= min(abs((y * 12 + mo) - (r["nam"] * 12 + r["thang"])), 24) * 0.1
            if y and mo and r["nam"] == y and r["thang"] == mo:
                sc += 2
                if d and r["ngay"] == d:
                    sc += 1
            m = re.match(r"\s*(\d{2})(\d{2})", r["so_hd"])
            if m and d and mo and int(m.group(1)) == d and int(m.group(2)) == mo:
                sc += 3
            if ry == fy:
                sc += 0.5
            cand.append((sc, si, r["row"], "Tên file + thời gian"))
    cand.sort(key=lambda x: -x[0])
    used_s, used_r, res = set(), set(), {}
    for sc, si, row, how in cand:
        if si in used_s or row in used_r:
            continue
        used_s.add(si)
        used_r.add(row)
        res[row] = (si, how, sc)
    return res


# ------------------------------------------------------------------ write helpers (format GMPc)
COMPANY = "Công ty cổ phần GMPc Việt Nam"
REPORT = "THỐNG KÊ KH ĐÃ KÝ HĐ 2011–2026"


def font(bold=False, size=10, color=None, italic=False):
    return Font(name=FONT_NAME, bold=bold, size=size, color=color, italic=italic)


def title_band(ws, row, text, ncols, height=22):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(ncols, 1))
    c = ws.cell(row=row, column=1, value=text)
    c.font = font(bold=True, size=14, color="FFFFFF")
    c.fill = FILL_NAVY
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = height


def header(ws, row, names, widths=None, start_col=1):
    for j, h in enumerate(names, start_col):
        c = ws.cell(row=row, column=j, value=h)
        c.font = font(bold=True, color=NAVY)
        c.fill = FILL_TEAL
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BORDER
    if widths:
        for j, w in enumerate(widths, start_col):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[row].height = 36


def body(ws, start, data, money=(), pct=(), wrap=(), fills=None, center=()):
    for i, r in enumerate(data, start):
        for j, v in enumerate(r):
            c = ws.cell(row=i, column=j + 1, value=None if v == "" else v)
            c.font = font()
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=j in wrap,
                                    horizontal="center" if j in center else None)
            if j in money:
                c.number_format = "#,##0"
            if j in pct:
                c.number_format = "0%"
            if fills and fills.get(i - start) is not None:
                c.fill = fills[i - start]


def page_setup(ws, title_rows="1:2"):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = title_rows
    ws.oddFooter.left.text = COMPANY
    ws.oddFooter.left.font = FONT_NAME
    ws.oddFooter.right.text = "Trang &P / &N"
    ws.oddFooter.right.font = FONT_NAME
    ws.page_margins.left = ws.page_margins.right = 0.4


def table_sheet(wb, name, title, cols, data, money=(), pct=(), wrap=(), freeze_col="A", fills=None, notes=(),
                center=(0,)):
    """Mẫu GMPc: dòng 1 tiêu đề navy, dòng 2 tiêu đề cột teal, dòng 3 đệm, dữ liệu từ dòng 4, ghi chú cuối bảng."""
    ws = wb.create_sheet(name)
    title_band(ws, 1, f"{REPORT} — {title}", len(cols))
    header(ws, 2, [c for c, _ in cols], [w for _, w in cols])
    ws.row_dimensions[3].height = 4
    body(ws, 4, data, money, pct, wrap, fills, center)
    ws.freeze_panes = f"{freeze_col}4"
    if data:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}{3 + len(data)}"
    for k, n in enumerate(notes):
        c = ws.cell(row=5 + len(data) + k, column=1, value=n)
        c.font = font(italic=True, size=9, color="595959")
    page_setup(ws)
    return ws


def mark_total(ws, row, ncols):
    for j in range(1, ncols + 1):
        ws.cell(row=row, column=j).font = font(bold=True, color=NAVY)
        ws.cell(row=row, column=j).fill = FILL_TOTAL


# ------------------------------------------------------------------ main
def threshold_sensitivity(groups):
    """Số KH quay lại / lượt quay lại theo các ngưỡng khoảng cách khác nhau."""
    out = []
    for gap in (12, 18, 24, 36):
        n_kh = n_ev = 0
        for rs in groups.values():
            idx = sorted(i for i, _ in (ym_index(r) for r in rs) if i is not None)
            ev = sum(1 for a, b in zip(idx, idx[1:]) if b - a >= gap)
            n_ev += ev
            n_kh += ev > 0
        out.append((gap, n_kh, n_ev))
    return out


def main():
    all_rows, year_heads = load_rows()
    for r in all_rows:
        r["lv_chuan"] = LINH_VUC_MAP.get(lv_key(r["linh_vuc"]), "")
        r["tt_chuan"] = TINH_TRANG_MAP.get(r["tinh_trang"].lower(), "")
        r["dich_vu"] = services(r["noi_dung"], r["so_hd"])
        r["khu_vuc"] = province(r["cong_ty"])
        r["year"] = r["nam"] or r["block"]
    groups, ids = assign_customers(all_rows)
    issues = checks(all_rows, year_heads, groups)
    excluded = [r for r in all_rows if r["row"] in EXCLUDE_ROWS]
    rows = [r for r in all_rows if r["row"] not in EXCLUDE_ROWS]
    for k in list(groups):
        groups[k] = [r for r in groups[k] if r["row"] not in EXCLUDE_ROWS]
        if not groups[k]:
            del groups[k]
    for r in excluded:
        r["loai_hd"] = "Không tính"
        issues.append(["Lỗi", r["row"], r["cong_ty"][:70], "Không tính vào thống kê", EXCLUDE_ROWS[r["row"]],
                       "Xoá dòng trùng trong file gốc (hoặc sửa nếu là HĐ khác)"])
    events = episodes(groups)
    sens = threshold_sensitivity(groups)

    # scan
    scans = list_scans()
    extract = load_scan_extract()
    smatch = match_scans(all_rows, scans, extract)
    row_by = {r["row"]: r for r in all_rows}
    for r in all_rows:
        if r["row"] in smatch:
            si, how, sc = smatch[r["row"]]
            r["scan"] = f"{scans[si][0]}\\{scans[si][1]}"
            r["scan_how"] = how
            ex = extract.get(scans[si])
            sy, smo, sd = scan_date(scans[si][1], scans[si][0])
            if ex:   # so sánh với dữ liệu đã đọc từ scan
                m = re.match(r"(\d{2}|__)/(\d{2}|__)/(\d{4})", ex.get("ngay_ky", ""))
                if m and m.group(2) != "__" and r["thang"] and (int(m.group(2)), int(m.group(3))) != (r["thang"], r["nam"]):
                    issues.append(["Cần kiểm tra", r["row"], r["cong_ty"][:70], "Lệch ngày với scan",
                                   f"Excel {fmt_ym(r['nam'], r['thang'])}, HĐ scan ghi {ex['ngay_ky']}", "Sửa theo HĐ"])
                g = ex.get("gia_tri")
                if g and r["gia_tri"] and abs(g - r["gia_tri"]) > 1000:
                    issues.append(["Cần kiểm tra", r["row"], r["cong_ty"][:70], "Lệch giá trị với scan",
                                   f"Excel {r['gia_tri']:,.0f}, HĐ scan {g:,.0f} ({ex.get('gia_tri_note', '')})", "Sửa theo HĐ"])
                if ex.get("so_hd") and r["so_hd"] and norm_contract(ex["so_hd"]) != norm_contract(r["so_hd"]):
                    issues.append(["Cần kiểm tra", r["row"], r["cong_ty"][:70], "Lệch số HĐ với scan",
                                   f"Excel '{r['so_hd']}', HĐ scan '{ex['so_hd']}'", "Sửa theo HĐ"])
            elif sy and smo and r["nam"] and r["thang"] and (sy, smo) != (r["nam"], r["thang"]) and sy > 2000:
                issues.append(["Cần kiểm tra", r["row"], r["cong_ty"][:70], "Lệch tháng với tên file scan",
                               f"Excel {fmt_ym(r['nam'], r['thang'])}, file scan '{scans[si][1]}'",
                               "Kiểm tra ngày trên HĐ (có thể đảo ngày/tháng hoặc tên file sai)"])
        else:
            r["scan"], r["scan_how"] = "", ""
    matched_s = {si for si, _, _ in smatch.values()}
    unmatched_scans = [s for i, s in enumerate(scans) if i not in matched_s]

    # ---- thống kê theo năm
    years = sorted({r["year"] for r in rows if r["year"]}, reverse=True)
    by_year = []
    for y in years:
        rs = [r for r in rows if r["year"] == y]
        kh = {r["kh_key"] for r in rs if r["kh_key"]}
        new = {r["kh_key"] for r in rs if r.get("loai_hd") == "HĐ đầu tiên"}
        back = {r["kh_key"] for r in rs if r.get("loai_hd") == "KH quay lại"}
        more = [r for r in rs if r.get("loai_hd") == "Ký thêm (cùng dự án)"]
        by_year.append([y, len(rs), len(kh), len(new), len(back), len(more),
                        sum(1 for r in rs if not r["gia_tri"]), sum(1 for r in rs if r["scan"])])
    tot = ["Tổng", sum(x[1] for x in by_year), len({r["kh_key"] for r in rows if r["kh_key"]}),
           sum(x[3] for x in by_year), sum(x[4] for x in by_year), sum(x[5] for x in by_year),
           sum(x[6] for x in by_year), sum(x[7] for x in by_year)]

    # ---- khách hàng
    cust = []
    for k, rs in groups.items():
        rs_sorted = sorted(rs, key=lambda r: (ym_index(r)[0] or 0, r["row"]))
        last = rs_sorted[-1]
        first = rs_sorted[0]
        names = [r["cong_ty"] for r in rs_sorted]
        name = names[-1]
        yrs = sorted({r["year"] for r in rs if r["year"]})
        n_ep = max(r["dot"] for r in rs)
        lvs = Counter(r["lv_chuan"] or r["linh_vuc"] for r in rs if (r["lv_chuan"] or r["linh_vuc"]))
        svc = []
        for r in rs_sorted:
            for s in r["dich_vu"]:
                if s not in svc and s != "Khác":
                    svc.append(s)

        def latest(field):
            for r in reversed(rs_sorted):
                if r[field]:
                    return r[field]
            return ""
        li, est = ym_index(last)
        months_since = (TODAY.year * 12 + TODAY.month - 1) - li if li is not None else None
        now_idx = TODAY.year * 12 + TODAY.month - 1
        ongoing = any(r["tt_chuan"] == "Đang thực hiện" and (ym_index(r)[0] or 0) >= now_idx - 36 for r in rs)
        has_contact = bool(latest("sdt") or latest("email"))
        # gợi ý chăm sóc (quy tắc đơn giản, xem sheet Hướng dẫn)
        if months_since is not None and (months_since < 12 or ongoing):
            seg = "1. Đang/ vừa làm dự án"
            tip = "Chăm sóc trong dự án; hỏi nhu cầu gói tiếp theo"
        elif months_since is not None and months_since < 48:
            seg = "2. Ưu tiên liên hệ (1–4 năm)"
            if "Đào tạo & hồ sơ GMP" in svc or "Đánh giá / tái đánh giá" in svc:
                tip = "Đã làm hồ sơ/đánh giá GMP: hỏi lịch tái đánh giá định kỳ, đào tạo lại, tự thanh tra"
            elif "Quy hoạch & thiết kế" in svc or "Lập dự án / BC đầu tư" in svc:
                tip = "Đã làm thiết kế/BCĐT: hỏi tiến độ xây dựng, chào thi công - giám sát - hồ sơ GMP"
            else:
                tip = "Hỏi thăm, giới thiệu dịch vụ mới"
        else:
            seg = "3. KH cũ (> 4 năm)"
            tip = "Xác minh còn hoạt động, cập nhật đầu mối; giới thiệu năng lực mới"
        if not has_contact:
            tip += " | Thiếu SĐT/email"
        cust.append({
            "id": ids[k], "name": name, "aliases": "; ".join(sorted(set(names) - {name}))[:300],
            "kv": latest("khu_vuc"), "lv": ", ".join(x for x, _ in lvs.most_common()),
            "first": fmt_ym(first["year"], first["thang"]), "last": fmt_ym(last["year"], last["thang"]),
            "first_y": first["year"], "last_y": last["year"], "n_years": len(yrs), "years": ", ".join(map(str, yrs)),
            "n": len(rs), "ep": n_ep, "back": n_ep - 1, "more": len(rs) - n_ep, "last_idx": li or 0,
            "svc": ", ".join(svc), "last_content": last["noi_dung"][:200], "ms": months_since if months_since is not None else "",
            "status": last["tt_chuan"] or last["tinh_trang"], "dm": latest("dau_moi"), "sdt": latest("sdt"),
            "email": latest("email"), "gmpc": latest("gmpc"), "seg": seg, "tip": tip, "key": k,
            "note": ALIAS_NOTE.get(k, ""), "group": GROUP_NOTE.get(k, ""),
        })
    cust.sort(key=lambda c: (-c["n"], -c["last_idx"]))

    # =============================================================== workbook
    wb = openpyxl.Workbook()
    n_kh = len(cust)
    n_back_kh = sum(1 for c in cust if c["back"] > 0)
    n_multi = sum(1 for c in cust if c["n"] > 1)
    n_noval = sum(1 for r in rows if not r["gia_tri"])

    # ---- Hướng dẫn (theo mẫu sheet 'Hướng dẫn' của biểu mẫu GMPc: cột B, dải tiêu đề navy)
    ws = wb.active
    ws.title = "Hướng dẫn"
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 125
    c = ws.cell(row=2, column=2, value=f"{REPORT} & DỮ LIỆU CHĂM SÓC KHÁCH HÀNG")
    c.font = font(bold=True, size=14, color="FFFFFF")
    c.fill = FILL_NAVY
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[2].height = 24
    c = ws.cell(row=3, column=2, value=f"{COMPANY} — Ngày lập: {TODAY:%d/%m/%Y} — Dữ liệu đến {DATA_TO[1]:02d}/{DATA_TO[0]} "
                                      "(năm 2026 chưa trọn năm)")
    c.font = font(size=10, color="555555")
    sections = [
        ("Nguồn dữ liệu", [
            f"• Nguồn chính: sheet 'Tổng hợp' của file {NAS_SRC} (bản copy, không sửa file gốc).",
            f"• Hợp đồng scan: {SCAN_ROOT} ({len(scans)} file PDF). Dùng tên file để đánh dấu 'đã có scan'; riêng 135 HĐ "
            "2021–2026 đã đọc nội dung (2026-09-27) được đối chiếu thêm số HĐ / ngày / giá trị.",
        ]),
        ("Định nghĩa", [
            "• Khách hàng (KH): gộp theo tên đã chuẩn hoá (bỏ dấu, bỏ 'Công ty/CP/TNHH…', bỏ phần trong ngoặc) + bảng gộp "
            "đã xác nhận (sheet 'Nhóm KH'). Công ty con cùng tập đoàn liên hệ riêng → tính là KH riêng, ghi ở cột 'Nhóm / tập đoàn'.",
            f"• Dự án (đợt): các HĐ của cùng KH ký cách nhau < {GAP_QUAY_LAI_THANG} tháng được coi là cùng 1 dự án (ký thêm gói).",
            f"• KH quay lại: HĐ ký cách HĐ trước đó của KH ≥ {GAP_QUAY_LAI_THANG} tháng (dự án cũ đã kết thúc).",
            "• Loại HĐ: 'HĐ đầu tiên' | 'Ký thêm (cùng dự án)' | 'KH quay lại'. Tái ký = Ký thêm + Quay lại.",
            "• Thiếu tháng ký → tính như giữa năm (tháng 7); khoảng cách khi đó là ước tính (ghi ở sheet 'KH quay lại').",
            "• Nhóm dịch vụ: phân loại tự động theo từ khoá trong 'Nội dung' (1 HĐ có thể thuộc nhiều nhóm) - mang tính tham khảo.",
            "• Khu vực: tỉnh/thành ghi trong tên KH (theo địa giới cũ như dữ liệu gốc); trống nếu tên không ghi.",
        ]),
        ("Giá trị hợp đồng", [
            f"• KHÔNG tổng hợp giá trị HĐ: {n_noval}/{len(rows)} HĐ chưa có giá trị (chủ yếu 2011–2015), một số HĐ ghi giá đã gồm VAT "
            "hoặc lệch so với HĐ scan → tổng sẽ không chính xác.",
            "• Giá trị từng HĐ vẫn giữ ở sheet 'Danh sách HĐ' (theo Excel gốc) để tra cứu. Sheet 'Theo năm' có cột 'Số HĐ thiếu giá trị' "
            "để theo dõi mức độ đầy đủ của dữ liệu.",
        ]),
        ("Nhóm chăm sóc (gợi ý, theo thời gian từ HĐ gần nhất)", [
            "1. Đang / vừa làm dự án: HĐ gần nhất < 12 tháng hoặc còn HĐ đang thực hiện (trong 3 năm gần đây).",
            "2. Ưu tiên liên hệ: HĐ gần nhất 1–4 năm - giai đoạn KH thường cần tái đánh giá, đào tạo lại, mở rộng dự án.",
            "3. KH cũ (> 4 năm): xác minh còn hoạt động và cập nhật đầu mối trước khi chào dịch vụ.",
            "Lưu ý: chu kỳ tái đánh giá GMP/CGMP thực tế phụ thuộc lĩnh vực và ngày cấp GCN của từng KH - cần bộ phận chuyên môn xác nhận.",
        ]),
        ("Các sheet", [
            "• Tổng quan - số liệu cho profile công ty • Theo năm • Theo lĩnh vực • Theo khu vực",
            "• Khách hàng - 1 dòng / KH, dùng để lập kế hoạch chăm sóc • KH quay lại • Danh sách HĐ - 1 dòng / HĐ",
            "• Lỗi cần sửa - các điểm cần sửa trong file gốc • Scan / Scan chưa khớp • Nhóm KH • Danh mục",
        ]),
    ]
    r = 5
    for head, items in sections:
        c = ws.cell(row=r, column=2, value=head)
        c.font = font(bold=True, size=11, color=NAVY)
        r += 1
        for t in items:
            c = ws.cell(row=r, column=2, value=t)
            c.font = font()
            c.alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        r += 1
    c = ws.cell(row=r, column=2, value="Màu")
    c.font = font(bold=True, size=11, color=NAVY)
    for t, f in (("Đỏ nhạt = Lỗi cần sửa trong file gốc", FILL_ERR), ("Vàng = Cần kiểm tra / xác nhận", FILL_WARN),
                 ("Xanh dương nhạt = Thông tin / cần bổ sung", FILL_INFO), ("Xanh lá nhạt = KH quay lại", FILL_BACK),
                 ("Xám = Không tính vào thống kê", FILL_OUT)):
        r += 1
        c = ws.cell(row=r, column=2, value=t)
        c.font = font()
        c.fill = f
    ws.sheet_view.showGridLines = False
    page_setup(ws, title_rows=None)

    # ---- Tổng quan
    ws = wb.create_sheet("Tổng quan")
    title_band(ws, 1, f"{REPORT} — TỔNG QUAN", 3)
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16

    def block(r, label, head, data):
        c = ws.cell(row=r, column=1, value=label)
        c.font = font(bold=True, size=11, color=NAVY)
        header(ws, r + 1, head)
        ws.row_dimensions[r + 1].height = 20
        body(ws, r + 2, data, center=(1, 2))
        return r + 2 + len(data) + 1

    kpi = [
        ("Giai đoạn", f"{min(years)} – {DATA_TO[1]:02d}/{DATA_TO[0]}"),
        ("Tổng số hợp đồng", len(rows)),
        ("Tổng số khách hàng", n_kh),
        ("KH ký từ 2 HĐ trở lên", f"{n_multi} ({n_multi / n_kh:.0%})"),
        (f"KH quay lại (HĐ mới sau ≥ {GAP_QUAY_LAI_THANG} tháng)", f"{n_back_kh} ({n_back_kh / n_kh:.0%})"),
        ("Số lượt quay lại", len(events)),
        ("Số HĐ ký thêm trong cùng dự án", sum(1 for r in rows if r.get("loai_hd") == "Ký thêm (cùng dự án)")),
        ("Số tỉnh/thành có KH (theo tên KH)", len({c["kv"] for c in cust if c["kv"]})),
    ]
    nr = block(3, "Chỉ số chính", ["Chỉ số", "Giá trị"], kpi)
    lv_kh, lv_hd = Counter(), Counter()
    for k, rs in groups.items():
        for lv in {r["lv_chuan"] or "(Chưa xác định)" for r in rs}:
            lv_kh[lv] += 1
    for r in rows:
        lv_hd[r["lv_chuan"] or "(Chưa xác định)"] += 1
    nr = block(nr + 1, "Theo lĩnh vực (1 KH có thể thuộc nhiều lĩnh vực)", ["Lĩnh vực", "Số KH", "Số HĐ"],
               [[k, lv_kh[k], lv_hd[k]] for k, _ in lv_hd.most_common()])
    svc_kh, svc_hd = Counter(), Counter()
    for r in rows:
        for s in r["dich_vu"]:
            svc_hd[s] += 1
    for k, rs in groups.items():
        for s in {s for r in rs for s in r["dich_vu"]}:
            svc_kh[s] += 1
    nr = block(nr + 1, "Theo nhóm dịch vụ (phân loại tự động theo nội dung HĐ)", ["Nhóm dịch vụ", "Số KH", "Số HĐ"],
               [[k, svc_kh[k], v] for k, v in svc_hd.most_common()])
    nr = block(nr + 1, "Độ nhạy định nghĩa 'KH quay lại' (khoảng cách tối thiểu giữa 2 HĐ) - tham khảo",
               ["Ngưỡng (tháng)", "Số KH quay lại", "Số lượt quay lại"],
               [[f"≥ {g}" + (" (đang dùng)" if g == GAP_QUAY_LAI_THANG else ""), k, e] for g, k, e in sens])
    c = ws.cell(row=nr + 1, column=1, value="Giá trị HĐ không tổng hợp do dữ liệu còn thiếu - xem sheet Hướng dẫn.")
    c.font = font(italic=True, size=9, color="595959")
    ws.sheet_view.showGridLines = False
    page_setup(ws)
    ws.page_setup.orientation = "portrait"

    # ---- Theo năm
    cols = [("Năm", 8), ("Số HĐ", 9), ("Số KH", 9), ("KH mới (HĐ đầu tiên)", 12), ("KH quay lại", 11),
            ("HĐ ký thêm (cùng dự án)", 13), ("Số HĐ thiếu giá trị", 11), ("Số HĐ đã có scan", 11)]
    ws = table_sheet(wb, "Theo năm", "THEO NĂM", cols, by_year + [tot], center=tuple(range(8)),
                     notes=["Số KH trong năm = số KH khác nhau có HĐ trong năm; dòng Tổng = số KH khác nhau toàn giai đoạn.",
                            f"Năm 2026 tính đến {DATA_TO[1]:02d}/{DATA_TO[0]}. Giá trị HĐ không tổng hợp (dữ liệu còn thiếu)."])
    last_row = 3 + len(by_year) + 1
    mark_total(ws, last_row, len(cols))
    ws.auto_filter.ref = None
    hc = len(cols) + 6   # bảng phụ (năm tăng dần) cho biểu đồ - cột ẩn
    for j, h in enumerate(["Năm", "Số HĐ", "Số KH", "KH quay lại"]):
        ws.cell(row=2, column=hc + j, value=h).font = font(bold=True, color="808080")
    for i, x in enumerate(sorted(by_year), 3):
        for j, v in enumerate([str(x[0]), x[1], x[2], x[4]]):
            ws.cell(row=i, column=hc + j, value=v).font = font(color="808080")
    ch = BarChart()
    ch.type = "col"
    ch.title = "Số HĐ, số KH và KH quay lại theo năm"
    ch.add_data(Reference(ws, min_col=hc + 1, max_col=hc + 3, min_row=2, max_row=2 + len(by_year)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=hc, min_row=3, max_row=2 + len(by_year)))
    for sr, col in zip(ch.series, (NAVY, TEAL, "F4B183")):
        sr.graphicalProperties.solidFill = col
        sr.graphicalProperties.line.solidFill = col
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.x_axis.delete = False   # openpyxl mặc định ẩn trục trên Excel mới
    ch.y_axis.delete = False
    ch.height, ch.width = 9, 24
    ch.visible_cells_only = False    # vẫn vẽ dù cột dữ liệu bị ẩn
    for j in range(4):
        ws.column_dimensions[get_column_letter(hc + j)].hidden = True
    ws.add_chart(ch, f"A{last_row + 5}")
    ws.page_setup.fitToHeight = 1

    # ---- Theo lĩnh vực
    lv_names = [k for k, _ in lv_hd.most_common()]
    data = []
    for y in years:
        c = Counter(r["lv_chuan"] or "(Chưa xác định)" for r in rows if r["year"] == y)
        data.append([y] + [c.get(n, 0) for n in lv_names] + [sum(c.values())])
    data.append(["Tổng"] + [lv_hd[n] for n in lv_names] + [len(rows)])
    ws = table_sheet(wb, "Theo lĩnh vực", "SỐ HĐ THEO NĂM VÀ LĨNH VỰC", [("Năm", 8)] + [(n, 12) for n in lv_names] + [("Tổng", 9)],
                     data, center=tuple(range(len(lv_names) + 2)),
                     notes=["Lĩnh vực đã chuẩn hoá - bảng quy đổi ở sheet 'Danh mục'."])
    mark_total(ws, 3 + len(data), len(lv_names) + 2)
    ws.auto_filter.ref = None

    # ---- Theo khu vực
    kv = defaultdict(lambda: [0, 0])
    for c in cust:
        kv[c["kv"] or "(Không ghi trong tên)"][0] += 1
    kv_of = {c["key"]: c["kv"] for c in cust}
    for r in rows:
        kv[kv_of.get(r["kh_key"]) or "(Không ghi trong tên)"][1] += 1
    data = sorted([[k, v[0], v[1]] for k, v in kv.items()], key=lambda x: (x[0].startswith("("), -x[1]))
    table_sheet(wb, "Theo khu vực", "THEO KHU VỰC", [("Tỉnh / thành", 26), ("Số KH", 10), ("Số HĐ", 10)], data,
                center=(1, 2), notes=["Khu vực lấy từ tên tỉnh/thành ghi trong tên KH (địa giới trước sáp nhập 2025)."])

    # ---- Khách hàng
    cols = [("Mã KH", 8), ("Khách hàng (tên gần nhất)", 42), ("Nhóm / tập đoàn", 14), ("Khu vực", 12), ("Lĩnh vực", 18),
            ("HĐ đầu tiên", 10), ("HĐ gần nhất", 10), ("Số năm có HĐ", 8), ("Các năm có HĐ", 22), ("Số HĐ", 7),
            ("Số dự án (đợt)", 8), ("Số lần quay lại", 8), ("Số HĐ ký thêm", 8), ("Dịch vụ đã dùng", 34),
            ("Nội dung HĐ gần nhất", 45), ("Số tháng từ HĐ gần nhất", 9), ("Tình trạng HĐ gần nhất", 13),
            ("Đầu mối KH", 18), ("SĐT", 14), ("Email", 26), ("Đầu mối GMPc", 12), ("Nhóm chăm sóc (gợi ý)", 20),
            ("Gợi ý liên hệ", 48), ("Tên khác trong dữ liệu", 40), ("Ghi chú gộp", 30)]
    data = [[c["id"], c["name"], c["group"], c["kv"], c["lv"], c["first"], c["last"], c["n_years"], c["years"], c["n"],
             c["ep"], c["back"], c["more"], c["svc"], c["last_content"], c["ms"], c["status"], c["dm"], c["sdt"],
             c["email"], c["gmpc"], c["seg"], c["tip"], c["aliases"], c["note"]] for c in cust]
    fl = {i: FILL_BACK for i, c in enumerate(cust) if c["back"] > 0}
    table_sheet(wb, "Khách hàng", "DANH SÁCH KHÁCH HÀNG (1 DÒNG / KH)", cols, data, wrap=(1, 13, 14, 22, 23), fills=fl,
                freeze_col="C", center=(0, 5, 6, 7, 9, 10, 11, 12, 15),
                notes=["Xanh lá nhạt = KH có quay lại. Nhóm chăm sóc là gợi ý theo thời gian từ HĐ gần nhất (xem Hướng dẫn)."])

    # ---- KH quay lại
    data = []
    name_of = {c["key"]: (c["id"], c["name"]) for c in cust}
    for k, a, b, gap, est in sorted(events, key=lambda e: (-(e[2]["year"] or 0), e[2]["thang"] or 0)):
        cid, cname = name_of[k]
        data.append([cid, cname, b["dot"], fmt_ym(a["year"], a["thang"]), a["noi_dung"][:150],
                     fmt_ym(b["year"], b["thang"]), b["so_hd"], b["noi_dung"][:150],
                     round(gap / 12, 1), "Ước tính (thiếu tháng)" if est else "", b["row"]])
    table_sheet(wb, "KH quay lại", "CÁC LƯỢT KHÁCH HÀNG QUAY LẠI",
                [("Mã KH", 8), ("Khách hàng", 40), ("Đợt thứ", 7), ("HĐ trước - thời điểm", 11), ("HĐ trước - nội dung", 42),
                 ("HĐ quay lại - thời điểm", 11), ("Số HĐ", 26), ("HĐ quay lại - nội dung", 42), ("Khoảng cách (năm)", 10),
                 ("Ghi chú", 14), ("Dòng Excel gốc", 9)],
                data, wrap=(1, 4, 7), center=(0, 2, 3, 5, 8, 10),
                notes=[f"Mỗi dòng = 1 lượt KH quay lại (HĐ mới cách HĐ trước ≥ {GAP_QUAY_LAI_THANG} tháng)."])

    # ---- Danh sách HĐ
    issue_rows = defaultdict(list)
    for lv, row, _, kind, desc, _ in issues:
        if isinstance(row, int) and row in row_by:
            issue_rows[row].append((lv, kind))
    cols = [("Năm", 7), ("Tháng ký", 9), ("Số hợp đồng", 28), ("Mã KH", 8), ("Khách hàng", 40), ("Khu vực", 11),
            ("Lĩnh vực (gốc)", 15), ("Lĩnh vực (chuẩn hoá)", 15), ("Nội dung", 55), ("Nhóm dịch vụ", 24),
            ("Giá trị trước VAT theo Excel gốc (VNĐ)", 17), ("Tình trạng", 13), ("Loại HĐ", 16), ("Đợt", 6),
            ("Cách HĐ trước (tháng)", 9), ("Đầu mối KH", 16), ("SĐT", 13), ("Email", 24), ("Đầu mối GMPc", 11),
            ("Ngày quyết toán (2011-13)", 11), ("Ngày đạt GCN (2011-13)", 13), ("File scan", 36),
            ("Điểm cần kiểm tra", 40), ("Dòng Excel gốc", 8)]
    srt = sorted(all_rows, key=lambda r: (-(r["year"] or 0), r["thang"] or 13, r["row"]))
    data, fl = [], {}
    for i, r in enumerate(srt):
        iss = issue_rows.get(r["row"], [])
        data.append([r["year"] or "", fmt_ym(r["nam"], r["thang"]) if r["thang"] else "", r["so_hd"], r["kh_id"],
                     r["cong_ty"], r["khu_vuc"], r["linh_vuc"], r["lv_chuan"], r["noi_dung"], ", ".join(r["dich_vu"]),
                     round(r["gia_tri"]) if r["gia_tri"] else "", r["tt_chuan"] or r["tinh_trang"], r.get("loai_hd", ""),
                     r.get("dot", ""), r.get("gap", ""), r["dau_moi"], r["sdt"], r["email"], r["gmpc"],
                     r["quyet_toan"], r["gcn"], r["scan"], "; ".join(sorted({k for _, k in iss})), r["row"]])
        if r["row"] in EXCLUDE_ROWS:
            fl[i] = FILL_OUT
        elif any(l == "Lỗi" for l, _ in iss):
            fl[i] = FILL_ERR
        elif r.get("loai_hd") == "KH quay lại":
            fl[i] = FILL_BACK
    table_sheet(wb, "Danh sách HĐ", "DANH SÁCH HỢP ĐỒNG (1 DÒNG / HĐ)", cols, data, money=(10,), wrap=(2, 4, 8, 22),
                fills=fl, freeze_col="F", center=(0, 1, 3, 13, 14, 23),
                notes=["Đỏ nhạt = dòng có lỗi; xanh lá nhạt = HĐ của KH quay lại; xám = không tính thống kê (trùng)."])

    # ---- Lỗi cần sửa
    order = {"Lỗi": 0, "Cần kiểm tra": 1, "Cần bổ sung": 2, "Thông tin": 3}
    issues.sort(key=lambda x: (order[x[0]], x[3], x[1] if isinstance(x[1], int) else 0))
    fl = {i: {"Lỗi": FILL_ERR, "Cần kiểm tra": FILL_WARN}.get(x[0], FILL_INFO) for i, x in enumerate(issues)}
    table_sheet(wb, "Lỗi cần sửa", "CÁC ĐIỂM CẦN SỬA TRONG FILE GỐC",
                [("Mức độ", 12), ("Dòng Excel gốc", 9), ("Khách hàng", 40), ("Loại lỗi", 24), ("Mô tả", 70), ("Đề xuất", 40)],
                issues, wrap=(2, 4, 5), fills=fl, center=(0, 1),
                notes=["Dòng = số dòng trong sheet 'Tổng hợp' của file gốc. Lỗi = nên sửa; Cần kiểm tra = cần đối chiếu HĐ; "
                       "Cần bổ sung = thiếu thông tin; Thông tin = lưu ý định dạng."])

    # ---- Scan
    data = []
    for r in sorted(all_rows, key=lambda r: (-(r["year"] or 0), r["row"])):
        data.append([r["year"] or "", r["row"], r["cong_ty"][:80], r["so_hd"], "Có" if r["scan"] else "Chưa có",
                     r["scan"], r["scan_how"]])
    fl = {i: FILL_WARN for i, d in enumerate(data) if d[4] == "Chưa có"}
    table_sheet(wb, "Scan", "ĐỐI CHIẾU HĐ SCAN", [("Năm", 7), ("Dòng Excel gốc", 9), ("Khách hàng", 45), ("Số HĐ", 28),
                                                 ("Scan", 9), ("File scan", 45), ("Cách khớp", 20)],
                data, wrap=(2,), fills=fl, center=(0, 1, 4),
                notes=[f"Tổng {len(scans)} file scan; khớp {len(smatch)}; chưa khớp {len(unmatched_scans)} (xem sheet 'Scan chưa khớp'); "
                       f"dòng Excel chưa có scan: {sum(1 for r in all_rows if not r['scan'])} (tô vàng).",
                       "Khớp tự động theo số HĐ đã đọc từ scan 2021–2026, hoặc theo tên file (tên KH + ngày)."])
    data = [[folder, f, "HĐ do OKS ký - không thuộc danh sách GMPc" if folder == "OKS"
             else SCAN_NOT_GMPC.get((folder, f), "Chưa có dòng tương ứng trong Excel, hoặc tên file khác tên KH - cần kiểm tra")]
            for folder, f in unmatched_scans]
    table_sheet(wb, "Scan chưa khớp", "FILE SCAN CHƯA KHỚP DÒNG EXCEL", [("Thư mục", 9), ("Tên file", 50), ("Ghi chú", 60)],
                data, notes=["File scan trên NAS chưa khớp được với dòng nào trong sheet 'Tổng hợp'."])

    # ---- Nhóm KH
    data = []
    for c in sorted(cust, key=lambda c: c["id"]):
        variants = sorted({r["cong_ty"] for r in groups[c["key"]]})
        if len(variants) > 1 or c["note"]:
            data.append([c["id"], c["name"], len(variants), "\n".join(variants), c["note"]])
    for a, b, why in NOT_MERGED:
        ca = next((c for c in cust if c["key"] == a), None)
        cb = next((c for c in cust if c["key"] == b), None)
        if ca and cb:
            data.append([f"{ca['id']} / {cb['id']}", f"{ca['name']}  |  {cb['name']}", "", "KHÔNG GỘP - khác KH", why])
    fl = {i: FILL_WARN for i, d in enumerate(data) if "KHÔNG GỘP" in str(d[3])}
    table_sheet(wb, "Nhóm KH", "GỘP TÊN KHÁCH HÀNG", [("Mã KH", 12), ("Tên dùng trong báo cáo", 45), ("Số cách ghi tên", 9),
                                                    ("Các cách ghi tên trong Excel gốc", 70), ("Ghi chú", 50)],
                data, wrap=(1, 3, 4), fills=fl, center=(0, 2),
                notes=["Các KH được gộp từ nhiều cách ghi tên. Dòng vàng: tên giống nhau nhưng KHÔNG gộp (đã xác nhận là KH khác nhau)."])

    # ---- Danh mục
    ws = wb.create_sheet("Danh mục")
    title_band(ws, 1, f"{REPORT} — DANH MỤC", 2)
    header(ws, 2, ["Lĩnh vực (gốc, viết thường)", "Lĩnh vực (chuẩn hoá)"], [36, 90])
    body(ws, 3, sorted(LINH_VUC_MAP.items()))
    s = len(LINH_VUC_MAP) + 5
    header(ws, s, ["Nhóm dịch vụ", "Từ khoá (không dấu) trong Nội dung"])
    body(ws, s + 1, [[n, p.replace("|", ", ")] for n, p in SERVICE_RULES], wrap=(1,))
    page_setup(ws)

    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, f"{TODAY:%y%m%d} Thong ke KH va CSKH 2011-2026.xlsx")
    wb.save(out)
    print("Saved:", out)
    print(f"HĐ {len(rows)} | KH {n_kh} | KH ≥2 HĐ {n_multi} | KH quay lại {n_back_kh} | lượt quay lại {len(events)}")
    print(f"Scan {len(scans)} | khớp {len(smatch)} | scan chưa khớp {len(unmatched_scans)} | HĐ chưa có scan {sum(1 for r in rows if not r['scan'])}")
    print("Issues:", Counter(x[0] for x in issues), Counter(x[3] for x in issues).most_common())


if __name__ == "__main__":
    main()
