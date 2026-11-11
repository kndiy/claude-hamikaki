"""Trích thư viện hạng mục từ bộ mẫu "2025 Mau chao gia" (chỉ đọc mẫu, không sửa).

Kết quả:
  catalog/catalog.json  - dữ liệu cho make_quote.py
  catalog/DANH_MUC.md   - danh mục dễ đọc để chọn mẫu / mục khi lập phiếu yêu cầu

Chạy từ thư mục project:  python scripts/build_catalog.py
"""
import json
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = ROOT.parent / "2025 Mau chao gia"
OUT_DIR = ROOT / "catalog"

SECTORS = {
    "1. DUOC PHAM": ("duoc_pham", "Dược phẩm (GMP)"),
    "2. MY PHAM": ("my_pham", "Mỹ phẩm (CGMP-ASEAN)"),
    "3. THUC PHAM BVSK": ("tpbvsk", "Thực phẩm bảo vệ sức khỏe (HS-GMP)"),
    "4. THUOC THU Y": ("thu_y", "Thuốc thú y (GMP)"),
    "5. KHO GSP DUOC PHAM": ("kho_gsp", "Kho GSP / GDP dược phẩm"),
    "6. DU DKSX MP": ("ddksx_mp", "Đủ điều kiện sản xuất mỹ phẩm"),
    "7. TAI DANH GIA": ("tai_danh_gia", "Tái đánh giá"),
    "8. LAP DU AN DAU TU": ("du_an_dau_tu", "Lập dự án đầu tư"),
}
# Files that are not reusable templates (customer examples, calculation sheets).
SKIP_FILES = {"Song ngu.xlsx"}
SKIP_PREFIX = ("~$", "3. Bảng tính")

ROMAN = re.compile(r"^(I|II|III|IV|V|VI|VII|VIII|IX|X)$")
# Wording that must not go into a quote unreviewed (could read as payments to officials).
BLOCKED = re.compile(r"ngoại giao|thúc đẩy hồ sơ|(chi phí|bao gồm)[^.]*tiếp đoàn", re.I)


def slug(text):
    text = unicodedata.normalize("NFD", text.replace("đ", "d").replace("Đ", "D"))
    text = "".join(c for c in text if unicodedata.category(c) != "Mn").lower()
    text = re.sub(r"^(2025|2024)\s+", "", text.strip())
    return re.sub(r"[^a-z0-9]+", "_", text).strip("_")


def clean(v):
    return None if v is None else str(v).strip()


def parse_sheet(ws):
    """Return the scope table, notes and closing text of one sheet."""
    rows = list(ws.iter_rows(min_row=1, max_col=5, values_only=True))
    head = next(i for i, r in enumerate(rows) if clean(r[0]) and re.match(r"^(STT|Item)", clean(r[0])))
    sections, i = [], head + 1
    while i < len(rows):
        a, c, d, e = clean(rows[i][0]), clean(rows[i][2]), rows[i][3], rows[i][4]
        if a and re.match(r"^(Tổng|Total)", a):
            break
        if a and ROMAN.match(a):
            sections.append({"code": a, "title": c, "price": e if isinstance(e, (int, float)) else None, "items": []})
        elif c:
            if not sections:  # Some templates start with an item row and no roman header.
                sections.append({"code": "I", "title": None, "price": None, "items": []})
            item = {"code": a, "text": c}
            if isinstance(e, (int, float)):
                item["price"] = e
            sections[-1]["items"].append(item)
        i += 1
    i += 3  # totals: subtotal, VAT, total with VAT
    footnotes, notes, closing = [], [], []
    for r in rows[i:]:
        a, c = clean(r[0]), clean(r[2])
        if a and a.startswith("*"):
            notes.append(c)
        elif not notes and (a or c):
            footnotes.append(a or c)
        elif notes and a:
            closing.append(a)
    keys = ["vat", "time", "payment"]
    return {
        "intro": clean(rows[head - 1][0]),
        "sections": sections,
        "footnotes": footnotes,
        "notes": dict(zip(keys, notes)),
        "closing": closing,
    }


def norm(s):
    return re.sub(r"\s+", " ", s).strip() if s else s


def candidates(vn, en, warn):
    """Suggest an English text for each Vietnamese string, only where the two sheets line up.

    The ENG sheets are not line-by-line translations (extra lines, different numbering),
    so a pair is suggested only when a section has the same number of lines on both sheets.
    The final English lives in catalog/song_ngu.json, which is reviewed by hand.
    """
    out = {}
    if en is None:
        return out
    if len(vn["sections"]) != len(en["sections"]):
        warn.append(f"số mục VN={len(vn['sections'])} khác ENG={len(en['sections'])}")
    else:
        for sv, se in zip(vn["sections"], en["sections"]):
            out[norm(sv["title"])] = norm(se["title"])
            if len(sv["items"]) != len(se["items"]):
                warn.append(f"mục {sv['code']}: {len(sv['items'])} dòng VN, {len(se['items'])} dòng ENG")
                continue
            for iv, ie in zip(sv["items"], se["items"]):
                out[norm(iv["text"])] = norm(ie["text"])
    for k in vn["notes"]:
        out[norm(vn["notes"][k])] = norm(en["notes"].get(k))
    out[norm(vn["intro"])] = norm(en["intro"])
    for a, b in zip(vn["closing"], en["closing"]):
        out[norm(a)] = norm(b)
    return {k: v for k, v in out.items() if k and v}


def main():
    catalog, suggest = {}, {}
    md = ["# Danh mục mẫu chào giá", "", f"Nguồn: `{TEMPLATE_DIR}` (chỉ đọc).", ""]
    for folder, (sector_id, sector_name) in SECTORS.items():
        md += [f"## {sector_name} (`{sector_id}`)", ""]
        for f in sorted((TEMPLATE_DIR / folder).rglob("*.xlsx")):
            if f.name in SKIP_FILES or f.name.startswith(SKIP_PREFIX):
                continue
            wb = openpyxl.load_workbook(f)
            vn_ws = next(ws for ws in wb.worksheets if "ENG" not in ws.title.upper())
            en_ws = next((ws for ws in wb.worksheets if "ENG" in ws.title.upper()), None)
            warn = []
            data = parse_sheet(vn_ws)
            en_data = parse_sheet(en_ws) if en_ws else None
            for vn_text, en_text in candidates(data, en_data, warn).items():
                suggest.setdefault(vn_text, {}).setdefault(en_text, []).append(f"{sector_id}/{slug(f.stem)}")
            name = f.parent.name if folder == "8. LAP DU AN DAU TU" else f.stem
            tid = f"{sector_id}/{slug(name)}"
            for s in data["sections"]:
                if BLOCKED.search(s["title"] or ""):
                    s["blocked"] = True
                for it in s["items"]:
                    if BLOCKED.search(it["text"]):
                        it["blocked"] = True
            for k, fn in enumerate(data["footnotes"]):
                if BLOCKED.search(fn):
                    warn.append(f"ghi chú cuối có câu cần sửa: \"{fn[:60]}…\"")
            catalog[tid] = {
                "sector": sector_id,
                "sector_name": sector_name,
                "source": str(f.relative_to(TEMPLATE_DIR)),
                "sheet_vn": vn_ws.title,
                "sheet_en": en_ws.title if en_ws else None,
                "warnings": warn,
                **data,
            }
            md.append(f"### `{tid}`  ")
            md.append(f"Nguồn: `{f.relative_to(TEMPLATE_DIR)}` · ENG: {'có' if en_ws else '**không**'}")
            md.append("")
            for s in data["sections"]:
                if s.get("blocked"):
                    md.append(f"- ⛔ **{s['code']}. {s['title']}** — *không dùng được (rủi ro pháp lý)*")
                    continue
                md.append(f"- **{s['code']}. {s['title']}** ({len(s['items'])} dòng)")
                for it in s["items"]:
                    if it.get("blocked"):
                        md.append(f"  - ⛔ `{it['code']}` {it['text'][:80]} — *bị loại mặc định*")
            for w in warn:
                md.append(f"- ⚠️ {w}")
            md.append("")
    OUT_DIR.mkdir(exist_ok=True)
    (OUT_DIR / "catalog.json").write_text(json.dumps(catalog, ensure_ascii=False, indent=1), encoding="utf-8")
    (OUT_DIR / "DANH_MUC.md").write_text("\n".join(md), encoding="utf-8")
    # Every Vietnamese string a quote can contain, with English suggestions from the ENG sheets.
    strings = []
    for t in catalog.values():
        for s in t["sections"]:
            strings += [s["title"]] + [it["text"] for it in s["items"]]
        strings += [t["intro"], *t["notes"].values(), *t["footnotes"], *t["closing"]]
    strings = list(dict.fromkeys(norm(s) for s in strings if s))
    song_ngu_path = OUT_DIR / "song_ngu.json"
    song_ngu = json.loads(song_ngu_path.read_text(encoding="utf-8")) if song_ngu_path.exists() else {}
    todo = {s: suggest.get(s, {}) for s in strings if s not in song_ngu}
    (OUT_DIR / "en_goi_y.json").write_text(json.dumps(todo, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(catalog)} mẫu -> {OUT_DIR}")
    print(f"{len(strings)} câu VN; {len(strings) - len(todo)} đã có EN trong song_ngu.json; {len(todo)} còn thiếu -> en_goi_y.json")
    for tid, t in catalog.items():
        for w in t["warnings"]:
            if "cần sửa" in w:
                print(f"  ! {tid}: {w}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
