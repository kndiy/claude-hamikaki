"""Duyệt thư viện song ngữ bằng Excel.

    python scripts/song_ngu_xlsx.py xuat   -> catalog/Duyet song ngu.xlsx
    python scripts/song_ngu_xlsx.py nhap   -> ghi lại catalog/song_ngu.json từ file Excel

Trong Excel: sửa cột "EN" nếu cần, đặt cột "Trạng thái" = da_duyet khi đã đọc xong.
Không sửa cột "VN" (là khóa tra cứu, phải khớp đúng câu trong mẫu).
"""
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
JSON = ROOT / "catalog" / "song_ngu.json"
XLSX = ROOT / "catalog" / "Duyet song ngu.xlsx"
CATALOG = json.loads((ROOT / "catalog" / "catalog.json").read_text(encoding="utf-8"))
HEAD = ["#", "VN", "EN", "Trạng thái", "Ghi chú", "Dùng trong mẫu"]


def used_in():
    import re
    norm = lambda s: re.sub(r"\s+", " ", s).strip() if s else s
    where = {}
    for tid, t in CATALOG.items():
        texts = [t["intro"], *t["notes"].values(), *t["footnotes"], *t["closing"]]
        for s in t["sections"]:
            texts += [s["title"]] + [it["text"] for it in s["items"]]
        for x in texts:
            if x:
                where.setdefault(norm(x), set()).add(tid)
    return where


def export():
    data = json.loads(JSON.read_text(encoding="utf-8"))
    where = used_in()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Song ngu"
    ws.append(HEAD)
    for i, (vn, v) in enumerate(data.items(), 1):
        ws.append([i, vn, v["en"], v.get("trang_thai", "cho_duyet"), v.get("ghi_chu", ""), ", ".join(sorted(where.get(vn, [])))])
    for col, w in zip("ABCDEF", (5, 70, 70, 12, 40, 40)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows():
        for c in row:
            c.font = Font(name="Calibri Light", b=c.row == 1)
            c.alignment = Alignment(wrap_text=True, vertical="top")
        if row[0].row > 1 and row[4].value:
            row[4].fill = PatternFill("solid", fgColor="FFFF00")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(XLSX)
    print(f"{len(data)} dòng -> {XLSX}")


def import_():
    data = json.loads(JSON.read_text(encoding="utf-8"))
    ws = openpyxl.load_workbook(XLSX).active
    changed = 0
    for _, vn, en, status, note, _ in ws.iter_rows(min_row=2, values_only=True):
        if vn not in data:
            print(f"  ! bỏ qua (cột VN đã bị sửa?): {str(vn)[:60]}")
            continue
        new = {"en": en, "trang_thai": status or "cho_duyet"}
        if note:
            new["ghi_chu"] = note
        if new != data[vn]:
            data[vn], changed = new, changed + 1
    JSON.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    done = sum(v["trang_thai"] == "da_duyet" for v in data.values())
    print(f"{changed} dòng thay đổi; đã duyệt {done}/{len(data)}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"xuat": export, "nhap": import_}[sys.argv[1]]()
