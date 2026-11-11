"""Bước 1: Lập danh mục PDF scan trên NAS (chỉ đọc), trích lớp chữ, chọn trang cần đọc và render ảnh.

- Không ghi gì lên NAS.
- Output:
    data/processed/scan_index.csv        danh mục file + trang được chọn
    data/processed/scan_text/<ID>.txt    lớp chữ (OCR sẵn trong PDF, chất lượng kém, chỉ để dò trang/đối chiếu số)
    <RENDER_DIR>/<ID>/pN.png             ảnh trang để đọc (thư mục tạm, không thuộc project)
"""
import csv
import os
import re
import sys

import pymupdf

SCAN_ROOT = r"L:\04. Tổng hợp Data khách hàng\01. Danh sách KH đã kí HĐ\Scan hợp đồng tư vấn"
PROJECT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROCESSED = os.path.join(PROJECT, "data", "processed")
RENDER_DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(PROCESSED, "render")

AMOUNT = re.compile(r"\d{1,3}(?:[.,]\s?\d{3}){2,}")
FEE_WORDS = re.compile(r"(ph[ií]|gi[aá]\s*tr[iị]|t[oô]ng\s*c[oô]ng|VN[DĐ]|thanh\s*to[aá]n|fee|payment)", re.I)


def pick_pages(doc):
    n = len(doc)
    texts = [(doc[i].get_text() or "") for i in range(n)]
    has_text = sum(len(t.strip()) for t in texts[:3]) > 200
    pages = list(range(min(3, n)))
    if has_text:
        scored = []
        for i in range(2, min(n, 14)):
            amounts = len(AMOUNT.findall(texts[i]))
            if amounts:
                scored.append((amounts * 2 + len(FEE_WORDS.findall(texts[i])), i))
        for _, i in sorted(scored, reverse=True)[:2]:
            if i not in pages:
                pages.append(i)
    else:
        pages = list(range(min(6, n)))
    return sorted(pages), texts, has_text


def main():
    os.makedirs(os.path.join(PROCESSED, "scan_text"), exist_ok=True)
    rows = []
    folders = sorted(d for d in os.listdir(SCAN_ROOT) if os.path.isdir(os.path.join(SCAN_ROOT, d)))
    idx = 0
    for folder in folders:
        for f in sorted(os.listdir(os.path.join(SCAN_ROOT, folder))):
            if not f.lower().endswith(".pdf"):
                continue
            idx += 1
            sid = f"S{idx:03d}"
            path = os.path.join(SCAN_ROOT, folder, f)
            doc = pymupdf.open(path)
            pages, texts, has_text = pick_pages(doc)
            with open(os.path.join(PROCESSED, "scan_text", f"{sid}.txt"), "w", encoding="utf-8") as fh:
                for i, t in enumerate(texts):
                    fh.write(f"\n===== PAGE {i + 1} =====\n{t}")
            out = os.path.join(RENDER_DIR, sid)
            os.makedirs(out, exist_ok=True)
            for i in pages:
                doc[i].get_pixmap(dpi=100, colorspace=pymupdf.csGRAY).save(os.path.join(out, f"p{i + 1}.png"))
            rows.append({"id": sid, "folder": folder, "file": f, "pages": len(doc),
                         "has_text": has_text, "rendered": " ".join(str(i + 1) for i in pages)})
            doc.close()
    with open(os.path.join(PROCESSED, "scan_index.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
    print(f"{len(rows)} PDFs indexed -> {RENDER_DIR}")


if __name__ == "__main__":
    main()
