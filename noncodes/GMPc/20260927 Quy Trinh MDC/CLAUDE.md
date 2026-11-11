# Quy Trinh GMPc

## Purpose
Draft the procedures (SOPs) for the MD&C department (Market Development & Communications; customers and partners) of Công ty Cổ phần GMPc Việt Nam. The SOPs must follow the company's document format. Deliverables are Word documents.

Migrated from the claude.ai project "Quy trình GMPc" ("Soạn quy trình của phòng MDC (phòng khách hàng, đối tác)"), created 2026-08-06, last updated 2026-09-06.

## Stack
- Reference docs are Markdown. Final deliverables are `.docx`.
- Python (recommended, listed in the root Tools table) is used for helper scripts, for example re-extracting from `source/` or generating the `.docx` SOPs.

## Structure
- `docs/`: reference documents as Markdown, extracted from the claude.ai export.
  - `20260702_SoDoToChuc_Phong_MDC_v0.2_2.md`: MD&C org chart and role duties (v0.2, 07/2026)
  - `260623_GMPc-CN-01_CHỨC NĂNG NHIỆM VỤ PHÒNG.md`: department functions and duties, GMPc-CN-01 V1.0, effective 23/06/2026
  - `GMPc-QT-01.BM01_Biểu mẫu trắng SOP.md`: blank SOP template, GMPc-QT-01.BM01. **Use this template for every new SOP.**
- `source/019fd74d-…json`: raw claude.ai project export, kept verbatim. Don't edit it.

## Commands
No build step. To re-extract the docs from the raw export (in WSL/Linux, use `python3` instead of `python`):
```
python -c "import json,os;d=json.load(open('source/019fd74d-9a72-70ab-827a-36db2d510d33.json',encoding='utf-8'));[open(os.path.join('docs',os.path.splitext(x['filename'])[0]+'.md'),'w',encoding='utf-8',newline='').write(x['content']) for x in d['docs']]"
```

## Conventions and constraints
- These are internal company documents that name staff and roles. Don't publish or share them outside the company.
- Document codes follow the pattern `GMPc-<type>-<nn>` (CN = chức năng, QT = quy trình, BM = biểu mẫu, TC = ?).
- Every SOP needs the drafter / reviewer / approver block and these sections: Mục đích, Phạm vi, Trách nhiệm, Định nghĩa…, following the BM01 template.

## Open questions and next steps
- The original `.docx` files are **not** in this folder. The export only had their text, so formatting, logos and layout are gone. Get the originals from the GMPc NAS (copy into `data/raw/` per the group rules, and record the source path here) before producing final Word output.
- The chat history of the claude.ai project was not in the export.
- Confirm what the `TC` document type means (GMPc-TC-001 is referenced).
- Next step: list the MD&C SOPs to write, based on GMPc-CN-01 §6.1 and the org chart.

## Tools

### Windows tools

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
<!-- launcher:tools:windows:end -->

## Change log
- 2026-09-27: Migrated from the claude.ai Projects export `C:\Users\Khiem Nguyen\Desktop\projects-000.zip` (extracted to `C:\Users\KHIEMN~1\AppData\Local\Temp\claude-migrate-20260927-154941\projects\019fd74d-9a72-70ab-827a-36db2d510d33.json`). Copied the raw JSON to `source/` and extracted its 3 docs to `docs/`. Other projects in the export were not migrated: Chinese Study, How to use Claude, an unnamed project, and Hợp đồng lao động. Verified that the SHA-256 hashes match and the docs round-trip exactly. The user chose "move", so the source zip was deleted afterwards.
- 2026-09-27: Moved from `C:\Users\Khiem Nguyen\Documents\projects\noncodes\20260927 Quy Trinh GMPc` into the GMPc group and renamed it `20260927 Quy Trinh MDC`. Copied with robocopy and changed no content. The NAS note in Open questions now refers to the group rules instead of `L:` only. Python was raised to recommended and moved to the root Tools table (removed from this project's table), since it is useful across projects. Fixed the re-extract command (`newline=''`) so it no longer writes CRLF on Windows; it now reproduces `docs/` byte for byte. Added Python to the root Linux tools table (`python3`), and noted in Commands that WSL/Linux uses `python3`.
