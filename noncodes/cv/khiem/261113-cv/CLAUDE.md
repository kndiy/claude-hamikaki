# 261113-cv

Created 2026-11-13 by the launcher. Fill in each section during the first session.

## Purpose and goals
New CV for Khiem Nguyen (2026-11), targeting a still-unconfirmed company that trains a Claude model to predict/trade gold on trading platforms. Must include Pricefx experience (customers, projects with skimmed process, Lead Configuration Engineer on all).

## Tech stack and key dependencies
Output: .docx, A4, 8 mm margins on all sides (Khiem's requirement), Calibri 10pt. Other design left to Claude.
Built with python-docx (node is not installed in WSL); render check via `soffice --headless --convert-to pdf` + `pdftoppm`.

## Structure and key files
- Source (read-only, Windows disk): `/mnt/d/Documents/Personal/00 Family Small/K Duy Khiem/CV/231010 RnD4U/`
  - `Khiem9112_Resume.pdf` (current 1-page Pricefx CE resume: skills, Firth NZ, KP Film-Aptly, Pricefx cert 10/01/2024, education), `231009 CV.indd`, `Cover Letters/`, `To Print/` (degree scans).
- Jira-assigned projects seen (partial, rnd4u-f): Firth (Phase 5 + maintenance), Stramit Pricing Transformation, Metcash maintenance, KP Film / Aptly.

## Commands (build / run / test / lint)
TBD.

## Conventions, constraints, decisions
- Confirm customer naming with Khiem before including (confidentiality).
- "Lead Configuration Engineer" on all Pricefx projects is Khiem's instruction.
- Gold-trading target unverified; tailor only after the company is confirmed.

## Open questions and next steps
- Full customer/project list with dates and one-line process each.
- Target company, role title, output format, length, language.
- Skills-gap framing: cybersecurity, training Claude, hosting, REST API + subscription access.

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:linux:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
<!-- launcher:tools:linux:end -->

## Change log

Older entries: `zClaudeDocs/CLAUDE_CHANGELOG.md`.

- 2026-11-13: Copied source CV + degree scans to `khiem/documents/` (cv-231010/, degrees/).
- 2026-11-13: `Khiem_Nguyen_CV_draft1.docx` created. Verified content from old resume only; yellow-highlighted placeholders mark everything still to confirm (Pricefx project list/dates, previous employer, +15%/-5% figures, target role, personal email).
- 2026-11-13: Project created. Filled purpose, sources, open items after reading the old resume and cover letter.
