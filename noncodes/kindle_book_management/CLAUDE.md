# kindle_book_management

Created 2026-10-30 by the launcher. Fill in each section during the first session.

## Purpose and goals
Convert .epub novels (or folders of .epub/.docx) into one Word .docx built on `template/Template.docx`. Saved for reuse from `kndiy-tools`.

## Tech stack and key dependencies
Groovy 4.0.23 + Java 21, Apache POI 5.2.5 (jars from `~/.m2`; commons-compress 1.26.2 needed, 1.23.0 breaks docx writing). **Not Python** (the original converter is Groovy).

## Structure and key files
- `src/services/novelParsingService/` (NovelProcessor, NovelParserEpub, NovelWriteToDocx, ...) and `src/services/kndiyLibraries/DataStructure.groovy`: copied unchanged from `~/projects/codes/kndiy/kndiy-tools`.
- `template/Template.docx`: copied from `/mnt/d/Documents/Kindle Books/RAW Chinese Books/`.
- `run_epub_to_docx.sh` / `.groovy`: CLI runner.
- `korean_book/<Author_Title>/`: per-book working folders. `Saenal_Overgeared/` holds `WN KR Overgeared.docx` and `.epub` (copied, not moved, from `/mnt/d/Documents/Kindle Books/WN KR Overgeared/`; originals still there) and `cover_no_book5_pursuit.png` (cover with "PURSUIT OF POWER" and "BOOK 5" inpainted out with OpenCV; slight smudges remain).
- `.claude/skills/epub-to-docx/SKILL.md`: project skill.

## Commands (build / run / test / lint)
`./run_epub_to_docx.sh in.epub out.docx "Book" "Author" ["Chapter "] [template] ["skip|skip"]`. No tests/lint. Verified 2026-10-30 on WN KR Overgeared (2059 chapters, 0 faulty).

## Conventions, constraints, decisions
- Copies are independent of kndiy-tools; fixes must be made in both if wanted.
- Skill is project-level, so it only loads in sessions started here.

## Open questions and next steps
- Port to Python if the owner wants it (not done).
- Promo lines removed only on exact match.

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

<!-- launcher:tools:linux:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
<!-- launcher:tools:linux:end -->

## Change log

Older entries: `zClaudeDocs/CLAUDE_CHANGELOG.md`.

- 2026-10-30: Added `korean_book/Saenal_Overgeared/` (Overgeared docx, epub, edited cover).
- 2026-10-30: Fixed `epub-to-docx` SKILL.md frontmatter (stray `# kindle_book_management` line replaced by opening `---`) so the description loads.
- 2026-10-30: Copied Groovy epub converter + Template.docx from kndiy-tools, added runner script and `epub-to-docx` skill.
