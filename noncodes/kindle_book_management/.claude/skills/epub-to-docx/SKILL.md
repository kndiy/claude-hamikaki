---
name: epub-to-docx
description: Convert an .epub novel (or a folder of .epub/.docx files) into a single Word .docx using the saved Template.docx. Use when the user asks to convert an epub to word/docx in the kindle_book_management project.
---

# epub-to-docx

Groovy converter (copied from `~/projects/codes/kndiy/kndiy-tools`, `services/novelParsingService`). Not Python.

## Run

```bash
cd ~/projects/noncodes/kindle_book_management
./run_epub_to_docx.sh <in.epub|folder> <out.docx> "<Book name>" "<Author>" ["Chapter "] [template] ["skip text 1|skip text 2"]
```

- Needs Java 21 and the jars in `~/.m2` (see the script). No Maven build.
- Defaults: chapter prefix `Chapter `, template `template/Template.docx` (SimSun Normal; Bookerly is set per run; indentation comes from template styles).

## How to apply

1. Ask for the epub path, output path, book name and author if not given. Don't guess.
2. Run the script; check the final line `N Fautly Chapters: [...]` (EmptyBody / SkippedChapter) and report it.
3. Epub parsing relies on `h2.chapter-title` pages; epubs with a different structure need `NovelParserEpub.groovy` changed.
4. Promo lines are removed only on exact text match; pass them as skip texts (arg 7).
