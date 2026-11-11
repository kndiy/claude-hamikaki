# kndiy-tools

Personal Groovy utility toolbox run from `Main.groovy`: sensor (LogTag) data parsing and reports, inventory (Nhap Xuat Ton) reports, file renaming, novel/text parsing, Pricefx price list validation, image helpers. Remote: https://github.com/khiemnd-rnd4u/KndiyTools.git (branch `master`; also `features/NewLogTagHandler`, `features/NewProcessedGoodsHandler`).

## Main tools (per owner, 2026-09-29)
1. **Inventory parser** (`services/inventories`): reads Nibot export data and builds Excel reports, auto-calculating out prices from inventory. A second input, `ProductMaster.xlsx`, groups SKUs and labels and holds the production norms.
2. **Sensor parser** (`services/sensorService`): reads LogTag PDFs and builds reports for GMP / food-storage inspection.
3. **Novel parser** (`services/novelParsingService`, `readTxtNovelParsingService`): builds a .docx novel from data scraped from novel sites (a hobby project). `NovelProcessor` accepts a folder or a single file of `.docx` or `.epub` (`NovelParserEpub`). Skip signs may be `null` for clean sources such as epubs. Example: `runEpubParser()` in `Main.groovy` (call commented out).

The other helpers (file renaming, Pricefx price list validation, image helpers) are secondary.

## Tech stack
- Groovy 4.0.23 (`groovy-all`), Java 21, Maven, `gmavenplus-plugin` 1.13.1
- Apache POI 5.2.5, iText 8.0.5, GroupDocs Parser 25.5 (custom repo `releases.groupdocs.com`), metadata-extractor 2.19.0, Lombok 1.18.44, Joda-Time, log4j/slf4j
- Version `3.3.0-SNAPSHOT`, groupId `com.kndiy`

## Structure
- `pom.xml`
- `src/main/groovy/Main.groovy`: entry point. `main()` calls one `runXxx()`; the others are commented out. Currently `runInventoryParser()`.
- `src/main/groovy/{entities,services,cache,chineseTextService}`: `services/` has `sensorService`, `inventories`, `pricefx`, `imageService`, `novelParsingService`, `readTxtNovelParsingService`, `kndiyLibraries`, `fasleDatabase`.
- `target/` is build output (git-ignored). `.idea/` is partly tracked.

## Commands
- Build/run: `mvn package`. **Warning:** the `gmavenplus` `execute` goal runs `Main.groovy` during the build, so it processes real files (see Constraints). Check which `runXxx()` is active before building.
- Tests / lint: none configured.

## Constraints and decisions
- `Main.groovy` lines 195-198 hard-code input/output paths on the Windows D: drive: `/mnt/d/Documents/Personal/00 Family Small/Projects/260503 Nhap Xuat Ton Ario/...`. Kept as-is on migration (decision 2026-09-29). This is not the company NAS.
- Work tree was migrated with uncommitted changes in 9 Groovy files and a staged `.idea/remote-targets.xml`. Nothing was committed.
- Never touch the company NAS (`L:\`).

- Running one service without Maven: `java -cp <groovy 4.0.23, groovy-xml, poi 5.2.5 + deps from ~/.m2>:src/main/groovy groovy.ui.GroovyMain script.groovy`. Maven only runs `Main.groovy` (no Groovy compile goal), so avoid `mvn package` for this.
- `NovelParserOdt.groovy` is fully commented out (odfdom is not in the pom). The `odt` branch in `NovelProcessor` was removed because the dangling class reference broke compilation.
- Epub run (WN KR Overgeared): template `/mnt/d/Documents/Kindle Books/RAW Chinese Books/Template.docx` (SimSun Normal, Bookerly set per run). The writer sets no indentation, so it comes from the template styles.

## Open items
- Epub promo lines are removed only on exact text match (`skippableTexts`). 13 variants like `6th Anniversary Event Link (Submission End Date: ...)` remain in the Overgeared docx.
- Move the hard-coded paths in `Main.groovy` into config or arguments.
- Decide whether to commit or discard the uncommitted changes.
- `.claude/settings.local.json` is copied from the old project. It references java-11 and `/mnt/c/Users`, so it is outdated.
- Build was not run after migration, because building executes `Main.groovy` against real data.

## Tools

### Linux tools
<!-- launcher:tools:linux:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| Java (JDK 21) | `java -version` | required | `sudo apt-get update && sudo apt-get install -y openjdk-21-jdk` | 21.0.12.1 |
| Maven | `mvn -version` | required | `sudo apt-get update && sudo apt-get install -y maven` | 3.9.12 |
| Git | `git --version` | recommended | `sudo apt-get update && sudo apt-get install -y git` | 2.53.0 |
<!-- launcher:tools:linux:end -->

## Change log
- 2026-09-29: Added "Main tools" section from owner's description.
- 2026-09-29: Migrated (copy, nothing excluded, `.git` kept) from `/home/kndiy/git/KndiyTools` to `projects/codes/kndiy/kndiy-tools`. Added `.claude-project` marker and this file. No source paths changed. Source left untouched.
- 2026-10-30: Added `NovelParserEpub` and epub/single-file/blank-skip-sign support in `NovelProcessor`; `runEpubParser()` in `Main.groovy`. Generated `WN KR Overgeared.docx` (2059 chapters, no faulty chapters). Removed the dead `odt` branch. Did not run `mvn package` or tests (none configured).
