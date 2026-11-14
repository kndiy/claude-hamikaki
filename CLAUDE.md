# Global Profile

These rules apply to every project and session.

## 1. Act as an advisor, not just an assistant

- Don't just carry out requests. Judge whether they are the right thing to do.
- Before acting, point out risks, trade-offs, and better alternatives when they exist. Push back on weak ideas and explain why.
- Give a clear recommendation along with your reasoning, not a neutral list of options.
- Bring up issues I didn't ask about when they matter, such as security, maintainability, performance, cost, or scope creep.
- Be direct and honest. Don't just agree to be agreeable.

## 2. Confirm unknowns. No guessing.

- If a requirement, name, value, path, version, business rule, or intent is unclear or missing, **stop and ask** before acting on it.
- Don't invent facts, APIs, file contents, config values, or assumptions. Check them in the code, docs, or tools first. If you can't check something, ask.
- If you have to state an assumption, label it as one and get my confirmation before building on it.
- Say plainly when you don't know something or aren't sure.

## 3. Keep CLAUDE.md in every project

- **New project:** create a `CLAUDE.md` at the project root before doing any other work. Include:
  - The project's purpose and goals
  - The tech stack and key dependencies
  - The structure and key files
  - Commands to build, run, test, and lint
  - Conventions, constraints, and decisions made so far
  - Open questions and next steps
- **After any change:** update the project's `CLAUDE.md` in Markdown with the new context. Record what changed, why, any decisions made, and any new open items. Keep it current and short.
- **Existing project without a `CLAUDE.md`:** create one before continuing.

## 4. Project and data locations

- All projects live in this root folder (the folder holding this `CLAUDE.md`) on the local disk. Never use OneDrive. Currently the root is `C:\Users\Khiem Nguyen\Documents\projects` on Windows, `~/projects` inside WSL on a WSL machine, or `~/projects` on macOS (not an iCloud-synced Desktop/Documents folder).
- **Claude-file backup (since 2026-11-11):** the sub-project repos exclude Claude files, so two mirror repos in this root back them up. Both use the same scripts; each one's `mirror.conf` sets its scope.

  | Mirror folder | GitHub repo | Scope |
  |---|---|---|
  | `_claude-rnd4u/` | `git@github.com:rnd4u/claude-rnd4u.git` | `codes/pricefx` (work) |
  | `_claude-hamikaki/` | `git@github.com:kndiy/claude-hamikaki.git` | everything else: `codes/` and `noncodes/` without `codes/pricefx`, plus `_launcher/`, this root `CLAUDE.md` and the root launch scripts (family repo: Ha, Minh Kha, Khiem) |

  - **What gets copied:**
    - Claude files: `CLAUDE.md`, `.claude-project`, `.claude/` and `zClaudeDocs/`, text files only.
    - In hamikaki, also helper scripts (`.py`, `.ps1`, `.sh`, `.groovy` and similar) that sit outside nested git repos.
    - Never copied: docx, xlsx, csv, images, `settings.local.json` and the `.lnk` shortcuts.
  - The project folders stay the source of truth. Never edit files inside `_claude-*/`, and ignore those folders when searching the tree.
  - **Claude runs these scripts when asked, from any session, including inside a sub-project.** For example: "sync/push/back up the Claude files", "update claude-rnd4u", "restore from claude-hamikaki".
    - Use absolute paths, because the session's folder is usually a sub-project.
    - Pick the mirror whose scope holds the files in question. If the user says just "the Claude files", run **both**: rnd4u, then hamikaki.
    - **Sync** (projects → a new `sync/<timestamp>` branch on GitHub, for review):
      - Windows: `powershell -NoProfile -ExecutionPolicy Bypass -File "<root>\_claude-<name>\sync.ps1"`
      - Linux/macOS: `bash <root>/_claude-<name>/sync.sh`
      - Afterwards, give the user the compare link it prints. They merge on GitHub. Nothing is pushed straight to `main`.
      - If the secret scan stops the run, show the hits. Only use `-AllowSecrets` if the user confirms they are false positives.
    - **Restore** (merged `origin/main` → project folders):
      - Always run `restore.ps1 -DryRun` (or `restore.sh --dry-run`) first and show the result.
      - Then run it for real.
      - On conflicts (local edits that aren't on `main`), stop and ask. Usually the fix is to sync first. Use `-Force` only with the user's explicit OK, and point them to the backups in `_claude-<name>/.restore-backup/`.
    - Suggest restoring at the start of work and syncing at the end.
    - On Windows, the scripts set git's `core.sshCommand` to the Windows OpenSSH client, so the passphrase key in the Windows ssh-agent works. That key (account `khiemnd-rnd4u`, a collaborator on `kndiy/claude-hamikaki`) reaches both repos. A push needs no extra setup.
    - **Changing a script:** edit it in one mirror, then copy it unchanged to the other one. Only `mirror.conf` may differ.
  - Details (options, `mirror.conf` format, the new-machine steps): `_claude-rnd4u/README.md`, the same file in both mirrors.

## 5. Projects tree

This file replaced `~/.claude/CLAUDE.md` on 2026-09-27 (backup: `~/.claude/CLAUDE.md.bak`). It applies only to sessions started inside this root.

- Structure: `root > subRoot > [group > group > ...] > project`.
  - subRoots: `codes` (software projects), `noncodes` (data, reports, documents). The list lives in `_launcher/config.json` (Windows), `_launcher/linux/config` (Linux) and `_launcher/mac/config` (macOS).
  - Groups can be nested to any depth. A folder is a **project** when it contains a `.claude-project` marker file. Every other folder under a subRoot is a **group**.
- CLAUDE.md layering: this file applies to everything under the root. A subRoot's or group's `CLAUDE.md` applies to everything below it. A project's `CLAUDE.md` applies to that project only. Put each rule at the highest level where it holds, and don't copy rules down.
- Create all new projects inside this root, either through the launcher (`_launcher\Launch.cmd` on Windows, `bash _launcher/linux/start-claude.sh` in WSL, `Launch.command` on macOS) or by hand with the same structure and marker.
- Folder names must work on Windows, Linux and macOS: none of `\ / : * ? " < > |`, no leading `.` or `_`, and no trailing `.` or space. The launcher doesn't list folders whose names start with `_` or `.`.
- **Portable:** copying the whole root to another machine is enough. On first run the launcher installs the missing tools listed in the tables below and in the project's own table. Claude login, `~/.claude` settings and auto-memory stay per machine, so put durable rules in CLAUDE.md files, not in memory.

## 6. Tools

These are the tools Claude Code relies on. Any `CLAUDE.md` in the tree (root, subRoot, group, project) can carry a Windows, Linux and/or macOS table between the markers below. The launcher checks the root tables at start and the tables along a project's path when that project is opened. Each Check runs in the folder holding that `CLAUDE.md`, and the detected version is written into Status. If a `required` or `recommended` tool is missing, the launcher offers to run its Install command from that folder.

- **When work starts depending on a new tool, add a row** at the level where it's needed (Level: `required` / `recommended` / `optional`). A Check passes when the command exists and exits with code 0. Inside a cell, escape `|` as `\|`. An Install cell that starts with `manual:` is only shown, never run.
- The launcher maintains the Status column, so don't edit it by hand.

### Windows tools

<!-- launcher:tools:windows:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| Claude Code | `claude --version` | required | `irm https://claude.ai/install.ps1 \| iex` | 2.1.322 |
| Git for Windows | `git --version` | recommended | `winget install --id Git.Git -e --source winget --accept-source-agreements --accept-package-agreements` | 2.45.1 |
| winget | `winget --version` | optional | `manual: install "App Installer" from the Microsoft Store` | 1.29.380 |
| Python | `python --version` | recommended | `winget install --id Python.Python.3.13 -e --scope user --source winget --accept-source-agreements --accept-package-agreements` | 3.13.15 |
<!-- launcher:tools:windows:end -->

### Linux (WSL Ubuntu) tools

<!-- launcher:tools:linux:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| curl | `curl --version` | required | `sudo apt-get update && sudo apt-get install -y curl` | 8.18.0 |
| Claude Code | `claude --version` | required | `curl -fsSL https://claude.ai/install.sh \| bash` | 2.1.325 |
| Git | `git --version` | recommended | `sudo apt-get update && sudo apt-get install -y git` | 2.53.0 |
| Python | `python3 --version` | recommended | `sudo apt-get update && sudo apt-get install -y python3 python3-pip python3-venv` | 3.14.4 |
| Java (JDK 21) | `java -version` | recommended | `sudo apt-get update && sudo apt-get install -y openjdk-21-jdk` | 21.0.13.0.3 |
| Maven | `mvn -version` | recommended | `sudo apt-get update && sudo apt-get install -y maven` | 3.9.12 |
<!-- launcher:tools:linux:end -->

### macOS tools

<!-- launcher:tools:macos:start -->
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| Homebrew | `brew --version` | recommended | `/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"` | - |
| Claude Code | `claude --version` | required | `curl -fsSL https://claude.ai/install.sh \| bash` | - |
| Git | `git --version` | recommended | `brew install git` | - |
| Python | `python3 --version` | recommended | `brew install python` | - |
<!-- launcher:tools:macos:end -->
