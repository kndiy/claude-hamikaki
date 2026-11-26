# _launcher

Scripts that check tools, let you pick a project under the projects root, and open Claude Code in it, on Windows, on WSL Ubuntu and on macOS. The root CLAUDE.md (sections 5 and 6) describes the tree and the tools tables. This file covers how the launcher works.

## Entry points
- `<root>\Launch.cmd` (Windows master): shows a menu to run the Windows launcher, the WSL launcher, or update the WSL copy. `-ExecutionPolicy Bypass` applies to this process only.
- `<root>/launch.sh` (Linux master): offers a `cproj` alias in `~/.bashrc`, then runs `linux/start-claude.sh --shell-after`.
- `<root>/Launch.command` (macOS master): double-click in Finder or run from Terminal. Offers a `cproj` alias in `~/.zshrc` (zsh) or `~/.bash_profile` (bash), then runs `mac/start-claude.sh --shell-after`.
- **`S) settings`** at the root, subRoot and group menus (both OSes): opens Claude Code in that folder with `--add-dir <_launcher>` and an `--append-system-prompt` that sets the scope (that level's CLAUDE.md, and the launcher when asked). When Claude exits, the same menu comes back and subRoots are reloaded from config. Project nodes don't have it, because picking a project already opens Claude there.
- **`M) migrate`** at subRoot and group menus (both OSes): asks for a source folder (it must exist and must not contain the current folder), then opens Claude Code in the current folder with `--add-dir <source>`, a rules prompt and a first message asking it to analyse the source. The rules prompt sets the steps: analyse without changing anything → summarise and ask (folder name, copy or move, excludes, paths, tools) → copy with the marker → CLAUDE.md with a tools table → verify. When Claude exits, if exactly one new project appeared, the launcher offers to open it.
- **`M)` also takes a claude.ai data export** (the .zip from Settings → Privacy → Export data, or its unzipped folder) for moving Claude Desktop chat Projects. It's detected by `projects.json` at the top or in the only subfolder. A .zip is extracted to `%TEMP%` / `$TMPDIR` (outside the tree) and deleted after the session. Claude then lists the projects and asks which to migrate, the folder names and which chats to include. It writes knowledge docs to `knowledge/`, the instructions verbatim to `knowledge/_project_instructions.md` and turned into CLAUDE.md rules, and confirmed chats to `chats/`. Finally it reports what couldn't be migrated.
- You can also run `Start-Claude.ps1 [-Root X] [-SkipToolCheck]` or `linux/start-claude.sh` or `mac/start-claude.sh [--root X] [--skip-tool-check] [--shell-after]` directly.

## Files
| File | Role |
|---|---|
| `Launch.ps1` | Master: registers a Start Menu shortcut "Claude Projects" once per machine/root (state in `%LOCALAPPDATA%\claude-projects-launcher`), then offers Windows / WSL |
| `Start-Claude.ps1` | Windows launcher: tool check, tree menu, starts `claude` |
| `Install-Tools.ps1` | `-ClaudeMd <file> [-Names ...]`: runs Install commands from that file's folder |
| `Launcher.Common.ps1` | Shared helpers (config, templates, table parse/write, `Get-ToolVersion`) |
| `config.json` | `root` (default `..`), `subRoots`, `wslDistro` (`Ubuntu`), `wslRoot` (`~/projects`) |
| `linux/start-claude.sh`, `install-tools.sh`, `lib.sh` | Bash port. Config: `linux/config` (`root=`, `subRoots=a,b`), default root `../..` |
| `mac/start-claude.sh`, `install-tools.sh`, `lib.sh` | macOS port of the Linux scripts. Config: `mac/config` (same keys), default root `../..`. Rewritten for bash 3.2 and BSD tools (see Decisions) |
| `linux/wsl-bridge.sh` | Called by `Launch.ps1` through `wsl.exe`: `status` / `deploy` / `run` |

## Decisions
- **The root is dynamic:** it's always the folder holding `_launcher` unless it was changed explicitly. A saved root that doesn't exist on this machine falls back to that default. Copying the tree to another machine needs no reconfiguration.
- **Project vs group:** a project has a `.claude-project` marker. Navigation uses a stack of folders rather than comparing path strings, because 8.3 short names broke the comparison.
- **Tools tables:** there are separate `windows` and `linux` marker blocks, so each OS only runs its own commands. A table can sit in any CLAUDE.md along the path. A Check passes when it exits 0 with no PowerShell error; native stderr output alone doesn't count as a failure. Status is written back only when it changes.
- **WSL root is separate** (Linux filesystem, `/mnt/*` refused). The user keeps WSL projects only in WSL and backs them up with restic. `deploy` copies the launcher, plus the root CLAUDE.md and `launch.sh` when they're missing. It never copies projects.
- **Start Menu shortcut, not Desktop:** the Desktop is often synced by OneDrive.
- PowerShell 5.1 quirks handled: ASCII-only `.ps1` files, UTF-8 without BOM for CLAUDE.md, `wsl.exe` calls under `ErrorActionPreference=Continue`, paths passed to WSL as arguments (5.1 mangles embedded quotes), and `Read-Line` treats end of input as quit.
- **Settings sessions:** they run in the node folder so that level's CLAUDE.md chain loads. `--add-dir` gives access to `_launcher` but doesn't load its CLAUDE.md, so the appended prompt tells Claude to read it first. The Windows prompt contains no `"` because PowerShell 5.1 mangles quotes in native arguments. Edits to launcher scripts apply on the next start, since the running launcher is already loaded.
- **Migration safety:** the default is to copy. The source is never modified. It's deleted only if the user chose move, the copy was verified, and the user confirmed again. Nothing on L: is ever changed or deleted. Claude does the copying (not the launcher) so it can ask about excludes, venvs and hard-coded paths first.
- **macOS port is a copy of the Linux scripts, not shared code,** so the tested Linux/WSL scripts stay untouched. Differences: bash 3.2 (no `mapfile`, no negative array indexes, no `set -u` because empty arrays fail under it, `local a=()` split into declare + assign); no `timeout` (uses `timeout`/`gtimeout` if present); `mktemp` always gets a template; the `macos` marker block in CLAUDE.md files; `refresh_path` also adds `/opt/homebrew/bin` and `/usr/local/bin`; root can't be in OneDrive/iCloud; no WSL `/mnt` check; `--shell-after` runs `$SHELL -l`. Keep the three ports in sync when changing behavior.
- **macOS gotchas:** `git` and `python3` on a clean Mac are shims that pop up the Command Line Tools install dialog when run, so the tool check can trigger it. Homebrew's installer asks you to add its `shellenv` line to your profile; the launcher adds the paths for its own run only. Files copied by zip lose the executable bit: `chmod +x Launch.command _launcher/mac/*.sh`.
- **Windows: Claude runs via `Diagnostics.Process`, not `& claude`,** for the settings and migrate options. Inside a function whose output is captured (`$chain = Select-Project ...`), `& claude` sends Claude's stdout into that output and breaks its UI. Testing with a stub found this. The final project launch stays `& claude` because it runs at the top level. Arguments are quoted by hand, with trailing backslashes doubled.
- **claude.ai export format is unverified.** Anthropic doesn't document it. According to a third-party tool (github.com/Brads777/ClaudeProjectExport), `projects.json` holds each project's instructions and knowledge docs as **extracted text only**: original PDFs and images aren't included. Chats aren't linked to projects, and memory isn't included in projects. The rules prompt therefore tells Claude to inspect the real structure, match chats only with the user's confirmation, list binaries to re-download by hand, and never copy the whole export (it holds every chat in the account) or `users.json`.
- **Security:** Install commands come from CLAUDE.md files and run through `Invoke-Expression` / `bash -c`. Each one is shown and needs Y before it runs. Review table edits made by others, including edits made by Claude.

## Testing
Tested on Windows (PowerShell 5.1) against scratch copies of the tree: navigation, nested groups, creating projects and groups, the `O` guard, `B` back, tool detection and write-back, a project install running in the project folder, the missing-required-tool exit, Check edge cases, the master runner (shortcut prompt, WSL-not-installed path, Windows path). The Linux scripts were tested under **Git Bash**, not real WSL: tables, install, alias registration, bridge `status`/`deploy`, `/mnt` refusal.

## Status (2026-10-02)
The Windows side is working and in use: the report project is in `noncodes\` and its tools check passes. The WSL side is written, tested under Git Bash only, and not yet deployed. The macOS port (2026-10-02) is written and syntax-checked, and its flow was tested on Linux with a stub `claude`; it has **never run on a real Mac**.

## Waiting on the user
- Review the reworded rule 4 in the root CLAUDE.md, which now says "this root folder" instead of an absolute path.
- Decide on a Python venv for the report project on Linux (see below).
- Check whether restic, started when Windows starts, can reach `\\wsl$` before the distro is running.
- Optional: put `_launcher` (or the whole root, minus data) under git.

## Open items
- **Run `Launch.command` on a real Mac** and check: bash 3.2 behaviour, BSD `awk`/`sed`/`sort -z`, the Homebrew installer, the Claude Code installer, `cproj` in zsh, the Finder double-click (Gatekeeper may block downloaded .command files: right-click > Open once).
- The macOS tools table has no Java/Maven rows (the Linux one has); add them if you need them there.
- Try `M)` with a real claude.ai export; adjust the export rules if its structure differs from the third-party description.
- **Not yet run on real WSL Ubuntu** (WSL isn't installed on this machine). Still to verify: `wsl.exe -l -q` parsing, `wslpath`, `bridge run`, apt installs with sudo, `claude` install into `~/.local/bin`.
- The report project has no Linux tools table yet. Ubuntu 24.04 blocks `pip install` system-wide (PEP 668), so a venv approach has to be decided first.
- Changing subRoots means editing `config.json` / `linux/config` by hand; there's no menu option.

## Change log

Older entries: `zClaudeDocs/CLAUDE_CHANGELOG.md`.

- 2026-10-02: Added the macOS launcher (`Launch.command`, `mac/*.sh`), a `macos` tools table in the root CLAUDE.md, and macOS mentions in the root rules and the settings prompts. Tested on Linux (bash 5) in a scratch tree with a stub `claude`: alias, tool check, new project, settings, migrate, back. Not tested on bash 3.2 or BSD tools.
- 2026-09-27: Created. Windows launcher, per-OS tools tables, path-wide project tools, dynamic root, master runners, bash port, WSL bridge.
- 2026-09-27: Migration. Copied `~/.claude/CLAUDE.md` into the root CLAUDE.md, then renamed the original to `CLAUDE.md.bak`, so it no longer loads outside the root. Moved `20260927 Thong Ke...` into `noncodes\` with a marker and a Windows tools table. Renamed the root `Projects` → `projects` for compatibility.
