# Claude-file mirror

This repo backs up the Claude-related files of the projects tree. The sub-project repos exclude those files, so they're kept here instead. The project folders stay the source of truth: Claude Code reads and writes the files there, never here.

There are two mirror repos. They use **the same scripts and this same README**. Only `mirror.conf` differs between them.

| Local folder (in the projects root) | GitHub repo | Scope (`mirror.conf`) |
|---|---|---|
| `_claude-rnd4u` | `git@github.com:rnd4u/claude-rnd4u.git` | `codes/pricefx` (work) |
| `_claude-hamikaki` | `git@github.com:kndiy/claude-hamikaki.git` | everything else: `codes/` and `noncodes/` without `codes/pricefx`, plus `_launcher/`, the root `CLAUDE.md` and the launch scripts (family: **ha**, **mi**nh **ka**, **ki**em) |

- `sync` copies the project files into the repo and sends them to GitHub on a review branch.
- `restore` copies the reviewed state (`origin/main`) back into the project folders.

## What gets mirrored

`mirror.conf` has these sections. Each item goes on its own line, paths are relative to the projects root, and `#` starts a comment:

| Section | Meaning |
|---|---|
| `[claude]` | Folders whose **Claude files** are mirrored: `CLAUDE.md`, `CLAUDE.local.md`, `.claude-project`, and anything in `.claude/` or `zClaudeDocs/`. |
| `[scripts]` | Helper-script extensions (e.g. `py ps1 sh groovy`) also mirrored inside `[claude]` folders, but **only outside nested git repos**. Code in a git repo is already backed up by that repo. |
| `[whole]` | Folders where every text file is mirrored (e.g. `_launcher`). |
| `[files]` | Single files at the projects root (e.g. `CLAUDE.md`, `Launch.cmd`). |
| `[exclude]` | Folders or single files skipped entirely (e.g. `codes/pricefx` in hamikaki, raw customer data files in rnd4u). |

Fixed rules, the same in every mirror:

- Text extensions only, up to 1 MB per file.
- docx, xlsx, csv, images, pdf, zip, certificates and data dumps with no extension are never copied, and neither are `settings.local.json` files.
- Folders that are always skipped: `.git`, `node_modules`, `target`, `build`, `dist`, `__pycache__`, `.venv` and `venv`.
- `.sh` and `.command` files are committed as executable.
- Mirrored paths are the same as under the projects root.

## Sync (projects → GitHub, for review)

- Windows: `powershell -ExecutionPolicy Bypass -File <mirror>\sync.ps1`
- Linux / macOS: `bash <mirror>/sync.sh`

Options: `-Remote <url>`, `-DryRun`, `-NoPush`, `-Message`, `-AllowSecrets`. `sync.sh` takes the same options in `--kebab-case`.

Each run:

1. Fetches `origin` and creates a new branch `sync/<yyyyMMdd-HHmmss>`.
   - It starts from `origin/main`.
   - If the newest `sync/*` branch is still open (not merged, and not deleted on GitHub), it stacks the new branch on top instead. Edits that exist only on an unmerged branch, such as script changes, are never dropped.
2. Deletes the scope folders in the mirror, copies the files in again, and scans them for secrets. It stops if it finds any.
3. Commits and pushes **only that branch**, then prints a GitHub compare link. You review it there and merge the pull request yourself. Nothing is ever pushed to `main` directly.

Other behaviour:

- Deleted or renamed source files drop out of the mirror too.
- If nothing changed since the newest `sync/*` branch, no new branch is made.
- Each new branch contains everything not yet on `main`, so older local `sync/*` branches are deleted. On GitHub, merge only the newest PR and close the older ones.
- After merging, delete the branch on GitHub, or turn on "Automatically delete head branches". To reject a sync, close its PR **and delete its branch**. Otherwise the next sync stacks on it.

## Restore (GitHub `origin/main` → projects)

- Windows: `powershell -ExecutionPolicy Bypass -File <mirror>\restore.ps1 -DryRun`, then run it again without `-DryRun`.
- Linux / macOS: `bash <mirror>/restore.sh --dry-run`, then without `--dry-run`.

How it handles each file in `origin/main` (the mirror's own files, meaning the scripts, this README and `mirror.conf`, are skipped):

| Local file | Action |
|---|---|
| missing | created, but only if its folder exists (see below) |
| identical | left alone |
| differs, but matches an older version in `origin/main`'s history | updated (the local copy was just outdated) |
| differs, and matches no version on `main` (unsynced or still-unmerged edits) | **conflict: skipped**, exit code 3 |

- `-Force` / `--force` overwrites conflicts and saves the local copies first in `.restore-backup/<timestamp>/` (ignored by git).
- **Folders that don't exist yet are skipped**, because `git clone` refuses a non-empty folder. On a new machine, clone the project repos first, then restore. `-CreateMissing` / `--create-missing` overrides this.
- Restore never deletes local files. Files removed from the mirror are only reported, and only for paths inside this mirror's scope.
- Recommended order on each machine: **restore before you start work, sync when you finish.**

## New machine

1. Copy or set up the projects root, then clone the project repos you need.
2. Clone each mirror with long paths on, because some paths are longer than 260 characters on Windows:
   - `git clone -c core.longpaths=true git@github.com:rnd4u/claude-rnd4u.git _claude-rnd4u`
   - `git clone -c core.longpaths=true git@github.com:kndiy/claude-hamikaki.git _claude-hamikaki`
3. In each mirror, run restore with `-DryRun` first, then for real. Restore `_claude-hamikaki` first: it brings back the launcher and the root `CLAUDE.md`.

## Notes

- Usually Claude runs these scripts when you ask from any session ("sync the Claude files", "restore from claude-hamikaki"). The instructions for that are in the backup section of the root `CLAUDE.md`, which every session under the root loads.
- **Changing a script:** edit it in one mirror, then copy it unchanged to the other. Each `.ps1` script must match its `.sh` twin, and the `.ps1` files are ASCII-only, for Windows PowerShell 5.1.
- On Windows, the `.ps1` scripts set `core.sshCommand` for the repo to `C:/Windows/System32/OpenSSH/ssh.exe`. The GitHub key has a passphrase and is unlocked in the Windows ssh-agent, which Git for Windows' bundled ssh can't reach. That key (`~/.ssh/kndiy_ed25519`, account `khiemnd-rnd4u`) is the company key and is used for company work, which now lives almost entirely under `codes/pricefx` (including `rnd4u/claude-rnd4u`). `kndiy/claude-hamikaki` is private/family use and uses its own key, `~/.ssh/hamikaki_ed25519` (account `kndiy`). On Linux/WSL the hamikaki repo gets a per-repo `core.sshCommand` that selects that key (`ssh -i ~/.ssh/hamikaki_ed25519 -o IdentitiesOnly=yes`), because the plain `github.com` host in `~/.ssh/config` offers the company key.
- Some mirrored paths are longer than 260 characters. The Windows scripts use `\\?\` paths and set `core.longpaths`.
- The mirror folders hold copies of `CLAUDE.md` files. When searching the projects tree, ignore `_claude-*/` so the copies don't show up as duplicates.
- Tested (2026-11-11) on Windows: the scripts in a scratch copy of the tree, covering create, folder skip, update, conflict kept and force with backup, plus the `.sh` twins under Git Bash. Not yet tested on real Linux or macOS.
