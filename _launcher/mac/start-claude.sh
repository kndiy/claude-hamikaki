#!/usr/bin/env bash
# Check tools, pick a project under the projects root, and open Claude Code in it (macOS).
#
# Tree: root > subRoot > [group > ...] > project. A project is a folder that
# contains a .claude-project marker; every other folder under a subRoot is a group.
# Tools come from the macOS tools table in any CLAUDE.md on the path: the root
# table is checked at start, the subRoot/group/project tables once a project is
# chosen. Detected versions are written back; missing required/recommended tools
# go to install-tools.sh.
#
# Usage: start-claude.sh [--root PATH] [--skip-tool-check] [--shell-after]
#   --root PATH        use this root for this run only (default: the folder holding _launcher)
#   --skip-tool-check  skip all tool checks
#   --shell-after      open a login shell ($SHELL) in the project folder after Claude exits
# Runs on the stock macOS bash 3.2 (no mapfile, no negative array indexes, no set -u).
set -o pipefail
LAUNCHER_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$LAUNCHER_DIR/lib.sh"

root_arg=''; skip_tools=0; shell_after=0
while [ $# -gt 0 ]; do
    case "$1" in
        --root) root_arg="${2:-}"; shift 2 ;;
        --skip-tool-check) skip_tools=1; shift ;;
        --shell-after) shell_after=1; shift ;;
        -h|--help) sed -n '2,16p' "$0"; exit 0 ;;
        *) echo "Unknown option: $1"; exit 2 ;;
    esac
done

load_config
ROOT="$(resolve_root "${root_arg:-$ROOT}")"

# ---------- tool check ----------

tool_check() {  # CLAUDE_MD...
    [ $skip_tools = 1 ] && return 0
    local m md name level install ans blocking to_install
    blocking=(); to_install=()
    test_tools "$@"
    for m in "${MISSING_TOOLS[@]}"; do
        IFS=$'\t' read -r md name level install <<< "$m"
        case "$level" in required|recommended) to_install+=("$m") ;; esac
    done
    if [ ${#to_install[@]} -gt 0 ]; then
        say "$c_yellow" $'\n''Missing tools will be installed with:'
        for m in "${to_install[@]}"; do
            IFS=$'\t' read -r md name level install <<< "$m"
            echo "  $name (in $(dirname "$md")): $install"
        done
        read -r -p 'Install now? [Y/n] ' ans || ans=n
        if [[ ! "$ans" =~ ^[Nn] ]]; then
            for m in "${to_install[@]}"; do
                IFS=$'\t' read -r md name level install <<< "$m"
                bash "$LAUNCHER_DIR/install-tools.sh" "$md" "$name"
            done
            refresh_path
            test_tools "$@"
        fi
    fi
    for m in "${MISSING_TOOLS[@]}"; do
        IFS=$'\t' read -r md name level install <<< "$m"
        [ "$level" = required ] && blocking+=("  $name (in $(dirname "$md")): $install")
    done
    if [ ${#blocking[@]} -gt 0 ]; then
        say "$c_red" $'\n''Required tools are missing. Install them, open a NEW terminal, and run the launcher again:'
        printf '%s\n' "${blocking[@]}"
        exit 1
    fi
}

# ---------- tree navigation ----------

children() {  # DIR -> CHILDREN array
    CHILDREN=()
    local d
    while IFS= read -r -d '' d; do CHILDREN+=("$d"); done < <(
        find "$1" -mindepth 1 -maxdepth 1 -type d ! -name '.*' ! -name '_*' -print0 | sort -z)
}

is_project() { [ -f "$1/$PROJECT_MARKER" ]; }

# Runs Claude Code in a tree folder, then comes back to the menu.
claude_in() {  # DIR LABEL CLAUDE_ARGS...
    local dir="$1" label="$2" keep_root="$ROOT"; shift 2
    say "$c_green" $'\n'"Opening Claude Code in $dir ($label)"
    printf '\e]0;Claude %s - %s\a' "$label" "$(basename "$dir")"
    (cd "$dir" && claude "$@")
    # The session may have changed the subRoots list; script edits apply on the next start.
    load_config; ROOT="$keep_root"
    say "$c_gray" 'Back in the launcher. Changes to launcher scripts apply the next time it starts.'
}

# Opens Claude Code in a root, subRoot or group folder to edit that level's CLAUDE.md
# and/or the launcher. The launcher folder is added with --add-dir; its CLAUDE.md
# does not load from there, hence the appended prompt.
settings_session() {  # DIR KIND
    local dir="$1" kind="$2" prompt ldir
    ldir="$(cd "$LAUNCHER_DIR/.." && pwd)"   # _launcher (this script lives in _launcher/mac)
    prompt="Settings session started from the Claude Projects launcher at the $kind folder $dir. \
Scope: this folder's CLAUDE.md (rules and tools tables for everything below it) and, when the user asks, \
the launcher in $ldir (read $ldir/CLAUDE.md before changing it, \
keep the Windows (.ps1, ASCII only), Linux (linux/*.sh) and macOS (mac/*.sh, bash 3.2 compatible) versions in sync, and update that CLAUDE.md). \
Do not change projects below this folder unless asked. When the user exits, the launcher menu comes back."
    claude_in "$dir" settings --add-dir "$ldir" --append-system-prompt "$prompt"
}

# Prints the folder holding projects.json when DIR is a claude.ai data export
# (at its top or in its only subfolder); returns 1 otherwise.
find_claude_export() {  # DIR
    local p subs; subs=()
    [ -f "$1/projects.json" ] && { printf '%s\n' "$1"; return 0; }
    while IFS= read -r -d '' p; do subs+=("$p"); done < <(find "$1" -mindepth 1 -maxdepth 1 -type d -print0)
    [ ${#subs[@]} -eq 1 ] && [ -f "${subs[0]}/projects.json" ] && { printf '%s\n' "${subs[0]}"; return 0; }
    return 1
}

# Rules for migrating a claude.ai (Claude Desktop) chat-project export. The format is
# not documented by Anthropic, so Claude is told to inspect it rather than trust this.
export_migrate_rules() {  # ORIGIN EXPORT KIND DIR TOOLS
    printf '%s' "Migration session started from the Claude Projects launcher. Source: $1, a claude.ai data export \
(Claude Desktop chat Projects), found at $2. Target: new project folder(s) inside the $3 folder $4 \
(the current folder). Expected content, unverified: projects.json (each project with its custom instructions \
and knowledge docs, the docs as extracted text only), conversations.json (every chat of the account, not linked \
to projects), memories.json, users.json. Inspect the real structure first with a small script (files can be \
large) and do not assume field names. Work in this order. \
1) Without changing anything, list the projects (name, dates, number and size of knowledge docs, whether it has \
instructions), then ask the user (AskUserQuestion) which to migrate, the folder name for each (propose names \
that meet the root naming rules), and whether to bring chats: chats can only be matched by name or topic, so \
show the candidates and let the user confirm them. Do not guess. \
2) For each chosen project: create the folder with an empty .claude-project marker file; write each knowledge \
doc as a UTF-8 file under knowledge/ with its original name; keep the custom instructions verbatim in \
knowledge/_project_instructions.md; write confirmed chats as Markdown under chats/. \
3) Write the project CLAUDE.md per root rule 3: purpose and rules taken from the custom instructions (say which \
parts need the user's review), the knowledge and chats layout, a $5, and a change log entry naming the \
export and date. \
4) Verify doc counts against projects.json and report what could not be migrated: original PDFs, images and \
other binaries (only their text is in the export; list them so the user can download them from claude.ai), \
memory, and anything else missing. \
5) Copy only the chosen projects and chats. The export holds the whole account, so never copy it wholesale, \
never copy users.json, and never modify the export. \
When the user exits, the launcher menu comes back and offers to open the new project."
}

folder_migrate_rules() {  # ORIGIN KIND DIR TOOLS
    printf '%s' "Migration session started from the Claude Projects launcher. Source: $1. \
Target: a new project folder inside the $2 folder $3 (the current folder). Work in this order. \
1) Analyse the source without changing anything: purpose, tech stack, size and file count, git state \
(uncommitted changes, remotes), existing CLAUDE.md, virtual environments, dependency and build folders, caches, \
secrets, and absolute paths or config tied to the old location (OneDrive or iCloud folders, home folders, /Volumes mounts, the NAS). \
2) Give a short summary, then ask the user (AskUserQuestion) everything you need before copying: the new folder \
name (propose one that meets the root naming rules), copy or move, what to exclude and rebuild instead, which \
hard-coded paths to change, which tools the project needs. Do not guess. \
3) Migrate: create the folder with an empty .claude-project marker file, copy with rsync -a or cp -a (keep .git), \
rebuild excluded environments if the user agreed, and change only the paths the user approved. \
4) Create or update the project CLAUDE.md per root rule 3, including a $4, and a change log entry \
naming the source path and date. \
5) Verify: compare file counts and sizes, run or build the project if that is cheap, and report differences. \
6) Never modify the source. Delete it only if the user chose move, after verification, and after asking again. \
Never change or delete anything on the company NAS. \
When the user exits, the launcher menu comes back and offers to open the new project."
}

# Asks for a folder (or a .zip, e.g. a claude.ai data export) elsewhere and opens
# Claude Code here to analyse it, ask the user, and copy it in as new project(s).
# A .zip is extracted to a temp folder outside the tree and deleted afterwards.
# Sets NEW_DIR when exactly one new project appeared and the user wants to open it;
# returns 1 otherwise.
migrate_session() {  # DIR KIND
    local dir="$1" kind="$2" src zip='' temp='' export origin tools rules first ans p b before new
    NEW_DIR=''; before=(); new=()
    read -r -p 'Project folder, or claude.ai export .zip, to migrate (blank to cancel): ' src || return 1
    src="${src%/}"; src="${src/#\~/$HOME}"
    [ -n "$src" ] || return 1
    if [ -f "$src" ] && [[ "$src" =~ \.[Zz][Ii][Pp]$ ]]; then
        zip="$(cd "$(dirname "$src")" && pwd -P)/$(basename "$src")"
        temp="$(mktemp -d "${TMPDIR:-/tmp}/claude-migrate-XXXXXX")" || return 1
        say "$c_gray" "Extracting to $temp (deleted after the session) ..."
        if command -v unzip >/dev/null; then unzip -q "$zip" -d "$temp"
        else python3 -m zipfile -e "$zip" "$temp"; fi || {
            say "$c_red" 'Could not extract the .zip.'; rm -rf "$temp"; return 1; }
        src="$temp"
    elif [ ! -d "$src" ]; then
        say "$c_red" "Folder or .zip not found: $src"; return 1
    else
        src="$(cd "$src" && pwd -P)"
        case "$(cd "$dir" && pwd -P)/" in
            "$src"/*) say "$c_red" 'That folder contains this one; pick a folder outside it.'; return 1 ;;
        esac
    fi
    children "$dir"; before=("${CHILDREN[@]}")
    origin="$src"
    [ -n "$zip" ] && origin="$src (extracted from $zip; the launcher deletes this copy after the session)"
    tools='Tools section whose macOS table sits between the lines <!-- launcher:tools:macos:start --> and <!-- launcher:tools:macos:end --> (columns Tool, Check, Level, Install, Status; Status -)'
    if export="$(find_claude_export "$src")"; then
        rules="$(export_migrate_rules "$origin" "$export" "$kind" "$dir" "$tools")"
        first="List the projects in the claude.ai export at $export and ask which to migrate here."
    else
        rules="$(folder_migrate_rules "$origin" "$kind" "$dir" "$tools")"
        first="Analyse $src and plan its migration into a new project here."
    fi
    claude_in "$dir" migrate --add-dir "$src" --append-system-prompt "$rules" "$first"
    [ -n "$temp" ] && rm -rf "$temp"
    children "$dir"
    for p in "${CHILDREN[@]}"; do
        is_project "$p" || continue
        for b in "${before[@]}"; do [ "$b" = "$p" ] && continue 2; done
        new+=("$p")
    done
    [ ${#new[@]} -eq 1 ] || return 1
    read -r -p "Open the new project '$(basename "${new[0]}")' now? [Y/n] " ans || return 1
    [[ "$ans" =~ ^[Nn] ]] && return 1
    NEW_DIR="${new[0]}"
}

new_tree_folder() {  # PARENT KIND -> NEW_DIR
    local parent="$1" kind="$2" name bad_chars='[\\/:*?"<>|]' bad_start='^[._]' bad_end='[. ]$'
    NEW_DIR=''
    read -r -p "New $kind name (blank to cancel): " name || return 1
    name="$(printf '%s' "$name" | sed 's/^[[:space:]]*//; s/[[:space:]]*$//')"
    [ -n "$name" ] || return 1
    # Same rules as the Windows launcher so the tree can move between systems.
    if [[ "$name" =~ $bad_chars || "$name" =~ $bad_start || "$name" =~ $bad_end ]]; then
        say "$c_red" 'Invalid name: no \ / : * ? " < > |, must not start with . or _, or end with . or space.'
        return 1
    fi
    if [ -e "$parent/$name" ]; then say "$c_red" "'$name' already exists."; return 1; fi
    mkdir -p "$parent/$name"
    new_folder_md "$parent/$name" "$kind"
    [ "$kind" = project ] && : > "$parent/$name/$PROJECT_MARKER"
    say "$c_green" "Created $kind '$name'"
    NEW_DIR="$parent/$name"
}

# Sets CHAIN to the folders from the subRoot down to the chosen project; returns 1 to go back.
select_project() {  # SUBROOT_DIR
    local stack dir c i crumb tag p
    stack=("$1")
    while true; do
        dir="${stack[${#stack[@]}-1]}"
        children "$dir"
        crumb=''; for p in "${stack[@]}"; do crumb="${crumb:+$crumb > }$(basename "$p")"; done
        say "$c_cyan" $'\n'"== $crumb =="
        [ ${#CHILDREN[@]} -eq 0 ] && say "$c_gray" '  (empty)'
        for i in "${!CHILDREN[@]}"; do
            if is_project "${CHILDREN[$i]}"; then tag='[project]'; else tag='[group]  '; fi
            printf '  %2d) %s %s\n' $((i + 1)) "$tag" "$(basename "${CHILDREN[$i]}")"
        done
        say "$c_cyan" '   N) new project here   G) new group here'
        say "$c_cyan" '   M) migrate a project here: folder or claude.ai export .zip (Claude analyses, asks first)'
        say "$c_cyan" '   S) settings: open Claude Code here (this CLAUDE.md, the launcher)'
        [ ${#stack[@]} -gt 1 ] && say "$c_cyan" '   O) open THIS folder as a project (marks it)'
        say "$c_cyan" '   B) back   Q) quit'
        read -r -p 'Choose: ' c || exit 0
        c="${c//[[:space:]]/}"
        if [[ "$c" =~ ^[0-9]+$ ]] && [ "$c" -ge 1 ] && [ "$c" -le ${#CHILDREN[@]} ]; then
            p="${CHILDREN[$((c - 1))]}"
            stack+=("$p")
            if is_project "$p"; then CHAIN=("${stack[@]}"); return 0; fi
            continue
        fi
        case "$c" in
            [Nn]) if new_tree_folder "$dir" project; then stack+=("$NEW_DIR"); CHAIN=("${stack[@]}"); return 0; fi ;;
            [Gg]) if new_tree_folder "$dir" group; then stack+=("$NEW_DIR"); fi ;;
            [Ss]) if [ ${#stack[@]} -gt 1 ]; then settings_session "$dir" group; else settings_session "$dir" subRoot; fi ;;
            [Mm])
                if [ ${#stack[@]} -gt 1 ]; then p=group; else p=subRoot; fi
                if migrate_session "$dir" "$p"; then stack+=("$NEW_DIR"); CHAIN=("${stack[@]}"); return 0; fi ;;
            [Oo])
                [ ${#stack[@]} -gt 1 ] || continue
                for p in "${CHILDREN[@]}"; do
                    if is_project "$p"; then say "$c_red" 'This folder contains projects, so it must stay a group.'; continue 2; fi
                done
                : > "$dir/$PROJECT_MARKER"
                new_folder_md "$dir" project
                CHAIN=("${stack[@]}"); return 0 ;;
            [Bb]) [ ${#stack[@]} -gt 1 ] || return 1; stack=("${stack[@]:0:${#stack[@]}-1}") ;;
            [Qq]) exit 0 ;;
            *) say "$c_red" 'Invalid choice.' ;;
        esac
    done
}

set_new_root() {
    local path err ans
    read -r -p 'New root path (blank to cancel): ' path || return 0
    path="${path%/}"; path="${path/#\~/$HOME}"
    [ -n "$path" ] || return 0
    if ! err="$(check_root_path "$path")"; then say "$c_red" "$err"; return 0; fi
    if [ ! -d "$path" ]; then
        read -r -p "'$path' does not exist. Create it? [y/N] " ans || return 0
        [[ "$ans" =~ ^[Yy] ]] || return 0
    fi
    # A new root starts from the current root's rules and tools tables.
    init_root "$path" "$ROOT/CLAUDE.md"
    ROOT="$path"
    save_config
    say "$c_green" "Root is now $ROOT (saved to $LAUNCHER_DIR/config)."
}

# ---------- main ----------

if ! err="$(check_root_path "$ROOT")"; then say "$c_red" "Root '$ROOT': $err"; exit 1; fi
init_root "$ROOT" ''
cd "$ROOT" || exit 1
say "$c_cyan" "Root: $ROOT"
tool_check "$ROOT/CLAUDE.md"

CHAIN=()
while [ ${#CHAIN[@]} -eq 0 ]; do
    say "$c_cyan" $'\n'"== $ROOT =="
    for i in "${!SUBROOTS[@]}"; do printf '  %2d) %s\n' $((i + 1)) "${SUBROOTS[$i]}"; done
    say "$c_cyan" '   C) change root   S) settings: open Claude Code at the root   Q) quit'
    read -r -p 'Choose subRoot: ' c || exit 0
    c="${c//[[:space:]]/}"
    if [[ "$c" =~ ^[0-9]+$ ]] && [ "$c" -ge 1 ] && [ "$c" -le ${#SUBROOTS[@]} ]; then
        select_project "$ROOT/${SUBROOTS[$((c - 1))]}" || CHAIN=()
    elif [[ "$c" =~ ^[Ss]$ ]]; then
        settings_session "$ROOT" root
    elif [[ "$c" =~ ^[Cc]$ ]]; then
        set_new_root
        tool_check "$ROOT/CLAUDE.md"
        cd "$ROOT" || exit 1
    elif [[ "$c" =~ ^[Qq]$ ]]; then
        exit 0
    fi
done

# Tools declared by the subRoot, groups and the project itself.
mds=(); for p in "${CHAIN[@]}"; do mds+=("$p/CLAUDE.md"); done
tool_check "${mds[@]}"

project="${CHAIN[${#CHAIN[@]}-1]}"
cd "$project" || exit 1
printf '\e]0;Claude - %s\a' "$(basename "$project")"
say "$c_green" $'\n'"Opening Claude Code in $project"
claude
[ $shell_after = 1 ] && exec "${SHELL:-/bin/zsh}" -l
exit 0
