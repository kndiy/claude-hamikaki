# Shared helpers for start-claude.sh and install-tools.sh. Source only.

TOOLS_START='<!-- launcher:tools:linux:start -->'
TOOLS_END='<!-- launcher:tools:linux:end -->'
PROJECT_MARKER='.claude-project'

c_cyan=$'\e[36m'; c_green=$'\e[32m'; c_yellow=$'\e[33m'; c_red=$'\e[31m'; c_gray=$'\e[90m'; c_off=$'\e[0m'
say() { printf '%s%s%s\n' "$1" "$2" "$c_off"; }

# ---------- config ----------

# Sets ROOT and SUBROOTS from $LAUNCHER_DIR/config (key=value lines).
# root may be relative to this folder ("../..", i.e. the folder holding _launcher,
# the default) so the tree can be copied to another machine; resolve_root
# handles a saved root that does not exist here.
load_config() {
    ROOT='../..'
    SUBROOTS=(codes noncodes)
    local cfg="$LAUNCHER_DIR/config" line
    [ -f "$cfg" ] || return 0
    while IFS= read -r line || [ -n "$line" ]; do
        case "$line" in
            root=*)     ROOT="${line#root=}" ;;
            subRoots=*) IFS=',' read -r -a SUBROOTS <<< "${line#subRoots=}" ;;
        esac
    done < "$cfg"
}

resolve_root() {  # ROOT -> prints an absolute path
    local r="$1" def
    def="$(cd "$LAUNCHER_DIR/../.." && pwd)"
    case "$r" in
        '')  printf '%s\n' "$def"; return ;;
        /*)  ;;
        *)   r="$LAUNCHER_DIR/$r" ;;
    esac
    if [ -d "$r" ]; then (cd "$r" && pwd); else
        printf '%sSaved root %s not found on this machine; using %s%s\n' "$c_yellow" "$r" "$def" "$c_off" >&2
        printf '%s\n' "$def"
    fi
}

save_config() {
    local IFS=','
    printf 'root=%s\nsubRoots=%s\n' "$ROOT" "${SUBROOTS[*]}" > "$LAUNCHER_DIR/config"
}

# Prints an error and returns 1 when the path is not an acceptable root.
check_root_path() {
    case "$1" in
        /mnt/*) echo 'The WSL root must be on the Linux filesystem, not /mnt/* (Windows drives).'; return 1 ;;
        /*)     return 0 ;;
        *)      echo 'Use an absolute path.'; return 1 ;;
    esac
}

# ---------- CLAUDE.md templates ----------

default_tool_block() {
    cat <<EOF
$TOOLS_START
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| curl | \`curl --version\` | required | \`sudo apt-get update && sudo apt-get install -y curl\` | - |
| Claude Code | \`claude --version\` | required | \`curl -fsSL https://claude.ai/install.sh \\| bash\` | - |
| Git | \`git --version\` | recommended | \`sudo apt-get update && sudo apt-get install -y git\` | - |
$TOOLS_END
EOF
}

new_folder_md() {  # DIR KIND(project|group|subRoot)
    local dir="$1" kind="$2" name today
    [ -f "$dir/CLAUDE.md" ] && return 0
    name="$(basename "$dir")"; today="$(date +%F)"
    if [ "$kind" = project ]; then
        cat > "$dir/CLAUDE.md" <<EOF
# $name

Created $today by the launcher. Fill in each section during the first session.

## Purpose and goals
TBD

## Tech stack and key dependencies
TBD

## Structure and key files
TBD

## Commands (build / run / test / lint)
TBD

## Conventions, constraints, decisions
TBD

## Open questions and next steps
TBD

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

$TOOLS_START
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
$TOOLS_END

## Change log
- $today: Project created.
EOF
    else
        cat > "$dir/CLAUDE.md" <<EOF
# $name ($kind)

Rules here apply to every project below this folder. Put project-specific rules in that project's own CLAUDE.md.

## Scope
TBD

## Shared rules
(none yet)
EOF
    fi
}

# Creates the root, its subRoots and their CLAUDE.md files when missing.
# A new root CLAUDE.md is seeded from SEED with {{ROOT}} replaced.
init_root() {  # ROOT SEED
    local root="$1" seed="$2" s text
    mkdir -p "$root"
    if [ ! -f "$root/CLAUDE.md" ]; then
        if [ -n "$seed" ] && [ -f "$seed" ]; then
            text="$(cat "$seed")"
            printf '%s\n' "${text//\{\{ROOT\}\}/$root}" > "$root/CLAUDE.md"
        else
            printf '# Root rules\n\nApplies to every project under this root.\n' > "$root/CLAUDE.md"
        fi
    fi
    if ! grep -qF "$TOOLS_START" "$root/CLAUDE.md"; then
        printf '\n### Linux (WSL Ubuntu) tools\n\n%s\n' "$(default_tool_block)" >> "$root/CLAUDE.md"
    fi
    for s in "${SUBROOTS[@]}"; do
        mkdir -p "$root/$s"
        new_folder_md "$root/$s" subRoot
    done
}

# ---------- tools table ----------

# Prints one tab-separated line per tool: name, check, level, install, status.
# Cells may wrap values in backticks; "\|" inside a cell is a literal pipe.
read_tool_table() {  # CLAUDE_MD
    [ -f "$1" ] || return 0
    awk -v s="$TOOLS_START" -v e="$TOOLS_END" '
        function cell(v) { gsub(/^[ \t]+|[ \t]+$/, "", v); if (v ~ /^`.*`$/) v = substr(v, 2, length(v) - 2); gsub(/\001/, "|", v); return v }
        { sub(/\r$/, "") }
        $0 == s { inb = 1; next }
        $0 == e { inb = 0; next }
        inb && /^[ \t]*\|/ {
            if (++rows <= 2) next
            line = $0; gsub(/\\\|/, "\001", line)
            sub(/^[ \t]*\|/, "", line); sub(/\|[ \t]*$/, "", line)
            n = split(line, c, "|")
            if (n < 5) next
            printf "%s\t%s\t%s\t%s\t%s\n", cell(c[1]), cell(c[2]), tolower(cell(c[3])), cell(c[4]), cell(c[5])
        }' "$1"
}

# Rewrites the Status cells from STATUS_FILE (lines "name<TAB>status"); only touches the file when something changed.
set_tool_status() {  # CLAUDE_MD STATUS_FILE
    local md="$1" tmp
    tmp="$(mktemp)"
    awk -v s="$TOOLS_START" -v e="$TOOLS_END" -F '\t' '
        NR == FNR { st[$1] = $2; next }
        function trim(v) { gsub(/^[ \t]+|[ \t]+$/, "", v); return v }
        { sub(/\r$/, "") }
        $0 == s { inb = 1; print; next }
        $0 == e { inb = 0; print; next }
        inb && /^[ \t]*\|/ && ++rows > 2 {
            line = $0; gsub(/\\\|/, "\001", line)
            sub(/^[ \t]*\|/, "", line); sub(/\|[ \t]*$/, "", line)
            n = split(line, c, "|")
            name = trim(c[1]); gsub(/`/, "", name)
            if (n >= 5 && (name in st)) {
                out = "|"
                for (i = 1; i <= n; i++) { v = (i == 5) ? st[name] : trim(c[i]); gsub(/\001/, "\\|", v); out = out " " v " |" }
                print out; next
            }
        }
        { print }' "$2" "$md" > "$tmp"
    if cmp -s "$tmp" "$md"; then rm -f "$tmp"; else cat "$tmp" > "$md"; rm -f "$tmp"; fi
}

# Runs CHECK in DIR. Prints the first version number, "installed" when none is
# printed, or nothing (status 1) when the command is missing or exits non-zero.
tool_version() {  # CHECK DIR
    local check="$1" dir="${2:-.}" exe out
    exe="${check%% *}"
    command -v "$exe" > /dev/null 2>&1 || return 1
    out="$(cd "$dir" && timeout 60 bash -c "$check" 2>&1)" || return 1
    out="$(printf '%s' "$out" | grep -oE '[0-9]+(\.[0-9]+)+' | head -n 1)"
    printf '%s\n' "${out:-installed}"
}

# Checks every tool in the given CLAUDE.md files, prints a table, and writes
# Status back. Missing tools go to MISSING_TOOLS as "md<TAB>name<TAB>level<TAB>install".
test_tools() {  # CLAUDE_MD...
    local md name check level install status v tmp color
    MISSING_TOOLS=()
    for md in "$@"; do
        [ -f "$md" ] || continue
        tmp="$(mktemp)"
        while IFS=$'\t' read -r name check level install status; do
            [ -s "$tmp" ] || say "$c_cyan" $'\n'"Tools in $md"
            if v="$(tool_version "$check" "$(dirname "$md")" < /dev/null)"; then color="$c_green"; else
                v=MISSING; MISSING_TOOLS+=("$md"$'\t'"$name"$'\t'"$level"$'\t'"$install")
                if [ "$level" = optional ]; then color="$c_gray"; else color="$c_yellow"; fi
            fi
            printf '%s\t%s\n' "$name" "$v" >> "$tmp"
            printf '%s  %-20s %-12s %s%s\n' "$color" "$name" "$level" "$v" "$c_off"
        done < <(read_tool_table "$md")
        [ -s "$tmp" ] && set_tool_status "$md" "$tmp"
        rm -f "$tmp"
    done
}

# The Claude Code native installer puts the binary in ~/.local/bin.
refresh_path() {
    case ":$PATH:" in *":$HOME/.local/bin:"*) ;; *) PATH="$HOME/.local/bin:$PATH"; export PATH ;; esac
    hash -r
}
