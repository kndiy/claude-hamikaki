#!/usr/bin/env bash
# Installs the missing tools listed in one CLAUDE.md's macOS tools table.
# Each Install command runs from the folder holding that CLAUDE.md. Tools that
# can't be installed automatically are listed at the end with the command to run.
# Exit code 1 when anything is still missing.
#
# Usage: install-tools.sh CLAUDE_MD [TOOL_NAME...]   (no names = every missing tool)
set -o pipefail
LAUNCHER_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$LAUNCHER_DIR/lib.sh"

[ $# -ge 1 ] || { echo "Usage: $0 CLAUDE_MD [TOOL_NAME...]"; exit 2; }
md="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"; shift
dir="$(dirname "$md")"
failed=()

# Read the table up front so install commands keep the terminal on stdin (sudo prompts).
rows=()
while IFS= read -r row; do rows+=("$row"); done < <(read_tool_table "$md")
for row in "${rows[@]}"; do
    IFS=$'\t' read -r name check level install status <<< "$row"
    if [ $# -gt 0 ]; then
        want=0; for n in "$@"; do [ "$n" = "$name" ] && want=1; done
        [ $want = 1 ] || continue
    elif tool_version "$check" "$dir" > /dev/null < /dev/null; then
        continue
    fi
    say "$c_cyan" $'\n'">> $name: $install"
    case "$install" in
        manual:*) failed+=("$name"$'\t'"${install#manual:}"); continue ;;
    esac
    (cd "$dir" && bash -c "$install") || say "$c_red" "Install command failed."
    refresh_path
    if tool_version "$check" "$dir" > /dev/null < /dev/null; then
        say "$c_green" "$name installed."
    else
        failed+=("$name"$'\t'"$install")
    fi
done

if [ ${#failed[@]} -gt 0 ]; then
    say "$c_yellow" $'\n'"These tools could not be installed automatically. Run the commands yourself (in $dir), then open a NEW terminal:"
    for f in "${failed[@]}"; do printf '  %-20s %s\n' "${f%%$'\t'*}" "${f#*$'\t'}"; done
    exit 1
fi
exit 0
