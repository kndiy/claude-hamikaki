#!/usr/bin/env bash
# Master runner (Linux / WSL Ubuntu). The root is this folder, wherever it is copied.
# First run on a machine (or after the tree moves) offers a "cproj" alias in ~/.bashrc.
set -uo pipefail
root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
tag='# claude-projects-launcher'
line="alias cproj='bash \"$root/launch.sh\"' $tag"

if ! grep -qxF "$line" "$HOME/.bashrc" 2>/dev/null; then
    echo "Root on this machine: $root"
    read -r -p "Add alias 'cproj' for this root to ~/.bashrc? [Y/n] " ans || ans=n
    if [[ ! "$ans" =~ ^[Nn] ]]; then
        [ -f "$HOME/.bashrc" ] && sed -i "/$tag\$/d" "$HOME/.bashrc"
        printf '%s\n' "$line" >> "$HOME/.bashrc"
        echo "Added. Open a new terminal and type: cproj"
    fi
fi

exec bash "$root/_launcher/linux/start-claude.sh" --shell-after "$@"
