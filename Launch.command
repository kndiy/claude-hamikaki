#!/usr/bin/env bash
# Master runner (macOS). Double-click in Finder (opens Terminal) or run it from a terminal.
# The root is this folder, wherever it is copied. First run on a machine (or after the tree
# moves) offers a "cproj" alias in the shell profile (~/.zshrc for zsh, ~/.bash_profile for bash).
# If double-click says "permission denied": chmod +x Launch.command _launcher/mac/*.sh
set -o pipefail
root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
tag='# claude-projects-launcher'
line="alias cproj='bash \"$root/Launch.command\"' $tag"
case "${SHELL##*/}" in
    zsh)  profile="$HOME/.zshrc" ;;
    bash) profile="$HOME/.bash_profile" ;;
    *)    profile='' ;;
esac

if [ -n "$profile" ] && ! grep -qxF "$line" "$profile" 2>/dev/null; then
    echo "Root on this machine: $root"
    read -r -p "Add alias 'cproj' for this root to $profile? [Y/n] " ans || ans=n
    if [[ ! "$ans" =~ ^[Nn] ]]; then
        if [ -f "$profile" ]; then
            grep -v "$tag\$" "$profile" > "$profile.tmp.$$"; cat "$profile.tmp.$$" > "$profile"; rm -f "$profile.tmp.$$"
        fi
        printf '%s\n' "$line" >> "$profile"
        echo "Added. Open a new terminal and type: cproj"
    fi
elif [ -z "$profile" ]; then
    echo "Shell $SHELL: add this alias to your profile yourself if you want it: $line"
fi

bash "$root/_launcher/mac/start-claude.sh" --shell-after "$@"
