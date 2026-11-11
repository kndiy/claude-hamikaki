#!/usr/bin/env bash
# Called by the Windows master runner (Launch.ps1) through wsl.exe.
# Usage: wsl-bridge.sh status|deploy|run WSL_ROOT [WINDOWS_LAUNCHER_DIR]
#   status  exit 0 when WSL_ROOT/_launcher/linux/start-claude.sh exists
#   deploy  copy the launcher (and the root CLAUDE.md + launch.sh when missing) into WSL_ROOT
#   run     start the Linux launcher, leaving a shell in the project folder afterwards
# Paths are passed as arguments (not inline bash) because Windows PowerShell 5.1
# mangles embedded double quotes in native-command arguments.
set -uo pipefail
mode="${1:-}"; root="${2:-}"; src="${3:-}"
root="${root/#\~/$HOME}"; root="${root%/}"

case "$mode" in
    status)
        [ -f "$root/_launcher/linux/start-claude.sh" ] ;;
    deploy)
        case "$root" in /mnt/*) echo 'The WSL root must be on the Linux filesystem, not /mnt/*.'; exit 1 ;; esac
        [ -d "$src" ] || { echo "Launcher source not found: $src"; exit 1; }
        mkdir -p "$root/_launcher"
        cp -r "$src/." "$root/_launcher/"
        # The Linux root keeps its own config; don't carry over the Windows one.
        rm -f "$root/_launcher/config.json"
        [ -f "$root/CLAUDE.md" ] || cp "$src/../CLAUDE.md" "$root/CLAUDE.md"
        [ -f "$root/launch.sh" ] || cp "$src/../launch.sh" "$root/launch.sh"
        find "$root/_launcher" "$root/launch.sh" -name '*.sh' -exec sed -i 's/\r$//' {} + -exec chmod +x {} +
        echo "Launcher installed in $root/_launcher" ;;
    run)
        exec bash "$root/_launcher/linux/start-claude.sh" --shell-after ;;
    *)
        echo "Usage: $0 status|deploy|run WSL_ROOT [WINDOWS_LAUNCHER_DIR]"; exit 2 ;;
esac
