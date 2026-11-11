#!/usr/bin/env bash
# restore.sh - fetch the reviewed mirror (origin/main) and copy its files back
# into the real projects tree. Linux (WSL) and macOS (bash 3.2 compatible).
# Windows twin: restore.ps1 (keep both in sync).
#
# Per file:
#   missing locally                         -> created
#   identical                               -> left alone
#   differs, local = an older mirrored copy -> updated (local was just stale)
#   differs, local never mirrored           -> CONFLICT: skipped (unsynced local edits);
#                                              --force backs it up, then overwrites
# Never deletes local files. Skips files whose folder does not exist yet (clone the
# project first, else git clone refuses the non-empty folder); --create-missing overrides.
#
# Usage:
#   bash restore.sh --dry-run          # show what would happen
#   bash restore.sh                    # apply
#   bash restore.sh --force            # overwrite conflicts (backups in .restore-backup/<timestamp>/)
#   bash restore.sh --create-missing   # also create files in folders that do not exist yet
#   bash restore.sh --ref <ref>        # restore from another ref (default origin/main)
set -euo pipefail

MIRROR_OWN="README.md mirror.conf sync.ps1 sync.sh restore.ps1 restore.sh .gitignore .gitattributes"

DRY=0; FORCE=0; CREATE=0; REF="origin/main"
while [ $# -gt 0 ]; do
  case "$1" in
    --dry-run) DRY=1; shift ;;
    --force) FORCE=1; shift ;;
    --create-missing) CREATE=1; shift ;;
    --ref) REF="$2"; shift 2 ;;
    *) echo "Unknown option: $1" >&2; exit 1 ;;
  esac
done

MIRROR="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(dirname "$MIRROR")"
STAMP="$(date '+%Y%m%d-%H%M%S')"
BACKUP="$MIRROR/.restore-backup/$STAMP"
KNOWN="$(mktemp)"; EVER="$(mktemp)"; ATREF="$(mktemp)"; TMP=""
cleanup() {
  [ -n "$TMP" ] && git -C "$MIRROR" worktree remove --force "$TMP" >/dev/null 2>&1 || true
  git -C "$MIRROR" worktree prune >/dev/null 2>&1 || true
  rm -f "$KNOWN" "$EVER" "$ATREF" "${SCOPE_FILE:-}"
}
trap cleanup EXIT
in_list() { case " $2 " in *" $1 "*) return 0 ;; esac; return 1; }
G() { git -c core.quotepath=off "$@"; }

# Scope from mirror.conf (same format as sync.sh), used to keep the
# "removed in mirror" report to this mirror's own paths.
SCOPE_FILE="$(mktemp)"
section=""
TAB="$(printf '\t')"
while IFS= read -r raw || [ -n "$raw" ]; do
  line="${raw%%#*}"; line="$(printf '%s' "$line" | tr -d '\r' | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')"
  [ -n "$line" ] || continue
  case "$line" in \[*\]) section="$(printf '%s' "$line" | tr -d '[]' | tr 'A-Z' 'a-z')"; continue ;; esac
  case "$section" in claude|whole|files|exclude) printf '%s\t%s\n' "$section" "${line%/}" >> "$SCOPE_FILE" ;; esac
done < "$MIRROR/mirror.conf"
in_scope() { # $1 = rel path
  local sec p yes=1
  while IFS="$TAB" read -r sec p; do
    case "$sec" in
      files) [ "$1" = "$p" ] && return 0 ;;
      exclude) case "$1" in "$p"/*) return 1 ;; esac ;;
      claude|whole) case "$1" in "$p"/*) yes=0 ;; esac ;;
    esac
  done < "$SCOPE_FILE"
  return $yes
}

cd "$MIRROR"
case "$REF" in origin/*) git fetch -q origin ;; esac
git rev-parse -q --verify "$REF^{commit}" >/dev/null || { echo "Ref not found: $REF" >&2; exit 1; }
echo "Restoring from $REF ($(git log -1 --format='%h %s' "$REF"))"

# Every version of every path in the reviewed history of $REF: "path<TAB>blob".
# Only $REF, not unmerged sync branches: a local file matching a pending
# (unmerged) version must count as a conflict, not be reverted.
G log --raw --no-abbrev --no-renames --format= "$REF" | while IFS= read -r line; do
  case "$line" in :*) ;; *) continue ;; esac
  meta="${line%%	*}"; path="${line#*	}"
  printf '%s\n' "$path" >> "$EVER"
  set -- $meta
  for b in "$3" "$4"; do case "$b" in *[!0]*) printf '%s\t%s\n' "$path" "$b" >> "$KNOWN" ;; esac; done
done

TMP="$(mktemp -d)"; rmdir "$TMP"
git worktree add -q --detach "$TMP" "$REF"

created=0; updated=0; same=0; conflict=0; forced=0; nofolder=0
while IFS= read -r line; do
  meta="${line%%	*}"; rel="${line#*	}"
  in_list "$rel" "$MIRROR_OWN" && continue
  printf '%s\n' "$rel" >> "$ATREF"
  set -- $meta; mode="$1"; blob="$3"
  dst="$ROOT/$rel"; src="$TMP/$rel"; parent="$(dirname "$dst")"
  if [ ! -f "$dst" ]; then
    if [ ! -d "$parent" ] && [ $CREATE -eq 0 ]; then action=nofolder; else action=created; fi
  else
    h="$(git hash-object --no-filters -- "$dst")"
    if [ "$h" = "$blob" ]; then action=same
    elif grep -qxF "$rel	$h" "$KNOWN"; then action=updated
    elif [ $FORCE -eq 1 ]; then action=forced
    else action=conflict; fi
  fi
  case "$action" in
    same) same=$((same + 1)); continue ;;
    nofolder) nofolder=$((nofolder + 1)); echo "  skip (folder missing)  $rel"; continue ;;
    conflict) conflict=$((conflict + 1)); echo "  CONFLICT (local edits) $rel"; continue ;;
    created) created=$((created + 1)) ;;
    updated) updated=$((updated + 1)) ;;
    forced) forced=$((forced + 1)) ;;
  esac
  printf '  %-22s %s\n' "$action" "$rel"
  [ $DRY -eq 1 ] && continue
  if [ "$action" = forced ]; then
    mkdir -p "$BACKUP/$(dirname "$rel")"; cp -p "$dst" "$BACKUP/$rel"
  fi
  mkdir -p "$parent"; cp "$src" "$dst"
  if [ "$mode" = 100755 ]; then chmod +x "$dst"; fi
done < <(G ls-tree -r "$REF")

# Paths the mirror once held but no longer does: report only.
sort -u "$EVER" | while IFS= read -r p; do
  grep -qxF "$p" "$ATREF" && continue
  in_list "$p" "$MIRROR_OWN" && continue
  in_scope "$p" || continue
  if [ -f "$ROOT/$p" ]; then echo "  removed in mirror, kept locally: $p"; fi
done

echo "Created $created, updated $updated, unchanged $same, conflicts $conflict, forced $forced, skipped (no folder) $nofolder."
if [ $DRY -eq 1 ]; then echo "Dry run: nothing changed."
elif [ $forced -gt 0 ]; then echo "Backups of overwritten files: $BACKUP"; fi
if [ $conflict -gt 0 ]; then
  echo "Conflicts were skipped. Run sync first to send those local edits for review, or rerun with --force."
  exit 3
fi
