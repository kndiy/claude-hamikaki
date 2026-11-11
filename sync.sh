#!/usr/bin/env bash
# sync.sh - mirror Claude-related text files from the projects tree into this
# repo, commit them on a NEW branch sync/<timestamp>, and push that branch for
# review. main only changes when you merge the pull request on GitHub.
# What gets mirrored is set in mirror.conf next to this script; the script itself
# is identical in every mirror repo (_claude-rnd4u, _claude-hamikaki).
# Linux (WSL) and macOS (bash 3.2 compatible). Windows twin: sync.ps1 (keep both in sync).
#
# Usage:
#   bash sync.sh                      # sync + commit on new branch + push branch
#   bash sync.sh --remote <url>       # first time: set origin
#   bash sync.sh --dry-run            # list what would be mirrored, change nothing
#   bash sync.sh --no-push            # sync + commit on new branch, no push
#   bash sync.sh --allow-secrets      # commit even if the secret scan finds something
#   bash sync.sh --message "text"     # custom commit message
set -euo pipefail

# ---------- fixed rules (same in every mirror; keep in step with sync.ps1) ----------
EXCLUDE_NAMES="settings.local.json"
TEXT_EXT="md txt json xml py groovy java sh ps1 patch diff yml yaml sql properties cmd command bat js"
EXEC_EXT="sh command"   # kept executable (git mode 755)
MAX_BYTES=1048576
SECRET_RE='-----BEGIN [A-Z ]*PRIVATE KEY-----|gh[pousr]_[A-Za-z0-9]{36}|github_pat_[A-Za-z0-9_]{20,}|AKIA[0-9A-Z]{16}|sk-ant-[A-Za-z0-9_-]{20,}|xox[abprs]-[A-Za-z0-9-]{10,}|"(password|passwd|pwd|secret|clientSecret|apiKey|api_key|accessToken)"[[:space:]]*:[[:space:]]*"[^"<>${}*]{6,}"'
SKIP_DIRS=".git node_modules target build dist __pycache__ .venv venv"
# -------------------------------------------------------------------------------------

REMOTE=""; MESSAGE=""; DRY=0; NOPUSH=0; ALLOW=0
while [ $# -gt 0 ]; do
  case "$1" in
    --remote) REMOTE="$2"; shift 2 ;;
    --message) MESSAGE="$2"; shift 2 ;;
    --dry-run) DRY=1; shift ;;
    --no-push) NOPUSH=1; shift ;;
    --allow-secrets) ALLOW=1; shift ;;
    *) echo "Unknown option: $1" >&2; exit 1 ;;
  esac
done

MIRROR="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(dirname "$MIRROR")"
WORK="$(mktemp -d)"; trap 'rm -rf "$WORK"' EXIT
LIST="$WORK/list"; : > "$LIST"

# mirror.conf -> one file per section ($WORK/conf.<section>), one item per line.
# Sections: [claude] [whole] [files] [exclude] [scripts] - see sync.ps1 for meanings.
CONF="$MIRROR/mirror.conf"
[ -f "$CONF" ] || { echo "Missing $CONF" >&2; exit 1; }
for sec in claude whole files exclude scripts; do : > "$WORK/conf.$sec"; done
section=""
while IFS= read -r raw || [ -n "$raw" ]; do
  line="${raw%%#*}"; line="$(printf '%s' "$line" | tr -d '\r' | sed -e 's/^[[:space:]]*//' -e 's/[[:space:]]*$//')"
  [ -n "$line" ] || continue
  case "$line" in
    \[*\]) section="$(printf '%s' "$line" | tr -d '[]' | tr 'A-Z' 'a-z')"
           [ -f "$WORK/conf.$section" ] || { echo "Unknown section [$section] in mirror.conf" >&2; exit 1; }
           continue ;;
  esac
  [ -n "$section" ] || { echo "mirror.conf: item outside a section: $line" >&2; exit 1; }
  line="${line%/}"
  case "$line" in /*|*..*) echo "mirror.conf: bad path: $line" >&2; exit 1 ;; esac
  if [ "$section" = scripts ]; then
    for x in $line; do x="${x#.}"; printf '%s\n' "$x" | tr 'A-Z' 'a-z' >> "$WORK/conf.scripts"; done
  else
    printf '%s\n' "$line" >> "$WORK/conf.$section"
  fi
done < "$CONF"
SCRIPT_EXT="$(tr '\n' ' ' < "$WORK/conf.scripts")"
cat "$WORK/conf.claude" "$WORK/conf.whole" > "$WORK/scopes"
cat "$WORK/scopes" "$WORK/conf.files" > "$WORK/targets"

in_list() { case " $2 " in *" $1 "*) return 0 ;; esac; return 1; }
excluded() { # $1 = rel path
  local x
  while IFS= read -r x; do case "$1" in "$x"|"$x"/*) return 0 ;; esac; done < "$WORK/conf.exclude"
  return 1
}
lower_ext() { case "$1" in *.*) printf '%s' "${1##*.}" | tr 'A-Z' 'a-z' ;; *) printf '' ;; esac; }

count=0
add_file() { # $1 = rel path
  local size
  size=$(wc -c < "$ROOT/$1" | tr -d ' ')
  if [ "$size" -gt "$MAX_BYTES" ]; then echo "  skipped (> 1 MB): $1"; return 0; fi
  printf '%s\n' "$1" >> "$LIST"; count=$((count + 1))
}

while IFS= read -r s; do
  if [ ! -d "$ROOT/$s" ]; then echo "WARN scope not found: $s"; continue; fi
  mode=claude; grep -qxF -- "$s" "$WORK/conf.whole" && mode=whole
  # Nested git repos inside this scope: helper scripts in them are already in git.
  : > "$WORK/repos"
  while IFS= read -r g; do g="${g#"$ROOT"/}"; printf '%s\n' "${g%/.git}" >> "$WORK/repos"; done < <(find "$ROOT/$s" -name .git -prune -print)
  # find does the pre-filtering (one process): prune skipped and excluded folders,
  # and in [claude] scopes print only Claude files and helper-script candidates.
  set -- "$ROOT/$s" "("
  first=1
  for d in $SKIP_DIRS; do
    if [ $first -eq 1 ]; then set -- "$@" -name "$d"; first=0; else set -- "$@" -o -name "$d"; fi
  done
  while IFS= read -r x; do set -- "$@" -o -path "$ROOT/$x"; done < "$WORK/conf.exclude"
  set -- "$@" ")" -prune -o -type f
  if [ $mode = claude ]; then
    set -- "$@" "(" -name CLAUDE.md -o -name CLAUDE.local.md -o -name .claude-project -o -path "*/.claude/*" -o -path "*/zClaudeDocs/*"
    for x in $SCRIPT_EXT; do set -- "$@" -o -iname "*.$x"; done
    set -- "$@" ")"
  fi
  set -- "$@" -print
  while IFS= read -r f; do
    rel="${f#"$ROOT"/}"; name="${f##*/}"
    excluded "$rel" && continue
    in_list "$name" "$EXCLUDE_NAMES" && continue
    ext="$(lower_ext "$name")"
    if [ $mode = whole ]; then
      if [ -n "$ext" ] && in_list "$ext" "$TEXT_EXT"; then add_file "$rel"; fi
      continue
    fi
    claude=0
    case "$name" in CLAUDE.md|CLAUDE.local.md|.claude-project) claude=1 ;; esac
    case "/$rel" in */.claude/*|*/zClaudeDocs/*) claude=1 ;; esac
    if [ $claude -eq 1 ]; then
      if [ "$name" = .claude-project ] || { [ -n "$ext" ] && in_list "$ext" "$TEXT_EXT"; }; then add_file "$rel"; fi
      continue
    fi
    if [ -z "$ext" ] || ! in_list "$ext" "$SCRIPT_EXT"; then continue; fi
    inrepo=0
    while IFS= read -r r; do case "$rel" in "$r"/*) inrepo=1; break ;; esac; done < "$WORK/repos"
    if [ $inrepo -eq 0 ]; then add_file "$rel"; fi
  done < <(find "$@")
done < "$WORK/scopes"
while IFS= read -r r; do
  if [ -f "$ROOT/$r" ]; then add_file "$r"; else echo "WARN root file not found: $r"; fi
done < "$WORK/conf.files"
desc="$(tr '\n' ' ' < "$WORK/targets")"
[ -s "$WORK/conf.exclude" ] && desc="$desc(excluding $(tr '\n' ' ' < "$WORK/conf.exclude" | sed 's/ $//'))"
echo "Found $count file(s) in: $desc"

if [ $DRY -eq 1 ]; then sed 's/^/  /' "$LIST"; echo "Dry run: nothing changed."; exit 0; fi

# Git prep: new branch from origin/main (or local main if there is no remote yet),
# or from the newest sync branch while that one is still open (stacked), so edits that
# exist only on an unmerged branch (e.g. script changes) are never dropped.
cd "$MIRROR"
[ -d .git ] || git init -q -b main
if [ -n "$REMOTE" ]; then
  if git remote get-url origin >/dev/null 2>&1; then git remote set-url origin "$REMOTE"; else git remote add origin "$REMOTE"; fi
fi
HAS_REMOTE=0; git remote get-url origin >/dev/null 2>&1 && HAS_REMOTE=1
if [ $HAS_REMOTE -eq 1 ] && [ $NOPUSH -eq 0 ]; then git fetch -q --prune origin; fi
if git rev-parse -q --verify refs/remotes/origin/main >/dev/null; then BASE="origin/main"
elif git rev-parse -q --verify refs/heads/main >/dev/null; then BASE="main"
else echo "No main branch yet. Commit the scripts to main and push it once (see README)." >&2; exit 1; fi
# The scope folders are regenerated below, so discard leftovers from a failed run.
while IFS= read -r s; do
  if git ls-files --error-unmatch -- "$s" >/dev/null 2>&1; then git checkout -q HEAD -- "$s"; fi
  git clean -fdq -- "$s"
done < "$WORK/targets"
# Newest existing sync branch: a run with nothing new since then makes no new branch.
PREV="$(git for-each-ref --sort=-refname --count=1 --format='%(refname:short)' refs/heads/sync/)"
BRANCH="sync/$(date '+%Y%m%d-%H%M%S')"
START="$BASE"
if [ -n "$PREV" ] && ! git merge-base --is-ancestor "$PREV" "$BASE"; then
  # Open = not merged, and still on GitHub (or never pushed). A branch closed and
  # deleted on GitHub without merging counts as rejected and is not stacked on.
  if git rev-parse -q --verify "refs/remotes/origin/$PREV" >/dev/null || ! git config --get "branch.$PREV.remote" >/dev/null; then
    START="$PREV"
    echo "Stacking on unmerged $PREV. Merge only the newest PR and close the older one."
  fi
fi
git checkout -q -b "$BRANCH" "$START"

# Replace each scope in the mirror with a fresh copy (handles deletes/renames).
while IFS= read -r s; do rm -rf "${MIRROR:?}/$s"; done < "$WORK/targets"
while IFS= read -r rel; do
  mkdir -p "$MIRROR/$(dirname "$rel")"
  cp -p "$ROOT/$rel" "$MIRROR/$rel"
done < "$LIST"

# Secret scan on the mirrored copies.
hits=""
while IFS= read -r rel; do
  if grep -qE -- "$SECRET_RE" "$MIRROR/$rel" 2>/dev/null; then hits="$hits
  $rel"; fi
done < "$LIST"
if [ -n "$hits" ]; then
  echo "Possible secrets found:$hits"
  if [ $ALLOW -eq 0 ]; then
    echo "Not committing. Remove them at the source, or rerun with --allow-secrets if they are false positives."
    exit 2
  fi
fi

# Commit on the sync branch and push it.
git add -A
while IFS= read -r rel; do
  ext="$(lower_ext "${rel##*/}")"
  if [ -n "$ext" ] && in_list "$ext" "$EXEC_EXT"; then git update-index --chmod=+x -- "$rel"; fi
done < "$LIST"
if git diff --cached --quiet "$BASE"; then
  echo "Nothing new compared with $BASE. Removing branch $BRANCH."
  git checkout -q --detach "$BASE"; git branch -q -D "$BRANCH"; exit 0
fi
if [ -n "$PREV" ] && git diff --cached --quiet "$PREV"; then
  echo "Nothing new since $PREV (not merged yet). Review that branch instead; no new branch made."
  git checkout -q "$PREV"; git branch -q -D "$BRANCH"; exit 0
fi
[ -n "$MESSAGE" ] || MESSAGE="Sync Claude files $(date '+%Y-%m-%d %H:%M')"
git commit -q -m "$MESSAGE"
git --no-pager log -1 --shortstat --format="Committed %h on $BRANCH"
# Each sync branch holds everything not yet on main, so older local ones are redundant.
git for-each-ref --format='%(refname:short)' refs/heads/sync/ | while IFS= read -r b; do
  [ "$b" = "$BRANCH" ] || git branch -q -D "$b"
done
if [ $NOPUSH -eq 1 ]; then echo "Push skipped (--no-push). Branch $BRANCH is local only."; exit 0; fi
if [ $HAS_REMOTE -eq 0 ]; then
  echo "No remote yet. Run once with: --remote git@github.com:<owner>/<repo>.git"; exit 0
fi
git push -q -u origin "$BRANCH"
URL="$(git remote get-url origin | sed -e 's#^git@github\.com:#https://github.com/#' -e 's#\.git$##')"
echo
echo "Pushed $BRANCH. Review and merge it here:"
echo "  $URL/compare/main...$BRANCH?expand=1"
