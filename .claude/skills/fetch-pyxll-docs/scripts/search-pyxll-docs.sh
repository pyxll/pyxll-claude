#!/usr/bin/env bash
# search-pyxll-docs.sh — Search PyXLL full docs, return matching page URLs.
#
# Downloads and caches the full docs locally (refreshed every 24 h) so only
# the relevant pages need to be loaded into the AI context.
#
# Usage: ./search-pyxll-docs.sh <keyword> [keyword2 ...] [-f|--fresh]
#
# Works on macOS, Linux, Git Bash (Windows), and WSL.

set -euo pipefail

FULL_URL="https://www.pyxll.com/llms-full.txt"
CACHE_DIR="${XDG_CACHE_HOME:-$HOME/.cache}/pyxll-docs"
CACHE_FILE="$CACHE_DIR/llms-full.txt"
CACHE_MAX_AGE=86400  # 24 hours

usage() { echo "Usage: $0 [-f|--fresh] <keyword> [keyword2 ...]"; exit 1; }

FRESH=false
KEYWORDS=()
while [[ $# -gt 0 ]]; do
    case $1 in
        -f|--fresh) FRESH=true; shift ;;
        -h|--help)  usage ;;
        -*)         echo "Unknown option: $1"; usage ;;
        *)          KEYWORDS+=("$1"); shift ;;
    esac
done
[[ ${#KEYWORDS[@]} -eq 0 ]] && { echo "Error: at least one keyword required"; usage; }

mkdir -p "$CACHE_DIR"

need_download() {
    [[ "$FRESH" == "true" ]] && return 0
    [[ ! -f "$CACHE_FILE" ]] && return 0
    local age
    if [[ "$(uname)" == "Darwin" ]]; then
        age=$(( $(date +%s) - $(stat -f %m "$CACHE_FILE") ))
    else
        age=$(( $(date +%s) - $(stat -c %Y "$CACHE_FILE") ))
    fi
    [[ $age -gt $CACHE_MAX_AGE ]]
}

if need_download; then
    echo "Downloading PyXLL docs (cached for 24 h)..." >&2
    curl -sL "$FULL_URL" -o "$CACHE_FILE"
fi

# Build pipe-separated lowercase keyword pattern for awk
kw_pattern=""
for kw in "${KEYWORDS[@]}"; do
    kw_pattern="${kw_pattern}|$(printf '%s' "$kw" | tr '[:upper:]' '[:lower:]')"
done
kw_pattern="${kw_pattern:1}"

echo "Searching for: ${KEYWORDS[*]}" >&2

# PyXLL llms-full.txt page format:
#   ---                                    <- page boundary
#   ## Title                               <- page header (## not ###)
#   [path.md](https://www.pyxll.com/...)  <- URL line immediately after title
#   ... content ...
awk -v kw="$kw_pattern" '
BEGIN { url=""; n=split(kw, K, "|"); state=0 }
/^---$/                                { state=1; next }
state==1 && /^## /                     { state=2; next }
state==2 && index($0, "www.pyxll.com") {
    if (split($0, a, "(") >= 2) {
        cp = index(a[2], ")")
        if (cp > 0) url = substr(a[2], 1, cp-1)
    }
    state=0; next
}
url != "" {
    line = tolower($0)
    for (i=1; i<=n; i++) {
        if (index(line, K[i]) > 0 && !seen[url]) {
            seen[url]=1; print url; break
        }
    }
}
' "$CACHE_FILE"
