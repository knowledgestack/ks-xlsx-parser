#!/usr/bin/env bash
# publish_wiki.sh — mirror docs/wiki/ into the GitHub wiki for this repo.
#
# GitHub wikis are separate git repos at <repo>.wiki.git. This script clones
# that repo into a tmpdir, copies the markdown files from docs/wiki/, commits,
# and pushes.
#
# Prereqs:
#   1. Enable the Wiki: Settings → General → Features → Wikis ✓
#   2. Seed the wiki with ONE page (any content) in the UI — GitHub won't
#      create the underlying repo until the first page exists.
#
# Usage:
#   scripts/publish_wiki.sh
#   scripts/publish_wiki.sh --dry-run   # show diff, don't push

set -euo pipefail

REPO="${REPO:-knowledgestack/ks-xlsx-parser}"
SRC_DIR="$(cd "$(dirname "$0")/.." && pwd)/docs/wiki"
TMP_DIR="$(mktemp -d)"
DRY_RUN=0
if [[ "${1:-}" == "--dry-run" ]]; then
    DRY_RUN=1
fi

trap 'rm -rf "$TMP_DIR"' EXIT

echo "→ publishing $SRC_DIR → wiki of $REPO"

git clone --depth 1 "git@github.com:${REPO}.wiki.git" "$TMP_DIR" 2>/dev/null || {
    echo "✗ Failed to clone ${REPO}.wiki.git"
    echo "  The wiki isn't initialised yet. Visit"
    echo "  https://github.com/${REPO}/wiki and create any page once."
    exit 1
}

cp -v "$SRC_DIR"/*.md "$TMP_DIR/"

cd "$TMP_DIR"
git add -A

if git diff --cached --quiet; then
    echo "→ wiki already up to date; nothing to push."
    exit 0
fi

if [[ $DRY_RUN -eq 1 ]]; then
    echo "→ dry run — diff below:"
    git diff --cached --stat
    exit 0
fi

git commit -m "docs(wiki): sync from docs/wiki/ ($(date -u +%Y-%m-%dT%H:%MZ))"
git push origin master 2>/dev/null || git push origin main
echo "✓ wiki published → https://github.com/${REPO}/wiki"
