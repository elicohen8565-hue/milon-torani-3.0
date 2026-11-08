#!/bin/bash
# Publishes the built program as a GitHub release - only for a NEW version number. The version is read from src/Core.cs (AppVersion);
# a push that does not change it builds and tests the program but publishes nothing. Needs GH_TOKEN (GitHub Actions gives it).
set -euo pipefail
ROOT=$(cd "$(dirname "$0")/.." && pwd)
cd "$ROOT"
EXE="dist/Milon-Torani-3.1.exe"
OLD="dist/Milon-Torani-3.0.exe"   # the same file under the name of the old download links
VER=$(grep -oP 'AppVersion = "\K[0-9.]+' src/Core.cs)
TAG="v$VER"
if gh release view "$TAG" >/dev/null 2>&1; then
  echo "release $TAG already exists - nothing to publish"
  exit 0
fi
NOTES=$(mktemp)
cat RELEASE_NOTES.md > "$NOTES"
{
  echo
  echo "---"
  echo "SHA-256 של הקובץ: \`$(cut -d' ' -f1 "$EXE.sha256")\`"
} >> "$NOTES"
gh release create "$TAG" "$EXE" "$OLD" --title "המילון התורני 3.1 – גרסה $VER" --notes-file "$NOTES" --latest
echo "published $TAG"
