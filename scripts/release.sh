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
  # a release whose program files did not all arrive (an upload that failed half way) is completed; a complete one is left alone
  HAVE=$(gh release view "$TAG" --json assets --jq '.assets[].name')
  UP=()
  for f in "$EXE" "$OLD"; do
    grep -qxF "$(basename "$f")" <<<"$HAVE" || UP+=("$f" "$f.sha256")
  done
  if [ ${#UP[@]} -eq 0 ]; then
    echo "release $TAG already exists - nothing to publish"
    exit 0
  fi
  gh release upload "$TAG" "${UP[@]}" --clobber
  echo "release $TAG completed with: ${UP[*]}"
  exit 0
fi
NOTES=$(mktemp)
cat RELEASE_NOTES.md > "$NOTES"
{
  echo
  echo "---"
  echo "SHA-256 של הקובץ: \`$(cut -d' ' -f1 "$EXE.sha256")\`"
} >> "$NOTES"
gh release create "$TAG" "$EXE" "$OLD" "$EXE.sha256" "$OLD.sha256" --title "המילון התורני 3.1 – גרסה $VER" --notes-file "$NOTES" --latest
echo "published $TAG"
