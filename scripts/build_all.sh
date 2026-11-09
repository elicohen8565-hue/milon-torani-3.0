#!/bin/bash
# Builds the program from this repository and runs every check. GitHub Actions runs exactly this; by hand (Ubuntu 24.04):
#   sudo apt-get install -y mono-devel mono-mcs libmono-cecil-private-cil python3-pil fonts-dejavu-core
#   bash scripts/build_all.sh
# Result: dist/Milon-Torani-3.1.exe (and the same file as Milon-Torani-3.0.exe, the name of the old download links) -
# only when every check below has passed (any failure stops the script).
set -euo pipefail
ROOT=$(cd "$(dirname "$0")/.." && pwd)
B="$ROOT/build"
EXE_NAME="Milon-Torani-3.1.exe"
OLD_NAME="Milon-Torani-3.0.exe"   # links to the 3.0 download (README of 3.0, posts) keep getting the newest version

# one version everywhere: the program's number (Core.cs), the file's details (UI.cs) and the download name
APPV=$(grep -oP 'AppVersion = "\K[0-9.]+' "$ROOT/src/Core.cs")
ASMV=$(grep -oP 'AssemblyVersion\("\K[0-9.]+' "$ROOT/src/UI.cs")
[ "$APPV" = "$ASMV" ] || { echo "version mismatch: AppVersion $APPV in Core.cs, AssemblyVersion $ASMV in UI.cs"; exit 1; }
[ "$EXE_NAME" = "Milon-Torani-${APPV%.*.*}.exe" ] || { echo "the download name $EXE_NAME does not match version $APPV"; exit 1; }

rm -rf "$B" "$ROOT/dist"
mkdir -p "$B" "$ROOT/dist"

cp "$ROOT"/src/Core.cs "$ROOT"/src/UI.cs "$ROOT"/src/Words.cs "$ROOT"/src/DevDump.cs "$ROOT"/src/CheckApi.cs "$ROOT"/src/build.sh "$ROOT"/src/make_assets.py "$B"/
cp "$ROOT/data/basic_list.txt" "$B/basic.txt"
cp "$ROOT/data/extended_only_list.txt" "$B/more.txt"
cp "$ROOT/data/hebrew_base.txt" "$B/hebrew_base.txt"
cd "$B"

echo "== 1. assets: icon, exact word count of each edition, compressed word lists"
python3 make_assets.py

# the key for sending suggestions (a GitHub token that may only open issues in this repository), from the secret ISSUE_TOKEN of
# GitHub Actions. It goes into the program hidden the same way Suggest.Encode in src/Core.cs hides it; it is never printed or kept.
if [ -n "${ISSUE_TOKEN:-}" ]; then
  trap 'rm -f "$B/issue.bin"' EXIT   # never left behind, also when a later step fails
  python3 - <<'PY'
import os
t = os.environ["ISSUE_TOKEN"].strip().encode("utf-8")
mask = b"milon-torani-3.0/issues"
out = bytes(b ^ mask[i % len(mask)] ^ ((i * 31 + 7) & 0xFF) for i, b in enumerate(t))
open("issue.bin", "wb").write(out)
PY
  echo "the key for sending suggestions: included"
else
  echo "the key for sending suggestions: none (suggestions go through the browser)"
fi

echo "== 2. build against the .NET Framework 4.5 reference assemblies, then the API check of the finished program"
bash build.sh out

echo "== 3. self-test (the install logic against a temporary folder and a pretend Word)"
mono out/TorahDictionary.exe --selftest "$B/selftest-folder" > selftest.txt 2>&1 || { cat selftest.txt; echo "SELF-TEST FAILED"; exit 1; }
grep -c "^  ok " selftest.txt | sed 's/^/checks passed: /'
grep -q "^SELFTEST PASSED" selftest.txt || { cat selftest.txt; exit 1; }

echo "== 4. the dictionary itself: the four editions byte for byte as in tests/dictionary.sha256"
mono out/DevDump.exe dump
(cd dump && sha256sum -c "$ROOT/tests/dictionary.sha256")

rm -f issue.bin
cp out/TorahDictionary.exe "$ROOT/dist/$EXE_NAME"
cp out/TorahDictionary.exe "$ROOT/dist/$OLD_NAME"
(cd "$ROOT/dist" && sha256sum "$EXE_NAME" > "$EXE_NAME.sha256" && sha256sum "$OLD_NAME" > "$OLD_NAME.sha256")
echo "OK -> dist/$EXE_NAME"
cat "$ROOT/dist/$EXE_NAME.sha256"
