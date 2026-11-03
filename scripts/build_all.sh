#!/bin/bash
# Builds the program from this repository and runs every check. GitHub Actions runs exactly this; by hand (Ubuntu 24.04):
#   sudo apt-get install -y mono-devel mono-mcs libmono-cecil-private-cil python3-pil fonts-dejavu-core
#   bash scripts/build_all.sh
# Result: dist/Milon-Torani-3.0.exe - only when every check below has passed (any failure stops the script).
set -euo pipefail
ROOT=$(cd "$(dirname "$0")/.." && pwd)
B="$ROOT/build"
EXE_NAME="Milon-Torani-3.0.exe"
rm -rf "$B" "$ROOT/dist"
mkdir -p "$B" "$ROOT/dist"

cp "$ROOT"/src/Core.cs "$ROOT"/src/UI.cs "$ROOT"/src/DevDump.cs "$ROOT"/src/CheckApi.cs "$ROOT"/src/build.sh "$ROOT"/src/make_assets.py "$B"/
cp "$ROOT/data/basic_list.txt" "$B/basic.txt"
cp "$ROOT/data/extended_only_list.txt" "$B/more.txt"
cp "$ROOT/data/hebrew_base.txt" "$B/hebrew_base.txt"
cd "$B"

echo "== 1. assets: icon, exact word count of each edition, compressed word lists"
python3 make_assets.py

echo "== 2. build against the .NET Framework 4.5 reference assemblies, then the API check of the finished program"
bash build.sh out

echo "== 3. self-test (the install logic against a temporary folder and a pretend Word)"
mono out/TorahDictionary.exe --selftest "$B/selftest-folder" > selftest.txt 2>&1 || { cat selftest.txt; echo "SELF-TEST FAILED"; exit 1; }
grep -c "^  ok " selftest.txt | sed 's/^/checks passed: /'
grep -q "^SELFTEST PASSED" selftest.txt || { cat selftest.txt; exit 1; }

echo "== 4. the dictionary itself: the four editions byte for byte as in tests/dictionary.sha256"
mono out/DevDump.exe dump
(cd dump && sha256sum -c "$ROOT/tests/dictionary.sha256")

cp out/TorahDictionary.exe "$ROOT/dist/$EXE_NAME"
(cd "$ROOT/dist" && sha256sum "$EXE_NAME" > "$EXE_NAME.sha256")
echo "OK -> dist/$EXE_NAME"
cat "$ROOT/dist/$EXE_NAME.sha256"
