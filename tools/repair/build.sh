#!/bin/bash
# Builds the repair tool (Word's list of custom dictionaries full of old entries; 9-11-2026): dist/Milon-Repair.exe
# Made from the program's own code, compiled like the program (BUILD.md), API check, self-test with a pretend Word 365.
#   bash tools/repair/build.sh            (expects build/ from scripts/build_all.sh: the word lists and the icon)
set -euo pipefail
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
B="$ROOT/build"
if [ ! -f "$B/basic.txt.gz" ]; then bash "$ROOT/scripts/build_all.sh"; fi
cp "$ROOT/src/Core.cs" "$ROOT/src/UI.cs" "$ROOT/src/Words.cs" "$ROOT/src/CheckApi.cs" "$ROOT/tools/repair/Repair.cs" "$B/"
cd "$B"
mkdir -p out
API=/usr/lib/mono/4.5-api
REFS="-nostdlib -noconfig -lib:$API -r:mscorlib.dll -r:System.dll -r:System.Core.dll -r:System.Drawing.dll -r:System.Windows.Forms.dll"
RES="-resource:basic.txt.gz,Otzaria.basic.txt.gz -resource:more.txt.gz,Otzaria.more.txt.gz -resource:hebrew_base.txt.gz,Otzaria.hebrew_base.txt.gz -resource:counts.txt,Otzaria.counts.txt"
mcs $REFS -codepage:utf8 -langversion:5 -optimize+ -target:winexe -main:OtzariaDict.RepairProgram -out:out/Repair.exe \
    -win32icon:app.ico $RES -resource:app.ico,Otzaria.app.ico Core.cs UI.cs Words.cs Repair.cs
CECIL=$(ls /usr/lib/mono/gac/Mono.Cecil/*/Mono.Cecil.dll 2>/dev/null | sort -V | tail -1 || true)
mcs -langversion:5 -r:"$CECIL" -out:out/CheckApi.exe CheckApi.cs
mono out/CheckApi.exe out/Repair.exe
mono out/Repair.exe --selftest "$B/repair-selftest" > repair-selftest.txt 2>&1 || { cat repair-selftest.txt; echo "REPAIR SELF-TEST FAILED"; exit 1; }
grep -q "^SELFTEST PASSED" repair-selftest.txt || { cat repair-selftest.txt; exit 1; }
grep -c "^  ok " repair-selftest.txt | sed 's/^/checks passed: /'
mkdir -p "$ROOT/dist"
cp out/Repair.exe "$ROOT/dist/Milon-Repair.exe"
(cd "$ROOT/dist" && sha256sum Milon-Repair.exe > Milon-Repair.exe.sha256)
echo "OK -> dist/Milon-Repair.exe"
