#!/bin/bash
# Builds the size check (a developer tool, not part of the program): dist/Milon-SizeCheck.exe
# It is made from the program's own code, so first the whole program is built and checked (scripts/build_all.sh),
# then the tool is compiled the same way (the .NET Framework 4.5 reference assemblies), its API is checked, and its self-test runs.
#   bash tools/build_sizecheck.sh
set -euo pipefail
ROOT=$(cd "$(dirname "$0")/.." && pwd)
B="$ROOT/build"
bash "$ROOT/scripts/build_all.sh"

cp "$ROOT/tools/SizeCheck.cs" "$B/"
cd "$B"
API=/usr/lib/mono/4.5-api
REFS="-nostdlib -noconfig -lib:$API -r:mscorlib.dll -r:System.dll -r:System.Core.dll -r:System.Drawing.dll -r:System.Windows.Forms.dll"
RES="-resource:basic.txt.gz,Otzaria.basic.txt.gz -resource:more.txt.gz,Otzaria.more.txt.gz -resource:hebrew_base.txt.gz,Otzaria.hebrew_base.txt.gz -resource:counts.txt,Otzaria.counts.txt"
echo "== size check: build"
mcs $REFS -codepage:utf8 -langversion:5 -optimize+ -target:winexe -main:OtzariaDict.SizeCheckProgram -out:out/SizeCheck.exe \
    -win32icon:app.ico $RES -resource:app.ico,Otzaria.app.ico Core.cs UI.cs Words.cs SizeCheck.cs
mono out/CheckApi.exe out/SizeCheck.exe

echo "== size check: self-test (a pretend Word)"
mono out/SizeCheck.exe --selftest "$B/sizecheck-selftest" > sizecheck-selftest.txt 2>&1 || { cat sizecheck-selftest.txt; echo "SIZE CHECK SELF-TEST FAILED"; exit 1; }
grep -c "^  ok " sizecheck-selftest.txt | sed 's/^/checks passed: /'
grep -q "^SELFTEST PASSED" sizecheck-selftest.txt || { cat sizecheck-selftest.txt; exit 1; }

cp out/SizeCheck.exe "$ROOT/dist/Milon-SizeCheck.exe"
(cd "$ROOT/dist" && sha256sum Milon-SizeCheck.exe > Milon-SizeCheck.exe.sha256)
echo "OK -> dist/Milon-SizeCheck.exe"
cat "$ROOT/dist/Milon-SizeCheck.exe.sha256"
