#!/bin/bash
# Builds WordCheck 4 (a developer tool, not part of the program): dist/WordCheck4.exe
# Compiled against the .NET Framework 4.5 reference assemblies like the program (BUILD.md), then CheckApi looks inside the finished
# program, then the self-test runs (a pretend Word; Mono only).
#   bash tools/wordcheck4/build.sh
set -euo pipefail
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
B="$ROOT/build/wordcheck4"
rm -rf "$B"
mkdir -p "$B" "$ROOT/dist"
cp "$ROOT/tools/wordcheck4/WordCheck.cs" "$ROOT/tools/wordcheck4/pool4.txt.gz" "$ROOT/src/CheckApi.cs" "$B/"
cd "$B"
API=/usr/lib/mono/4.5-api
if [ ! -d "$API" ]; then echo "missing $API  (apt-get install -y mono-devel)"; exit 1; fi
REFS="-nostdlib -noconfig -lib:$API -r:mscorlib.dll -r:System.dll -r:System.Core.dll -r:System.Drawing.dll -r:System.Windows.Forms.dll"
mcs $REFS -codepage:utf8 -langversion:5 -optimize+ -target:winexe -out:WordCheck4.exe -resource:pool4.txt.gz,WordCheck.pool.txt.gz WordCheck.cs

CECIL=$(ls /usr/lib/mono/gac/Mono.Cecil/*/Mono.Cecil.dll 2>/dev/null | sort -V | tail -1 || true)
if [ -z "$CECIL" ]; then echo "missing Mono.Cecil (the API check cannot run) - apt-get install -y mono-devel"; exit 1; fi
mcs -langversion:5 -r:"$CECIL" -out:CheckApi.exe CheckApi.cs
mono CheckApi.exe WordCheck4.exe

mono WordCheck4.exe --selftest > selftest.txt 2>&1 || { cat selftest.txt; echo "WORDCHECK 4 SELF-TEST FAILED"; exit 1; }
grep -q "^SELFTEST PASSED" selftest.txt || { cat selftest.txt; exit 1; }
grep "^embedded list" selftest.txt
cp WordCheck4.exe "$ROOT/dist/WordCheck4.exe"
(cd "$ROOT/dist" && sha256sum WordCheck4.exe > WordCheck4.exe.sha256)
echo "OK -> dist/WordCheck4.exe"
cat "$ROOT/dist/WordCheck4.exe.sha256"
