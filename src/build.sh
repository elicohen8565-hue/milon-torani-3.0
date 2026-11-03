#!/bin/bash
# Builds TorahDictionary.exe (an ASCII name on purpose; it is handed to the user under the Hebrew name "המילון התורני 3.0.exe", see BUILD.md)
# and the developer tools DevDump.exe / CheckApi.exe (the files this folder must contain: see BUILD.md).
#   ./build.sh [output folder]        (default: the current folder)
#
# IMPORTANT: this compiles against the .NET Framework 4.5 REFERENCE assemblies (/usr/lib/mono/4.5-api), NOT against Mono's own libraries.
# Mono's own library has methods that .NET Framework (Windows) does not have, for example String.Split(char, StringSplitOptions).
# A program compiled against it runs fine under Mono and the self-test passes, but on Windows it dies with MissingMethodException
# (this is what happened to 3.0.0.3: the window never opened). With the reference assemblies such a call either does not compile
# or is bound to the method that Windows really has. After the build, CheckApi.exe looks inside the finished program once more.
set -e
set -o pipefail
OUT="${1:-.}"
mkdir -p "$OUT"
API=/usr/lib/mono/4.5-api
if [ ! -d "$API" ]; then echo "missing $API  (apt-get install -y mono-devel)"; exit 1; fi

REFS="-nostdlib -noconfig -lib:$API -r:mscorlib.dll -r:System.dll -r:System.Core.dll -r:System.Drawing.dll -r:System.Windows.Forms.dll"
RES="-resource:basic.txt.gz,Otzaria.basic.txt.gz -resource:more.txt.gz,Otzaria.more.txt.gz -resource:hebrew_base.txt.gz,Otzaria.hebrew_base.txt.gz -resource:counts.txt,Otzaria.counts.txt"
OPT="-codepage:utf8 -langversion:5 -optimize+"

mcs $REFS $OPT -target:winexe -out:"$OUT/TorahDictionary.exe" -win32icon:app.ico $RES -resource:app.ico,Otzaria.app.ico Core.cs UI.cs
mcs $REFS $OPT -target:exe -main:OtzariaDict.DevDump -out:"$OUT/DevDump.exe" $RES Core.cs DevDump.cs

# second guard, on the finished program: look inside it (Mono.Cecil) and check every .NET library member it uses against the reference assemblies
CECIL=$(ls /usr/lib/mono/gac/Mono.Cecil/*/Mono.Cecil.dll 2>/dev/null | sort -V | tail -1 || true)
if [ -n "$CECIL" ] && [ -f CheckApi.cs ]; then
  mcs -langversion:5 -r:"$CECIL" -out:"$OUT/CheckApi.exe" CheckApi.cs
  mono "$OUT/CheckApi.exe" "$OUT/TorahDictionary.exe"
else
  echo "WARNING: Mono.Cecil or CheckApi.cs not found - only the quick check ran"
  if command -v ikdasm >/dev/null 2>&1 && ikdasm "$OUT/TorahDictionary.exe" | grep -q -E "String::Split\(char,|Path::Join|String::Contains\(char|String::Join\(char"; then
    echo "BAD: the program calls a method that .NET Framework does not have"; exit 1
  fi
fi
echo "built OK -> $OUT"
