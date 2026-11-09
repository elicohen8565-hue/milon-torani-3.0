#!/bin/bash
# the sources scan, two chunks at a time (2 cores); resumable: a chunk whose out_*.tsv exists is skipped
# usage: bash run_scan.sh <work folder with targets.txt and list_a_1.tsv .. list_b_3.tsv>
HERE=$(cd "$(dirname "$0")" && pwd)
cd "$1"
run() { [ -s out_$1.tsv ] && { echo "skip $1"; return; }; python3 -I "$HERE/scan_sources.py" targets.txt $1.tsv out_$1.tsv.tmp > log_$1.txt 2>&1 && mv out_$1.tsv.tmp out_$1.tsv; echo "done $1 $(date +%H:%M:%S)"; }
for k in 1 2 3; do
  run list_a_$k & run list_b_$k & wait
done
echo ALL_DONE
