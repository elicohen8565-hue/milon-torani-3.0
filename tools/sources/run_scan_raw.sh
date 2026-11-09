#!/bin/bash
# the places of the words that tokens.norm_token can never produce (a geresh inside: פלאג'י), matched as written; two chunks at a
# time, resumable. usage: bash run_scan_raw.sh <work folder with recount/unstable.txt and list_a_1.tsv .. list_b_3.tsv>
# (merge_sources.py takes the out_raw_*.tsv files after the six out_*.tsv)
HERE=$(cd "$(dirname "$0")" && pwd)
cd "$1"
run() { [ -s out_raw_$1.tsv ] && { echo "skip $1"; return; }; python3 -I "$HERE/scan_sources.py" --raw recount/unstable.txt $1.tsv out_raw_$1.tsv.tmp > log_raw_$1.txt 2>&1 && mv out_raw_$1.tsv.tmp out_raw_$1.tsv; echo "done raw $1 $(date +%H:%M:%S)"; }
for k in 1 2 3; do
  run list_a_$k & run list_b_$k & wait
done
echo ALL_DONE
