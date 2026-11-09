#!/bin/bash
# the count "as written" (recount.py) of every word of the extended list and of the words that left, two chunks at a time;
# resumable: a chunk whose recount_*.tsv exists is skipped
# usage: bash run_recount.sh <work folder with recount/words.txt and list_a_1.tsv .. list_b_3.tsv>
HERE=$(cd "$(dirname "$0")" && pwd)
cd "$1"
run() { [ -s recount/out_$1.tsv ] && { echo "skip $1"; return; }; python3 -I "$HERE/recount.py" recount/words.txt $1.tsv recount/out_$1.tsv.tmp > recount/log_$1.txt 2>&1 && mv recount/out_$1.tsv.tmp recount/out_$1.tsv; echo "done $1 $(date +%H:%M:%S)"; }
for k in 1 2 3; do
  run list_a_$k & run list_b_$k & wait
done
echo ALL_DONE
