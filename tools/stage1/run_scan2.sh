#!/bin/bash
# Scan the additional sources of the Otzaria/otzaria-library repository, one source at a time (download, count, move on).
# Resumable: a source whose out/new_<name>.tsv exists is skipped. Log: stage5/scan2.log
R=/home/claude/work/ext2/otz2
O=/home/claude/work/out
L=/home/claude/work/stage5/scan2.log
cd $R || exit 1
git sparse-checkout init --no-cone >> $L 2>&1
src() {  # name, then folders
  name=$1; shift
  if [ -s "$O/new_$name.tsv" ]; then echo "skip $name" >> $L; return; fi
  printf '/%s/\n' "$@" > /home/claude/work/stage5/pattern_$name.txt
  git sparse-checkout set --no-cone --stdin < /home/claude/work/stage5/pattern_$name.txt >> $L 2>&1
  git checkout >> $L 2>&1
  roots=(); for d in "$@"; do roots+=("$R/$d"); done
  echo "checked out $name $(date +%H:%M:%S) size=$(du -sh $R --exclude=.git 2>/dev/null | cut -f1)" >> $L
  python3 -I /home/claude/work/stage5/scan_multi.py "$O/new_$name.tsv.tmp" "${roots[@]}" >> $L 2>&1 && mv "$O/new_$name.tsv.tmp" "$O/new_$name.tsv"
  echo "done $name $(date +%H:%M:%S)" >> $L
}
src dicta "DictaToOtzaria/ערוך/ספרים" "extraBooks/דיקטה" "extraBooks/דיקטה ערוך"
src orayta "OraytaToOtzaria/ספרים"
src morebooks "MoreBooks/ספרים/אוצריא"
src beitaharon "BeitAharonVeYisraelToOtzaria/ספרים"
src natlib "National-LibraryToOtzaria/ספרים/אוצריא" "extraBooks/National-LibraryToOtzaria"
src toratemet "ToratEmetToOtzaria/ספרים/אוצריא"
src ksk "KSK/ספרים/אוצריא"
src onyourway "OnYourWayToOtzaria/ספרים/אוצריא"
src small "pninimToOtzaria" "wikisourceToOtzaria" "extraBooks/ויקיטקסט" "wikiJewishBooksToOtzaria" "yam-HaHachmaToOtzaria" "tashmaToOtzaria" "Ben-YehudaToOtzaria/ספרים/אוצריא" "extraBooks/לא מסודר לפי מקור" "extraBooks/ספרים פרטיים ועוד" "extraBooks/האוצר השיתופי"
echo SCAN2_DONE >> $L
