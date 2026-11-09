# Where each word appears (3.1, 9-11-2026)

The word-editing window shows next to each word where it appears (the user: "the online libraries are full of typos").
The data file is `data/sources.txt.gz`; the build puts it into the program as the resource `Otzaria.sources.gz`.

How it was made (in the developer's workspace; the paths in the scripts are those of that workspace):

1. `make_lists.py <repo> <out>`: the words to look for (every word of the biggest edition) and the files to scan, in order of
   priority: the Otzaria copy of the four books, the Sefaria copy of the four books, the rest of the Otzaria main library
   (Sivan22/otzaria-library), the rest of Sefaria (Sefaria-Export, Hebrew merged texts), then the other sources of the maintained
   Otzaria library (Dicta, Orayta, the National Library, Torat Emet ...), as scanned in stage 5.
   `make_lists.py <repo> <out> --targets-only` writes only targets.txt (the --words list of step 4).
2. `need_missing.py`: which blobs of the lists are not in the local partial clones (fetch them first; 9-11: one, skipped).
3. `split_lists.py`: six chunks of about equal text; `run_scan.sh <work folder>`: two at a time (two cores), `scan_sources.py`
   each - the files are read straight from git (`git cat-file --batch`, nothing is checked out), words cut exactly as in every
   earlier scan (`tokens.py`). About 25 minutes.
4. `merge_sources.py --words targets.txt out.txt chunk1 ... chunk6` (in order), then `gzip -n -9`.

Per word: per core collection (Tanakh / Mishna / Bavli / Yerushalmi) the number of occurrences and the first 2 places, in one
copy (Otzaria's; the Sefaria copy only for a word that Otzaria's does not have); the 3
other books where the word appears most (each with its count and its first place; a book in several files, or its copies in
several libraries, is one entry); how many other books in all; the total. A place = book, heading (the h2..h6 headings of
Otzaria, or the section line of Sefaria: "דף ב עמוד א", "פרק ג"), the verse / mishna mark or the line number under the heading,
and the word number in the line.

A word of 3.0 also needs at least one place here (tools/stage1/build_lists_v2.py, its 5th argument) - the window shows it.

The result of 9-11 (3.1): 620,371 words looked for; 585,442 found (every word of the two lists; the rest are abbreviations and
"chibur torani" forms that the program makes by itself and no book has - the window marks them in grey).

Which words of the extended list stay is decided by recount.py (run_recount.sh): at least 3 times as written, in all the text
(footnotes and headings too) - the scan above counts the main text through tokens.norm_token, which joins a word across an inner
quote mark. The words with a geresh inside, which norm_token can never produce (פלאג'י), get their places from
scan_sources.py --raw (run_scan_raw.sh); merge_sources.py takes those out_raw_*.tsv files after the six out_*.tsv.
