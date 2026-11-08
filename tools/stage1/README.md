# Stage 1 (the dictionary itself) - scripts and results of 3-11-2026

These scripts ran in the developer's environment. The paths in them are that environment's (`/home/claude/work/...`). The scan outputs they read (`out/*.tsv`, several hundred MB) are not kept here. The full account (in Hebrew) is `claude/STAGE1_DICT.md` in the claude.ai project.

## Scripts

- `split_check.py`: checks the basic/extended split against four copies of the four books (Otzaria, Sefaria, Mamre Tanakh, Mamre Bavli). Its log is `split_check.log`.
  - `move_to_basic.tsv`: 796 words that are only in the extended edition but appear in the four books.
  - `missing_from_basic.tsv.gz`: the forms of the four books that no edition has, with what Word said about them. 5,279 of them Word marks.
- `scan_sefaria4.py`: scans the four books in the Sefaria export.
- `scan_multi.py` and `run_scan2.sh`: scan the nine additional libraries of Otzaria/otzaria-library (Dicta, Orayta, the National Library and more).
- `rank_new.py`: joins all the sources into `candidates.tsv`. Its log is `rank_new.log`. `tools/wordcheck4/make_pool4.py` then makes WordCheck 4's list from it.
