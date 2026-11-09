"""New word lists for version 3.1 (9-11-2026). The basic edition gets every form of the four books that Word marks, and the extended
edition grows with the most valuable forms of the scanned libraries - up to a total of 300,000 words for the biggest edition
(extended + chibur torani). No further Word check (the user asked to go on without one): a form goes in when an earlier Word check
marked it, or when no check asked about it and Hspell does not know it either (82% of such forms were marked before). A form Word
knows anyway costs only a place in the file.

usage: python3 -I build_lists_v2.py <repo> <stage5 folder> <target total> <out folder for the record>
  The lists in <repo>/data must be those of 3.0 (the starting point). Writes <repo>/data/basic_list.txt and
  <repo>/data/extended_only_list.txt (sorted, one word per line, as before) and, for the record, added_basic.tsv /
  added_extended.tsv / moved_to_basic.txt and a summary on stdout.

Rules:
  - a whole form (the user, 3-11: no fragments): at least 3 letters, or at least 2 letters with a quote mark (" or ': an
    abbreviation); a final letter only at the very end (not ןאם, ך"א); not ending in a non-final כ מ נ פ צ unless it has a quote
    mark (not אבנ, אעפכ: a lost final letter or an abbreviation without its mark); at most one quote mark; no letter three times
    in a row (not עייין) - except a prefix before an abbreviation with a doubled letter (בבב"ק) and a few named forms (מממון);
    always kept: לםרבה (Isaiah 9:6, the closed mem) and the divine-name sign ייי, with prefixes;
  - "chibur torani" forms are left to that option of the dictionary;
  - the four books (basic edition):
    - the 796 words of the four books that were only in the extended list move to the basic list, except the words that are in
      the four books only inside verse or page references in parentheses (קוהלת, קיג ...): they stay in extended;
    - a form of the four books goes in when Word marked it, or when no check asked about it and Hspell does not know it - if it
      has a second witness: at least two of the four copies, or at least 3 occurrences, or a library other than the Otzaria
      library (which holds the same files as the Otzaria copy). A form of one copy, seen once or twice, with no other witness is
      more likely a typo of that copy;
    - the four-book forms are taken before the cap, so the cap limits only the extended edition;
  - extended: the remaining forms of the libraries (marked, or untested and unknown to Hspell), best first by the red marks each
    removes (occurrences in all the libraries x the chance Word marks it), until the biggest edition reaches the target;
  - left out: the typos an independent check found in a sample, and text that is not the four books (a reading Sefaria's editors
    mark as deleted, a scribe's colophon).

3.1 changes after the independent check of the categories (9-11, /home/claude/verify/category_report.md): the final-letter
rule also sees a quote mark (42 forms like ך"א); the four-book forms come before the cap (before, 2,138 of them were left out by
it); one rule for a single witness, marked or not (before, 843 marked forms of one copy went in without a second witness); the
non-final ending rule; the exceptions; references and deleted readings are not four-book text.
"""
import itertools
import os
import re
import sys

sys.path.insert(0, '/home/claude/work/scripts')
from hs import hspell_flags  # noqa: E402

repo, st5, target, rec = sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4]
os.makedirs(rec, exist_ok=True)


def read_words(p):
    with open(p, encoding='utf-8') as f:
        return [l.strip() for l in f if l.strip()]


basic = read_words(os.path.join(repo, 'data/basic_list.txt'))
more = read_words(os.path.join(repo, 'data/extended_only_list.txt'))
base = read_words(os.path.join(repo, 'data/hebrew_base.txt'))
assert len(basic) == 52116 and len(more) == 72739, 'the lists in data/ must be those of 3.0 (52,116 + 72,739 lines)'

CONNECTORS = ['ד', 'וד', 'כד', 'וכד', 'מד', 'ומד']
FIRST = 'אבגדהוזחטיכלמנסעפצקרשת'
LAST = 'אבגדהוזחטיךלםןסעףץקרשת'


def acronyms(n):
    return [''.join(h) + '"' + l for h in itertools.product(FIRST, repeat=n - 1) for l in LAST]


acro = set(acronyms(2) + acronyms(3))
comp = set(c + b for b in base for c in CONNECTORS)
LETTER = re.compile('[א-ת]')

FINAL_INSIDE = re.compile('[ךםןףץ](?=.*[א-ת])')   # a final letter with any letter after it, also after a quote mark
NONFINAL_END = re.compile('[כמנפצ]$')
THRICE = re.compile(r'([א-ת])\1\1')
PREFIX = 'והבכלמשד'
ALWAYS = re.compile('^[' + PREFIX + ']{0,2}(לםרבה|ייי\'?)$')

# found by the independent check in its sample: typos of one copy (לככריז ...), a ד/ר slip (יהוידיב for יהויריב), a ח read as ת
TYPOS = {'לככריז', 'תעבדיבה', 'מאלהיהו', 'גגפן', 'והתננן', 'אשכ', 'איסת', 'יהוידיב', 'בכתומ', 'האתרונים'}
# in the four books only inside references in parentheses (the Mechon Mamre copies print "(תהילים קיג)"): stay in extended
REFERENCES = {'קוהלת', 'דנייאל', 'קיג', 'קז', 'קכב', 'קיח', 'קכט', 'קה', 'קיז', 'קכד', 'קכג', 'קסב', 'קסז'}
# not the text of the four books: readings Sefaria's editors mark as deleted (<...>), a scribe's colophon at the end of the Yerushalmi
NOT_TEXT = REFERENCES | {'איצטלי', 'דאתאמר', 'והחיטין', 'חנילא', 'דוקניות'}

have = set(basic) | set(more) | acro | comp
# a prefix before a word with a doubled letter, named by the independent check (most other forms of this shape are typos:
# בבבית, דדדוקא, מממה)
DOUBLED = {'מממוניה', 'מממון', "מממונ'", 'מממרא'}


def no_fragment(w):
    """a whole form (see the rules above)"""
    if ALWAYS.match(w):
        return True
    n = len(LETTER.findall(w))
    quotes = w.count('"') + w.count("'")
    if quotes > 1 or FINAL_INSIDE.search(w):
        return False
    if quotes == 0 and NONFINAL_END.search(w):
        return False
    # three in a row: only a prefix before an abbreviation that starts with a doubled letter (בבב"ק = ב + בב"ק), or the named forms
    if THRICE.search(w) and not (w in DOUBLED or (w[0] in PREFIX and w[0] == w[1] == w[2] and w[1:] in acro)):
        return False
    return n >= 3 or (n >= 2 and quotes == 1)


print('before: basic %d | extended %d | extended + chibur %d' % (
    len(set(basic) | acro), len(set(basic) | acro | set(more)), len(set(basic) | acro | set(more) | comp)))

# the four books
moved_all = [l.split('\t')[0] for l in read_words(os.path.join(st5, 'move_to_basic.tsv'))[1:]]
moved = [w for w in moved_all if w in set(more) and w not in NOT_TEXT]
rows = [l.split('\t') for l in read_words(os.path.join(st5, 'missing_from_basic.tsv'))[1:]]
untested = [r[0] for r in rows if r[1] == 'untested']
hs = dict(zip(untested, hspell_flags(untested)))

# the libraries
cand_rows = []
with open(os.path.join(st5, 'candidates.tsv'), encoding='utf-8') as f:
    f.readline()
    for line in f:
        cand_rows.append(line.rstrip('\n').split('\t'))
lib = {}      # form -> (score, sources, files, status, occurrences, in the four books, a witness besides the Otzaria copy)
for w, st, src, occ, files, hk, ch, onlynew, names in cand_rows:
    if ch == '1' or w in have or w in TYPOS or not no_fragment(w):
        continue
    if st == 'marked':
        p = 1.0
    elif st == 'untested' and hk == '0':
        p = 0.82
    else:
        continue
    srcs = set(names.split(','))
    four = 'four' in srcs and w not in NOT_TEXT
    lib[w] = (int(occ) * p, int(src), int(files), st, int(occ), four, len(srcs - {'four', 'library'}) > 0)

add_basic = []      # (form, how, copies, occurrences)
for w, status, copies, occ, books in rows:
    if w in have or w in TYPOS or w in NOT_TEXT or not no_fragment(w):
        continue
    evidence = status == 'marked' or (status == 'untested' and not hs[w])
    if not evidence:
        continue
    if int(copies) >= 2 or int(occ) >= 3:
        add_basic.append((w, status, int(copies), int(occ)))
    elif w in lib and lib[w][6]:
        add_basic.append((w, status + '+libraries', int(copies), int(occ)))
basic_set = set(basic) | set(moved) | set(w for w, _, _, _ in add_basic)
# a four-book form of the library list that the four-book list does not have as missing (it has no row there): the same rule
row_words = set(r[0] for r in rows)
for w, r in sorted(lib.items()):
    if r[5] and w not in basic_set and r[6] and w not in row_words:
        basic_set.add(w)
        add_basic.append((w, r[3] + '+libraries', 0, r[4]))
more_set = set(more) - set(moved)


def total(b, m):
    return len(b | acro | m | comp)


# extended: the forms that are not four-book forms, best first, up to the target
cands = sorted(((r[0], r[1], r[2], w, r[3], r[4]) for w, r in lib.items() if not r[5] and w not in basic_set),
               key=lambda r: (-r[0], -r[1], -r[2], r[3]))
room = target - total(basic_set, more_set)
chosen = cands[:max(0, room)]
while total(basic_set, more_set | set(r[3] for r in chosen)) > target:
    chosen = chosen[:-1]
more_set |= set(r[3] for r in chosen)

basic_new = sorted(basic_set)
more_new = sorted(more_set - basic_set)
with open(os.path.join(repo, 'data/basic_list.txt'), 'w', encoding='utf-8', newline='\n') as f:
    f.write('\n'.join(basic_new) + '\n')
with open(os.path.join(repo, 'data/extended_only_list.txt'), 'w', encoding='utf-8', newline='\n') as f:
    f.write('\n'.join(more_new) + '\n')
with open(os.path.join(rec, 'moved_to_basic.txt'), 'w', encoding='utf-8', newline='\n') as f:
    f.write('\n'.join(moved) + '\n')
with open(os.path.join(rec, 'added_basic.tsv'), 'w', encoding='utf-8', newline='\n') as f:
    f.write('word\tstatus\tcopies\toccurrences\n')
    for r in add_basic:
        f.write('%s\t%s\t%d\t%d\n' % r)
with open(os.path.join(rec, 'added_extended.tsv'), 'w', encoding='utf-8', newline='\n') as f:
    f.write('rank\tword\tstatus\tsources\tfiles\toccurrences\texpected_red\n')
    for i, r in enumerate(chosen):
        f.write('%d\t%s\t%s\t%d\t%d\t%d\t%.1f\n' % (i, r[3], r[4], r[1], r[2], r[5], r[0]))

lost = (set(basic) | set(more)) - (set(basic_new) | set(more_new))
via_lib = sum(1 for r in add_basic if r[1].endswith('+libraries'))
print('moved from extended to basic: %d (%d stay in extended: in the four books only inside references)' % (
    len(moved), len([w for w in moved_all if w in NOT_TEXT and w in set(more)])))
print('added to basic: %d (marked by Word before %d, untested and unknown to Hspell %d; through the libraries as second witness %d)' % (
    len(add_basic), sum(1 for r in add_basic if r[1].startswith('marked')), sum(1 for r in add_basic if r[1].startswith('untested')), via_lib))
print('added to extended: %d of %d candidates (marked before %d); last score %.2f' % (
    len(chosen), len(cands), sum(1 for r in chosen if r[4] == 'marked'), chosen[-1][0] if chosen else 0))
print('after: basic %d | basic + chibur %d | extended %d | extended + chibur %d' % (
    len(set(basic_new) | acro), len(set(basic_new) | acro | comp), len(set(basic_new) | acro | set(more_new)), total(set(basic_new), set(more_new))))
print('lists: basic_list.txt %d lines, extended_only_list.txt %d lines; old words lost: %d' % (len(basic_new), len(more_new), len(lost)))
print('kept by the exceptions: ' + ' '.join(sorted(w for w in (set(basic_new) | set(more_new)) - (set(basic) | set(more))
                                                   if ALWAYS.match(w) or THRICE.search(w))[:30]))
