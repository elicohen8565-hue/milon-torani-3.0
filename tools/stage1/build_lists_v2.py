"""New word lists (version 3.0.0.10, 9-11-2026): the basic edition gets every form of the four books that Word marks, and the
extended edition grows with the most valuable forms of the scanned libraries - up to a total of about 300,000 words for the
biggest edition (extended + chibur torani). No further Word check (the user asked to go on without one): a form goes in when an
earlier Word check marked it, or when no check asked about it and Hspell does not know it either (82% of such forms were marked
before). A form Word knows anyway costs only a place in the file.

usage: python3 -I build_lists_v2.py <repo> <stage5 folder> <target total> <out folder for the record>
  writes <repo>/data/basic_list.txt and <repo>/data/extended_only_list.txt (sorted, one word per line, as before) and, for the
  record, added_basic.tsv / added_extended.tsv / moved_to_basic.txt and a summary on stdout.

Rules:
  - no fragments (the user, 3-11): at least 3 letters, or at least 2 letters with a quote mark (" or ': an abbreviation); and no
    typo-like forms: a final letter inside the form, a letter three times in a row, two quote marks;
  - "chibur torani" forms are left to that option of the dictionary;
  - basic: the 796 words of the four books that were only in the extended list move to the basic list; every form of the four
    books that an earlier check found marked goes in; an untested form goes in when Hspell does not know it and it is in at
    least two of the four copies or occurs at least 3 times (a form of one copy, seen once or twice, may be a typo of that copy);
  - extended: the remaining forms (marked, or untested and unknown to Hspell), best first by the red marks each removes
    (occurrences in all the libraries x the chance Word marks it), until the biggest edition reaches the target.
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

CONNECTORS = ['ד', 'וד', 'כד', 'וכד', 'מד', 'ומד']
FIRST = 'אבגדהוזחטיכלמנסעפצקרשת'
LAST = 'אבגדהוזחטיךלםןסעףץקרשת'


def acronyms(n):
    return [''.join(h) + '"' + l for h in itertools.product(FIRST, repeat=n - 1) for l in LAST]


acro = set(acronyms(2) + acronyms(3))
comp = set(c + b for b in base for c in CONNECTORS)
LETTER = re.compile('[א-ת]')


FINAL_INSIDE = re.compile('[ךםןףץ][א-ת]')
THRICE = re.compile(r'([א-ת])\1\1')


def no_fragment(w):
    """a whole form: at least 3 letters (2 with a quote mark: an abbreviation), a final letter only at the end (otherwise it is
    a typo such as ןאם, לםרבה), no letter three times in a row (קיייץ, אאא), at most one quote mark (not הרא"ש')"""
    n = len(LETTER.findall(w))
    quotes = w.count('"') + w.count("'")
    if quotes > 1 or FINAL_INSIDE.search(w) or THRICE.search(w):
        return False
    return n >= 3 or (n >= 2 and quotes == 1)


def total(b, m, with_comp):
    s = set(b) | acro | set(m)
    if with_comp:
        s |= comp
    return len(s)


have = set(basic) | set(more) | acro | comp
print('before: basic %d | extended %d | extended + chibur %d' % (total(basic, [], False), total(basic, more, False), total(basic, more, True)))

# 1. the four books
moved = [l.split('\t')[0] for l in read_words(os.path.join(st5, 'move_to_basic.tsv'))[1:]]
moved = [w for w in moved if w in set(more)]
rows = [l.split('\t') for l in read_words(os.path.join(st5, 'missing_from_basic.tsv'))[1:]]
untested = [r[0] for r in rows if r[1] == 'untested']
hs = dict(zip(untested, hspell_flags(untested)))
add_basic = []
for w, status, copies, occ, books in rows:
    if w in have or w in comp or not no_fragment(w):
        continue
    if status == 'marked' or (status == 'untested' and not hs[w] and (int(copies) >= 2 or int(occ) >= 3)):
        add_basic.append((w, status, int(copies), int(occ)))
basic_set = set(basic) | set(moved) | set(w for w, _, _, _ in add_basic)
more_set = set(more) - set(moved)

# 2. the libraries
cands = []
with open(os.path.join(st5, 'candidates.tsv'), encoding='utf-8') as f:
    f.readline()
    for line in f:
        w, st, src, occ, files, hk, ch, onlynew, names = line.rstrip('\n').split('\t')
        if ch == '1' or w in have or w in basic_set or not no_fragment(w):
            continue
        if st == 'marked':
            p = 1.0
        elif st == 'untested' and hk == '0':
            p = 0.82
        else:
            continue
        cands.append((int(occ) * p, int(src), int(files), w, st, int(occ), 'four' in names.split(',')))
cands.sort(key=lambda r: (-r[0], -r[1], -r[2], r[3]))

now = total(basic_set, more_set, True)
room = target - now
chosen = cands[:max(0, room)]
# the forms already counted (none of them is in the editions) - one exact count to be sure
while total(basic_set, more_set | set(r[3] for r in chosen), True) > target:
    chosen = chosen[:-1]
# a chosen form that also appears in the four books belongs to the basic edition (the user: every word of Tanakh, Mishna and Talmud
# in the basic one) - e.g. a form of one copy of the four books, rare there, that the other libraries confirm
to_basic = [r for r in chosen if r[6]]
basic_set |= set(r[3] for r in to_basic)
for r in to_basic:
    add_basic.append((r[3], r[4] + '+libraries', 0, r[5]))
more_set |= set(r[3] for r in chosen if not r[6])

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
        if not r[6]:
            f.write('%d\t%s\t%s\t%d\t%d\t%d\t%.1f\n' % (i, r[3], r[4], r[1], r[2], r[5], r[0]))

marked_b = sum(1 for r in add_basic if r[1] == 'marked')
marked_e = sum(1 for r in chosen if r[4] == 'marked' and not r[6])
cov_all = sum(r[0] for r in cands)
cov = sum(r[0] for r in chosen)
print('moved from extended to basic: %d' % len(moved))
print('added to basic: %d (marked by Word before %d, untested and unknown to Hspell %d, of them %d from the four books that the libraries confirm)'
      % (len(add_basic), marked_b, len(add_basic) - marked_b, len(to_basic)))
n_ext = len(chosen) - len(to_basic)
print('added to extended: %d of %d candidates (marked before %d, untested %d) - with the four-book ones, %.1f%% of the expected red marks of all candidates'
      % (n_ext, len(cands), marked_e, n_ext - marked_e, 100.0 * cov / cov_all if cov_all else 0))
print('after: basic %d | basic + chibur %d | extended %d | extended + chibur %d'
      % (total(basic_new, [], False), total(basic_new, [], True), total(basic_new, more_new, False), total(basic_new, more_new, True)))
print('lists: basic_list.txt %d lines, extended_only_list.txt %d lines' % (len(basic_new), len(more_new)))
print('last ones added: ' + ' '.join(r[3] for r in chosen[-12:]))
