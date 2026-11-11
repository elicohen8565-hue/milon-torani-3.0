"""Version 3.1.0.1 (11-11-2026): the words that "chibur torani" builds on (data/hebrew_base.txt), checked against the Torah library.

usage: python3 -I rebuild_base.py <repo> <lib_all.tsv>
lib_all.tsv: every form of the Otzaria main library - form <TAB> times <TAB> books (the scan of stage 1; in the developer's
workspace, not in the repository). Run after drop_typos.py; running it again changes nothing (it reads its own first input from
<repo>/tools/stage1/hebrew_base_3.1.0.1.txt, the list as it was).

"chibur torani" puts the connectors ד, וד, כד, וכד, מד, ומד before each word of the list. The list came from a frequency list of
today's internet Hebrew, so (the owner, 11-11-2026: "check that there is no nonsense there and that nothing is missing"):
- 988 of its words are hardly ever written in Torah books (in fewer than 20 of the ~9,800 books): פייסבוק, הקואליציה, כדורגל,
  אוגוסט, מייקל ... nobody writes דהקואליציה in a kuntres. They leave.
- 282 are a word with ו or ש already before it (ולכן, שהיא): a connector does not come before those (דולכן). Of the words of
  this kind, only those whose ד form the library itself writes (10 times or more, and once per 2,000 times of the word) stay (ודאי - דודאי).
- missing: the words that Torah books use most were not in an internet list. Added: every word that 1,200 books or more have,
  whose ד form the library writes 10 times or more (so it is a word that takes a connector), and that is not itself a word with
  ו, ש or ד before it (דלא, והא): הגמרא, קושיא, סברא, מצוה, פליגי, בירושלמי ...
The forms that are no longer made go to data/retired_list.txt."""
import os
import re
import sys

repo, lib_path = sys.argv[1], sys.argv[2]
here = os.path.dirname(os.path.abspath(__file__))
data = os.path.join(repo, 'data')
CONNECTORS = ['ד', 'וד', 'כד', 'וכד', 'מד', 'ומד']   # as Core.Connectors
MIN_BOOKS, ADD_BOOKS, MIN_D = 20, 1200, 10


def read(p):
    with open(p, encoding='utf-8') as f:
        return [l.strip() for l in f if l.strip()]


def write(p, words):
    with open(p, 'w', encoding='utf-8', newline='\n') as f:
        f.write(''.join(w + '\n' for w in words))


old = read(os.path.join(here, 'hebrew_base_3.1.0.1.txt'))
olds = set(old)
times, books = {}, {}
with open(lib_path, encoding='utf-8') as f:
    for line in f:
        w, c, b = line.rstrip('\n').split('\t')
        times[w], books[w] = int(c), int(b)


def known(x):
    return x in olds or books.get(x, 0) >= 300


def prefixed(w):   # a word with ו, ש or ד before it
    return len(w) >= 3 and w[0] in 'ושד' and known(w[1:])


def takes_d(w):
    return times.get('ד' + w, 0) >= MIN_D


def takes_d_often(w):   # for a word that looks prefixed: the ד form at least once per 2,000 times of the word (דולא is a typo, דודאי is not)
    return takes_d(w) and times.get('ד' + w, 0) * 2000 >= times.get(w, 0)


rare = [w for w in old if books.get(w, 0) < MIN_BOOKS]
pre = [w for w in old if books.get(w, 0) >= MIN_BOOKS and prefixed(w) and not takes_d_often(w)]
keep = [w for w in old if books.get(w, 0) >= MIN_BOOKS and not (prefixed(w) and not takes_d_often(w))]
ok = re.compile(r'^[א-ת]{3,}$')
add = sorted((w for w in books if books[w] >= ADD_BOOKS and ok.match(w) and w not in olds and w[-1] not in 'כמנפצ'
              and not prefixed(w) and takes_d(w)), key=lambda w: (-books[w], w))
new = keep + add
assert len(new) == len(set(new))

retired = read(os.path.join(data, 'retired_list.txt'))
made = {c + w for w in new for c in CONNECTORS}
# every form that is no longer made is retired - also one that the extended list has by itself: a person with the basic edition
# loses it, and a retired word that is still in the edition a person has is simply a word of the dictionary (Core.ForeignWords)
gone = sorted({c + w for w in rare + pre for c in CONNECTORS} - made)
write(os.path.join(data, 'hebrew_base.txt'), new)
new_retired = sorted(set(retired) | set(gone))
write(os.path.join(data, 'retired_list.txt'), new_retired)
print('base: %d -> %d  (out: %d hardly written in Torah books, %d with a prefix; in: %d)' % (len(old), len(new), len(rare), len(pre), len(add)))
print('forms no longer made: %d; retired list: %d words' % (len(gone), len(new_retired)))
