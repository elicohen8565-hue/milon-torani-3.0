"""Join the outputs of scan_sources.py (chunks, in order of priority) into one file for the program (9-11-2026).
usage: python3 -I merge_sources.py [--words <words.txt>] <out.txt> <chunk1.tsv> [<chunk2.tsv> ...]
  --words: keep only these words (one per line: every word of the biggest edition, see make_lists.py --targets-only)
out.txt (UTF-8, \n):
  #B <n>      then n lines: book name <TAB> collection code (0-3 core, - = other) <TAB> library
  #H <n>      then n lines: heading
  #L <n>      then n lines: label (a verse / mishna mark, or a line number)
  #W <n>      then n lines, sorted by code point: word <TAB> total <TAB> core <TAB> others <TAB> number of other books
              core   = entries "c:count:place,place" joined by ";" (c = 0 Tanakh, 1 Mishna, 2 Bavli, 3 Yerushalmi), each
                       from one copy (Otzaria's, or the other copy when Otzaria's does not have the word)
              others = entries "book:count:place" joined by ";" - the 3 books where the word appears most, most first (books of
              the same name - a book in several files, or its copies in several libraries - are one entry)
              place  = book.heading.label.word number (indexes into the tables above)
Only the books, headings and labels that a kept place uses are written (renumbered)."""
import re
import sys

CORE = {'תנ"ך': 0, 'משנה': 1, 'תלמוד בבלי': 2, 'תלמוד ירושלמי': 3}
args = sys.argv[1:]
keep = None
if args[:1] == ['--words']:
    keep = set(l.strip() for l in open(args[1], encoding='utf-8') if l.strip())
    args = args[2:]
out_path, chunks = args[0], args[1:]
rec = {}     # word -> [total, {c: [count, [place]]}, [[book, count, place]], nother]   (places as global tuples)
gbooks, gbook_ix = [], {}
gheads, ghead_ix = [], {}
glabels, glabel_ix = [], {}


LEAD = re.compile(r'^\(([^()]*)\)\s*(.+)$')


def display_name(n):
    """'(דרשות) בית דוד' -> 'בית דוד (דרשות)': the name starts with the book, not with a note in parentheses (Dicta names 9 books
    so, and has one of them under both names)"""
    m = LEAD.match(n)
    return '%s (%s)' % (m.group(2), m.group(1)) if m else n


def keep_top(lst, ent, n=3):
    """lst keeps the n entries with the biggest count (ent[1]); of equal counts, the one that came first stays"""
    for e in lst:
        if gbooks[e[0]][0] == gbooks[ent[0]][0]:   # a book of the same name: one entry, its first place kept
            if e[0] == ent[0]:
                e[1] += ent[1]                     # more files of the same book in the same library: added up
            else:
                e[1] = max(e[1], ent[1])           # its copy in another library: the same text, not more occurrences
            return
    if len(lst) < n:
        lst.append(ent)
        return
    j = 0
    for k in range(1, len(lst)):
        if lst[k][1] <= lst[j][1]:
            j = k
    if ent[1] > lst[j][1]:
        del lst[j]          # the list stays in the order things came, so "the last of the smallest" is the latest of them
        lst.append(ent)


def gid(table, ix, key):
    i = ix.get(key)
    if i is None:
        i = ix[key] = len(table)
        table.append(key)
    return i


for path in chunks:
    with open(path, encoding='utf-8') as f:
        def section(tag):
            line = f.readline().rstrip('\n')
            assert line.startswith(tag), (path, tag, line[:40])
            return int(line.split('\t')[1])
        books = [tuple(f.readline().rstrip('\n').split('\t')) for _ in range(section('#BOOKS'))]
        heads = [f.readline().rstrip('\n') for _ in range(section('#HEADS'))]
        labels = [f.readline().rstrip('\n') for _ in range(section('#LABELS'))]
        bmap = [gid(gbooks, gbook_ix, (display_name(b[0]), str(CORE[b[1]]) if b[1] in CORE else '-', b[2])) for b in books]
        hmap = [gid(gheads, ghead_ix, h) for h in heads]
        lmap = [gid(glabels, glabel_ix, l) for l in labels]

        def place(s):
            b, h, l, p = s.split('.')
            return (bmap[int(b)], hmap[int(h)], lmap[int(l)], int(p))

        n = section('#WORDS')
        for _ in range(n):
            w, total, core, other, nother = f.readline().rstrip('\n').split('\t')
            if keep is not None and w not in keep:
                continue
            r = rec.get(w)
            if r is None:
                r = rec[w] = [0, {}, [], 0]
            r[0] += int(total)
            if core:
                for ent in core.split(';'):
                    coll, cnt, places = ent.split(':', 2)
                    name, lib = coll.split('|', 1) if '|' in coll else (coll, '')
                    key = (CORE[name], lib)          # each copy apart; one is chosen below
                    e = r[1].get(key)
                    if e is None:
                        e = r[1][key] = [0, []]
                    e[0] += int(cnt)
                    for p in places.split(','):
                        if p and len(e[1]) < 2:
                            e[1].append(place(p))
            if other:
                for ent in other.split(';'):
                    b, cnt, p = ent.split(':', 2)
                    keep_top(r[2], [bmap[int(b)], int(cnt), place(p)])
            r[3] += int(nother)
    print('read', path, 'words so far', len(rec), flush=True)

# only what is used, renumbered in order of first use
used_b, used_h, used_l = {}, {}, {}


def use(d, i):
    j = d.get(i)
    if j is None:
        j = d[i] = len(d)
    return j


lines = []
def one_copy(core):
    """per collection one copy of the four books: Otzaria's, or else the first other copy that has the word (its places then say
    whose copy it is) - the counts and places are those of one copy, not of two copies added up"""
    res = {}
    for (c, lib), v in core.items():
        if c not in res or lib == 'אוצריא':
            res[c] = v
    return res


for w in sorted(rec):
    total, core, other, nother = rec[w]
    core = one_copy(core)
    cs = ';'.join('%d:%d:%s' % (c, v[0], ','.join('%d.%d.%d.%d' % (use(used_b, p[0]), use(used_h, p[1]), use(used_l, p[2]), p[3]) for p in v[1]))
                  for c, v in sorted(core.items()))
    os_ = ';'.join('%d:%d:%d.%d.%d.%d' % (use(used_b, b), cnt, use(used_b, p[0]), use(used_h, p[1]), use(used_l, p[2]), p[3])
                   for b, cnt, p in sorted(other, key=lambda e: -e[1]))   # most first (stable: ties keep their order)
    lines.append('%s\t%d\t%s\t%s\t%d' % (w, total, cs, os_, nother))


def ordered(d, table):
    res = [None] * len(d)
    for i, j in d.items():
        res[j] = table[i]
    return res


with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
    bl = ordered(used_b, gbooks)
    f.write('#B\t%d\n' % len(bl))
    for b in bl:
        f.write('%s\t%s\t%s\n' % b)
    hl = ordered(used_h, gheads)
    f.write('#H\t%d\n' % len(hl))
    for h in hl:
        f.write(h + '\n')
    ll = ordered(used_l, glabels)
    f.write('#L\t%d\n' % len(ll))
    for l in ll:
        f.write(l + '\n')
    f.write('#W\t%d\n' % len(lines))
    for l in lines:
        f.write(l + '\n')
print('DONE words %d books %d heads %d labels %d' % (len(lines), len(used_b), len(used_h), len(used_l)))
