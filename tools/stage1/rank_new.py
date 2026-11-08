"""Words that the libraries bring and the dictionary lacks, with the evidence for each (how many independent sources, files).
Sources: four books (4 copies, one source), the main Otzaria library (lib_all), Sefaria (ext_sefaria) and every out/new_*.tsv.
Earlier Word checks (WordCheck part 1) tell which forms Word marks or accepts; Hspell is asked about the rest.
Output: stage5/candidates.tsv (forms with real evidence: 2+ sources or 3+ files), summary on stdout.
usage: python3 -I rank_new.py"""
import collections, glob, os, re, sys
sys.path.insert(0, '/home/claude/work/scripts')
from hs import hspell_flags
O = '/home/claude/work/out/'
DUMP = '/home/claude/milon-torani-3.0/build/dump/'
def dic(p): return set(w for w in open(p, encoding='utf-16').read().splitlines() if w.strip())
D = dic(DUMP + 'extended.dic')
COMP = dic(DUMP + 'extended_comp.dic') - D
RX = re.compile(r"[\"']")
def ok_len(w): n = len(RX.sub('', w)); return 2 <= n <= 12

names = ['four', 'library', 'sefaria'] + [os.path.basename(p)[4:-4] for p in sorted(glob.glob(O + 'new_*.tsv'))]
BIT = {n: 1 << i for i, n in enumerate(names)}
NEWMASK = sum(BIT[n] for n in names[3:])
agg = {}   # word -> [mask, occurrences, files]
def add(word, name, occ, files):
    if word in D or not ok_len(word): return
    a = agg.get(word)
    if a is None: a = agg[word] = [0, 0, 0]
    a[0] |= BIT[name]; a[1] += occ; a[2] += files
for p in ('raw_counts_v2.tsv', 'ext_sefaria4.tsv'):
    for line in open(O + p, encoding='utf-8'):
        q = line.rstrip('\n').split('\t'); add(q[0], 'four', sum(int(x) for x in q[1:5]), int(q[5]))
for p in ('ext_mamre_tanakh.tsv', 'ext_mamre_bavli.tsv'):
    for line in open(O + p, encoding='utf-8'):
        w, a_, b = line.rstrip('\n').split('\t'); add(w, 'four', int(a_), int(b))
for name, p in [('library', 'lib_all.tsv'), ('sefaria', 'ext_sefaria.tsv')] + [(n, 'new_%s.tsv' % n) for n in names[3:]]:
    for line in open(O + p, encoding='utf-8'):
        w, a_, b = line.rstrip('\n').split('\t'); add(w, name, int(a_), int(b))
def nsrc(m): return bin(m).count('1')
print('sources:', ', '.join(names))
print('forms not in the dictionary (2-12 letters), all sources together: %d' % len(agg))

flagged = set(l.split('\t')[1] for l in open(O + 'flagged_part1.tsv', encoding='utf-8'))
tested = set()
for line in open(O + 'pool_index.tsv', encoding='utf-8'):
    q = line.split('\t')
    if int(q[0]) < 179750: tested.add(q[1])

keep = [w for w, a in agg.items() if nsrc(a[0]) >= 2 or a[2] >= 3]
print('with real evidence (2+ sources or 3+ files): %d | rare (1 source, 1-2 files): %d' % (len(keep), len(agg) - len(keep)))
only_new = [w for w in keep if agg[w][0] & ~NEWMASK == 0]
print('  of them ONLY in the additional libraries: %d (in 2+ of them: %d)' % (len(only_new), sum(1 for w in only_new if nsrc(agg[w][0]) >= 2)))
gained = [w for w in keep if agg[w][0] & NEWMASK and agg[w][0] & ~NEWMASK and w not in tested]
print('  untested forms of the old sources that the additional libraries confirm: %d' % len(gained))
untested = [w for w in keep if w not in tested]
print('Hspell is asked about %d untested forms...' % len(untested))
hs = dict(zip(untested, hspell_flags(untested)))
rows = []
for w in keep:
    m, occ, files = agg[w]
    st = 'marked' if w in flagged else ('accepted' if w in tested else 'untested')
    rows.append((w, st, nsrc(m), occ, files, int(hs.get(w, False)), int(w in COMP), int(m & ~NEWMASK == 0),
                 ','.join(n for n in names if m & BIT[n])))
rows.sort(key=lambda r: (-r[2], -r[4], -r[3]))
with open('/home/claude/work/stage5/candidates.tsv', 'w', encoding='utf-8') as f:
    f.write('word\tword_check\tsources\toccurrences\tfiles\thspell_knows\tin_chibur\tonly_new_libraries\tsource_names\n')
    for r in rows: f.write('%s\t%s\t%d\t%d\t%d\t%d\t%d\t%d\t%s\n' % r)
c = collections.Counter(r[1] for r in rows)
print('by the earlier Word check:', dict(c))
ut = [r for r in rows if r[1] == 'untested']
hk = sum(r[5] for r in ut)
print('untested: %d | Hspell knows %d, does not know %d | already in chibur torani: %d' % (len(ut), hk, len(ut) - hk, sum(r[6] for r in ut)))
# expected red forms among the untested, with the rates measured before (four books: 2.5%% of Hspell-known were marked; tier 2: 82%% of Hspell-unknown)
print('expected that Word marks: about %d of the Hspell-known + about %d of the Hspell-unknown' % (round(0.025 * hk), round(0.82 * (len(ut) - hk))))
for k in (2, 3, 4, 5, 6, 8):
    sel = [r for r in ut if r[2] >= k]
    print('  untested in %d+ sources: %d (Hspell does not know %d)' % (k, len(sel), sum(1 for r in sel if not r[5])))
print('marked earlier and still missing: %d' % c['marked'])
print('examples, untested, Hspell does not know, most sources:', ' '.join('%s(%d/%d)' % (r[0], r[2], r[4]) for r in [r for r in ut if not r[5]][:40]))
print('examples, only in the additional libraries:', ' '.join('%s(%d/%d)' % (r[0], r[2], r[4]) for r in [r for r in rows if r[7]][:40]))
