"""Is the split between the basic and the extended edition right? Basic = Tanakh, Mishnah, Bavli, Yerushalmi.
Four independent copies of the four books: Otzaria (raw_counts_v2.tsv), Sefaria (ext_sefaria4.tsv), Mechon Mamre Tanakh and Bavli.
1. extended-only words that appear in the four books  -> belong in basic
2. words of the four books that are in no edition and Word marks them  -> missing from basic
3. words of the basic list that appear in none of the four books  -> perhaps belong in extended
usage: python3 -I split_check.py"""
import collections, re, sys
O = '/home/claude/work/out/'
DUMP = '/home/claude/milon-torani-3.0/build/dump/'
DATA = '/home/claude/milon-torani-3.0/data/'

def dic(p): return [w for w in open(p, encoding='utf-16').read().splitlines() if w.strip()]
basic_ed = dic(DUMP + 'basic.dic'); ext_ed = dic(DUMP + 'extended.dic')
B, D = set(basic_ed), set(ext_ed)
E = D - B
basic_list = [l.strip() for l in open(DATA + 'basic_list.txt', encoding='utf-8') if l.strip()]
BL = set(basic_list)
generated = B - BL                       # the 2-3 letter acronyms that the installer makes
print('basic edition %d (list %d + generated acronyms %d) | extended edition %d | extended-only %d' % (len(B), len(BL), len(generated), len(D), len(E)))

K, sec = set(), 'K'
for line in open('/home/claude/wc3/an/out/word_knows.txt', encoding='utf-8'):
    line = line.rstrip('\n')
    if line.startswith('#H:'): sec = 'H'; continue
    if line.startswith('#') or not line: continue
    if sec == 'K': K.add(line)

BOOKS = ['tanakh', 'mishnah', 'bavli', 'yerushalmi']
src = {}   # name -> {word: [per book counts]}
def load4(path):
    d = {}
    for line in open(path, encoding='utf-8'):
        p = line.rstrip('\n').split('\t'); d[p[0]] = [int(x) for x in p[1:5]]
    return d
src['otzaria'] = load4(O + 'raw_counts_v2.tsv')
src['sefaria'] = load4(O + 'ext_sefaria4.tsv')
for name, book in (('mamre_tanakh', 0), ('mamre_bavli', 2)):
    d = {}
    for line in open(O + 'ext_%s.tsv' % name, encoding='utf-8'):
        w, a, b = line.rstrip('\n').split('\t'); v = [0, 0, 0, 0]; v[book] = int(a); d[w] = v
    src[name] = d
ALL4 = set().union(*[set(d) for d in src.values()])
def nsrc(w): return sum(1 for d in src.values() if w in d)
def occ(w): return sum(sum(d[w]) for d in src.values() if w in d)
def books(w):
    s = set()
    for d in src.values():
        if w in d:
            for i, c in enumerate(d[w]):
                if c: s.add(BOOKS[i])
    return s
print('forms in the four books: ' + ', '.join('%s %d' % (n, len(d)) for n, d in src.items()) + ' | all together %d' % len(ALL4))

# 1 ---------------------------------------------------------------------------------------------------------
move = sorted(E & ALL4, key=lambda w: (-nsrc(w), -occ(w)))
print('\n1. EXTENDED-ONLY words that appear in the four books: %d (of %d extended-only)' % (len(move), len(E)))
by_n = collections.Counter(nsrc(w) for w in move)
print('   in how many of the 4 copies:', dict(sorted(by_n.items())))
by_book = collections.Counter(b for w in move for b in books(w))
print('   in which books:', dict(by_book))
print('   Word knows by itself (K):', sum(1 for w in move if w in K))
print('   in 2+ copies: %d | in 1 copy only: %d' % (sum(1 for w in move if nsrc(w) >= 2), sum(1 for w in move if nsrc(w) == 1)))
print('   examples (most copies, most frequent):', ' '.join('%s(%d/%d)' % (w, nsrc(w), occ(w)) for w in move[:30]))
one = [w for w in move if nsrc(w) == 1]
print('   examples found in one copy only:', ' '.join('%s(%d)' % (w, occ(w)) for w in sorted(one, key=lambda w: -occ(w))[:20]))
open('/home/claude/work/stage5/move_to_basic.tsv', 'w', encoding='utf-8').write(
    'word\tcopies\toccurrences\tbooks\tword_knows\n' + ''.join('%s\t%d\t%d\t%s\t%d\n' % (w, nsrc(w), occ(w), ','.join(sorted(books(w))), w in K) for w in move))

# 2 ---------------------------------------------------------------------------------------------------------
flagged = set()
for line in open(O + 'flagged_part1.tsv', encoding='utf-8'):
    flagged.add(line.rstrip('\n').split('\t')[1])
tested = set()
for line in open(O + 'pool_index.tsv', encoding='utf-8'):
    p = line.rstrip('\n').split('\t')
    if int(p[0]) < 179750: tested.add(p[1])
def nl(w): return len(re.sub(r"[\"']", '', w))
missing = [w for w in ALL4 if w not in D]
m_flag = [w for w in missing if w in flagged]
m_acc = [w for w in missing if w in tested and w not in flagged]
m_untested = [w for w in missing if w not in tested]
print('\n2. forms of the four books in NO edition: %d' % len(missing))
print('   Word marks them (tested): %d | Word accepts them (tested): %d | not tested yet: %d' % (len(m_flag), len(m_acc), len(m_untested)))
print('   marked, by copies:', dict(sorted(collections.Counter(nsrc(w) for w in m_flag).items())))
print('   marked, most frequent:', ' '.join('%s(%d/%d)' % (w, nsrc(w), occ(w)) for w in sorted(m_flag, key=lambda w: (-nsrc(w), -occ(w)))[:30]))
ut2 = [w for w in m_untested if 2 <= nl(w) <= 12]
print('   not tested, 2-12 letters: %d (in 2+ copies %d) | 1 letter or 13+: %d' % (len(ut2), sum(1 for w in ut2 if nsrc(w) >= 2), len(m_untested) - len(ut2)))
print('   not tested, most frequent:', ' '.join('%s(%d/%d)' % (w, nsrc(w), occ(w)) for w in sorted(ut2, key=lambda w: (-nsrc(w), -occ(w)))[:30]))
with open('/home/claude/work/stage5/missing_from_basic.tsv', 'w', encoding='utf-8') as f:
    f.write('word\tstatus\tcopies\toccurrences\tbooks\n')
    for w in sorted(missing, key=lambda w: (-nsrc(w), -occ(w))):
        st = 'marked' if w in flagged else ('accepted' if w in tested else 'untested')
        f.write('%s\t%s\t%d\t%d\t%s\n' % (w, st, nsrc(w), occ(w), ','.join(sorted(books(w)))))

# 3 ---------------------------------------------------------------------------------------------------------
not4 = [w for w in basic_list if w not in ALL4]
print('\n3. words of the BASIC list (not the generated acronyms) that appear in none of the four books: %d of %d' % (len(not4), len(BL)))
lib = {}
for line in open(O + 'lib_all.tsv', encoding='utf-8'):
    w, a, b = line.rstrip('\n').split('\t'); lib[w] = int(a)
inlib = [w for w in not4 if w in lib]
print('   of them in the rest of the library: %d | nowhere at all: %d | Word knows by itself: %d' % (len(inlib), len(not4) - len(inlib), sum(1 for w in not4 if w in K)))
print('   examples (most frequent in the library):', ' '.join('%s(%d)' % (w, lib[w]) for w in sorted(inlib, key=lambda w: -lib[w])[:25]))
print('   examples (nowhere):', ' '.join([w for w in not4 if w not in lib][:25]))
open('/home/claude/work/stage5/basic_not_in_4books.tsv', 'w', encoding='utf-8').write(
    'word\tlibrary_occurrences\tword_knows\n' + ''.join('%s\t%d\t%d\n' % (w, lib.get(w, 0), w in K) for w in not4))
gen4 = [w for w in generated if w in ALL4]
print('\n(the generated acronyms: %d, of them appear in the four books: %d - they stay in basic by design)' % (len(generated), len(gen4)))
