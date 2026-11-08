"""Build the embedded word list (pool) of WordCheck 4: words from the books that are NOT in the Torah dictionary, most valuable first.

usage: python3 -I make_pool4.py <candidates.tsv> <pool3.txt.gz> <dump folder> <how many> <out pool4.txt.gz> <out pool4_index.tsv>

  candidates.tsv  stage5/rank_new.py: every form of the scanned libraries that the dictionary lacks (2-12 letters, with real
                  evidence), with what earlier Word checks said (marked / accepted / untested), whether Hspell knows it,
                  whether it is a "chibur torani" form, and in which sources it appears.
  pool3.txt.gz    WordCheck 3's list: its 13 control words (@C) and speed-test words (@E, @A) are used again as they are.
  dump folder     DevDump.exe output: the four editions (basic.dic, basic_comp.dic, extended.dic, extended_comp.dic). No word of
                  the list may be in any of them.
  how many        how many untested forms to take, best first.

Which forms: untested by Word, not "chibur torani" forms (that edition option covers them). Order = the expected number of red
marks the form removes: its occurrences in all the libraries x the chance that Word marks it (measured before: 82% of the forms
Hspell does not know, 2.5% of the forms it knows). Then every untested form of the four books (Tanakh, Mishna, Bavli,
Yerushalmi) that is not in the list yet: those decide the basic edition.

Lines of the pool (gzip, UTF-8), as WordCheck 3 reads them:
  @C<TAB>word   control words: dictionary words that Word does not know by itself (they show that the dictionary is disconnected)
  @D<TAB>word   the forms Word is asked about, best first
  @E<TAB>word   speed test, words that Word marks;  @A<TAB>word  speed test, words that Word accepts
"""
import gzip
import os
import sys

cand_path, pool3_path, dump_dir, how_many, out_path, index_path = sys.argv[1:7]
how_many = int(how_many)

old = {'@C': [], '@D': [], '@E': [], '@A': []}
for line in gzip.open(pool3_path, 'rt', encoding='utf-8'):
    line = line.rstrip('\n')
    if line[:2] in old and line[2:3] == '\t':
        old[line[:2]].append(line[3:])
assert len(old['@C']) == 13 and len(old['@E']) == 1000 and len(old['@A']) == 1000, 'unexpected pool3'


def dic(p):
    return set(w for w in open(p, encoding='utf-16').read().splitlines() if w.strip())


editions = set()
for name in ('basic.dic', 'basic_comp.dic', 'extended.dic', 'extended_comp.dic'):
    editions |= dic(os.path.join(dump_dir, name))
assert all(c in editions for c in old['@C']), 'a control word is not in the dictionary'
speed = set(old['@E']) | set(old['@A'])

rows = []
with open(cand_path, encoding='utf-8') as f:
    head = f.readline().rstrip('\n').split('\t')
    assert head[:9] == ['word', 'word_check', 'sources', 'occurrences', 'files', 'hspell_knows', 'in_chibur',
                        'only_new_libraries', 'source_names'], head
    for line in f:
        w, st, src, occ, files, hk, ch, onlynew, names = line.rstrip('\n').split('\t')
        if st != 'untested' or ch == '1':
            continue
        if w in editions or w in speed or any(c in w for c in '\t\r\n '):
            continue
        hk = hk == '1'
        weight = int(occ) * (0.025 if hk else 0.82)
        rows.append((weight, int(src), int(files), w, int(occ), hk, 'four' in names.split(',')))

rows.sort(key=lambda r: (-r[0], -r[1], -r[2], r[3]))
chosen = rows[:how_many]
have = set(r[3] for r in chosen)
four_extra = [r for r in rows[how_many:] if r[6]]
chosen += four_extra
words = [r[3] for r in chosen]
assert len(words) == len(set(words)), 'a form twice'

lines = ['@C\t' + c for c in old['@C']] + ['@D\t' + w for w in words] + ['@E\t' + e for e in old['@E']] + ['@A\t' + a for a in old['@A']]
with gzip.open(out_path, 'wb', 9) as f:
    f.write(('\n'.join(lines) + '\n').encode('utf-8'))
with open(index_path, 'w', encoding='utf-8') as f:
    f.write('rank\tword\tsources\toccurrences\tfiles\thspell_knows\tfour_books\texpected_red\n')
    for i, r in enumerate(chosen):
        f.write('%d\t%s\t%d\t%d\t%d\t%d\t%d\t%.1f\n' % (i, r[3], r[1], r[4], r[2], int(r[5]), int(r[6]), r[0]))

n_four = sum(1 for r in chosen if r[6])
print('untested forms (not chibur, not in any edition): %d' % len(rows))
print('pool: %d forms = the best %d + %d more of the four books (four books in all: %d) | Hspell does not know %d'
      % (len(words), min(how_many, len(rows)), len(four_extra), n_four, sum(1 for r in chosen if not r[5])))
print('controls %d | speed test %d + %d' % (len(old['@C']), len(old['@E']), len(old['@A'])))
print('first: ' + ' '.join(words[:25]))
print('last of the best: ' + ' '.join(words[how_many - 10:how_many]))
