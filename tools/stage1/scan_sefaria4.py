"""Count the words of the four basic books in the Sefaria copy (Tanakh: Torah/Prophets/Writings; Mishnah: the six orders;
Bavli and Yerushalmi: the tractates, without commentaries), with the same tokens.py as every other scan.
Output: out/ext_sefaria4.tsv  word <TAB> tanakh <TAB> mishnah <TAB> bavli <TAB> yerushalmi <TAB> files"""
import os, re, sys, collections
sys.path.insert(0, '/home/claude/work/scripts')
import tokens
ROOT = '/home/claude/work/ext/TorahBibleCodes_Sefaria-Export/'
paths = [l.rstrip('\n') for l in open('/home/claude/work/ext/TorahBibleCodes_Sefaria-Export_paths.txt', encoding='utf-8')]
groups = [('tanakh', re.compile(r'^txt/Tanakh/(Torah|Prophets|Writings)/[^/]+/Hebrew/merged\.txt$')),
          ('mishnah', re.compile(r'^txt/Mishnah/Seder [^/]+/[^/]+/Hebrew/merged\.txt$')),
          ('bavli', re.compile(r'^txt/Talmud/Bavli/Seder [^/]+/[^/]+/Hebrew/merged\.txt$')),
          ('yerushalmi', re.compile(r'^txt/Talmud/Yerushalmi/Seder [^/]+/[^/]+/Hebrew/merged\.txt$'))]
counts = collections.defaultdict(lambda: [0, 0, 0, 0, 0])
nfiles = collections.Counter()
missing = 0
for p in paths:
    for gi, (g, rx) in enumerate(groups):
        if rx.match(p):
            full = ROOT + p
            if not os.path.exists(full):
                missing += 1
                break
            nfiles[g] += 1
            seen = set()
            for tok in tokens.file_tokens(full):
                c = counts[tok]
                c[gi] += 1
                if tok not in seen:
                    seen.add(tok); c[4] += 1
            break
print('files:', dict(nfiles), '| not checked out:', missing, '| distinct forms:', len(counts))
for gi, (g, rx) in enumerate(groups):
    print('  %-10s occurrences=%9d distinct=%7d' % (g, sum(v[gi] for v in counts.values()), sum(1 for v in counts.values() if v[gi])))
with open('/home/claude/work/out/ext_sefaria4.tsv', 'w', encoding='utf-8') as f:
    for w in sorted(counts):
        c = counts[w]
        f.write('%s\t%d\t%d\t%d\t%d\t%d\n' % (w, c[0], c[1], c[2], c[3], c[4]))
