"""Count words in every .txt file under several folders (one source = several folders), skipping folders marked not relevant.
usage: python3 -I scan_multi.py <out.tsv> <folder> [<folder> ...]
out.tsv: token <TAB> total occurrences <TAB> number of files that contain it"""
import os, sys
sys.path.insert(0, '/home/claude/work/scripts')
import tokens
out, roots = sys.argv[1], sys.argv[2:]
SKIP = ('לא רלוונטי',)
counts, nfiles = {}, 0
for root in roots:
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith('.') and d not in SKIP]
        for fn in sorted(filenames):
            if not fn.endswith('.txt'):
                continue
            nfiles += 1
            seen = set()
            for tok in tokens.file_tokens(os.path.join(dirpath, fn)):
                c = counts.get(tok)
                if c is None:
                    c = counts[tok] = [0, 0]
                c[0] += 1
                if tok not in seen:
                    seen.add(tok); c[1] += 1
with open(out, 'w', encoding='utf-8') as f:
    for tok in sorted(counts):
        c = counts[tok]
        f.write('%s\t%d\t%d\n' % (tok, c[0], c[1]))
print('files', nfiles, 'distinct', len(counts), 'occurrences', sum(c[0] for c in counts.values()))
