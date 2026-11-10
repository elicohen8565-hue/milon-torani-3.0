"""How often each word appears AS WRITTEN in each book (3.1.0.2, 11-11-2026) - for the rule "a book counts once".

usage: python3 -I recount_books.py <words.txt> <list.tsv> <out.tsv>

recount.py adds up every file of every library, and one book often sits in two or three libraries (Otzaria, Sefaria, Dicta,
Orayta ...): "וביהידות" is written once, in Pele Yoetz, and was counted 3 times. The owner (11-11-2026, 01:05): "of course a book
enters only once". This is recount.py (the same text, the same cutting of words), but the count is kept per book: the name of
the book as scan_sources.py takes it (the <h1> of an Otzaria file, the title line of a Sefaria file) and its library.
out.tsv: word <TAB> book <TAB> library <TAB> count. dedup_count.py reads the six outputs.
Give it only the words in doubt (a word in four books or more needs no count)."""
import html
import os
import re
import subprocess
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tokens  # noqa: E402

words_path, list_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
targets = set(l.strip() for l in open(words_path, encoding='utf-8') if l.strip())
RE_NOTE = re.compile(r'<i class="footnote">(.*?)</i>|<sup[^>]*>(.*?)</sup>', re.S)
RE_TAGS = re.compile(r'<[^>]+>')


def clean(line):
    s = RE_NOTE.sub(lambda m: ' ' + (m.group(1) or m.group(2) or '') + ' ', line)
    s = tokens.RE_BR.sub(' ', s)
    s = tokens.RE_TAG.sub('', s)
    s = html.unescape(s)
    s = tokens.RE_CURLY.sub(' ', s)
    s = tokens.RE_LEAD.sub('', s)
    s = unicodedata.normalize('NFKD', s)
    s = s.translate(tokens.SPECIAL)
    return tokens.RE_MARKS.sub('', s)


def tree(repo):
    out = subprocess.run(['git', '-C', repo, 'ls-tree', '-r', '-z', 'HEAD'], stdout=subprocess.PIPE, check=True).stdout
    res = {}
    for ent in out.split(b'\0'):
        if ent:
            meta, path = ent.split(b'\t', 1)
            res[path.decode('utf-8')] = meta.split()[2].decode()
    return res


count = {}     # (word, book, library) -> count
trees, readers = {}, {}
done = missing = 0
env = dict(os.environ, GIT_NO_LAZY_FETCH='1')
for row in (l.rstrip('\n').split('\t') for l in open(list_path, encoding='utf-8') if l.strip()):
    repo, path, library, fmt = row[0], row[1], row[4], row[5]
    if repo not in trees:
        trees[repo] = tree(repo)
        readers[repo] = subprocess.Popen(['git', '-C', repo, 'cat-file', '--batch'], stdin=subprocess.PIPE, stdout=subprocess.PIPE, env=env)
    sha = trees[repo].get(path)
    if sha is None:
        missing += 1
        continue
    rd = readers[repo]
    rd.stdin.write((sha + '\n').encode())
    rd.stdin.flush()
    header = rd.stdout.readline().decode().split()
    if len(header) < 3 or header[1] != 'blob':
        missing += 1
        continue
    data = rd.stdout.read(int(header[2]))
    rd.stdout.read(1)
    text = data.decode('utf-8', errors='replace')
    name = os.path.splitext(os.path.basename(path))[0]       # (the name of the book: as in scan_sources.py)
    if fmt == 'otz':
        m = re.search(r'<h1[^>]*>(.*?)</h1>', text[:2000], re.S)
        if m:
            t = RE_TAGS.sub('', m.group(1)).strip()
            if t:
                name = t[:80]
    else:
        parts = text.split('\n', 3)
        if len(parts) > 1 and parts[1].strip():
            name = parts[1].strip()[:80]
    lines = text.split('\n')
    if fmt == 'sef':
        lines = lines[9:] if len(lines) > 9 else []
    for raw in lines:
        for tok in tokens.RE_TOKEN.findall(clean(raw)):
            t = tok.replace('׳', "'").replace('״', '"')
            if t in targets:
                k = (t, name, library)
                count[k] = count.get(k, 0) + 1
    done += 1
    if done % 500 == 0:
        print('files %d (missing %d) entries %d' % (done, missing, len(count)), flush=True)

with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
    for k in sorted(count):
        f.write('%s\t%s\t%s\t%d\n' % (k[0], k[1].replace('\t', ' '), k[2], count[k]))
print('DONE files %d missing %d entries %d' % (done, missing, len(count)))
