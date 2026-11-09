"""How often each word appears in the libraries AS WRITTEN (3.1, after the 4th independent check of the categories, 9-11-2026).

usage: python3 -I recount.py <words.txt> <list.tsv> <out.tsv>

The sources scan (scan_sources.py, like every earlier scan) counts the main text only - tokens.clean_line drops footnotes, <sup>
and headings - and through tokens.norm_token, which joins a word across an inner quote mark: "ז"לכי" (ז"ל and כי run together in
one book) is counted as "זלכי", a form no book has. The independent check found both: real words that appear 3 times or more only
when the footnotes are read (מיקרותם, 10 times) were dropped, and forms made by the joining (זלכי, גלק) were kept.

This count, for the rule "at least 3 times in all the libraries":
- every line of every file of list.tsv - also footnotes, <sup> and headings, which are text of the book too (the 9 header lines of
  a Sefaria file are skipped);
- a footnote or a <sup> is cut off from the word before it by a space; every other tag is removed (a tag inside a word does not
  split the word), as in the scan;
- a word counts only as written: the token, with ׳ and ״ read as ' and ", must equal the word exactly.
out.tsv: word <TAB> count, for every word of words.txt that was found at least once."""
import html
import os
import re
import subprocess
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))   # tokens.py next to this file
import tokens  # noqa: E402

words_path, list_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
targets = set(l.strip() for l in open(words_path, encoding='utf-8') if l.strip())
RE_NOTE = re.compile(r'<i class="footnote">(.*?)</i>|<sup[^>]*>(.*?)</sup>', re.S)


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
    """path -> blob id of every file of HEAD"""
    out = subprocess.run(['git', '-C', repo, 'ls-tree', '-r', '-z', 'HEAD'], stdout=subprocess.PIPE, check=True).stdout
    res = {}
    for ent in out.split(b'\0'):
        if ent:
            meta, path = ent.split(b'\t', 1)
            res[path.decode('utf-8')] = meta.split()[2].decode()
    return res


count = {}
trees, readers = {}, {}
done = missing = 0
env = dict(os.environ, GIT_NO_LAZY_FETCH='1')
for row in (l.rstrip('\n').split('\t') for l in open(list_path, encoding='utf-8') if l.strip()):
    repo, path, fmt = row[0], row[1], row[5]
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
    lines = data.decode('utf-8', errors='replace').split('\n')
    if fmt == 'sef':
        lines = lines[9:]
    for raw in lines:
        for tok in tokens.RE_TOKEN.findall(clean(raw)):
            t = tok.replace('׳', "'").replace('״', '"')
            if t in targets:
                count[t] = count.get(t, 0) + 1
    done += 1
    if done % 500 == 0:
        print('files %d (missing %d) words found %d' % (done, missing, len(count)), flush=True)

with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
    for w in sorted(count):
        f.write('%s\t%d\n' % (w, count[w]))
print('DONE files %d missing %d words %d' % (done, missing, len(count)))
