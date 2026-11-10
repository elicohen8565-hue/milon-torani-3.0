"""Which words of the extended list only ONE book has, fewer than 3 times (3.1.0.2, 11-11-2026).

usage: python3 -I dedup_count.py <words.txt> <out.txt> <recount_books output> [<recount_books output> ...]

The owner (11-11-2026): a book counts once, however many libraries carry it; and of the words that are then written fewer than 3
times, those that two different books have stay (two authors wrote the form), those that one book alone has go (most are typos of
that one digital copy: "וביהידות", once, in Pele Yoetz, which sits in three libraries).
A book = its name (recount_books.py). The files of one name in one library are added up (a book in several files); between
libraries the largest count is taken (copies of one book). A copy under another name counts as another book - the word stays.
words.txt = the words that were counted (a word that is not in it is in many books and is not judged).
out.txt: the words that go, sorted; printed: how many of each kind."""
import collections
import sys

words_path, out_path, outs = sys.argv[1], sys.argv[2], sys.argv[3:]
targets = [l.strip() for l in open(words_path, encoding='utf-8') if l.strip()]
per = {}   # word -> {book: {library: count}}
for p in outs:
    for line in open(p, encoding='utf-8'):
        w, book, lib, n = line.rstrip('\n').split('\t')
        d = per.setdefault(w, {}).setdefault(book, {})
        d[lib] = d.get(lib, 0) + int(n)
kinds = collections.Counter()
go = []
for w in targets:
    books = per.get(w, {})
    once = sum(max(c.values()) for c in books.values())
    if once >= 3:
        kinds['3 times or more, a book counted once'] += 1
    elif len(books) >= 2:
        kinds['fewer than 3, in two books (stays)'] += 1
    else:
        kinds['fewer than 3, in one book: %d time(s) (goes)' % once] += 1
        go.append(w)
with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
    f.write(''.join(w + '\n' for w in sorted(go)))
for k in sorted(kinds):
    print('%7d  %s' % (kinds[k], k))
print('%7d  words go' % len(go))
