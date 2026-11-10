"""Version 3.1.0.2 (11-11-2026): the words of the extended list that one book alone has, fewer than 3 times.

usage: python3 -I drop_single_book.py <repo>
Run after drop_typos.py. Takes the words of single_book_words.txt (next to this file; made by tools/sources/dedup_count.py from
the count per book, recount_books.py) out of <repo>/data/extended_only_list.txt and adds them to <repo>/data/retired_list.txt.
Running it again changes nothing.

Why: the rule of the extended list is "written exactly so at least 3 times", and the count added up every library - but one book
often sits in two or three of them, so a word written once in one book was counted 3 times. The owner: "of course a book enters
only once"; a word that two different books have stays, a word that only one book has (once or twice) goes."""
import os
import sys


def read(p):
    with open(p, encoding='utf-8') as f:
        return [l.strip() for l in f if l.strip()]


def write(p, words):
    with open(p, 'w', encoding='utf-8', newline='\n') as f:
        f.write(''.join(w + '\n' for w in words))


repo = sys.argv[1]
data = os.path.join(repo, 'data')
go = set(read(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'single_book_words.txt')))
more = read(os.path.join(data, 'extended_only_list.txt'))
basic = set(read(os.path.join(data, 'basic_list.txt')))
retired = read(os.path.join(data, 'retired_list.txt'))
assert not (go & basic), 'a word of the basic list is not judged by this rule'
out = [w for w in more if w in go]
assert not (go - set(out) - set(retired)), 'a word of the list is neither in the extended list nor retired'
write(os.path.join(data, 'extended_only_list.txt'), [w for w in more if w not in go])
new_retired = sorted(set(retired) | go)
write(os.path.join(data, 'retired_list.txt'), new_retired)
print('single-book words taken out: %d (retired list: %d words)' % (len(out), len(new_retired)))
