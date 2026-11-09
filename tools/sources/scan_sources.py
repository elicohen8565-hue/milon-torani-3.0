"""Where does each word of the dictionary appear? (3.1, 9-11-2026: the user wants the source of every word next to it in the
word-editing window, because the online libraries are full of typos.)

usage: python3 -I scan_sources.py <targets.txt> <list.tsv> <out.tsv>
  targets.txt: the words of the dictionary, one per line
  list.tsv:    the files to scan, in order of priority: repo <TAB> path <TAB> kind <TAB> collection <TAB> library <TAB> format
               kind = core (Tanakh / Mishna / Bavli / Yerushalmi: collection = its name) or book (collection = "")
               format = otz (Otzaria: <h1>..<h6> headings) or sef (Sefaria merged.txt)
  out.tsv:     one line per word that was found (see write() below), and the tables of books / headings / labels first

The files are read straight from git (cat-file --batch), nothing is checked out. Words are cut exactly as in every earlier scan
(tokens.py), so a word of the lists is found as it is.
For each word: per core collection and copy (collection|library) the number of occurrences and the first 2 places; the 3 other books where it appears most
(each with its number of occurrences and its first place; of books with the same number, the one scanned first); how many
other books in all; the total number of occurrences.
A place = (book, heading, label, word number); label = the verse / mishna mark for Tanakh and Mishna, else the line number
under the heading."""
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))   # tokens.py next to this file
import tokens  # noqa: E402

targets_path, list_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
targets = set(l.strip() for l in open(targets_path, encoding='utf-8') if l.strip())

RE_H = re.compile(r'^\s*<h([1-6])[^>]*>(.*?)</h\1>', re.S)
RE_TAGS = re.compile(r'<[^>]*>')
RE_MARK = re.compile(r'^\s*\(\s*([א-ת"\'׳״]{1,4})\s*\)')
RE_SEF_SECTION = re.compile(r'^(Daf|Chapter|Siman|Seif|Halakhah|Mishnah|Section|Paragraph|Verse|Volume|Part|Book|Gate|Klal|Teshuva|Letter|Page|Line|Psalm|Parasha|Sha\'ar|Ma\'amar|Introduction)\s+(\S+)\s*$')
HEB_SECTION = {'Daf': 'דף', 'Chapter': 'פרק', 'Siman': 'סימן', 'Seif': 'סעיף', 'Halakhah': 'הלכה', 'Mishnah': 'משנה',
               'Section': 'חלק', 'Paragraph': 'פסקה', 'Verse': 'פסוק', 'Volume': 'כרך', 'Part': 'חלק', 'Book': 'ספר',
               'Gate': 'שער', 'Klal': 'כלל', 'Teshuva': 'תשובה', 'Letter': 'אות', 'Page': 'עמוד', 'Line': 'שורה',
               'Psalm': 'מזמור', 'Parasha': 'פרשה', "Sha'ar": 'שער', "Ma'amar": 'מאמר', 'Introduction': 'הקדמה'}
ONES = ['', 'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט']
TENS = ['', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ']
HUNDREDS = ['', 'ק', 'ר', 'ש', 'ת', 'תק', 'תר', 'תש', 'תת', 'תתק']


def gematria(n):
    if n <= 0 or n >= 1000:
        return str(n)
    s = HUNDREDS[n // 100] + TENS[(n % 100) // 10] + ONES[n % 10]
    return s.replace('יה', 'טו').replace('יו', 'טז')


def daf(text):
    """'דף ב.' -> 'דף ב עמוד א', 'דף ב:' -> 'דף ב עמוד ב'"""
    t = text.strip()
    if t.startswith('דף ') and t[-1:] in ('.', ':'):
        return t[:-1].strip() + (' עמוד א' if t[-1] == '.' else ' עמוד ב')
    return t


def sef_heading(line):
    """'Daf 2a' -> 'דף ב עמוד א', 'Chapter 3' -> 'פרק ג'; None = not a section line"""
    m = RE_SEF_SECTION.match(line.strip())
    if not m:
        return None
    word, val = m.group(1), m.group(2)
    heb = HEB_SECTION.get(word, word)
    if word == 'Daf':
        dm = re.match(r'^(\d+)([ab])$', val)
        if dm:
            return 'דף %s עמוד %s' % (gematria(int(dm.group(1))), 'א' if dm.group(2) == 'a' else 'ב')
    if val.isdigit():
        return heb + ' ' + gematria(int(val))
    return heb + ' ' + val


books, book_ix = [], {}            # (name, collection, library)
heads, head_ix = [''], {'': 0}
labels, label_ix = [''], {'': 0}


def intern(table, ix, s):
    i = ix.get(s)
    if i is None:
        i = ix[s] = len(table)
        table.append(s)
    return i


rec = {}   # word -> [total, {collection: [count, [places]]}, [[book, count, place]], number of other books]


def scan_text(text, book, kind, collection, fmt):
    """count the words of one file; returns {word: [count, [places]]} (places: the first 2)"""
    found = {}
    path_heads = {}          # level -> text (below h1)
    head_id = 0
    line_no = 0
    lines = text.split('\n')
    if fmt == 'sef':
        lines = lines[9:] if len(lines) > 9 else []    # the header of merged.txt (titles, versions, sources)
    for raw in lines:
        if fmt == 'otz':
            m = RE_H.match(raw)
            if m:
                level = int(m.group(1))
                if level >= 2:
                    t = RE_TAGS.sub('', m.group(2)).strip()
                    t = daf(t)[:60]
                    path_heads = {k: v for k, v in path_heads.items() if k < level}
                    path_heads[level] = t
                    head_id = intern(heads, head_ix, ', '.join(path_heads[k] for k in sorted(path_heads)))
                    line_no = 0
                continue
        else:
            h = sef_heading(raw)
            if h is not None:
                head_id = intern(heads, head_ix, h)
                line_no = 0
                continue
        s = tokens.clean_line(raw)
        if not s or not s.strip():
            continue
        line_no += 1
        toks = [tokens.norm_token(t) for t in tokens.RE_TOKEN.findall(s)]
        if not toks:
            continue
        label = None
        for pos, tok in enumerate(toks, 1):
            if tok not in targets:
                continue
            if label is None:
                lab = str(line_no)
                if kind == 'core' and collection in ('תנ"ך', 'משנה'):
                    mm = RE_MARK.match(RE_TAGS.sub('', raw))
                    if mm:
                        lab = mm.group(1).replace('׳', '').replace('״', '').replace('"', '').replace("'", '')
                label = intern(labels, label_ix, lab)
            f = found.get(tok)
            if f is None:
                f = found[tok] = [0, []]
            f[0] += 1
            if len(f[1]) < 2:
                f[1].append((book, head_id, label, pos))
    return found


def keep_top(lst, ent, n=3):
    """lst keeps the n entries [book, count, place] with the biggest count; of equal counts, the one that came first stays"""
    for e in lst:
        if e[0] == ent[0]:      # another file of the same book (a book in several files): one entry, the counts added up
            e[1] += ent[1]
            return
    if len(lst) < n:
        lst.append(ent)
        return
    j = 0
    for k in range(1, len(lst)):
        if lst[k][1] <= lst[j][1]:    # the last of the smallest: what came first is kept
            j = k
    if ent[1] > lst[j][1]:
        del lst[j]          # the list stays in the order things came, so "the last of the smallest" is the latest of them
        lst.append(ent)


def merge_file(found, book, kind, collection):
    for w, (cnt, places) in found.items():
        r = rec.get(w)
        if r is None:
            r = rec[w] = [0, {}, [], 0, -1]   # total, core, other books (top 3), number of other books, the last book counted
        r[0] += cnt
        if kind == 'core':
            # each copy of the four books apart (the Otzaria copy and the Sefaria copy): a word the Bavli has once is not shown
            # as "twice", and its two places are not one place in two copies (merge_sources.py shows one copy)
            key = collection + '|' + books[book][2]
            c = r[1].get(key)
            if c is None:
                c = r[1][key] = [0, []]
            c[0] += cnt
            for p in places:
                if len(c[1]) < 2:
                    c[1].append(p)
        else:
            if r[4] != book:   # the files of one book come one after another: the book is counted once
                r[3] += 1
                r[4] = book
            keep_top(r[2], [book, cnt, places[0]])


def blob_reader(repo):
    env = dict(os.environ, GIT_NO_LAZY_FETCH='1')
    p = subprocess.Popen(['git', '-C', repo, 'cat-file', '--batch'], stdin=subprocess.PIPE, stdout=subprocess.PIPE, env=env)
    return p


def tree(repo):
    """path -> blob id of every file of HEAD"""
    out = subprocess.run(['git', '-C', repo, 'ls-tree', '-r', '-z', 'HEAD'], stdout=subprocess.PIPE, check=True).stdout
    res = {}
    for ent in out.split(b'\0'):
        if not ent:
            continue
        meta, path = ent.split(b'\t', 1)
        res[path.decode('utf-8')] = meta.split()[2].decode()
    return res


items = [l.rstrip('\n').split('\t') for l in open(list_path, encoding='utf-8') if l.strip()]
trees, readers = {}, {}
done = missing = 0
for repo, path, kind, collection, library, fmt in items:
    if repo not in trees:
        trees[repo] = tree(repo)
        readers[repo] = blob_reader(repo)
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
    size = int(header[2])
    data = rd.stdout.read(size)
    rd.stdout.read(1)
    text = data.decode('utf-8', errors='replace')
    name = os.path.splitext(os.path.basename(path))[0]
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
    book = intern(books, book_ix, (name, collection, library))
    merge_file(scan_text(text, book, kind, collection, fmt), book, kind, collection)
    done += 1
    if done % 500 == 0:
        print('files %d (missing %d) words found %d' % (done, missing, len(rec)), flush=True)


def place(p):
    return '%d.%d.%d.%d' % p


with open(out_path, 'w', encoding='utf-8') as f:
    f.write('#BOOKS\t%d\n' % len(books))
    for b in books:
        f.write('%s\t%s\t%s\n' % b)
    f.write('#HEADS\t%d\n' % len(heads))
    for h in heads:
        f.write(h.replace('\t', ' ').replace('\n', ' ') + '\n')
    f.write('#LABELS\t%d\n' % len(labels))
    for l in labels:
        f.write(l + '\n')
    f.write('#WORDS\t%d\n' % len(rec))
    for w in sorted(rec):
        total, core, other, nother = rec[w][:4]
        cs = ';'.join('%s:%d:%s' % (c, v[0], ','.join(place(p) for p in v[1])) for c, v in core.items())
        os_ = ';'.join('%d:%d:%s' % (b, cnt, place(p)) for b, cnt, p in sorted(other, key=lambda e: -e[1]))   # stable: ties keep their order
        f.write('%s\t%d\t%s\t%s\t%d\n' % (w, total, cs, os_, nother))
print('DONE files %d missing %d words %d books %d heads %d' % (done, missing, len(rec), len(books), len(heads)), flush=True)
