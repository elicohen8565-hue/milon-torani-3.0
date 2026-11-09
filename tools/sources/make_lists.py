"""The words to look for and the files to scan, for scan_sources.py (9-11-2026).
usage: python3 -I make_lists.py <milon repo> <out folder> [--targets-only]
  --targets-only: only targets.txt (also the --words list of merge_sources.py: every word of the biggest edition)
writes targets.txt (every word of the biggest edition: the lists, the generated abbreviations and the chibur forms),
list_a.tsv (Otzaria core four books, Sefaria core four books, the rest of the Otzaria main library, the rest of Sefaria) and
list_b.tsv (the other sources of the maintained Otzaria library, as scanned in stage 5)."""
import itertools
import os
import sys

repo, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)


def words(p):
    return [l.strip() for l in open(p, encoding='utf-8') if l.strip()]


basic = words(os.path.join(repo, 'data/basic_list.txt'))
more = words(os.path.join(repo, 'data/extended_only_list.txt'))
base = words(os.path.join(repo, 'data/hebrew_base.txt'))
FIRST = 'אבגדהוזחטיכלמנסעפצקרשת'
LAST = 'אבגדהוזחטיךלםןסעףץקרשת'
acro = set(''.join(h) + '"' + l for n in (2, 3) for h in itertools.product(FIRST, repeat=n - 1) for l in LAST)
comp = set(c + b for b in base for c in ['ד', 'וד', 'כד', 'וכד', 'מד', 'ומד'])
targets = set(basic) | set(more) | acro | comp
with open(os.path.join(out, 'targets.txt'), 'w', encoding='utf-8') as f:
    f.write('\n'.join(sorted(targets)) + '\n')
print('targets', len(targets))
if '--targets-only' in sys.argv[3:]:
    sys.exit(0)

OLD = '/home/claude/work/repo'
SEF = '/home/claude/work/ext/TorahBibleCodes_Sefaria-Export'
NEW = '/home/claude/work/ext2/otz2'
SKIP = 'לא רלוונטי'
CORE_OTZ = {'תנך': 'תנ"ך', 'משנה': 'משנה', 'תלמוד בבלי': 'תלמוד בבלי', 'תלמוד ירושלמי': 'תלמוד ירושלמי'}


def otz_core(p):
    parts = p.split('/')
    if len(parts) < 4 or parts[0] != 'אוצריא' or not p.endswith('.txt'):
        return None
    if parts[1] == 'תנך' and parts[2] in ('תורה', 'נביאים', 'כתובים'):
        return 'תנ"ך'
    if parts[1] in ('משנה', 'תלמוד בבלי', 'תלמוד ירושלמי') and parts[2].startswith('סדר '):
        return CORE_OTZ[parts[1]]
    return None


def sef_core(p):
    parts = p.split('/')
    if not p.endswith('/Hebrew/merged.txt') or len(parts) < 5:
        return None
    if parts[1] == 'Tanakh' and parts[2] in ('Torah', 'Prophets', 'Writings'):
        return 'תנ"ך'
    if parts[1] == 'Mishnah' and parts[2].startswith('Seder '):
        return 'משנה'
    if parts[1] == 'Talmud' and parts[2] in ('Bavli', 'Yerushalmi') and parts[3].startswith('Seder '):
        return 'תלמוד בבלי' if parts[2] == 'Bavli' else 'תלמוד ירושלמי'
    return None


old_paths = words('/home/claude/work/all_paths.txt')
sef_paths = words('/home/claude/work/ext/TorahBibleCodes_Sefaria-Export_paths.txt')
rows_a = []
rows_a += [(OLD, p, 'core', otz_core(p), 'אוצריא', 'otz') for p in old_paths if otz_core(p)]
rows_a += [(SEF, p, 'core', sef_core(p), 'ספריא', 'sef') for p in sef_paths if sef_core(p)]
rows_a += [(OLD, p, 'book', '', 'אוצריא', 'otz') for p in old_paths
           if p.startswith('אוצריא/') and p.endswith('.txt') and not otz_core(p) and SKIP not in p]
rows_a += [(SEF, p, 'book', '', 'ספריא', 'sef') for p in sef_paths
           if p.startswith('txt/') and p.endswith('/Hebrew/merged.txt') and not sef_core(p)]

SOURCES = [
    ('דיקטה', ['DictaToOtzaria/ערוך/ספרים', 'extraBooks/דיקטה', 'extraBooks/דיקטה ערוך']),
    ('אורייתא', ['OraytaToOtzaria/ספרים']),
    ('ספרים נוספים', ['MoreBooks/ספרים/אוצריא']),
    ('בית אהרן וישראל', ['BeitAharonVeYisraelToOtzaria/ספרים']),
    ('הספרייה הלאומית', ['National-LibraryToOtzaria/ספרים/אוצריא', 'extraBooks/National-LibraryToOtzaria']),
    ('תורת אמת', ['ToratEmetToOtzaria/ספרים/אוצריא']),
    ('KSK', ['KSK/ספרים/אוצריא']),
    ('ובלכתך בדרך', ['OnYourWayToOtzaria/ספרים/אוצריא']),
    ('פנינים', ['pninimToOtzaria']), ('ויקיטקסט', ['wikisourceToOtzaria', 'extraBooks/ויקיטקסט']),
    ('ויקי-ספרים', ['wikiJewishBooksToOtzaria']), ('ים החכמה', ['yam-HaHachmaToOtzaria']), ('תשמ"א', ['tashmaToOtzaria']),
    ('פרויקט בן-יהודה', ['Ben-YehudaToOtzaria/ספרים/אוצריא']),
    ('אוצריא - ספרים נוספים', ['extraBooks/לא מסודר לפי מקור', 'extraBooks/ספרים פרטיים ועוד', 'extraBooks/האוצר השיתופי']),
]
new_paths = words('/tmp/claude-0/-home-claude/e84253b5-2401-5767-b94f-79aceb04db21/scratchpad/otz2_all.txt')
rows_b = []
for lib, folders in SOURCES:
    for p in new_paths:
        if p.endswith('.txt') and SKIP not in p and any(p.startswith(fo + '/') for fo in folders) and '/.' not in p:
            rows_b.append((NEW, p, 'book', '', lib, 'otz'))
for name, rows in (('list_a.tsv', rows_a), ('list_b.tsv', rows_b)):
    with open(os.path.join(out, name), 'w', encoding='utf-8') as f:
        for r in rows:
            f.write('\t'.join(r) + '\n')
    print(name, len(rows), 'core', sum(1 for r in rows if r[2] == 'core'))
