"""Version 3.1.0.1 (the deep check of 9-11-2026, 22:40): forms of the word lists that are typos of the online libraries.

usage: python3 -I drop_typos.py <repo>
Run after build_lists_v2.py. Takes the words below out of <repo>/data/basic_list.txt and <repo>/data/extended_only_list.txt and
adds them to <repo>/data/retired_list.txt (so that an update does not keep them in the person's file as if the person had added
them). Running it again changes nothing.

Why they were in: build_lists_v2.py checks the shape of every NEW form (no letter three times in a row, no regular כ מ נ פ צ at the
end ...), but the words that 3.0 already had were only counted ("at least 3 times as written"), not checked for shape. These forms
are written 3 times or more in the libraries, but they are typos of the digital copies (most of them a yod or a vav written three
times: הייינו is written so 176 times, in 162 books - a copying error that repeats), so Word must go on marking them.

Second part (the second deep check, 10-11-2026): 18 forms that are two or more words run together in the digital copies
(אסורלהתרפאות, עובדכוכבים, הכתובשנאמר ...): each is written so 3-12 times in 2-4 books, so the count let it in.

Third part (the same check): "chibur torani" puts the connectors ד, וד, כד ... before 7,000 common words of today's Hebrew
(data/hebrew_base.txt, made once from a frequency list of the internet). 27 of them have no place in a Torah dictionary: crude
words, names from other religions, and slang calls (וואו, יאללה). They leave that list, and their six forms each go to the retired
list. Words that are also words of the Torah stay (זונה, אונס, מין, מגעיל ...).

Kept on purpose (they look alike, but are right): the divine-name sign ייי with its prefixes (לייי, בייי ...); a prefix before a word
or an abbreviation that starts with a doubled letter (ובבבא, דבבבלי, כככבים, כככר, מממרים, דמממון, בבב"ק, מממ"ה ...); the Aramaic
forms of the Targum (תיצבייין, תהוייין ...); and every form of the basic list that the four books themselves have (התתתונים in the
Yerushalmi, מממרו in the Bavli ...). Not decided here: the forms of two letters without a quote mark (numbers such as תב, יט) that
3.0 had - they serve references (סימן תב).
"""
import os
import sys

TYPOS = {
    # a letter three times in a row inside a common word (the right word in brackets)
    'בסוגייין': 'בסוגיין', 'בריייתא': 'ברייתא', 'דבריייתא': 'דברייתא', 'הבריייתא': 'הברייתא', 'דאוריייתא': 'דאורייתא',
    'מדאוריייתא': 'מדאורייתא', 'דמייין': 'דמיין', 'דמייירי': 'דמיירי', 'דקייי"ל': 'דקיי"ל', 'דשיייך': 'דשייך', 'שיייך': 'שייך',
    'הייינו': 'היינו', 'והייינו': 'והיינו', 'ודיייק': 'ודייק', "ועייי'": "ועיי'", 'ועייין': 'ועיין', 'עייי"ש': 'עיי"ש',
    'עדייין': 'עדיין', 'חיייב': 'חייב', 'חיייבים': 'חייבים', 'יתחיייב': 'יתחייב', 'לחיייבו': 'לחייבו', 'מחיייבין': 'מחייבין',
    'לקייים': 'לקיים', 'מתקייים': 'מתקיים', 'קייימא': 'קיימא', 'קייימי': 'קיימי', 'לשנוייי': 'לשנויי', 'מתיייחס': 'מתייחס',
    'נתייישב': 'נתיישב', 'ענייים': 'עניים', 'מכווונים': 'מכוונים', 'מתפללל': 'מתפלל', 'דהההלכה': 'דההלכה',
    # a prefix written twice
    'בבבוקר': 'בבוקר', 'לללמוד': 'ללמוד', 'מממשנתנו': 'ממשנתנו', 'ומממילא': 'וממילא', 'שששחטו': 'ששחטו', 'הבבבלי': 'הבבלי',
    # not words
    'חחח': '', 'ללל': '',
    # ב read as כ at the end (a regular כ never ends a word): הכתוב
    'כתוכ': 'כתוב', 'הכתוכ': 'הכתוב', 'שכתוכ': 'שכתוב',
}
# words run together (the words they are, for the reader)
TYPOS.update({
    'אסורלהתרפאות': 'אסור להתרפאות', 'בברייתחבסמוך': 'בברייתא בסמוך', 'דאפילובמתכון': 'דאפילו במתכוון', 'הראיהמכתובות': 'הראיה מכתובות',
    'אמאיאמרינן': 'אמאי אמרינן', 'דתפלתציבור': 'דתפלת ציבור', 'הכתובשנאמר': 'הכתוב שנאמר', 'ישראלינתנו': 'ישראל ינתנו',
    'מידידאכילה': 'מידי דאכילה', 'נחלאארנונא': 'נחלא ארנונא', 'עובדכוכבים': 'עובד כוכבים', 'רבינודלשכת': 'רבינו דלשכת',
    'שהיינומילין': 'שהיינו מילין', 'שלישיויאמר': 'שלישי ויאמר', 'יהודהאומר': 'יהודה אומר', 'מאיסוראדאם': 'מאיסורא דאם',
    'עשרהובתינוקת': 'עשרה ובתינוקת', 'שמתחילהקדירה': 'שמתחילה קדירה',
})

# common words of today's Hebrew that "chibur torani" must not build on
BASE_OUT = ('סקס פורנו חרא החרא הומו בתחת התחת מינית מטומטם אידיוט אקס '
            'ישו מוחמד הכנסייה המולד נוצרים הנוצרים הנוצרית מוסלמים המוסלמים היטלר הנאצים '
            'וואו אוקיי יאללה וואלה אהה').split()
CONNECTORS = ['ד', 'וד', 'כד', 'וכד', 'מד', 'ומד']   # as Core.Connectors


def read(p):
    with open(p, encoding='utf-8') as f:
        return [l.strip() for l in f if l.strip()]


def write(p, words):
    with open(p, 'w', encoding='utf-8', newline='\n') as f:
        f.write(''.join(w + '\n' for w in words))


repo = sys.argv[1]
data = os.path.join(repo, 'data')
basic, more = read(os.path.join(data, 'basic_list.txt')), read(os.path.join(data, 'extended_only_list.txt'))
retired = read(os.path.join(data, 'retired_list.txt'))
base = read(os.path.join(data, 'hebrew_base.txt'))
out_b = [w for w in basic if w in TYPOS]
out_m = [w for w in more if w in TYPOS]
missing = sorted(set(TYPOS) - set(out_b) - set(out_m) - set(retired))
assert not missing, 'not in the lists and not retired either: %s' % ' '.join(missing)
write(os.path.join(data, 'basic_list.txt'), [w for w in basic if w not in TYPOS])
write(os.path.join(data, 'extended_only_list.txt'), [w for w in more if w not in TYPOS])
out_base = [w for w in base if w in BASE_OUT]
left = [w for w in base if w not in BASE_OUT]
# the forms that the program no longer makes (a form that the word lists have, or that another base word still gives, stays)
still = set(basic) | set(more) | {c + w for w in left for c in CONNECTORS}
forms = [c + w for w in BASE_OUT for c in CONNECTORS if c + w not in still]
write(os.path.join(data, 'hebrew_base.txt'), left)
new_retired = sorted(set(retired) | set(TYPOS) | set(forms))
write(os.path.join(data, 'retired_list.txt'), new_retired)
print('typos taken out: basic %d, extended %d; common words: %d (%d forms); retired list: %d words'
      % (len(out_b), len(out_m), len(out_base), len(forms), len(new_retired)))
