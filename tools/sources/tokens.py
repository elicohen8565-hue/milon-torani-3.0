"""Word extraction used for every text of the Otzaria library (same rules as extract.py of the first scan).
Re-typed on 2-11-2026 (evening) from pipeline/tokens.py of the Project: every invisible or look-alike character is written as a \\u escape."""
import html
import re
import unicodedata

RE_HEADING = re.compile(r'\s*<h[1-6][ >]')
RE_FOOTNOTE = re.compile(r'<i class="footnote">.*?</i>', re.S)
RE_SUP = re.compile(r'<sup[^>]*>.*?</sup>', re.S)
RE_BR = re.compile(r'<br\s*/?>', re.I)
RE_TAG = re.compile(r'<[^>]*>')
RE_CURLY = re.compile(r'\{[^{}]*\}')
RE_LEAD = re.compile(r'^\s*\(\s*[א-ת"\'׳״]{1,4}\s*\)\s*')
# Hebrew points and cantillation (U+0591-05C7), grapheme joiner, dots, zero-width and bidi controls, varika
RE_MARKS = re.compile('[֑-ׇ͏̇̈​-‏‪-‮⁠ﬞ]')
SPECIAL = {0x05F0: 'וו', 0x05F1: 'וי', 0x05F2: 'יי', 0x05BE: ' ', 0x05C0: ' ', 0x05C3: ' ', 0x05C6: ' '}
RE_TOKEN = re.compile(r"[א-ת]+(?:[\"'׳״][א-ת]+)*[׳']?")


def norm_token(tok):
    tok = tok.replace('׳', "'").replace('״', '"')
    trail = ''
    if tok.endswith("'"):
        trail, tok = "'", tok[:-1]
    segs = re.split(r"[\"']", tok)
    marks = re.findall(r"[\"']", tok)
    res = segs[0]
    for i, m in enumerate(marks):
        nxt = segs[i + 1]
        if m == '"' and i == len(marks) - 1 and len(nxt) == 1:
            res += '"' + nxt
        else:
            res += nxt
    return res + trail


def clean_line(line):
    if RE_HEADING.match(line):
        return None
    s = RE_FOOTNOTE.sub(' ', line)
    s = RE_SUP.sub(' ', s)
    s = RE_BR.sub(' ', s)
    s = RE_TAG.sub('', s)
    s = html.unescape(s)
    s = RE_CURLY.sub(' ', s)
    s = RE_LEAD.sub('', s)
    s = unicodedata.normalize('NFKD', s)
    s = s.translate(SPECIAL)
    s = RE_MARKS.sub('', s)
    return s


def file_tokens(path):
    """yield every normalised token of one text file (main text only)"""
    with open(path, encoding='utf-8', errors='replace') as f:
        for line in f:
            s = clean_line(line.rstrip('\n'))
            if not s:
                continue
            for tok in RE_TOKEN.findall(s):
                yield norm_token(tok)
