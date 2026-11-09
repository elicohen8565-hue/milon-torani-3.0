import subprocess


def hspell_flags(words, chunk=200000):
    """Return list of booleans (True = Hspell recognises the word) aligned with `words`.
    Needs:  apt-get install hspell   (open-source Hebrew spell checker).
    Hspell reads ISO-8859-8 (not UTF-8); pipe mode `hspell -a` prints one record per input line:
    '*' / '+' = recognised, '&' / '#' = not recognised."""
    res = []
    for i in range(0, len(words), chunk):
        part = words[i:i + chunk]
        data = ('\n'.join(part) + '\n').encode('iso-8859-8', errors='replace')
        p = subprocess.run(['hspell', '-a'], input=data, capture_output=True)
        lines = p.stdout.decode('iso-8859-8', errors='replace').split('\n')[1:]
        recs = [l for l in lines if l != '']
        if len(recs) != len(part):
            raise SystemExit(f'alignment problem: {len(recs)} records for {len(part)} words')
        res.extend(r.startswith('*') or r.startswith('+') for r in recs)
    return res
