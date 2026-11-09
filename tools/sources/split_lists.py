"""sizes of the files of the scan lists (local objects only) and a split into chunks of about equal text, in order"""
import os, subprocess, sys
env = dict(os.environ, GIT_NO_LAZY_FETCH='1')
def tree(repo):
    out = subprocess.run(['git', '-C', repo, 'ls-tree', '-r', '-z', 'HEAD'], stdout=subprocess.PIPE, check=True, env=env).stdout
    t = {}
    for ent in out.split(b'\0'):
        if ent:
            meta, path = ent.split(b'\t', 1)
            t[path.decode()] = meta.split()[2].decode()
    return t
for lst, n in (('list_a.tsv', 3), ('list_b.tsv', 3)):
    rows = [l.rstrip('\n').split('\t') for l in open(lst, encoding='utf-8') if l.strip()]
    sizes = []
    for repo in sorted(set(r[0] for r in rows)):
        t = tree(repo)
        shas = [t.get(r[1], '') for r in rows if r[0] == repo]
        out = subprocess.run(['git', '-C', repo, 'cat-file', '--batch-check'], input=('\n'.join(shas) + '\n').encode(),
                             stdout=subprocess.PIPE, env=env).stdout.decode().split('\n')
        sz = {}
        for line in out:
            p = line.split()
            if len(p) == 3 and p[1] == 'blob':
                sz[p[0]] = int(p[2])
        for r in rows:
            if r[0] == repo:
                r.append(sz.get(t.get(r[1], ''), 0))
    total = sum(r[6] for r in rows)
    print(lst, 'files', len(rows), 'MB', total // 1000000)
    chunk, acc, k = [], 0, 1
    for r in rows:
        chunk.append(r); acc += r[6]
        if acc >= total * k / n and k < n:
            open('%s_%d.tsv' % (lst[:-4], k), 'w', encoding='utf-8').write(''.join('\t'.join(map(str, x[:6])) + '\n' for x in chunk))
            print('  chunk', k, len(chunk), 'files'); chunk = []; k += 1
    open('%s_%d.tsv' % (lst[:-4], k), 'w', encoding='utf-8').write(''.join('\t'.join(map(str, x[:6])) + '\n' for x in chunk))
    print('  chunk', k, len(chunk), 'files')
