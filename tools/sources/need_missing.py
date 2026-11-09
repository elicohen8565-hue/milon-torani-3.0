"""which blobs of the scan lists are not in the local clones yet -> need_<n>.txt (one id per line, per repo)"""
import subprocess, sys
lists = sys.argv[1:]
rows = [l.rstrip('\n').split('\t') for f in lists for l in open(f, encoding='utf-8') if l.strip()]
repos = sorted(set(r[0] for r in rows))
for i, repo in enumerate(repos):
    out = subprocess.run(['git', '-C', repo, 'ls-tree', '-r', '-z', 'HEAD'], stdout=subprocess.PIPE, check=True).stdout
    tree = {}
    for ent in out.split(b'\0'):
        if ent:
            meta, path = ent.split(b'\t', 1)
            tree[path.decode()] = meta.split()[2].decode()
    miss = set(l[1:].strip() for l in subprocess.run(['git', '-C', repo, 'rev-list', '--objects', '--missing=print', 'HEAD'],
                                                      stdout=subprocess.PIPE, check=True).stdout.decode().split('\n') if l.startswith('?'))
    need = sorted(set(tree[r[1]] for r in rows if r[0] == repo and r[1] in tree and tree[r[1]] in miss))
    open('need_%d.txt' % i, 'w').write('\n'.join(need) + ('\n' if need else ''))
    print(i, repo, 'files in list', sum(1 for r in rows if r[0] == repo), 'missing blobs', len(need))
