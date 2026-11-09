# Rebuilds the binary/derived inputs of TorahDictionary.exe (run in the build folder before build.sh; scripts/build_all.sh does it all, see BUILD.md)
#   pip install wordfreq pillow
#   basic.txt / more.txt  = data/basic_list.txt / data/extended_only_list.txt  (copy them here first)
#   retired.txt           = data/retired_list.txt (the words earlier versions had and this one dropped; may be empty)
#   hebrew_base.txt       = if the file is already in this folder (copy data/hebrew_base.txt) it is used exactly as it is;
#                           only when it is missing it is generated from wordfreq. Using the saved copy keeps the
#                           "chibur torani" list - and with it the word counts - identical from build to build.
import io, itertools, os, re, struct, subprocess
from PIL import Image, ImageDraw, ImageFont

# 1. base Hebrew words for "chibur torani" (Aramaic connectors are put in front of them inside the program)
if not os.path.exists('hebrew_base.txt'):
    from wordfreq import top_n_list
    heb = re.compile(r'^[א-ת]+$')
    words = [w for w in top_n_list('he', 12000) if heb.match(w) and 3 <= len(w) <= 11 and not w.startswith('ד')][:7000]
    open('hebrew_base.txt', 'w', encoding='utf-8', newline='\n').write('\n'.join(words) + '\n')

# 2. icon: the blue of Word, a rounded square with the letter alef. Every size is drawn 8 times larger and then shrunk (LANCZOS), so the corners
#    and the letter are smooth; each size is made on its own (not shrunk from the biggest), so the small ones stay sharp.
SIZES = [16, 20, 24, 32, 40, 48, 64, 128, 256]
def icon(size):
    big = size * 8
    im = Image.new('RGBA', (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((0, 0, big - 1, big - 1), radius=int(big * 0.22), fill=(0x18, 0x5A, 0xBD, 255))
    f = ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', int(big * 0.78))
    d.text((big / 2, big / 2), 'א', font=f, fill=(0xFF, 0xFF, 0xFF, 255), anchor='mm')
    return im.resize((size, size), Image.LANCZOS)
frames = [icon(n) for n in SIZES]

# The .ico file is written by hand (not by Pillow): Pillow stores EVERY picture as PNG, and the .NET class Icon (and Mono) cannot always load a small
# PNG picture from an icon file. So, as in every professional icon: the 256 picture is PNG, all the smaller ones are plain 32-bit bitmaps with transparency.
def ico_picture(im):
    w, h = im.size
    if w >= 256:
        buf = io.BytesIO()
        im.save(buf, 'PNG')
        return buf.getvalue()
    px = im.tobytes('raw', 'BGRA', 0, -1)             # bottom row first, as in a bitmap
    mask = bytes(((w + 31) // 32 * 4) * h)             # the "AND" mask: all zero (the transparency comes from the alpha channel)
    head = struct.pack('<IiiHHIIiiII', 40, w, h * 2, 1, 32, 0, len(px) + len(mask), 0, 0, 0, 0)
    return head + px + mask

pictures = [ico_picture(f) for f in frames]
ico = struct.pack('<HHH', 0, 1, len(frames))
offset = 6 + 16 * len(frames)
for f, p in zip(frames, pictures):
    ico += struct.pack('<BBBBHHII', f.size[0] % 256, f.size[1] % 256, 0, 0, 1, 32, len(p), offset)   # a size of 256 is stored as 0
    offset += len(p)
with open('app.ico', 'wb') as fh:
    fh.write(ico + b''.join(pictures))

# 3. exact number of words of each edition (shown on the cards without opening the word lists at start-up).
#    Same rules as Core.BuildWords: basic + acronyms of 2-3 letters [+ extended] [+ composition], without duplicates.
#    The program's self-test builds every edition for real and checks these numbers, so a stale file is caught.
CONNECTORS = ['ד', 'וד', 'כד', 'וכד', 'מד', 'ומד']
FIRST = 'אבגדהוזחטיכלמנסעפצקרשת'      # 22 letters, regular form
LAST = 'אבגדהוזחטיךלםןסעףץקרשת'       # 22 letters, final forms where Hebrew writes them

def acronyms(n):
    out = []
    for head in itertools.product(FIRST, repeat=n - 1):
        for last in LAST:
            out.append(''.join(head) + '"' + last)
    return out

def read_words(path):
    with open(path, encoding='utf-8') as f:
        return [l.strip() for l in f if l.strip()]

basic = read_words('basic.txt')
more = read_words('more.txt')
base = read_words('hebrew_base.txt')
acro = acronyms(2) + acronyms(3)
comp = [c + b for b in base for c in CONNECTORS]

def count(*parts):
    seen = set()
    for p in parts:
        seen.update(p)
    return len(seen)

counts = [
    ('basic', count(basic, acro)),
    ('basic_comp', count(basic, acro, comp)),
    ('extended', count(basic, acro, more)),
    ('extended_comp', count(basic, acro, more, comp)),
]
with open('counts.txt', 'w', encoding='utf-8', newline='\n') as f:
    for k, v in counts:
        f.write('%s %d\n' % (k, v))
print('counts:', ', '.join('%s=%d' % kv for kv in counts))

# 4. gzip the word lists that are embedded as resources
if not os.path.exists('retired.txt'):
    open('retired.txt', 'w').close()
subprocess.check_call(['gzip', '-kf', '-9', 'basic.txt', 'more.txt', 'hebrew_base.txt', 'retired.txt'])
