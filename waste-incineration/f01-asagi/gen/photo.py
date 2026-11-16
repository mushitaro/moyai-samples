# -*- coding: utf-8 -*-
"""手書き帳票のスキャン（写真）を作る。帳票の PDF の頁を画像にし、傾き・影・ぼけ・明るさのむら・紙の折れ・手書きの丸やチェックを足す。"""
import os, glob, subprocess, io, math, random, re, shutil
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageEnhance, ImageOps
from xlutil import TMP

FONT = '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'
LO = ['/usr/bin/soffice', '-env:UserInstallation=file:///tmp/lo_F01', '--headless', '--convert-to', 'pdf', '--outdir']

def to_pdf(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(paths), 40):
        subprocess.run(LO + [outdir] + paths[i:i + 40], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=1800)

def page_words(pdf, page):
    """pdftotext -bbox でその頁の単語と位置（pt）"""
    out = subprocess.run(['pdftotext', '-bbox', '-f', str(page), '-l', str(page), pdf, '-'], capture_output=True, text=True).stdout
    m = re.search(r'<page width="([\d.]+)" height="([\d.]+)"', out)
    W, H = (float(m.group(1)), float(m.group(2))) if m else (842, 595)
    words = [(float(a), float(b), float(c), float(d), w) for a, b, c, d, w in
             re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', out)]
    return W, H, words

def raster(pdf, page, width):
    tmp = os.path.join(TMP, '_rast')
    os.makedirs(tmp, exist_ok=True)
    for f in glob.glob(tmp + '/*'):
        os.remove(f)
    subprocess.run(['pdftoppm', '-png', '-scale-to-x', str(width), '-scale-to-y', '-1', '-f', str(page), '-l', str(page), pdf, tmp + '/p'], check=True)
    return Image.open(sorted(glob.glob(tmp + '/*.png'))[0]).convert('RGB')

def _pen(draw, pts, color, w):
    for a, b in zip(pts, pts[1:]):
        draw.line([a, b], fill=color, width=w)

def handwrite(img, rnd, targets, scale, person):
    """手書きの丸・チェック・書き込み"""
    d = ImageDraw.Draw(img)
    ink = rnd.choice([(25, 35, 120), (30, 30, 40), (160, 30, 30)])
    for (x0, y0, x1, y1) in targets[:rnd.randint(1, 3)]:
        cx, cy = (x0 + x1) / 2 * scale, (y0 + y1) / 2 * scale
        rx, ry = max(18, (x1 - x0) * scale / 2 + 8), max(12, (y1 - y0) * scale / 2 + 6)
        pts = [(cx + rx * math.cos(t) * (1 + 0.06 * math.sin(3 * t)), cy + ry * math.sin(t)) for t in [i * 0.2 for i in range(34)]]
        _pen(d, pts, ink, 3)
    for _ in range(rnd.randint(1, 4)):   # チェック
        x, y = rnd.randint(60, img.width - 80), rnd.randint(140, img.height - 80)
        _pen(d, [(x, y), (x + 9, y + 12), (x + 28, y - 16)], ink, 3)
    f = ImageFont.truetype(FONT, int(24 * rnd.uniform(0.9, 1.2)))
    note = rnd.choice(['済', '確認 ' + person, '了', '要確認', person + ' 確認済', 'OK'])
    tx = Image.new('RGBA', (220, 60), (0, 0, 0, 0))
    ImageDraw.Draw(tx).text((5, 5), note, font=f, fill=ink + (230,))
    tx = tx.rotate(rnd.uniform(-12, 12), expand=1, resample=Image.BICUBIC)
    img.paste(tx, (rnd.randint(img.width - 330, img.width - 240), rnd.randint(30, 110)), tx)
    return img

def stain(img, box, scale, rnd):
    """コーヒーのしみ・強いぼけでその欄を読めなくする"""
    x0, y0, x1, y1 = [int(v * scale) for v in box]
    pad = 10
    reg = img.crop((x0 - pad, y0 - pad, x1 + pad, y1 + pad)).filter(ImageFilter.GaussianBlur(7))
    ov = Image.new('RGBA', reg.size, (120, 80, 40, 120))
    reg = Image.alpha_composite(reg.convert('RGBA'), ov).convert('RGB')
    img.paste(reg, (x0 - pad, y0 - pad))
    return img

def augment(img, rnd):
    w, h = img.size
    # 紙の色・明るさのむら
    grad = Image.new('L', (w, h))
    gd = ImageDraw.Draw(grad)
    cx, cy = rnd.uniform(0.2, 0.8) * w, rnd.uniform(0.2, 0.8) * h
    for i in range(0, max(w, h), 40):
        v = int(255 - 55 * i / max(w, h))
        gd.ellipse([cx - i, cy - i, cx + i, cy + i], outline=v, width=42)
    grad = grad.filter(ImageFilter.GaussianBlur(60))
    paper = Image.new('RGB', (w, h), (rnd.randint(236, 248), rnd.randint(232, 244), rnd.randint(214, 230)))
    img = Image.composite(img, paper, ImageOps.invert(img.convert('L')).point(lambda p: 255 if p > 40 else p * 6))
    img = Image.composite(img, Image.new('RGB', (w, h), (150, 145, 135)), grad.point(lambda p: max(170, p)))
    # 紙の折れ
    d = ImageDraw.Draw(img)
    if rnd.random() < 0.8:
        y = int(h * rnd.uniform(0.35, 0.65))
        d.line([(0, y), (w, y + rnd.randint(-15, 15))], fill=(175, 170, 160), width=2)
        d.line([(0, y + 3), (w, y + 3 + rnd.randint(-15, 15))], fill=(250, 248, 242), width=2)
    # ぼけ
    img = img.filter(ImageFilter.GaussianBlur(rnd.uniform(0.5, 1.1)))
    # 傾き（±3〜8度）と影（台の色）
    ang = rnd.choice([-1, 1]) * rnd.uniform(3, 8)
    bg = (rnd.randint(55, 90),) * 3
    img = img.rotate(ang, expand=1, fillcolor=bg, resample=Image.BICUBIC)
    sh = Image.new('L', img.size, 0)
    ImageDraw.Draw(sh).rectangle([0, 0, img.width, img.height], fill=0)
    edge = Image.linear_gradient('L').resize((img.width, img.height))
    img = Image.composite(img, ImageEnhance.Brightness(img).enhance(0.78), edge.point(lambda p: 255 if p > 70 else 140 + p))
    img = ImageEnhance.Contrast(img).enhance(rnd.uniform(0.9, 1.05))
    return img

def save_jpeg(img, path, comment, target=100_000):
    long = max(img.size)
    tgt = random.Random(path).randint(1568, 2000)
    if long != tgt:
        k = tgt / long
        img = img.resize((int(img.width * k), int(img.height * k)), Image.LANCZOS)
    lo, hi, best = 20, 85, None
    for _ in range(7):
        q = (lo + hi) // 2
        b = io.BytesIO()
        img.save(b, 'JPEG', quality=q, comment=comment.encode('utf-8'), optimize=True)
        n = b.tell()
        best = b
        if n > target * 1.1:
            hi = q - 1
        elif n < target * 0.9:
            lo = q + 1
        else:
            break
    open(path, 'wb').write(best.getvalue())
    return img.size, os.path.getsize(path)
