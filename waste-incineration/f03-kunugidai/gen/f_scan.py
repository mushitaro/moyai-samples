# -*- coding: utf-8 -*-
"""市直営補修班 d1 の手書き日報（H13〜H28）のスキャン PDF。年度 1 冊。
様式（B5 縦の印刷物）に手書きで記入した頁を画像にし、スキャンの傾き（±3〜8 度）・影・ぼけ・明るさのむら・紙の折れ・
手書きの丸・チェック・書き込みを加える。1 頁 1,600×2,000px、JPEG（100KB 前後）。
H29 に d2 が綴りをスキャンしたとき、直営の補修作業があった日と、ほかに年 10 日ほどを残した（全部ではない）。
手書きの字形を持つフォントが無いので、IPAゴシックを 1 字ずつ揺らして書く。"""
import io, random
import datetime as dt
from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageChops
from pypdf import PdfWriter, PdfReader
from world import SEED
from rec import Rec
from f_daily import direct_text

FONT = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
PW, PH = 1350, 1906
CW, CH = 1600, 2000
WD = "月火水木金土日"
_FC = {}


def _f(sz):
    if sz not in _FC:
        _FC[sz] = ImageFont.truetype(FONT, sz)
    return _FC[sz]


def _hand(draw, xy, text, rng, size=36, ink=(25, 35, 85)):
    x, y = xy
    for ch in text:
        s = size + rng.randint(-4, 2)
        draw.text((x + rng.randint(-1, 1), y + rng.randint(-4, 4)), ch, font=_f(s), fill=ink)
        x += (s if ord(ch) > 0x2000 else s * 0.56) + rng.randint(-2, 1)
    return x


def _form(dr):
    k = (0, 0, 0)
    dr.text((70, 50), "【架空】くぬぎ台市 環境センター", font=_f(26), fill=k)
    dr.text((420, 100), "補 修 作 業 日 報", font=_f(52), fill=k)
    dr.text((700, 205), "平成　　年　　月　　日（　）", font=_f(30), fill=k)
    dr.text((800, 255), "天候", font=_f(28), fill=k)
    dr.rectangle((1040, 40, 1280, 180), outline=k, width=3)
    dr.line((1160, 40, 1160, 180), fill=k, width=2)
    dr.line((1040, 78, 1280, 78), fill=k, width=2)
    dr.text((1072, 46), "係長", font=_f(24), fill=k)
    dr.text((1192, 46), "班長", font=_f(24), fill=k)
    top = 320
    cols = [70, 390, 990, 1130, 1280]
    dr.rectangle((70, top, 1280, top + 64), outline=k, width=4)
    for i, t in enumerate(["機器・場所", "作　業　内　容", "人　員", "時　間"]):
        dr.text((cols[i] + 24, top + 16), t, font=_f(28), fill=k)
    rows, rh = 14, 78
    for i in range(rows + 1):
        dr.line((70, top + 64 + i * rh, 1280, top + 64 + i * rh), fill=k, width=2)
    for c in cols:
        dr.line((c, top, c, top + 64 + rows * rh), fill=k, width=3 if c in (70, 1280) else 2)
    y2 = top + 64 + rows * rh + 26
    dr.rectangle((70, y2, 1280, y2 + 160), outline=k, width=3)
    dr.text((90, y2 + 12), "使用材料", font=_f(26), fill=k)
    dr.rectangle((70, y2 + 186, 1280, y2 + 400), outline=k, width=3)
    dr.text((90, y2 + 198), "所見・連絡", font=_f(26), fill=k)
    dr.text((70, PH - 70), "様式 環セ-補-1（H15.4 改）" if True else "", font=_f(20), fill=k)
    return top + 64, rh, cols, y2


def _scan(paper, rng):
    """紙の画像をスキャナの台に置いた画像にする"""
    ang = rng.choice([-1, 1]) * rng.uniform(3, 8)
    # 明るさのむら（紙の上）
    grad = Image.new("L", (PW, PH), 0)
    gd = ImageDraw.Draw(grad)
    cx, cy = rng.randint(0, PW), rng.randint(0, PH)
    for r in range(0, 1800, 60):
        v = min(60, r // 30)
        gd.ellipse((cx - r, cy - r, cx + r, cy + r), outline=v, width=60)
    paper = ImageChops.subtract(paper, Image.merge("RGB", (grad, grad, grad)))
    # 紙の折れ（横に 1 本）
    d = ImageDraw.Draw(paper)
    fy = rng.randint(PH // 3, PH * 2 // 3)
    for k in range(6):
        g = 205 + k * 7
        d.line((0, fy + k, PW, fy + k + rng.randint(-6, 6)), fill=(g, g, g), width=1)
    bed = Image.new("RGB", (CW, CH), (60, 60, 62))
    rot = paper.rotate(ang, resample=Image.BICUBIC, expand=True, fillcolor=(60, 60, 62))
    ox = (CW - rot.width) // 2 + rng.randint(-20, 20)
    oy = (CH - rot.height) // 2 + rng.randint(-20, 20)
    # 影
    sh = Image.new("L", rot.size, 0)
    sh.paste(255, (0, 0, rot.width, rot.height), rot.convert("L").point(lambda p: 255 if p > 70 else 0))
    sh = sh.filter(ImageFilter.GaussianBlur(14))
    bed.paste((25, 25, 25), (ox + 14, oy + 18), sh)
    mask = rot.convert("L").point(lambda p: 255 if p > 66 else 0)
    bed.paste(rot, (ox, oy), mask)
    bed = bed.filter(ImageFilter.GaussianBlur(rng.uniform(0.7, 1.3)))
    px = bed.load()
    for _ in range(5000):
        x, y = rng.randrange(CW), rng.randrange(CH)
        g = rng.randint(120, 200)
        px[x, y] = (g, g, g)
    return bed.convert("L")


def make_year(fy, model, path, fname, parts_by_wid, tmpdir):
    rng = random.Random("%s-scan-%d" % (SEED, fy))
    rec = Rec(fname)
    rows = [r for r in model.direct if (r["date"].year if r["date"].month >= 4 else r["date"].year - 1) == fy]
    byd = {}
    for r in rows:
        byd.setdefault(r["date"], []).append(r)
    work_days = sorted(d for d, lst in byd.items() if any(x["wid"] or x["phase"] == "fixed" for x in lst))
    others = sorted(d for d in byd if d not in work_days)
    keep = sorted(set(work_days) | set(rng.sample(others, min(10, len(others)))))
    if not keep:
        return 0
    imgs = []
    for pi, d in enumerate(keep):
        prng = random.Random("%s-p-%s" % (SEED, d))
        im = Image.new("RGB", (PW, PH), (252, 251, 246))
        dr = ImageDraw.Draw(im)
        y0, rh, cols, y2 = _form(dr)
        wx = getattr(model, "h25_fixed", {}).get(d) or prng.choice("晴晴曇雨")
        _hand(dr, (782, 196), str(fy - 1988 if d.month >= 4 else d.year - 1988), prng, 38)
        _hand(dr, (880, 196), str(d.month), prng, 38)
        _hand(dr, (968, 196), str(d.day), prng, 38)
        _hand(dr, (1078, 198), WD[d.weekday()], prng, 34)
        _hand(dr, (880, 248), wx, prng, 36)
        cx, cy = 1220, 130
        dr.ellipse((cx - 40, cy - 40, cx + 40, cy + 40), outline=(200, 40, 40), width=4)
        dr.text((cx - 22, cy - 18), "d1", font=_f(34), fill=(200, 40, 40))
        mats = []
        lst = sorted(byd[d], key=lambda x: (x["wid"] is None,))[:13]
        for i, it in enumerate(lst):
            yy = y0 + i * rh + 18
            e = model.tag[it["tag"]]
            eq = it.get("eqname") or ("%s %s" % ("%d号" % e.unit if e.unit else "", e.name.replace(" ", ""))).strip()
            txt = direct_text(it, prng)
            men = it.get("men") or ("d1 外%d名" % (it["n"] - 1) if it["n"] > 1 else "1名")
            if it.get("tm"):
                tm = it["tm"]
            else:
                st = prng.choice([8.5, 9.0, 10.0, 13.0, 13.5])
                en = st + it["h"]
                tm = "%d:%02d〜%d:%02d" % (int(st), int(st % 1 * 60), int(en), int(round(en % 1 * 60)) % 60)
            _hand(dr, (82, yy), eq[:11], prng, 32)
            unread = prng.random() < 0.06
            if len(txt) > 18:
                _hand(dr, (402, yy - 10), txt[:18], prng, 30)
                _hand(dr, (402, yy + 26), txt[18:36], prng, 30)
            else:
                _hand(dr, (402, yy), txt, prng, 32)
            _hand(dr, (1000, yy), men, prng, 24)
            _hand(dr, (1138, yy - 8), tm.split("〜")[0] + "〜", prng, 22)
            _hand(dr, (1168, yy + 24), tm.split("〜")[1], prng, 22)
            if unread:
                # インクのにじみ（コーヒーの輪じみ）で読めない
                bx, by = 420 + prng.randint(0, 200), yy - 14
                for k in range(8):
                    g = 120 + k * 12
                    dr.ellipse((bx - 90 + k * 6, by - 30 + k * 3, bx + 260 - k * 6, by + 60 - k * 3), fill=(g, g - 10, g - 30))
            if prng.random() < 0.25:
                dr.text((1290 - 30, yy - 4), "✓", font=_f(34), fill=(30, 40, 160))
            p = parts_by_wid.get(it["wid"]) if it["wid"] and it["phase"] in ("first", "fixed") else None
            if p:
                mats.append("控 %s" % p)
            f = {"page": pi + 1, "日付": d.isoformat(), "天候": wx, "機器・場所": eq[:11], "作業内容": "読めない" if unread else txt[:36],
                 "人員": men, "時間": tm}
            rec.addp("item", f, works=[it["wid"]] if it["wid"] else None)
        if d == dt.date(2013, 5, 21):
            mats.append("SS400 t9 端材（工作室在庫）")
        if d == dt.date(2013, 5, 22):
            mats.append("SS400 t9 端材　溶接棒 B-14 φ3.2 約2kg")
        mt = "　".join(mats)
        if mt:
            _hand(dr, (250, y2 + 60), mt[:40], prng, 32)
            rec.addp("item", {"page": pi + 1, "日付": d.isoformat(), "使用材料": mt[:40]})
        # 丸囲み・書き込み
        if prng.random() < 0.5 and lst:
            yy = y0 + 0 * rh
            dr.ellipse((380, yy + 2, 1000, yy + rh - 2), outline=(200, 40, 40), width=3)
        memo = prng.choice(["", "", "済", "部品 発注済", "係長 確認", "次回 要点検"])
        if memo:
            _hand(dr, (100, y2 + 250), memo, prng, 34, ink=(170, 30, 30))
            rec.addp("item", {"page": pi + 1, "日付": d.isoformat(), "所見・連絡": memo})
        im = _scan(im, prng)
        imgs.append(im)
    # JPEG で PDF に（1 頁ずつ 100KB 前後に合わせる）
    w = PdfWriter()
    for im in imgs:
        q = 38
        bio = io.BytesIO()
        im.save(bio, "PDF", resolution=200, quality=q)
        bio.seek(0)
        for p in PdfReader(bio).pages:
            w.add_page(p)
    w.add_metadata({"/Title": "【架空】くぬぎ台環境センター 補修作業日報（直営 %s スキャン）" % ("平成%d年度" % (fy - 1988)),
                    "/Author": "くぬぎ台市 環境センター 直営補修班 d1（H29 d2 スキャン）", "/Subject": "作業日報（手書きスキャン）"})
    with open(path, "wb") as fh:
        w.write(fh)
    rec.save()
    return len(imgs)
