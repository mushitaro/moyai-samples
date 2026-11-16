# -*- coding: utf-8 -*-
"""作業日報（本作り）
 R1〜R3：現場の手書き複写式日報をスキャンした PDF（1 日 1 枚、書ききれない日は続紙を 2 枚目に）
 R4〜：事務所（e2）が Excel に清書。1 日 1 シート・縦型"""
import os, io, random, datetime as dt
from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance
from openpyxl import Workbook
from common import *
import world as W
import fullworld as FW

FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
WD = FW.WD
# 現場の略し方（手書きだけ）
ABBR = [("誘引通風機", "IDF"), ("押込送風機", "FDF"), ("バグフィルタ", "BF"), ("二次送風機", "2次FAN"), ("コンベヤ", "CV"),
        ("ボイラ給水ポンプ", "BFP"), ("スートブロワ", "SB"), ("減温塔", "減温塔"), ("クリンカ", "クリンカー")]

def abbr(t):
    for a, b in ABBR:
        t = t.replace(a, b)
    return t

def wrap(t, n):
    out = []
    while len(t) > n:
        out.append(t[:n]); t = "  " + t[n:]
    out.append(t)
    return out

def fname(job, style, ext):
    return "【架空】さくら坂焼却場_作業日報_%s%s_%s.%s" % (W.nendo_name(job.fy), job.c["short"], style, ext)

# ---------------- 手書きスキャン ----------------
def hand(img, xy, text, size, rnd, color=(30, 40, 110)):
    x, y = xy
    for ch in text:
        if ch == " ":
            x += size * 0.45; continue
        s = int(size * rnd.uniform(0.9, 1.08))
        f = ImageFont.truetype(FONT_PATH, s)
        im = Image.new("RGBA", (s + 14, s + 14), (0, 0, 0, 0))
        ImageDraw.Draw(im).text((5, 3), ch, font=f, fill=color + (rnd.randint(200, 245),))
        im = im.rotate(rnd.uniform(-7, 7), resample=Image.BICUBIC)
        img.alpha_composite(im, (int(x), int(y + rnd.uniform(-3, 3))))
        x += s * (0.95 if ord(ch) > 255 else 0.58) + rnd.uniform(-2, 2)
    return x

def form(job, dd, lines, rnd, pageno, cont):
    W_, H_ = 1654, 2339
    img = Image.new("RGBA", (W_, H_), (250, 250, 246, 255))
    dr = ImageDraw.Draw(img)
    fP = lambda s: ImageFont.truetype(FONT_PATH, s)
    ink = (40, 40, 40)
    dr.text((90, 70), "【架空】", font=fP(28), fill=ink)
    dr.text((560, 90), "作　業　日　報" + ("（続紙）" if cont else ""), font=fP(64 if not cont else 54), fill=ink)
    dr.text((1180, 120), "【架空】さくら坂機工株式会社", font=fP(30), fill=ink)
    L, R = 100, W_ - 100
    def box(y0, y1): dr.rectangle([L, y0, R, y1], outline=ink, width=3)
    box(220, 420); dr.line([L, 320, R, 320], fill=ink, width=2)
    dr.text((L + 15, 245), "工事名", font=fP(30), fill=ink); dr.line([L + 150, 220, L + 150, 420], fill=ink, width=2)
    dr.text((L + 15, 345), "月　日", font=fP(30), fill=ink)
    dr.text((L + 420, 345), "月", font=fP(30), fill=ink); dr.text((L + 560, 345), "日（　）", font=fP(30), fill=ink)
    dr.line([900, 320, 900, 420], fill=ink, width=2)
    dr.text((920, 345), "天候　晴・曇・雨・雪", font=fP(30), fill=ink)
    box(420, 620)
    dr.text((L + 15, 450), "作業時間", font=fP(30), fill=ink); dr.text((L + 260, 450), "：　　〜　　：", font=fP(30), fill=ink)
    dr.line([L, 520, R, 520], fill=ink, width=2)
    heads = ["職長", "機械工", "電工", "溶接工", "作業員", "計"]
    cw = (R - L) / 6
    for i, h in enumerate(heads):
        dr.text((L + i * cw + 20, 530), h, font=fP(26), fill=ink)
        if i: dr.line([L + i * cw, 520, L + i * cw, 620], fill=ink, width=2)
    box(620, 1720)
    dr.text((L + 15, 635), "作業内容（場所・数量・測定値）", font=fP(30), fill=ink)
    for k in range(11):
        dr.line([L + 10, 700 + k * 92, R - 10, 700 + k * 92], fill=(120, 120, 140), width=1)
    box(1720, 1960)
    dr.text((L + 15, 1735), "ＫＹ（危険予知）", font=fP(28), fill=ink); dr.line([L, 1840, R, 1840], fill=ink, width=2)
    dr.text((L + 15, 1860), "記入者", font=fP(28), fill=ink); dr.line([1000, 1840, 1000, 1960], fill=ink, width=2)
    dr.text((1020, 1860), "所長確認", font=fP(28), fill=ink)
    dr.text((L, 1990), "※この用紙は現場で記入し、翌朝事務所へ提出（複写の白は事務所、黄は現場控え）", font=fP(24), fill=ink)
    dr.text((R - 160, 2240), "No.%03d" % pageno, font=fP(24), fill=ink)
    d = dd["d"]
    hand(img, (L + 175, 240), job.c["short"].replace("業務委託", "") if len(job.c["short"]) < 16 else job.c["short"][:14], 40, rnd)
    hand(img, (L + 330, 335), str(d.month), 44, rnd); hand(img, (L + 480, 335), str(d.day), 44, rnd); hand(img, (L + 625, 335), WD[d.weekday()], 40, rnd)
    cx = {"晴": 1027, "曇": 1087, "雨": 1147}[dd["w"]]
    dr.ellipse([cx - 24, 337, cx + 26, 390], outline=(30, 40, 110, 220), width=4)
    hand(img, (L + 290, 440), "8", 40, rnd); hand(img, (L + 360, 440), "00", 40, rnd)
    hand(img, (L + 470, 440), "17" if not dd["ot"] else "19", 40, rnd); hand(img, (L + 560, 440), "00", 40, rnd)
    cnt = dict(dd["crew"]); cnt["作業員"] = cnt.pop("普通作業員", 0)
    tot = 0
    for i, k in enumerate(["職長", "機械工", "電工", "溶接工", "作業員"]):
        v = cnt.get(k, 0); tot += v
        hand(img, (L + i * cw + 100, 565), str(v) if v else "—", 44 if v else 36, rnd)
    hand(img, (L + 5 * cw + 100, 565), str(tot), 44, rnd)
    smudged = []
    for k, t in enumerate(lines):
        y = 650 + (k + 1) * 92 - 55
        hand(img, (L + 30, y), t, 38, rnd)
        if rnd.random() < 0.045:
            # インクのにじみ・汚れで読めない
            blot = Image.new("RGBA", (700, 80), (0, 0, 0, 0))
            bd = ImageDraw.Draw(blot)
            for _ in range(40):
                x0, y0 = rnd.randint(0, 640), rnd.randint(0, 50)
                bd.ellipse([x0, y0, x0 + rnd.randint(30, 80), y0 + rnd.randint(15, 30)], fill=(60, 60, 90, rnd.randint(120, 200)))
            blot = blot.filter(ImageFilter.GaussianBlur(6))
            img.alpha_composite(blot, (L + 20, y - 10))
            smudged.append(k)
        if rnd.random() < 0.12:  # 所長のチェック
            dr.line([L + 5, y + 30, L + 18, y + 48, L + 36, y + 8], fill=(200, 40, 40, 220), width=5)
    if not cont:
        hand(img, (L + 300, 1762), dd.get("ky", "炉内 酸欠・落下物 注意"), 36, rnd)
        hand(img, (L + 180, 1870), "f1" if dd["team"] == "A班" else "f2", 44, rnd)
    dr.ellipse([1250, 1855, 1350, 1950], outline=(200, 40, 40, 200), width=5)
    dr.text((1268, 1880), "e1", font=fP(40), fill=(200, 40, 40, 200))
    return img, smudged

def scanify(img, rnd):
    img = img.convert("RGB")
    W_, H_ = img.size
    grad = Image.new("L", (W_, H_)); gd = ImageDraw.Draw(grad)
    side = rnd.random() < 0.5
    for x in range(0, W_, 8):
        f = x / W_ if side else 1 - x / W_
        gd.rectangle([x, 0, x + 8, H_], fill=int(255 - rnd.uniform(18, 34) * f ** 1.5))
    img = Image.composite(img, Image.new("RGB", (W_, H_), (205, 205, 200)), grad)
    dr = ImageDraw.Draw(img)
    if rnd.random() < 0.7:
        fy = int(H_ * rnd.uniform(0.4, 0.55))
        for k in range(6):
            dr.line([0, fy + k, W_, fy + k + rnd.randint(-6, 6)], fill=(225 - k * 6, 225 - k * 6, 220 - k * 6))
    ang = rnd.choice([-1, 1]) * rnd.uniform(3, 8)
    img = img.rotate(ang, resample=Image.BICUBIC, expand=True, fillcolor=(60, 60, 60))
    W2, H2 = img.size
    sh = Image.new("RGB", (W2 + 40, H2 + 40), (40, 40, 40)); sh.paste(img, (rnd.randint(4, 16), rnd.randint(4, 16)))
    img = sh.filter(ImageFilter.GaussianBlur(rnd.uniform(0.8, 1.4)))
    img = ImageEnhance.Brightness(img).enhance(rnd.uniform(0.92, 1.04))
    img.thumbnail((rnd.randint(1568, 2000),) * 2)
    return img, ang

KY = ["入炉時 酸欠・落下物 注意", "クリンカ落下 保護メガネ", "高所 安全帯 使用", "重量物 はさまれ注意", "雨天 足元 すべり", "火気使用 消火器 準備", "回転機 停止・施錠 確認"]

def build_scan(job, rnd):
    pages, recs, info = [], [], []
    seq = 0
    pageno = 0
    for dd in job.days:
        lines = []
        for t in dd["lines"]:
            lines += wrap(abbr(t), 31)
        dd["ky"] = rnd.choice([k for k in KY if dd["w"] == "雨" or "雨" not in k])
        chunks = [lines[i:i + 10] for i in range(0, len(lines), 10)] or [[]]
        cnt = dict(dd["crew"])
        recs.append(dict(seq=seq, kind="item", page=pageno + 1, fields={"日付": dd["d"].isoformat(), "天候": dd["w"], "工事名": job.title, "欄": "人員",
                                                                    "職長": cnt.get("職長", 0), "機械工": cnt.get("機械工", 0), "電工": cnt.get("電工", 0),
                                                                    "溶接工": cnt.get("溶接工", 0), "作業員": cnt.get("普通作業員", 0),
                                                                    "計": sum(cnt.values()), "manDays": dd["md"], "manDaysText": "%d名" % sum(cnt.values()),
                                                                    "作業時間": "8:00〜%s" % ("19:00" if dd["ot"] else "17:00")}))
        seq += 1
        for ci, ch in enumerate(chunks):
            pageno += 1
            img, sm = form(job, dd, ch, rnd, pageno, ci > 0)
            img, ang = scanify(img, rnd)
            buf = io.BytesIO(); q = 40
            img.save(buf, "JPEG", quality=q)
            while buf.getbuffer().nbytes > 125000 and q > 20:
                q -= 6; buf = io.BytesIO(); img.save(buf, "JPEG", quality=q)
            buf.seek(0)
            pages.append(Image.open(buf).convert("RGB"))
            info.append(dict(page=pageno, image="p%03d" % pageno, px=list(img.size), kb=buf.getbuffer().nbytes // 1024, angle=round(ang, 1)))
            # 行（折り返した続きは前の行にまとめる）
            joined = []
            for k, t in enumerate(ch):
                if t.startswith("  ") and joined:
                    joined[-1][0] += t.strip(); joined[-1][1].append(k)
                else:
                    joined.append([t, [k]])
            for t, ks in joined:
                bad = any(k in sm for k in ks)
                recs.append(dict(seq=seq, kind="item", page=pageno, fields={"日付": dd["d"].isoformat(), "天候": dd["w"], "欄": "作業内容",
                                                                            "作業内容": "読めない" if bad else t, "原文": t, "続紙": ci > 0}))
                seq += 1
    fn = os.path.join(OUT, fname(job, "手書きスキャン", "pdf"))
    pages[0].save(fn, "PDF", save_all=True, append_images=pages[1:], resolution=200.0,
                  title="【架空】さくら坂焼却場 %s 作業日報（手書きスキャン）" % job.title, author="【架空】さくら坂機工 事務所", subject="手書き日報スキャン")
    return fn, recs, info


# ---------------- Excel 清書 ----------------
def build_xlsx(job, rnd):
    recs = []
    wb = Workbook(); wb.remove(wb.active)
    seen = {}
    for i, dd in enumerate(job.days):
        d = dd["d"]
        nm = "%d.%d" % (d.month, d.day)
        if dd["team"] != "B班" and any(x["d"] == d and x["team"] == "B班" for x in job.days):
            nm += "A"
        ws = wb.create_sheet(nm)
        setw(ws, {"A": 3, "B": 11, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11})
        cell(ws, "B1", "【架空】", F(8))
        ws.merge_cells("B2:H2"); cell(ws, "B2", "作　業　日　報", F(18, True), CENTER)
        cell(ws, "F3", "【架空】さくら坂機工株式会社", F(9))
        wdx = WD[d.weekday()]
        # 清書のくせ：前のシートをコピーして曜日を直し忘れる（年に 1 回くらい）
        typo = rnd.random() < 0.03
        if typo:
            wdx = WD[(d.weekday() - 1) % 7]
        places = sorted({("%s号炉 " % t.unit if t.unit in ("1", "2") else "共通 ") + t.name.split("（")[0].split("・")[0] for t, *_ in dd["items"]})
        place = dd.get("place") or "、".join(places)[:60]
        rows = [("工事件名", job.title), ("日　付", "%s%d月%d日（%s）" % (W.wareki_k(d.year if d.month >= 1 else d.year).replace("年", "年"), d.month, d.day, wdx)),
                ("天　候", dd["w"]), ("作業時間", "8:00 〜 17:00（休憩 12:00〜13:00）" + ("　残業 17:00〜19:00" if dd["ot"] else "")),
                ("作業責任者", "%s 職長 %s" % (dd["team"], "f1" if dd["team"] == "A班" else "f2")), ("作業箇所", place)]
        r = 4
        for k, v in rows:
            cell(ws, "B%d" % r, k, F(10, True), CENTER, B_ALL, fill("F2F2F2"))
            ws.merge_cells("C%d:H%d" % (r, r)); cell(ws, "C%d" % r, v, F(10), LEFT, B_ALL)
            for c in "DEFGH": ws["%s%d" % (c, r)].border = B_ALL
            r += 1
        base = {"日付": d.isoformat(), "曜日（記載）": wdx, "天候": dd["w"], "班": dd["team"]}
        r += 1; cell(ws, "B%d" % r, "人　員", F(10, True)); r += 1
        for j, h in enumerate(["職　種", "人　数", "時　間", "工数（人日）"]):
            cell(ws, "%s%d" % ("BCDE"[j], r), h, F(9, True), CENTER, B_ALL, fill("F2F2F2"))
        r += 1
        tot = 0
        for job_, n in dd["crew"]:
            otp = job_ == "機械工" and dd["ot"]
            hrs = 8 + (dd["ot"] if otp else 0)
            md = (n * 8 + (min(n, 3) * dd["ot"] if otp else 0)) / 8.0
            tot += md
            for j, v in enumerate([job_, n, "8h＋残 %dh" % dd["ot"] if otp else "8h", md]):
                c = cell(ws, "%s%d" % ("BCDE"[j], r), v, F(9), CENTER, B_ALL)
                if j == 3: c.number_format = "0.00"
            recs.append(dict(sheet=nm, r=r, kind="item", fields=dict(base, 欄="人員", 職種=job_, 人数=n, 時間="8h" if not otp else "8h＋残 %dh" % dd["ot"], manDays=md, manDaysText="%.2f" % md)))
            r += 1
        cell(ws, "B%d" % r, "計", F(9, True), CENTER, B_ALL, fill("F2F2F2"))
        cell(ws, "C%d" % r, sum(n for _, n in dd["crew"]), F(9, True), CENTER, B_ALL, fill("F2F2F2"))
        cell(ws, "D%d" % r, None, F(9), CENTER, B_ALL, fill("F2F2F2"))
        c = cell(ws, "E%d" % r, tot, F(9, True), CENTER, B_ALL, fill("F2F2F2")); c.number_format = "0.00"
        recs.append(dict(sheet=nm, r=r, kind="subtotal", fields=dict(base, 欄="人員計", 人数=sum(n for _, n in dd["crew"]), manDays=tot, manDaysText="%.2f" % tot)))
        r += 2; cell(ws, "B%d" % r, "作業内容", F(10, True)); r += 1
        top = r
        for k, t in enumerate(dd["lines"], 1):
            ws.merge_cells("B%d:H%d" % (r, r))
            cell(ws, "B%d" % r, "%d. %s" % (k, t.strip("　")), F(10), LEFT)
            ws.row_dimensions[r].height = 30 if len(t) > 38 else 18
            recs.append(dict(sheet=nm, r=r, kind="item", fields=dict(base, 欄="作業内容", No=k, 作業内容=t.strip("　"))))
            r += 1
        while r < top + 5:
            ws.merge_cells("B%d:H%d" % (r, r)); cell(ws, "B%d" % r, None, F(10)); r += 1
        box(ws, "B%d:H%d" % (top, r - 1))
        r += 1
        prog = dd.get("prog")
        if not prog:
            ps = []
            for t, ii, n, part in dd["items"]:
                ps.append("%s %d/%d日" % (t.name.split("（")[0][:12], ii + 1, n))
            prog = "、".join(ps[:4]) + (" ほか" if len(ps) > 4 else "")
        note = dd.get("note")
        if note is None:
            ms = [t.memo for t, ii, n, part in dd["items"] if t.memo and ii == n - 1 and "延長" in t.memo]
            note = "、".join(ms)
        for lab, v in (("進　捗", prog), ("特記事項", note or "なし")):
            cell(ws, "B%d" % r, lab, F(10, True), CENTER, B_ALL, fill("F2F2F2"))
            ws.merge_cells("C%d:H%d" % (r, r)); cell(ws, "C%d" % r, v, F(10), LEFT, B_ALL)
            if len(v) > 40: ws.row_dimensions[r].height = 30
            r += 1
        cell(ws, "B%d" % r, "安　全", F(10, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("C%d:H%d" % (r, r))
        ky = "KY：" + rnd.choice([k for k in KY if dd["w"] == "雨" or "雨" not in k])
        prev = seen.get(dd["team"])
        if prev == ky and rnd.random() < 0.6:
            ky = "〃"
        seen[dd["team"]] = ky if ky != "〃" else prev
        cell(ws, "C%d" % r, ky, F(10), LEFT, B_ALL)
        r += 2
        for j, h in enumerate(["清書", "所長", "市 確認"]):
            col = "FGH"[j]
            cell(ws, "%s%d" % (col, r), h, F(8), CENTER, B_ALL, fill("F2F2F2"))
            cell(ws, "%s%d" % (col, r + 1), ["e2", "e1", ("l1" if job.fy <= 2022 else "l2") if rnd.random() < 0.3 else ""][j], F(9, color="C00000" if j else "000000"), CENTER, B_ALL)
            ws.row_dimensions[r + 1].height = 34
        cell(ws, "B%d" % (r + 1), "（原本：現場手書き日報 No.%03d）" % (i + 1), F(8, color="7F7F7F"))
        page(ws, "portrait", 9, 1, 1, area="A1:H%d" % (r + 1))
    props(wb, "【架空】さくら坂焼却場 %s 作業日報（清書）" % job.title, "【架空】さくら坂機工 事務所 e2")
    fn = os.path.join(OUT, fname(job, "清書様式", "xlsx"))
    wb.save(fn)
    return fn, recs


def build(job, rnd):
    if job.fy <= 2021:
        return build_scan(job, rnd)
    fn, recs = build_xlsx(job, rnd)
    return fn, recs, None
