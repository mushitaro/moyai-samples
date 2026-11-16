# -*- coding: utf-8 -*-
"""工事報告書（全年度）。r1 様式（H13〜H19：表型）、r2 旧様式（H20〜R1：文章型）、r2 新様式（R2〜：サマリー＋機器別表＋写真）。
Word で作って LibreOffice で文字のある PDF にする。表の見出し行はページごとにくり返す。原価は載せない。"""
import os, re, random
import datetime as dt
import rep_base as RB
from rep_base import _doc, _p, _table, _footer, _header, _cover, _photos, _pdf, body
from rec import Rec, LINKS
from model import WORKS, agent, reporter, city_person, workday
from world import SEED, CLS, wareki, SYS

TAGRE = re.compile(r"\d[A-Z]{2}-\d{3}")

PARTS = {"stoker": ("火格子", "SCH13", "枚"), "furnace": ("キャスタブル", "高アルミナ質", "t"), "sb": ("ランスチューブ・パッキン", "", "式"),
         "sv": ("弁座パッキン類", "", "式"), "bf": ("ろ布", "PPS φ150×6,000", "本"), "cat": ("触媒ブロック", "ハニカム", "ブロック"),
         "fan_big": ("軸受", "22324 CAME4", "個"), "fan": ("軸受・Vベルト", "", "式"), "pump_s": ("メカニカルシール", "", "組"),
         "pump_m": ("メカニカルシール・軸受", "", "組"), "mov": ("グランドパッキン", "", "式"), "cv": ("グランドパッキン・ダイヤフラム", "", "式"),
         "damper": ("軸受", "", "個"), "conv": ("チェーン・スプロケット", "", "式"), "feeder": ("摺動板", "", "枚"), "hyd": ("作動油・フィルタ", "", "式"),
         "inst": ("検出器", "", "台"), "inst_a": ("伝送器", "", "台"), "anal": ("分析計 検出部", "", "式"), "panel": ("リレー・端子台", "", "式"),
         "tower": ("噴霧ノズル", "二流体 SUS316", "本"), "hx": ("ガスケット", "", "枚"), "sh": ("過熱器管", "STB410 φ38.1×4.5t", "本"),
         "eco": ("エコノマイザ管", "STB340 φ38.1", "本"), "scrub": ("ライニング材", "ビニルエステル FRP", "式"), "crane": ("ワイヤロープ", "IWRC φ22", "本"),
         "turbine": ("パッキン類", "", "式"), "tr": ("絶縁油", "", "L"), "comp": ("エレメント類", "", "式"), "misc": ("消耗部品", "", "式")}


def _res(w, e, rng):
    k = w["kind"]
    if k == "INSPECT":
        if e.cls in ("boiler", "sh", "eco") or "肉厚" in w["what"]:
            return "肉厚 最小 %.1fmm" % rng.uniform(2.9, 3.9), "○"
        if "水圧" in w["what"]:
            return "2.6MPa 30 分保持 漏れなし", "◎"
        return rng.choice(["異常なし", "異常なし", "摩耗 軽微", "汚れ 中、清掃", "腐食 軽微"]), rng.choice(["◎", "○"])
    if k == "REPLACE":
        return "新品に更新、試運転 良", "◎"
    if k == "REPAIR":
        return "補修完了", "○"
    return rng.choice(["%s 実施、作動 良" % w["what"], "%s 実施" % w["what"], "整備後 試運転 良"]), "◎"


def _next(w, e):
    c = CLS[e.cls]
    if w["kind"] == "INSPECT":
        return "毎年" if c.get("ins") == 1 else "－"
    if c.get("oh"):
        return "%d 年後" % c["oh"]
    return "－"


def _wy(d, west=False):
    if west:
        return "%d年%d月%d日" % (d.year, d.month, d.day)
    if d.year >= 2019 and not (d.year == 2019 and d.month < 5):
        return "令和%d年%d月%d日" % (d.year - 2018, d.month, d.day)
    return "平成%d年%d月%d日" % (d.year - 1988, d.month, d.day)


def _attach(rec, n0, wl):
    for r, ws in zip(rec.records[n0:], wl):
        for w in ws or []:
            LINKS.append(dict(file=rec.file, seq=r["seq"], work=w))


def _rows(c, model):
    out = []
    for t in c.tasks:
        for wid in t.works:
            w = WORKS[wid]
            out.append((t, w, model.tag[w["tag"]]))
    return out


def generic(c, model, path_pdf, tmp):
    rng = random.Random("%s-rep-%s" % (SEED, c.key))
    fname = os.path.basename(path_pdf)
    rec = Rec(fname)
    RB.REC = rec
    st = "r1" if c.fy <= 2007 else ("r2o" if c.fy <= 2019 else "r2n")
    west = st == "r2n"
    title = {"r1": "整備報告書", "r2o": "完了報告書", "r2n": "定期整備報告書" if c.kind != "hoshu" else "工事報告書"}[st]
    d = _doc("【架空】%s %s" % (c.name, title), "汐見重工環境 %s" % reporter(c.fy), "工事報告書")
    no = "SH-KN-%d-%02d" % (c.fy, list(model.contracts).index(c.key) % 97) if st == "r2n" else None
    nm = c.name.split(" くぬぎ台環境センター ")
    nm2 = c.name.replace("%s " % wareki(c.fy), "")
    _cover(d, [(wareki(c.fy), 16), (nm2, 16), ("%s" % title, 22)], no)
    _p(d, ("%d年%d月" % (c.end.year, c.end.month)) if west else _wy(c.end).rsplit("月", 1)[0].rsplit("年", 1)[0] + "年%d月" % c.end.month, 12, align="c")
    _p(d, "汐見重工環境株式会社", 13, True, "c")
    if st != "r1":
        _p(d, "環境プラント本部 保守サービス部", 11, align="c")
    else:
        _p(d, "環境プラント事業部 サービス課", 11, align="c")
    d.add_page_break()
    body()
    _header(d, "【架空】%s %s" % (c.name, title) + ("　%s" % no if no else ""))
    _footer(d, "汐見重工環境株式会社")
    rows = _rows(c, model)
    wk = [t for t in c.tasks if t.kind not in ("city",)]
    a = min(t.a for t in wk)
    b = max(t.b for t in wk)
    ov = [("工事名", c.name), ("工期", "%s ～ %s" % (_wy(c.start, west), _wy(c.end, west))),
          ("作業期間", "%s ～ %s" % (_wy(a, west), _wy(b, west))), ("発注者", "くぬぎ台市 環境部 環境センター"),
          ("受注者", "汐見重工環境株式会社"), ("現場代理人", agent(c.fy)), ("報告者", reporter(c.fy))]
    if c.amount2:
        add = [t for t in c.tasks if t.kind == "add"]
        ov.insert(2, ("変更契約", "%s の追加" % (add[0].name.replace("【追加】", "") if add else "作業")))
    if st == "r1":
        d.add_heading("1. 工事概要", 1)
        _table(d, ["項目", "内容"], ov, [35, 130], 9)
        d.add_heading("2. 作業実施一覧", 1)
        tr = []
        for i, (t, w, e) in enumerate(rows, 1):
            res, j = _res(w, e, rng)
            tr.append((i, ("%d号炉 " % e.unit if e.unit else "") + e.name, w["what"], res, t.note or ""))
        n0 = len(rec.records)
        _table(d, ["No", "機器", "作業内容", "結果", "備考"], tr, [10, 45, 45, 40, 25], 8)
        _attach(rec, n0, [[w["id"]] for t, w, e in rows])
        d.add_heading("3. 使用部品", 1)
        _table(d, ["品名", "仕様", "数量", "使用箇所"], _parts(rows, rng, simple=True), [40, 45, 20, 60], 8)
        d.add_heading("4. 所見", 1)
        _p(d, _shoken(c, rows, rng), 9.5)
    elif st == "r2o":
        d.add_heading("1. 工事概要", 1)
        _table(d, ["項目", "内容"], ov, [35, 130], 9)
        d.add_heading("2. 工事工程", 1)
        dl = [t for t in c.tasks if t.act and t.plan and t.act[1] != t.plan[1]]
        txt = "工事は作業工程表のとおり実施した。%sに炉を立ち下げ、%sから整備を行い、%sに作業を終えた。" % (_wy(c.tasks[0].a), _wy(a), _wy(b))
        if c.kind == "kyotsu":
            txt = "全炉停止（%s〜%s）を中心に、共通設備の点検整備を行った。" % (_wy(c.zentei[0]), _wy(c.zentei[1])) if c.zentei else txt
        if c.kind == "hoshu":
            txt = "%s の工事を、同炉の定期整備の停止期間中に行った。" % c.name
        if dl:
            txt += "なお、%s は %s のため予定より遅れたが、全体の工程に大きな影響はなかった。" % (dl[0].name, dl[0].note or "作業量の増")
        else:
            txt += "工程の遅れはなかった。"
        _p(d, txt)
        d.add_heading("3. 整備内容", 1)
        bys = {}
        for t, w, e in rows:
            bys.setdefault(SYS[e.sys][0], []).append((t, w, e))
        for i, (s, lst) in enumerate(bys.items(), 1):
            d.add_heading("3.%d %s" % (i, s), 2)
            parts = []
            for t, w, e in lst[:8]:
                res, j = _res(w, e, rng)
                parts.append("%sは%sを行った（%s）。" % (e.name.replace(" ", ""), w["what"], res))
            if len(lst) > 8:
                parts.append("ほか %d 件の点検整備を行い、異常は認められなかった。" % (len(lst) - 8))
            n0 = len(rec.records)
            _p(d, "".join(parts))
            _attach(rec, n0, [[w["id"] for t, w, e in lst[:8]]])
        bl = [(t, w, e) for t, w, e in rows if e.cls in ("boiler", "sh", "eco") and w["kind"] == "INSPECT"]
        if bl:
            d.add_heading("4. 測定結果", 1)
            _p(d, "ボイラ肉厚測定結果（最小値、単位 mm）")
            mr = []
            for t, w, e in bl:
                des = {"boiler": 3.5, "sh": 4.5, "eco": 3.5}[e.cls]
                now = round(des - rng.uniform(0.3, 0.9), 1)
                lim = {"boiler": 2.4, "sh": 2.6, "eco": 2.2}[e.cls]
                mr.append((e.name, "%s" % des, "%.1f" % (now + rng.choice([0.1, 0.2, 0.3])), "%.1f" % now, "%.1f" % lim,
                           "良" if now - lim > 0.8 else "良（減肉傾向）"))
            _table(d, ["部位", "設計肉厚", "前回", "今回", "管理値", "判定"], mr, [40, 22, 22, 22, 22, 37], 8.5)
        d.add_heading("5. 交換部品・材料", 1)
        _table(d, ["品名", "仕様", "数量", "備考"], _parts(rows, rng), [35, 55, 20, 55], 8.5)
        d.add_heading("6. 所見", 1)
        _p(d, _shoken(c, rows, rng))
        _p(d, "以上", 10, align="r")
        _p(d, "（別冊）写真帳 1 冊", 9)
    else:
        d.add_heading("1. 報告サマリー", 1)
        _p(d, "工期 %s〜%s、作業 %s〜%s。主な結果を下表に示す（判定 ◎良 ○軽微 △要監視 ×要対策）。" % (
            "%d/%d/%d" % (c.start.year, c.start.month, c.start.day), "%d/%d" % (c.end.month, c.end.day), "%d/%d" % (a.month, a.day), "%d/%d" % (b.month, b.day)), 9)
        sr = []
        for i, (t, w, e) in enumerate(rows, 1):
            res, j = _res(w, e, rng)
            sr.append((i, "%s（%s）" % (e.name, e.tag), w["what"], res, j, _next(w, e)))
        _table(d, ["No", "対象機器（タグ）", "実施作業", "結果", "判定", "次回推奨"], sr, [8, 44, 40, 40, 9, 24], 7.5)
        d.add_heading("2. 工事概要", 1)
        _table(d, ["項目", "内容"], ov, [35, 130], 9)
        d.add_heading("3. 工程実績", 1)
        pr = []
        for t in c.tasks:
            if t.kind == "city" and "立上げ" not in t.name and "立下げ" not in t.name:
                continue
            p = "%d/%d〜%d/%d" % (t.plan[0].month, t.plan[0].day, t.plan[1].month, t.plan[1].day) if t.plan else "－"
            ac = t.act or t.plan
            q = "%d/%d〜%d/%d" % (ac[0].month, ac[0].day, ac[1].month, ac[1].day)
            pr.append((t.name, p, q, t.note or ("追加作業" if t.plan is None else ("予定どおり" if not t.act else "－"))))
        _table(d, ["作業", "予定", "実績", "差異・理由"], pr, [60, 28, 28, 49], 8)
        d.add_heading("4. 機器別 実施内容", 1)
        main = [(t, w, e) for t, w, e in rows if w["kind"] != "INSPECT"][:4] or rows[:2]
        for i, (t, w, e) in enumerate(main, 1):
            d.add_heading("4.%d %s（%s）" % (i, e.name, e.tag), 2)
            res, j = _res(w, e, rng)
            _table(d, ["区分", "内容"], [("作業前", rng.choice(["摩耗・汚れ 中", "前回整備から %d 年経過" % CLS[e.cls].get("oh", 4) if CLS[e.cls].get("oh") else "点検周期", "運転中 異常なし（周期整備）"])),
                                       ("実施", w["what"]), ("作業後", res), ("判定", j)], [22, 143], 8.5, "EDEDED")
            if i <= 2:
                _photos(d, ["%s 作業前" % e.name, "%s 作業後" % e.name], tmp, abs(hash((c.key, i))) % 100000)
        d.add_heading("5. 交換部品・支給品", 1)
        _table(d, ["品名", "仕様", "数量", "区分", "使用先・扱い"], _parts(rows, rng, kubun=True), [26, 38, 28, 24, 49], 7.5)
        d.add_heading("6. 所見", 1)
        _p(d, _shoken(c, rows, rng), 9.5)
        _p(d, "以上", 10, align="r")
    dx = os.path.join(tmp, os.path.splitext(fname)[0] + ".docx")
    d.save(dx)
    _link(rec, c)
    rec.save()
    RB.REC = None
    return _pdf(dx, path_pdf)


def _parts(rows, rng, simple=False, kubun=False):
    out = []
    seen = set()
    for t, w, e in rows:
        if w["kind"] not in ("OVERHAUL", "REPLACE", "REPAIR") or e.cls not in PARTS:
            continue
        nm, spec, unit = PARTS[e.cls]
        key = (nm, e.tag)
        if key in seen:
            continue
        seen.add(key)
        if w["kind"] == "REPLACE":
            nm, spec, unit, q = e.name + "（本体）", e.spec, "台", 1
        else:
            q = {"枚": rng.randint(12, 44), "本": rng.randint(2, 12), "t": round(rng.uniform(0.5, 1.6), 1), "個": 2, "L": 200}.get(unit, 1)
        qs = "%s %s" % (q, unit)
        if e.cls == "bf" and w["kind"] == "OVERHAUL":
            qs = "288 本"
        if simple:
            out.append((nm, spec or "－", qs, "%s（%s）" % (e.name, e.tag)))
        elif kubun:
            kb = "市支給品" if (e.cls == "bf" and w["kind"] == "OVERHAUL") else "当社手配"
            out.append((nm, spec or "－", qs, kb, e.tag))
        else:
            out.append((nm, spec or "－", qs, "当社手配（%s）" % e.tag))
    if not out:
        out.append(("（交換部品なし）", "－", "－", "－") if not kubun else ("（交換部品なし）", "－", "－", "－", "－"))
    return out[:24]


def _shoken(c, rows, rng):
    worst = [(t, w, e) for t, w, e in rows if e.cls in ("sh", "boiler", "stoker", "fan_big", "scrub")]
    if worst:
        t, w, e = rng.choice(worst)
        return "%s は %s。次回以降も点検・測定を継続し、整備・取替の時期を判断されたい。" % (e.name, rng.choice(["摩耗が進んでいる", "減肉傾向がみられる", "おおむね良好である"]))
    return "各機器とも整備後の試運転で異常は認められなかった。"


def _link(rec, c):
    """表の行に書いた機器（タグ・名前）から、その契約の作業へ結ぶ"""
    byt = {}
    for t in c.tasks:
        for wid in t.works:
            byt.setdefault(WORKS[wid]["tag"], []).append(wid)
    for r in rec.records:
        txt = " ".join(str(v) for v in r["fields"].values())
        ws = []
        for tg in TAGRE.findall(txt):
            ws += byt.get(tg, [])
        for wid in dict.fromkeys(ws):
            LINKS.append(dict(file=rec.file, seq=r["seq"], work=wid))


def sample(kind, c, path_pdf, tmp):
    fname = os.path.basename(path_pdf)
    rec = Rec(fname)
    RB.REC = rec
    {"H25-1": RB.make_old, "R6-2": RB.make_new_2, "R6-C": RB.make_new_c}[kind](path_pdf, tmp)
    _link(rec, c)
    rec.save()
    RB.REC = None
