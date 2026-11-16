# -*- coding: utf-8 -*-
"""工程から日報の行を作る。作業の種類と工程の中の位置（初日・途中・最終日）で書き方を変え、測定値などを入れる。"""
import random
import datetime as dt
from model import WORKS, workday, wdays, next_workday
from world import SEED

WX = {1: "晴晴曇雪", 2: "晴晴曇雪", 3: "晴曇雨", 4: "晴曇雨", 5: "晴晴曇雨", 6: "雨曇雨晴", 7: "晴雨曇晴", 8: "晴晴曇雨",
      9: "晴雨曇", 10: "晴晴曇雨", 11: "晴晴曇", 12: "晴曇晴"}


def weather(d):
    r = random.Random("w%s" % d.isoformat())
    w = r.choice(WX[d.month])
    if r.random() < 0.15:
        w = w + "のち" + r.choice("晴曇雨")
        if w[0] == w[-1]:
            w = w[0]
    return w


FAM = {
    "stoker": "stoker", "furnace": "refr", "boiler": "ut", "sh": "tube", "eco": "tube", "sb": "sb", "sv": "sv",
    "bf": "bf", "cat": "cat", "scrub": "lining", "fan_big": "fan", "fan": "fan", "pump_s": "pump", "pump_m": "pump",
    "comp": "pump", "mov": "valve", "cv": "valve", "damper": "damper", "inst": "inst", "inst_a": "inst", "anal": "anal",
    "panel": "panel", "dcs": "panel", "camera": "inst", "conv": "conv", "feeder": "conv", "hyd": "hyd", "tank": "vessel",
    "hx": "vessel", "tower": "vessel", "turbine": "turbine", "gen": "turbine", "acc": "fan", "crane": "crane",
    "hoist": "crane", "tr": "panel", "egen": "pump", "struct": "struct", "misc": "conv", "scale": "inst",
    "door": "hyd", "hopper": "vessel", "burner": "valve",
}

P = {
    "stoker": (["ストーカ 開放、火格子 摩耗測定（最大 {a}mm）", "火格子 外観点検・打診、脱落なし"],
               ["火格子 取外し {k} 枚", "新火格子 取付け {k} 枚", "可動梁・サイドシール 点検、摺動面 清掃", "火格子 取付けボルト 増締め、隙間 調整（{g}mm）"],
               ["空打ち確認、異音・引っかかりなし", "取付け完了、隙間 最終確認"]),
    "refr": (["損傷部 確認・マーキング（{s}m²）", "炉内 清掃、クリンカ除去"],
             ["損傷部 はつり {s}m²", "アンカー 溶接 {k} 本、型枠", "キャスタブル 施工 {s}m²"],
             ["養生、型枠 撤去", "補修部 目視確認"]),
    "ut": (["外面清掃（ワイヤブラシ）、測定点マーキング {n} 点", "炉内・ボイラ 清掃、スケール除去"],
           ["肉厚測定（UT）{n} 点", "水管 外観点検、膨出・割れなし", "測定値 再確認（最小部 {t}mm）"],
           ["測定結果 整理（最小 {t}mm）、市へ報告", "測定記録 まとめ"]),
    "tube": (["取替範囲 マーキング、保温材 撤去", "管 切断準備、養生"],
             ["既設管 切断・撤去、開先加工", "新管 取付け・TIG 溶接 {k} 箇所", "溶接部 放射線透過試験 {k} 箇所"],
             ["保温復旧、検査記録 整理", "溶接部 PT、合格"]),
    "sb": (["スートブロワ 取外し、ランスチューブ 曲がり測定", "駆動部 分解"],
           ["パッキン・軸受 交換", "ランスチューブ 交換", "ノズル 清掃"],
           ["組付け、作動確認（前進・後退）", "リミットスイッチ 調整、作動良"]),
    "sv": (["安全弁 取外し・分解", "弁体 外観点検"], ["弁座 すり合わせ", "ばね 点検、組立"], ["吹出し圧力 調整（{p}MPa）", "吹出し試験 立会（市）、良"]),
    "bf": (["BF {r} 室 ろ布 取外し {k} 本", "ろ布 目視点検、破損 {j} 本"],
           ["ケージ 点検・清掃（曲がり {j} 本）", "新ろ布 取付け {k} 本", "パルス弁 作動点検"],
           ["蛍光粉 漏れ試験、漏れなし", "差圧計 導圧管 清掃、復旧"]),
    "cat": (["触媒反応塔 開放、触媒 入口面 点検", "触媒 抜取り（試験用 {k} 個）"],
            ["触媒 ブロック 取外し {k} 段", "新触媒 装填 {k} ブロック", "反応塔 内部 清掃"],
            ["閉止、気密確認", "装填記録 整理"]),
    "lining": (["既設ライニング 調査（ふくれ・ピンホール）", "養生・換気設備 設置"],
               ["既設ライニング はつり {s}m²", "素地調整（ブラスト）", "ライニング 施工 {s}m²", "トップコート 施工"],
               ["スパーク試験（ピンホール なし）", "膜厚測定 {n} 点（平均 {t}mm）"]),
    "fan": (["ケーシング 開放", "羽根車 外観点検、ダスト付着 清掃"],
            ["羽根車 摩耗測定（最大 {t}mm）", "軸受 交換", "羽根車 清掃・肉盛", "Vベルト・カップリング 点検"],
            ["芯出し、試運転（振動 {v}mm/s）", "組立、手回し確認"]),
    "pump": (["分解、羽根車・ケーシング 点検", "取外し、工場へ搬出"],
             ["羽根車・ライナ 摩耗 確認（{t}mm）", "メカシール 交換", "軸受 交換、軸 振れ測定（{g}mm）"],
             ["組立・芯出し（{g}mm）", "試運転、吐出圧 確認"]),
    "valve": (["取外し・分解", "作動確認（全開・全閉）"], ["弁座・弁体 すり合わせ", "グランドパッキン 交換", "アクチュエータ 点検"],
              ["組付け、作動確認 良", "開度 調整"]),
    "damper": (["ダンパ 作動確認、軸受 点検", "ケーシング 内部 清掃"], ["軸受 給脂、グランド 交換", "羽根 ダスト除去"], ["開度 調整、作動 良"]),
    "inst": (["既設 取外し、配線 確認", "検出部 清掃"], ["新品 取付け、配線", "校正（{n} 点）"], ["ループ確認、指示 良"]),
    "anal": (["分析計 サンプリング系 点検", "フィルタ・ポンプ 交換"], ["ゼロ・スパン 校正", "検出器 交換"], ["指示 確認、DCS と照合"]),
    "panel": (["盤内 清掃・増締め", "停電・検電"], ["絶縁抵抗測定（電動機 {n} 台）", "保護継電器 試験", "シーケンス 試験"], ["復電、各部 確認"]),
    "conv": (["点検口 開放、内部 清掃", "チェーン・スクリュー 摩耗 点検"], ["摩耗部品 交換", "駆動部 減速機 油交換", "張り 調整"], ["試運転、異音なし"]),
    "hyd": (["油圧 ユニット 点検、作動油 分析 採取", "シリンダ 漏れ 点検"], ["フィルタ 交換、作動油 補給", "パッキン 交換"], ["作動確認、油圧 {p}MPa"]),
    "vessel": (["マンホール 開放、内部 清掃", "内部 換気・酸欠測定"], ["内面 点検、肉厚測定 {n} 点", "ノズル・内部品 点検"], ["閉止、ガスケット 交換", "水張り・気密 確認"]),
    "turbine": (["保温 撤去、上半ケーシング 開放準備", "上半ケーシング 吊上げ（天井クレーン）"],
                ["動翼・静翼 外観点検", "動翼 MT（指示なし）", "軸受・ジャーナル 点検", "調速装置 点検", "ラビリンス パッキン 隙間測定"],
                ["上半 組立、ボルト 締付け", "無負荷 試運転、振動 {v}μm"]),
    "crane": (["ワイヤロープ 摩耗・素線切れ 点検", "ブレーキ 分解"], ["ワイヤロープ 交換", "バケット 爪・シリンダ 点検", "走行車輪 点検"], ["荷重試験、作動 良"]),
    "struct": (["外観点検（腐食・塗装）", "点検用 仮設"], ["肉厚測定 {n} 点", "腐食部 ケレン・補修塗装"], ["点検記録 整理"]),
}
SCAF = (["足場組立（炉内）", "足場組立（ボイラ {k} パス）", "足場組立（BF 天井部）"], ["足場 盛替え"], ["足場解体（炉内）", "足場解体・搬出"])
ELEC = (["MCC 点検", "現場盤 清掃・増締め"], ["絶縁抵抗測定（電動機 {n} 台）", "計器 校正（温度計 {k} 点）", "圧力計・流量計 校正"],
        ["インターロック 試験（{k} 項目）", "DCS 入出力 確認"])
INSP = (["点検口 開放、目視点検", "内部 清掃"], ["点検、異常なし", "摩耗・腐食 点検、軽微"], ["閉止・復旧"])
TRIAL = (["試運転 立会、各部 漏れ・異音 点検", "燃焼 調整"], ["負荷運転 確認", "排ガス 値 確認"], ["片付け、残材搬出"])
HOSHU = {"h01": (["資材搬入、養生", "仮設 換気・照明 設置"],), "h05": (["復旧、片付け", "残材 搬出"],)}


def _fill(t, rng):
    return t.format(a=rng.randint(6, 14), k=rng.randint(6, 48), g=rng.choice([2, 3, 4]), s=round(rng.uniform(0.6, 3.5), 1),
                    n=rng.randint(20, 140), t=round(rng.uniform(2.8, 4.2), 1), p=round(rng.uniform(2.0, 2.3), 2),
                    r=rng.randint(1, 6), j=rng.randint(0, 4), v=round(rng.uniform(0.8, 2.4), 1))


def phrase(task, eqname, i, n, rng):
    fam = FAM.get(task.cls, "conv") if task.cls else None
    if task.kind == "scaf":
        lst = SCAF[2] if "解体" in task.name else SCAF[0] + SCAF[1]
        return _fill(rng.choice(lst), rng), (100 if i == n - 1 else min(95, int(100.0 * (i + 1) / n)))
    elif task.kind == "trial":
        pool = TRIAL
    elif task.crew == "E" and (task.kind == "std" or fam in ("panel", None)):
        pool = ELEC
    elif task.kind == "std" and task.crew in ("A", "C", "D") and "肉厚" not in task.name:
        pool = INSP
    elif "肉厚" in task.name or "水圧" in task.name:
        pool = P["ut"] if "水圧" not in task.name else (["水圧試験 準備（仮設ポンプ・盲板）"], ["水圧試験 {p}MPa 保持 30 分"], ["水圧試験 合格、復旧"])
        if "水圧" in task.name:
            return "水圧試験 %.1fMPa 30分保持、漏れなし（市 立会）" % rng.choice([2.6, 2.6, 2.5]), 100
    else:
        pool = P.get(fam, INSP)
    if task.kind == "hoshu" and task.id in HOSHU:
        pool = (HOSHU[task.id][0], HOSHU[task.id][0], HOSHU[task.id][0])
    ph = 0 if i == 0 else (2 if i == n - 1 else 1)
    if n == 1:
        ph = rng.choice([0, 2])
    lst = pool[ph] if len(pool) > ph else pool[-1]
    txt = _fill(rng.choice(lst), rng)
    prog = 100 if i == n - 1 else min(95, int(round(100.0 * (i + 1) / n / 5)) * 5 + rng.choice([0, 0, 5, -5]))
    return txt, max(5, prog)


def place_of(task, model, unit):
    if task.tags:
        e = model.tag.get(task.tags[0])
        if e:
            pre = ("%d号炉 " % e.unit) if e.unit else ""
            more = " 他" if len(task.tags) > 1 and task.kind != "std" else ""
            return pre + e.name + more, e
    if task.kind == "scaf":
        return ("%d号炉 炉内・ボイラ" % unit) if unit else "共通 タービン室", None
    if task.crew == "E":
        return ("%d号炉 MCC・計装" % unit) if unit else "受変電・共通盤", None
    return ("%d号炉 現場" % unit) if unit else "共通", None


def gen_daily(c, model):
    """c.daily を作る。{date: dict(weather, note, rows=[(tid, crew, place, text, n, h, prog)])}"""
    rng = random.Random("%s-%s" % (SEED, c.key))
    tasks = [t for t in c.tasks if t.kind != "city"]
    if not tasks:
        return
    a = min(t.a for t in tasks)
    b = max(t.b for t in tasks)
    fixed = getattr(c, "daily_fixed", {}) or {}
    days = wdays(a, b)
    out = {}
    for d in days:
        if d in fixed:
            out[d] = dict(fixed[d])
            out[d]["fixed"] = True
            continue
        rows = []
        for t in tasks:
            if not (t.a <= d <= t.b) or not workday(d):
                continue
            td = wdays(t.a, t.b)
            i, n = td.index(d), len(td)
            place, e = place_of(t, model, c.unit)
            txt, prog = phrase(t, e.name if e else "", i, n, rng)
            if t.kind == "add" and not txt.startswith("【追加】"):
                txt = "【追加】" + txt
            nn = max(2, t.size + rng.choice([-1, 0, 0, 1]))
            h = rng.choice([8, 8, 8, 8, 8, 9, 10]) if t.kind not in ("trial",) else rng.choice([4, 6, 8])
            rows.append((t.id, t.crew, place, txt, nn, h, prog))
        if not rows:
            continue
        note = ""
        for t in tasks:
            if t.act and t.plan and t.plan[1] == d and t.act[1] > d and t.note:
                note = "%s：%s（予定より遅れ）。" % (t.name, t.note)
            if t.kind == "add" and t.a == d:
                note = "%s を追加で実施（開放後の点検で損傷を発見し、市と協議）。" % t.name
            if "水圧" in t.name and t.a == d:
                note = (note + " " if note else "") + "ボイラ水圧試験 市立会。"
        if c.zentei and c.zentei[0] == d:
            note = (note + " " if note else "") + "本日から全炉停止。"
        if not note and rng.random() < 0.18:
            note = rng.choice(["朝礼で高所作業の注意を確認。", "市 施設係 巡回。", "資材 搬入（午後）。", "炉内 酸欠測定 O2 20.9%。",
                               "熱中症 対策で休憩を増やした。" if d.month in (7, 8, 9) else "強風のため屋外の揚重を中止。"])
        out[d] = dict(weather=weather(d), note=note, rows=rows)
    c.daily = out
    # 工数を作業に配る
    md = {}
    for d, v in out.items():
        for r in v["rows"]:
            md[r[0]] = md.get(r[0], 0) + r[4] * r[5] / 8.0
    for t in c.tasks:
        if t.works and t.id in md:
            per = md[t.id] / len(t.works)
            for w in t.works:
                WORKS[w]["manDays"] = round(WORKS[w]["manDays"] + per, 1)
        elif t.works:
            for w in t.works:
                WORKS[w]["manDays"] = WORKS[w]["manDays"] or 0.0
