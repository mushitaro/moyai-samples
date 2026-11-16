# -*- coding: utf-8 -*-
"""地元業者（g1、いずみ川設備工業）の補修工事 10 件（2019–2024）。工程 → 週報 → 報告書の材料をここに持つ。"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import rng  # noqa: E402
import jobs as J  # noqa: E402
from model import D, DAY, add_work_days, is_off, next_work  # noqa: E402
from model2 import thickness  # noqa: E402
from works import BYID, CONTRACTS, actual_works  # noqa: E402


def _works(cid, pred=lambda w: True):
    return [w for w in actual_works() if w["contract"] == cid and pred(w)]


def segment(label, stop, steps, r, operator=True):
    """steps：(作業項目, 日数, 作業の一覧, 日ごとの文 or None, 予定と実績の差)。差は 'slip'（1 日遅れ）・'add'（予定に無い）。"""
    T = []
    if operator:
        T.append(dict(name="立下げ・冷却（運転業者）", ps=stop, pe=stop + 2 * DAY, as_=stop, ae=stop + 2 * DAY,
                      works=[], texts=None, note="", han="運転"))
        d = next_work(stop + 3 * DAY)
    else:
        d = next_work(stop)
    pd = d
    for (name, n, ws, texts, diff) in steps:
        if diff == "add":
            a0, a1 = d, add_work_days(d, n)
            T.append(dict(name=name, ps=None, pe=None, as_=a0, ae=a1, works=ws, texts=texts, note="追加", han="g1"))
            d = next_work(a1 + DAY)
            continue
        p0, p1 = pd, add_work_days(pd, n)
        a0 = d
        a1 = add_work_days(d, n + (1 if diff == "slip" else 0))
        T.append(dict(name=name, ps=p0, pe=p1, as_=a0, ae=a1, works=ws, texts=texts, note="", han="g1"))
        pd = next_work(p1 + DAY)
        d = next_work(a1 + DAY)
    if operator:
        last = max(t["ae"] for t in T)
        lp = max(t["pe"] for t in T if t["pe"])
        T.append(dict(name="立上げ（運転業者）", ps=lp + DAY, pe=lp + 2 * DAY, as_=last + DAY, ae=last + 2 * DAY,
                      works=[], texts=None, note="", han="運転"))
    return dict(label=label, tasks=T, span=(T[0]["as_"], max(t["ae"] for t in T)))


def daily_rows(areas, r, place_of):
    rows = []
    for a in areas:
        for t in a["tasks"]:
            if t["han"] == "運転":
                continue
            d, k = t["as_"], 0
            wd = []
            while d <= t["ae"]:
                if not is_off(d):
                    wd.append(d)
                d += DAY
            for k, d in enumerate(wd):
                if t["texts"]:
                    txt = t["texts"][min(k, len(t["texts"]) - 1)]
                else:
                    txt = t["name"] + ("" if len(wd) == 1 else ("（続き）" if 0 < k < len(wd) - 1 else ("" if k == 0 else " 完了")))
                rows.append([d, r.choice(["晴", "晴", "曇", "雨", "晴"]), r.randint(4, 6), 8.0, txt,
                             place_of(a, t), "", [w["wid"] for w in t["works"]], a["label"]])
    rows.sort(key=lambda x: x[0])
    merged = []
    for row in rows:                       # 同じ日に 2 つの作業があれば 1 行にまとめる（週報は 1 日 1 行）
        if merged and merged[-1][0] == row[0]:
            m = merged[-1]
            m[4] += "。" + row[4]
            m[2] = max(m[2], row[2]) + 1
            m[7] = m[7] + row[7]
        else:
            merged.append(row)
    return merged


def boiler_steps(line, fy, ws_line, extra=None, r=None, slip=False):
    th = thickness(line, fy)
    st = [("養生・仮設（照明・換気）", 1, [], ["養生・照明・換気ファン設置。ボイラ マンホール開放"], ""),
          ("ボイラ開放・水洗・ダスト除去", 2, [], ["ボイラ水洗（第1〜第3パス）", "水洗 完了、ダスト払い・排水処理"], "slip" if slip else ""),
          ("水管肉厚測定", 1, [], [f"水管肉厚測定（超音波 84 点）最小 {th:.1f}mm"], "")]
    if extra:
        st += extra
    st += [("性能検査準備・受検", 2, ws_line, ["性能検査準備（付属品点検・安全弁取外し整備）", "ボイラ性能検査 受検 合格"], ""),
           ("復旧・片付け", 1, [], ["マンホール閉止、保温復旧、片付け"], "")]
    return st


def boiler_job(cid, fy, stop1, stop2, extras1=(), extras2=(), slip1=False, cyl_in=2):
    r = rng(cid)
    W = _works(cid)
    b = lambda line: [w for w in W if BYID[w["eq"]]["base"] == "廃熱ボイラ" and BYID[w["eq"]]["line"] == line]
    cyl = [w for w in W if BYID[w["eq"]]["base"] in ("脱気器", "高圧蒸気だめ")]
    ex2 = list(extras2)
    ex2.append(("脱気器・高圧蒸気だめ 開放点検（バイパス運転）", 2, cyl,
                ["脱気器・高圧蒸気だめ マンホール開放、内面点検", "第一種圧力容器 性能検査 受検 合格、閉止"], ""))
    a1 = segment("1号炉", stop1, boiler_steps(1, fy, b(1), list(extras1), r, slip1), r)
    a2 = segment("2号炉", stop2, boiler_steps(2, fy, b(2), ex2, r), r)
    return [a1, a2], W


def find(line, base):
    return next(e["id"] for e in BYID.values() if e["line"] == line and e["base"] == base)


def wsel(W, eid, kind=None):
    return [w for w in W if w["eq"] == eid and (kind is None or w["kind"] == kind)]


def build():
    J_ = []
    # ---------------- R1 1号炉ボイラ水管補修工事（両炉のボイラ性能検査整備を含む）
    cid = "L20190"
    W = _works(cid)
    r = rng(cid)
    rep = wsel(W, "1D-001", "OVERHAUL")
    a1 = segment("1号炉", D(2019, 11, 8), [
        ("養生・仮設（照明・換気・足場）", 1, [], ["養生・照明・換気ファン設置、ボイラ室足場組立"], ""),
        ("ボイラ開放・水洗", 2, [], ["ボイラ マンホール開放、ダスト払い", "ボイラ水洗（第1〜第3パス）"], ""),
        ("漏えい箇所確認・肉厚測定", 1, [], [f"漏えい箇所確認（第1パス 曲がり部 2 本）。肉厚測定 最小 {thickness(1, 2019):.1f}mm"], ""),
        ("水管部分取替（2本）", 4, rep, ["水管 2 本 切断・撤去", "開先加工、新管 仮付け", "新管 溶接（TIG＋被覆アーク）", "溶接 仕上げ、外観検査"], "slip"),
        ("放射線透過試験・水圧試験", 2, [], ["放射線透過試験（継手 4 箇所 合格）", "水圧試験 1.5MPa 30 分 漏れなし"], ""),
        ("性能検査準備・受検", 2, wsel(W, "1D-001", "INSPECT"), ["性能検査準備（安全弁整備）", "ボイラ性能検査 受検 合格"], ""),
        ("復旧・片付け", 1, [], ["マンホール閉止、保温復旧、足場解体"], "")], r)
    a2 = segment("2号炉", D(2020, 1, 17), boiler_steps(2, 2019, wsel(W, "2D-001"), [
        ("脱気器・高圧蒸気だめ 開放点検（バイパス運転）", 2, [w for w in W if w["eq"] in ("0G-003", "0G-006")],
         ["脱気器・高圧蒸気だめ 開放、内面点検", "第一種圧力容器 性能検査 受検 合格"], "")], r), r)
    J_.append(dict(cid=cid, areas=[a1, a2], works=W, amount=7480000, cdate=D(2019, 10, 21), climit=D(2020, 2, 28),
                   target="1号炉 廃熱ボイラ（1D-001）水管、1・2号炉 ボイラ、脱気器（0G-003）・高圧蒸気だめ（0G-006）",
                   summary=["1号炉 ボイラ 漏えい水管 2 本の部分取替（第1パス 曲がり部）", "放射線透過試験・水圧試験",
                            "1・2号炉 ボイラ性能検査 受検整備", "脱気器・高圧蒸気だめ 性能検査 受検整備"],
                   result="水管 2 本を取替、放射線透過試験・水圧試験 合格。1号炉 11/21、2号炉 1/23 にボイラ性能検査 合格。",
                   supply="なし（材料は持込）",
                   tables=[("水管取替記録", ["管番号", "位置", "状態", "処置"],
                            [["1-T-14", "第1パス 曲がり部", "ピンホール漏えい（外面腐食）", "部分取替 1.2m"],
                             ["1-T-15", "第1パス 曲がり部", "減肉 2.8mm（隣接管）", "部分取替 1.2m"]]),
                           ("水管肉厚測定結果（最小値）", ["号炉", "第1パス", "第2パス上部", "第2パス下部", "管理値"],
                            [["1号炉", "3.9mm", "3.9mm", f"{thickness(1, 2019):.1f}mm", "3.2mm"],
                             ["2号炉", "4.0mm", "3.6mm", f"{thickness(2, 2019):.1f}mm", "3.2mm"]])],
                   materials=[["ボイラ水管", "STB340 φ50.8×3.5t", "2.4m", "持込", ""],
                              ["溶接棒", "TIG・被覆アーク", "一式", "持込", ""],
                              ["マンホールガスケット", "", "18枚", "持込", "両炉＋脱気器"]],
                   shoken=["1号炉 第1パス曲がり部は外面腐食が進んでおり、残りの管も次回の開放時に重点的に確認する。",
                           f"2号炉 第2パス下部の最小肉厚は {thickness(2, 2019):.1f}mm で管理値に近い。停止まで年 1 回の測定を続ける。"],
                   place=lambda a, t: f"{a['label']} ボイラ" if "脱気" not in t["name"] else "脱気器室"))
    # ---------------- R2 1号炉誘引通風機補修工事
    cid = "L20200"
    W = _works(cid)
    r = rng(cid)
    a = segment("1号炉", D(2020, 8, 21), [
        ("養生・誘引通風機 開放", 1, [], ["養生、誘引通風機 ケーシング開放、羽根車点検"], ""),
        ("羽根車 取外し・搬出", 1, [], ["羽根車 取外し、工場へ搬出（肉盛・動的バランス）"], ""),
        ("ケーシング内面清掃・補修", 2, [], ["ケーシング内面 付着物除去", "ケーシング 摩耗部 肉盛補修"], ""),
        ("羽根車 据付・軸受交換", 2, W, ["羽根車 搬入・据付", "軸受 2 個交換、給油"], "slip"),
        ("芯出し・試運転・振動測定", 1, [], ["芯出し、試運転、振動測定（2.3mm/s）"], ""),
        ("片付け", 1, [], ["片付け、引渡し"], "")], r)
    J_.append(dict(cid=cid, areas=[a], works=W, amount=4180000, cdate=D(2020, 7, 27), climit=D(2020, 9, 30),
                   target="1号炉 誘引通風機（1F-005）",
                   summary=["誘引通風機 羽根車 取外し・工場での肉盛溶接・動的バランス", "ケーシング内面 補修", "軸受交換、芯出し、試運転"],
                   result="振動値 7.1mm/s → 2.3mm/s（管理値 4.5mm/s）。羽根車の摩耗部を肉盛し、残留不釣合いを許容値内とした。",
                   supply="なし",
                   tables=[("振動測定結果", ["測定点", "方向", "補修前（R2.7）", "補修後（R2.9）", "判定"],
                            [["負荷側軸受", "水平", "7.1", "2.3", "良"], ["負荷側軸受", "垂直", "5.6", "1.9", "良"],
                             ["負荷側軸受", "軸方向", "3.8", "1.2", "良"], ["反負荷側軸受", "水平", "6.2", "2.1", "良"],
                             ["反負荷側軸受", "垂直", "4.9", "1.7", "良"], ["反負荷側軸受", "軸方向", "3.1", "1.0", "良"]]),
                           ("羽根車 羽根厚さ（最小）", ["羽根No.", "補修前(mm)", "補修後(mm)", "新品(mm)"],
                            [[f"{k}", f"{round(6.1 + 0.2 * (k % 3), 1)}", "9.0", "9.0"] for k in range(1, 13)])],
                   materials=[["軸受", "22224", "2個", "持込", ""], ["肉盛溶接材", "耐摩耗用", "一式", "持込", "工場施工"],
                              ["バランスウェイト", "", "3個", "持込", ""]],
                   shoken=["羽根車の摩耗は入口側に集中しており、停止までは現状で使用できる。", "2号炉の誘引通風機は本工事の対象外。振動は 3.0mm/s 前後で推移している（運転業者記録）。"],
                   place=lambda a, t: "1号炉 誘引通風機室"))
    # ---------------- R2 ボイラ性能検査整備工事
    areas, W = boiler_job("L20201", 2020, D(2020, 10, 16), D(2020, 12, 4))
    J_.append(_bjob("L20201", 2020, areas, W, 3080000, D(2020, 9, 28), D(2021, 2, 26)))
    # ---------------- R3 前処理破砕機補修工事
    cid = "L20210"
    W = _works(cid)
    r = rng(cid)
    a = segment("前処理破砕機", D(2021, 6, 14), [
        ("養生・No.1 減速機 取外し", 1, wsel(W, "0A-018"), ["養生、No.1 減速機 取外し（ギヤ破損確認）"], ""),
        ("No.2 減速機 取外し", 1, wsel(W, "0A-019"), ["予備機 No.2 から減速機 取外し、点検"], ""),
        ("No.1 へ据付・芯出し", 2, [], ["No.2 の減速機を No.1 に据付", "芯出し、カップリング締結"], ""),
        ("試運転・負荷運転", 1, [], ["無負荷・負荷運転、異音・温度確認 良好"], "")], r, operator=False)
    J_.append(dict(cid=cid, areas=[a], works=W, amount=2420000, cdate=D(2021, 6, 7), climit=D(2021, 7, 30),
                   target="前処理破砕機 No.1（0A-018）、No.2（0A-019）",
                   summary=["No.1 減速機（ギヤ破損）の取外し", "予備機 No.2 の減速機を取外し No.1 へ転用・据付", "芯出し、試運転"],
                   result="No.1 は転用した減速機で復旧。No.2 は減速機が無いため使用不可（部品取り）。工事中の 4 日間は破砕ごみの貯留分で運転。",
                   supply="No.2 減速機（市の予備機から転用）",
                   tables=[("減速機 点検記録", ["項目", "No.1（取外し品）", "No.2（転用品）", "判定基準"],
                            [["歯面", "2段目 歯欠け 3 枚", "摩耗 軽微", "欠け・ピッチング無し"],
                             ["軸受", "異音・がた有り", "良", "異音無し"], ["潤滑油", "金属粉 多量", "良", "—"],
                             ["バックラッシ", "0.82mm", "0.31mm", "0.5mm 以下"]])],
                   materials=[["減速機", "No.2 予備機より転用", "1台", "市所有", ""], ["潤滑油", "ギヤ油 VG220", "40L", "持込", ""],
                              ["カップリングボルト", "M16", "8本", "持込", ""]],
                   shoken=["No.1 が故障すると前破砕ができず、流動床炉の運転が続けられない。停止まで予備機が無い状態となるため、日常点検で減速機の温度・異音を記録すること。"],
                   place=lambda a, t: "破砕機室"))
    # ---------------- R3 ボイラ性能検査整備工事（1号炉 BF 内部点検、2号炉 散気ノズル目視点検を含む）
    W = _works("L20211")
    ex1 = [("バグフィルタ 内部点検", 1, wsel(W, "1E-007"), ["バグフィルタ 内部点検、ろ布破損 8 本に閉止板"], "")]
    ex2 = [("流動砂抜出・散気ノズル目視点検", 2, wsel(W, "2B-002"),
            ["流動砂 一部抜出", "散気ノズル目視点検（閉塞 11・摩耗 7 本）"], "")]
    areas, W = boiler_job("L20211", 2021, D(2021, 10, 15), D(2021, 12, 3), ex1, ex2)
    j = _bjob("L20211", 2021, areas, W, 3410000, D(2021, 9, 27), D(2022, 2, 25))
    j["summary"] += ["1号炉 バグフィルタ 内部点検（破損ろ布 8 本 閉止）", "2号炉 散気ノズル 目視点検"]
    j["tables"].append(("付帯点検結果", ["号炉", "対象", "結果"],
                        [["1号炉", "バグフィルタ ろ布", "破損 8 本（3室 5本・5室 3本）、閉止板で処置"],
                         ["2号炉", "散気ノズル（一部抜砂して目視）", "閉塞 11 本・摩耗 7 本。全数点検を要す"]]))
    j["shoken"] += ["1号炉バグフィルタの破損ろ布は予備ろ布での部分交換を勧める。", "2号炉散気ノズルは全数点検のうえ取替が必要。予備品の手配に 5 か月程度かかる。"]
    J_.append(j)
    # ---------------- R4 1号炉バグフィルタろ布部分交換工事
    cid = "L20220"
    W = _works(cid)
    r = rng(cid)
    a = segment("1号炉", D(2022, 8, 19), [
        ("養生・バグフィルタ 開放", 1, [], ["養生、BF 点検口開放、パージ"], ""),
        ("ろ布点検・マーキング", 1, [], ["ろ布全数点検、破損・劣化 32 本マーキング（閉止板 8 本含む）"], ""),
        ("ろ布取替（32本）", 2, W, ["ろ布取替 16 本（3室・4室）", "ろ布取替 16 本（5室・6室）、閉止板撤去"], ""),
        ("リーク試験", 1, [], ["蛍光粉によるリーク試験、漏れなし"], ""),
        ("復旧・片付け", 1, [], ["点検口閉止、片付け"], "")], r)
    rows = []
    rr = rng("bf")
    for k in range(32):
        room = [3, 3, 3, 3, 3, 4, 4, 4, 5, 5, 5, 5, 5, 5, 6, 6][k % 16]
        st = "閉止板（R3）" if k < 8 else rr.choice(["破れ", "摩耗", "目詰まり", "破れ"])
        rows.append([k + 1, f"{room}室", f"{rr.randint(1, 8)}列-{rr.randint(1, 10):02d}", st, "取替"])
    J_.append(dict(cid=cid, areas=[a], works=W, amount=1650000, cdate=D(2022, 6, 10), climit=D(2022, 8, 31),
                   target="1号炉 バグフィルタ（1E-007）ろ布", summary=["ろ布全数点検", "破損・劣化ろ布 32 本の取替（予備ろ布）", "リーク試験"],
                   result="ろ布 32 本を取替（R3 閉止の 8 本を含む）。リーク試験で漏れなし。", supply="ろ布 32 本（市の予備ろ布、H28 購入分）",
                   tables=[("ろ布取替記録", ["No", "室", "位置", "状態", "処置"], rows)],
                   materials=[["ろ布", "PTFE 予備品", "32本", "市支給", "残 16 本"], ["リテーナ", "", "4本", "持込", "曲がり"],
                              ["蛍光粉", "", "1kg", "持込", ""]],
                   shoken=["ろ布は 7 年目で、3・5 室に劣化が多い。停止まで残り 2 年半、予備ろ布 16 本で対応できる見込み。"],
                   place=lambda a, t: "1号炉 バグフィルタ"))
    # ---------------- R4 ボイラ性能検査整備工事
    areas, W = boiler_job("L20221", 2022, D(2022, 11, 11), D(2023, 1, 20), slip1=True)
    J_.append(_bjob("L20221", 2022, areas, W, 3080000, D(2022, 10, 3), D(2023, 2, 28)))
    # ---------------- R5 2号炉散気ノズル補修工事（見本の工程・日報のまま）
    cid = "L20230"
    W = _works(cid)
    T = []
    for (no, name, ps, pe, as_, ae, note) in J.J23:
        T.append(dict(name=name, ps=ps, pe=pe, as_=as_, ae=ae, works=W if name.startswith("散気ノズル取替") else [],
                      texts=None, note=note, han="運転" if "運転業者" in name else "g1"))
    a = dict(label="2号炉", tasks=T, span=J.J23_SPAN)
    J_.append(dict(cid=cid, areas=[a], works=W, amount=5940000, cdate=D(2023, 10, 2), climit=D(2023, 11, 30),
                   target="2号炉 散気ノズル（散気管）120 本（市機器番号 2B-002 散気装置）", sample=True,
                   daily=[[d, w, n, h, txt, place, noz, [x["wid"] for x in W] if noz else [], "2号炉"]
                          for (d, w, n, h, txt, place, noz) in J.J23_DAILY]))
    # ---------------- R5 ボイラ性能検査整備工事
    areas, W = boiler_job("L20231", 2023, D(2024, 1, 12), D(2024, 2, 9))
    J_.append(_bjob("L20231", 2023, areas, W, 3190000, D(2023, 10, 16), D(2024, 2, 29)))
    # ---------------- R6 ボイラ性能検査整備補修工事（最終年度、1 件のみ）
    W = _works("L20240")
    ex2 = [("ボイラ給水ポンプ No.1 メカニカルシール交換", 1, wsel(W, "2D-003"), ["給水ポンプ No.1 メカニカルシール交換、試運転 漏れなし"], "add")]
    areas, W = boiler_job("L20240", 2024, D(2024, 10, 11), D(2024, 12, 6), (), ex2)
    j = _bjob("L20240", 2024, areas, W, 3520000, D(2024, 9, 24), D(2025, 2, 28))
    j["summary"].append("2号炉 ボイラ給水ポンプ No.1 メカニカルシール交換（追加）")
    j["materials"].append(["メカニカルシール", "TFB 純正", "1組", "持込", "追加"])
    j["shoken"] = ["本年度末で停止のため、次回の性能検査は受けない。検査証の有効期限は 1号炉 R7.10、2号炉 R7.12。",
                   f"1号炉 最小肉厚 {thickness(1, 2024):.1f}mm、2号炉 {thickness(2, 2024):.1f}mm。停止まで使用に支障なし。"]
    J_.append(j)
    for j in J_:
        j["contract"] = CONTRACTS[j["cid"]]
        j["fy"] = j["contract"]["fy"]
        j["vendor"] = "g1"
        if "daily" not in j:
            j["daily"] = daily_rows(j["areas"], rng("d" + j["cid"]), j["place"])
        _mandays(j)
    return J_


def _bjob(cid, fy, areas, W, amount, cdate, climit):
    th1, th2 = thickness(1, fy), thickness(2, fy)
    r = rng("b" + cid)
    rows = []
    for ln, th in ((1, th1), (2, th2)):
        rows.append([f"{ln}号炉", f"{round(th + 0.3 + r.random() * 0.2, 1)}mm", f"{round(th + 0.1 + r.random() * 0.2, 1)}mm",
                     f"{th:.1f}mm", "3.2mm"])
    return dict(cid=cid, areas=areas, works=W, amount=amount, cdate=cdate, climit=climit,
                target="1・2号炉 廃熱ボイラ（1D-001・2D-001）、脱気器（0G-003）・高圧蒸気だめ（0G-006）",
                summary=["1・2号炉 ボイラ開放・水洗・水管肉厚測定", "ボイラ性能検査 受検整備（安全弁整備・付属品点検）",
                         "脱気器・高圧蒸気だめ 第一種圧力容器 性能検査 受検整備"],
                result="1・2号炉 ボイラ、脱気器・高圧蒸気だめとも性能検査に合格。",
                supply="なし",
                tables=[("水管肉厚測定結果（最小値）", ["号炉", "第1パス", "第2パス上部", "第2パス下部", "管理値"], rows),
                        ("付属品点検", ["号炉", "安全弁 吹出し試験", "水面計", "圧力計", "判定"],
                         [["1号炉", f"{0.96 + 0.01 * (fy % 3):.2f}MPa", "良", "校正済", "良"],
                          ["2号炉", f"{0.97 + 0.01 * (fy % 2):.2f}MPa", "良", "校正済", "良"]])],
                materials=[["マンホールガスケット", "", "18枚", "持込", ""], ["安全弁 パッキン", "", "4組", "持込", ""]],
                shoken=[f"2号炉 第2パス下部の最小肉厚は {th2:.1f}mm で管理値に近い。", "脱気器の内面に軽い腐食あり、停止まで経過観察。"],
                place=lambda a, t: f"{a['label']} ボイラ" if "脱気" not in t["name"] else "脱気器室")


def _mandays(j):
    """週報の人数（人日）を作業へ割り振る。どの作業とも書いていない日（養生・片付けなど）は、
    その工事の主な作業（点検以外、無ければ全部）に按分する。"""
    per = {}
    main = [w["wid"] for w in j["works"] if w["kind"] != "INSPECT"] or [w["wid"] for w in j["works"]]
    for row in j["daily"]:
        ids = sorted(set(row[7])) or main
        for wid in ids:
            per.setdefault(wid, 0)
            per[wid] += row[2] / max(1, len(ids))
    for a in j["areas"]:
        for t in a["tasks"]:
            for w in t["works"]:
                w["start"], w["end"] = t["as_"], t["ae"]
                w["manDays"] = round(per.get(w["wid"], 0), 1)
                w["task"] = (j["cid"], a["label"], t["name"])


if __name__ == "__main__":
    js = build()
    for j in js:
        print(j["cid"], [a["span"] for a in j["areas"]], len(j["daily"]), [w["manDays"] for w in j["works"]][:4])
    lw = [w for w in actual_works() if w["contract"].startswith("L")]
    print(len(lw), sum(1 for w in lw if "manDays" in w))
