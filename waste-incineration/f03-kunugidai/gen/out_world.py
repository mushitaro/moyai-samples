# -*- coding: utf-8 -*-
"""answer/world.json を書く。"""
import os, json
import xlhelp as X
import model as M
from rec import LINKS
from world import SYS, MAKERS, KIKAN_FY, DXN_FY, wareki
from f_ledger import K2_NAMES

ALIAS_DOC = {
    "誘引通風機": {"旧台帳（H15・H26・R3）": "誘引送風機", "旧報告書（r1・r2旧）": "IDF（誘引通風機）", "工程表・日報": "誘引通風機"},
    "バグフィルタ": {"旧台帳": "ろ過式集じん器", "日報・工程表": "BF", "報告書": "バグフィルタ"},
    "湿式洗煙塔": {"旧台帳": "洗煙塔", "旧報告書": "洗煙塔", "工程表": "洗煙塔"},
    "触媒反応塔": {"旧台帳": "脱硝反応塔", "工程表": "触媒"},
    "ごみクレーン 1号": {"旧台帳": "ごみクレーン No.1"}, "ごみクレーン 2号": {"旧台帳": "ごみクレーン No.2"},
    "減温塔": {"旧台帳": "ガス減温塔"}, "蒸気式空気予熱器": {"旧台帳": "空気予熱器"}, "灰溶融炉本体": {"旧台帳": "灰溶融炉"},
    "主灰押出装置": {"直営日報（d1）": "灰押出装置"},
}

PLAN_DEFECTS = [
    dict(file="中長期整備計画_H23_長寿命化計画原案_k2様式.xlsx", where="小計（燃焼ガス冷却設備）H28 列", what="小計が明細の合計より 3,000 千円多い（手入力の小計）"),
    dict(file="中長期整備計画_H23_長寿命化計画原案_k2様式.xlsx / H30版", where="法定点検等", what="作業環境測定（DXN）と計量機定期検査の行が無い"),
    dict(file="中長期整備計画_H30版_k2様式.xlsx", where="小計（排ガス処理設備）H33 列", what="小計が明細の合計より 1,200 千円少ない（行を直して小計を直し忘れ）"),
    dict(file="中長期整備計画_R1版_k3様式.xlsx", where="法定点検シート", what="計量機 定期検査の行が無い（R2 版で追加）"),
    dict(file="中長期整備計画_R4版_k3様式.xlsx", where="総括 合計 R9 列", what="炉別シートの合計より 12.0 百万円多い（総括の手直し漏れ）"),
    dict(file="中長期整備計画_H13_整備年次計画_k1様式.xlsx", where="全体", what="系統単位の計画で、法定点検の行が無い（「別に管理」とだけ注記）。台帳の Excel 化（H15）前のため全台は載らない"),
    dict(file="中長期整備計画（各版）", where="期限超えの扱い", what="大物（2,000 万円以上）の更新は期限を過ぎても 5 年先に置いたまま毎年送られている（周期の超え）"),
    dict(file="中長期整備計画_H30版", where="3号炉 湿式洗煙塔・非常用発電機・2号炉 過熱器", what="周期どおりに置いた計画（H30 更新・R6 OH・H35 取替）が実績と合わない（実績は R1 部分補修→R7 全面更新、R7 更新判断、H30・R6 部分取替）"),
]


def write(m, eqs, files):
    ws = []
    for w in M.WORKS.values():
        kind = w["kind"]
        ws.append(dict(id=w["id"], tag=w["tag"], kind="OVERHAUL" if kind == "REPAIR" else kind, detail="補修" if kind == "REPAIR" else None,
                       what=w["what"], start=str(w["start"]) if w["start"] else None, end=str(w["end"]) if w["end"] else None,
                       manDays=w["manDays"], vendor=w["vendor"], contract=w["contract"], fy=w["fy"],
                       contractName=(m.contracts[w["contract"]].name if w["contract"] in m.contracts else
                                     {"DIRECT": "市直営（契約なし）"}.get(w["contract"], w["contract"]))))
    eq = []
    for e in eqs:
        eq.append(dict(tag=e.tag, name=e.name, unit=e.unit_name, system=SYS[e.sys][0], subsystem=SYS[e.sys][1], duty=e.duty or None,
                       idle=e.idle, installFY=e.install, spec=e.spec, kW=e.kw or None, maker=MAKERS[e.maker], importance=e.imp,
                       legal=e.legal or None, names={"k3台帳・計画": e.name, "k2台帳（H26・R3）": K2_NAMES.get(e.name, e.name),
                                                     "k1台帳（H15）": K2_NAMES.get(e.name, e.name)}, aliases=ALIAS_DOC.get(e.name)))
    cs = []
    for c in m.contracts.values():
        cs.append(dict(key=c.key, fy=c.fy, kind=c.kind, name=c.name, unit=c.unit, start=str(c.start), end=str(c.end), amount=c.amount,
                       amountChanged=c.amount2, docs=bool(c.docs and not getattr(c, "planned_only", False)),
                       plannedOnly=bool(getattr(c, "planned_only", False))))
    unrec = list(m.unrecorded)
    unrec.append(dict(what="H12（2000）以前の整備", reason="庁舎移転時に資料が散逸（未記録）。対象期間は H13 から"))
    unrec.append(dict(what="直営班の定常作業（巡視・給油・清掃）", reason="直営の日報・週報には書いてあるが、整備作業（world の作業）としては数えない"))
    unrec.append(dict(what="d1 手書き日報（H13〜H28）のうち綴りに残らなかった日", reason="H29 のスキャン時に補修作業があった日と年 10 日ほどだけを残した"))
    unrec.append(dict(what="定期整備の一般点検（各機器の目視・給油）", reason="定修仕様書に含まれ、日報・報告書には主な作業だけが出る。工程の標準点検だけを作業として持つ"))
    prof = dict(id="f03", name="【架空】くぬぎ台環境センター", city="【架空】くぬぎ台市", capacity="100t/24h×3炉=300t/日", type="並行揺動式ストーカ・全連続",
                completed="1991-03", boiler="2MPa・250℃・10.5t/h×3", power="3,000kW", flue="BF＋触媒反応塔＋湿式洗煙",
                history=["2000–2002 DXN 恒久対策（1号 H12・2号 H13・3号 H14）", "2003 灰溶融炉追設→2013.10 休止（台帳に残置）", "2011–2014 基幹的設備改良（共通→1→2→3号）"],
                operation="市直営運転（4 班 3 交代）", contracts="炉別定期整備 3 件＋全炉停止時共通 1 件／年＋補修請負・改良工事", period="2001–2026 年度（26 年）",
                recordedAt="2026-11（2026 年度は計画＋上期実績）", equipmentCount=len(eqs), countRule="タグ番号 1 件＝1 台（電動弁・調節弁・主要計装を含む、手動弁は含まない）。灰溶融 41 台は休止区分で残置",
                persons=dict(city=["k1 2001–2010", "k2 2011–2017", "k3 2018–"], direct=["d1 2001–2016", "d2 2017–"], maker=["m1 2001–2012", "m2 2013–"], report=["r1 2001–2007", "r2 2008–"]))
    out = dict(profile=prof, equipment=eq, works=ws, contracts=cs, links=LINKS, unrecorded=unrec, planDefects=PLAN_DEFECTS,
               files=[dict(資料=a, file=b, fy=c) for a, b, c in files])
    with open(os.path.join(X.BASE, "answer", "world.json"), "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=0, default=str)
