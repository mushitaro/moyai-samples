# -*- coding: utf-8 -*-
"""answer/world.json：機器・実際の作業・資料の行との結び・どの資料にも無い作業・計画の欠け・施設の profile。"""
import json
import os
from datetime import date
from common import ROOT
from world import ALL, BY_CODE, maker_of
from works import WORKS, UNRECORDED, BY_ID
from contracts_full import get
import gen_ledger_full as L
from gen_report_full import maker_name

ABBR = {"誘引通風機": ["IDF", "誘引ファン"], "押込送風機": ["FDF"], "バグフィルタ本体": ["BF", "バグ"], "3次過熱器": ["3SH"], "2次過熱器": ["2SH"],
        "1次過熱器": ["1SH"], "エコノマイザ": ["ECO"], "空冷式復水器": ["ACC"], "空冷復水器ファン減速機": ["ACCファン減速機", "減速機"],
        "燃焼火格子": ["ストーカ（燃焼段）", "燃焼段"], "乾燥火格子": ["ストーカ（乾燥段）"], "後燃焼火格子": ["ストーカ（後燃焼段）"]}


def fy(d):
    d = date.fromisoformat(d)
    return d.year if d.month >= 4 else d.year - 1


def write(links, actlinks):
    L.ASOF, L.VER = date(2026, 11, 5), 2026
    eq = []
    for e in ALL:
        iy, im = L.install_ym(e)
        eq.append(dict(code=e["code"], name=e["name"], unit={"1": "1号炉", "2": "2号炉", "0": "共通"}[e["unit"]], category=e["cat"],
                       system=e["sub"], kind=e["kind"], motorOf=e["motor_of"], role=e["role"] or None, importance=e["imp"],
                       spec=e["spec"], maker=maker_of(e), installed=f"{iy}-{im:02d}",
                       names={"機器台帳・計画・SPC様式報告書・日報(2019-)": e["name"], "メーカー様式報告書": maker_name(e["code"]),
                              "略称（日報・メモ）": ABBR.get(e["base"], [])}))
    # 計画の実績セル → 作業
    by = {}
    for w in WORKS:
        by.setdefault((w["code"], fy(w["start"])), []).append(w["id"])
    for (fname, sheet, r, code, y, mark) in actlinks:
        links.append(dict(file=fname, sheet=sheet, r=r, year=y, works=by.get((code, y), [])))
    contracts = []
    for c in get():
        S, E = c["period"]
        contracts.append(dict(key=c["key"], name=c["name"], fy=c["fy"], no=c["contract_no"], start=S.isoformat(), end=E.isoformat(),
                              status=c["status"], foreman=c["foreman"], spc=c["spc"],
                              events=[{"date": d.isoformat(), "text": t} for d, t in c["events"]],
                              changes=[{"what": a, "area_m2": b} for a, b in c.get("changes", [])]))
    pend = {c["key"] for c in get() if c["status"] != "完了"}
    unrec = [dict(work=w, reason=r) for w, r in UNRECORDED]
    for w in WORKS:
        if w["contract_key"] in pend:
            unrec.append(dict(work=w["id"], reason="報告書提出前（2026 年度下期）のため、計画の実績欄・機器台帳には未反映。工程表・日報にはある"))
    gaps = [
        dict(file="中長期整備計画 sp1 2015–2019年度版", what="計量機（0UK-010〜030）の定期検査（計量法 2年毎）が 2020 年度だけ抜けている", kind="計画の欠け"),
        dict(file="中長期整備計画 sp2 2020–2026年度版（費用集計シート）", what="燃焼設備 2027 年度の費用が乾燥火格子更新（62,000千円）の前倒し分を入れ忘れ。行の和と合計が 62,000 食い違う", kind="費用の小計の食い違い"),
        dict(file="中長期整備計画 両様式", what="受変電の月次点検は年に記号 1 つだけで、月ごとには持たない", kind="法定点検の粗さ"),
        dict(file="中長期整備計画 sp2 版（第4期）", what="0YN-400 脱気器 2030 の計画に旧記号 ◇ が残る（他は ☆）", kind="資料の汚さ"),
        dict(file="工事帳票（2023 1号炉）", what="1KT-110 誘引通風機インバータ盤 盤内精密点検が工程・日報・報告書に無い（PM-NAVI と計画の実績には有る）", kind="記載漏れ"),
    ]
    profile = dict(id="f07", name="【架空】かざはやエコプラザ", owner="かざはや市ほか2市の広域", operator="【架空】かざはや環境パートナーズ株式会社（SPC）",
                   completed="2015-03", capacity="150t/24h×2炉=300t/日", furnace="階段式ストーカ・全連続・低空気比",
                   boiler="4MPa・400℃・17.5t/h", power="6,500kW", gas="消石灰・活性炭＋BF＋触媒脱硝", contract="DBO（運営20年 2015–2034年度）",
                   period="2015–2026年度（12年）", asof="2026-11-13", equipmentCount=len(ALL),
                   counting="SPC 保全システムの保全対象。機器と電動機を別に1台ずつ。主要計装を含む。弁は入れない",
                   people=dict(spc=["sp1（保全 2015–2019）", "sp2（保全 2020–）", "sp3（運転）", "sp4（管理）"],
                               maker=["q1（職長 2015–2018）", "q2（職長 2019–）", "q3（副職長 全停の日報 2021–）"]))
    out = dict(profile=profile, equipment=eq, contracts=contracts, works=WORKS, links=links, unrecorded=unrec, planGaps=gaps)
    with open(os.path.join(ROOT, "answer", "world.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=0)
