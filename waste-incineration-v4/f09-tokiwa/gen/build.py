# -*- coding: utf-8 -*-
"""F09 本作り：2009–2026 年度の全資料と正解を作る。  python3 -I build.py（full/gen/ で）"""
import os, sys, json, shutil, datetime as dt
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import world as W, contracts as C
import g_plan, g_ledger, g_sched, g_daily, g_report, g_order, render, rec
from plan_gen import u_name
from ledger_gen import spc_name

FULL = os.path.normpath(os.path.join(HERE, ".."))
FILES = os.path.join(FULL, "files"); ANS = os.path.join(FULL, "answer")
TMP = os.path.join(HERE, "tmp"); PNG = os.path.join(HERE, "png")
P = "【架空】ときわ北部クリーンセンター_"
OUT = []   # (file, 資料, 年度, 様式, 契約)

def cname(c):
    if c.kind == "teishu": return f"{c.fy}年度{c.furn}号炉定期整備工事"
    if c.kind == "dcs": return "2019年度DCS更新工事"
    return f"基幹的設備改良工事{'R7' if c.fy == 2025 else 'R8'}施工分"

def emit(fn, doc, fy, form, cid=None):
    OUT.append({"file": fn, "資料": doc, "年度": fy, "様式": form, "契約": cid})
    return os.path.join(FILES, fn), fn

def main():
    for d in (FILES, ANS, TMP, PNG):
        if os.path.exists(d): shutil.rmtree(d)
        os.makedirs(d)
    os.makedirs(os.path.join(ANS, "files"))
    # 発注（PO 番号を報告書へ渡すため先に）
    for y in range(2009, 2027):
        if y <= 2015:
            p, fn = emit(P + f"発注台帳_{y}年度_SPC旧購買システム.csv", "発注台帳", y, "SPC旧購買システム"); g_order.old_csv(y, p, fn)
        else:
            p, fn = emit(P + f"発注台帳_{y}年度_SPC新購買システム.csv", "発注台帳", y, "SPC新購買システム"); g_order.new_csv(y, p, fn)
    for fyw, y in (("令和07", 2025), ("令和08", 2026)):
        p, fn = emit(P + f"発注台帳_基幹的設備改良事業{'R7' if y == 2025 else 'R8'}_組合財務会計.csv", "発注台帳", y, "組合財務会計"); g_order.union_csv(fyw, p, fn)
    # 計画
    for V in range(2009, 2027):
        if V <= 2012:
            p, fn = emit(P + f"中長期整備計画_長期修繕計画{V}年度版_SPCsv1初版様式.xlsx", "中長期整備計画", V, "SPC sv1 初版"); g_plan.sv1(V, p, fn)
        elif V <= 2016:
            p, fn = emit(P + f"中長期整備計画_長期修繕計画{V}年度改訂_SPCsv1様式.xlsx", "中長期整備計画", V, "SPC sv1 2013 様式"); g_plan.sv1(V, p, fn)
        else:
            p, fn = emit(P + f"中長期整備計画_長期修繕計画{V}年度改訂_SPCsv2様式.xlsx", "中長期整備計画", V, "SPC sv2 様式"); g_plan.sv2(V, p, fn)
    p, fn = emit(P + "中長期整備計画_令和6年度策定_第2期移行_組合u1様式.xlsx", "中長期整備計画", 2024, "組合 u1"); g_plan.union(2024, p, fn)
    p, fn = emit(P + "中長期整備計画_令和7年度改訂_組合u1様式.xlsx", "中長期整備計画", 2025, "組合 u1"); g_plan.union(2025, p, fn)
    # 台帳
    for V in (2009, 2012, 2015, 2019, 2023):
        p, fn = emit(P + f"機器台帳_{V}年度末_SPC保全システム.xlsx", "機器台帳", V, "SPC 設備ツリー"); g_ledger.spc_ledger(V, p, fn)
    for V in (2024, 2025, 2026):
        tag = "_第2期移行引渡し" if V == 2024 else ""
        p, fn = emit(P + f"機器台帳_令和{V - 2018}年4月1日現在{tag}_組合様式.xlsx", "機器台帳", V, "組合様式"); g_ledger.union_ledger(V, p, fn)
    # 契約ごと（工程表 → 日報 → 報告書）
    for c in C.ALL:
        nm = cname(c)
        if c.kind == "kikan":
            p, fn = emit(P + f"工程表_{nm}_メーカー様式.xlsx", "工程表", c.fy, "メーカー", c.cid); g_sched.maker(c, p, fn)
            p, fn = emit(P + f"作業日報_{nm}_メーカー週報様式.xlsx", "作業日報", c.fy, "メーカー週報", c.cid); g_daily.maker(c, p, fn)
            if c.status == "完了":
                p, fn = emit(P + f"工事報告書_{nm}_メーカー出来形写真帳.pdf", "工事報告書", c.fy, "メーカー出来形写真帳", c.cid); g_report.maker(c, p, fn)
            continue
        if c.fy <= 2016:
            p, fn = emit(P + f"工程表_{nm}_SPCsv1様式.xlsx", "工程表", c.fy, "SPC sv1", c.cid); g_sched.sv1(c, p, fn, "sv1")
        elif c.fy <= 2023:
            p, fn = emit(P + f"工程表_{nm}_SPCsv2様式.xlsx", "工程表", c.fy, "SPC sv2", c.cid); g_sched.sv1(c, p, fn, "sv2")
        else:
            p, fn = emit(P + f"工程表_{nm}_SPC2024様式.xlsx", "工程表", c.fy, "SPC 2024", c.cid); g_sched.s2024(c, p, fn)
        if c.fy <= 2013:
            p, fn = emit(P + f"作業日報_{nm}_x1の2009様式.xlsx", "作業日報", c.fy, "x1 2009", c.cid); g_daily.x1_v(c, p, fn, "2009")
        elif c.fy <= 2023:
            p, fn = emit(P + f"作業日報_{nm}_x1の2014様式.xlsx", "作業日報", c.fy, "x1 2014", c.cid); g_daily.x1_v(c, p, fn, "2014")
        else:
            p, fn = emit(P + f"作業日報_{nm}_x1の2024様式.xlsx", "作業日報", c.fy, "x1 2024", c.cid); g_daily.x1_2024(c, p, fn)
        if c.fy <= 2017:
            p, fn = emit(P + f"工事報告書_{nm}_SPCsv1文章型.pdf", "工事報告書", c.fy, "SPC sv1 文章型", c.cid); g_report.sv1(c, p, TMP, fn)
        else:
            p, fn = emit(P + f"工事報告書_{nm}_SPCsv2表型2018様式.pdf", "工事報告書", c.fy, "SPC sv2 2018", c.cid); g_report.sv2(c, p, fn, g_order.PO_OF)
    write_world()
    with open(os.path.join(HERE, "files_index.json"), "w", encoding="utf-8") as fp:
        json.dump(OUT, fp, ensure_ascii=False, indent=0)
    print(len(OUT), "files")

def write_world():
    eqs = []
    for e in W.EQS:
        eqs.append({"tag": e.tag, "name": e.name, "furnace": e.furn, "group": e.group, "system": e.sub, "class": e.cls, "crit": e.crit, "duty": e.duty,
                    "installed": e.install, "spcAssetNo": e.spc_no, "spec": e.spec, "maker": e.maker,
                    "aliases": {"SPC台帳": spc_name(e), "SPC計画": e.name, "組合計画・台帳": u_name(e)},
                    "replaced": W.replaced_before(e, 2027)})
    works = []
    for c in C.ALL:
        if c.kind == "kikan":
            for n_, (kou, name, tags, attr, ps, pe, as_, ae, crew) in enumerate(c.tasks):
                started = as_ <= W.TODAY
                done_days = sum(1 for d in W.daterange(as_, min(ae, W.TODAY)) if d.weekday() != 6) if started else 0
                for t in (tags or [None]):
                    works.append({"id": f"{c.cid}:K{n_:02d}" + (f":{t}" if t else ""), "contract": c.cid, "tag": t, "work": name, "attr": attr if attr != "-" else None,
                                  "start": as_.isoformat() if started else None, "end": ae.isoformat() if (started and ae <= W.TODAY) else None,
                                  "plannedStart": ps.isoformat(), "plannedEnd": pe.isoformat(), "manDays": crew * done_days if crew else 0,
                                  "contractor": W.MAKER, "status": "完了" if ae <= W.TODAY else ("施工中" if started else "未着手")})
            continue
        for t in c.tasks:
            md = sum(sum(cr.values()) for _, _, cr in t.steps)
            for tg in (t.tags or [None]):
                works.append({"id": f"{c.cid}:{t.tid}" + (f":{tg}" if tg else ""), "contract": c.cid, "tag": tg, "work": t.name, "attr": t.attr if t.attr != "-" else None,
                              "start": t.actual[0].isoformat(), "end": t.actual[1].isoformat(), "manDays": round(md / max(1, len(t.tags or [1])), 2),
                              "contractor": t.company, "note": t.note})
    undoc = []
    for c in C.ALL:
        for tag, attr in c.direct:
            undoc.append({"contract": c.cid, "tag": tag, "attr": attr, "period": [c.start.isoformat(), c.end.isoformat()],
                          "reason": "SPC 保全班の直営（弁の一部・C 重要度の小物・タンク等、または協力会社の範囲 12 件を超えた分）。協力会社の日報・報告書に載らず、SPC 保全システムの作業記録のみ（本サンプルには含めない）"})
    for y in range(2009, 2027):
        undoc.append({"contract": None, "tag": None, "attr": "INSPECT", "period": [f"{y}-04-01", f"{y + 1}-03-31"],
                      "reason": "法定点検（受変電 年次・月次、消防設備、DXN・ばい煙測定、クレーン性能検査ほか）は外部の検査機関・SPC が別途行い、計画の法定点検行にのみ載る"})
    for e in W.EQS:
        if e.furn == 1:
            c = C.raw_code(e, 2026)
            if c in ("O", "R"):
                undoc.append({"contract": None, "tag": e.tag, "attr": C.ATTR[c], "period": ["2026-04-01", "2027-03-31"],
                              "reason": "1号炉は R8 基幹改良で停止中のため、SPC 定期整備分は 2027 年度へ繰越（未実施）"})
    contracts = [{"id": c.cid, "fy": c.fy, "furnace": c.furn, "kind": c.kind, "title": c.title, "orderer": c.orderer, "contractor": c.contractor,
                  "start": c.start.isoformat(), "end": c.end.isoformat(), "status": c.status} for c in C.ALL]
    gaps = [
        {"file": "長期修繕計画 2015・2016 年度改訂（SPC sv1）", "kind": "周期の超え", "what": "精密機能検査 2014→2018（3 年周期を 4 年）"},
        {"file": "長期修繕計画 2009–2016（SPC sv1）", "kind": "法令点検の抜け", "what": "第一種圧力容器 性能検査（脱気器）の行が無い（sv2 2017 で追加）"},
        {"file": "長期修繕計画 2017–2019 年度改訂（SPC sv2）費用集計", "kind": "小計の食い違い", "what": "2019 年度合計が行の合計より 88,000 千円少ない（DCS 更新分を後から足し合計未修正）"},
        {"file": "中長期整備計画 令和6年度策定・令和7年度改訂（組合）", "kind": "法令点検の抜け", "what": "計量機 定期検査の行が無い"},
        {"file": "中長期整備計画 令和6年度策定・令和7年度改訂（組合）総括表", "kind": "小計の食い違い", "what": "燃焼溶融設備 R8 だけ税抜のまま転記（シート小計と合わない）"},
        {"file": "長期修繕計画・組合計画", "kind": "周期の超え", "what": "2号 不燃物抜出装置A 2 年周期を R7→R8 へ 1 年延長（基幹改良と重なり）"},
    ]
    prof = {"id": "f09", "name": W.FAC, "union": W.UNION, "operator": W.SPC, "maker": W.MAKER, "capacity": "95t/24h×2炉＝190t/日",
            "type": "流動床式ガス化溶融炉（ガス化炉＋旋回溶融炉）全連続", "completed": "2009-03", "period": [2009, 2026], "asOf": W.TODAY.isoformat(),
            "boiler": "4MPa 400℃ 13t/h", "power": "3,800kW", "idf": "22,000m3N/h", "equipmentCount": {"total": len(W.EQS), "perFurnace": 205, "common": 238},
            "counting": "タグ番号単位。弁は電動・調節弁のみ、計装は分析計・主要流量計のみ。SPC 台帳の付属機器・電動機行は数えない。2018 年度追設の水銀計 2 台は 2017 年度以前の台帳・計画に無い",
            "contracts": "第1期 2009–2023 長期包括（SPC、整備込み）／第2期 2024–2033 運営委託（年度上限内）／基幹改良 R7–R9 組合→メーカー直接（継続費）"}
    world = {"profile": prof, "equipment": eqs, "contracts": contracts, "works": works, "links": rec.LINKS, "undocumented": undoc, "planGaps": gaps}
    with open(os.path.join(ANS, "world.json"), "w", encoding="utf-8") as fp:
        json.dump(world, fp, ensure_ascii=False, indent=0)

if __name__ == "__main__":
    main()
