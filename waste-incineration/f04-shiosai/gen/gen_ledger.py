# -*- coding: utf-8 -*-
"""機器台帳：2015 t1（竣工図書の機器表を Excel に写したもの、1 シート、同型は数量でまとめる）
          2024 t3（1 台 1 行、管理番号、炉別・共通・建築のシート）"""
import datetime as dt
from collections import OrderedDict
import world as W
from xl import *

MAKERS = {"crane": "【架空】大鷲クレーン製作所", "scale": "【架空】さざなみ衡機", "pc": "（市販品）",
          "hv": "【架空】東陽電機", "tr": "【架空】東陽電機", "panel": "【架空】東陽電機", "gen": "【架空】八雲発動機",
          "dc": "【架空】東陽電機", "ups": "【架空】東陽電機", "pump": "【架空】ナミキポンプ", "ac": "（建築工事）",
          "fire": "（建築工事）", "elev": "【架空】昇和エレベータ", "hoist": "【架空】大鷲クレーン製作所", "itv": "【架空】東陽電機"}

def maker_of(e, style):
    if e.sys == "Z" and e.cls not in ("elev",): return "（建築工事）"
    mk = MAKERS.get(e.cls)
    if e.cls == "panel" and e.sys in ("B", "E", "P"): mk = None
    if e.cls == "pump" and e.sys in ("C",): mk = None
    if not mk:
        mk = "水樹プラント（株）" if style == 2015 else "水樹プラント"
    return mk

def ym(fy):
    if fy == 1994: return "H7.3"
    if fy == 2001: return "H14.3"
    return "H%d" % (fy - 1988) if fy <= 2018 else "R%d" % (fy - 2018)

def ledger_2015():
    fname = "【架空】しおさい町ごみ処理場_機器台帳_H27作成_t1竣工図書転記様式.xlsx"
    rec_open(fname)
    wb = new_wb("【架空】しおさい町ごみ処理場 機器台帳（H27）", "しおさい町 環境衛生課 t1", dt.datetime(2015, 5, 22, 9, 0), "t1")
    ws = wb.active; ws.title = "機器一覧"
    widths(ws, [9, 12, 26, 8, 6, 34, 18, 8, 26])
    put(ws, 1, 1, "【架空】しおさい町ごみ処理場　機器台帳", F(14, True))
    put(ws, 2, 1, "（竣工図書 機器表〔H7.3〕・DXN対策工事 完成図書〔H14.3〕より転記　H27.5 作成）", F(9))
    put(ws, 2, 8, "作成：t1", F(9))
    hdr = ["図書No", "設備名", "機器名称", "号炉", "数量", "形式・仕様", "製作所", "設置", "備考"]
    for i, h in enumerate(hdr, 1):
        put(ws, 4, i, h, F(10, True), AC, B_ALL, fill("D9D9D9"))
    groups = OrderedDict()
    late = []
    for e in W.EQ:
        if e.added_2015:
            late.append(e); continue
        key = (e.sys, e.dno, e.oname, e.spec)
        groups.setdefault(key, []).append(e)
    r = 5
    sys_start = {}
    prev_maker = None
    rows_by_sys = OrderedDict()
    for key, es in groups.items():
        rows_by_sys.setdefault(key[0], []).append(es)
    for sysc in W.SYS_ORDER:
        if sysc not in rows_by_sys: continue
        r0 = r
        for es in rows_by_sys[sysc]:
            e = es[0]
            fs = sorted(set(x.furnace for x in es))
            if fs == [0]: gk = "共通"
            elif fs == [1, 2]: gk = "1・2号"
            else: gk = "%d号" % fs[0]
            n = len(es)
            per = n // len(fs) if fs != [0] else n
            put(ws, r, 1, e.dno, F(9), AC, B_HAIR)
            put(ws, r, 3, e.oname, F(10), AL, B_HAIR)
            put(ws, r, 4, gk, F(9), AC, B_HAIR)
            q = put(ws, r, 5, n, F(10), AC, B_HAIR)
            put(ws, r, 6, e.spec, F(9), AL, B_HAIR)
            mk = maker_of(e, 2015)
            put(ws, r, 7, "〃" if mk == prev_maker else mk, F(9), AC if mk == prev_maker else AL, B_HAIR)
            prev_maker = mk
            put(ws, r, 8, ym(e.inst), F(9), AC, B_HAIR)
            note = []
            if fs == [1, 2] and per > 1: note.append("各炉%d台（常用1・予備%d）" % (per, per - 1))
            elif fs == [0] and n > 1 and e.name[-1] in "AB": note.append("常用1・予備1")
            if e.name in W.EN_SCOPE: note.append("H27・28 長寿命化で更新予定")
            if e.name == "灰出しコンベヤNo.1": note.append("H21 チェーン取替")
            if e.name == "ごみクレーン": note.append("予備機なし")
            if e.name == "誘引通風機": note.append("H27・28 更新予定（インバータ化）")
            if e.name == "バグフィルタ": note.append("H14 EP撤去→BF新設。ろ布 H20取替")
            if e.name == "空気圧縮機No.1": note.append("H22 更新")
            if e.oname == "空気槽（雑用）" or "第二種" in e.spec: note.append("2圧 自主検査 年1")
            put(ws, r, 9, "、".join(note), F(9), AL, B_HAIR)
            rec(ws.title, r, "item", {"図書No": e.dno, "設備名": W.SYS[sysc], "機器名称": e.oname, "号炉": gk, "数量": n, "形式・仕様": e.spec,
                                      "製作所": "〃" if ws.cell(row=r, column=7).value == "〃" else mk, "設置": ym(e.inst), "備考": "、".join(note)})
            eqlink(ws.title, r, [x.mid for x in es])
            r += 1
        merge(ws, r0, 2, r - 1, 2, W.SYS[sysc], F(10, True), AC)
        # 設備の区切り（下線を太く）
        for c in range(1, 10):
            ws.cell(row=r - 1, column=c).border = Border(left=thin, right=thin, top=hair, bottom=thin)
    # H28 に t1 が追記（インバータ盤）。並びの途中に入れず末尾に足した
    r += 1
    put(ws, r, 1, "【H28 追加】", F(10, True, "C00000"))
    r += 1
    put(ws, r, 1, "E-405", F(9, color="C00000"), AC, B_HAIR)
    put(ws, r, 2, "通風設備", F(9, color="C00000"), AC, B_HAIR)
    put(ws, r, 3, "誘引送風機インバータ盤", F(10, color="C00000"), AL, B_HAIR)
    put(ws, r, 4, "1・2号", F(9, color="C00000"), AC, B_HAIR)
    put(ws, r, 5, 2, F(10, color="C00000"), AC, B_HAIR)
    put(ws, r, 6, "37kW用 屋内自立", F(9, color="C00000"), AL, B_HAIR)
    put(ws, r, 7, "水樹プラント（株）", F(9, color="C00000"), AL, B_HAIR)
    put(ws, r, 8, "H27・H28", F(9, color="C00000"), AC, B_HAIR)
    put(ws, r, 9, "長寿命化工事で新設（1号H27、2号H28）", F(9, color="C00000"), AL, B_HAIR)
    comment(ws, r, 3, "完成図書が届いたら図書Noを確認すること（仮にE-405）", "t1")
    rec(ws.title, r, "item", {"図書No": "E-405", "設備名": "通風設備", "機器名称": "誘引送風機インバータ盤", "号炉": "1・2号", "数量": 2, "形式・仕様": "37kW用 屋内自立",
                              "製作所": "水樹プラント（株）", "設置": "H27・H28", "備考": "長寿命化工事で新設（1号H27、2号H28）",
                              "comments": {"C%d" % r: "完成図書が届いたら図書Noを確認すること（仮にE-405）"}})
    eqlink(ws.title, r, [e.mid for e in W.EQ if e.added_2015])
    r += 2
    total = len(W.EQ)
    put(ws, r, 4, "計", F(10, True), AC, B_ALL)
    put(ws, r, 5, total, F(10, True), AC, B_ALL)
    rec(ws.title, r, "total", {"数量": total})
    put(ws, r, 6, "（H28追加 2台を含む）", F(9))
    r += 1
    put(ws, r, 1, "※弁類・計装品（検出端・指示計）は除く。建築設備は竣工図書 建築機械設備機器表より。", F(8))
    ws.freeze_panes = "A5"
    page(ws, landscape=False, paper=9, fit_w=1, fit_h=0, titles="4:4", area="A1:I%d" % r)
    ws.oddFooter.center.text = "&P / &N"
    return save(wb, fname)

def last_work(e, upto=2023):
    ws_ = [w for w in W.works_of(e.mid) if w.fy <= upto and w.kind != "INSPECT"]
    if not ws_: return ""
    w = ws_[-1]
    return "%s %s" % (W.wareki(w.fy), w.desc)

def role(e):
    nm = e.name
    if nm.endswith("A") and nm[:-1] + "B" in [x.name for x in W.EQ]: return "常用"
    if nm.endswith("B") and nm[:-1] + "A" in [x.name for x in W.EQ]: return "予備"
    if nm in ("井水揚水ポンプNo.1", "空気圧縮機No.1"): return "常用"
    if nm in ("井水揚水ポンプNo.2", "空気圧縮機No.2"): return "予備"
    if nm in ("中央制御室空調機No.1", "中央制御室空調機No.2"): return "交互"
    if nm == "ごみクレーン": return "単独（予備なし）"
    return ""

def ledger_2024():
    fname = "【架空】しおさい町ごみ処理場_機器台帳_R6改訂_t3一台一行様式.xlsx"
    rec_open(fname)
    wb = new_wb("【架空】しおさい町ごみ処理場 機器台帳（R6改訂）", "t1", dt.datetime(2015, 5, 22, 9, 0), "しおさい町 環境衛生課 t3")
    first = True
    sheets = [("1号炉", lambda e: e.furnace == 1), ("2号炉", lambda e: e.furnace == 2),
              ("共通", lambda e: e.furnace == 0 and e.sys != "Z"), ("建築設備", lambda e: e.sys == "Z")]
    for sname, cond in sheets:
        ws = wb.active if first else wb.create_sheet(sname)
        ws.title = sname; first = False
        widths(ws, [10, 8, 13, 24, 20, 9, 30, 7, 30, 6, 22])
        merge(ws, 1, 1, 1, 8, "【架空】しおさい町ごみ処理場　機器台帳（%s）" % sname, F(13, True), AL, None)
        put(ws, 1, 9, "R6.4 改訂（t3）／H27.5 初版（t1）", F(8), AR)
        put(ws, 2, 1, "管理番号＝号炉(0:共通)-設備記号-連番。旧図書Noは竣工図書の機器表番号。弁・計装品は除く。", F(8))
        hdr = ["管理番号", "旧図書No", "設備区分", "機器名称", "竣工図書名称", "常用/予備", "形式・仕様", "設置", "直近の整備・更新（R5まで）", "重要度", "備考"]
        for i, h in enumerate(hdr, 1):
            put(ws, 4, i, h, F(9, True), AC, B_ALL, fill("DDEBF7"))
        r = 5
        es = [e for e in W.EQ if cond(e)]
        prev_sys = None
        for e in es:
            if e.sys != prev_sys and prev_sys is not None:
                pass
            put(ws, r, 1, e.mid, F(9), AC, B_HAIR)
            put(ws, r, 2, e.dno, F(8, color="808080"), AC, B_HAIR)
            put(ws, r, 3, W.SYS[e.sys] if e.sys != prev_sys else "", F(9), AL, B_HAIR)
            put(ws, r, 4, e.name, F(10), AL, B_HAIR)
            put(ws, r, 5, "" if e.oname == e.name else e.oname, F(8, color="808080"), AL, B_HAIR)
            put(ws, r, 6, role(e), F(8), AC, B_HAIR)
            put(ws, r, 7, e.spec, F(8), AL, B_HAIR)
            put(ws, r, 8, ym(e.inst), F(9), AC, B_HAIR)
            put(ws, r, 9, last_work(e), F(8), AL, B_HAIR)
            put(ws, r, 10, e.imp, F(10, True), AC, B_HAIR, fill("FCE4D6") if e.imp == "A" else None)
            nt = ""
            if e.name == "ごみクレーン": nt = "予備なし→停止＝全炉停止。性能検査 奇数年度"
            if e.name in ("ごみクレーン操作盤", "ごみクレーン動力盤"): nt = "部品供給終了（R3通知）R6更新予定"
            if e.name == "計量機": nt = "定期検査 偶数年度"
            if "第二種" in e.spec: nt = "第二種圧力容器 自主検査 年1"
            if e.name == "誘引通風機用インバータ盤": nt = "H27/H28 新設（旧台帳では末尾に追記）"
            if e.name == "灰出しコンベヤNo.1": nt = "R4 チェーン全数取替"
            put(ws, r, 11, nt, F(8), AL, B_HAIR)
            rec(sname, r, "item", {"管理番号": e.mid, "旧図書No": e.dno, "設備区分": W.SYS[e.sys] if e.sys != prev_sys else "", "機器名称": e.name, "竣工図書名称": "" if e.oname == e.name else e.oname,
                                   "常用/予備": role(e), "形式・仕様": e.spec, "設置": ym(e.inst), "直近の整備・更新": last_work(e), "重要度": e.imp, "備考": nt})
            eqlink(sname, r, [e.mid])
            prev_sys = e.sys
            r += 1
        box(ws, 4, 1, r - 1, 11, B_HAIR)
        for c in range(1, 12):
            ws.cell(row=4, column=c).border = B_ALL
        put(ws, r + 1, 1, "計 %d 台" % len(es), F(10, True))
        rec(sname, r + 1, "total", {"台数": len(es)})
        ws.freeze_panes = "E5"
        page(ws, landscape=True, paper=9, fit_w=1, fit_h=0, titles="4:4", area="A1:K%d" % (r + 1))
        ws.oddFooter.center.text = "&A　&P / &N"
    ws = wb.create_sheet("改訂履歴")
    widths(ws, [10, 10, 70])
    put(ws, 1, 1, "【架空】しおさい町ごみ処理場　機器台帳　改訂履歴", F(12, True))
    rows = [("H27.5", "t1", "竣工図書機器表・DXN対策完成図書から転記（1シート、同型機は数量でまとめた）"),
            ("H28.3", "t1", "誘引送風機インバータ盤 2面を末尾に追加"),
            ("R6.4", "t3", "1台1行に分割。管理番号を付番。炉別・共通・建築でシート分割。名称を長寿命化計画の呼び方にそろえ、旧名称は残した。重要度（A/B/C）を整備計画から転記。直近整備は修繕執行管理表・点検報告書から拾った（R2・R3は管理表のみ）"),
            ]
    for i, h in enumerate(["改訂", "担当", "内容"], 1):
        put(ws, 3, i, h, F(10, True), AC, B_ALL, fill("DDEBF7"))
    for j, row in enumerate(rows):
        for i, v in enumerate(row, 1):
            put(ws, 4 + j, i, v, F(10), AL if i == 3 else AC, B_ALL)
        rec("改訂履歴", 4 + j, "item", {"改訂": row[0], "担当": row[1], "内容": row[2]})
    put(ws, 8, 1, "台数：1号炉 58／2号炉 58／共通 118／建築設備 22　合計 256 台", F(10))
    page(ws, landscape=False)
    return save(wb, fname)

if __name__ == "__main__":
    print(ledger_2015()); print(ledger_2024())
