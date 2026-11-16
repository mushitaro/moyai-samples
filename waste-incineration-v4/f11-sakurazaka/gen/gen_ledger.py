# -*- coding: utf-8 -*-
"""機器台帳：l1 2019 作成（1 行に台数）→ l2 2025 改訂（列を足した）"""
import os, copy
from openpyxl import Workbook
from openpyxl.comments import Comment
from common import *
import world as W

FN_OLD = "【架空】さくら坂焼却場_機器台帳_令和元年度作成_l1様式.xlsx"
FN_NEW = "【架空】さくら坂焼却場_機器台帳_令和7年度改訂_l2様式.xlsx"

UNIT_TXT = {"1": "1号", "2": "2号", "共通": "共通"}
UNIT_TXT_H14 = {"1": "No.1", "2": "No.2", "共通": "共通"}

def maker_of(e):
    if "H21 更新" in e.renew and e.name == "ごみクレーン":
        return W.MAKER_S + "（常用機：港南クレーン製作所）"
    return W.MAKER_S

# --- l2 が 2025 に足した列の中身 ---------------------------------------
def supply(e):
    n = e.name
    if any(k in n for k in ("火格子",)) and "油圧" not in n:
        return "生産中止（R4.12 最終受注）"
    if n in ("中央監視装置",):
        return "生産中止（保守部品なし）"
    if n == "ごみクレーン":
        return "予備機：生産中止／常用機：供給可"
    if "スートブロワ" in n:
        return "生産中止（代替品あり）"
    if n in ("誘引通風機",):
        return "長納期（羽根車 10か月）"
    if n in ("給じん装置", "ホッパゲート", "ブリッジ解除装置"):
        return "長納期（6〜8か月）"
    if n in ("油圧ポンプ", "灰押出機", "押込送風機", "二次送風機", "飛灰混練機"):
        return "長納期（4〜6か月）"
    if n in ("廃熱ボイラ本体", "節炭器"):
        return "管材は市販（図面は市保管）"
    if any(k in n for k in ("ダクト", "煙道", "ピット", "槽", "タンク", "煙突", "本体", "ホッパ")) and "ポンプ" not in n:
        return "—"
    if any(k in n for k in ("ポンプ", "ファン", "ブロワ", "圧縮機", "コンベヤ", "変圧器", "ホイスト")):
        return "汎用品で代替可"
    return "供給可"

STATE_OVR = {"1B-03": "B", "2B-03": "B", "2F-07": "C", "1C-01": "B", "0A-07": "B", "0L-01": "C",
             "0E-01": "B", "1C-12": "B", "1G-01": "C", "2G-01": "B", "0G-08": "A", "1D-05": "A", "2D-05": "A",
             "0K-04": "B", "0H-06": "B", "2C-12": "B"}

def state(e):
    if e.no in STATE_OVR:
        return STATE_OVR[e.no]
    h = sum(ord(c) for c in e.no + e.name) % 7
    return "A" if h < 4 else "B"

def disposal(e):
    if e.name in ("トラックスケール",):
        return "転用検討（新施設 計量棟）"
    if e.name in ("ごみクレーン",):
        return "常用機は売却検討"
    if e.name in ("非常用発電機",):
        return "解体まで保安用に残す"
    if e.name in ("受電用変圧器",):
        return "解体工事まで使用"
    return "撤去（解体工事）"

# l2 が更新・改造の欄に追記した文（R1〜R6 の実績）
RENEW_ADD = {
    "1D-05": "R2 ろ布全数取替（540本）",
    "2D-05": "R3 ろ布全数取替（540本）",
    "1C-01": "R4 第1パス前面水管 8本部分取替",
    "0L-01": "R2 監視操作端末 3台更新",
    "1B-03": "R5 火格子片 72枚取替（予備 18枚 倉庫保管）",
    "0G-08": "R6 常用機パドル全数取替（地元製作）",
    "2F-07": "R6 羽根車購入（R7.3 納入・倉庫保管）→ R7.10 取替予定",
    "1C-12": "R4 下部コーン当て板",
    "0E-01": "R3 常用機 管3本施栓",
}


# ---------------------------------------------------------------------------
# l1 は完成図書の「機器仕様書」（1 機器 1 欄：型式／容量／電動機／付属品）の書き方のまま写した。
# 1 機器 = 見出し 1 行＋仕様 3〜4 行のブロック。H14 の追加機器は各設備の後ろに「H14 追加」の見出しで差し込む。
# ---------------------------------------------------------------------------
ROTATING = ("エアカーテン", "ポンプ", "ファン", "ブロワ", "送風機", "通風機", "圧縮機", "コンベヤ", "フィーダ", "供給機", "混練機", "ホイスト", "クレーン", "押出機", "磁選機", "脱水機", "スケール")
ACC = {
    "誘引通風機": "潤滑油ユニット（別項）、インバータ盤（H22）",
    "押込送風機": "吸込消音器、Vベルト駆動",
    "バグフィルタ本体": "パルス用ヘッダタンク、差圧計",
    "廃熱ボイラ本体": "安全弁 2、水面計 2、ドラム",
    "ごみクレーン": "グラブバケット（常用・予備 各 1）、給電ケーブルリール",
    "灰クレーン": "グラブバケット 1",
    "給じん装置": "油圧シリンダ 2",
    "燃焼火格子": "可動フレーム、支持ローラ 12",
    "乾燥火格子": "可動フレーム、支持ローラ 8",
    "後燃焼火格子": "可動フレーム、支持ローラ 6",
    "火格子油圧ユニット": "油圧ポンプ（別項）、油冷却器（別項）",
    "減温塔本体": "2 流体噴霧ノズル 4",
    "温水熱交換器（温水プール用）": "安全弁、温度調節弁",
    "非常用発電機": "始動用空気槽、燃料小出槽（別項）",
    "中央監視装置": "プリンタ、無停電電源（別項）",
    "飛灰混練機": "パドル 二軸",
    "トラックスケール": "指示計（H22）、印字装置",
}

def split_spec(e):
    toks = e.spec.split()
    kata = toks[0] if toks else ""
    k = 1
    import re
    while k < len(toks) and (toks[k] in ("製", "造") or (toks[k].endswith("式") and not re.search(r"\d", toks[k]))):
        kata += (" " if toks[k] not in ("製", "造") else "") + toks[k]
        k += 1
    cap, mot = [], []
    for t in toks[k:]:
        if "kW" in t and any(k in e.name for k in ROTATING) and "×" not in t:
            mot.append(t)
        elif "kW" in t and "×" in t and any(k in e.name for k in ROTATING):
            a, b = t.rsplit("×", 1)
            cap.append(a); mot.append(b) if "kW" in b else cap.append(b)
        else:
            cap.append(t)
    # 容量の中の電動機（「… 30kW」が最後に付く形）
    if cap and "kW" in cap[-1] and any(k in e.name for k in ROTATING):
        mot.append(cap.pop())
    return kata, " ".join(cap), " ".join(mot)

def unit_txt(e):
    return {"1": "1号炉", "2": "2号炉", "共通": "共通"}[e.unit]

def block(ws, r, e, new):
    """1 機器のブロックを r 行目から書き、次の行を返す"""
    kata, cap, mot = split_spec(e)
    h14 = e.oldno.startswith("D14")
    # 見出し行
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    t = "【%s】　%s　（%s）" % (e.no, e.name, unit_txt(e))
    cell(ws, "A%d" % r, t, F(10, True), LEFT, B_ALL, fill("DDEBF7") if not h14 else fill("E2EFDA"))
    for c in "BCDEF":
        ws["%s%d" % (c, r)].border = B_ALL
    cell(ws, "G%d" % r, "図書 " + e.oldno, F(8), CENTER, B_ALL, fill("DDEBF7") if not h14 else fill("E2EFDA"))
    left = [("型　式", kata), ("容量・寸法", cap), ("電 動 機", mot or "—")]
    right = [("数　量", "%d 台" % e.units + ("（%s）" % e.remark if e.remark else "")),
             ("製造・設置", "%s　%s" % (maker_of(e), e.inst)),
             ("更新・改造", e.renew or "")]
    acc = ACC.get(e.name)
    if acc:
        left.append(("付 属 品", acc))
        right.append(("", ""))
    n = len(left)
    for i in range(n):
        rr = r + 1 + i
        lab, val = left[i]
        cell(ws, "A%d" % rr, lab, F(8), CENTER, B_HAIR, fill("F2F2F2"))
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
        cell(ws, "B%d" % rr, val, F(9), LEFT, B_HAIR); ws["C%d" % rr].border = B_HAIR
        lab, val = right[i]
        cell(ws, "D%d" % rr, lab, F(8), CENTER, B_HAIR, fill("F2F2F2") if lab else None)
        ws.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=7)
        cell(ws, "E%d" % rr, val, F(9), LEFT, B_HAIR)
        for c in "FG":
            ws["%s%d" % (c, rr)].border = B_HAIR
        if new and lab == "更新・改造" and RENEW_ADD.get(e.no):
            c = ws["E%d" % rr]
            c.value = (e.renew + "／" if e.renew else "") + RENEW_ADD[e.no]
            c.font = F(9, color="0000CC")
    if new:
        add = [("部品供給", supply(e)), ("状態 R6", state(e)), ("廃止時", disposal(e)), ("", "")]
        cell(ws, "H%d" % r, "R7.6 追記（l2）", F(8, True, "0000CC"), CENTER, B_ALL, fill("FFF2CC"))
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        ws["I%d" % r].border = B_ALL
        for i in range(n):
            rr = r + 1 + i
            lab, val = add[i]
            cell(ws, "H%d" % rr, lab, F(8), CENTER, B_HAIR, fill("FFF2CC") if lab else None)
            cell(ws, "I%d" % rr, val, F(9, color="C00000" if (lab == "状態 R6" and val == "C") or "生産中止" in val else "0000CC"), LEFT, B_HAIR)
    return r + 1 + n + 0


def build(new=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "集計"
    cell(ws, "A1", "【架空】さくら坂焼却場　機器台帳（機器仕様台帳）", F(16, True))
    cell(ws, "A2", "令和元年10月作成　施設管理課（l1）", F(10))
    if new:
        cell(ws, "A3", "令和7年6月改訂（l2）：各機器の右に部品供給・状態・廃止時の処置の欄を追加、R1〜R6 の更新を追記（青字）", F(10, color="0000CC"))
    notes = [
        "・竣工時（S63.3）の完成図書「機器仕様書」の書き方に合わせ、1 機器ごとに型式・容量・電動機・付属品を写した（Excel の元データは H30 以前分が残っていない）。",
        "・同じ型式の並列機は 1 欄にまとめ、数量欄に台数と常用・予備の別を書いた。",
        "・H14 ダイオキシン類対策工事の追加機器は、各設備の後ろに「H14 追加」として差し込んだ（図書番号 D14-）。",
        "・弁類、計装機器、盤類は本台帳に含めない（完成図書 計装機器リスト・単線結線図を参照）。電気集じん器（H14 撤去）は欠番。",
    ]
    for i, t in enumerate(notes):
        cell(ws, "A%d" % (5 + i), t, F(9))
    r0 = 10
    heads = ["設備", "1号炉\n欄数", "1号炉\n台数", "2号炉\n欄数", "2号炉\n台数", "共通\n欄数", "共通\n台数", "計\n欄数", "計\n台数"]
    for i, h in enumerate(heads):
        cell(ws, "%s%d" % (get_column_letter(1 + i), r0), h, F(9, True), CENTER, B_ALL, fill("D9E1F2"))
    tot = [0] * 8
    r = r0 + 1
    for cat in W.CAT_ORDER:
        vals = []
        for u in ("1", "2", "共通"):
            es = [e for e in W.EQUIP if e.cat == cat and e.unit == u]
            vals += [len(es), sum(e.units for e in es)]
        vals += [vals[0] + vals[2] + vals[4], vals[1] + vals[3] + vals[5]]
        tot = [a + b for a, b in zip(tot, vals)]
        cell(ws, "A%d" % r, W.CAT[cat], F(9), LEFT, B_ALL)
        for i, v in enumerate(vals):
            cell(ws, "%s%d" % (get_column_letter(2 + i), r), v if v else "", F(9), CENTER, B_ALL)
        r += 1
    cell(ws, "A%d" % r, "合計", F(9, True), CENTER, B_ALL, fill("F2F2F2"))
    for i, v in enumerate(tot):
        cell(ws, "%s%d" % (get_column_letter(2 + i), r), v, F(9, True), CENTER, B_ALL, fill("F2F2F2"))
    setw(ws, {"A": 18, **{get_column_letter(i): 8 for i in range(2, 10)}})
    page(ws, "portrait", 9)

    ws = wb.create_sheet("機器仕様台帳")
    cell(ws, "A1", "【架空】さくら坂焼却場　機器仕様台帳", F(13, True))
    cell(ws, "A2", "作成 R1.10 施設管理課（竣工時 完成図書「機器仕様書」・H14 工事 完成図書より）" + ("　／　改訂 R7.6（l2）" if new else ""), F(9))
    setw(ws, {"A": 10, "B": 30, "C": 14, "D": 10, "E": 22, "F": 14, "G": 11})
    if new:
        setw(ws, {"H": 9, "I": 30})
    r = 4
    for ci, cat in enumerate(W.CAT_ORDER, 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9 if new else 7)
        cell(ws, "A%d" % r, "%d.　%s" % (ci, W.CAT[cat]), F(12, True), LEFT, Border(bottom=med))
        r += 1
        if cat == "D":
            cell(ws, "A%d" % r, "（竣工時の電気集じん器・EP ダストコンベヤ 各炉 図書 D-191・D-192 は H14 撤去により欠番）", F(8, color="808080"))
            r += 1
        for grp_h14 in (False, True):
            es = [e for e in W.EQUIP if e.cat == cat and e.oldno.startswith("D14") == grp_h14]
            if not es:
                continue
            if grp_h14:
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
                cell(ws, "A%d" % r, "―― H14 追加（ダイオキシン類対策工事 完成図書より）――", F(9, True, "375623"), CENTER)
                r += 1
            for u in ("1", "2", "共通"):
                for e in [x for x in es if x.unit == u]:
                    r = block(ws, r, e, new)
                    r += 0
        r += 1
    page(ws, "portrait", 9 if not new else 8, 1, 0)
    ws.print_title_rows = None
    if new:
        for row in ws.iter_rows(min_row=4, max_col=1):
            v = row[0].value
            if isinstance(v, str) and v.startswith("【2F-07】"):
                row[0].comment = Comment("R6 羽根車購入（負担行為 R06-0108）。R7.3.14 納入、倉庫 B-2。取替は R7 2号定修で。l2", "l2")
            if isinstance(v, str) and v.startswith("【0L-01】"):
                row[0].comment = Comment("端末は R2 更新だがサーバは H22 のまま。OS サポート切れ。廃止まで持たせる。l2", "l2")
            if isinstance(v, str) and v.startswith("【1B-03】"):
                row[0].comment = Comment("朝凪炉工 R4.9 通知：火格子片 R4.12 受注分で生産終了", "l2")
        wsh = wb.create_sheet("変更履歴")
        cell(wsh, "A1", "【架空】さくら坂焼却場　機器台帳　変更履歴", F(12, True))
        hist = [
            ("R1.10", "l1", "新規作成（完成図書の機器仕様書の書き方で転記、198 欄・273 台）"),
            ("R3.4", "l1", "1号 BF ろ布取替（R2）は紙の台帳にだけ手書きで記入（Excel は未更新、R7 改訂で追記）"),
            ("R7.6", "l2", "各機器の右に追記欄：部品供給（R5.11 朝凪炉工の供給可否回答）、状態（R6 定修結果）、廃止時の処置（案）"),
            ("R7.6", "l2", "更新・改造欄に R1〜R6 の実績を青字で追記（BF ろ布、1号ボイラ水管、中央監視端末、1号火格子、混練機パドル ほか）"),
            ("R7.6", "l2", "2号誘引通風機 羽根車：R6 購入品が R7.3 に納入。取替は R7 定修予定（取替後に更新欄を直す）"),
        ]
        for i, h in enumerate(["日付", "担当", "内容"]):
            cell(wsh, "%s3" % "ABC"[i], h, F(9, True), CENTER, B_ALL, fill("FFF2CC"))
        for i, (d, p, t) in enumerate(hist):
            for j, v in enumerate((d, p, t)):
                cell(wsh, "%s%d" % ("ABC"[j], 4 + i), v, F(9), LEFT, B_ALL)
        setw(wsh, {"A": 8, "B": 6, "C": 90})
        page(wsh, "landscape", 9)
    props(wb, "【架空】さくら坂焼却場 機器台帳（機器仕様台帳）", "【架空】さくら坂市 施設管理課 " + ("l2" if new else "l1"))
    fn = os.path.join(OUT, FN_NEW if new else FN_OLD)
    wb.save(fn)
    return fn


def main():
    return [build(False), build(True)]

if __name__ == "__main__":
    main()
