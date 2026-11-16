# -*- coding: utf-8 -*-
"""工事報告書（全契約、報告書を提出済みのもの）
 メーカー様式 MP-R-03（2015–2020）：水樹テクノサービスが Word で作成 → LibreOffice で PDF
 SPC 様式 KZ-H-01（2021–）：sp2 が作った帳票。1 枚目サマリー＋続紙。reportlab で PDF
見本の 3 件（2018 1号炉・2023 1号炉・2023 全停）は見本の記載内容をそのまま使う。"""
import os
import copy
import zlib
from datetime import timedelta
from reportlab.lib.units import mm
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, PageBreak
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from common import *
from world import BY_CODE, rule
from contracts_full import get, CREWS
from works import TASK_WORKS, BY_ID
from supplies import supplies
import measures as M
import sample_report_new as SN
import sample_report_old as SOLD

LINKS = []
KUBUN = {"I": "点検", "L": "法定", "O": "整備", "R": "補修", "X": "更新"}
MAKER_NAME = {"燃焼火格子": "ストーカ（燃焼段）", "乾燥火格子": "ストーカ（乾燥段）", "後燃焼火格子": "ストーカ（後燃焼段）", "給じん装置": "給じんプッシャ",
              "主灰押出装置": "灰押出機", "焼却炉本体": "炉本体 耐火物", "誘引通風機": "IDF（誘引送風機）", "押込送風機": "FDF（押込送風機）",
              "二次送風機": "二次空気送風機", "排ガス再循環送風機": "EGR送風機", "バグフィルタ本体": "バグフィルタ", "触媒反応塔": "脱硝反応塔",
              "減温塔噴霧ノズルユニット": "減温塔 噴霧ノズル", "ボイラ給水ポンプ": "ボイラ給水ポンプ", "長抜差型スートブロワ": "ボイラ 長抜差スートブロワ",
              "蒸気タービン": "蒸気タービン", "空冷式復水器": "ACC（空冷復水器）", "空冷復水器ファン": "ACCファン", "空冷復水器ファン減速機": "ACCファン減速機",
              "3次過熱器": "ボイラ 3SH", "2次過熱器": "ボイラ 2SH", "1次過熱器": "ボイラ 1SH", "エコノマイザ": "ボイラ ECO", "水管壁（炉頂・放射室）": "ボイラ 水冷壁"}


def maker_name(code):
    e = BY_CODE[code]
    u = e["unit"]
    base = MAKER_NAME.get(e["base"], e["base"])
    pre = f"No.{u} " if u in ("1", "2") else ""
    if e["kind"] == "電動機":
        return pre + MAKER_NAME.get(BY_CODE[e["motor_of"]]["base"], BY_CODE[e["motor_of"]]["base"]) + e["sfx"] + " 電動機"
    return pre + base + (e["sfx"] if e["sfx"] else "")


def result_of(c, code, t, w):
    e = BY_CODE[code]
    y = c["fy"]
    u = e["unit"]
    b = e["base"]
    if w == "火格子摩耗測定":
        v = M.grate_wear(u, b.replace("火格子", ""), y) or 0
        return f"最大 {v}mm／管理15mm " + ("要注意" if v >= 10 else "良")
    if w == "燃焼火格子更新":
        return "良（クリアランス2～5mm内）"
    if w == "肉厚測定":
        v, pv, nom, lim = M.boiler_min(u, b, y)
        if v is None:
            return "良"
        return f"最小 {v}mm／限界{lim} " + M.boiler_judge(b, v, lim)
    if t == "L":
        return "合格" if ("検査" in w and "自主" not in w) else "良"
    if b == "誘引通風機" or "送風機" in b or "ポンプ" in b or "ファン" in b:
        if t in "OR":
            bv, av = M.vib(code, y)
            return f"振動 {bv}→{av}mm/s 良" if av else f"振動 {bv}mm/s 要観察"
    if w == "ろ布抜取試験":
        lo, hi = M.bf_strength(u, y)
        return f"保持率 {lo}～{hi}%"
    if w == "触媒性能試験":
        a, b2 = M.cat_rate(u, y)
        return f"脱硝率 {a}%" + (f"／{b2}%" if b2 else "")
    if "絶縁診断" in w:
        pi, td = M.pi_index(code, y)
        return f"PI {pi}、tanδ {td}% 良"
    return "良"


def ev_suffix(t):
    ev = t.get("event")
    if not ev:
        return ""
    return {"追加": "（追加補修あり）", "手直し": "（手直し後 良）", "遅れ": "（着手 2日繰下げ）"}[ev[0]]


def summary_rows(c):
    rows = []
    for t in c["tasks"]:
        for (code, tt, w) in t["works"]:
            e = BY_CODE[code]
            res = result_of(c, code, tt, w)
            rows.append(dict(code=code, name=e["name"], kubun=KUBUN[tt], work=w, result=res + (ev_suffix(t) if t["works"][0][0] == code else ""),
                             pm="済", task=t))
    return rows


def recommendations(c):
    y = c["fy"]
    out = []
    keys = set()
    for t in c["tasks"]:
        for (code, tt, w) in t["works"]:
            e = BY_CODE[code]
            if e["kind"] == "電動機" or code in keys:
                continue
            nxt = [(yy, ww) for yy, lst in sorted(rule(e, y >= 2020, y + 1).items()) if yy > y for (t2, ww) in lst if t2 in "XR"]
            if nxt and e["imp"] == "A" and len(out) < 4:
                yy, ww = nxt[0]
                txt = f"次回 {yy}年度 {ww}（長期修繕計画どおり）。"
                if w == "火格子摩耗測定":
                    v = M.grate_wear(e["unit"], e["base"].replace("火格子", ""), y)
                    txt = f"最大 {v}mm。" + txt
                if w == "肉厚測定":
                    v, pv, nom, lim = M.boiler_min(e["unit"], e["base"], y)
                    if v:
                        txt = f"最小 {v}mm（前回 {pv}mm）。" + txt
                out.append((e["name"], txt))
                keys.add(code)
    for t in c["tasks"]:
        ev = t.get("event")
        if ev and ev[0] == "追加":
            out.append(("耐火物", "今回の追加補修部（側壁下部）を次回定修で重点点検とする。"))
    if not out:
        out.append(("全般", "特記事項なし。次回定修は長期修繕計画どおり。"))
    return out


def materials(c):
    rows = []
    for s in supplies(c):
        q, used, sp = s["qty"], s["used"], s["spare"]
        if c["key"] == "C20230" and "減速機" in s["name"]:
            q, used, sp = 2, 1, 1
        rows.append([s["name"].replace(f"{c['unit']}号炉 ", ""), s["spec"], "SPC支給", f"{q}{s['unit']}", f"{used}{s['unit']}",
                     f"{sp}{s['unit']}" if sp else "—", "SPC予備品倉庫" if sp else ""])
    cats = {}
    for t in c["tasks"]:
        n = t["name"]
        for k, row in [("耐火物", ["キャスタブル（高アルミナ質）", "東雲築炉 標準品", "受注者（協力会社）", "1.5t", "1.4t", "0.1t", "持帰り"]),
                       ("シリンダ", ["油圧シリンダ パッキンキット", "純正", "受注者", "1式", "1式", "—", ""]),
                       ("油圧装置", ["作動油フィルタエレメント", "10μm", "受注者", "2個", "2個", "—", ""]),
                       ("送風機", ["送風機 軸受", "純正", "受注者", "1台分", "1台分", "—", ""]),
                       ("誘引通風機", ["誘引通風機 軸受", "22228 CCK／W33 ほか", "受注者", "1組", "1組", "—", ""]),
                       ("分解整備", ["メカシール・軸受（ポンプ類）", "純正", "受注者", "1式", "1式", "—", ""]),
                       ("スートブロワ", ["スートブロワ グランドパッキン", "黒鉛系", "受注者", "1式", "1式", "—", ""]),
                       ("排ガス処理", ["BF電磁弁ダイヤフラム・ダンパシール", "純正", "受注者", "1式", "1式", "—", ""]),
                       ("計装", ["熱電対・検出器", "純正", "受注者", "1式", "1式", "—", ""]),
                       ("電動機", ["電動機軸受", "6312ZZ ほか", "受注者", "1式", "1式", "—", ""]),
                       ("摺動部", ["摺動板（ライナ）", "SS400＋硬化肉盛 t25", "受注者", "4枚", "4枚", "—", ""])]:
            if k in n and k not in cats:
                cats[k] = row
    for k, row in cats.items():
        if k == "耐火物" and c.get("changes"):
            area = c["changes"][0][1]
            add = round(area * 0.5, 1)
            row = [row[0], row[1], row[2], f"当初1.5t＋追加{add}t", f"{round(1.4 + add, 1)}t", "0.1t", "持帰り"]
        rows.append(row)
    return rows


# ================================================================ SPC 様式
def spc_content(c):
    y = c["fy"]
    S, E = c["period"]
    no = f"KZ-{y}-H-{ {'2': 6, '1': 14, '0': 15, 'T': 18}[c['unit']]:03d}"
    sub = E + timedelta(days=20 + zlib.crc32(c["key"].encode()) % 6)
    stop = c["stop"]
    head = [("工事件名", "【架空】" + c["name"]),
            ("契約区分", f"定期整備（請負）　SPC発注 契約番号 {c['contract_no']}" + ("（変更1回）" if c.get("changes") else "")),
            ("工期（契約）", f"{S.year}年{S.month}月{S.day}日 ～ {E.year}年{E.month}月{E.day}日"),
            ("工期（実績）", f"{stop.year}年{stop.month}月{stop.day}日 炉立下げ ～ {E.year}年{E.month}月{E.day}日 完了確認"),
            ("受注者", f"【架空】水樹テクノサービス株式会社　職長 {c['foreman']}" + ("（日報 q3）" if c["unit"] == "0" and y >= 2021 else "")),
            ("SPC担当", f"{c['spc']}（保全）　提出日 {sub.year}年{sub.month}月{sub.day}日")]
    srows = [(r["code"], r["name"], r["kubun"], r["work"], r["result"], r["pm"]) for r in summary_rows(c)]
    changes = []
    for (what, area) in c.get("changes", []):
        ev = [t for t in c["tasks"] if t.get("event") and t["event"][0] == "追加"]
        d = ev[0]["event"][1]
        changes.append(("変更1", f"{d.month}/{d.day} {ev[0]['event'][2]}。変更契約により工期内で施工。"))
    for t in c["tasks"]:
        ev = t.get("event")
        if ev and ev[0] in ("手直し", "遅れ"):
            changes.append(({"手直し": "手直し", "遅れ": "工程"}[ev[0]], f"{ev[1].month}/{ev[1].day} {t['name'].split('（')[0]}：{ev[2]}。"))
    if not changes:
        changes = [("変更", "なし")]
    tables = []
    thick = [(code, w) for t in c["tasks"] for (code, tt, w) in t["works"] if w == "肉厚測定"]
    if thick:
        rows = []
        for code, w in thick:
            e = BY_CODE[code]
            v, pv, nom, lim = M.boiler_min(e["unit"], e["base"], y)
            if v is None:
                continue
            rows.append([code, e["name"], f"{nom}", f"{lim}", f"{v}", f"{pv}" if pv else "—", M.boiler_judge(e["base"], v, lim)])
        tables.append(("ボイラ肉厚測定（超音波厚さ計、【架空】かざはや非破壊検査、単位 mm）",
                       ["保全ｺｰﾄﾞ", "部位", "公称", "限界", "最小", "前回最小", "判定"], rows, [18, 48, 18, 16, 16, 22, 18]))
    grate = [(code, w) for t in c["tasks"] for (code, tt, w) in t["works"] if w == "火格子摩耗測定"]
    if grate:
        rows = []
        for code, w in grate:
            e = BY_CODE[code]
            v = M.grate_wear(e["unit"], e["base"].replace("火格子", ""), y) or 0
            rows.append([code, e["name"], f"{v}", f"{round(v * 0.72, 1)}", "15.0", "要注意" if v >= 10 else "良"])
        tables.append(("火格子摩耗測定（デプスゲージ、単位 mm、新品基準面から）", ["保全ｺｰﾄﾞ", "段", "最大", "平均", "管理値", "判定"], rows,
                       [20, 40, 20, 20, 20, 20]))
    rot = [(code, tt, w) for t in c["tasks"] for (code, tt, w) in t["works"]
           if tt in "OR" and ("送風機" in BY_CODE[code]["base"] or "ポンプ" in BY_CODE[code]["base"] or BY_CODE[code]["base"] == "誘引通風機")
           and BY_CODE[code]["kind"] != "電動機"]
    if rot:
        rows = []
        for code, tt, w in rot:
            bv, av = M.vib(code, y)
            rows.append([code, BY_CODE[code]["name"], "軸受振動（速度 rms）", f"{bv} mm/s", f"{av} mm/s" if av else "—", "4.5 以下", "良"])
        tables.append(("回転機の振動（整備前・整備後）", ["保全ｺｰﾄﾞ", "機器", "項目", "整備前", "整備後", "基準", "判定"], rows,
                       [18, 40, 34, 20, 26, 22, 14]))
    pics = [t["name"].split("（")[0][:14] for t in c["tasks"] if t["works"]][:6]
    while len(pics) < 3:
        pics.append("作業状況")
    return dict(no=no, head=head, summary=srows, changes=changes, recs=recommendations(c), tables=tables,
                materials=materials(c), photos=pics, title=f"【架空】{c['name']} 工事報告書",
                author=f"【架空】水樹テクノサービス {c['foreman']}（SPC様式 {c['spc']}）", footer=f"【架空】かざはやエコプラザ　{c['fy']}年度 {c['short']}",
                note=("※1号炉定期整備工事（別契約）の範囲は含まない" if c["unit"] == "0" else ""))


def c23_content(c):
    """見本（2023 1号炉）の記載。"""
    base = spc_content(c)
    base["no"] = "KZ-2023-H-014"
    base["head"] = [("工事件名", "【架空】" + c["name"]), ("契約区分", "定期整備（請負）　SPC発注 契約番号 KZ-C-2023-021（変更1回）"),
                    ("工期（契約）", "2023年10月2日 ～ 2023年11月2日"),
                    ("工期（実績）", "2023年10月1日 炉立下げ ～ 2023年11月1日 ごみ投入・定格到達、11月2日 完了確認"),
                    ("受注者", "【架空】水樹テクノサービス株式会社　職長 q2"), ("SPC担当", "sp2（保全）　提出日 2023年11月24日")]
    base["summary"] = list(SN.SUM_C23)
    base["changes"] = [("変更1", "10/11 燃焼段 左右側壁下部（火格子上0～400mm）に耐火物の剥離・欠損 計2.4m²（最大深さ85mm）を発見。"
                                "10/12 協議のうえ追加補修を承認（変更契約 2023/10/20、工期変更なし、耐火 2名増員）。"),
                       ("手直し", "10/19 燃焼火格子 取付後のクリアランス測定で 6・7・8列目が基準（2～5mm）を超過（最大7.5mm）。"
                                 "固定火格子受け金具の締付順による位置ずれ。10/19～20 再調整し全列 2.5～4.5mm。"),
                       ("工程", "火格子取付 2日、耐火補修 4日、水銀計 1日 予定より延びたが、足場解体以降は予定どおり。")]
    base["recs"] = [("乾燥火格子", "最大9.6mm。現在の摩耗速度（約1.1mm/年）では長期修繕計画の更新年 2027 年度で妥当。2025 年度に再測定。"),
                    ("3次過熱器", "下段管 最小5.6mm（2021：6.3mm）。減肉速度 約0.35mm/年。2027 年度の下段管部分更新を推奨（計画どおり）。"),
                    ("誘引通風機", "インペラ前縁の肉盛は今回で2回目。次回は肉盛でなくインペラ更新とする（計画 2029 年度）。"),
                    ("耐火物", "燃焼段側壁下部は今回初めて損傷。2024 年度定修で同部位を重点点検。")]
    import zlib as Z
    cl = []
    for col in range(1, 9):
        vals = [round(2.6 + (Z.crc32(f"cl{col}{k}".encode()) % 17) / 10, 1) for k in range(5)]
        note = "手直し前 最大" + {6: "7.5", 7: "6.8", 8: "6.1"}[col] if col in (6, 7, 8) else ""
        cl.append([f"{col}列"] + [str(v) for v in vals] + ["良", note])
    base["tables"] = [
        ("燃焼火格子 取付後クリアランス（可動／固定間、単位 mm、基準 2～5）", ["列", "左端", "1/4", "中央", "3/4", "右端", "判定", "備考"], cl,
         [12, 16, 16, 16, 16, 16, 14, 52]),
        ("ボイラ肉厚測定（超音波厚さ計、【架空】かざはや非破壊検査、単位 mm）", ["保全ｺｰﾄﾞ", "部位", "公称", "限界", "測定点数", "最小", "前回(2021)最小", "判定"],
         [["1BO-030", "水管壁 放射室下部（肉盛部）", "5.0+肉盛", "4.0", "120", "4.9", "5.2", "要観察"],
          ["1BO-030", "水管壁 炉頂", "5.0", "4.0", "60", "4.8", "4.9", "良"], ["1BO-040", "蒸発管群", "4.5", "3.5", "80", "4.3", "4.3", "良"],
          ["1BO-050", "1次過熱器", "5.0", "4.5", "60", "4.9", "4.9", "良"], ["1BO-060", "2次過熱器", "6.0", "4.5", "60", "5.8", "5.9", "良"],
          ["1BO-070", "3次過熱器 下段管", "7.1", "4.5", "72", "5.6", "6.3", "要観察"], ["1BO-070", "3次過熱器 上段管", "7.1", "4.5", "48", "6.8", "6.9", "良"],
          ["1BO-090", "エコノマイザ", "4.0", "3.0", "60", "3.9", "3.9", "良"], ["1BO-020", "蒸気ドラム 胴板", "32.0", "28.0", "12", "31.8", "31.8", "良"]],
         [17, 48, 18, 14, 18, 14, 26, 18]),
        ("回転機の振動・絶縁（据付後・立上げ時）", ["保全ｺｰﾄﾞ", "機器", "項目", "整備前", "整備後", "判定基準", "判定"],
         [["1KT-100", "誘引通風機", "軸受振動（速度 rms）", "5.8 mm/s", "2.1 mm/s（10/31 負荷運転）", "4.5 以下", "良"],
          ["1KT-100M", "誘引通風機 電動機", "成極指数 PI／tanδ", "—", "3.8／1.2%", "2.0以上／5%以下", "良"],
          ["1BO-100", "ボイラ給水ポンプA", "軸受振動", "3.9 mm/s", "1.6 mm/s", "4.5 以下", "良"],
          ["1NS-130", "燃焼装置油圧ポンプB", "吐出圧力", "13.2 MPa", "14.0 MPa", "14.0±0.5", "良"]], [17, 32, 36, 20, 40, 22, 14]),
    ]
    base["text_after"] = ("試験（委託）　ろ布：引張強度保持率 81～86%（交換目安 70%）、通気度 良。触媒：脱硝率 1層 92%・2層 89%（初期 95%）。"
                          "予備層充填は計画どおり 2025 年度でよい。（【架空】かざはや分析センター 2023/11/17 報告）")
    base["materials"] = [
        ["燃焼火格子（可動・固定）", "高クロム鋳鋼 MZ-G150-C", "SPC支給", "200枚", "192枚", "8枚", "SPC予備品倉庫（No.3棚）"],
        ["火格子取付ボルト・ナット", "SUS310S M20", "受注者", "420組", "384組", "36組", "受注者持帰り"],
        ["水銀連続分析計", "【架空】青嵐計器 SR-HG7", "SPC支給", "1台", "1台", "—", "旧機はSPCで廃棄"],
        ["誘引通風機 軸受", "22228 CCK／W33 ほか", "受注者", "1組", "1組", "—", "※SPC予備 1組は別途購入・保管"],
        ["インペラ肉盛溶接材", "硬化肉盛 HF-600 φ3.2", "受注者", "35kg", "31kg", "4kg", "受注者持帰り"],
        ["ボイラ給水ポンプ メカニカルシール", "【架空】千鳥ポンプ CP 純正", "SPC支給", "2組", "1組", "1組", "SPC予備品倉庫"],
        ["同 軸受・ライナリング", "純正", "SPC支給", "1式", "1式", "—", ""],
        ["キャスタブル（高アルミナ質）", "東雲築炉 標準品", "受注者（協力会社）", "当初1.5t＋追加1.2t", "2.6t", "0.1t", "持帰り"],
        ["BF電磁弁ダイヤフラム", "1-1/2\" 用", "受注者", "120個", "120個", "—", ""], ["R熱電対", "φ8 L=1,500 保護管付", "受注者", "3本", "3本", "—", ""],
        ["油圧ポンプ シールキット", "純正", "受注者", "2式", "2式", "—", ""],
        ["電動機軸受（油圧ポンプB×2・噴霧水ポンプB）", "6312ZZ ほか", "受注者", "3台分", "3台分", "—", ""]]
    base["photos"] = ["燃焼段 撤去後（受台）", "新品火格子 取付状況", "側壁下部 損傷部（追加補修前）", "同 補修後", "IDF インペラ 肉盛後", "水銀計 新機"]
    return base


def c23c_content(c):
    base = spc_content(c)
    base["no"] = "KZ-2023-H-015"
    base["head"] = [("工事件名", "【架空】" + c["name"]), ("契約区分", "定期整備（請負）　SPC発注 契約番号 KZ-C-2023-022"),
                    ("工期（契約）", "2023年10月14日 ～ 2023年10月22日"),
                    ("工期（実績）", "10月13日 21:00 2号炉立下げ（全炉停止）～ 10月22日 2号炉立上げ、10月23日 タービン併入"),
                    ("受注者", "【架空】水樹テクノサービス株式会社　職長 q2（日報 q3）"), ("SPC担当", "sp2（保全）　提出日 2023年11月10日")]
    base["summary"] = list(SN.SUM_C23C)
    base["changes"] = [("変更", "なし"), ("手直し", "10/20 減速機3 組立後の軸端振れ 0.08mm（基準0.05mm以下）。10/21 芯出し再調整し 0.03mm。工期内。")]
    base["recs"] = [("ACC減速機3", "低速段歯面に軽微なピッチング（面積率 約3%）。異音の主因は出力側軸受の外輪剥離で、軸受交換で解消。"
                                   "歯面は2025年度全停で再点検し、進行していれば歯車交換（予備歯車の手配 納期約8か月）。"),
                    ("ACC減速機（他5台）", "同型のため、2025・2027年度の分解整備を前倒しする場合は振動傾向で判断。")]
    base["tables"] = [("測定結果", ["保全ｺｰﾄﾞ", "機器", "項目", "結果", "基準", "判定"],
                       [["0YN-300", "ACCファン減速機3", "出力軸 振れ（再調整後）", "0.03 mm", "0.05 以下", "良"],
                        ["0YN-300", "同", "振動（速度 rms、10/23 運転）", "1.9 mm/s（整備前 6.4）", "4.5 以下", "良"],
                        ["0DE-020", "主変圧器", "油中ガス（可燃性ガス総量）", "185 ppm", "要注意 1,000", "良"],
                        ["0DE-270", "非常用発電機", "負荷試験 電圧・周波数", "6.6kV／50.0Hz 安定", "—", "良"],
                        ["0YN-400", "脱気器", "胴板 肉厚 最小", "11.7 mm（公称12）", "—", "合格"]], [18, 36, 46, 38, 26, 16])]
    base["materials"] = [["ACC減速機 軸受一式（1台分）", "【架空】水樹プラント 純正", "SPC支給", "2台分", "1台分", "1台分", "SPC予備品倉庫"],
                         ["減速機 潤滑油", "ISO VG220", "受注者", "60L", "48L", "12L", "持帰り"], ["主油ポンプ ブッシュ", "純正", "受注者", "2個", "2個", "—", ""],
                         ["ポンプ用メカシール・軸受", "純正（脱気器給水B・冷却水B）", "受注者", "2式", "2式", "—", ""],
                         ["電動機軸受（復水ポンプA）", "6310ZZ", "受注者", "1台分", "1台分", "—", ""],
                         ["脱気器 マンホールガスケット", "うず巻形", "受注者", "2枚", "2枚", "—", ""]]
    base["photos"] = ["減速機3 低速段歯面", "出力側軸受 外輪剥離", "脱気器 内部"]
    base["note"] = "※1号炉定期整備工事（別契約 KZ-C-2023-021）の範囲は含まない"
    return base


def code_links(c, code):
    out = []
    for t in c["tasks"]:
        for w in TASK_WORKS.get((c["key"], t["no"]), []):
            if BY_ID[w]["code"] == code or (code.startswith("0DE-0") and "～" in code and BY_ID[w]["code"].startswith("0DE")):
                out.append(w)
    return out


def build_spc(c, ct, fname):
    P, STS, STH, STT, ST = SN.P, SN.STS, SN.STH, SN.STT, SN.ST
    GRID, HEAD = SN.GRID, SN.HEAD
    meta = dict(title=ct["title"], author=ct["author"], subject="【架空】工事報告書 様式KZ-H-01", no=ct["no"], pages=1, footer=ct["footer"])
    seq = [0]
    hf = {k: v for k, v in ct["head"]}
    hf["報告書番号"] = ct["no"]

    def R(kind, fields):
        rec(fname, kind, fields, seq=seq[0])
        seq[0] += 1

    el = [P("【架空】工 事 報 告 書", STT), Spacer(1, 3 * mm)]
    top = Table([[SN.kv_table(ct["head"], 24, 110), SN.stamp()]], colWidths=[112 * mm, 72 * mm])
    top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    el.append(top)
    el.append(Spacer(1, 2 * mm))
    el.append(P(f"1. 対象機器と実施内容（サマリー）　※全{len(ct['summary'])}件。1枚目に載らない分は続紙へ続く" + ("　" + ct["note"] if ct.get("note") else ""), STH))
    el.append(SN.summary_table(ct["summary"], 20))
    for row in ct["summary"]:
        R("item", dict(section="サマリー", code=row[0], name=row[1], kubun=row[2], work=row[3], result=row[4], pmnavi=row[5],
                       contract=hf["工事件名"], reportNo=ct["no"]))
        LINKS.append(dict(file=fname, seq=seq[0] - 1, works=code_links(c, row[0])))
    el.append(Spacer(1, 2 * mm))
    el.append(P("2. 変更・追加", STH))
    el.append(SN.kv_table(ct["changes"], 18))
    for k, v in ct["changes"]:
        R("item", dict(section="変更・追加", label=k, text=v, reportNo=ct["no"]))
    el.append(Spacer(1, 2 * mm))
    el.append(P("3. 次回への提言", STH))
    el.append(SN.kv_table(ct["recs"], 22))
    for k, v in ct["recs"]:
        R("item", dict(section="次回への提言", label=k, text=v, reportNo=ct["no"]))
    el.append(PageBreak())
    el.append(P("4. 測定結果", STH))
    if not ct["tables"]:
        el.append(P("測定を伴う作業なし（目視・作動確認のみ）。", ST))
    for i, (title, head, rows, w) in enumerate(ct["tables"], 1):
        el.append(P(f"4.{i} {title}", ST))
        t = Table([[P(x, STS) for x in head]] + [[P(x, STS) for x in r] for r in rows], colWidths=[x * mm for x in w], repeatRows=1)
        t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (-1, 0), HEAD)]))
        el.append(t)
        el.append(Spacer(1, 3 * mm))
        for r in rows:
            R("item", dict(section="測定結果", table=title, **{h_: v for h_, v in zip(head, r)}, reportNo=ct["no"]))
    if ct.get("text_after"):
        el.append(P(ct["text_after"], ST))
        R("item", dict(section="測定結果", text=ct["text_after"], reportNo=ct["no"]))
    el.append(Spacer(1, 3 * mm))
    el.append(P("5. 使用部品・材料", STH))
    mh = ["品名", "仕様・型式", "手配", "数量", "使用", "予備・残", "保管・扱い"]
    t = Table([[P(x, STS) for x in mh]] + [[P(x, STS) for x in r] for r in ct["materials"]],
              colWidths=[x * mm for x in [40, 34, 20, 22, 14, 16, 36]], repeatRows=1)
    t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (-1, 0), HEAD)]))
    el.append(t)
    for r in ct["materials"]:
        R("item", dict(section="使用部品・材料", **dict(zip(mh, r)), reportNo=ct["no"]))
    el.append(P("※本様式には金額を記載しない（契約・精算は別途）。", STS))
    el.append(Spacer(1, 3 * mm))
    el.append(P("6. 写真（写真代替図）", STH))
    cap = ct["photos"]
    rows = [cap[i:i + 3] for i in range(0, len(cap), 3)]
    cells = [[P("【架空】見本用の写真代替図<br/><br/><br/><br/>" + f"写真{k * 3 + j + 1} {x}", STS) for j, x in enumerate(r)] +
             [""] * (3 - len(r)) for k, r in enumerate(rows)]
    t = Table(cells, colWidths=[60 * mm] * 3, rowHeights=[34 * mm] * len(cells))
    t.setStyle(TableStyle(GRID + [("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    el.append(t)
    path = outpath(fname)
    for _ in range(2):
        SN.Doc(path, meta).build(copy.deepcopy(el))
        meta["pages"] = pdf_pages(path)
    return fname


# ================================================================ メーカー様式（Word → PDF）
def maker_content(c):
    y = c["fy"]
    S, E = c["period"]
    works = []
    for t in c["tasks"]:
        if not t["works"]:
            continue
        codes = list(dict.fromkeys(code for code, _, _ in t["works"]))
        nm = "・".join(dict.fromkeys(maker_name(x) for x in codes[:3])) + ("ほか" if len(codes) > 3 else "")
        res = "、".join(dict.fromkeys(result_of(c, code, tt, w) for code, tt, w in t["works"]))[:60] + ev_suffix(t)
        works.append(dict(name=nm, work=t["name"], result=res, codes=codes, task=t))
    wear = None
    if any(w == "火格子摩耗測定" for t in c["tasks"] for (_, _, w) in t["works"]):
        wear = wear_table(c["unit"] if c["unit"] in ("1", "2") else "1", y)
    remarks = [f"({i + 1}) {a}：{b}" for i, (a, b) in enumerate(recommendations(c))]
    return dict(works=works, wear=wear, remarks=remarks, materials=materials(c))


def wear_table(u, y):
    rows = []
    for st, ncol in (("乾燥", 6), ("燃焼", 8), ("後燃焼", 6)):
        mx = M.grate_wear(u, st, y) or 0.5
        for side in ("可動", "固定"):
            for col in range(1, ncol + 1):
                vals = []
                for p in range(1, 6 if st == "燃焼" else 4):
                    k = zlib.crc32(f"{u}{y}{st}{side}{col}{p}".encode()) % 100 / 100
                    v = mx * (1.0 if side == "可動" else 0.55) * (0.72 + 0.25 * k)
                    vals.append(round(v, 1))
                if side == "可動" and col == (3 if st == "燃焼" else 2):
                    vals[1] = mx
                m_ = max(vals)
                rows.append([st + "段" if (side == "可動" and col == 1) else "", side if col == 1 else "", f"{col}列"] +
                            [str(v) for v in vals] + [""] * (5 - len(vals)) + [str(m_), "要注意" if m_ >= 10 else "良"])
    return rows


def build_maker(c, fname):
    S, E = c["period"]
    y = c["fy"]
    sample = c.get("sample") == "C18"
    stem = os.path.splitext(fname)[0]
    docx_path = os.path.join(WORK, stem + ".docx")
    doc = Document()
    cp = doc.core_properties
    cp.title = f"【架空】{c['name']} 完成報告書"
    cp.author = f"【架空】水樹テクノサービス株式会社 {c['foreman']}"
    cp.subject = "【架空】完成報告書（水樹プラント様式 MP-R-03）"
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Mm(297), Mm(210)
    sec.left_margin = sec.right_margin = Mm(20)
    sec.top_margin = Mm(18)
    sec.bottom_margin = Mm(15)
    SOLD.set_font(sec.header.paragraphs[0].add_run(f"【架空】水樹プラント様式 MP-R-03　　　　　　　　　　【架空】かざはやエコプラザ {c['short']}"), 7.5)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run()
    SOLD.set_font(run, 8)
    for kind, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        el = OxmlElement("w:fldChar" if kind else "w:instrText")
        if kind:
            el.set(qn("w:fldCharType"), kind)
        else:
            el.text = txt
        run._r.append(el)
    para, table, page_break = SOLD.para, SOLD.table, SOLD.page_break
    seq = [0]

    def R(fields):
        rec(fname, "item", fields, seq=seq[0])
        seq[0] += 1
    sub = E + timedelta(days=14)
    for _ in range(5):
        para(doc, "")
    para(doc, "【架空】", 12, False, "c")
    para(doc, f"{y}年度", 16, True, "c")
    para(doc, f"かざはやエコプラザ　{c['short']}", 20, True, "c", 8)
    para(doc, "完　成　報　告　書", 26, True, "c", 30)
    for _ in range(6):
        para(doc, "")
    para(doc, f"工　期　　{S.year}年{S.month}月{S.day}日 ～ {E.year}年{E.month}月{E.day}日", 12, False, "c")
    para(doc, f"提出日　　{sub.year}年{sub.month}月{sub.day}日", 12, False, "c", 40)
    para(doc, "提出先　【架空】かざはや環境パートナーズ株式会社　御中", 12, False, "c", 30)
    para(doc, "【架空】水樹テクノサービス株式会社", 14, True, "c")
    para(doc, "（監修　【架空】水樹プラント株式会社 環境プラント部）", 10, False, "c")
    page_break(doc)
    para(doc, "目　次", 13, True, "c", 6)
    chaps = ["1. 工事概要", "2. 工程実績", "3. 作業内容および結果"]
    ct = maker_content(c)
    if ct["wear"] or sample:
        chaps.append("4. 火格子摩耗測定記録")
    chaps += ["5. 使用材料", "6. 所見および提言", "7. 写真（写真代替図）"]
    for t in chaps:
        para(doc, "　　" + t, 10)
    para(doc, "")
    para(doc, "1. 工事概要", 12, True)
    scope = {"1": "1号炉 ストーカ・ボイラ・排ガス処理設備ほか", "2": "2号炉 ストーカ・ボイラ・排ガス処理設備ほか",
             "0": "受変電・余熱利用（タービン・ACC）・共通設備"}[c["unit"]]
    ov = [("工事件名", "【架空】" + c["name"]), ("工事場所", "【架空】かざはやエコプラザ（かざはや市）"),
          ("工期", f"{S.year}年{S.month}月{S.day}日 ～ {E.year}年{E.month}月{E.day}日"),
          ("発注者", f"【架空】かざはや環境パートナーズ株式会社（担当 {c['spc']}）"),
          ("受注者", f"【架空】水樹テクノサービス株式会社（職長 {c['foreman']}）"),
          ("協力会社", "、".join(dict.fromkeys(CREWS[t["crew"]][0] for t in c["tasks"] if CREWS[t["crew"]][0].find("水樹") < 0 and t["crew"] != "SPC運転"))),
          ("工事範囲", f"{scope}の点検整備（仕様書 KZ-{y % 100:02d}-{ {'1': '01', '2': '02', '0': '03'}[c['unit']]} による）")]
    if sample:
        ov[-1] = ("工事範囲", "1号炉 ストーカ・給じん装置・灰押出装置・耐火物・スートブロワ・減温塔・バグフィルタ・炉別MCC の点検整備（仕様書 KZ-18-01 による）")
        ov[2] = ("工期", "2018年9月3日 ～ 2018年9月28日（炉停止 9月1日 6:00、ごみ投入 9月26日 10:20）")
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    for k, v in ov:
        cells = t.add_row().cells
        SOLD.cell_text(cells[0], k, 9, True, "c")
        SOLD.shade(cells[0], "EDEDED")
        SOLD.cell_text(cells[1], v, 9)
    SOLD.set_widths(t, [30, 140])
    hdr = dict(ov)
    para(doc, "")
    para(doc, "2. 工程実績", 12, True)
    rows = []
    for tk in c["tasks"]:
        note = tk["event"][2][:24] if tk.get("event") else ""
        rows.append([tk["no"], tk["name"], f"{tk['ps'].month}/{tk['ps'].day}～{tk['pe'].month}/{tk['pe'].day}",
                     f"{tk['as_'].month}/{tk['as_'].day}～{tk['ae'].month}/{tk['ae'].day}", note])
    table(doc, ["No", "作業項目", "予定", "実績", "備考"], rows, [10, 80, 24, 24, 32], 8)
    for r_, tk in zip(rows, c["tasks"]):
        R(dict(section="工程実績", processNo=r_[0], name=r_[1], plan=r_[2], actual=r_[3], note=r_[4], contract=hdr["工事件名"]))
        LINKS.append(dict(file=fname, seq=seq[0] - 1, works=TASK_WORKS.get((c["key"], tk["no"]), [])))
    page_break(doc)
    para(doc, "3. 作業内容および結果", 12, True)
    if sample:
        wl = [["1", "No.1 ストーカ（乾燥段・燃焼段・後燃焼段）", "火格子摩耗測定（可動・固定 全列）", "燃焼段 可動側の摩耗が大きい（最大8.6mm）。4章・6章参照", "1NS-020 1NS-030 1NS-040"],
              ["2", "No.1 炉本体 耐火物", "打診・目視点検、部分補修（燃焼段右側壁 0.9m²、炉出口天井れんが 2枚）", "補修後 異常なし", "1NS-010"],
              ["3", "No.1 給じんプッシャ", "摺動部ライナ（摺動板）取替、ストローク確認", "良", "1UK-040"],
              ["4", "No.1 灰押出機", "プッシャ底板ライナ 2枚 溶接補修", "良", "1NS-170"],
              ["5", "No.1 ボイラ 長抜差スートブロワ No.4", "ランス・グランドパッキン交換、手回し・作動確認", "良", "1BO-190"],
              ["6", "No.1 減温塔 噴霧ノズル", "ノズルチップ 6個交換", "噴霧パターン確認 良", "1HG-050"],
              ["7", "No.1 バグフィルタ", "ろ布抜取 6室×2本 → 強度・通気度試験（分析センター）", "試験結果 2018/10/26 別紙報告（引張強度保持率 88～92%）", "1HG-130"],
              ["8", "No.1 動力制御盤 MCC-3", "盤内清掃・端子増締め・絶縁抵抗測定", "良", "1DK-030"],
              ["9", "No.1 ボイラ給水ポンプ（B）電動機", "絶縁診断（PI・tanδ）", "良（PI 4.1）", "1BO-110M"],
              ["10", "No.1 炉 全体", "復旧・単体／連動試運転、立上げ立会", "9/26 ごみ投入、9/27 定格負荷到達", ""]]
    else:
        wl = [[str(i + 1), w["name"], w["work"], w["result"], " ".join(w["codes"])] for i, w in enumerate(ct["works"])]
    table(doc, ["No", "機器（呼称）", "作業内容", "結果"], [x[:4] for x in wl], [8, 46, 62, 54], 8.5)
    for x in wl:
        R(dict(section="作業内容および結果", no=x[0], equipment=x[1], work=x[2], result=x[3]))
        ws_ = []
        for cd in x[4].split():
            ws_ += code_links(c, cd)
        LINKS.append(dict(file=fname, seq=seq[0] - 1, works=ws_))
    para(doc, "")
    para(doc, "※機器の呼称は当社の完成図書（機器リスト）による。", 8)
    if ct["wear"] or sample:
        page_break(doc)
        para(doc, f"4. 火格子摩耗測定記録（{y}年 定修時、デプスゲージ、単位 mm、新品基準面からの摩耗量）", 11, True)
        para(doc, "管理値：摩耗量 15mm で取替（メーカー基準）。測定点：乾燥段・後燃焼段は各列3点、燃焼段は各列5点。", 8.5)
        wr = SOLD.wear_rows() if sample else ct["wear"]
        table(doc, ["段", "側", "列", "点1", "点2", "点3", "点4", "点5", "最大", "判定"], [[str(v) for v in r] for r in wr],
              [18, 12, 12, 14, 14, 14, 14, 14, 16, 18], 8)
        stage = ""
        for r in wr:
            stage = r[0] or stage
            R(dict(section="火格子摩耗測定記録", stage=stage, side=r[1], column=r[2], points=[v for v in r[3:8] if v != ""], max=r[8], judge=r[9]))
    page_break(doc)
    para(doc, "5. 使用材料", 12, True)
    if sample:
        mats = [["給じん装置 摺動板（ライナ）", "SS400＋硬化肉盛 t25", "4枚", "当社手配", "使用"], ["灰押出機 底板ライナ", "SS400 t19", "2枚", "当社手配", "使用"],
                ["キャスタブル（高アルミナ質）", "東雲築炉 標準品", "1.8t", "協力会社手配", "使用 1.6t／残 0.2t は持帰り"], ["耐火れんが", "SK34 並形", "2枚", "協力会社手配", "使用"],
                ["スートブロワ グランドパッキン", "黒鉛系", "1式", "当社手配", "使用"], ["減温塔 ノズルチップ", "二流体 SUS316", "6個", "SPC支給（予備品）", "使用"]]
    else:
        mats = [[m[0], m[1], m[3], "SPC支給" if m[2] == "SPC支給" else ("協力会社手配" if "協力" in m[2] else "当社手配"),
                 "使用" if m[5] in ("—", "") else f"使用 {m[4]}／予備 {m[5]}"] for m in ct["materials"]]
        if not mats:
            mats = [["消耗品（ウエス・グリス等）", "—", "1式", "当社手配", "使用"]]
    table(doc, ["品名", "仕様", "数量", "手配", "使用区分"], mats, [46, 40, 16, 30, 38], 8.5)
    for m in mats:
        R(dict(section="使用材料", name=m[0], spec=m[1], qty=m[2], supply=m[3], use=m[4]))
    para(doc, "")
    para(doc, "6. 所見および提言", 12, True)
    if sample:
        remarks = ["(1) 燃焼段の可動火格子は、運転3年半で最大8.6mm（1列目・3列目）の摩耗があった。乾燥段（最大3.9mm）・後燃焼段（最大5.1mm）に比べ大きい。"
                   "摩耗速度は約2.5mm/年で、管理値15mmに達するのは2021～2022年度と推定する。長期修繕計画の更新年（2027年度）の見直しを提言する。",
                   "(2) 当面の対策として、次回定修（2019年度）で摩耗の大きい列（1・3列目）の部分交換を行うことを推奨する。",
                   "(3) 燃焼段右側壁の耐火物損傷は、火格子上0～300mmの範囲でクリンカ付着によるもの。次回も同部位を重点点検とする。",
                   "(4) その他の機器に異常はなかった。"]
    else:
        remarks = ct["remarks"]
    for tx in remarks:
        para(doc, tx, 9.5, space_after=4)
        R(dict(section="所見および提言", text=tx))
    page_break(doc)
    para(doc, "7. 写真（写真代替図）", 12, True)
    para(doc, "※本見本では写真の代わりに【架空】見本用の写真代替図（枠と説明）を置く。", 8)
    tt = doc.add_table(rows=3, cols=2)
    tt.style = "Table Grid"
    caps = [w["work"].split("（")[0][:20] for w in ct["works"]][:6] if not sample else \
        ["燃焼段 可動火格子 3列目（摩耗部）", "摩耗測定状況（デプスゲージ）", "燃焼段右側壁 損傷部（補修前）", "同 補修後", "給じん装置 摺動板 取替後", "スートブロワ No.4 ランス交換"]
    while len(caps) < 6:
        caps.append("作業状況")
    for i, cap in enumerate(caps):
        SOLD.cell_text(tt.rows[i // 2].cells[i % 2], f"【架空】見本用の写真代替図\n\n\n\n\n\n\n写真{i + 1} {cap}", 9, False, "c")
    SOLD.set_widths(tt, [85, 85])
    doc.save(docx_path)
    pdf = to_pdf(docx_path, WORK)
    os.replace(pdf, outpath(fname))
    return fname


def run():
    out = []
    for c in get():
        if c["status"] != "完了":
            continue
        y = c["fy"]
        if y <= 2020:
            fn = f"{FN}_工事報告書_{y}年度{c['short']}_メーカー様式.pdf"
            build_maker(c, fn)
        else:
            fn = f"{FN}_工事報告書_{y}年度{c['short']}_SPC様式.pdf"
            ct = c23_content(c) if c.get("sample") == "C23" else (c23c_content(c) if c.get("sample") == "C23C" else spc_content(c))
            build_spc(c, ct, fn)
        out.append((c, fn))
    return out


if __name__ == "__main__":
    import time
    t = time.time()
    for c, f in run():
        print(f)
    print(time.time() - t)
