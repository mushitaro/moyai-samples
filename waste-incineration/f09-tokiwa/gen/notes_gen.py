# -*- coding: utf-8 -*-
"""notes/ の集計部分（台数内訳・代表機器の根拠記録・作業の対応表）を世界から書き出す"""
import os, sys, json
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import world as W, works as K
from collections import Counter, OrderedDict
NOTES = os.path.join(HERE, "..", "notes")

def main():
    os.makedirs(NOTES, exist_ok=True)
    L = ["# 設備階層と管理対象機器の台数（生成器から自動出力）", "",
         f"合計 {len(W.EQS)} 台＝炉別 {W.counts()[1]}×2＋共通 {W.counts()[0]}（プロファイル 2.5：205×2＋238＝648）", "",
         "数え方：タグ番号 1 件＝1 台。弁は電動弁・調節弁（電磁弁・安全弁を含む）のみ、計装は分析計・主要流量計のみ。手動弁・一般計装（温度・圧力・レベル）は台帳外。",
         "常用・予備は別の台（A/B、1号/2号、No.1〜3）。共通設備は 1 組のみ（炉へ重複計上しない）。部品（ろ布・散気ノズル・耐火物・カッタ刃）は機器にしない。", "",
         "## 設備（大分類）× 号炉", "", "| 設備 | 1号炉 | 2号炉 | 共通 | 計 |", "|---|---:|---:|---:|---:|"]
    groups = OrderedDict()
    for e in W.EQS:
        groups.setdefault(e.group, Counter())[e.furn] += 1
    for g, c in groups.items():
        L.append(f"| {g} | {c[1]} | {c[2]} | {c[0]} | {sum(c.values())} |")
    L += ["", "## 系統（中分類）ごと（1 炉分と共通）", "", "| 系統コード | 設備 | 系統 | 1 炉あたり | 共通 |", "|---|---|---|---:|---:|"]
    subs = OrderedDict()
    for e in W.EQS:
        k = (e.sys, e.group, e.sub)
        subs.setdefault(k, Counter())[e.furn] += 1
    for (sy, g, sb), c in sorted(subs.items()):
        L.append(f"| {sy} | {g} | {sb} | {c[1]} | {c[0]} |")
    L += ["", "## 常用・予備の組（抜粋）", ""]
    for e in W.EQS:
        if e.duty in ("常用", "予備") and e.furn == 0:
            L.append(f"- {e.tag} {e.name}：{e.duty}")
    open(os.path.join(NOTES, "01_設備階層と台数.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")

    L = ["# 代表機器の根拠記録（12 台）", "", "過去の作業 → 現在の状態 → 放置すると → なぜその年度・その作業か → 採らなかった案。計画の版ごとの印は plan_gen.rep_marks にある。", ""]
    for r in W.REP:
        e = W.BY_TAG[r["tag"]]
        L += [f"## {r['key']} {e.tag} {e.furn_label()} {e.name}（重要度 {e.crit}、{e.spec}）", "",
              "過去の作業：", *[f"- {y} 年度 {a}：{t}" for y, a, t in r["history"]], "",
              f"- 現在の状態：{r['state']}", f"- 放置すると：{r['risk']}", f"- 選んだ年度・作業：{r['plan']}", f"- 採らなかった案と理由：{r['why_not']}", ""]
    open(os.path.join(NOTES, "02_根拠記録.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")

    L = ["# 工程 → 日報 → 報告書 → 発注 の対応（生成器から自動出力）", ""]
    for key, tasks in (("K1", K.K1), ("K2", K.K2)):
        c = W.CONTRACTS[key]
        L += [f"## {key} {c['title']}（{c['kind']}）", "", "| 工程 | 作業 | 機器タグ | 属性 | 予定 | 実績（日報） | 日報の日数 | 延べ人数 |", "|---|---|---|---|---|---|---:|---:|"]
        for t in tasks:
            ppl = sum(sum(cr.values()) for _, _, cr in t.steps)
            L.append(f"| {t.tid} | {t.name} | {' '.join(t.tags)} | {t.attr} | {t.plan[0]:%m/%d}〜{t.plan[1]:%m/%d} | {t.actual[0]:%m/%d}〜{t.actual[1]:%m/%d} | {len(t.steps)} | {ppl} |")
        L.append("")
    L += ["## K3 基幹的設備改良工事（令和7年度 2号炉施工分）", "", "| 工種 | 作業 | 機器タグ | 属性 | 予定 | 実績 | 1日人員 |", "|---|---|---|---|---|---|---:|"]
    for kou, name, tags, attr, ps, pe, as_, ae, crew in K.K3:
        L.append(f"| {kou} | {name} | {' '.join(tags)} | {attr} | {ps:%Y-%m-%d}〜{pe:%m/%d} | {as_:%Y-%m-%d}〜{ae:%m/%d} | {crew} |")
    open(os.path.join(NOTES, "03_工程と作業の対応_自動.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")

if __name__ == "__main__":
    main()
