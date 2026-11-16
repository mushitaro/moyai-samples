# -*- coding: utf-8 -*-
"""F09 帳票見本を全部作り直す：python3 -I run_all.py（scripts/ で実行）"""
import os, sys, json, shutil, random
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
random.seed(909)
import world as W, plan_gen, ledger_gen, sched_gen, daily_gen, report_gen, order_gen, render

OUT = os.path.normpath(os.path.join(HERE, "..", "out"))
TMP = os.path.normpath(os.path.join(HERE, "..", "tmp"))
P = "【架空】ときわ北部クリーンセンター_"

# (file, 生成関数, PNG にするページ, 資料, 使用者, 使用開始年度, 一行説明, infilter)
ITEMS = [
    (P + "中長期整備計画_長期修繕計画2016年度改訂_SPCsv1様式.xlsx", plan_gen.spc_sv1, [1, 14], "中長期整備計画", "SPC 保全 sv1", 2013,
     "SPC sv1 が 2013 年に作った長期修繕計画（炉別ブロック・周期列・事業 n 年目行・設備小計 百万円）", None),
    (P + "中長期整備計画_長期修繕計画2017年度改訂_SPCsv2様式.xlsx", plan_gen.spc_sv2, [1, 14, 15], "中長期整備計画", "SPC 保全 sv2", 2017,
     "sv2 が 2017 年に sv1 版をコピーして直した様式（設備順・号炉列・第2期参考列・費用は別シート）", None),
    (P + "中長期整備計画_令和7年度改訂_組合u1様式.xlsx", plan_gen.union_plan, [1, 2, 5], "中長期整備計画", "組合 施設課 u1", 2024,
     "組合 u1 が 2024 年度に新設した計画（和暦 10 年・年度ごと計画／実績／費用の 3 段・設備別シート＋総括）", None),
    (P + "工程表_2015年度2号炉定期整備工事_SPCsv1様式.xlsx", sched_gen.sv1_2015, [1], "工程表", "SPC 保全 sv1", 2009,
     "sv1 の定期整備工程表（日単位・1 工事 1 枚・予定のみの塗り）", None),
    (P + "工程表_2025年度1号炉定期整備工事_SPC2024様式.xlsx", sched_gen.sv2_2025, [1], "工程表", "SPC 保全 sv2", 2024,
     "第2期開始の 2024 年に sv2 が直した工程表（予定／実績の 2 段・班列・炉状態の帯）", None),
    (P + "工程表_基幹的設備改良工事R7施工分_メーカー様式.xlsx", sched_gen.maker_k3, [1], "工程表", "メーカー 現場代理人 m1", 2025,
     "基幹改良でメーカーが作る週単位の実施工程表（予定線と実績塗り、変更は橙）", None),
    (P + "作業日報_2015年度2号炉定期整備工事_x1の2014様式.xlsx", daily_gen.x1_2015, [6, 7], "作業日報", "協力会社 職長 x1", 2014,
     "x1 が 2014 年から使う縦型 1 日 1 枚の日報（職種別人数・KY・特記）", None),
    (P + "作業日報_2025年度1号炉定期整備工事_x1の2024様式.xlsx", daily_gen.x1_2025, [7, 9], "作業日報", "協力会社 職長 x1", 2024,
     "2024 年に工数記入のため横型へ直した日報（班別・工数・進捗・写真No.、右端に 2014 様式の職種別人員）", None),
    (P + "作業日報_基幹的設備改良工事R7施工分_メーカー週報様式.xlsx", daily_gen.maker_k3, [8, 11], "作業日報", "メーカー m2", 2025,
     "基幹改良でメーカーが書く 1 週 1 枚の週報（行＝工種・作業、列＝日、出来高・人員・立会）", None),
    (P + "工事報告書_2015年度2号炉定期整備工事_SPCsv1文章型.pdf", None, [1, 2, 3], "工事報告書", "SPC 保全 sv1", 2009,
     "sv1 が Word で書く文章型の完了報告書（表紙・本文・部品表）", None),
    (P + "工事報告書_2025年度1号炉定期整備工事_SPCsv2表型2018様式.pdf", None, [1, 2, 3], "工事報告書", "SPC 保全 sv2", 2018,
     "sv2 が 2018 年に直した報告書（1 枚目は件名・記書き・印欄、添付にサマリー表・詳細・部品・測定）", None),
    (P + "工事報告書_基幹的設備改良工事R7施工分_メーカー出来形写真帳.pdf", None, [1, 2, 3], "工事報告書", "メーカー m1", 2025,
     "メーカーが年度ごとに出す出来形総括と写真帳中心の施工報告書", None),
    (P + "発注台帳_2015年度_SPC旧購買システム.csv", order_gen.old_csv, [1], "発注台帳", "SPC 購買", 2009,
     "SPC 旧購買システムの発注明細 CSV（Shift_JIS・英字列名・半角カナ）", "CSV:44,34,64"),
    (P + "発注台帳_2025年度_SPC新購買システム.csv", order_gen.new_csv, [1], "発注台帳", "SPC 購買", 2016,
     "2016 年更新の SPC 購買システム CSV（UTF-8 BOM、発注ヘッダ行 H と明細行 D、工事件名・設備タグ）", "CSV:44,34,76"),
    (P + "発注台帳_基幹的設備改良事業R7-R8_組合財務会計.csv", order_gen.union_csv, [1], "発注台帳", "組合 財務（施設課 u1 出力）", 2025,
     "組合の財務会計システムの支出負担行為・支出一覧 CSV（予算科目はコード＋名称 1 列、和暦）", "CSV:44,34,64"),
    (P + "機器台帳_2023年度末_SPC保全システム.xlsx", ledger_gen.spc_ledger, [1, 7], "機器台帳", "SPC 保全", 2009,
     "SPC 保全システムの設備ツリー出力（系統→設備→機器→付属→電動機、台帳対象は機器行）", None),
    (P + "機器台帳_令和8年4月1日現在_組合様式.xlsx", ledger_gen.union_ledger, [1, 6, 22], "機器台帳", "組合 施設課 u1", 2024,
     "2024 年に SPC 台帳を変換した組合の設備別シート台帳（R7 更新で新旧番号を併記、記録シート追加）", None),
]

def main():
    import notes_gen; notes_gen.main()
    os.makedirs(OUT, exist_ok=True); os.makedirs(TMP, exist_ok=True)
    for f in os.listdir(OUT):
        os.remove(os.path.join(OUT, f))
    rw = os.path.join(TMP, "render")
    if os.path.exists(rw): shutil.rmtree(rw)
    manifest = []
    for fn, gen, pages, doc, user, fy, desc, inf in ITEMS:
        path = os.path.join(OUT, fn)
        if "SPCsv1文章型" in fn:
            report_gen.sv1_2015(path, TMP)
        elif "SPCsv2表型" in fn:
            report_gen.sv2_2025(path)
        elif "メーカー出来形" in fn:
            report_gen.maker_k3(path)
        else:
            gen(path)
        pngs, n = render.to_png(path, pages, OUT, infilter=inf)
        print(f"{fn}  pages={n}  png={pngs}")
        manifest.append({"file": fn, "png": pngs, "資料": doc, "様式の使用者": user, "使用開始年度": fy, "一行説明": desc})
    with open(os.path.join(OUT, "manifest.json"), "w", encoding="utf-8") as fp:
        json.dump(manifest, fp, ensure_ascii=False, indent=1)

if __name__ == "__main__":
    main()
