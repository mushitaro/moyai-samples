# -*- coding: utf-8 -*-
"""正解の世界：契約、実際に行った作業（2005–2024 年度）、計画の 2 つの版、2016 年度の区分。"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import rng, wareki_year, MAKER, LOCAL  # noqa: E402
from world import build_equipment  # noqa: E402

FY0, FY1 = 2005, 2024
EQ = build_equipment()
BYID = {e["id"]: e for e in EQ}


def find(line, base, no=None):
    for e in EQ:
        if e["line"] == line and e["base"] == base and (no is None or e["name"].endswith(f"No.{no}")):
            return e["id"]
    raise KeyError((line, base, no))


# ---------------------------------------------------------------- 契約
def contracts():
    c = {}
    for fy in range(FY0, 2019):
        c[f"T{fy}"] = dict(fy=fy, vendor=MAKER, kind="定期整備業務委託",
                           name=f"{wareki_year(fy, 'full')} いずみ川清掃センター 定期整備業務委託")
    c["P2008"] = dict(fy=2008, vendor=MAKER, kind="改良工事（請負・2 か年）",
                      name="いずみ川清掃センター 部分改良工事（平成20〜21年度）")
    c["R2014"] = dict(fy=2014, vendor=MAKER, kind="補修工事（請負）",
                      name="平成26年度 いずみ川清掃センター 1号炉ボイラ水管補修工事")
    local = {
        2019: ["1号炉ボイラ水管補修工事"],
        2020: ["1号炉誘引通風機補修工事", "ボイラ性能検査整備工事"],
        2021: ["前処理破砕機補修工事", "ボイラ性能検査整備工事"],
        2022: ["1号炉バグフィルタろ布部分交換工事", "ボイラ性能検査整備工事"],
        2023: ["2号炉散気ノズル補修工事", "ボイラ性能検査整備工事"],
        2024: ["ボイラ性能検査整備補修工事"],      # 最終年度は 1 件のみ
    }
    for fy in range(2019, 2025):
        c[f"K{fy}"] = dict(fy=fy, vendor="黒潮クレーン株式会社", kind="点検委託",
                           name=f"{wareki_year(fy, 'full')} ごみクレーン年次自主検査業務委託")
    c["S2020a"] = dict(fy=2020, vendor="紫雲電機株式会社", kind="修繕（需用費）", name="中央監視制御盤修繕")
    c["S2020b"] = dict(fy=2020, vendor="紫雲電機株式会社", kind="工事請負", name="計装用空気圧縮機更新工事")
    for fy, names in local.items():
        for k, nm in enumerate(names):
            c[f"L{fy}{k}"] = dict(fy=fy, vendor=LOCAL, kind="補修工事（請負）",
                                  name=f"{wareki_year(fy, 'full')} いずみ川清掃センター {nm}")
    return c


CONTRACTS = contracts()

# ------------------------------------------------- 機器の種別ごとの整備の立て方
# (点検周期, 分解整備周期)  None は周期で入れない（状態を見て入れる）
STRAT = {
    "流動床炉本体": (1, None), "散気装置": (1, None), "廃熱ボイラ": (1, None),
    "エコノマイザ": (1, None), "バグフィルタ": (1, None), "減温塔": (1, None),
    "誘引通風機": (1, 4), "押込送風機": (1, 5), "定量給じん機": (1, 3),
    "不燃物抜出装置": (1, 3), "振動ふるい": (2, 4), "砂循環エレベータ": (1, None),
    "ボイラ給水ポンプ": (2, 4), "前処理破砕機": (1, 3), "ごみクレーン": (1, None),
    "混練機": (1, 4), "計装用空気圧縮機": (1, 4), "非常用発電機": (1, 6),
    "脱気器": (1, None), "高圧蒸気だめ": (1, None), "余剰蒸気復水器": (2, None),
}
DEFAULT = {"A": (2, 6), "B": (3, None), "C": (None, None)}

# 法定点検（計画の「法定点検等」の行）。確度は notes/05 に書く。
LEGAL = [
    ("ボイラー性能検査（1号炉）", 1, "ボイラー及び圧力容器安全規則"),
    ("ボイラー性能検査（2号炉）", 1, "同上"),
    ("第一種圧力容器性能検査（脱気器・高圧蒸気だめ）", 1, "同上"),
    ("第二種圧力容器定期自主検査（空気槽）", 1, "同上"),
    ("クレーン性能検査（ごみクレーン）", 2, "クレーン等安全規則"),
    ("クレーン年次自主検査（ごみクレーン・ホイスト）", 1, "同上"),
    ("受変電設備 年次点検（停電）", 1, "電気事業法・保安規程"),
    ("非常用発電機 総合点検", 1, "消防法・保安規程"),
    ("ばい煙測定（年2回）", 1, "大気汚染防止法"),
    ("ダイオキシン類測定（排ガス・飛灰）", 1, "ダイオキシン類対策特別措置法"),
    ("作業環境測定（ダイオキシン類）", 1, "労働安全衛生規則 第592条の2"),
    ("計量機 定期検査", 2, "計量法"),
    ("消防用設備等点検（総合）", 1, "消防法"),
    ("精密機能検査", 3, "廃棄物処理法施行規則"),
]


# ---------------------------------------------------------------- 実績
def _w(eq, fy, kind, label, contract, md=None, days=1, mh=None, vendor=None):
    return dict(eq=eq, fy=fy, kind=kind, label=label, contract=contract, md=md,
                days=days, mh=mh, vendor=vendor or CONTRACTS.get(contract, {}).get("vendor", ""))


def special_works():
    """代表機器の作業（根拠記録の対象）。"""
    W = []
    sanki2, sanki1 = find(2, "散気装置"), find(1, "散気装置")
    W += [_w(find(1, "流動床炉本体"), 2010, "OVERHAUL", "フリーボード部耐火物部分補修（点検時発見・追加 変更契約）",
             "T2010", md=(9, 14), days=4)]
    for e in EQ:
        if e["year"] in (2008, 2009) and e["line"] in (1, 2) and e["base"] not in ("散気装置", "砂循環エレベータ"):
            W.append(_w(e["id"], e["year"], "UPGRADE",
                        ("更新（部分改良）" if e.get("old_tag") else "新設（部分改良）"), "P2008"))
    W += [_w(sanki1, 2008, "UPGRADE", "散気装置更新（散気ノズル 120 本・ヘッダ）", "P2008"),
          _w(sanki2, 2009, "UPGRADE", "散気装置更新（散気ノズル 120 本・ヘッダ）", "P2008"),
          _w(sanki2, 2021, "INSPECT", "散気ノズル目視点検（閉塞 11・摩耗 7 本）", "L20211"),
          _w(sanki2, 2023, "REPLACE", "散気ノズル 96 本取替（予備品支給）・風箱清掃", "L20230",
             md=(10, 16), days=12, mh=None)]      # 工数は日報から（jobs.J23_DAILY）
    b1 = find(1, "廃熱ボイラ")
    W += [_w(b1, 2014, "OVERHAUL", "水管 6 本部分取替（第2パス）", "R2014"),
          _w(b1, 2019, "OVERHAUL", "水管 2 本部分取替（漏えい箇所）", "L20190")]
    cr1, cr2 = find(0, "前処理破砕機", 1), find(0, "前処理破砕機", 2)
    W += [_w(cr1, 2021, "OVERHAUL", "減速機取替（No.2 予備機から転用）", "L20210"),
          _w(cr2, 2021, "INSPECT", "減速機取外し（部品取り）", "L20210")]
    idf1 = find(1, "誘引通風機")
    W += [_w(idf1, 2020, "OVERHAUL", "羽根車肉盛・動的バランス・軸受交換", "L20200")]
    bf1, bf2 = find(1, "バグフィルタ"), find(2, "バグフィルタ")
    W += [_w(bf1, 2021, "INSPECT", "内部点検（ろ布破損 8 本・閉止板処置）", "L20211"),
          _w(bf1, 2005, "REPLACE", "ろ布全数交換（480 本）", "T2005"),
          _w(bf2, 2006, "REPLACE", "ろ布全数交換（480 本）", "T2006"),
          _w(bf1, 2011, "REPLACE", "ろ布全数交換（480 本）", "T2011"),
          _w(bf2, 2012, "REPLACE", "ろ布全数交換（480 本）", "T2012"),
          _w(bf1, 2016, "REPLACE", "ろ布全数交換（480 本）", "T2016"),
          _w(bf2, 2017, "REPLACE", "ろ布全数交換（480 本）", "T2017"),
          _w(bf1, 2022, "OVERHAUL", "ろ布部分交換 32 本（破損）", "L20220")]
    el1 = find(1, "砂循環エレベータ")
    W += [_w(el1, 2008, "UPGRADE", "砂循環エレベータ更新（部分改良）", "P2008"),
          _w(el1, 2014, "OVERHAUL", "チェーン・バケット交換", "T2014")]
    W += [_w(find(2, "砂循環エレベータ"), 2009, "UPGRADE", "砂循環エレベータ更新（部分改良）", "P2008"),
          _w(find(2, "砂循環エレベータ"), 2015, "OVERHAUL", "チェーン・バケット交換", "T2015")]
    gt2 = find(2, "減温塔")
    W += [_w(gt2, 2013, "OVERHAUL", "下部コーン内面腐食部 当て板補修", "T2013")]
    w1 = find(0, "計量機", 1)
    W += [_w(w1, 2014, "REPLACE", "ロードセル 4 個更新", "T2014")]
    cms = find(0, "中央監視制御盤")
    W += [_w(cms, 2020, "OVERHAUL", "CPU 基板故障 予備基板へ交換", "S2020a", vendor="紫雲電機")]
    ac1 = find(0, "計装用空気圧縮機", 1)
    W += [_w(ac1, 2020, "REPLACE", "空気圧縮機更新（故障・代替なし）", "S2020b", vendor="紫雲電機")]
    # 2019–2024：法定の年次検査に付く開放点検（地元業者の補修工事・市の直接契約）
    for fy in range(2019, 2025):
        lc = f"L{fy}0" if fy in (2019, 2024) else f"L{fy}1"
        for ln in (1, 2):
            W.append(_w(find(ln, "廃熱ボイラ"), fy, "INSPECT", "性能検査受検整備（開放・清掃）", lc))
        for b in ("脱気器", "高圧蒸気だめ"):
            W.append(_w(find(0, b), fy, "INSPECT", "性能検査受検整備", lc))
        for k in (1, 2):
            W.append(_w(find(0, "ごみクレーン", k), fy, "INSPECT",
                        "年次自主検査" + ("・性能検査" if fy % 2 == 0 else ""), f"K{fy}", vendor="黒潮クレーン"))
    W += [_w(find(2, "ボイラ給水ポンプ", 1), 2024, "OVERHAUL", "メカニカルシール交換（漏れ・追加）", "L20240")]
    crc = find(0, "ごみクレーン制御盤")
    W += [_w(crc, 2013, "REPLACE", "クレーン制御盤更新（インバータ化）", "T2013")]
    return W


def _season_contract(fy):
    return f"T{fy}" if fy <= 2018 else None


def periodic_works():
    """メーカー定期整備（2005–2018）の点検・分解整備。2017–18 は停止方針で縮小。"""
    r = rng("actual")
    W = []
    for e in EQ:
        ins, oh = STRAT.get(e["base"], DEFAULT[e["imp"]])
        if e["duty"] == "予備" and ins:
            ins = ins + 1
        ph_i = r.randint(0, 2)
        ph_o = r.randint(0, 5)
        for fy in range(FY0, 2019):
            if fy >= 2017 and e["imp"] != "A":
                continue
            # 乱数は従来どおり引き（見本 H22 と同じ作業になる）、まだ設置されていない機器は作業を記録しない
            ok = not (fy < e["year"] and not e.get("old_tag"))
            if oh and (fy + ph_o + e["line"]) % oh == 0 and fy < 2017:
                if r.random() > 0.06:
                    if ok:
                        W.append(_w(e["id"], fy, "OVERHAUL", "分解整備", f"T{fy}"))
                    continue
            if ins and (fy + ph_i + e["line"]) % ins == 0:
                if r.random() > 0.08 and ok:
                    W.append(_w(e["id"], fy, "INSPECT", "点検", f"T{fy}"))
            elif e["imp"] in "AB" and r.random() < 0.03 and ok:
                W.append(_w(e["id"], fy, "OVERHAUL", "補修（点検時発見）", f"T{fy}"))
    return W


_CACHE = []


def actual_works():
    if _CACHE:
        return _CACHE
    W = periodic_works() + special_works()
    W.sort(key=lambda w: (w["fy"], w["eq"], w["kind"]))
    for k, w in enumerate(W, 1):
        w["wid"] = f"W{k:04d}"
    _CACHE.extend(W)
    return _CACHE


# ---------------------------------------------------------------- 計画
SYM_OLD = {"INSPECT": "○", "OVERHAUL": "◎", "REPLACE": "●", "REPAIR": "△"}


def plan_old(ver=2015):
    """H21 策定（i1）→ H27 改訂（i2 が H27 以降の欄だけ直した）。対象 H22–H31。
    過去の欄は当初計画のまま残す。戻り値 {eq: {fy: 記号}}、改訂コメント。"""
    r = rng("plan_old")
    P, comments = {}, {}
    for e in EQ:
        ins, oh = STRAT.get(e["base"], DEFAULT[e["imp"]])
        if e["duty"] == "予備" and ins:
            ins += 1
        ph_i, ph_o = r.randint(0, 2), r.randint(0, 5)
        row = {}
        for fy in range(2010, 2020):
            if oh and (fy + ph_o + e["line"]) % oh == 0:
                row[fy] = "◎"
            elif ins and (fy + ph_i + e["line"]) % ins == 0:
                row[fy] = "○"
        P[e["id"]] = row
    # 当初計画（メーカー提案を i1 が写した）の更新予定
    def put(eid, fy, s, cm=None):
        P[eid][fy] = s
        if cm:
            comments[(eid, fy)] = cm
    # ver：版の年度（2009＝H21.12 策定、2012＝H24.5 改訂、2015＝H27.5 改訂）。その版で知っていたことだけ。
    put(find(1, "バグフィルタ"), 2011, "●")
    put(find(2, "バグフィルタ"), 2012, "●")
    put(find(1, "バグフィルタ"), 2017, "●")
    put(find(2, "バグフィルタ"), 2018, "●")
    put(find(0, "中央監視制御盤"), 2013, "●", "H25予算見送り。H29へ（H27改訂 i2）" if ver >= 2015 else None)
    if ver >= 2015:
        put(find(0, "中央監視制御盤"), 2017, "●")
    put(find(0, "ごみクレーン", 1), 2015, "●", "H27→H30へ延期（H27改訂）" if ver >= 2015 else None)
    put(find(0, "ごみクレーン", 2), 2016, "●", "H28→H31へ延期（H27改訂）" if ver >= 2015 else None)
    if ver >= 2015:
        put(find(0, "ごみクレーン", 1), 2018, "●")
        put(find(0, "ごみクレーン", 2), 2019, "●")
    put(find(0, "前処理破砕機", 2), 2012, "●", "H24見送り（予算）i1" if ver >= 2012 else None)
    if ver >= 2012:
        put(find(1, "廃熱ボイラ"), 2014, "△", "第2パス水管 部分取替（H22定修所見）")
    if ver >= 2015:
        put(find(2, "廃熱ボイラ"), 2016, "△", "第2パス水管 部分取替（予定）")
    put(find(0, "ごみクレーン制御盤"), 2013, "●")
    return P, comments


SYM_NEW = {"○": "点検", "◎": "分解整備", "△": "補修", "▲": "停止までの最小限の補修",
           "◇": "予備品で交換", "●": "更新（やむを得ないもの）"}


def plan_new(ver=2022):
    """H28 作り替え（i2）→ R4 改訂（i2、令和4年10月）。対象 H29–R6。
    記録は計画と実績の 2 行。R5 以降は更新されない（未更新期間）。"""
    r = rng("plan_new")
    P = {}
    for e in EQ:
        row = {}
        ins, _ = STRAT.get(e["base"], DEFAULT[e["imp"]])
        if e["imp"] == "A" and ins == 1 and e["base"] in (
                "流動床炉本体", "散気装置", "廃熱ボイラ", "バグフィルタ", "減温塔", "誘引通風機",
                "ごみクレーン", "脱気器", "高圧蒸気だめ"):
            for fy in range(2017, 2025):
                if fy <= 2018 or e["base"] in ("廃熱ボイラ", "ごみクレーン", "脱気器", "高圧蒸気だめ") \
                        or fy % 2 == e["line"] % 2:
                    row[fy] = "○"
        elif e["imp"] == "A":
            # H29・H30 はメーカー定期整備が続く（A 機器のみ点検）。R1 以降は入れない
            row[2017] = row[2018] = "○"
        elif e["imp"] == "B" and r.random() < 0.12:
            row[2017 + r.randint(0, 3)] = "▲"
        P[e["id"]] = row

    def put(eid, fy, s):
        P[eid][fy] = s
    # ver：2016＝H28.6 作成、2019＝R1.10 改訂、2022＝R4.10 改訂（最後）
    put(find(2, "バグフィルタ"), 2017, "●")           # ろ布全数交換（停止まで持たせる最後の交換）
    put(find(1, "廃熱ボイラ"), 2019, "▲")
    put(find(1, "砂循環エレベータ"), 2020, "▲")       # 実施せず（延び量に余裕）
    put(find(0, "前処理破砕機", 1), 2021, "▲")
    put(find(0, "中央監視制御盤"), 2017, "◇")         # 予備基板の確保（更新は中止）
    if ver >= 2019:
        put(find(1, "誘引通風機"), 2020, "▲")         # R1 の振動上昇
    if ver >= 2022:
        put(find(1, "バグフィルタ"), 2022, "◇")       # R3 点検で破損
        put(find(2, "散気装置"), 2023, "◇")           # R3 点検、R4 予備品購入
        for k in (1, 2):
            P[find(0, "ごみクレーン", k)].pop(2024, None)  # 抜け：R6 の性能検査を入れていない
    return P


# ---------------------------------------------------------------- 2016 年度の区分（台帳新様式）
def policy():
    pol = {}
    spare = {"散気装置", "定量給じん機", "中央監視制御盤", "誘引通風機速度制御装置",
             "ごみクレーン制御盤", "バグフィルタ", "混練機"}
    for e in EQ:
        if e["base"] in spare:
            pol[e["id"]] = "予備品で延命"
        else:
            pol[e["id"]] = "停止まで使用"
    pol[find(0, "前処理破砕機", 2)] = "部品取り（R3〜）"
    pol[find(0, "計装用空気圧縮機", 1)] = "停止まで使用（R2更新）"
    return pol


if __name__ == "__main__":
    W = actual_works()
    print(len(W), sum(1 for w in W if w["fy"] >= 2019))
    po, cm = plan_old()
    print(sum(len(v) for v in po.values()), len(cm))
    print(sum(len(v) for v in plan_new().values()))
