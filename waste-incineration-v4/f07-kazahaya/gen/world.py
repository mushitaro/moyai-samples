# -*- coding: utf-8 -*-
"""【架空】かざはやエコプラザ（F07）正解の世界：機器・整備方針・計画（版別）・実績。

物語の「今日」= 2026-11-13。乱数は使わず、機器コードの crc32 で決まる値で揺らす（再現性）。
作業の属性：I=点検・測定(INSPECT) O=分解整備・補修(OVERHAUL) R=部分補修(OVERHAUL扱い)
            X=更新(REPLACE) L=法定検査(INSPECT, 法定) U=改良(UPGRADE、本施設は 2026 年度まで無し)
"""
import zlib
from datetime import date
from world_equipment import build

TODAY = date(2026, 11, 13)
FAC = "【架空】かざはやエコプラザ"
SPC = "【架空】かざはや環境パートナーズ株式会社"
MAKER = "【架空】水樹プラント株式会社"
SUBS = "【架空】水樹テクノサービス株式会社"
YEARS = list(range(2015, 2035))  # 事業期間 20 年（2015–2034 年度）


def h(code, mod):
    return zlib.crc32(code.encode("utf-8")) % mod


F1, F2, CM = build()
ALL = F1 + F2 + CM
BY_CODE = {e["code"]: e for e in ALL}


def find(unit, base, sfx=""):
    for e in ALL:
        if e["unit"] == unit and e["base"] == base and e["sfx"] == sfx:
            return e
    raise KeyError((unit, base, sfx))


# ---------------------------------------------------------------- 整備方針（規則）
def rule(e, rev, ver=2026):
    """機器 e の 2015–2034 年度の整備予定 {年度: [(属性, 作業名)]}。
    rev=False: 事業契約時の長期修繕計画（〜2019 年度版まで）。
    rev=True : 2020 年度 5 年目見直し以降の知見（摩耗実績、部品供給終了の通知など）を反映。"""
    b, u, kind, imp = e["base"], e["unit"], e["kind"], e["imp"]
    code = e["code"]
    off = h(code, 6)
    m = {}

    def add(y, t, w):
        if 2015 <= y <= 2034:
            m.setdefault(y, []).append((t, w))

    fu = int(u) if u in ("1", "2") else 0
    # ---- 電動機 ----
    if kind == "電動機":
        kw = e["kw"]
        if kw >= 200:  # 高圧（誘引通風機）
            for y in ([2019, 2025, 2031] if not rev else [2019, 2023, 2029]):
                add(y + (fu == 2), "O", "電動機軸受交換・精密点検")
            for y in range(2017 + (fu == 2), 2035, 2):
                add(y, "I", "絶縁診断")
            return m
        if kw >= 37:
            for y in range(2018 + off % 3, 2035, 6):
                add(y, "I", "絶縁診断")
            if b.startswith("ごみクレーン巻上") or b.startswith("空冷復水器ファン"):
                add(2027 + off % 3, "O", "電動機分解整備")
            return m
        if kw >= 11 and imp == "A":
            add(2022 + off % 4, "O", "電動機軸受交換")
        return m  # 小形電動機は事後保全（計画に載せない）

    # ---- 燃焼設備：火格子 ----
    if b == "燃焼火格子":
        for y in YEARS:
            add(y, "I", "火格子摩耗測定")
        if not rev:
            add(2019 + (fu == 2), "R", "火格子部分交換")
            add(2023 + (fu == 2), "R", "火格子部分交換")
            add(2027 + (fu == 2), "X", "燃焼火格子更新")
        else:
            add(2019 + (fu == 2), "R", "火格子部分交換")
            add(2023 + (fu == 2), "X", "燃焼火格子更新")
            add(2029 + (fu == 2), "R", "火格子部分交換")
            add(2033 + (fu == 2), "X", "燃焼火格子更新")
        return m
    if b == "乾燥火格子":
        for y in range(2016, 2035, 2):
            add(y + (fu == 2), "I", "火格子摩耗測定")
        add((2029 if not rev else 2027) + (fu == 2), "X", "乾燥火格子更新")
        return m
    if b == "後燃焼火格子":
        for y in range(2016, 2035, 2):
            add(y + (fu == 2), "I", "火格子摩耗測定")
        if not rev:
            add(2030 + (fu == 2), "X", "後燃焼火格子更新")
        else:
            add(2025 + (fu == 2), "R", "火格子部分交換")
            add(2030 + (fu == 2), "X", "後燃焼火格子更新")
        return m
    if "油圧シリンダ" in b:
        base_y = 2020 + off % 3
        for y in range(base_y, 2035, 5):
            add(y, "O", "シリンダパッキン交換")
        return m
    if b in ("燃焼装置油圧ポンプ", "灰押出油圧ポンプ", "投入扉油圧ポンプ"):
        y0 = 2020 + (0 if e["sfx"] == "A" else 3) + (fu == 2)
        for y in range(y0, 2035, 6):
            add(y, "O", "ポンプ分解整備")
        return m
    if b in ("燃焼装置油圧ユニット", "灰押出油圧ユニット", "投入扉油圧ユニット"):
        for y in range(2018 + off % 2, 2035, 3):
            add(y, "I", "作動油分析・フィルタ交換")
        return m
    if b == "焼却炉本体":
        for y in YEARS[1:]:
            add(y, "I", "炉内耐火物点検")
        for y in range(2018 + (fu == 2), 2035, 2):
            add(y, "R", "耐火物部分補修")
        return m
    if b in ("給じん装置", "主灰押出装置", "ホッパブリッジ解除装置"):
        for y in range(2017 + off % 3, 2035, 4):
            add(y, "O", "摺動部・ライナ補修")
        return m
    # ---- 送風機 ----
    if b == "誘引通風機":
        if not rev:
            add(2019 + (fu == 2), "O", "軸受交換・インペラ点検")
            add(2025 + (fu == 2), "O", "軸受交換・インペラ点検")
            add(2031 + (fu == 2), "X", "インペラ更新")
        else:
            add(2019 + (fu == 2), "O", "軸受交換・インペラ点検")
            add(2023 + (fu == 2), "O", "軸受交換・インペラ肉盛補修")
            add(2027 + (fu == 2), "I", "インペラ摩耗測定")
            add(2029 + (fu == 2), "X", "インペラ更新")
        return m
    if b in ("押込送風機", "二次送風機", "排ガス再循環送風機", "炉側壁冷却ファン", "脱臭送風機", "環境集じんファン"):
        y0 = 2020 + off % 3
        for y in range(y0, 2035, 6 if b != "排ガス再循環送風機" else 4):
            add(y, "O", "軸受交換・インペラ清掃")
        if b == "排ガス再循環送風機" and rev:
            add(2028 + (fu == 2), "X", "インペラ更新（耐食材）")
        return m
    # ---- ボイラ ----
    if b in ("ボイラ本体", "水管壁（炉頂・放射室）", "蒸発管群", "1次過熱器", "2次過熱器", "3次過熱器", "エコノマイザ", "蒸気ドラム"):
        start = 2017 if fu == 1 else 2016
        for y in range(start, 2035, 2):
            add(y, "L" if b == "ボイラ本体" else "I", "定期事業者検査" if b == "ボイラ本体" else "肉厚測定")
        if b == "3次過熱器":
            add((2029 if not rev else 2027) + (fu == 2), "R", "下段管部分更新・プロテクタ交換")
        if b == "水管壁（炉頂・放射室）":
            add((2026 if not rev else 2025) + (fu == 2), "R", "肉盛補修")
        if b == "2次過熱器":
            add(2031 + (fu == 2), "R", "プロテクタ交換")
        return m
    if b == "ボイラ給水ポンプ":
        y0 = (2017 if e["sfx"] == "A" else 2020) + (fu == 2)
        for y in range(y0, 2035, 6):
            add(y, "O", "ポンプ分解整備")
        return m
    if b in ("長抜差型スートブロワ",):
        y0 = 2018 + int(e["sfx"]) % 4 + (fu == 2)
        for y in range(y0, 2035, 4):
            add(y, "O", "ランス・パッキン交換")
        return m
    if b == "定置型スートブロワ":
        if rev and int(e["sfx"]) in (5, 6):
            add(2026 + (fu == 2), "X", "ノズル管更新")
        return m
    if b == "過熱器減温器":
        add(2024 + (fu == 2), "I", "スプレーノズル点検")
        return m
    # ---- 排ガス処理 ----
    if b == "バグフィルタ本体":
        if not rev:
            add(2021 + (fu == 2), "X", "ろ布全数交換")
            add(2028 + (fu == 2), "X", "ろ布全数交換")
            for y in (2018, 2020, 2025):
                add(y + (fu == 2), "I", "ろ布抜取試験")
        else:
            add(2021 + (fu == 2), "X", "ろ布全数交換")
            add(2027 + (fu == 2), "X", "ろ布全数交換")
            for y in (2018, 2020, 2023, 2025):
                add(y + (fu == 2), "I", "ろ布抜取試験")
        return m
    if b == "バグフィルタ逆洗装置":
        for y in range(2019 + (fu == 2), 2035, 4):
            add(y, "O", "電磁弁ダイヤフラム交換")
        return m
    if b == "触媒反応塔":
        for y in (2017, 2020, 2023, 2026, 2029, 2032):
            add(y + (fu == 2), "I", "触媒性能試験")
        if not rev:
            add(2024 + (fu == 2), "X", "予備層触媒充填")
            add(2030 + (fu == 2), "X", "触媒交換（1層）")
        else:
            add(2025 + (fu == 2), "X", "予備層触媒充填")
            add(2031 + (fu == 2), "X", "触媒交換（1層）")
        return m
    if b == "減温塔噴霧ノズルユニット":
        for y in range(2016 + (fu == 2), 2035, 2):
            add(y, "O", "ノズルチップ交換")
        return m
    if b == "減温塔噴霧水ポンプ":
        y0 = 2019 + (0 if e["sfx"] == "A" else 3) + (fu == 2)
        for y in range(y0, 2035, 5):
            add(y, "O", "ポンプ分解整備")
        return m
    if b in ("減温塔", "煙突内筒"):
        for y in range(2019 + (fu == 2), 2035, 4):
            add(y, "I", "内面点検・肉厚測定")
        return m
    if b == "排ガス再加熱器":
        add(2022 + (fu == 2), "I", "管漏れ点検")
        add(2030 + (fu == 2), "X", "伝熱管更新")
        return m
    if b in ("HCl計", "SOx・NOx・O2・CO計", "ばいじん計", "水銀連続分析計", "NH3計"):
        life = {"HCl計": 10, "SOx・NOx・O2・CO計": 12, "ばいじん計": 11, "水銀連続分析計": 9, "NH3計": 10}[b]
        y = 2014 + life + (fu == 2)
        if rev and b == "水銀連続分析計":
            y = 2023 + (fu == 2)  # 2021 年に後継機種への切替通知
        add(y, "X", "分析計更新")
        add(y + 10, "X", "分析計更新")
        return m
    if b in ("炉出口温度計", "二次燃焼室ガス温度計"):
        for y in range(2017 + (fu == 2), 2035, 3):
            add(y, "X", "熱電対交換")
        return m
    # ---- 灰出し ----
    if b in ("主灰振動コンベヤ", "飛灰搬送コンベヤ", "バグフィルタ飛灰コンベヤ", "ボイラダスト搬出コンベヤ", "ボイラ灰搬送コンベヤ",
             "飛灰集合コンベヤ", "主灰搬出コンベヤ", "飛灰バケットエレベータ"):
        y0 = 2019 + off % 4
        for y in range(y0, 2035, 5):
            add(y, "O", "チェーン・スプロケット交換" if "チェーン" in e["spec"] or "エレベータ" in b else "トラフ・ベルト補修")
        return m
    if b == "飛灰混練機":
        y0 = 2018 + (0 if e["sfx"] == "A" else 2)
        for y in range(y0, 2035, 4):
            add(y, "O", "パドル・ライナ交換")
        return m
    # ---- 盤 ----
    if kind == "盤":
        if "DCS" in b or b in ("自動燃焼制御装置（ACC）", "データロガ・帳票装置"):
            add(2030 if (rev and ver >= 2025) else 2032, "X", "DCS更新")  # 2024.7 部品供給終了予告
            for y in (2020, 2025):
                add(y, "I", "ハード定期点検")
            return m
        if b in ("受電盤（特別高圧）", "高圧配電盤", "系統連系保護盤", "発電機盤"):
            for y in YEARS:
                add(y, "L", "年次点検")
            if b == "高圧配電盤":
                add(2033 if not rev else 2031, "X", "VCB更新")
            return m
        if b == "直流電源装置":
            add(2022 if not rev else 2021, "X", "蓄電池更新")
            add(2032, "X", "蓄電池更新")
            return m
        if b == "無停電電源装置":
            add(2021 + off % 2, "X", "UPS蓄電池更新")
            add(2028 + off % 2, "X", "UPS本体更新")
            return m
        if imp == "A":
            for y in range(2017 + off % 4, 2035, 4):
                add(y, "I", "盤内精密点検")
        return m
    # ---- 共通：受入 ----
    if b == "計量機":
        yrs = list(range(2016, 2035, 2))
        if not rev:
            yrs.remove(2020)  # sp1 の計画で 2020 年度を落とした（計画の欠け）
        for y in yrs:
            add(y, "L", "定期検査（計量法）")
        add(2027, "X", "ロードセル更新")
        return m
    if b == "ごみクレーン":
        for y in YEARS[1:]:
            add(y, "L", "年次自主検査" if y % 2 == 0 else "性能検査・年次自主検査")
        for y in range(2018 + (e["sfx"] == "2号"), 2035, 3):
            add(y, "X", "ワイヤロープ交換")
        return m
    if b in ("ごみクレーン巻上装置", "ごみクレーン走行装置", "ごみクレーン横行装置"):
        y0 = (2023 if e["sfx"] == "1号" else 2025) if rev else (2022 if e["sfx"] == "1号" else 2024)
        add(y0 + (b != "ごみクレーン巻上装置") * 2, "O", "減速機分解整備")
        return m
    if b == "ごみクレーンバケット":
        for y in range(2019 + (e["sfx"] == "2号"), 2035, 4):
            add(y, "O", "爪・シリンダ整備")
        return m
    if b == "灰クレーン":
        for y in YEARS[1:]:
            add(y, "L", "年次自主検査" if y % 2 == 0 else "性能検査・年次自主検査")
        return m
    if b == "ごみ投入扉":
        if int(e["sfx"]) in (2, 3, 4, 5):
            add(2024 + int(e["sfx"]) % 2, "R", "扉板補修")
        return m
    # ---- 余熱利用 ----
    if b == "蒸気タービン":
        add(2018, "L", "定期事業者検査（開放点検）")
        add(2022, "L", "定期事業者検査（開放点検）")
        add(2026, "L", "定期事業者検査")
        add(2026, "O", "分解精密点検")
        add(2030, "L", "定期事業者検査（開放点検）")
        add(2034, "L", "定期事業者検査")
        return m
    if b in ("タービン減速機", "発電機", "タービン調速装置", "主蒸気止め弁油筒"):
        for y in (2018, 2022, 2026, 2030, 2034):
            add(y, "O" if y == 2026 else "I", "分解点検" if y == 2026 else "開放点検")
        if b == "発電機" and rev:
            add(2032, "O", "固定子絶縁診断・ワニス処理")
        return m
    if b == "空冷復水器ファン減速機":
        n = int(e["sfx"])
        if not rev or ver < 2023:  # 2022 の異音（3 号）を受けて 2023 年度版から見直し
            add(2024 + n % 3, "O", "減速機分解整備")
        else:
            add({3: 2023, 1: 2025, 2: 2025, 4: 2027, 5: 2027, 6: 2028}[n], "O", "減速機分解整備")
        return m
    if b == "空冷式復水器":
        for y in range(2019, 2035, 2):
            add(y, "I", "フィン洗浄・目視点検")
        return m
    if b in ("脱気器",):
        for y in YEARS[1:]:
            add(y, "L", "性能検査")
        return m
    if b in ("高圧蒸気だめ", "低圧蒸気だめ", "空気槽"):
        for y in YEARS[1:]:
            add(y, "L", "定期自主検査")
        return m
    if b in ("復水ポンプ", "脱気器給水ポンプ", "機器冷却水ポンプ", "真空ポンプ", "プラント用水揚水ポンプ"):
        y0 = 2019 + (0 if e["sfx"] == "A" else 3) + off % 2
        for y in range(y0, 2035, 6):
            add(y, "O", "ポンプ分解整備")
        return m
    if b == "純水装置":
        for y in range(2018, 2035, 4):
            add(y, "X", "RO膜交換")
        return m
    if b == "機器冷却水冷却塔":
        add(2025, "R", "充填材交換")
        return m
    # ---- 雑設備・電気 ----
    if b in ("計装用空気圧縮機", "雑用空気圧縮機"):
        y0 = 2018 + (0 if e["sfx"] == "A" else 1)
        for y in range(y0, 2035, 4 if b == "雑用空気圧縮機" else 5):
            add(y, "O", "オーバーホール")
        return m
    if b == "主変圧器":
        for y in YEARS:
            add(y, "L", "年次点検")
        for y in range(2017, 2035, 3):
            add(y, "I", "油中ガス分析")
        return m
    if b == "所内変圧器":
        for y in YEARS:
            add(y, "L", "年次点検")
        return m
    if b == "非常用発電機":
        for y in YEARS:
            add(y, "L", "年次点検（負荷試験）")
        add(2020, "O", "機関オーバーホール")
        add(2026 if rev else 2027, "O", "機関オーバーホール")
        return m
    if b == "タービン室天井クレーン":
        for y in YEARS[1:]:
            add(y, "L", "年次自主検査")
        return m
    if b == "煙突外筒":
        add(2024, "I", "外筒劣化診断")
        return m
    # ---- 主要計装（重要度 A、その他） ----
    if kind == "計装" and imp == "A" and b not in ("炉内圧力計",):
        add(2024 + off % 4, "X", "検出器更新")
        return m
    # ---- その他の重要度 A・B 機器：分解整備を 1〜2 回 ----
    if kind == "機器" and imp in ("A", "B"):
        span = 7 if imp == "A" else 9
        y0 = 2019 + off % 5
        for y in range(y0, 2035, span):
            add(y, "O", "分解整備")
        return m
    return m  # 重要度 C の多く・構造物：事後保全または日常点検のみ


# ---------------------------------------------------------------- 実績（計画どおりでなかったもの）
# (機器コード, 年度) : 実績の上書き。None は「計画にあったが実施しなかった」。
def actual(e):
    """2015〜2026 年度上期の実績 {年度: [(属性, 作業名, 備考)]}。"""
    out = {}
    for y in range(2015, 2027):
        src = rule(e, rev=(y >= 2020))
        for (t, w) in src.get(y, []):
            out.setdefault(y, []).append((t, w, ""))
    b, u = e["base"], e["unit"]
    # 2026 年度は上期分のみ実績。下期（1 号炉定修 10–11 月、全停 10–11 月）は施工中・未完。
    if 2026 in out:
        # 1 号炉定修（2026-10-05〜11-10）と全停（2026-10-19〜11-20、余熱利用・電気）は下期：実績未反映
        if u == "1" or (u == "0" and e["sys"] in ("YN", "DE")):
            out[2026] = [(t, w, "施工中") for (t, w, _) in out[2026]]
    # ---- 計画からのずれ（実績） ----
    if b == "誘引通風機" and u == "1":
        out.setdefault(2021, []).append(("I", "振動上昇のため臨時点検（軸受温度・振動測定）", "臨時"))
    if b == "バグフィルタ本体" and u == "2":
        pass
    if b == "空冷復水器ファン減速機" and e["sfx"] == "3":
        out.setdefault(2022, []).append(("I", "異音のため臨時点検（油分析・振動）", "臨時"))
    if b == "ボイラ給水ポンプ" and u == "1" and e["sfx"] == "A":
        # 2017 分解整備後、2023 に計画どおり分解整備（メカシール漏れ再発）
        pass
    if b == "長抜差型スートブロワ" and u == "2" and e["sfx"] == "3":
        # 2022 予定を 2023 に延期（部品納期）
        for k in (2022,):
            out[k] = [x for x in out.get(k, []) if x[1] != "ランス・パッキン交換"]
            if not out[k]:
                del out[k]
        out.setdefault(2023, []).append(("O", "ランス・パッキン交換", "前年度から延期（部品納期）"))
    return out


# ---------------------------------------------------------------- 費用（千円、概算・税抜）
COST = {
    "燃焼火格子更新": 98000, "乾燥火格子更新": 62000, "後燃焼火格子更新": 58000, "火格子部分交換": 14500,
    "ろ布全数交換": 21600, "ろ布抜取試験": 900, "予備層触媒充填": 38000, "触媒交換（1層）": 41000,
    "触媒性能試験": 1200, "インペラ更新": 16500, "インペラ更新（耐食材）": 7800, "定期事業者検査": 4200,
    "定期事業者検査（開放点検）": 18000, "分解精密点検": 46000, "肉厚測定": 1300, "下段管部分更新・プロテクタ交換": 32000,
    "肉盛補修": 18500, "プロテクタ交換": 9600, "耐火物部分補修": 12800, "炉内耐火物点検": 0, "DCS更新": 0,
    "分析計更新": 9200, "ワイヤロープ交換": 2600, "減速機分解整備": 3400, "RO膜交換": 3800,
    "機関オーバーホール": 12500, "VCB更新": 0, "蓄電池更新": 7800, "UPS蓄電池更新": 1900, "UPS本体更新": 6400,
    "年次点検": 0, "ポンプ分解整備": 1800, "ロードセル更新": 5400, "熱電対交換": 180, "検出器更新": 950,
    "伝熱管更新": 21000, "分解点検": 6800, "開放点検": 0, "軸受交換・インペラ点検": 3800,
    "軸受交換・インペラ肉盛補修": 6200, "インペラ摩耗測定": 300, "電動機軸受交換・精密点検": 2400,
    "絶縁診断": 160, "電動機分解整備": 900, "電動機軸受交換": 450,
}
DCS_TOTAL = 480000  # DCS 一式更新（共通計上、機器行には内訳を持たない）


def cost_of(work, e):
    if work in COST:
        c = COST[work]
    else:
        c = 600 + h(e["code"] + work, 9) * 150
    return c


if __name__ == "__main__":
    n_rows = sum(1 for e in ALL if rule(e, True))
    print("計画に作業がある機器", n_rows, "/", len(ALL))


# ---------------------------------------------------------------- 台帳の属性（メーカー・型式・製造番号・設置年月）
MAKERS = {
    "電動機": "【架空】大楠電機", "盤": "【架空】大楠電機", "計装": "【架空】青嵐計器",
}
def maker_of(e):
    b = e["base"]
    if e["kind"] == "電動機":
        return "【架空】大楠電機"
    if "DCS" in b or b in ("自動燃焼制御装置（ACC）", "データロガ・帳票装置", "ITV監視装置"):
        return "【架空】汀システムズ"
    if e["kind"] == "盤":
        return "【架空】大楠電機"
    if e["kind"] == "計装":
        return "【架空】青嵐計器"
    if "ポンプ" in b:
        return "【架空】千鳥ポンプ"
    if "送風機" in b or "ファン" in b or b == "誘引通風機":
        return "【架空】柊送風機" if "減速機" not in b else "【架空】水樹プラント"
    if "クレーン" in b:
        return "【架空】白峰クレーン"
    if b.startswith("タービン") or b in ("蒸気タービン", "ターニング装置", "主蒸気止め弁油筒", "グランド蒸気復水器", "グランド排風機"):
        return "【架空】鳴瀬タービン"
    if b.startswith("発電機"):
        return "【架空】鳴瀬電機"
    if "空気圧縮機" in b or b == "空気除湿装置":
        return "【架空】蒼空コンプレッサ"
    if b in ("主変圧器", "所内変圧器", "非常用発電機"):
        return "【架空】大楠電機"
    if b == "計量機":
        return "【架空】秤屋計測"
    return "【架空】水樹プラント"


def model_of(e):
    k = h(e["code"] + "m", 900) + 100
    pre = {"【架空】大楠電機": "OK", "【架空】千鳥ポンプ": "CP", "【架空】柊送風機": "HF", "【架空】白峰クレーン": "SC",
           "【架空】鳴瀬タービン": "NT", "【架空】鳴瀬電機": "NG", "【架空】蒼空コンプレッサ": "SK", "【架空】青嵐計器": "SR",
           "【架空】汀システムズ": "MS", "【架空】秤屋計測": "HK", "【架空】水樹プラント": "MZ"}[maker_of(e)]
    if e["kind"] == "電動機":
        return f"{pre}-{'H' if e['kw'] >= 200 else 'T'}{e['kw']:g}-{k}"
    return f"{pre}-{e['sys']}{k}"


def serial_of(e):
    return f"{14 + h(e['code'], 1)}{h(e['code'] + 's', 9000) + 1000:04d}"


def install_ym(e):
    """設置年月（台帳の現在値）。更新した機器は更新年月。"""
    base = (2014, 10 + h(e["code"], 3))  # 2014-10〜12 据付
    ym = base
    acts = actual(e)
    for y in sorted(acts):
        for (t, w, note) in acts[y]:
            if t == "X" and note != "施工中" and any(k in w for k in ("更新", "交換")) and "ろ布" not in w and "ワイヤ" not in w \
                    and "熱電対" not in w and "蓄電池" not in w and "RO膜" not in w and "充填" not in w and "触媒" not in w:
                mon = 10 if e["unit"] in ("1", "0") else 6
                ym = (y, mon)
    return ym


# ---------------------------------------------------------------- 定修の暦（実績・2026 は予定を含む）
from datetime import timedelta
TEISHU = {}
for y in range(2015, 2027):
    # 1 号炉：9〜10 月、2 号炉：5〜6 月、全停（共通）：10 月中旬（1 号炉定修の期間中に 2 号炉を止める）
    s1 = {2015: date(2015, 10, 5), 2018: date(2018, 9, 3), 2023: date(2023, 10, 2), 2026: date(2026, 10, 5)}.get(y, date(y, 9, 7 + (y % 3) * 7))
    TEISHU[("1", y)] = (s1, s1 + timedelta(days={2018: 25, 2026: 36}.get(y, 31)))
    s2 = date(y, 5, 8 + (y % 4) * 7) if y > 2015 else date(2015, 11, 16)
    TEISHU[("2", y)] = (s2, s2 + timedelta(days=28))
    s0 = {2018: date(2018, 9, 15), 2023: date(2023, 10, 14), 2026: date(2026, 10, 19)}.get(y, s1 + timedelta(days=12))
    TEISHU[("0", y)] = (s0, s0 + timedelta(days=(8 if y not in (2018, 2022, 2026) else 20 if y != 2026 else 32)))
