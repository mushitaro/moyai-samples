# -*- coding: utf-8 -*-
"""【架空】くぬぎ台環境センター（F03）の「正解の世界」。
機器（タグ番号 1 件＝1 台）、各機器の履歴（実際に行った作業）、契約、代表機器の根拠。
資料はすべてここから書き分ける。乱数の種と物語の今日は固定。
"""
import random
from collections import OrderedDict

SEED = 20261113
TODAY = (2026, 11, 13)
FAC = "【架空】くぬぎ台環境センター"
CITY = "【架空】くぬぎ台市"

# 年度 ------------------------------------------------------------------
def wareki(fy, short=False):
    if fy >= 2019:
        n = fy - 2018
        return ("R%d" % n) if short else ("令和%s年度" % ("元" if n == 1 else n))
    n = fy - 1988
    return ("H%d" % n) if short else ("平成%d年度" % n)

def wa_y(fy):  # 年度の数字だけ
    return fy - 2018 if fy >= 2019 else fy - 1988

# 改良の年度 -------------------------------------------------------------
DXN_FY = {1: 2000, 2: 2001, 3: 2002}      # DXN 恒久対策（EP→BF、減温塔・触媒追設）
KIKAN_FY = {0: 2011, 1: 2012, 2: 2013, 3: 2014}  # 基幹的設備改良（0=共通・中央）
AM_FY = 2003                               # 灰溶融炉追設
AM_STOP_FY = 2013                          # 灰溶融炉休止

MAKERS = {
    "SH": "汐見重工環境", "YZ": "杠葉ポンプ製作所", "TK": "鴇田送風機製作所",
    "SW": "椹木電機", "YR": "楪計装", "SN": "篠目クレーン", "TG": "柘植バルブ工業",
    "AH": "葦原タービン", "HI": "柊分析計器", "IZ": "伊沢原動機", "ON": "恩田精機",
    "KR": "黒鉄コンベヤ", "NB": "野火止産業",
}

# 系統の名前（設備 → 系統）
SYS = OrderedDict([
    ("RF", ("受入供給設備", "ごみ受入・供給")),
    ("IN", ("燃焼設備", "焼却炉・ストーカ")),
    ("BL", ("燃焼ガス冷却設備", "ボイラ・給水")),
    ("GT", ("排ガス処理設備", "減温塔・BF・触媒・洗煙")),
    ("FD", ("通風設備", "送風機・煙道")),
    ("AS", ("灰出し設備", "主灰・飛灰")),
    ("FA", ("灰出し設備", "飛灰処理")),
    ("ST", ("余熱利用設備", "蒸気タービン発電・復水")),
    ("CH", ("薬品供給設備", "薬品貯留")),
    ("WW", ("排水処理設備", "洗煙排水・プラント排水")),
    ("UT", ("用水・空気設備", "用水・冷却水・圧縮空気")),
    ("EL", ("電気設備", "受変電・配電・動力")),
    ("IC", ("計装制御設備", "DCS・監視")),
    ("MS", ("雑設備", "煙突・荷役・消火")),
    ("AM", ("灰溶融設備（休止）", "電気抵抗式灰溶融炉")),
])

# 機器の種類（周期と費用の目安。千円）
#   oh: 分解整備・部分取替の周期（年）、rep: 更新の目安（年）、insp: 特別な点検・測定の周期
CLS = {
    "hopper":  dict(oh=None, rep=32, coh=0,    crep=18000, imp="B", rep_name="更新", ins=None),
    "feeder":  dict(oh=5,  rep=22, coh=2800,  crep=26000, imp="A"),
    "stoker":  dict(oh=3,  rep=25, coh=9500,  crep=98000, imp="A", oh_name="火格子部分交換"),
    "hyd":     dict(oh=6,  rep=20, coh=1800,  crep=14000, imp="A"),
    "pump_s":  dict(oh=5,  rep=18, coh=450,   crep=1600,  imp="B"),
    "pump_m":  dict(oh=4,  rep=18, coh=1200,  crep=5200,  imp="B"),
    "fan_big": dict(oh=4,  rep=22, coh=4200,  crep=38000, imp="A", ins=1, ins_name="振動診断"),
    "fan":     dict(oh=5,  rep=24, coh=1500,  crep=9000,  imp="B"),
    "furnace": dict(oh=2,  rep=None, coh=6800, crep=0,     imp="A", oh_name="耐火物部分補修"),
    "burner":  dict(oh=5,  rep=20, coh=700,   crep=5500,  imp="B"),
    "boiler":  dict(oh=None, rep=None, coh=0, crep=0,     imp="A", ins=1, ins_name="肉厚測定"),
    "sh":      dict(oh=None, rep=14, coh=0,   crep=42000, imp="A", ins=1, ins_name="肉厚測定"),
    "eco":     dict(oh=None, rep=24, coh=0,   crep=30000, imp="A", ins=2, ins_name="肉厚測定"),
    "sb":      dict(oh=3,  rep=15, coh=550,   crep=3800,  imp="B"),
    "sv":      dict(oh=2,  rep=20, coh=300,   crep=1500,  imp="A", oh_name="吹出し調整"),
    "mov":     dict(oh=6,  rep=20, coh=350,   crep=1900,  imp="B"),
    "cv":      dict(oh=4,  rep=16, coh=280,   crep=1400,  imp="B"),
    "damper":  dict(oh=6,  rep=22, coh=600,   crep=3500,  imp="C"),
    "inst":    dict(oh=None, rep=12, coh=0,   crep=650,   imp="C"),
    "inst_a":  dict(oh=None, rep=13, coh=0,   crep=900,   imp="B"),
    "anal":    dict(oh=None, rep=11, coh=0,   crep=16000, imp="A"),
    "panel":   dict(oh=None, rep=22, coh=0,   crep=6500,  imp="B", ins=6, ins_name="精密点検"),
    "dcs":     dict(oh=None, rep=15, coh=0,   crep=35000, imp="A"),
    "conv":    dict(oh=4,  rep=18, coh=900,   crep=7000,  imp="B"),
    "bf":      dict(oh=7,  rep=32, coh=16000, crep=160000, imp="A", oh_name="ろ布全数交換"),
    "cat":     dict(oh=10, rep=None, coh=48000, crep=0,   imp="A", oh_name="触媒交換", ins=3, ins_name="活性試験"),
    "scrub":   dict(oh=12, rep=None, coh=36000, crep=0,   imp="A", oh_name="ライニング更新", ins=2, ins_name="内面点検"),
    "tower":   dict(oh=8,  rep=None, coh=5200, crep=0,    imp="B", oh_name="ノズル・内面補修"),
    "tank":    dict(oh=None, rep=32, coh=0,   crep=4000,  imp="C", ins=5, ins_name="内部点検"),
    "hx":      dict(oh=6,  rep=25, coh=900,   crep=6000,  imp="B", oh_name="開放清掃"),
    "camera":  dict(oh=None, rep=10, coh=0,   crep=1200,  imp="C"),
    "crane":   dict(oh=3,  rep=22, coh=3200,  crep=72000, imp="A", oh_name="ワイヤロープ交換・整備"),
    "hoist":   dict(oh=None, rep=25, coh=0,   crep=2200,  imp="C"),
    "turbine": dict(oh=4,  rep=None, coh=38000, crep=0,   imp="A", oh_name="開放点検"),
    "gen":     dict(oh=8,  rep=None, coh=9000, crep=0,    imp="A", oh_name="精密点検"),
    "acc":     dict(oh=6,  rep=26, coh=4500,  crep=60000, imp="A"),
    "tr":      dict(oh=None, rep=30, coh=0,   crep=12000, imp="A", ins=6, ins_name="絶縁診断"),
    "egen":    dict(oh=8,  rep=30, coh=3500,  crep=42000, imp="A", oh_name="オーバーホール"),
    "comp":    dict(oh=3,  rep=16, coh=800,   crep=5500,  imp="B"),
    "struct":  dict(oh=None, rep=None, coh=0, crep=0,     imp="B", ins=5, ins_name="外観・肉厚点検"),
    "misc":    dict(oh=8,  rep=25, coh=500,   crep=3000,  imp="C"),
    "scale":   dict(oh=None, rep=20, coh=0,   crep=9000,  imp="B"),
    "door":    dict(oh=6,  rep=25, coh=500,   crep=4200,  imp="C"),
}

# ------------------------------------------------------------------------
# 炉別 165 台（1 炉分）。sys|名称|cls|仕様|kW|メーカー|由来|区分
# 由来: 1990=竣工時、dxn=DXN 対策で追設、kikan=基幹改良で更新、kikan+=基幹改良で追設
FURNACE_TXT = """
IN|ごみ投入ホッパ|hopper|容量 40m³、SS400 内張 SUS|0|SH|1990|
IN|ホッパブリッジ解除装置|hyd|油圧押出式|0|SH|1990|
IN|給じん装置|feeder|プッシャ式 油圧駆動 4.2t/h|0|SH|kikan|
IN|乾燥ストーカ|stoker|並行揺動式 幅 2.4m×長 3.2m|0|SH|1990|
IN|燃焼ストーカ|stoker|並行揺動式 幅 2.4m×長 4.0m|0|SH|1990|
IN|後燃焼ストーカ|stoker|並行揺動式 幅 2.4m×長 2.4m|0|SH|1990|
IN|ストーカ駆動油圧ユニット|hyd|14MPa、油タンク 1.2m³|0|SH|kikan|
IN|ストーカ油圧ポンプ A|pump_s|可変容量形 140L/min×14MPa|37|YZ|kikan|常用
IN|ストーカ油圧ポンプ B|pump_s|可変容量形 140L/min×14MPa|37|YZ|kikan|予備
IN|作動油冷却器|hx|多管式 水冷|0|SH|kikan|
IN|焼却炉本体|furnace|耐火物 SK34・キャスタブル、空冷壁|0|SH|1990|
IN|助燃バーナ|burner|灯油 120L/h|0|ON|1990|
IN|再燃バーナ|burner|灯油 80L/h|0|ON|1990|
IN|バーナ燃料油ポンプ|pump_s|ギヤポンプ 6L/min|1.5|YZ|1990|
IN|炉内監視カメラ|camera|水冷ケーシング カラー|0|SW|kikan|
IN|落じんコンベヤ|conv|スクリュー式 φ300|3.7|KR|1990|
IN|炉側壁冷却ファン|fan|3,000m³N/h×3kPa|7.5|TK|1990|
IN|乾燥段 燃焼空気調節ダンパ|damper|電動 バタフライ|0.4|TG|kikan|
IN|燃焼段 燃焼空気調節ダンパ|damper|電動 バタフライ|0.4|TG|kikan|
IN|後燃焼段 燃焼空気調節ダンパ|damper|電動 バタフライ|0.4|TG|kikan|
IN|二次空気調節ダンパ|damper|電動 バタフライ|0.4|TG|kikan|
IN|炉出口温度計 A|inst_a|熱電対 K 0〜1,200℃|0|YR|kikan|
IN|炉出口温度計 B|inst_a|熱電対 K 0〜1,200℃|0|YR|kikan|
IN|炉内圧力計|inst|差圧伝送器 ±500Pa|0|YR|kikan|
IN|乾燥段 火格子温度計|inst|熱電対 K|0|YR|1990|
IN|燃焼段 火格子温度計|inst|熱電対 K|0|YR|1990|
IN|後燃焼段 火格子温度計|inst|熱電対 K|0|YR|1990|
IN|ホッパレベル計|inst|マイクロ波式|0|YR|kikan|
IN|自動燃焼制御装置|dcs|ACC（DCS 内機能＋専用演算器）|0|SW|kikan|
IN|ストーカ油圧圧力計|inst|圧力伝送器 0〜25MPa|0|YR|kikan|
IN|落じんシュートダンパ|damper|電動 フラップ|0.2|TG|1990|
IN|燃焼室温度計|inst_a|熱電対 R|0|YR|kikan|
IN|二次燃焼室温度計|inst_a|熱電対 R|0|YR|kikan|
BL|ボイラ本体|boiler|自然循環水管式 10.5t/h、2.0MPa・250℃|0|SH|1990|
BL|過熱器|sh|1次・2次 STB410/SUS 吊下げ式|0|SH|1990|
BL|エコノマイザ|eco|裸管式 STB340|0|SH|1990|
BL|スートブロワ SB-1|sb|長抜差し式 蒸気|0.4|SH|1990|
BL|スートブロワ SB-2|sb|長抜差し式 蒸気|0.4|SH|1990|
BL|スートブロワ SB-3|sb|長抜差し式 蒸気|0.4|SH|1990|
BL|スートブロワ SB-4|sb|長抜差し式 蒸気|0.4|SH|1990|
BL|スートブロワ SB-5|sb|定置旋回式 蒸気|0.2|SH|1990|
BL|スートブロワ SB-6|sb|定置旋回式 蒸気|0.2|SH|1990|
BL|スートブロワ SB-7|sb|定置旋回式 蒸気|0.2|SH|1990|
BL|スートブロワ SB-8|sb|定置旋回式 蒸気|0.2|SH|1990|
BL|ドラム安全弁 A|sv|吹出し 2.2MPa|0|TG|1990|
BL|ドラム安全弁 B|sv|吹出し 2.25MPa|0|TG|1990|
BL|過熱器安全弁|sv|吹出し 2.1MPa|0|TG|1990|
BL|主蒸気止弁|mov|電動 80A 2.0MPa|0.75|TG|kikan|
BL|給水調節弁|cv|空気作動 50A|0|TG|kikan|
BL|連続ブロー調節弁|cv|空気作動 20A|0|TG|kikan|
BL|間欠ブロー弁|mov|電動 40A|0.4|TG|kikan|
BL|過熱低減器|hx|スプレー式|0|SH|1990|
BL|過熱低減調節弁|cv|空気作動 25A|0|TG|kikan|
BL|ボイラ給水ポンプ A|pump_m|多段渦巻 13m³/h×3.2MPa|30|YZ|kikan|常用
BL|ボイラ給水ポンプ B|pump_m|多段渦巻 13m³/h×3.2MPa|30|YZ|kikan|予備
BL|ドラム水位計（差圧式）|inst_a|差圧伝送器|0|YR|kikan|
BL|ドラム水位計（電極式）|inst_a|電極式 警報用|0|YR|1990|
BL|ドラム水面計 A|inst|二色水面計|0|TG|1990|
BL|ドラム水面計 B|inst|二色水面計|0|TG|1990|
BL|主蒸気流量計|inst_a|オリフィス＋差圧伝送器|0|YR|kikan|
BL|主蒸気圧力計|inst|圧力伝送器 0〜3MPa|0|YR|kikan|
BL|主蒸気温度計|inst|測温抵抗体|0|YR|kikan|
BL|ドラム圧力計|inst|圧力伝送器 0〜3MPa|0|YR|kikan|
BL|給水流量計|inst|差圧式|0|YR|kikan|
BL|ボイラ水質計|inst_a|pH・導電率|0|HI|kikan|
BL|清缶剤注入ポンプ A|pump_s|ダイヤフラム 2L/h|0.2|ON|1990|常用
BL|清缶剤注入ポンプ B|pump_s|ダイヤフラム 2L/h|0.2|ON|1990|予備
BL|清缶剤タンク|tank|0.5m³ PE|0|ON|1990|
BL|ボイラダストコンベヤ|conv|スクリュー式 φ250|2.2|KR|1990|
BL|ボイラドレン弁|mov|電動 25A|0.2|TG|kikan|
BL|過熱器出口弁|mov|電動 80A|0.75|TG|kikan|
BL|起動用ベント弁|mov|電動 25A|0.2|TG|kikan|
BL|スートブロワ蒸気元弁|mov|電動 50A|0.4|TG|kikan|
BL|空気予熱器用減圧弁|cv|自力式 2.0→0.8MPa|0|TG|1990|
BL|ブロータンク|tank|0.8m³|0|SH|1990|
BL|ボイラ出口温度計|inst|熱電対 K|0|YR|kikan|
BL|エコノマイザ出口温度計|inst|熱電対 K|0|YR|kikan|
BL|給水入口電動弁|mov|電動 50A|0.4|TG|kikan|
BL|連続ブロー流量計|inst|面積式|0|YR|1990|
GT|減温塔|tower|水噴霧式 φ3.0m×H14m|0|SH|dxn|
GT|減温塔噴霧水ポンプ A|pump_m|渦巻 2.5m³/h×1.5MPa|5.5|YZ|dxn|常用
GT|減温塔噴霧水ポンプ B|pump_m|渦巻 2.5m³/h×1.5MPa|5.5|YZ|dxn|予備
GT|噴霧水調節弁|cv|空気作動 25A|0|TG|dxn|
GT|噴霧空気調節弁|cv|空気作動 40A（二流体）|0|TG|dxn|
GT|噴霧水流量計|inst|電磁流量計|0|YR|dxn|
GT|減温塔出口温度計|inst_a|熱電対 K|0|YR|dxn|
GT|減温塔ダスト排出装置|conv|ロータリバルブ φ300|0.75|KR|dxn|
GT|減温塔ダストコンベヤ|conv|スクリュー式|2.2|KR|dxn|
GT|消石灰定量供給装置|feeder|テーブルフィーダ 30kg/h|0.4|ON|dxn|
GT|活性炭定量供給装置|feeder|スクリューフィーダ 3kg/h|0.2|ON|dxn|
GT|薬剤吹込みブロワ A|fan|ルーツ 3m³/min|7.5|TK|dxn|常用
GT|薬剤吹込みブロワ B|fan|ルーツ 3m³/min|7.5|TK|dxn|予備
GT|バグフィルタ|bf|パルスジェット式 6 室 ろ布 288 本（φ150×6m）|0|SH|dxn|
GT|バグフィルタ制御盤|panel|PLC パルス制御|0|SW|dxn|
GT|バグフィルタ差圧計|inst_a|差圧伝送器 0〜5kPa|0|YR|dxn|
GT|バグフィルタ入口温度計|inst_a|熱電対 K|0|YR|dxn|
GT|BF ダスト排出装置|conv|ロータリバルブ φ300×6|0.75|KR|dxn|
GT|BF ダストコンベヤ|conv|スクリュー式|3.7|KR|dxn|
GT|BF 入口ダンパ|damper|電動 ルーバ|0.75|TG|dxn|
GT|BF 出口ダンパ|damper|電動 ルーバ|0.75|TG|dxn|
GT|BF バイパスダンパ|damper|電動 ルーバ|0.75|TG|dxn|
GT|BF ホッパヒータ|misc|電気式 6kW×6|0|SW|dxn|
GT|BF プレコート装置|feeder|定量フィーダ|0.4|ON|dxn|
GT|ガス再加熱器|hx|蒸気式 フィンチューブ|0|SH|dxn|
GT|再加熱器蒸気調節弁|cv|空気作動 50A|0|TG|dxn|
GT|触媒反応塔|cat|ハニカム触媒 2 層 210℃|0|SH|dxn|
GT|アンモニア水注入ポンプ A|pump_s|ダイヤフラム 10L/h|0.2|ON|dxn|常用
GT|アンモニア水注入ポンプ B|pump_s|ダイヤフラム 10L/h|0.2|ON|dxn|予備
GT|アンモニア注入調節弁|cv|空気作動 15A|0|TG|dxn|
GT|触媒入口温度計|inst|熱電対 K|0|YR|dxn|
GT|触媒バイパスダンパ|damper|電動 ルーバ|0.75|TG|dxn|
GT|触媒出口温度計|inst|熱電対 K|0|YR|dxn|
GT|BF 逆洗用空気槽|tank|1.5m³ 0.6MPa|0|SH|dxn|
GT|湿式洗煙塔|scrub|充填塔 FRP ライニング φ2.8m×H12m|0|SH|1990|
GT|洗煙循環ポンプ A|pump_m|耐食渦巻 60m³/h×0.3MPa|15|YZ|1990|常用
GT|洗煙循環ポンプ B|pump_m|耐食渦巻 60m³/h×0.3MPa|15|YZ|1990|予備
GT|苛性ソーダ注入ポンプ|pump_s|ダイヤフラム 40L/h|0.4|ON|1990|
GT|苛性ソーダ注入調節弁|cv|空気作動 15A|0|TG|kikan|
GT|洗煙塔 pH 計|inst_a|ガラス電極式|0|HI|kikan|
GT|洗煙塔液位計|inst|差圧式|0|YR|kikan|
GT|洗煙排水抜出ポンプ|pump_s|耐食渦巻 3m³/h|2.2|YZ|1990|
GT|洗煙塔出口温度計|inst|測温抵抗体|0|YR|kikan|
GT|循環液冷却器|hx|プレート式 チタン|0|SH|1990|
GT|洗煙塔補給水調節弁|cv|空気作動 25A|0|TG|kikan|
GT|洗煙塔差圧計|inst|差圧伝送器|0|YR|kikan|
GT|洗煙塔バイパスダンパ|damper|電動 ルーバ|0.75|TG|1990|
GT|排ガス分析計|anal|HCl・SOx・NOx・O2・CO 連続|0|HI|kikan|
GT|ばいじん計|anal|光散乱式|0|HI|kikan|
GT|排ガス流量計|inst_a|ピトー管式|0|YR|kikan|
GT|HCl 計（イオン電極）|anal|イオン電極式 予備計|0|HI|1990|
FD|押込送風機|fan|20,000m³N/h×4.5kPa|45|TK|kikan|
FD|二次送風機|fan|8,000m³N/h×4.0kPa|18.5|TK|kikan|
FD|蒸気式空気予熱器|hx|ベアチューブ 0.8MPa 蒸気|0|SH|1990|
FD|押込送風機入口ダンパ|damper|電動 ベーン|0.4|TG|kikan|
FD|二次送風機ダンパ|damper|電動 ベーン|0.4|TG|kikan|
FD|誘引通風機|fan_big|30,000m³N/h×7.5kPa 両吸込 ターボ|220|TK|kikan|
FD|誘引通風機インバータ盤|panel|6.6kV 高圧インバータ 250kW|0|SW|kikan|
FD|誘引通風機入口ダンパ|damper|電動 ベーン|0.75|TG|kikan|
FD|誘引通風機出口ダンパ|damper|電動 ルーバ|0.75|TG|1990|
FD|煙道|struct|炉別 SS400 保温付|0|SH|1990|
FD|煙突内筒|struct|SUS 内径 1.0m（外筒共通）|0|SH|1990|
FD|一次空気流量計|inst|ベンチュリ＋差圧|0|YR|kikan|
FD|二次空気流量計|inst|ベンチュリ＋差圧|0|YR|kikan|
FD|誘引通風機振動計|inst_a|速度形 2 点|0|YR|kikan|
FD|誘引通風機軸受温度計|inst|測温抵抗体 2 点|0|YR|kikan|
FD|押込送風機振動計|inst|速度形|0|YR|kikan|
FD|燃焼用 O2 計|anal|ジルコニア式（ボイラ出口）|0|HI|kikan|
FD|誘引通風機入口圧力計|inst|圧力伝送器|0|YR|kikan|
FD|減温塔入口温度計|inst|熱電対 K|0|YR|dxn|
AS|主灰押出装置|conv|水封式プッシャ|5.5|SH|1990|
AS|主灰振動コンベヤ|conv|振動式 幅 600|3.7|KR|1990|
AS|飛灰集合コンベヤ|conv|スクリュー式|3.7|KR|dxn|
AS|灰押出装置水位計|inst|フロート式|0|YR|1990|
UT|機器冷却水電動弁|mov|電動 80A|0.4|TG|kikan|
UT|機器冷却水流量計|inst|電磁流量計|0|YR|kikan|
UT|計装空気レシーバ（炉別）|tank|0.5m³|0|SH|1990|
EL|炉別動力制御盤 1|panel|コントロールセンタ 440V 12 面|0|SW|kikan|
EL|炉別動力制御盤 2|panel|コントロールセンタ 440V 10 面|0|SW|kikan|
EL|ストーカ現場操作盤|panel|屋内自立形|0|SW|kikan|
EL|バグフィルタ現場盤|panel|屋内自立形|0|SW|dxn|
EL|洗煙現場盤|panel|屋内自立形|0|SW|kikan|
EL|炉別計装電源盤|panel|AC100V 分電|0|SW|kikan|
IC|DCS 炉別コントローラ|dcs|二重化コントローラ|0|SW|kikan|
IC|投入ホッパ監視 ITV|camera|固定カメラ|0|SW|kikan|
IC|主灰コンベヤ監視 ITV|camera|固定カメラ|0|SW|kikan|
"""

COMMON_TXT = """
RF|計量機（入口）|scale|ロードセル式 30t|0|NB|kikan|
RF|計量機（出口）|scale|ロードセル式 30t|0|NB|kikan|
RF|計量データ処理装置|dcs|料金計算・帳票|0|NB|kikan|
RF|投入扉 No.1|door|観音開き 油圧|0|SH|1990|
RF|投入扉 No.2|door|観音開き 油圧|0|SH|1990|
RF|投入扉 No.3|door|観音開き 油圧|0|SH|1990|
RF|投入扉 No.4|door|観音開き 油圧|0|SH|1990|
RF|投入扉 No.5|door|観音開き 油圧|0|SH|1990|
RF|投入扉 No.6|door|観音開き 油圧|0|SH|1990|
RF|投入扉油圧ユニット|hyd|7MPa|11|SH|1990|
RF|ごみピット|struct|7,200m³ RC|0|SH|1990|
RF|ごみクレーン 1号|crane|天井走行 バケット 5.0m³ 吊上げ 8.5t|0|SN|kikan|常用
RF|ごみクレーン 2号|crane|天井走行 バケット 5.0m³ 吊上げ 8.5t|0|SN|kikan|予備
RF|ごみクレーン制御盤|panel|インバータ 自動運転|0|SW|kikan|
RF|ごみクレーン予備バケット|misc|フォーク式 5.0m³|0|SN|kikan|
RF|粗大ごみ切断機|hyd|竪型せん断式 300kN|22|NB|1990|
RF|ピット放水銃 A|misc|遠隔操作 400L/min|0|NB|2009|
RF|ピット放水銃 B|misc|遠隔操作 400L/min|0|NB|2009|
RF|ピット脱臭装置|hx|活性炭吸着式|0|SH|1990|
RF|脱臭ファン|fan|20,000m³/h|22|TK|1990|
RF|エアカーテン A|fan|プラットホーム出入口|3.7|TK|1990|
RF|エアカーテン B|fan|プラットホーム出入口|3.7|TK|1990|
RF|ごみピット汚水ポンプ A|pump_s|水中 0.5m³/min|3.7|YZ|1990|常用
RF|ごみピット汚水ポンプ B|pump_s|水中 0.5m³/min|3.7|YZ|1990|予備
RF|ごみピット汚水貯留槽|tank|20m³ RC|0|SH|1990|
RF|汚水炉内噴霧ポンプ A|pump_s|渦巻 1m³/h|1.5|YZ|1990|常用
RF|汚水炉内噴霧ポンプ B|pump_s|渦巻 1m³/h|1.5|YZ|1990|予備
ST|蒸気タービン|turbine|単気筒 背圧抽気復水 3,000kW|0|AH|1990|
ST|減速機|gen|平行軸 歯車 6,500/1,500rpm|0|AH|1990|
ST|発電機|gen|三相同期 3,750kVA 6.6kV|0|SW|1990|
ST|主油ポンプ|pump_s|軸直結 歯車|0|AH|1990|
ST|補助油ポンプ|pump_s|電動 歯車|7.5|AH|1990|
ST|非常用油ポンプ|pump_s|直流電動|3.7|AH|1990|
ST|潤滑油冷却器 A|hx|多管式|0|AH|1990|
ST|潤滑油冷却器 B|hx|多管式|0|AH|1990|
ST|調速装置|cv|電子ガバナ|0|AH|kikan|
ST|主塞止弁|mov|油圧 150A|0|AH|1990|
ST|タービンバイパス減圧弁|cv|空気作動 150A|0|TG|1990|
ST|空冷復水器|acc|A フレーム 36t/h|0|SH|1990|
ST|空冷復水器ファン No.1|fan|軸流 φ4.0m|30|TK|kikan|
ST|空冷復水器ファン No.2|fan|軸流 φ4.0m|30|TK|kikan|
ST|空冷復水器ファン No.3|fan|軸流 φ4.0m|30|TK|kikan|
ST|空冷復水器ファン No.4|fan|軸流 φ4.0m|30|TK|kikan|
ST|空冷復水器ファン No.5|fan|軸流 φ4.0m|30|TK|kikan|
ST|空冷復水器ファン No.6|fan|軸流 φ4.0m|30|TK|kikan|
ST|復水タンク|tank|8m³|0|SH|1990|
ST|復水ポンプ A|pump_m|渦巻 40m³/h|11|YZ|1990|常用
ST|復水ポンプ B|pump_m|渦巻 40m³/h|11|YZ|1990|予備
ST|抽気エゼクタ|hx|二段 蒸気式|0|AH|1990|
ST|グランド蒸気復水器|hx|多管式|0|AH|1990|
ST|グランド排風機|fan|小型ターボ|1.5|AH|1990|
ST|ターニング装置|misc|電動|3.7|AH|1990|
ST|潤滑油清浄機|misc|遠心式|2.2|AH|2008|
ST|高圧蒸気だめ|tank|φ600×4m|0|SH|1990|
ST|高圧蒸気だめ安全弁|sv|吹出し 2.1MPa|0|TG|1990|
ST|低圧蒸気だめ|tank|φ500×3m|0|SH|1990|
ST|低圧蒸気減圧弁|cv|空気作動 100A|0|TG|1990|
ST|脱気器|tank|加圧スプレー式 40t/h 0.2MPa|0|SH|1990|
ST|脱気器水位調節弁|cv|空気作動 80A|0|TG|kikan|
ST|純水装置|hx|2 床 3 塔 5m³/h|0|SH|1990|
ST|純水タンク|tank|30m³ SUS|0|SH|1990|
ST|純水移送ポンプ A|pump_s|渦巻 10m³/h|3.7|YZ|1990|常用
ST|純水移送ポンプ B|pump_s|渦巻 10m³/h|3.7|YZ|1990|予備
ST|脱酸剤注入装置|pump_s|ダイヤフラム|0.2|ON|1990|
ST|タービン監視盤|panel|振動・伸び差・軸位置|0|AH|kikan|
ST|発電機盤|panel|AVR・同期|0|SW|kikan|
ST|タービン振動監視装置|inst_a|渦電流式 4 点|0|AH|kikan|
ST|タービン入口蒸気流量計|inst_a|オリフィス＋差圧伝送器|0|YR|kikan|
ST|場内給湯熱交換器|hx|プレート式|0|SH|1990|
ST|場内温水ポンプ A|pump_s|渦巻|2.2|YZ|1990|常用
ST|場内温水ポンプ B|pump_s|渦巻|2.2|YZ|1990|予備
FA|飛灰搬送コンベヤ A|conv|チェーンコンベヤ|5.5|KR|dxn|
FA|飛灰搬送コンベヤ B|conv|チェーンコンベヤ|5.5|KR|dxn|
FA|飛灰貯留サイロ|tank|60m³|0|SH|dxn|
FA|飛灰サイロ集じん機|misc|パルス式|1.5|SH|dxn|
FA|飛灰定量供給機|feeder|ロータリフィーダ|1.5|ON|dxn|
FA|混練機 A|conv|二軸パドル 3t/h|22|NB|dxn|常用
FA|混練機 B|conv|二軸パドル 3t/h|22|NB|dxn|予備
FA|キレート剤タンク|tank|10m³|0|ON|dxn|
FA|キレート注入ポンプ A|pump_s|ダイヤフラム|0.4|ON|dxn|常用
FA|キレート注入ポンプ B|pump_s|ダイヤフラム|0.4|ON|dxn|予備
FA|処理灰コンベヤ|conv|ベルトコンベヤ|3.7|KR|dxn|
FA|灰ピット|struct|主灰・処理灰 600m³|0|SH|1990|
FA|灰クレーン|crane|天井走行 グラブ 1.2m³ 吊上げ 2.8t|0|SN|2006|
FA|主灰集合コンベヤ|conv|ベルトコンベヤ 3 炉集合|7.5|KR|1990|
FA|主灰磁選機|misc|吊下げ磁選|3.7|NB|2008|
CH|消石灰サイロ|tank|50m³|0|SH|dxn|
CH|活性炭サイロ|tank|10m³|0|SH|dxn|
CH|消石灰サイロ集じん機|misc|パルス式|1.5|SH|dxn|
CH|活性炭サイロ集じん機|misc|パルス式|0.75|SH|dxn|
CH|苛性ソーダタンク|tank|25m³ SS+ライニング|0|SH|1990|
CH|苛性ソーダ移送ポンプ A|pump_s|マグネット 3m³/h|1.5|ON|1990|常用
CH|苛性ソーダ移送ポンプ B|pump_s|マグネット 3m³/h|1.5|ON|1990|予備
CH|アンモニア水タンク|tank|15m³ SUS|0|SH|dxn|
CH|燃料油タンク|tank|地下 灯油 20kL|0|SH|1990|
CH|燃料油移送ポンプ A|pump_s|ギヤ 60L/min|1.5|YZ|1990|常用
CH|燃料油移送ポンプ B|pump_s|ギヤ 60L/min|1.5|YZ|1990|予備
WW|洗煙排水受槽|tank|30m³ RC ライニング|0|SH|1990|
WW|洗煙排水移送ポンプ A|pump_s|耐食渦巻 4m³/h|2.2|YZ|1990|常用
WW|洗煙排水移送ポンプ B|pump_s|耐食渦巻 4m³/h|2.2|YZ|1990|予備
WW|第一反応槽攪拌機|misc|竪型 2 段|2.2|NB|1990|
WW|第二反応槽攪拌機|misc|竪型 2 段|2.2|NB|1990|
WW|凝集沈殿槽|tank|φ5.0m|0|SH|1990|
WW|汚泥掻寄機|misc|中央駆動|0.4|NB|1990|
WW|汚泥引抜ポンプ A|pump_s|一軸ねじ|1.5|YZ|1990|常用
WW|汚泥引抜ポンプ B|pump_s|一軸ねじ|1.5|YZ|1990|予備
WW|汚泥脱水機|misc|スクリュープレス|3.7|NB|2010|
WW|キレート樹脂塔 A|tank|1.5m³|0|SH|1990|
WW|キレート樹脂塔 B|tank|1.5m³|0|SH|1990|
WW|砂ろ過器|tank|圧力式 φ1.2m|0|SH|1990|
WW|処理水槽|tank|20m³|0|SH|1990|
WW|放流ポンプ A|pump_s|渦巻 4m³/h|2.2|YZ|1990|常用
WW|放流ポンプ B|pump_s|渦巻 4m³/h|2.2|YZ|1990|予備
WW|反応槽 pH 計|inst_a|ガラス電極式|0|HI|kikan|
WW|放流水 pH 計|inst_a|ガラス電極式|0|HI|kikan|
WW|凝集剤注入ポンプ|pump_s|ダイヤフラム|0.2|ON|1990|
WW|苛性ソーダ注入ポンプ（排水）|pump_s|ダイヤフラム|0.2|ON|1990|
WW|硫酸注入ポンプ|pump_s|ダイヤフラム|0.2|ON|1990|
WW|プラント排水ピットポンプ A|pump_s|水中|2.2|YZ|1990|常用
WW|プラント排水ピットポンプ B|pump_s|水中|2.2|YZ|1990|予備
WW|放流水流量計|inst|電磁流量計|0|YR|kikan|
WW|排水処理現場盤|panel|屋内自立形|0|SW|kikan|
UT|上水受水槽|tank|60m³ FRP|0|SH|1990|
UT|プラント用水ポンプ A|pump_s|渦巻 20m³/h|7.5|YZ|1990|常用
UT|プラント用水ポンプ B|pump_s|渦巻 20m³/h|7.5|YZ|1990|予備
UT|井水ポンプ|pump_s|深井戸水中|11|YZ|1990|
UT|冷却塔 A|hx|開放式 800kW|7.5|NB|kikan|
UT|冷却塔 B|hx|開放式 800kW|7.5|NB|kikan|
UT|機器冷却水ポンプ A|pump_m|渦巻 120m³/h|22|YZ|kikan|常用
UT|機器冷却水ポンプ B|pump_m|渦巻 120m³/h|22|YZ|kikan|常用
UT|機器冷却水ポンプ C|pump_m|渦巻 120m³/h|22|YZ|kikan|予備
UT|冷却水薬注装置|pump_s|ダイヤフラム|0.2|ON|kikan|
UT|冷却水ろ過器|misc|オートストレーナ|0.4|NB|kikan|
UT|雑用空気圧縮機 A|comp|スクリュー 5.5m³/min|37|ON|kikan|常用
UT|雑用空気圧縮機 B|comp|スクリュー 5.5m³/min|37|ON|kikan|予備
UT|計装用空気圧縮機 A|comp|オイルフリー 3.0m³/min|22|ON|kikan|常用
UT|計装用空気圧縮機 B|comp|オイルフリー 3.0m³/min|22|ON|kikan|予備
UT|除湿装置 A|misc|冷凍式|1.5|ON|kikan|常用
UT|除湿装置 B|misc|冷凍式|1.5|ON|kikan|予備
UT|雑用空気槽|tank|3m³ 0.7MPa|0|SH|1990|
UT|計装用空気槽|tank|3m³ 0.7MPa|0|SH|1990|
UT|用水流量計|inst|電磁流量計|0|YR|kikan|
EL|受電盤|panel|6.6kV VCB 引込|0|SW|kikan|
EL|発電機連系盤|panel|6.6kV VCB・保護継電器|0|SW|kikan|
EL|高圧動力変圧器 No.1|tr|6.6kV/440V 1,500kVA 油入|0|SW|1990|
EL|高圧動力変圧器 No.2|tr|6.6kV/440V 1,500kVA 油入|0|SW|1990|
EL|低圧動力変圧器 No.1|tr|440V/210V 300kVA モールド|0|SW|kikan|
EL|低圧動力変圧器 No.2|tr|440V/210V 300kVA モールド|0|SW|kikan|
EL|照明変圧器|tr|6.6kV/210-105V 200kVA|0|SW|1990|
EL|高圧配電盤 No.1|panel|6.6kV フィーダ|0|SW|kikan|
EL|高圧配電盤 No.2|panel|6.6kV フィーダ|0|SW|kikan|
EL|高圧配電盤 No.3|panel|6.6kV フィーダ|0|SW|kikan|
EL|高圧配電盤 No.4|panel|6.6kV フィーダ|0|SW|kikan|
EL|高圧配電盤 No.5|panel|6.6kV フィーダ|0|SW|kikan|
EL|高圧配電盤 No.6|panel|6.6kV フィーダ（灰溶融用 休止）|0|SW|2003|
EL|低圧主幹盤 No.1|panel|440V ACB|0|SW|kikan|
EL|低圧主幹盤 No.2|panel|440V ACB|0|SW|kikan|
EL|低圧主幹盤 No.3|panel|210V ACB|0|SW|kikan|
EL|共通動力制御盤 No.1|panel|コントロールセンタ|0|SW|kikan|
EL|共通動力制御盤 No.2|panel|コントロールセンタ|0|SW|kikan|
EL|共通動力制御盤 No.3|panel|コントロールセンタ|0|SW|kikan|
EL|共通動力制御盤 No.4|panel|コントロールセンタ|0|SW|kikan|
EL|直流電源装置|panel|DC110V 鉛蓄電池 400Ah|0|SW|kikan|
EL|無停電電源装置|panel|UPS 30kVA|0|SW|kikan|
EL|非常用発電機|egen|ディーゼル 500kVA 6.6kV|0|IZ|1990|
EL|非常用発電機盤|panel|自動始動・同期|0|SW|1990|
EL|非常用発電機燃料小出槽|tank|950L|0|IZ|1990|
EL|進相コンデンサ No.1|tr|6.6kV 300kvar|0|SW|kikan|
EL|進相コンデンサ No.2|tr|6.6kV 300kvar|0|SW|kikan|
EL|共通計装電源盤|panel|AC100V 分電|0|SW|kikan|
EL|クレーン現場盤|panel|屋内自立形|0|SW|kikan|
EL|飛灰処理現場盤|panel|屋内自立形|0|SW|dxn|
IC|DCS オペレータコンソール No.1|dcs|液晶 2 画面|0|SW|kikan|
IC|DCS オペレータコンソール No.2|dcs|液晶 2 画面|0|SW|kikan|
IC|DCS オペレータコンソール No.3|dcs|液晶 2 画面|0|SW|kikan|
IC|DCS 共通コントローラ|dcs|二重化|0|SW|kikan|
IC|DCS ヒストリアンサーバ|dcs|データ収集・帳票|0|SW|kikan|
IC|運転帳票システム|dcs|日報・月報作成|0|SW|kikan|
IC|ITV 監視盤|panel|モニタ 12 面|0|SW|kikan|
IC|環境表示盤|panel|屋外 LED 表示|0|SW|2006|
IC|公害監視データ処理装置|dcs|測定値記録・報告|0|HI|kikan|
IC|煙突監視 ITV|camera|屋外 旋回|0|SW|kikan|
MS|煙突外筒|struct|RC H59m（内筒 3 本）|0|SH|1990|
MS|タービン室天井クレーン|crane|天井走行 10t|0|SN|1990|
MS|炉室電動ホイスト 1号|hoist|2.8t|0|SN|1990|
MS|炉室電動ホイスト 2号|hoist|2.8t|0|SN|1990|
MS|炉室電動ホイスト 3号|hoist|2.8t|0|SN|1990|
MS|洗車装置|misc|門型 自動|5.5|NB|2005|
MS|屋内消火栓ポンプ|pump_m|渦巻 消火用 30m³/h|15|YZ|1990|
"""

AM_TXT = """
AM|灰溶融炉本体|furnace|電気抵抗式 30t/日 耐火物内張|0|SH|2003|
AM|電極昇降装置 No.1|hyd|電動ウインチ式|3.7|SH|2003|
AM|電極昇降装置 No.2|hyd|電動ウインチ式|3.7|SH|2003|
AM|電極昇降装置 No.3|hyd|電動ウインチ式|3.7|SH|2003|
AM|溶融炉用変圧器|tr|6.6kV/400V 2,000kVA|0|SW|2003|
AM|溶融炉電源盤|panel|タップ切替・電流制御|0|SW|2003|
AM|灰供給コンベヤ A|conv|チェーンコンベヤ|3.7|KR|2003|
AM|灰供給コンベヤ B|conv|チェーンコンベヤ|3.7|KR|2003|
AM|溶融灰貯留ホッパ|tank|40m³|0|SH|2003|
AM|灰定量供給機|feeder|スクリュー|2.2|ON|2003|
AM|灰乾燥機|hx|回転式 蒸気加熱|5.5|SH|2003|
AM|スラグ水砕槽|tank|8m³|0|SH|2003|
AM|スラグコンベヤ A|conv|スクレーパ|2.2|KR|2003|
AM|スラグコンベヤ B|conv|ベルト|2.2|KR|2003|
AM|スラグ磁選機|misc|ドラム式|1.5|NB|2003|
AM|スラグ振動ふるい|misc|振動式|1.5|NB|2003|
AM|スラグストックホッパ|tank|30m³|0|SH|2003|
AM|溶融排ガス冷却塔|tower|水噴霧式|0|SH|2003|
AM|溶融排ガスバグフィルタ|bf|パルス式 ろ布 96 本|0|SH|2003|
AM|溶融 BF ダスト排出装置|conv|ロータリバルブ|0.75|KR|2003|
AM|溶融飛灰搬送コンベヤ|conv|スクリュー|2.2|KR|2003|
AM|溶融飛灰サイロ|tank|20m³|0|SH|2003|
AM|溶融排ガス誘引ファン|fan|6,000m³N/h|22|TK|2003|
AM|溶融排ガス分析計|anal|SOx・HCl・O2|0|HI|2003|
AM|溶融排ガス調節ダンパ|damper|電動|0.4|TG|2003|
AM|消石灰吹込装置（溶融）|feeder|テーブルフィーダ|0.4|ON|2003|
AM|活性炭吹込装置（溶融）|feeder|スクリュー|0.2|ON|2003|
AM|冷却塔噴霧水ポンプ A|pump_s|渦巻|3.7|YZ|2003|常用
AM|冷却塔噴霧水ポンプ B|pump_s|渦巻|3.7|YZ|2003|予備
AM|スラグ冷却水ポンプ A|pump_s|渦巻|5.5|YZ|2003|常用
AM|スラグ冷却水ポンプ B|pump_s|渦巻|5.5|YZ|2003|予備
AM|溶融系冷却塔|hx|開放式|3.7|NB|2003|
AM|窒素発生装置|comp|PSA 式 30m³N/h|7.5|ON|2003|
AM|溶融炉冷却ファン|fan|炉殻冷却|7.5|TK|2003|
AM|電極継足し用ホイスト|hoist|2t|0|SN|2003|
AM|溶融炉温度計 A|inst_a|放射温度計|0|YR|2003|
AM|溶融炉温度計 B|inst_a|熱電対 B|0|YR|2003|
AM|溶融炉現場盤|panel|屋内自立形|0|SW|2003|
AM|溶融設備動力制御盤|panel|コントロールセンタ|0|SW|2003|
AM|溶融炉監視カメラ|camera|水冷|0|SW|2003|
AM|溶融冷却水電動弁|mov|電動 50A|0.4|TG|2003|
"""

def _parse(txt):
    out = []
    for line in txt.strip().splitlines():
        p = line.split("|")
        out.append(dict(sys=p[0], name=p[1], cls=p[2], spec=p[3], kw=float(p[4]),
                        maker=p[5], origin=p[6], duty=p[7] if len(p) > 7 else ""))
    return out

# 代表機器（根拠記録の対象）の名前と号炉
REP_KEYS = [
    (2, "誘引通風機"), (2, "バグフィルタ"), (2, "過熱器"), (1, "燃焼ストーカ"),
    (0, "ごみクレーン 1号"), (0, "ごみクレーン 2号"), (0, "蒸気タービン"), (3, "湿式洗煙塔"),
    (2, "洗煙循環ポンプ A"), (0, "灰溶融炉本体"), (0, "非常用発電機"), (1, "ボイラ給水ポンプ A"),
    (1, "主灰押出装置"),
]


class Equip:
    __slots__ = ("tag", "unit", "sys", "name", "cls", "spec", "kw", "maker", "origin",
                 "duty", "install", "events", "health", "idle", "imp", "legal", "old_no", "aliases")

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k))

    @property
    def unit_name(self):
        return {0: "共通", 1: "1号炉", 2: "2号炉", 3: "3号炉"}[self.unit]

    def last(self, before_fy, kinds=("OVERHAUL", "REPLACE", "UPGRADE")):
        ys = [e["fy"] for e in self.events if e["fy"] < before_fy and e["kind"] in kinds]
        return max(ys) if ys else None


def _install_year(origin, unit):
    if origin == "1990":
        return 1990
    if origin == "dxn":
        return DXN_FY[unit] if unit else 2001
    if origin == "kikan":
        return 1990
    try:
        return int(origin)
    except ValueError:
        return 1990


def build_equipment():
    rng = random.Random(SEED)
    eqs = []
    ftab = _parse(FURNACE_TXT)
    assert len(ftab) == 165, len(ftab)
    ctab = _parse(COMMON_TXT)
    atab = _parse(AM_TXT)
    assert len(atab) == 41, len(atab)
    assert len(ctab) + len(atab) == 230, (len(ctab), len(atab))
    for unit in (1, 2, 3):
        cnt = {}
        for r in ftab:
            cnt[r["sys"]] = cnt.get(r["sys"], 0) + 1
            tag = "%d%s-%03d" % (unit, r["sys"], cnt[r["sys"]])
            eqs.append(Equip(tag=tag, unit=unit, idle=False, **r))
    cnt = {}
    for r in ctab + atab:
        cnt[r["sys"]] = cnt.get(r["sys"], 0) + 1
        tag = "0%s-%03d" % (r["sys"], cnt[r["sys"]])
        eqs.append(Equip(tag=tag, unit=0, idle=(r["sys"] == "AM"), **r))
    for e in eqs:
        e.install = _install_year(e.origin, e.unit)
        e.events = []
        e.imp = CLS[e.cls]["imp"]
        if e.duty == "予備" and e.imp == "A":
            e.imp = "B"
        e.health = rng.choice([-1, 0, 0, 0, 1])   # 劣化の速さ（-1 早い、+1 遅い）
        e.legal = ""
        e.aliases = {}
    _legal(eqs)
    return eqs


def _legal(eqs):
    for e in eqs:
        n = e.name
        if e.cls == "boiler":
            e.legal = "電事法 定期事業者検査（ボイラー）"
        elif n == "蒸気タービン":
            e.legal = "電事法 定期事業者検査（蒸気タービン）"
        elif n == "脱気器":
            e.legal = "安衛法 第一種圧力容器 性能検査"
        elif n.startswith("ごみクレーン") and "制御" not in n and "予備" not in n:
            e.legal = "クレーン則 性能検査・年次自主検査"
        elif n == "タービン室天井クレーン":
            e.legal = "クレーン則 性能検査・年次自主検査"
        elif n == "灰クレーン" or "ホイスト" in n:
            e.legal = "クレーン則 年次自主検査"
        elif n.startswith("計量機（"):
            e.legal = "計量法 定期検査"
        elif n == "非常用発電機":
            e.legal = "消防法 点検・電事法 保安規程"
        elif n == "屋内消火栓ポンプ" or n.startswith("ピット放水銃"):
            e.legal = "消防法 消防用設備等点検"
        elif n in ("受電盤", "発電機連系盤") or n.startswith("高圧"):
            e.legal = "電事法 保安規程（年次点検）"
        elif n in ("排ガス分析計", "ばいじん計"):
            e.legal = "大防法 測定（精度管理）"


# ------------------------------------------------------------------------
# 履歴（実際に行った作業）。2001 年度以前は散逸のため記録しない（Equip.install だけ持つ）。
def _unit_of_events(e):
    return e.unit


def gen_history(eqs):
    rng = random.Random(SEED + 1)
    for e in eqs:
        c = CLS[e.cls]
        ev = []
        start = e.install
        # 基幹改良で更新・追設
        if e.origin == "kikan":
            ky = KIKAN_FY[e.unit]
            ev.append(dict(fy=ky, kind="REPLACE", what="基幹的設備改良で更新", contract="KIKAN-%d" % e.unit))
            start = ky
        if e.origin == "dxn" and e.install >= 2001:
            ev.append(dict(fy=e.install, kind="UPGRADE", what="ダイオキシン類対策で追設", contract="DXN-%d" % e.unit))
        if e.idle:
            ev.append(dict(fy=AM_FY, kind="UPGRADE", what="灰溶融設備追設", contract="AM-2003"))
        last_oh = start
        last_rep = start
        stop = AM_STOP_FY if e.idle else 2026
        fy = max(start + 1, 2001)
        while fy <= stop:
            done = False
            if c.get("rep") and c["crep"] < 50000 and fy - last_rep >= c["rep"] + e.health * 2 + rng.choice([0, 0, 1, 2]):
                if not (e.unit and fy == KIKAN_FY[e.unit] and e.origin != "kikan" and False):
                    ev.append(dict(fy=fy, kind="REPLACE", what="更新", contract=None))
                    last_rep = last_oh = fy
                    done = True
            if not done and c.get("oh") and fy - last_oh >= c["oh"] + e.health + rng.choice([0, 0, 0, 1, -1]):
                ev.append(dict(fy=fy, kind="OVERHAUL", what=c.get("oh_name", "分解整備"), contract=None))
                last_oh = fy
                done = True
            if c.get("ins") and fy >= 2003 and (fy - start) % c["ins"] == 0 and not done:
                ev.append(dict(fy=fy, kind="INSPECT", what=c.get("ins_name", "精密点検"), contract=None))
            fy += 1
        e.events = sorted(ev, key=lambda x: x["fy"])
    _overrides(eqs)
    # 2026 年度は上期（9 月まで）の実績だけ：1号炉 定修（9 月）と共通のうち予備機のある機器
    for e in eqs:
        e.events = [x for x in e.events if x["fy"] < 2026 or e.unit == 1 or x.get("keep26")]


def find(eqs, unit, name):
    for e in eqs:
        if e.unit == unit and e.name == name:
            return e
    raise KeyError((unit, name))


def _overrides(eqs):
    """代表機器の履歴は手で書く（根拠記録と一致させる）。"""
    E = lambda u, n: find(eqs, u, n)
    # R1 2号炉 誘引通風機
    e = E(2, "誘引通風機"); e.health = 0
    e.events = [dict(fy=2013, kind="REPLACE", what="基幹的設備改良で更新（インバータ化）", contract="KIKAN-2"),
                dict(fy=2017, kind="OVERHAUL", what="軸受交換・羽根車清掃"),
                dict(fy=2020, kind="INSPECT", what="振動診断 4.6mm/s（ダスト付着によるアンバランス）"),
                dict(fy=2021, kind="OVERHAUL", what="羽根車清掃・現地バランス"),
                dict(fy=2022, kind="INSPECT", what="振動診断 3.1mm/s"),
                dict(fy=2023, kind="INSPECT", what="振動診断 4.2mm/s、羽根入口側 摩耗 最大 2.8mm"),
                dict(fy=2024, kind="OVERHAUL", what="羽根車 肉盛補修・現地バランス、軸受（両側）交換", contract="R6-2"),
                dict(fy=2025, kind="INSPECT", what="振動診断 1.9mm/s")]
    # R2 2号炉 バグフィルタ
    e = E(2, "バグフィルタ")
    e.events = [dict(fy=2001, kind="UPGRADE", what="ダイオキシン類対策で追設", contract="DXN-2"),
                dict(fy=2008, kind="OVERHAUL", what="ろ布全数交換（PTFE 混紡）"),
                dict(fy=2013, kind="OVERHAUL", what="ろ布全数交換・パルス弁交換（基幹改良と同時）", contract="KIKAN-2"),
                dict(fy=2022, kind="INSPECT", what="ろ布 抜取り強度試験（引張 残存 62%）"),
                dict(fy=2023, kind="REPAIR", what="ろ布破損 3 本交換（予備品）、差圧 1.65kPa"),
                dict(fy=2024, kind="OVERHAUL", what="ろ布全数交換 288 本（市支給品）", contract="R6-2")]
    # R3 2号炉 過熱器
    e = E(2, "過熱器")
    e.events = [dict(fy=2004, kind="INSPECT", what="肉厚測定"),
                dict(fy=2009, kind="REPLACE", what="2 次過熱器 管 全数取替（STB410→SUS310S）"),
                dict(fy=2013, kind="INSPECT", what="肉厚測定（基幹改良時）"),
                dict(fy=2018, kind="REPAIR", what="1 次過熱器 下段 管 4 本 部分取替"),
                dict(fy=2021, kind="INSPECT", what="肉厚測定 最小 3.6mm"),
                dict(fy=2022, kind="INSPECT", what="肉厚測定 最小 3.4mm"),
                dict(fy=2023, kind="INSPECT", what="肉厚測定 最小 3.3mm"),
                dict(fy=2024, kind="REPAIR", what="1 次過熱器 管 3 本 部分取替（追加）、肉厚測定 最小 3.1mm", contract="R6-2"),
                dict(fy=2025, kind="INSPECT", what="肉厚測定 最小 3.0mm")]
    # R4 1号炉 燃焼ストーカ
    e = E(1, "燃焼ストーカ")
    e.events = [dict(fy=2004, kind="OVERHAUL", what="火格子部分交換"),
                dict(fy=2007, kind="OVERHAUL", what="火格子部分交換"),
                dict(fy=2010, kind="OVERHAUL", what="火格子部分交換 30 枚"),
                dict(fy=2012, kind="INSPECT", what="基幹改良時 点検（駆動部のみ更新）", contract="KIKAN-1"),
                dict(fy=2013, kind="OVERHAUL", what="火格子部分交換 24 枚・サイドシール交換", contract="H25-1"),
                dict(fy=2016, kind="OVERHAUL", what="火格子部分交換 28 枚"),
                dict(fy=2019, kind="OVERHAUL", what="火格子部分交換 36 枚・可動フレーム補修"),
                dict(fy=2022, kind="OVERHAUL", what="火格子部分交換 40 枚（脱落 2 枚あり）"),
                dict(fy=2024, kind="INSPECT", what="火格子 摩耗測定（最大 9mm）"),
                dict(fy=2026, kind="OVERHAUL", what="火格子部分交換 44 枚")]
    # R5 ごみクレーン
    for n, evs in (("ごみクレーン 1号", [2011, 2014, 2017, 2020, 2023]), ("ごみクレーン 2号", [2011, 2015, 2019, 2024])):
        e = E(0, n)
        e.events = [dict(fy=2011, kind="REPLACE", what="基幹的設備改良で更新", contract="KIKAN-0")]
        for y in evs[1:]:
            e.events.append(dict(fy=y, kind="OVERHAUL", what="巻上ワイヤロープ交換・ブレーキ整備"))
        for y in range(2012, 2027):
            e.events.append(dict(fy=y, kind="INSPECT", what="年次自主検査" + ("・性能検査" if y % 2 == 1 else ""), keep26=False))
        e.events.sort(key=lambda x: x["fy"])
    # R6 蒸気タービン
    e = E(0, "蒸気タービン")
    e.events = [dict(fy=y, kind="OVERHAUL", what="開放点検（定期事業者検査）") for y in (2004, 2008, 2012, 2016, 2020)]
    e.events.append(dict(fy=2024, kind="OVERHAUL", what="開放点検（定期事業者検査）、動翼 先端エロージョン軽微", contract="R6-C"))
    # R7 3号炉 湿式洗煙塔
    e = E(3, "湿式洗煙塔")
    e.events = [dict(fy=2006, kind="OVERHAUL", what="ライニング全面更新（FRP）"),
                dict(fy=2014, kind="INSPECT", what="内面点検（基幹改良時）"),
                dict(fy=2019, kind="REPAIR", what="ライニング部分補修 4.5m²（気液接触部）"),
                dict(fy=2021, kind="INSPECT", what="内面点検 ふくれ 6 箇所"),
                dict(fy=2023, kind="INSPECT", what="内面点検 ピンホール 2 箇所（スパーク試験）・応急補修"),
                dict(fy=2024, kind="INSPECT", what="内面点検 ふくれ拡大・ピンホール 1 箇所"),
                dict(fy=2025, kind="OVERHAUL", what="ライニング全面更新（ビニルエステル FRP）")]
    # R8 2号炉 洗煙循環ポンプ
    e = E(2, "洗煙循環ポンプ A")
    e.events = [dict(fy=2008, kind="OVERHAUL", what="分解整備"), dict(fy=2012, kind="OVERHAUL", what="分解整備"),
                dict(fy=2016, kind="REPLACE", what="ポンプ本体更新（同型）"),
                dict(fy=2020, kind="OVERHAUL", what="分解整備・メカシール交換"),
                dict(fy=2024, kind="OVERHAUL", what="分解整備・羽根車交換（摩耗 腐食）・ライナ交換", contract="R6-2")]
    e = E(2, "洗煙循環ポンプ B")
    e.events = [dict(fy=2009, kind="OVERHAUL", what="分解整備"), dict(fy=2016, kind="REPLACE", what="ポンプ本体更新（同型）"),
                dict(fy=2022, kind="OVERHAUL", what="分解整備（予備機 運転 1,200h）")]
    # R9 灰溶融炉（休止）
    e = E(0, "灰溶融炉本体")
    e.events = [dict(fy=2003, kind="UPGRADE", what="灰溶融設備追設", contract="AM-2003"),
                dict(fy=2007, kind="OVERHAUL", what="耐火物部分補修"), dict(fy=2010, kind="OVERHAUL", what="耐火物部分補修・電極ホルダ交換"),
                dict(fy=2013, kind="INSPECT", what="休止（電力費のため 2013.10 停止）、内部清掃・乾燥保管")]
    for y in (2015, 2018, 2021, 2024):
        e.events.append(dict(fy=y, kind="INSPECT", what="休止設備 目視点検（直営）"))
    # R10 非常用発電機
    e = E(0, "非常用発電機")
    e.events = [dict(fy=2005, kind="OVERHAUL", what="オーバーホール"), dict(fy=2016, kind="OVERHAUL", what="オーバーホール・始動電動機交換"),
                dict(fy=2025, kind="INSPECT", what="負荷運転試験、メーカーより部品供給終了の通知（2025.9）")]
    # R11 1号炉 ボイラ給水ポンプ A
    e = E(1, "ボイラ給水ポンプ A")
    e.events = [dict(fy=2012, kind="REPLACE", what="基幹的設備改良で更新", contract="KIKAN-1"),
                dict(fy=2017, kind="OVERHAUL", what="分解整備"), dict(fy=2021, kind="OVERHAUL", what="分解整備"),
                dict(fy=2024, kind="REPAIR", what="メカニカルシール交換（直営、漏れ）", contract="DIRECT"),
                dict(fy=2025, kind="OVERHAUL", what="分解整備・軸スリーブ交換")]
    # R12 1号炉 主灰押出装置
    e = E(1, "主灰押出装置")
    e.events = [dict(fy=2005, kind="OVERHAUL", what="分解整備"), dict(fy=2009, kind="OVERHAUL", what="分解整備"),
                dict(fy=2013, kind="REPAIR", what="シュート当て板補修（直営）", contract="DIRECT"),
                dict(fy=2014, kind="OVERHAUL", what="分解整備・プッシャ板交換"),
                dict(fy=2018, kind="OVERHAUL", what="分解整備"), dict(fy=2021, kind="REPAIR", what="シュート当て板補修（直営）", contract="DIRECT"),
                dict(fy=2022, kind="OVERHAUL", what="分解整備"),
                dict(fy=2024, kind="REPAIR", what="シュート側板 穴あき 当て板補修（直営）", contract="DIRECT"),
                dict(fy=2026, kind="REPLACE", what="シュート更新（SUS304 12t）")]
    # 健全度（R8 版計画の 4 段階）
    for (u, n), h in {(2, "誘引通風機"): 3, (2, "バグフィルタ"): 4, (2, "過熱器"): 2, (1, "燃焼ストーカ"): 2,
                      (0, "ごみクレーン 1号"): 3, (0, "ごみクレーン 2号"): 3, (0, "蒸気タービン"): 3, (3, "湿式洗煙塔"): 4,
                      (2, "洗煙循環ポンプ A"): 4, (0, "灰溶融炉本体"): 1, (0, "非常用発電機"): 2,
                      (1, "ボイラ給水ポンプ A"): 3, (1, "主灰押出装置"): 3}.items():
        E(u, n).aliases["health4"] = h


# ------------------------------------------------------------------------
# 契約（見本で使うもの）
CONTRACTS = {
    "H25-1": dict(fy=2013, name="平成25年度 くぬぎ台環境センター 1号炉定期整備業務委託",
                  vendor="汐見重工環境株式会社", start=(2013, 5, 13), end=(2013, 6, 14), kind="業務委託"),
    "H25-3": dict(fy=2013, name="平成25年度 くぬぎ台環境センター 3号炉定期整備業務委託",
                  vendor="汐見重工環境株式会社", start=(2013, 11, 5), end=(2013, 12, 13), kind="業務委託"),
    "H25-C": dict(fy=2013, name="平成25年度 くぬぎ台環境センター 全炉停止時共通設備整備業務委託",
                  vendor="汐見重工環境株式会社", start=(2013, 9, 2), end=(2013, 9, 30), kind="業務委託"),
    "KIKAN-2": dict(fy=2013, name="くぬぎ台環境センター基幹的設備改良工事（2号炉）",
                    vendor="汐見重工環境株式会社", kind="請負"),
    "R6-1": dict(fy=2024, name="令和6年度 くぬぎ台環境センター 1号炉定期整備業務委託",
                 vendor="汐見重工環境株式会社", start=(2024, 9, 2), end=(2024, 9, 30), kind="業務委託"),
    "R6-2": dict(fy=2024, name="令和6年度 くぬぎ台環境センター 2号炉定期整備業務委託",
                 vendor="汐見重工環境株式会社", start=(2024, 10, 7), end=(2024, 11, 15), kind="業務委託"),
    "R6-C": dict(fy=2024, name="令和6年度 くぬぎ台環境センター 全炉停止時共通設備整備業務委託",
                 vendor="汐見重工環境株式会社", start=(2024, 10, 9), end=(2024, 11, 1), kind="業務委託"),
    "R6-3": dict(fy=2024, name="令和6年度 くぬぎ台環境センター 3号炉定期整備業務委託",
                 vendor="汐見重工環境株式会社", start=(2024, 11, 11), end=(2024, 12, 13), kind="業務委託"),
}


def world():
    eqs = build_equipment()
    gen_history(eqs)
    return eqs


if __name__ == "__main__":
    eqs = world()
    from collections import Counter
    print(len(eqs), Counter(e.unit for e in eqs), sum(e.idle for e in eqs))
    for u, n in REP_KEYS:
        e = find(eqs, u, n)
        print(e.tag, e.unit_name, e.name, e.install, [(x["fy"], x["kind"]) for x in e.events][-3:])
