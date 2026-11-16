# -*- coding: utf-8 -*-
"""【架空】しおさい町ごみ処理場（F04）の「正解の世界」。
機器 256 台、実際に行った作業、契約。帳票の生成器はすべてここから書き分ける。
今日 = 2026-11-13（2026 年度は計画＋上期実績）。乱数の種は固定。
"""
import random, datetime as dt

TODAY = dt.date(2026, 11, 13)
SEED = 4004
FAC = "【架空】しおさい町ごみ処理場"
FAC_PLAIN = "しおさい町ごみ処理場"
TOWN = "しおさい町"

# ---- 架空の会社・人（人は記号） ----
MAKER = "【架空】水樹プラント株式会社"
MAKER_S = "水樹プラント"
LOCAL = "【架空】有限会社 浜風機工"
LOCAL_S = "浜風機工"
CRANE = "【架空】港南クレーン整備株式会社"
CRANE_S = "港南クレーン整備"
ELEC = "【架空】みさき電気管理事務所"
ELEC_S = "みさき電気管理事務所"
ELEC_W = "【架空】しおさい電設株式会社"
ANAL = "【架空】東和環境分析センター株式会社"
FIRE = "【架空】しおさい防災設備"
SETSUBI = "【架空】有限会社 しおさい設備"
CONSUL = "【架空】株式会社 青潮技術コンサルタント"

def wareki(fy):
    """年度 → 和暦の短い表記（H26 / R4）"""
    if fy <= 2018:
        return "H%d" % (fy - 1988)
    return "R%d" % (fy - 2018)

def wareki_long(fy):
    if fy <= 2018:
        return "平成%d年度" % (fy - 1988)
    n = fy - 2018
    return "令和%s年度" % ("元" if n == 1 else str(n))

# ---- 設備記号（t3 が 2024 台帳で付けた管理番号に使う） ----
SYS = {
 "A": "受入供給設備", "B": "燃焼設備", "C": "燃焼ガス冷却設備", "D": "排ガス処理設備",
 "E": "通風設備", "F": "灰出し設備", "G": "飛灰処理設備", "H": "薬剤供給設備",
 "J": "給水設備", "K": "排水処理設備", "L": "燃料設備", "M": "圧縮空気設備",
 "N": "電気設備", "P": "計装設備", "R": "雑設備", "Z": "建築設備",
}
SYS_ORDER = list("ABCEDFGHJKLMNPRZ")  # 竣工図書の章立て順（通風が排ガス処理より前）

# 炉別機器 58 台（1 炉あたり）。
# (sys, 名称(計画・2024台帳), 竣工図書/2015台帳の名称, 仕様, 種別cls, 設置年度, 重要度, 図書No)
F_EQ = [
 ("B","ごみ投入ホッパ","ごみ投入ホッパ","鋼板製 容量 8m³","struct",1994,"B","M-201"),
 ("B","ホッパゲート","ホッパゲート","油圧式 1,600×1,400","hyd",1994,"B","M-202"),
 ("B","ホッパブリッジ解除装置","ブリッジ解除装置","油圧プッシャ式","hyd",1994,"C","M-203"),
 ("B","給じん装置","給じん装置","油圧プッシャ式 1.25t/h","feeder",1994,"A","M-204"),
 ("B","乾燥火格子","乾燥ストーカ","階段式 2.4m²","grate",1994,"A","M-205"),
 ("B","燃焼火格子","燃焼ストーカ","階段式 4.8m²","grate",1994,"A","M-206"),
 ("B","後燃焼火格子","後燃焼ストーカ","階段式 2.0m²","grate",1994,"A","M-207"),
 ("B","火格子駆動用油圧ユニット","ストーカ用油圧装置","7.5kW×1 14MPa","hyd",1994,"A","M-208"),
 ("B","給じん装置用油圧ユニット","給じん用油圧装置","3.7kW×1 14MPa","hyd",1994,"A","M-209"),
 ("B","焼却炉本体","焼却炉本体","耐火物・ケーシング 20t/16h","furnace",1994,"A","M-210"),
 ("B","炉壁冷却ファン","炉壁冷却用送風機","40m³/min 2.2kW","fan",1994,"B","M-211"),
 ("B","助燃バーナ","助燃バーナ","灯油 60L/h","burner",1994,"B","M-212"),
 ("B","助燃バーナ用送風機","助燃用送風機","3.7kW","fan",1994,"C","M-213"),
 ("B","再燃バーナ","再燃バーナ","灯油 40L/h","burner",1994,"B","M-214"),
 ("B","再燃バーナ用送風機","再燃用送風機","2.2kW","fan",1994,"C","M-215"),
 ("B","バーナ用燃料ポンプ","燃料ポンプ","歯車式 0.4kW","pump",1994,"C","M-216"),
 ("B","炉内監視用ITV","炉内監視カメラ","水冷ケース付","itv",1994,"C","E-411"),
 ("B","ホッパ監視用ITV","ホッパ監視カメラ","屋内用","itv",1994,"C","E-412"),
 ("B","落じんコンベヤ","落じんコンベヤ","スクリュー式 0.75kW","conv",1994,"B","M-217"),
 ("E","押込送風機","押込送風機","6,000m³N/h×4.0kPa 11kW","fan",1994,"A","M-401"),
 ("E","押込送風機入口ダンパ","押込用ダンパ","電動 φ450","damper",1994,"C","M-402"),
 ("E","二次送風機","二次送風機","2,000m³N/h×3.5kPa 5.5kW","fan",1994,"B","M-403"),
 ("E","空気予熱器","空気予熱器","鋼管式 ガス-空気","hx",1994,"B","M-404"),
 ("E","誘引通風機","誘引送風機","11,000m³N/h×4.5kPa 37kW","idf",1994,"A","M-405"),
 ("E","誘引通風機入口ダンパ","誘引用ダンパ","電動 φ600","damper",1994,"B","M-406"),
 ("E","誘引通風機用インバータ盤","（なし）","37kW用 屋内自立","panel",None,"A","E-405"),
 ("E","煙道","煙道","鋼板製 ガス冷却室出口～煙突","struct",1994,"B","M-407"),
 ("E","煙突内筒","煙突内筒","鋼板製 φ500 H40m","struct",1994,"B","M-408"),
 ("E","BFバイパスダンパ","バイパスダンパ","電動 φ600","damper",2001,"A","M-409"),
 ("C","ガス冷却室","ガス冷却室","水噴射式 900→200℃","cooler",1994,"A","M-301"),
 ("C","噴射水ポンプA","噴射水ポンプ","2m³/h×1.5MPa 3.7kW","pump",1994,"A","M-302"),
 ("C","噴射水ポンプB","噴射水ポンプ","2m³/h×1.5MPa 3.7kW","pump",1994,"A","M-302"),
 ("C","噴霧用空気圧縮機A","噴霧用空気圧縮機","スクリュー 7.5kW","comp",1994,"A","M-303"),
 ("C","噴霧用空気圧縮機B","噴霧用空気圧縮機","スクリュー 7.5kW","comp",1994,"A","M-303"),
 ("C","噴射水自動ストレーナ","噴射水ストレーナ","自動洗浄式 0.2kW","misc",1994,"C","M-304"),
 ("C","ガス冷却室ハンマリング装置","ハンマリング装置","エアハンマ式","misc",1994,"C","M-305"),
 ("C","ガス冷却室ダスト排出装置","ダスト排出装置","ロータリーバルブ 0.4kW","rv",1994,"B","M-306"),
 ("C","ガス冷却室ダストコンベヤ","ダストコンベヤ","スクリュー式 0.75kW","conv",1994,"B","M-307"),
 ("D","バグフィルタ","ろ過式集じん器","パルスジェット式 ろ布144本 ろ過面積305m²","bf",2001,"A","M-501"),
 ("D","BF加温ヒータ","ろ過式集じん器加温ヒータ","電気式 15kW","heater",2001,"B","M-502"),
 ("D","BFパルス用空気槽","払落し用空気槽","0.3m³ 0.7MPa（第二種圧力容器）","tank",2001,"B","M-503"),
 ("D","BFダスト排出装置","ろ過式集じん器ダスト排出装置","ロータリーバルブ 0.4kW","rv",2001,"B","M-504"),
 ("D","BFダストコンベヤ","ろ過式集じん器ダストコンベヤ","スクリュー式 0.75kW","conv",2001,"B","M-505"),
 ("D","BF保護用希釈空気ファン","希釈空気ファン","50m³/min 3.7kW","fan",2001,"B","M-506"),
 ("D","消石灰定量供給機","消石灰定量供給機","テーブルフィーダ 0.2kW","feeder",2001,"A","M-507"),
 ("D","活性炭定量供給機","活性炭定量供給機","スクリューフィーダ 0.1kW","feeder",2001,"A","M-508"),
 ("D","薬剤輸送ブロワ","薬剤輸送用ブロワ","ルーツ式 3.7kW","blower",2001,"A","M-509"),
 ("F","灰押出機","灰押出装置","油圧プッシャ式 水封","pusher",1994,"A","M-601"),
 ("F","灰押出機補給水ポンプ","灰押出装置補給水ポンプ","0.4kW","pump",1994,"C","M-602"),
 ("F","飛灰搬送コンベヤ","飛灰搬送コンベヤ","スクリュー式 1.5kW","conv",2001,"B","M-603"),
 ("B","油圧ユニット現場盤","油圧装置現場盤","屋内壁掛","panel",1994,"C","E-310"),
 ("N","炉別動力制御盤No.1","炉動力制御盤（No.1）","MCC 屋内自立","panel",1994,"A","E-301"),
 ("N","炉別動力制御盤No.2","炉動力制御盤（No.2）","MCC 屋内自立","panel",1994,"A","E-302"),
 ("N","炉現場操作盤","炉現場操作盤","屋内自立","panel",1994,"B","E-303"),
 ("P","燃焼制御盤","自動燃焼制御装置","PLC 屋内自立","panel",1994,"A","E-401"),
 ("P","BF制御盤","ろ過式集じん器制御盤","PLC 屋内自立","panel",2001,"A","E-402"),
 ("P","炉別計装盤","炉計装盤","屋内自立","panel",1994,"B","E-403"),
 ("J","炉別冷却水ヘッダ流量監視ユニット","冷却水監視装置","流量スイッチ付","misc",1994,"C","M-701"),
]
assert len(F_EQ) == 58, len(F_EQ)

# 共通設備 118 台 + 建築設備 22 台
C_EQ = [
 ("A","計量機","トラックスケール","ロードセル式 20t 4点支持","scale",1994,"A","M-101"),
 ("A","計量データ処理装置","計量データ処理装置","PC＋プリンタ","pc",2009,"B","E-101"),
 ("A","プラットホーム投入扉No.1","投入扉（No.1）","油圧観音開き 2.4×4.5m","door",1994,"B","M-102"),
 ("A","プラットホーム投入扉No.2","投入扉（No.2）","油圧観音開き 2.4×4.5m","door",1994,"B","M-102"),
 ("A","プラットホーム投入扉No.3","投入扉（No.3）","油圧観音開き 2.4×4.5m","door",1994,"B","M-102"),
 ("A","プラットホーム投入扉No.4","投入扉（No.4）","油圧観音開き 2.4×4.5m","door",1994,"C","M-102"),
 ("A","投入扉用油圧ユニット","投入扉用油圧装置","5.5kW","hyd",1994,"B","M-103"),
 ("A","エアカーテン","エアカーテン","上吹出し 1.5kW×2","fan",1994,"C","M-104"),
 ("A","ごみクレーン","ごみクレーン","吊上荷重3.2t バケット1.0m³油圧開閉 揚程18m","crane",1994,"A","M-105"),
 ("A","ごみクレーン操作盤","ごみクレーン操作盤","中央操作室 椅子式","panel",1994,"A","E-102"),
 ("A","ごみクレーン動力盤","ごみクレーン動力盤","クレーン電気室 自立","panel",1994,"A","E-103"),
 ("A","ごみピット脱臭用送風機","脱臭用送風機","150m³/min 7.5kW","fan",1994,"B","M-106"),
 ("A","脱臭装置","脱臭装置","活性炭吸着塔","misc",1994,"B","M-107"),
 ("A","ごみピット消臭剤噴霧装置","消臭剤噴霧装置","ポンプ 0.4kW","misc",1994,"C","M-108"),
 ("A","ごみピット監視用ITV","ピット監視カメラ","屋内用","itv",1994,"C","E-104"),
 ("A","ピット火災検知装置","ピット火災検知器","赤外線式 2基","misc",2012,"B","E-105"),
 ("A","ピット放水銃","放水銃","手動旋回式","misc",1994,"C","M-109"),
 ("A","粗大ごみ切断機","粗大ごみ切断機","油圧せん断式","shear",1994,"C","M-110"),
 ("A","粗大ごみ切断機用油圧ユニット","切断機用油圧装置","11kW","hyd",1994,"C","M-111"),
 ("F","灰出しコンベヤNo.1","灰搬出コンベヤ（No.1）","スクレーパ式 幅400 機長18.5m 2.2kW","ashconv",1994,"A","M-604"),
 ("F","灰出しコンベヤNo.2","灰搬出コンベヤ（No.2）","スクレーパ式 幅400 機長12m 1.5kW","ashconv",1994,"A","M-605"),
 ("F","灰バンカ","灰バンカ","鋼板製 12m³","struct",1994,"B","M-606"),
 ("F","灰バンカゲート","灰バンカゲート","油圧カットゲート","hyd",1994,"B","M-607"),
 ("F","灰バンカ用油圧ユニット","灰バンカ用油圧装置","3.7kW","hyd",1994,"B","M-608"),
 ("F","磁選機","磁選機","吊下げ式 1.5kW","misc",1994,"C","M-609"),
 ("F","灰出し散水ポンプ","灰散水ポンプ","1.5kW","pump",1994,"C","M-610"),
 ("F","灰汚水ポンプ","灰汚水ポンプ","水中 0.75kW","pump",1994,"C","M-611"),
 ("F","灰出し設備現場盤","灰出し現場盤","屋内壁掛","panel",1994,"C","E-501"),
 ("G","飛灰貯留槽","飛灰貯留槽","鋼板製 20m³","struct",2001,"B","M-801"),
 ("G","飛灰貯留槽用バグフィルタ","飛灰貯留槽集じん器","小型パルス式","bf_s",2001,"C","M-802"),
 ("G","飛灰排出装置","飛灰切出装置","ロータリーバルブ 0.4kW","rv",2001,"B","M-803"),
 ("G","飛灰定量供給機","飛灰定量供給機","スクリュー式 0.75kW","feeder",2001,"A","M-804"),
 ("G","混練機","混練機","二軸パドル 300kg/h 5.5kW","mixer",2001,"A","M-805"),
 ("G","キレート薬剤タンク","薬剤タンク","FRP 2m³","tank_s",2001,"C","M-806"),
 ("G","キレート薬剤希釈槽撹拌機","薬剤希釈槽撹拌機","0.4kW","misc",2001,"C","M-807"),
 ("G","キレート薬剤ポンプA","薬剤注入ポンプ","ダイヤフラム 0.2kW","pump",2001,"B","M-808"),
 ("G","キレート薬剤ポンプB","薬剤注入ポンプ","ダイヤフラム 0.2kW","pump",2001,"B","M-808"),
 ("G","加湿水ポンプA","加湿水ポンプ","0.75kW","pump",2001,"B","M-809"),
 ("G","加湿水ポンプB","加湿水ポンプ","0.75kW","pump",2001,"B","M-809"),
 ("G","処理灰コンベヤ","処理物コンベヤ","ベルト式 0.75kW","conv",2001,"B","M-810"),
 ("H","消石灰貯槽","消石灰貯槽","鋼板製 15m³","struct",2001,"B","M-901"),
 ("H","消石灰貯槽用バグフィルタ","消石灰貯槽集じん器","小型パルス式","bf_s",2001,"C","M-902"),
 ("H","消石灰切出装置","消石灰切出装置","ロータリーバルブ 0.4kW","rv",2001,"B","M-903"),
 ("H","活性炭貯槽","活性炭貯槽","鋼板製 5m³","struct",2001,"B","M-904"),
 ("H","活性炭貯槽用バグフィルタ","活性炭貯槽集じん器","小型パルス式","bf_s",2001,"C","M-905"),
 ("H","活性炭切出装置","活性炭切出装置","ロータリーバルブ 0.2kW","rv",2001,"B","M-906"),
 ("J","井水揚水ポンプNo.1","井水ポンプ（No.1）","水中 5.5kW","pump",1994,"A","M-1001"),
 ("J","井水揚水ポンプNo.2","井水ポンプ（No.2）","水中 5.5kW","pump",1994,"A","M-1001"),
 ("J","井水除鉄装置","除鉄ろ過器","接触酸化ろ過","misc",1994,"B","M-1002"),
 ("J","井水ろ過装置","砂ろ過器","圧力式","misc",1994,"C","M-1003"),
 ("J","受水槽","受水槽","FRP 30m³","tank_s",1994,"B","M-1004"),
 ("J","高置水槽","高置水槽","FRP 5m³","tank_s",1994,"C","M-1005"),
 ("J","プラント用水ポンプA","プラント用水ポンプ","3.7kW","pump",1994,"B","M-1006"),
 ("J","プラント用水ポンプB","プラント用水ポンプ","3.7kW","pump",1994,"B","M-1006"),
 ("J","機器冷却水ポンプA","機器冷却水ポンプ","5.5kW","pump",1994,"A","M-1007"),
 ("J","機器冷却水ポンプB","機器冷却水ポンプ","5.5kW","pump",1994,"A","M-1007"),
 ("J","冷却塔","冷却塔","開放式 20RT","misc",1994,"B","M-1008"),
 ("J","噴射水槽","噴射水槽","SUS 10m³","tank_s",1994,"B","M-1009"),
 ("K","汚水受槽ポンプA","汚水受槽ポンプ","水中 1.5kW","pump",1994,"B","M-1101"),
 ("K","汚水受槽ポンプB","汚水受槽ポンプ","水中 1.5kW","pump",1994,"B","M-1101"),
 ("K","ピット汚水ポンプ","ピット汚水ポンプ","水中 1.5kW","pump",1994,"B","M-1102"),
 ("K","ピット汚水ろ過器","ピット汚水ろ過器","スクリーン式","misc",1994,"C","M-1103"),
 ("K","反応槽撹拌機No.1","反応槽撹拌機","0.75kW","mixer_s",1994,"C","M-1104"),
 ("K","反応槽撹拌機No.2","反応槽撹拌機","0.75kW","mixer_s",1994,"C","M-1104"),
 ("K","凝集沈殿槽汚泥掻寄機","凝集沈殿槽","中心駆動 0.4kW","misc",1994,"B","M-1105"),
 ("K","汚泥引抜ポンプ","汚泥引抜ポンプ","0.75kW","pump",1994,"C","M-1106"),
 ("K","処理水ろ過器","ろ過器","圧力式","misc",1994,"C","M-1107"),
 ("K","処理水ポンプA","処理水ポンプ","1.5kW","pump",1994,"B","M-1108"),
 ("K","処理水ポンプB","処理水ポンプ","1.5kW","pump",1994,"B","M-1108"),
 ("K","凝集剤注入ポンプ","薬注ポンプ（凝集剤）","ダイヤフラム 0.1kW","pump",1994,"C","M-1109"),
 ("K","苛性ソーダ注入ポンプ","薬注ポンプ（苛性ソーダ）","ダイヤフラム 0.1kW","pump",1994,"C","M-1109"),
 ("K","硫酸注入ポンプ","薬注ポンプ（硫酸）","ダイヤフラム 0.1kW","pump",1994,"C","M-1109"),
 ("K","排水処理盤","排水処理制御盤","屋内自立","panel",1994,"B","E-601"),
 ("L","灯油貯留タンク","灯油タンク","地下 10kL","tank",1994,"B","M-1201"),
 ("L","灯油移送ポンプA","灯油移送ポンプ","0.4kW","pump",1994,"B","M-1202"),
 ("L","灯油移送ポンプB","灯油移送ポンプ","0.4kW","pump",1994,"B","M-1202"),
 ("L","灯油サービスタンク","サービスタンク","屋内 950L","tank_s",1994,"C","M-1203"),
 ("M","空気圧縮機No.1","空気圧縮機（No.1）","スクリュー 15kW 0.7MPa","comp",2010,"A","M-1301"),
 ("M","空気圧縮機No.2","空気圧縮機（No.2）","レシプロ 15kW 0.7MPa","comp",1994,"A","M-1301"),
 ("M","冷凍式除湿器","除湿装置","冷凍式","misc",1994,"B","M-1302"),
 ("M","空気槽No.1","空気槽（雑用）","1.0m³ 0.7MPa（第二種圧力容器）","tank",1994,"B","M-1303"),
 ("M","空気槽No.2","空気槽（計装用）","0.5m³ 0.7MPa（第二種圧力容器）","tank",1994,"B","M-1304"),
 ("N","高圧受電盤","高圧受電盤","6.6kV 屋内キュービクル VCB","hv",1994,"A","E-001"),
 ("N","高圧き電盤","高圧き電盤","6.6kV VCB","hv",1994,"A","E-002"),
 ("N","動力変圧器盤","動力用変圧器","6.6kV/420V 500kVA","tr",1994,"A","E-003"),
 ("N","照明変圧器盤","照明用変圧器","6.6kV/210-105V 75kVA","tr",1994,"B","E-004"),
 ("N","進相コンデンサ盤","高圧進相コンデンサ","6.6kV 100kvar","hv",1994,"B","E-005"),
 ("N","低圧主幹盤","低圧主幹盤","420V","panel",1994,"A","E-006"),
 ("N","共通動力制御盤No.1","共通動力制御盤（No.1）","MCC","panel",1994,"A","E-007"),
 ("N","共通動力制御盤No.2","共通動力制御盤（No.2）","MCC","panel",1994,"A","E-008"),
 ("N","共通動力制御盤No.3","共通動力制御盤（No.3）","MCC","panel",1994,"B","E-009"),
 ("N","灰出し・飛灰動力制御盤","灰・飛灰動力盤","MCC","panel",2001,"B","E-010"),
 ("N","非常用発電機","非常用発電装置","ディーゼル 75kVA 200V","gen",1994,"A","E-011"),
 ("N","非常用発電機盤","発電機盤","自立","panel",1994,"A","E-012"),
 ("N","直流電源装置","直流電源装置","DC100V 鉛蓄電池","dc",1994,"A","E-013"),
 ("N","無停電電源装置","無停電電源装置","5kVA","ups",2008,"B","E-014"),
 ("N","電灯主幹盤","電灯主幹盤","210-105V","panel",1994,"B","E-015"),
 ("N","屋外照明盤","屋外照明盤","屋外壁掛","panel",1994,"C","E-016"),
 ("N","計量棟分電盤","計量棟分電盤","屋内壁掛","panel",1994,"C","E-017"),
 ("P","中央監視操作盤","中央監視盤","PLC＋監視PC","panel",1994,"A","E-701"),
 ("P","データ処理装置","データロガ","PC 日報・月報","pc",1994,"B","E-702"),
 ("P","共通計装盤","共通計装盤","屋内自立","panel",1994,"B","E-703"),
 ("P","環境情報表示盤","排ガス表示盤","屋外LED","panel",2002,"C","E-704"),
 ("P","ITV監視盤","ITVモニタ盤","モニタ4面","itv",1994,"C","E-705"),
 ("R","煙突（外筒）","煙突","RC造 H40m","struct",1994,"B","M-1401"),
 ("R","保守用ホイスト","保守用電動ホイスト","2.8t 炉室天井","hoist",1994,"B","M-1402"),
 ("R","機器搬入用ホイスト","搬入用電動ホイスト","1.0t 薬品室","hoist",2001,"C","M-1403"),
 ("R","洗車装置","洗車設備","高圧洗浄 3.7kW","misc",1994,"C","M-1404"),
 ("R","真空掃除装置","真空掃除装置","据置式 7.5kW","misc",1994,"C","M-1405"),
 ("R","床洗浄ポンプ","床洗浄ポンプ","2.2kW","pump",1994,"C","M-1406"),
 ("R","炉室給気ファンNo.1","炉室給気ファン","5.5kW","fan",1994,"C","M-1407"),
 ("R","炉室給気ファンNo.2","炉室給気ファン","5.5kW","fan",1994,"C","M-1407"),
 ("R","炉室排気ファンNo.1","炉室排気ファン","5.5kW","fan",1994,"C","M-1408"),
 ("R","炉室排気ファンNo.2","炉室排気ファン","5.5kW","fan",1994,"C","M-1408"),
 ("R","灰出し室排風機","灰出し室排風機","3.7kW","fan",1994,"C","M-1409"),
 ("R","飛灰処理室排風機","飛灰処理室排風機","2.2kW","fan",2001,"C","M-1410"),
 ("R","薬品室換気ファン","薬品室排風機","1.5kW","fan",2001,"C","M-1411"),
 ("R","プラットホーム洗浄ポンプ","プラットホーム洗浄ポンプ","2.2kW","pump",1994,"C","M-1412"),
 # 建築設備 22
 ("Z","中央制御室空調機No.1","中央制御室空調機","パッケージ 10kW","ac",2006,"B","Z-01"),
 ("Z","中央制御室空調機No.2","中央制御室空調機","パッケージ 10kW","ac",2006,"B","Z-01"),
 ("Z","事務室空調機","事務室空調機","パッケージ 8kW","ac",1994,"C","Z-02"),
 ("Z","電気室空調機","電気室空調機","パッケージ 12.5kW","ac",1994,"B","Z-03"),
 ("Z","見学者ホール空調機","見学者ホール空調機","パッケージ 8kW","ac",1994,"C","Z-04"),
 ("Z","管理棟換気ファンNo.1","管理棟換気扇","0.4kW","fan",1994,"C","Z-05"),
 ("Z","管理棟換気ファンNo.2","管理棟換気扇","0.4kW","fan",1994,"C","Z-05"),
 ("Z","加圧給水ポンプユニット","給水ポンプユニット","1.5kW×2 交互","pump",2003,"B","Z-06"),
 ("Z","雑排水ポンプA","雑排水ポンプ","水中 0.75kW","pump",1994,"C","Z-07"),
 ("Z","雑排水ポンプB","雑排水ポンプ","水中 0.75kW","pump",1994,"C","Z-07"),
 ("Z","汚水ポンプ","汚水ポンプ","水中 0.75kW","pump",1994,"C","Z-08"),
 ("Z","合併処理浄化槽","合併処理浄化槽","30人槽","tank_s",1994,"B","Z-09"),
 ("Z","浄化槽ブロワA","浄化槽ブロワ","0.1kW","blower",1994,"C","Z-10"),
 ("Z","浄化槽ブロワB","浄化槽ブロワ","0.1kW","blower",1994,"C","Z-10"),
 ("Z","屋内消火栓ポンプ","消火栓ポンプ","7.5kW","pump",1994,"A","Z-11"),
 ("Z","消火栓ポンプ制御盤","消火ポンプ制御盤","屋内","panel",1994,"A","Z-12"),
 ("Z","自動火災報知設備受信機","自火報受信機","P型1級 15回線","fire",1994,"A","Z-13"),
 ("Z","非常放送設備","非常放送設備","アンプ 120W","fire",1994,"B","Z-14"),
 ("Z","電気温水器","電気温水器","370L","misc",1994,"C","Z-15"),
 ("Z","人荷用エレベーター","エレベーター","1,000kg 4停止","elev",1994,"B","Z-16"),
 ("Z","管理棟自動ドア","自動ドア","片引き","misc",2011,"C","Z-17"),
 ("Z","雨水排水ポンプ","雨水ポンプ","水中 1.5kW","pump",1994,"C","Z-18"),
]
assert len(C_EQ) == 140, len(C_EQ)
assert sum(1 for e in C_EQ if e[0] == "Z") == 22

class Eq:
    def __init__(s, **k): s.__dict__.update(k)
    def __repr__(s): return "<%s %s>" % (s.mid, s.name)

def build_equipment():
    eqs = []
    no = 0
    # 計画・台帳の並び：竣工図書の章立て順。炉別は 1号→2号の順に設備ごと。
    for sysc in SYS_ORDER:
        for (sy, nm, onm, spec, cls, inst, imp, dno) in C_EQ:
            pass
    # 管理番号は 号炉-設備記号-連番（2024 t3）
    seq = {}
    for sysc in SYS_ORDER:
        for f in (1, 2):
            for (sy, nm, onm, spec, cls, inst, imp, dno) in F_EQ:
                if sy != sysc: continue
                k = (f, sy); seq[k] = seq.get(k, 0) + 1
                eqs.append(Eq(furnace=f, sys=sy, name=nm, oname=onm, spec=spec, cls=cls,
                              inst=(inst if inst else (2015 if f == 1 else 2016)),
                              imp=imp, dno=dno, mid="%d-%s-%02d" % (f, sy, seq[k]),
                              added_2015=(inst is None)))
        for (sy, nm, onm, spec, cls, inst, imp, dno) in C_EQ:
            if sy != sysc: continue
            k = (0, sy); seq[k] = seq.get(k, 0) + 1
            eqs.append(Eq(furnace=0, sys=sy, name=nm, oname=onm, spec=spec, cls=cls, inst=inst,
                          imp=imp, dno=dno, mid="0-%s-%02d" % (sy, seq[k]), added_2015=False))
    for i, e in enumerate(eqs, 1):
        e.idx = i
    return eqs

EQ = build_equipment()
assert len(EQ) == 256
BYMID = {e.mid: e for e in EQ}

def find(name, furnace=0):
    for e in EQ:
        if e.name == name and e.furnace == furnace:
            return e
    raise KeyError((name, furnace))

# ---- 作業（実際に行ったこと）----
# kind: INSPECT / OVERHAUL / REPAIR(=補修, 属性は OVERHAUL 扱い) / REPLACE / UPGRADE
class Work:
    def __init__(s, mid, fy, kind, desc, contract, start=None, end=None, mh=None, vendor=None, src="hand"):
        s.mid, s.fy, s.kind, s.desc, s.contract = mid, fy, kind, desc, contract
        s.start, s.end, s.mh, s.vendor, s.src = start, end, mh, vendor, src
    def __repr__(s): return "<W %s %s %s %s>" % (s.fy, s.mid, s.kind, s.desc)

D = dt.date

# 契約（見本に使うものは詳細を持つ）
CONTRACTS = {
 "H26-MK": dict(fy=2014, name="平成26年度 しおさい町ごみ処理場 定期点検整備業務委託", vendor=MAKER, type="委託",
                start=D(2014,7,1), end=D(2015,1,30), site=(D(2014,10,14), D(2014,10,31))),
 "H27-28-EN": dict(fy=2015, name="しおさい町ごみ処理場 長寿命化改良工事（平成27〜28年度）", vendor=MAKER, type="請負",
                start=D(2015,6,15), end=D(2017,3,17)),
 "H29-MK": dict(fy=2017, name="平成29年度 しおさい町ごみ処理場 定期点検整備業務委託", vendor=MAKER, type="委託",
                start=D(2017,7,3), end=D(2018,1,31), site=(D(2017,10,10), D(2017,10,27))),
 "H30-ASH": dict(fy=2018, name="平成30年度 しおさい町ごみ処理場 灰出しコンベヤ補修工事", vendor=LOCAL, type="請負",
                start=D(2018,10,15), end=D(2018,11,30), site=(D(2018,10,24), D(2018,10,25))),
 "R4-MK": dict(fy=2022, name="令和4年度 しおさい町ごみ処理場 定期点検整備業務委託", vendor=MAKER, type="委託",
                start=D(2022,7,1), end=D(2023,1,31), site=(D(2022,10,11), D(2022,10,28))),
 "R4-CH": dict(fy=2022, name="令和4年度 しおさい町ごみ処理場 灰出しコンベヤチェーン取替修繕", vendor=LOCAL, type="請負",
                start=D(2022,9,12), end=D(2022,11,30), site=(D(2022,10,18), D(2022,10,21))),
}

def m(name, f=0):
    return find(name, f).mid

WORKS = []
def W(name, f, fy, kind, desc, contract, **k):
    WORKS.append(Work(m(name, f), fy, kind, desc, contract, **k))

# ---------------- 代表機器の履歴（根拠記録は notes） ----------------
# R1 ごみクレーン（予備なし）
W("ごみクレーン",0,2013,"OVERHAUL","巻上ワイヤロープ取替","H25-地元",vendor=CRANE)
W("ごみクレーン",0,2017,"OVERHAUL","バケット油圧ユニット整備（ポンプ・シリンダパッキン）","H29-CR",vendor=CRANE,start=D(2017,10,18),end=D(2017,10,19))
W("ごみクレーン",0,2019,"OVERHAUL","走行車輪 2輪取替","R1-CR",vendor=CRANE)
W("ごみクレーン",0,2021,"OVERHAUL","巻上ワイヤロープ取替","R3-CR",vendor=CRANE)
W("ごみクレーン",0,2022,"INSPECT","年次自主検査（横行レール摩耗・制御部品供給終了の確認）","R4-MK",vendor=MAKER,start=D(2022,10,19),end=D(2022,10,19))
W("ごみクレーン操作盤",0,2024,"REPLACE","クレーン制御装置更新（操作盤PLC・インバータ）","R6-CRUP",vendor=MAKER)
W("ごみクレーン動力盤",0,2024,"REPLACE","クレーン制御装置更新（動力盤インバータ）","R6-CRUP",vendor=MAKER)
W("ごみクレーン",0,2025,"INSPECT","クレーン性能検査（検査証更新）","R7-MK",vendor=MAKER)
# R2 灰出しコンベヤNo.1
W("灰出しコンベヤNo.1",0,2018,"OVERHAUL","不良リンク8リンク交換・駆動スプロケット肉盛・スクレーパ6枚取替","H30-ASH",vendor=LOCAL,start=D(2018,10,23),end=D(2018,10,25),mh=72.5)
W("灰出しコンベヤNo.1",0,2021,"INSPECT","チェーン伸び測定 2.1%（管理値2.0%超）","R3-MK",vendor=MAKER)
W("灰出しコンベヤNo.1",0,2022,"REPLACE","チェーン全数・スプロケット取替、テール軸受取替（追加）","R4-CH",vendor=LOCAL,start=D(2022,10,18),end=D(2022,10,21),mh=92.0)
W("灰出しコンベヤNo.1",0,2025,"INSPECT","チェーン伸び測定 0.6%・スクレーパ点検","R7-MK",vendor=MAKER)
W("燃焼火格子",2,2025,"OVERHAUL","火格子片 14枚取替（摩耗限度到達分）","R7-MK",vendor=MAKER)
W("灰バンカゲート",0,2026,"OVERHAUL","油圧シリンダパッキン取替","R8-地元2",vendor=LOCAL,start=D(2026,6,16),end=D(2026,6,16))
# R3 1号炉 燃焼火格子
W("燃焼火格子",1,2012,"OVERHAUL","火格子片 24枚取替","H24-MK",vendor=MAKER)
W("燃焼火格子",1,2015,"REPLACE","火格子全面更新（長寿命化）","H27-28-EN",vendor=MAKER)
W("燃焼火格子",1,2022,"OVERHAUL","火格子片 12枚取替（摩耗限度到達分）","R4-MK",vendor=MAKER,start=D(2022,10,13),end=D(2022,10,14))
W("燃焼火格子",1,2024,"OVERHAUL","火格子片 18枚取替","R6-MK",vendor=MAKER)
# R4 2号炉 ガス冷却室
W("ガス冷却室",2,2016,"REPLACE","内面耐火物・噴射ノズル更新（長寿命化）","H27-28-EN",vendor=MAKER)
W("ガス冷却室",2,2022,"INSPECT","下部ホッパ板厚測定 最小3.4mm（原厚6.0）","R4-MK",vendor=MAKER,start=D(2022,10,24),end=D(2022,10,24))
W("ガス冷却室",2,2023,"OVERHAUL","下部ホッパ当て板補修","R5-地元",vendor=LOCAL)
# R5 バグフィルタ
W("バグフィルタ",1,2015,"OVERHAUL","ろ布全数取替（長寿命化）","H27-28-EN",vendor=MAKER)
W("バグフィルタ",2,2016,"OVERHAUL","ろ布全数取替（長寿命化）","H27-28-EN",vendor=MAKER)
W("バグフィルタ",1,2021,"OVERHAUL","破損ろ布2本取替（予備品）","R3-MK",vendor=MAKER)
W("バグフィルタ",1,2022,"OVERHAUL","ろ布全数取替 144本","R4-MK",vendor=MAKER,start=D(2022,10,13),end=D(2022,10,15))
W("バグフィルタ",2,2022,"OVERHAUL","破損ろ布2本取替（町支給予備品）","R4-MK",vendor=MAKER,start=D(2022,10,22),end=D(2022,10,22))
W("バグフィルタ",2,2023,"OVERHAUL","ろ布全数取替 144本","R5-MK",vendor=MAKER)
# R6 1号炉 誘引通風機
W("誘引通風機",1,2015,"REPLACE","誘引通風機更新（インバータ化）","H27-28-EN",vendor=MAKER)
W("誘引通風機用インバータ盤",1,2015,"UPGRADE","インバータ盤新設","H27-28-EN",vendor=MAKER)
W("誘引通風機",2,2016,"REPLACE","誘引通風機更新（インバータ化）","H27-28-EN",vendor=MAKER)
W("誘引通風機用インバータ盤",2,2016,"UPGRADE","インバータ盤新設","H27-28-EN",vendor=MAKER)
W("誘引通風機",1,2022,"INSPECT","振動測定 3.1mm/s（R3 2.6）","R4-MK",vendor=MAKER,start=D(2022,10,17),end=D(2022,10,17))
# 2026年度は10月定修で軸受交換（計画）。上期実績には無い。
# R7 1号炉 噴射水ポンプA/B
W("噴射水ポンプA",1,2012,"OVERHAUL","分解整備","H24-地元",vendor=LOCAL)
W("噴射水ポンプB",1,2012,"OVERHAUL","分解整備","H24-地元",vendor=LOCAL)
W("噴射水ポンプA",1,2019,"OVERHAUL","メカニカルシール取替","R1-地元",vendor=LOCAL)
W("噴射水ポンプA",1,2022,"OVERHAUL","メカニカルシール取替（再漏れ）","R4-PA",vendor=LOCAL,start=D(2022,6,14),end=D(2022,6,14))
W("噴射水ポンプA",1,2024,"REPLACE","ポンプ更新","R6-地元",vendor=LOCAL)
W("噴射水ポンプB",1,2025,"REPLACE","ポンプ更新（R6から繰延べ）","R7-地元",vendor=LOCAL)
# R8 計量機
for fy in range(2012, 2027, 2):
    W("計量機",0,fy,"INSPECT","計量法 定期検査","%s-計量" % wareki(fy),vendor="検査機関")
W("計量機",0,2018,"OVERHAUL","ロードセル1個取替（定期検査前点検で器差不良）","H30-計量",vendor=SETSUBI)
W("計量データ処理装置",0,2025,"REPLACE","計量データ処理装置更新（OSサポート終了）","R7-計量",vendor=SETSUBI)
# R9 非常用発電機
W("非常用発電機",0,2019,"OVERHAUL","冷却水ホース・始動用蓄電池取替","R1-発電機",vendor=SETSUBI)
W("非常用発電機",0,2023,"INSPECT","負荷運転（消防点検）","R5-消防",vendor=FIRE)
# R10 空気圧縮機No.2（予備）
W("空気圧縮機No.2",0,2014,"OVERHAUL","吸排気弁・ピストンリング取替","H26-MK",vendor=MAKER,start=D(2014,10,22),end=D(2014,10,22))
W("空気圧縮機No.2",0,2026,"REPLACE","空気圧縮機更新（スクリュー式へ）R7から繰延べ","R8-地元",vendor=LOCAL,start=D(2026,7,8),end=D(2026,7,9))
# R11 高圧受電盤
W("高圧受電盤",0,2016,"INSPECT","年次点検でVCB絶縁抵抗低下指摘","H28-保安",vendor=ELEC)
W("高圧受電盤",0,2017,"REPLACE","真空遮断器（VCB）取替","H29-VCB",vendor=ELEC_W,start=D(2017,10,19),end=D(2017,10,19))
# R12 2号炉 火格子駆動用油圧ユニット
W("火格子駆動用油圧ユニット",2,2020,"REPLACE","油圧ポンプ取替（油漏れ・圧力低下）","R2-地元",vendor=LOCAL)
W("火格子駆動用油圧ユニット",2,2022,"INSPECT","作動油分析・交換","R4-MK",vendor=MAKER,start=D(2022,10,26),end=D(2022,10,26))

REP_MIDS = sorted(set(w.mid for w in WORKS))

# ---------------- 長寿命化改良工事 H27–28 ----------------
EN_SCOPE = ["乾燥火格子","燃焼火格子","後燃焼火格子","ガス冷却室","誘引通風機","燃焼制御盤","炉別計装盤"]
for f, fy in ((1, 2015), (2, 2016)):
    for nm in EN_SCOPE:
        if any(w.mid == m(nm, f) and w.fy == fy for w in WORKS): continue
        WORKS.append(Work(m(nm, f), fy, "REPLACE", {"ガス冷却室":"内面耐火物・噴射ノズル更新（長寿命化）"}.get(nm, "更新（長寿命化）"),
                          "H27-28-EN", vendor=MAKER))
WORKS.append(Work(m("中央監視操作盤"), 2016, "REPLACE", "中央監視盤更新（長寿命化）", "H27-28-EN", vendor=MAKER))
WORKS.append(Work(m("データ処理装置"), 2016, "REPLACE", "データロガ更新（長寿命化）", "H27-28-EN", vendor=MAKER))

# ---------------- 見本年度の手作り作業 ----------------
# H26（2014） メーカー委託で行った整備
for (nm, f, d, desc) in [
    ("燃焼火格子",1,D(2014,10,17),"火格子片 16枚取替"),
    ("燃焼火格子",2,D(2014,10,27),"火格子片 20枚取替"),
    ("押込送風機",2,D(2014,10,28),"軸受グリース交換・芯出し"),
    ("BFダスト排出装置",1,D(2014,10,20),"ロータ・シール取替"),
]:
    WORKS.append(Work(m(nm, f), 2014, "OVERHAUL", desc, "H26-MK", start=d, end=d, vendor=MAKER))
WORKS.append(Work(m("ガス冷却室",1), 2014, "INSPECT", "内面耐火物 剥離・脱落 確認（長寿命化で更新予定）", "H26-MK", start=D(2014,10,16), end=D(2014,10,16), vendor=MAKER))
WORKS.append(Work(m("ガス冷却室",2), 2014, "INSPECT", "内面耐火物 剥離 確認", "H26-MK", start=D(2014,10,27), end=D(2014,10,27), vendor=MAKER))
WORKS.append(Work(m("誘引通風機",1), 2014, "INSPECT", "振動 6.8mm/s 羽根車摩耗", "H26-MK", start=D(2014,10,21), end=D(2014,10,21), vendor=MAKER))
WORKS.append(Work(m("ごみクレーン"), 2014, "INSPECT", "年次自主検査", "H26-MK", start=D(2014,10,23), end=D(2014,10,23), vendor=MAKER))
# H26 地元補修
WORKS += [
 Work(m("灰出しコンベヤNo.2"),2014,"OVERHAUL","スクレーパ8枚取替","H26-地元1",vendor=LOCAL),
 Work(m("プラットホーム投入扉No.1"),2014,"OVERHAUL","油圧シリンダパッキン取替","H26-地元2",vendor=LOCAL),
 Work(m("事務室空調機"),2014,"OVERHAUL","圧縮機取替","H26-地元3",vendor=SETSUBI),
]
# H29（2017）
WORKS += [
 Work(m("押込送風機",2),2017,"OVERHAUL","軸受取替","H29-地元1",vendor=LOCAL,start=D(2017,6,20),end=D(2017,6,20)),
 Work(m("灰出しコンベヤNo.2"),2017,"OVERHAUL","スクレーパ取替（10枚）","H29-地元2",vendor=LOCAL,start=D(2017,10,18),end=D(2017,10,18)),
 Work(m("混練機"),2017,"OVERHAUL","パドル肉盛補修","H29-地元3",vendor=LOCAL,start=D(2017,8,8),end=D(2017,8,10)),
 Work(m("事務室空調機"),2017,"REPLACE","空調機取替","H29-地元4",vendor=SETSUBI,start=D(2018,2,14),end=D(2018,2,14)),
 Work(m("燃焼火格子",2),2017,"OVERHAUL","火格子片 8枚取替（予備品）","H29-MK",vendor=MAKER),
 Work(m("給じん装置",1),2017,"OVERHAUL","プッシャ摺動板取替","H29-MK",vendor=MAKER),
]
# H30（2018）
WORKS += [
 Work(m("BF加温ヒータ",2),2018,"REPLACE","エレメント取替","H30-地元2",vendor=LOCAL),
 Work(m("汚水受槽ポンプB"),2018,"REPLACE","ポンプ取替","H30-地元3",vendor=SETSUBI),
]
# R4（2022） 町の修繕執行管理表と一致させる
WORKS += [
 Work(m("プラットホーム投入扉No.3"),2022,"OVERHAUL","油圧シリンダパッキン取替","R4-DR",vendor=LOCAL,start=D(2023,1,17),end=D(2023,1,17)),
 Work(m("洗車装置"),2022,"REPLACE","高圧洗浄ポンプ取替","R4-WA",vendor=SETSUBI,start=D(2022,8,23),end=D(2022,8,23)),
 Work(m("浄化槽ブロワA"),2022,"REPLACE","ブロワ取替","R4-JO",vendor=SETSUBI,start=D(2023,3,7),end=D(2023,3,7)),
 Work(m("火格子駆動用油圧ユニット",1),2022,"INSPECT","作動油分析・交換","R4-MK",vendor=MAKER,start=D(2022,10,17),end=D(2022,10,17)),
 Work(m("燃焼火格子",2),2022,"INSPECT","摩耗測定（交換不要）","R4-MK",vendor=MAKER,start=D(2022,10,22),end=D(2022,10,22)),
 Work(m("押込送風機",1),2022,"INSPECT","振動測定・軸受グリース交換","R4-MK",vendor=MAKER,start=D(2022,10,17),end=D(2022,10,17)),
 Work(m("押込送風機",2),2022,"INSPECT","振動測定・軸受グリース交換","R4-MK",vendor=MAKER,start=D(2022,10,25),end=D(2022,10,25)),
 Work(m("誘引通風機",2),2022,"INSPECT","振動測定 1.8mm/s","R4-MK",vendor=MAKER,start=D(2022,10,25),end=D(2022,10,25)),
 Work(m("ガス冷却室",1),2022,"INSPECT","内部点検・噴射ノズル点検（ノズルチップ2本取替）","R4-MK",vendor=MAKER,start=D(2022,10,15),end=D(2022,10,15)),
 Work(m("灰押出機",1),2022,"OVERHAUL","プッシャ摺動部ライナ取替","R4-MK",vendor=MAKER,start=D(2022,10,14),end=D(2022,10,14)),
 Work(m("燃焼制御盤",1),2022,"INSPECT","計装ループ確認・調節計校正","R4-MK",vendor=MAKER,start=D(2022,10,17),end=D(2022,10,17)),
 Work(m("燃焼制御盤",2),2022,"INSPECT","計装ループ確認・調節計校正","R4-MK",vendor=MAKER,start=D(2022,10,26),end=D(2022,10,26)),
 Work(m("中央監視操作盤"),2022,"INSPECT","停電時点検・清掃","R4-MK",vendor=MAKER,start=D(2022,10,20),end=D(2022,10,20)),
 Work(m("BFパルス用空気槽",1),2022,"INSPECT","第二種圧力容器 定期自主検査","R4-MK",vendor=MAKER,start=D(2022,10,15),end=D(2022,10,15)),
 Work(m("BFパルス用空気槽",2),2022,"INSPECT","第二種圧力容器 定期自主検査","R4-MK",vendor=MAKER,start=D(2022,10,22),end=D(2022,10,22)),
 Work(m("空気槽No.1"),2022,"INSPECT","第二種圧力容器 定期自主検査","R4-MK",vendor=MAKER,start=D(2022,10,19),end=D(2022,10,19)),
 Work(m("空気槽No.2"),2022,"INSPECT","第二種圧力容器 定期自主検査","R4-MK",vendor=MAKER,start=D(2022,10,19),end=D(2022,10,19)),
 Work(m("保守用ホイスト"),2022,"INSPECT","年次自主検査","R4-MK",vendor=MAKER,start=D(2022,10,19),end=D(2022,10,19)),
 Work(m("高圧受電盤"),2022,"INSPECT","年次点検（停電）","R4-保安",vendor=ELEC,start=D(2022,10,20),end=D(2022,10,20)),
]

WORKS += [
 Work(m("灰出しコンベヤNo.2"),2015,"OVERHAUL","テール軸受取替","H27-地元1",vendor=LOCAL),
 Work(m("井水揚水ポンプNo.2"),2015,"REPLACE","水中ポンプ取替","H27-地元2",vendor=LOCAL),
 Work(m("雑排水ポンプA"),2015,"REPLACE","ポンプ取替","H27-地元3",vendor=SETSUBI),
 Work(m("プラント用水ポンプA"),2016,"OVERHAUL","分解整備（メカシール・軸受）","H28-地元1",vendor=LOCAL),
 Work(m("落じんコンベヤ",1),2016,"OVERHAUL","スクリュー肉盛補修","H28-地元2",vendor=LOCAL),
 Work(m("電気室空調機"),2016,"OVERHAUL","冷媒漏れ修理・ガス補充","H28-地元3",vendor=SETSUBI),
]
FIXED_YEARS = {2014, 2015, 2016, 2017, 2018, 2022}

# ---------------- 2014 コンサル計画（当初計画） ----------------
# 記号：● 更新 ◎ 整備（オーバーホール） ○ 補修 △ 点検・調査
LIFE = {"pump":(22,8),"fan":(25,10),"idf":(25,8),"conv":(20,7),"ashconv":(22,8),"hyd":(25,9),"panel":(25,None),
        "struct":(None,10),"grate":(None,None),"furnace":(None,5),"burner":(25,None),"itv":(15,None),"damper":(25,None),
        "hx":(None,10),"cooler":(None,None),"comp":(20,7),"misc":(25,None),"rv":(20,6),"bf":(None,6),"heater":(20,None),
        "tank":(None,None),"tank_s":(None,None),"feeder":(25,8),"blower":(20,7),"pusher":(None,7),"scale":(None,None),
        "pc":(10,None),"door":(None,8),"crane":(None,6),"shear":(None,None),"mixer":(None,6),"mixer_s":(25,None),
        "bf_s":(None,None),"hv":(30,None),"tr":(None,None),"gen":(None,10),"dc":(15,None),"ups":(10,None),"hoist":(None,None),
        "ac":(18,None),"fire":(25,None),"elev":(None,None)}

def consultant_plan():
    """2014 年度末（H27.3）にコンサルが作った整備スケジュール。H26〜H45。
    2014 時点の知識だけで作る（後年の故障は知らない）。"""
    rnd = random.Random(SEED + 1)
    plan = {}  # mid -> {fy: (sym, cost千円)}
    for e in EQ:
        p = {}
        if e.added_2015:
            plan[e.mid] = p; continue
        life, oh = LIFE[e.cls]
        hl = rnd.choice([2, 3, 3, 3, 4])  # 2014 調査の健全度（4=健全 1=劣化大）
        if e.cls == "panel" and e.imp != "A": hl = rnd.choice([3, 3, 4])
        e.health2014 = hl
        if e.name in EN_SCOPE:
            p[2015 if e.furnace == 1 else 2016] = ("●", None)
            e.health2014 = 2
        elif e.cls == "bf":
            p[2015 if e.furnace == 1 else 2016] = ("◎", None)  # ろ布（長寿命化）
            p[2021 if e.furnace == 1 else 2022] = ("◎", 6800)
            p[2027 if e.furnace == 1 else 2028] = ("◎", 7100)
        elif e.name in ("中央監視操作盤", "データ処理装置"):
            p[2016] = ("●", None)
        else:
            if life and e.imp != "C":
                y = e.inst + life + (hl - 3) * 2
                y = max(y, 2017 + rnd.randint(0, 4))
                if y <= 2033: p[y] = ("●", None)
            if oh and e.imp != "C":
                start = 2015 + rnd.randint(0, oh - 1)
                y = start
                while y <= 2033:
                    if y not in p and not any(s == "●" and abs(yy - y) < 3 for yy, (s, _) in p.items()):
                        p[y] = ("◎", None)
                    y += oh
            if e.cls in ("struct", "hx", "furnace") and e.imp in ("A", "B"):
                y = 2017 + rnd.randint(0, 2)
                while y <= 2033:
                    p.setdefault(y, ("△", None)); y += 5
        if e.name == "灰出しコンベヤNo.1":
            p = {2018: ("●", None), 2027: ("◎", None)}  # チェーン取替（コンサルは更新扱い）
        if e.name == "ごみクレーン":
            p = {2017: ("◎", None), 2020: ("◎", None), 2023: ("◎", None), 2026: ("●", None)}
        if e.name in ("ごみクレーン操作盤","ごみクレーン動力盤"):
            p = {2026: ("●", None)}
        if e.name == "高圧受電盤":
            p = {2019: ("●", None)}
        plan[e.mid] = p
    return plan

COST_UNIT = {  # 千円。コンサルの概算単価
 "pump":(1200,350),"fan":(2600,600),"idf":(18000,1800),"conv":(3500,800),"ashconv":(9800,2200),"hyd":(4200,900),
 "panel":(6500,0),"struct":(0,1500),"furnace":(0,4000),"burner":(2800,0),"itv":(900,0),"damper":(1500,0),"hx":(0,1200),
 "comp":(3800,700),"misc":(1500,300),"rv":(1100,300),"bf":(0,6800),"heater":(1600,0),"feeder":(2200,500),"blower":(1500,400),
 "pusher":(0,1800),"pc":(1800,0),"door":(0,700),"crane":(65000,3500),"mixer":(0,1600),"mixer_s":(500,0),"hv":(9500,0),
 "gen":(0,1300),"dc":(2600,0),"ups":(1500,0),"ac":(1800,0),"fire":(3500,0),"tr":(0,0),"tank":(0,0),"tank_s":(0,0),
 "grate":(0,0),"cooler":(0,0),"scale":(0,0),"shear":(0,0),"bf_s":(0,0),"hoist":(0,0),"elev":(0,0),
}
EN_COST = {"乾燥火格子":14500,"燃焼火格子":26800,"後燃焼火格子":12800,"ガス冷却室":21500,"誘引通風機":19800,
           "燃焼制御盤":16200,"炉別計装盤":7400,"中央監視操作盤":28500,"データ処理装置":4800}

def plan_cost(e, sym):
    if e.name in EN_COST and sym == "●": return EN_COST[e.name]
    if e.name == "バグフィルタ" and sym == "◎": return 6500
    if e.name == "灰出しコンベヤNo.1": return 9800 if sym == "●" else 2400
    if e.name in ("ごみクレーン操作盤","ごみクレーン動力盤"): return 14000
    if e.name == "高圧受電盤": return 9500
    if sym == "△": return 300
    a, b = COST_UNIT.get(e.cls, (0, 0))
    return a if sym == "●" else b

CPLAN = consultant_plan()

# ---------------- その他の年の実績（ランダム、種固定） ----------------

OH_TXT = {"pump":"分解整備（メカシール・軸受）","fan":"軸受取替・芯出し","idf":"軸受取替","conv":"スクリュー・軸受補修","ashconv":"スクレーパ・チェーン補修",
 "hyd":"油圧ポンプ・シール取替","struct":"腐食部 当て板補修","furnace":"耐火物 部分補修","hx":"伝熱管 部分補修","comp":"オーバーホール",
 "rv":"ロータ・シール取替","feeder":"駆動部整備","blower":"オーバーホール","pusher":"摺動部ライナ取替","door":"油圧シリンダ整備",
 "mixer":"パドル肉盛補修","gen":"内部点検整備","misc":"部品取替","damper":"駆動部整備","panel":"部品取替","itv":"カメラ取替"}
REP_TXT = {"panel":"盤内機器更新","itv":"カメラ更新","pc":"装置更新","ac":"機器取替"}
def wdesc(e, kind, tag=""):
    if kind == "REPLACE": t = REP_TXT.get(e.cls, "機器更新")
    elif kind == "INSPECT": t = "点検・調査"
    else: t = OH_TXT.get(e.cls, "補修")
    return t + tag

FIXED_LOCAL = {}
def _count_fixed_local():
    out = {}
    for w in WORKS:
        k = w.contract
        if (w.vendor in (LOCAL, SETSUBI, CRANE, ELEC_W) and not k.endswith("-MK") and "計量" not in k and "性能検査" not in k) or k in ("H30-計量", "R7-計量"):
            out.setdefault(w.fy, set()).add(k)
    return out

LOCAL_TARGET_RNG = random.Random(SEED + 7)
LOCAL_TARGET = {fy: LOCAL_TARGET_RNG.randint(3, 6) for fy in range(2012, 2027)}
LOCAL_TARGET[2026] = 2  # 上期だけ

def vendor_for(e):
    if e.sys == "Z": return SETSUBI
    if e.cls in ("hv", "tr", "dc", "ups") or (e.cls == "panel" and e.furnace == 0): return ELEC_W
    return LOCAL

def assign_year(fy, cands, rnd, mk_cap=5):
    """cands: [(e, kind, desc)]。炉の整備・点検はメーカー定修へ、更新と共通機器は地元の修繕契約へ（年の件数に収める）"""
    out = []
    fixed = _count_fixed_local().get(fy, set())
    slots = max(0, LOCAL_TARGET[fy] - len(fixed))
    SIMPLE = ("pump", "fan", "blower", "comp", "misc", "rv", "conv", "heater", "damper", "itv", "door", "ac", "fire", "panel", "pc", "dc", "ups")
    cands = [c for c in cands if not (c[1] == "REPLACE" and c[0].cls not in SIMPLE)]  # 大物の更新は単独契約が要るので見送り
    recent = {}
    for w in WORKS + DONE_RAND:
        if w.kind in ("REPLACE", "OVERHAUL") and w.fy < fy: recent[w.mid] = max(recent.get(w.mid, 0), w.fy)
    cands = [c for c in cands if fy - recent.get(c[0].mid, 0) >= 5]  # 5 年以内に手を入れた機器は外す
    mk = [c for c in cands if c[0].furnace and c[1] in ("OVERHAUL", "INSPECT")]
    loc = [c for c in cands if c not in mk]
    for (e, kind, desc) in mk[:mk_cap]:
        out.append(Work(e.mid, fy, kind, desc, "%s-MK" % wareki(fy), vendor=MAKER, src="rand"))
    seen = set(); loc2 = []
    for c in loc:
        if c[0].mid not in seen: seen.add(c[0].mid); loc2.append(c)
    loc = loc2
    loc.sort(key=lambda c: ({"A": 0, "B": 1, "C": 2}[c[0].imp], c[0].idx))
    groups = []
    for c in loc:
        v = vendor_for(c[0])
        placed = False
        for g in groups:
            if g["v"] == v and len(g["w"]) < 2 and g["cls"] == c[0].cls and g["kind"] == c[1]:
                g["w"].append(c); placed = True; break
        if not placed:
            if len(groups) < slots:
                groups.append({"v": v, "cls": c[0].cls, "kind": c[1], "w": [c]})
    for i, g in enumerate(groups, 1):
        for (e, kind, desc) in g["w"]:
            out.append(Work(e.mid, fy, kind, desc, "%s-L%d" % (wareki(fy), i), vendor=g["v"], src="rand"))
    return out

RAND_YEARS_A = [2012, 2013, 2019, 2020, 2021]
DONE_RAND = []
def gen_other_works():
    rnd = random.Random(SEED + 2)
    out = []
    have = {(w.mid, w.fy) for w in WORKS}
    for fy in RAND_YEARS_A:
        cands = []
        for e in EQ:
            if e.mid in REP_MIDS or e.cls in ("ashconv", "crane") or (e.mid, fy) in have: continue
            p = CPLAN[e.mid]
            sym = p.get(fy, (None,))[0]
            if sym and fy >= 2014:
                p_do = {2019: .55, 2020: .45, 2021: .5}.get(fy, .7) + (.15 if e.imp == "A" else 0)
                if sym == "●": p_do = .3 if e.imp == "A" else .15
                if rnd.random() < p_do:
                    kind = {"●": "REPLACE", "◎": "OVERHAUL", "○": "OVERHAUL", "△": "INSPECT"}[sym]
                    cands.append((e, kind, wdesc(e, kind)))
            # 計画外の故障修理
            if e.cls in ("pump", "conv", "fan", "hyd", "rv", "misc", "door") and rnd.random() < (.08 if fy < 2014 else .035):
                cands.append((e, "OVERHAUL", "故障修理（" + wdesc(e, "OVERHAUL") + "）"))
        rnd.shuffle(cands)
        got = assign_year(fy, cands, rnd)
        out += got; DONE_RAND.extend(got)
    return out

WORKS += gen_other_works()

# ---------------- t3 の計画（2023 作成、以後毎年見直し） ----------------
# 状態から選ぶ。重要度 C は原則 事後保全（計画に載せない）。廃止 2033 まで残り年数を見る。
def last_major(mid, upto):
    ws = [w for w in WORKS if w.mid == mid and w.fy <= upto and w.kind in ("REPLACE", "OVERHAUL", "UPGRADE")]
    return max(ws, key=lambda w: w.fy) if ws else None

def t3_plan():
    rnd = random.Random(SEED + 3)
    plan = {}
    for e in EQ:
        p = {}
        lm = last_major(e.mid, 2022)
        base = lm.fy if lm else e.inst
        life, oh = LIFE[e.cls]
        age = 2023 - base
        if e.imp == "C":
            plan[e.mid] = p; continue
        if life and e.inst + life <= 2029 and (lm is None or lm.kind != "REPLACE") and e.imp == "A":
            y = max(2024, min(2031, e.inst + life + rnd.randint(-1, 3)))
            if y <= 2033 - 3: p[y] = "更"
        upd = [y for y, v in p.items() if v == "更"]
        if upd and min(upd) <= 2027:
            oh = None  # 近く更新するものは整備を入れない
        if oh and age >= oh - 1:
            y = 2023 + rnd.randint(0, 3)
            p.setdefault(y, "整")
            if y + oh <= 2032 and e.imp == "A": p.setdefault(y + oh, "整")
        elif oh and e.imp == "A":
            y = base + oh + rnd.randint(0, 2)
            if 2023 <= y <= 2032: p.setdefault(y, "整")
        if e.cls in ("struct", "furnace", "hx") and e.imp in ("A","B") and rnd.random() < .6:
            p.setdefault(2023 + rnd.randint(1, 5), "点")
        plan[e.mid] = p
    # 代表機器（根拠記録どおり）
    def S(name, f, d): plan[m(name, f)] = d
    S("ごみクレーン",0,{2025:"点",2027:"整"})
    S("ごみクレーン操作盤",0,{2024:"更"}); S("ごみクレーン動力盤",0,{2024:"更"})
    S("灰出しコンベヤNo.1",0,{2025:"点",2028:"補"})
    S("燃焼火格子",1,{2024:"補",2027:"補"}); S("燃焼火格子",2,{2025:"補",2028:"補"})
    S("ガス冷却室",2,{2023:"補",2026:"更"}); S("ガス冷却室",1,{2025:"点",2027:"補"})
    S("バグフィルタ",1,{2028:"整"}); S("バグフィルタ",2,{2023:"整",2029:"整"})
    S("誘引通風機",1,{2026:"整"}); S("誘引通風機",2,{2027:"整"})
    S("噴射水ポンプA",1,{2024:"更"}); S("噴射水ポンプB",1,{2024:"更"})
    S("噴射水ポンプA",2,{2025:"更"}); S("噴射水ポンプB",2,{2025:"更"})
    S("計量機",0,{}); S("計量データ処理装置",0,{2025:"更"})
    S("非常用発電機",0,{2027:"整"})
    S("空気圧縮機No.2",0,{2025:"更"}); S("空気圧縮機No.1",0,{2027:"整"})
    S("高圧受電盤",0,{2027:"補"})
    S("火格子駆動用油圧ユニット",2,{2026:"整"}); S("火格子駆動用油圧ユニット",1,{2025:"整"})
    return plan

T3PLAN = t3_plan()
T3_COST = {"更": None, "整": None, "補": None, "点": 300}

def t3_actuals():
    """2023–2026上期の実績。代表機器は手書きの WORKS。その他は t3 計画から（年の件数に収め、残りは繰延べ）。"""
    rnd = random.Random(SEED + 4)
    out = []
    for fy in (2023, 2024, 2025):
        cands = []
        have = {(w.mid, w.fy) for w in WORKS}
        for e in EQ:
            if e.mid in REP_MIDS or (e.mid, fy) in have: continue
            s_ = T3PLAN[e.mid].get(fy)
            if s_ and rnd.random() < .85:
                kind = {"更": "REPLACE", "整": "OVERHAUL", "補": "OVERHAUL", "点": "INSPECT"}[s_]
                cands.append((e, kind, wdesc(e, kind)))
        got = assign_year(fy, cands, rnd)
        out += got; DONE_RAND.extend(got)
    return out

WORKS += t3_actuals()

def works_of(mid):
    return sorted([w for w in WORKS if w.mid == mid], key=lambda w: (w.fy, w.start or D(w.fy, 4, 1)))

# ---- 法定点検（t3 の計画に載る。周期の確度は notes） ----
LEGAL = [
 ("受変電設備 年次点検（停電）", "毎年", "電気事業法・保安規程", lambda y: True),
 ("受変電設備 月次点検", "毎月", "電気事業法・保安規程", lambda y: True),
 ("非常用発電機 点検（機器・総合）", "6か月・1年", "消防法／電気事業法", lambda y: True),
 ("ごみクレーン 性能検査", "2年", "クレーン則", lambda y: y % 2 == 1),
 ("ごみクレーン 年次自主検査", "毎年", "クレーン則", lambda y: True),
 ("保守用ホイスト 年次自主検査", "毎年", "クレーン則", lambda y: True),
 ("第二種圧力容器 定期自主検査（空気槽4基）", "毎年", "ボイラー及び圧力容器安全規則", lambda y: True),
 ("排ガス測定（ばいじん・NOx・SOx・HCl）", "年2回", "大気汚染防止法", lambda y: True),
 ("排ガス測定（ダイオキシン類）", "年1回", "ダイオキシン類対策特別措置法", lambda y: True),
 ("ダイオキシン類 作業環境測定", "6か月", "労働安全衛生規則", lambda y: True),
 ("計量機 定期検査", "2年", "計量法", lambda y: y % 2 == 0),
 ("消防用設備等点検", "6か月（報告3年）", "消防法", lambda y: True),
 ("精密機能検査", "3年", "廃棄物処理法施行規則", lambda y: y in (2025,)),  # t3 は次を R11 と書いた（周期超え：わざと）
 ("エレベーター 定期検査", "毎年", "建築基準法", lambda y: True),
 ("浄化槽 法定検査（11条）", "毎年", "浄化槽法", lambda y: True),
]
