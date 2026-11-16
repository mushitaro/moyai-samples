# -*- coding: utf-8 -*-
"""F10【架空】あかね台資源化センター：正解の世界（機器・状態・履歴・契約・実施作業）。
帳票はここから書き分ける。乱数の種と物語の「今日」は固定。"""
import random, datetime as dt

SEED = 1010
TODAY = dt.date(2026, 11, 13)
FAC = '【架空】あかね台資源化センター'
FACN = 'あかね台資源化センター'
CITY = 'あかね台市'
MAKER = '千草環境システム株式会社'
rng = random.Random(SEED)

def wareki(y):
    if y >= 2019: return '令和元' if y == 2019 else f'令和{y-2018}'
    return f'平成{y-1988}'
def wa_short(y):
    return f'R{y-2018}' if y >= 2019 else f'H{y-1988}'

# ------------------------------------------------------------------
# 機器。系統図番号（メーカー完成図書の体系）：炉-系統-設備-機器[-付属][-M]
#   炉 0=共通 1=1号 2=2号。電動機は親機器の下に -M1,-M2 で 1 台ずつ数える。
# ------------------------------------------------------------------
UNITS = []
_ser = [0]
def _serial(prefix='C'):
    _ser[0] += 1
    return f'{prefix}{rng.randint(10,99)}-{_ser[0]:05d}'

MAKERS_BY_KIND = {
    '機器': ['千草環境システム'], '電動機': ['【架空】東和電機', '【架空】三栄モータ', '【架空】東和電機'],
    '電動弁': ['【架空】大洋バルブ', '【架空】北斗弁工業'], '調節弁': ['【架空】北斗弁工業', '【架空】大洋バルブ'],
    '計装': ['【架空】北辰計測器', '【架空】光陽計装'], '盤': ['【架空】あかね電設', '【架空】東和電機'],
}

def motor_spec(kw):
    volt = 3300 if kw >= 200 else 400
    pole = 4 if kw < 150 else 6
    enc = '全閉外扇' if kw < 200 else '全閉内冷'
    return volt, pole, enc

class U(dict):
    pass

def add(furn, sys, sysname, sub, subname, name, kind='機器', spec='', kw=None, n=1,
        duty=None, parent=None, maker=None, year=2006, cat='static', imp=None, short=None,
        nmot=1, ident=None):
    """台数 n の同型機を 1 台ずつ登録。kw があれば電動機を子として nmot 台ずつ。"""
    out = []
    for i in range(n):
        u = U()
        u['furn'] = furn; u['sys'] = sys; u['sysname'] = sysname; u['sub'] = sub; u['subname'] = subname
        suffix = '' if n == 1 else ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'][i]
        nm = name if n == 1 else f'{name}（{suffix}）' if n <= 10 else f'{name}（{i+1}）'
        if furn in (1, 2) and not name.startswith(('1号', '2号')):
            disp = f'{furn}号{nm}'
        else:
            disp = nm
        u['name'] = disp; u['base'] = name; u['kind'] = kind; u['spec'] = spec; u['kw'] = None
        if duty is None:
            u['duty'] = '' if n == 1 else ('常用' if n > 1 and i < n - (1 if n == 2 or n == 3 else 0) else '予備')
        else:
            u['duty'] = duty[i] if isinstance(duty, (list, tuple)) else duty
        u['_par'] = parent; u['parent'] = None
        u['year'] = year; u['cat'] = cat; u['imp'] = imp; u['short'] = short
        u['maker'] = maker or rng.choice(MAKERS_BY_KIND.get(kind, ['千草環境システム']))
        if kind == '機器' and maker is None: u['maker'] = '千草環境システム'
        u['serial'] = _serial('S' if kind == '機器' else kind[0] if kind != '計装' else 'K')
        u['ident'] = ident
        UNITS.append(u); out.append(u)
        if kw:
            for k in range(nmot):
                m = U(u)
                m['kind'] = '電動機'; m['parent'] = '__pending__'; m['_p'] = u; m['_par'] = None
                v, p, enc = motor_spec(kw)
                m['name'] = disp + ('用電動機' if nmot == 1 else f'用電動機（{k+1}）')
                m['base'] = name + '用電動機'
                m['spec'] = f'{enc} {p}P {v}V'
                m['kw'] = kw; m['volt'] = v; m['pole'] = p
                m['maker'] = rng.choice(MAKERS_BY_KIND['電動機'])
                m['serial'] = _serial('M'); m['cat'] = 'motor'; m['short'] = None; m['ident'] = None
                UNITS.append(m)
    return out[0] if n == 1 else out

def furnace(f):
    F = f
    S = lambda *a, **k: add(F, *a, **k)
    # 10 給じん・副資材投入
    s, sn = '10', '給じん設備'
    S(s, sn, '1', 'ごみ供給装置', 'ごみ投入ホッパ', spec='容量 25m³ 鋼板製', cat='static')
    S(s, sn, '1', 'ごみ供給装置', 'ホッパゲート', spec='油圧開閉式 1.8m×1.8m', cat='mech', imp='B')
    S(s, sn, '1', 'ごみ供給装置', '上部シールダンパ', spec='油圧揺動式', cat='mech', imp='A')
    S(s, sn, '1', 'ごみ供給装置', '下部シールダンパ', spec='油圧揺動式', cat='mech', imp='A')
    hp = S(s, sn, '1', 'ごみ供給装置', '給じん油圧ユニット', spec='14MPa タンク 400L', cat='mech', imp='A')
    S(s, sn, '1', 'ごみ供給装置', '給じん油圧ポンプ', spec='可変容量形 80L/min', kw=11, n=2, parent=hp, cat='rot', imp='B')
    S(s, sn, '2', '副資材投入装置', 'コークス計量ホッパ', spec='ロードセル式 2.0m³', cat='static')
    S(s, sn, '2', '副資材投入装置', '石灰石計量ホッパ', spec='ロードセル式 1.2m³', cat='static')
    S(s, sn, '2', '副資材投入装置', 'コークス切出しフィーダ', spec='電磁振動式 3t/h', kw=1.5, cat='rot', imp='A')
    S(s, sn, '2', '副資材投入装置', '石灰石切出しフィーダ', spec='電磁振動式 2t/h', kw=1.5, cat='rot', imp='A')
    S(s, sn, '2', '副資材投入装置', '炉頂レベル計', kind='計装', spec='マイクロ波式', cat='instr')
    # 20 溶融炉
    s, sn = '20', '溶融炉設備'
    mf = S(s, sn, '1', '溶融炉本体', '溶融炉本体', spec='シャフト炉式 120t/24h 内径 2.6m（羽口部）', cat='furnace', imp='A', short='溶融炉')
    S(s, sn, '1', '溶融炉本体', '炉体冷却ジャケット', spec='水冷ジャケット 朝顔部・炉床部', cat='furnace', imp='A')
    S(s, sn, '1', '溶融炉本体', '主羽口装置', spec='水冷銅羽口 8本', cat='tuyere', imp='A', short='主羽口')
    S(s, sn, '1', '溶融炉本体', '副羽口装置', spec='水冷羽口 8本', cat='tuyere', imp='A', short='副羽口')
    S(s, sn, '1', '溶融炉本体', '炉頂ガス圧力調節ダンパ', kind='調節弁', spec='バタフライ形 電動', cat='valve', imp='A')
    S(s, sn, '1', '溶融炉本体', '炉頂温度計', kind='計装', spec='熱電対 K 3点', cat='instr')
    S(s, sn, '1', '溶融炉本体', '炉体温度計', kind='計装', spec='熱電対 K 24点（炉体表面）', cat='instr')
    S(s, sn, '1', '溶融炉本体', '炉内圧力計', kind='計装', spec='差圧伝送器', cat='instr')
    S(s, sn, '1', '溶融炉本体', '炉内監視ITV', kind='計装', spec='水冷ケーシング付', cat='instr')
    S(s, sn, '2', '出湯設備', '出湯口', spec='SiC質れんが・水冷スリーブ', cat='taphole', imp='A', short='出湯口')
    S(s, sn, '2', '出湯設備', '開孔機', spec='電動回転打撃式', kw=7.5, cat='rot', imp='A')
    S(s, sn, '2', '出湯設備', 'マッドガン', spec='油圧式 吐出量 0.15m³', cat='mech', imp='A')
    ht = S(s, sn, '2', '出湯設備', '出湯口油圧ユニット', spec='21MPa', cat='mech', imp='A')
    S(s, sn, '2', '出湯設備', '出湯口油圧ポンプ', spec='ピストン式 40L/min', kw=15, n=2, parent=ht, cat='rot', imp='B')
    S(s, sn, '2', '出湯設備', '出湯樋', spec='キャスタブル内張 L=4.5m', cat='static')
    S(s, sn, '2', '出湯設備', '出湯樋予熱バーナ', spec='灯油 30L/h', cat='mech')
    S(s, sn, '2', '出湯設備', '出湯口集じんフード', spec='鋼板製', cat='static')
    S(s, sn, '2', '出湯設備', '出湯口集じんファン', spec='ターボ形 300m³/min', kw=30, cat='rot', imp='B')
    S(s, sn, '2', '出湯設備', '出湯温度計', kind='計装', spec='二色放射温度計', cat='instr')
    S(s, sn, '3', '送風設備', '羽口送風機', spec='多段ターボ 9,000m³N/h×35kPa', kw=160, n=2, cat='rot', imp='B')
    S(s, sn, '3', '送風設備', '送風空気予熱器', spec='蒸気式 200℃', cat='static', imp='B')
    S(s, sn, '3', '送風設備', '送風流量調節弁', kind='調節弁', spec='バタフライ形 空気作動', cat='valve', imp='A')
    S(s, sn, '3', '送風設備', '酸素富化流量調節弁', kind='調節弁', spec='玉形 酸素用（禁油）', cat='valve', imp='A')
    S(s, sn, '3', '送風設備', '酸素遮断弁', kind='電動弁', spec='ボール形 酸素用', cat='valve', imp='A')
    S(s, sn, '3', '送風設備', '送風流量計', kind='計装', spec='ベンチュリ＋差圧', cat='instr')
    S(s, sn, '3', '送風設備', '送風酸素濃度計', kind='計装', spec='ジルコニア式', cat='instr')
    S(s, sn, '4', '炉体冷却水設備', '炉体冷却水ポンプ', spec='渦巻 120m³/h×45m', kw=30, n=2, cat='rot', imp='B')
    S(s, sn, '4', '炉体冷却水設備', '羽口冷却水ポンプ', spec='多段渦巻 60m³/h×90m', kw=37, n=2, cat='rot', imp='B')
    S(s, sn, '4', '炉体冷却水設備', '羽口冷却水熱交換器', spec='プレート式', cat='static', imp='B')
    S(s, sn, '4', '炉体冷却水設備', '炉体冷却塔', spec='開放式 150m³/h', kw=5.5, cat='rot', imp='B')
    S(s, sn, '4', '炉体冷却水設備', '炉体冷却水槽', spec='SUS 15m³', cat='static')
    S(s, sn, '4', '炉体冷却水設備', '冷却水流量計（羽口）', kind='計装', spec='電磁式', cat='instr')
    S(s, sn, '4', '炉体冷却水設備', '冷却水流量計（ジャケット）', kind='計装', spec='電磁式', cat='instr')
    # 30 スラグ水砕
    s, sn = '30', 'スラグ水砕設備'
    S(s, sn, '1', '水砕装置', '水砕樋', spec='SUS 水砕水噴射式', cat='static', imp='A')
    S(s, sn, '1', '水砕装置', '水砕槽', spec='鋼板製 30m³', cat='static', imp='A')
    S(s, sn, '1', '水砕装置', 'スラグ掻揚コンベヤ', spec='スクレーパ式 3t/h', kw=5.5, cat='rot', imp='A')
    S(s, sn, '1', '水砕装置', '水砕水循環ポンプ', spec='渦巻 200m³/h×25m', kw=30, n=2, cat='rot', imp='B')
    S(s, sn, '1', '水砕装置', '水砕水冷却塔', spec='開放式 200m³/h', kw=7.5, cat='rot', imp='B')
    S(s, sn, '1', '水砕装置', '水砕水自動ストレーナ', spec='逆洗式', kw=0.4, cat='rot')
    S(s, sn, '1', '水砕装置', '水砕ミスト排気ファン', spec='軸流 150m³/min', kw=3.7, cat='rot')
    S(s, sn, '1', '水砕装置', '水砕槽水位計', kind='計装', spec='投込式', cat='instr')
    S(s, sn, '1', '水砕装置', '水砕水温度計', kind='計装', spec='測温抵抗体', cat='instr')
    # 40 燃焼室
    s, sn = '40', '燃焼設備'
    S(s, sn, '1', '燃焼室', '燃焼室本体', spec='耐火物内張 容積 95m³', cat='furnace2', imp='A')
    S(s, sn, '1', '燃焼室', '燃焼用空気送風機', spec='ターボ 14,000m³N/h×6kPa', kw=45, cat='rot', imp='A')
    S(s, sn, '1', '燃焼室', '助燃バーナ', spec='灯油 150L/h', cat='mech', imp='B')
    S(s, sn, '1', '燃焼室', 'バーナ用送風機', spec='ターボ', kw=11, cat='rot')
    S(s, sn, '1', '燃焼室', '燃焼室ダスト排出装置', spec='ダブルダンパ', kw=0.75, cat='rot')
    S(s, sn, '1', '燃焼室', '燃焼空気流量調節ダンパ', kind='調節弁', spec='電動', cat='valve', imp='A')
    S(s, sn, '1', '燃焼室', '燃焼室出口温度計', kind='計装', spec='熱電対 R 2点', cat='instr')
    S(s, sn, '1', '燃焼室', '燃焼室圧力計', kind='計装', spec='差圧伝送器', cat='instr')
    S(s, sn, '1', '燃焼室', '燃焼室監視ITV', kind='計装', spec='水冷ケーシング付', cat='instr')
    # 50 ボイラ
    s, sn = '50', 'ボイラ設備'
    S(s, sn, '1', 'ボイラ本体', 'ボイラ本体', spec='自然循環水管 4MPa 400℃ 16t/h', cat='boiler', imp='A', short='ボイラ')
    S(s, sn, '1', 'ボイラ本体', '1次過熱器', spec='STB410 φ42.7', cat='sh', imp='A')
    S(s, sn, '1', 'ボイラ本体', '2次過熱器', spec='SUS310S/STBA24 φ42.7 t5.0', cat='sh2', imp='A', short='2次過熱器')
    S(s, sn, '1', 'ボイラ本体', '過熱器減温器', spec='スプレー式', cat='static', imp='A')
    S(s, sn, '1', 'ボイラ本体', 'エコノマイザ', spec='裸管 STB340', cat='sh', imp='A')
    S(s, sn, '1', 'ボイラ本体', 'ボイラ安全弁（ドラム）', kind='機器', spec='ばね式 吹出 4.4MPa', n=2, duty=['', ''], cat='safety', imp='A')
    S(s, sn, '1', 'ボイラ本体', 'ボイラ安全弁（過熱器出口）', kind='機器', spec='ばね式 吹出 4.2MPa', cat='safety', imp='A')
    S(s, sn, '2', 'スートブロワ', '長抜差式スートブロワ', spec='蒸気式 L=6m', kw=1.5, n=4, duty='', cat='rot')
    S(s, sn, '2', 'スートブロワ', '定置回転式スートブロワ', spec='蒸気式', kw=0.4, n=6, duty='', cat='rot')
    S(s, sn, '2', 'スートブロワ', 'スートブロワ現場盤', kind='盤', spec='自立形', cat='panel')
    S(s, sn, '3', 'ボイラダスト搬送', 'ボイラダストコンベヤ', spec='スクリュー式', kw=2.2, n=2, duty='', cat='rot')
    S(s, sn, '4', 'ボイラ給水・薬注', '清缶剤注入ポンプ', spec='ダイヤフラム式', kw=0.4, n=2, cat='rot')
    S(s, sn, '5', 'ボイラ計装・弁', '主蒸気止弁', kind='電動弁', spec='玉形 150A 4MPa', cat='valve', imp='A')
    S(s, sn, '5', 'ボイラ計装・弁', '給水調節弁', kind='調節弁', spec='玉形 50A', cat='valve', imp='A')
    S(s, sn, '5', 'ボイラ計装・弁', '減温水調節弁', kind='調節弁', spec='玉形 25A', cat='valve', imp='A')
    S(s, sn, '5', 'ボイラ計装・弁', '連続ブロー調節弁', kind='調節弁', spec='玉形 20A', cat='valve')
    S(s, sn, '5', 'ボイラ計装・弁', '缶底ブロー電動弁', kind='電動弁', spec='玉形 40A', cat='valve')
    S(s, sn, '5', 'ボイラ計装・弁', 'ドラム水位計（差圧式）', kind='計装', spec='差圧伝送器', cat='instr', imp='A')
    S(s, sn, '5', 'ボイラ計装・弁', 'ドラム水位計（二色式）', kind='計装', spec='二色水面計', cat='instr', imp='A')
    S(s, sn, '5', 'ボイラ計装・弁', '主蒸気流量計', kind='計装', spec='オリフィス＋差圧', cat='instr')
    S(s, sn, '5', 'ボイラ計装・弁', '主蒸気圧力計', kind='計装', spec='圧力伝送器', cat='instr')
    S(s, sn, '5', 'ボイラ計装・弁', '主蒸気温度計', kind='計装', spec='熱電対 K', cat='instr')
    S(s, sn, '5', 'ボイラ計装・弁', '給水流量計', kind='計装', spec='オリフィス＋差圧', cat='instr')
    # 60 減温塔
    s, sn = '60', '排ガス冷却設備'
    S(s, sn, '1', '減温塔', '減温塔本体', spec='鋼板製 φ3.2m H=18m', cat='static', imp='A')
    S(s, sn, '1', '減温塔', '減温塔噴霧ノズル', spec='二流体ノズル 6本', cat='nozzle', imp='A')
    S(s, sn, '1', '減温塔', '噴霧水ポンプ', spec='多段渦巻 6m³/h×150m', kw=7.5, n=2, cat='rot', imp='B')
    S(s, sn, '1', '減温塔', '噴霧水調節弁', kind='調節弁', spec='玉形 25A', cat='valve', imp='A')
    S(s, sn, '1', '減温塔', 'ノズル冷却空気ブロワ', spec='ルーツ', kw=3.7, cat='rot')
    S(s, sn, '1', '減温塔', '減温塔ダストコンベヤ', spec='スクリュー式', kw=2.2, cat='rot')
    S(s, sn, '1', '減温塔', '減温塔ロータリーバルブ', spec='φ300', kw=0.75, cat='rot')
    S(s, sn, '1', '減温塔', '減温塔出口温度計', kind='計装', spec='熱電対 K', cat='instr')
    S(s, sn, '1', '減温塔', '減温塔入口温度計', kind='計装', spec='熱電対 K', cat='instr')
    # 70 薬剤吹込
    s, sn = '70', '有害ガス除去設備'
    S(s, sn, '1', '薬剤吹込装置', '消石灰サービスホッパ', spec='1.5m³', cat='static')
    S(s, sn, '1', '薬剤吹込装置', '消石灰切出しロータリーバルブ', spec='φ150', kw=0.4, cat='rot')
    S(s, sn, '1', '薬剤吹込装置', '消石灰定量供給機', spec='テーブルフィーダ 30kg/h', kw=0.75, cat='rot', imp='A')
    S(s, sn, '1', '薬剤吹込装置', '消石灰吹込ブロワ', spec='ルーツ 3m³/min', kw=5.5, cat='rot', imp='A')
    S(s, sn, '1', '薬剤吹込装置', '活性炭サービスホッパ', spec='0.8m³', cat='static')
    S(s, sn, '1', '薬剤吹込装置', '活性炭切出しロータリーバルブ', spec='φ100', kw=0.4, cat='rot')
    S(s, sn, '1', '薬剤吹込装置', '活性炭定量供給機', spec='スクリューフィーダ 5kg/h', kw=0.4, cat='rot', imp='A')
    S(s, sn, '1', '薬剤吹込装置', '活性炭吹込ブロワ', spec='ルーツ 1.5m³/min', kw=3.7, cat='rot', imp='A')
    S(s, sn, '1', '薬剤吹込装置', '薬剤供給現場盤', kind='盤', spec='自立形', cat='panel')
    # 80 BF
    s, sn = '80', 'ろ過式集じん設備'
    S(s, sn, '1', 'ろ過式集じん器', 'ろ過式集じん器本体', spec='パルスジェット 6室 ろ布 PTFE 432本', cat='bf', imp='A', short='BF')
    S(s, sn, '1', 'ろ過式集じん器', 'BF逆洗制御盤', kind='盤', spec='自立形', cat='panel')
    S(s, sn, '1', 'ろ過式集じん器', 'BF逆洗空気槽', spec='1.0m³ 0.7MPa', cat='static')
    S(s, sn, '1', 'ろ過式集じん器', 'BF加温ヒータ', spec='電気式 60kW', cat='mech')
    S(s, sn, '1', 'ろ過式集じん器', 'BFダストスクリューコンベヤ', spec='φ300', kw=2.2, n=2, duty='', cat='rot')
    S(s, sn, '1', 'ろ過式集じん器', 'BFロータリーバルブ', spec='φ300', kw=0.75, n=2, duty='', cat='rot')
    S(s, sn, '1', 'ろ過式集じん器', '飛灰集合コンベヤ', spec='チェーン式', kw=3.7, cat='rot', imp='A')
    S(s, sn, '1', 'ろ過式集じん器', '飛灰搬送コンベヤ', spec='チェーン式（共通系へ）', kw=3.7, cat='rot', imp='A')
    S(s, sn, '1', 'ろ過式集じん器', 'BF入口ダンパ', kind='電動弁', spec='電動', cat='valve')
    S(s, sn, '1', 'ろ過式集じん器', 'BF出口ダンパ', kind='電動弁', spec='電動', cat='valve')
    S(s, sn, '1', 'ろ過式集じん器', 'BFバイパスダンパ', kind='電動弁', spec='電動', cat='valve', imp='A')
    S(s, sn, '1', 'ろ過式集じん器', 'BF差圧計', kind='計装', spec='差圧伝送器', cat='instr')
    S(s, sn, '1', 'ろ過式集じん器', 'BF入口温度計', kind='計装', spec='熱電対 K', cat='instr')
    # 90 脱硝
    s, sn = '90', '触媒脱硝設備'
    S(s, sn, '1', '触媒脱硝装置', '排ガス再加熱器', spec='蒸気式 フィン付管', cat='sh', imp='A')
    S(s, sn, '1', '触媒脱硝装置', '触媒反応塔', spec='ハニカム触媒 2層', cat='cat', imp='A', short='脱硝触媒')
    S(s, sn, '1', '触媒脱硝装置', 'アンモニア水注入ポンプ', spec='ダイヤフラム 10L/h', kw=0.4, n=2, cat='rot', imp='B')
    S(s, sn, '1', '触媒脱硝装置', 'アンモニア水流量調節弁', kind='調節弁', spec='小形', cat='valve')
    S(s, sn, '1', '触媒脱硝装置', '再加熱器蒸気調節弁', kind='調節弁', spec='玉形 50A', cat='valve')
    S(s, sn, '1', '触媒脱硝装置', '触媒入口温度計', kind='計装', spec='熱電対 K', cat='instr')
    S(s, sn, '1', '触媒脱硝装置', '触媒出口温度計', kind='計装', spec='熱電対 K', cat='instr')
    # A0 通風
    s, sn = 'A0', '通風設備'
    S(s, sn, '1', '誘引通風機', '誘引通風機', spec='片吸込ターボ 25,000m³N/h×7.5kPa', kw=132, cat='rot', imp='A', short='IDF')
    S(s, sn, '1', '誘引通風機', '誘引通風機インバータ盤', kind='盤', spec='400V 160kVA', cat='panel', imp='A')
    S(s, sn, '1', '誘引通風機', '誘引通風機入口ダンパ', kind='電動弁', spec='電動', cat='valve', imp='A')
    S(s, sn, '1', '誘引通風機', '誘引通風機軸受振動計', kind='計装', spec='速度形', cat='instr')
    S(s, sn, '1', '誘引通風機', '誘引通風機入口圧力計', kind='計装', spec='差圧伝送器', cat='instr')
    S(s, sn, '2', '煙道・煙突', '煙道', spec='鋼板製 保温付', cat='static')
    S(s, sn, '2', '煙道・煙突', '煙突内筒', spec='SUS 内径 0.9m', cat='static')
    S(s, sn, '3', '排ガス分析計', '排ガスO2計', kind='計装', spec='ジルコニア式', cat='analyzer')
    S(s, sn, '3', '排ガス分析計', '排ガスCO計', kind='計装', spec='赤外線式', cat='analyzer')
    S(s, sn, '3', '排ガス分析計', '排ガスHCl計', kind='計装', spec='イオン電極式', cat='analyzer')
    S(s, sn, '3', '排ガス分析計', '排ガスSOx計', kind='計装', spec='赤外線式', cat='analyzer')
    S(s, sn, '3', '排ガス分析計', '排ガスNOx計', kind='計装', spec='化学発光式', cat='analyzer')
    S(s, sn, '3', '排ガス分析計', 'ばいじん計', kind='計装', spec='光散乱式', cat='analyzer')
    S(s, sn, '3', '排ガス分析計', '排ガス流量計', kind='計装', spec='超音波式', cat='analyzer')
    S(s, sn, '3', '排ガス分析計', '水銀連続測定装置', kind='計装', spec='還元気化 原子吸光', cat='analyzer')
    # B0 炉別電気計装
    s, sn = 'B0', '炉別電気計装設備'
    S(s, sn, '1', '炉別盤', '炉別動力制御盤（MCC）', kind='盤', spec='400V 12面', cat='panel', imp='A')
    S(s, sn, '1', '炉別盤', '給じん現場操作盤', kind='盤', spec='壁掛', cat='panel')
    S(s, sn, '1', '炉別盤', '出湯口現場操作盤', kind='盤', spec='自立 防熱形', cat='panel', imp='A')
    S(s, sn, '1', '炉別盤', '水砕現場操作盤', kind='盤', spec='壁掛', cat='panel')
    S(s, sn, '2', '炉別制御装置', '炉別制御ステーション（FCS）', kind='盤', spec='DCS コントローラ二重化', cat='dcs', imp='A')
    S(s, sn, '2', '炉別制御装置', '炉別入出力盤（I/O-1）', kind='盤', spec='リモートI/O', cat='dcs', imp='A')
    S(s, sn, '2', '炉別制御装置', '炉別入出力盤（I/O-2）', kind='盤', spec='リモートI/O', cat='dcs', imp='A')

FURN_VALVES = ['給水入口電動弁', 'エコノマイザ出口電動弁', '主蒸気ドレン電動弁', '過熱器出口ドレン電動弁', 'ドラムベント電動弁',
    '起動用蒸気放出弁', '送風予熱器蒸気調節弁', '送風予熱器ドレン電動弁', '炉体冷却水温度調節弁', '羽口冷却水圧力調節弁',
    '羽口冷却水遮断弁', 'ジャケット冷却水遮断弁', '水砕水補給水調節弁', '水砕水供給電動弁', '噴霧水遮断弁',
    '噴霧空気調節弁', '燃焼室入口ダンパ', 'ボイラ入口ダンパ', '減温塔入口ダンパ', '減温塔バイパスダンパ',
    '再加熱器入口ダンパ', '触媒バイパスダンパ', '煙突入口ダンパ', '炉頂ガス緊急放散弁', '炉内圧力緊急放散ダンパ',
    'バーナ灯油遮断弁', 'バーナ灯油流量調節弁', '出湯樋予熱灯油遮断弁', '窒素パージ電動弁', '酸素ラインパージ電動弁',
    '送風機吐出ダンパ（A）', '送風機吐出ダンパ（B）', '冷却塔補給水調節弁', 'スートブロワ蒸気元弁', 'スートブロワ蒸気減圧弁',
    'アンモニア水遮断弁', '清缶剤注入電動弁', 'ボイラ水サンプリング電動弁', '燃焼空気予熱器蒸気調節弁', '排ガス再循環ダンパ']

def build():
    UNITS.clear(); _ser[0] = 0
    rng.seed(SEED)
    for f in (1, 2):
        start = len(UNITS)
        furnace(f)
        cnt = len(UNITS) - start
        need = 260 - cnt
        assert 0 <= need <= len(FURN_VALVES), (cnt, need)
        for nm in FURN_VALVES[:need]:
            kind = '調節弁' if '調節' in nm or '減圧' in nm else '電動弁'
            add(f, 'C0', '炉別弁類', '1', '電動弁・調節弁', nm, kind=kind, spec='電動' if kind == '電動弁' else '空気作動',
                cat='valve', imp='A' if ('緊急' in nm or '酸素' in nm) else None)
    common_build()
    # 親の解決・番号付け
    numbering()
    return UNITS

def common_build():
    S = lambda *a, **k: add(0, *a, **k)
    s, sn = '01', '受入供給設備'
    S(s, sn, '1', '計量設備', 'トラックスケール（入口）', spec='ロードセル 30t 目量 10kg', cat='scale', imp='A')
    S(s, sn, '1', '計量設備', 'トラックスケール（出口）', spec='ロードセル 30t 目量 10kg', cat='scale', imp='A')
    S(s, sn, '1', '計量設備', '計量システム盤', kind='盤', spec='データ処理装置', cat='panel')
    S(s, sn, '2', 'プラットホーム', 'プラットホーム出入口扉', spec='高速シートシャッタ', kw=2.2, n=2, duty='', cat='rot')
    S(s, sn, '2', 'プラットホーム', 'ごみ投入扉', spec='観音開き 油圧', n=6, duty='', cat='mech')
    hu = S(s, sn, '2', 'プラットホーム', '投入扉油圧ユニット', spec='7MPa', cat='mech', imp='A')
    S(s, sn, '2', 'プラットホーム', '投入扉油圧ポンプ', spec='ギヤポンプ', kw=7.5, n=2, parent=hu, cat='rot', imp='B')
    S(s, sn, '2', 'プラットホーム', 'ダンピングボックス', spec='油圧傾動', cat='mech')
    S(s, sn, '3', 'ごみピット', 'ごみピット', spec='RC 6,500m³', cat='static')
    S(s, sn, '3', 'ごみピット', 'ピット火災検知装置', kind='計装', spec='赤外線カメラ 2台', cat='analyzer')
    S(s, sn, '3', 'ごみピット', '放水銃', spec='遠隔操作 2基', n=2, duty='', cat='mech')
    for i, d in ((1, '常用'), (2, '予備')):
        c = S(s, sn, '4', 'ごみクレーン', f'{i}号ごみクレーン', spec='天井走行 バケット 4.5m³ 定格 8.5t', duty=d, cat='crane', imp='B', short=f'{i}号ごみクレーン')
        S(s, sn, '4', 'ごみクレーン', f'{i}号ごみクレーン走行装置', spec='', kw=11, parent=c, duty=d, cat='rot', nmot=2)
        S(s, sn, '4', 'ごみクレーン', f'{i}号ごみクレーン横行装置', spec='', kw=5.5, parent=c, duty=d, cat='rot')
        S(s, sn, '4', 'ごみクレーン', f'{i}号ごみクレーン巻上装置', spec='ワイヤロープ φ22×4本', kw=75, parent=c, duty=d, cat='rot')
        S(s, sn, '4', 'ごみクレーン', f'{i}号ごみクレーンバケット', spec='油圧開閉 ポリップ 4.5m³', kw=22, parent=c, duty=d, cat='rot')
        S(s, sn, '4', 'ごみクレーン', f'{i}号ごみクレーンインバータ盤', kind='盤', spec='巻上・走行・横行', parent=c, duty=d, cat='panel', imp='B')
    S(s, sn, '4', 'ごみクレーン', 'クレーン操作卓', kind='盤', spec='自動運転装置付', cat='panel', imp='A')
    S(s, sn, '5', '脱臭設備', '脱臭装置', spec='活性炭吸着式 600m³/min', cat='static')
    S(s, sn, '5', '脱臭設備', '脱臭ファン', spec='ターボ', kw=37, cat='rot')
    s, sn = '02', '副資材供給設備'
    S(s, sn, '1', 'コークス受入', 'コークス受入ホッパ', spec='20m³ グリズリ付', cat='static')
    S(s, sn, '1', 'コークス受入', 'コークス受入コンベヤ', spec='ベルト 450W 20t/h', kw=5.5, cat='rot', imp='A')
    S(s, sn, '1', 'コークス受入', 'コークスバケットエレベータ', spec='遠心排出式 20t/h H=32m', kw=11, cat='rot', imp='A')
    S(s, sn, '1', 'コークス受入', 'コークスバンカ', spec='鋼板製 200m³（約 90t）', cat='static', imp='A')
    S(s, sn, '1', 'コークス受入', 'コークスバンカレベル計', kind='計装', spec='マイクロ波式', cat='instr')
    S(s, sn, '1', 'コークス受入', 'コークス振動フィーダ', spec='電磁式', kw=2.2, cat='rot', imp='A')
    S(s, sn, '2', '石灰石受入', '石灰石受入ホッパ', spec='15m³', cat='static')
    S(s, sn, '2', '石灰石受入', '石灰石受入コンベヤ', spec='ベルト 450W 15t/h', kw=3.7, cat='rot', imp='A')
    S(s, sn, '2', '石灰石受入', '石灰石バケットエレベータ', spec='遠心排出式 15t/h H=32m', kw=7.5, cat='rot', imp='A')
    S(s, sn, '2', '石灰石受入', '石灰石バンカ', spec='鋼板製 120m³（約 160t）', cat='static', imp='A')
    S(s, sn, '2', '石灰石受入', '石灰石バンカレベル計', kind='計装', spec='マイクロ波式', cat='instr')
    S(s, sn, '2', '石灰石受入', '石灰石振動フィーダ', spec='電磁式', kw=1.5, cat='rot', imp='A')
    S(s, sn, '3', '炉頂供給', '副資材炉頂供給コンベヤ', spec='ベルト 両炉振分 10t/h', kw=5.5, n=2, duty=['常用', '予備'], cat='rot', imp='B')
    S(s, sn, '3', '炉頂供給', '副資材振分ダンパ', kind='電動弁', spec='電動 2方向', cat='valve', imp='A')
    S(s, sn, '4', '副資材集じん', '副資材集じん機', spec='バグ式 120m³/min', cat='static')
    S(s, sn, '4', '副資材集じん', '副資材集じんファン', spec='ターボ', kw=15, cat='rot')
    S(s, sn, '4', '副資材集じん', '副資材現場盤', kind='盤', spec='自立形', cat='panel')
    s, sn = '03', '酸素発生設備'
    S(s, sn, '1', 'PSA酸素発生装置', 'PSA原料空気ブロワ', spec='ルーツ 3,800m³N/h', kw=110, cat='rot', imp='A', short='PSA原料空気ブロワ')
    S(s, sn, '1', 'PSA酸素発生装置', 'PSA真空ポンプ', spec='ルーツ 真空 -60kPa', kw=160, n=2, duty=['常用', '常用'], cat='rot', imp='A')
    S(s, sn, '1', 'PSA酸素発生装置', 'PSA吸着塔', spec='ゼオライト充填 φ2.8m', n=2, duty=['常用', '常用'], cat='psa', imp='A', short='吸着塔')
    S(s, sn, '1', 'PSA酸素発生装置', 'PSA製品酸素槽', spec='20m³ 0.2MPa', cat='static', imp='A')
    S(s, sn, '1', 'PSA酸素発生装置', '酸素圧縮機', spec='ルーツ 800m³N/h 0.1MPa', kw=55, n=2, cat='rot', imp='B')
    S(s, sn, '1', 'PSA酸素発生装置', 'PSA原料空気冷却器', spec='水冷式', cat='static')
    S(s, sn, '1', 'PSA酸素発生装置', 'PSA切替弁', kind='調節弁', spec='空気作動 バタフライ', n=6, duty='', cat='valve', imp='A')
    S(s, sn, '1', 'PSA酸素発生装置', 'PSA制御盤', kind='盤', spec='PLC', cat='dcs', imp='A')
    S(s, sn, '1', 'PSA酸素発生装置', '製品酸素濃度計（PSA出口）', kind='計装', spec='磁気式', cat='analyzer', imp='A', ident='ak_new')
    S(s, sn, '1', 'PSA酸素発生装置', 'PSA原料空気ブロワインバータ盤', kind='盤', spec='400V 160kVA', cat='panel', imp='A', ident='ak_new')
    S(s, sn, '2', '液体酸素設備', '液体酸素貯槽（CE）', spec='真空断熱 20m³', cat='static', imp='A', short='CE')
    S(s, sn, '2', '液体酸素設備', '液体酸素蒸発器', spec='空温式 500m³N/h', n=2, cat='static', imp='B')
    S(s, sn, '2', '液体酸素設備', '液酸供給圧力調節弁', kind='調節弁', spec='酸素用', cat='valve', imp='A')
    S(s, sn, '2', '液体酸素設備', '酸素供給ヘッダ流量計', kind='計装', spec='熱式', cat='instr')
    S(s, sn, '2', '液体酸素設備', '酸素漏えい検知器', kind='計装', spec='ガルバニ式 4点', cat='analyzer')
    s, sn = '04', 'スラグ・メタル処理設備'
    S(s, sn, '1', 'スラグ搬送', 'スラグ搬送コンベヤ', spec='ベルト 両炉合流', kw=3.7, n=2, duty='', cat='rot')
    S(s, sn, '1', 'スラグ搬送', 'スラグ水切りスクリーン', spec='振動', kw=2.2, cat='rot')
    S(s, sn, '2', '選別装置', '磁選機', spec='吊下げ式 永磁', kw=3.7, cat='rot', imp='A')
    S(s, sn, '2', '選別装置', 'スラグ振動ふるい', spec='2段 5mm', kw=5.5, cat='rot', imp='A')
    S(s, sn, '2', '選別装置', 'スラグ粒度調整機', spec='ロール式', kw=15, cat='rot')
    S(s, sn, '2', '選別装置', 'メタル搬送コンベヤ', spec='ベルト', kw=1.5, cat='rot')
    S(s, sn, '2', '選別装置', 'メタルホッパ', spec='5m³', cat='static')
    S(s, sn, '2', '選別装置', 'スラグ貯留ホッパ', spec='60m³ 2槽', n=2, duty='', cat='static')
    S(s, sn, '2', '選別装置', 'スラグ・メタル現場盤', kind='盤', spec='自立形', cat='panel')
    s, sn = '05', '余熱利用設備'
    S(s, sn, '1', 'ボイラ給水', '脱気器', spec='加圧式 35t/h 第一種圧力容器', cat='vessel', imp='A', short='脱気器')
    S(s, sn, '1', 'ボイラ給水', 'ボイラ給水ポンプ', spec='多段渦巻 20m³/h×600m', kw=75, n=3, duty=['常用', '常用', '予備'], cat='rot', imp='B')
    S(s, sn, '1', 'ボイラ給水', '脱気器給水ポンプ', spec='渦巻', kw=11, n=2, cat='rot', imp='B')
    S(s, sn, '1', 'ボイラ給水', '純水装置', spec='イオン交換 5m³/h', cat='static')
    S(s, sn, '1', 'ボイラ給水', '純水ポンプ', spec='渦巻', kw=3.7, n=2, cat='rot')
    S(s, sn, '1', 'ボイラ給水', '脱酸素剤注入ポンプ', spec='ダイヤフラム', kw=0.2, n=2, cat='rot')
    S(s, sn, '1', 'ボイラ給水', '復水タンク', spec='SUS 30m³', cat='static')
    S(s, sn, '1', 'ボイラ給水', '高圧蒸気だめ', spec='4MPa φ600', cat='vessel')
    S(s, sn, '1', 'ボイラ給水', '低圧蒸気だめ', spec='0.5MPa φ500', cat='vessel')
    S(s, sn, '2', '蒸気タービン発電機', '蒸気タービン', spec='抽気復水 3.8MPa 390℃ 5,000kW', cat='turbine', imp='A', short='タービン')
    S(s, sn, '2', '蒸気タービン発電機', 'タービン減速機', spec='平行軸 6,600/1,500rpm', cat='mech', imp='A')
    S(s, sn, '2', '蒸気タービン発電機', '発電機', spec='三相同期 6.6kV 6,250kVA', cat='gen', imp='A')
    S(s, sn, '2', '蒸気タービン発電機', 'タービン補助油ポンプ', spec='歯車', kw=15, cat='rot', imp='A')
    S(s, sn, '2', '蒸気タービン発電機', 'タービン非常油ポンプ', spec='歯車 DC', kw=5.5, cat='rot', imp='A')
    S(s, sn, '2', '蒸気タービン発電機', 'ターニング装置', spec='', kw=3.7, cat='rot')
    S(s, sn, '2', '蒸気タービン発電機', 'タービン油冷却器', spec='シェルチューブ', n=2, cat='static')
    S(s, sn, '2', '蒸気タービン発電機', 'グランド蒸気復水器', spec='', cat='static')
    S(s, sn, '2', '蒸気タービン発電機', 'グランド排風機', spec='', kw=1.5, cat='rot')
    S(s, sn, '2', '蒸気タービン発電機', 'タービン制御盤', kind='盤', spec='電子ガバナ', cat='panel', imp='A')
    S(s, sn, '2', '蒸気タービン発電機', '発電機盤', kind='盤', spec='6.6kV', cat='panel', imp='A')
    S(s, sn, '2', '蒸気タービン発電機', 'タービン主蒸気止弁', kind='電動弁', spec='油圧＋電動', cat='valve', imp='A')
    S(s, sn, '2', '蒸気タービン発電機', 'タービンバイパス弁', kind='調節弁', spec='減圧減温', cat='valve', imp='A')
    S(s, sn, '2', '蒸気タービン発電機', 'タービン軸振動計', kind='計装', spec='非接触 4点', cat='instr', imp='A')
    S(s, sn, '3', '空冷復水器', '空冷復水器本体', spec='A形 4ベイ', cat='static', imp='A')
    S(s, sn, '3', '空冷復水器', '空冷復水器ファン', spec='軸流 φ6.1m 減速機付', kw=37, n=8, duty='', cat='rot', short='ACCファン')
    S(s, sn, '3', '空冷復水器', '復水ポンプ', spec='立形多段', kw=18.5, n=2, cat='rot', imp='B')
    S(s, sn, '3', '空冷復水器', '空気抽出器', spec='蒸気エゼクタ 2段', n=2, cat='static', imp='B')
    S(s, sn, '3', '空冷復水器', '復水器真空計', kind='計装', spec='', cat='instr')
    s, sn = '06', '灰処理設備'
    S(s, sn, '1', '飛灰処理', '飛灰貯槽', spec='60m³', cat='static', imp='A')
    S(s, sn, '1', '飛灰処理', '飛灰定量供給機', spec='', kw=1.5, cat='rot', imp='A')
    S(s, sn, '1', '飛灰処理', '飛灰混練機', spec='2軸パドル 2t/h', kw=15, n=2, cat='rot', imp='B')
    S(s, sn, '1', '飛灰処理', 'キレート薬剤タンク', spec='3m³', cat='static')
    S(s, sn, '1', '飛灰処理', 'キレート薬剤ポンプ', spec='', kw=0.4, n=2, cat='rot')
    S(s, sn, '1', '飛灰処理', '処理灰搬送コンベヤ', spec='ベルト', kw=2.2, cat='rot')
    S(s, sn, '1', '飛灰処理', '飛灰貯槽バグフィルタ', spec='', cat='static')
    S(s, sn, '1', '飛灰処理', '飛灰処理現場盤', kind='盤', spec='', cat='panel')
    s, sn = '07', '薬剤貯留設備'
    S(s, sn, '1', '薬剤貯留', '消石灰貯槽', spec='60m³ 受入バグ付', cat='static')
    S(s, sn, '1', '薬剤貯留', '消石灰移送ブロワ', spec='', kw=7.5, cat='rot')
    S(s, sn, '1', '薬剤貯留', '活性炭貯槽', spec='20m³', cat='static')
    S(s, sn, '1', '薬剤貯留', 'アンモニア水貯槽', spec='25%  10m³', cat='static', imp='A')
    S(s, sn, '1', '薬剤貯留', 'アンモニア水移送ポンプ', spec='', kw=0.75, n=2, cat='rot')
    S(s, sn, '1', '薬剤貯留', 'アンモニア漏えい検知器', kind='計装', spec='', cat='analyzer')
    s, sn = '08', '給水設備'
    S(s, sn, '1', '用水', 'プラント用水受水槽', spec='RC 300m³', cat='static')
    S(s, sn, '1', '用水', 'プラント用水ポンプ', spec='渦巻', kw=15, n=2, cat='rot', imp='B')
    S(s, sn, '1', '用水', '生活用水ポンプ', spec='', kw=3.7, n=2, cat='rot')
    S(s, sn, '2', '機器冷却水', '機器冷却水ポンプ', spec='渦巻 250m³/h', kw=37, n=2, cat='rot', imp='B')
    S(s, sn, '2', '機器冷却水', '機器冷却塔', spec='開放式 250m³/h', kw=7.5, n=2, duty='', cat='rot')
    S(s, sn, '2', '機器冷却水', '冷却水薬注装置', spec='', kw=0.2, cat='rot')
    s, sn = '09', '排水処理設備'
    S(s, sn, '1', 'プラント排水', '排水受槽', spec='RC 80m³', cat='static')
    S(s, sn, '1', 'プラント排水', '排水移送ポンプ', spec='水中', kw=3.7, n=2, cat='rot')
    S(s, sn, '1', 'プラント排水', '凝集沈殿槽', spec='', cat='static')
    S(s, sn, '1', 'プラント排水', '凝集槽撹拌機', spec='', kw=1.5, n=2, duty='', cat='rot')
    S(s, sn, '1', 'プラント排水', '砂ろ過器', spec='', cat='static')
    S(s, sn, '1', 'プラント排水', '処理水ポンプ', spec='', kw=3.7, n=2, cat='rot')
    S(s, sn, '1', 'プラント排水', '排水薬注ポンプ', spec='', kw=0.2, n=3, duty='', cat='rot')
    S(s, sn, '1', 'プラント排水', '排水pH計', kind='計装', spec='', cat='instr')
    S(s, sn, '1', 'プラント排水', '排水処理現場盤', kind='盤', spec='', cat='panel')
    s, sn = '10', '圧縮空気設備'
    S(s, sn, '1', '圧縮空気', '空気圧縮機', spec='スクリュー 10m³/min', kw=55, n=3, duty=['常用', '常用', '予備'], cat='rot', imp='B')
    S(s, sn, '1', '圧縮空気', '除湿器', spec='冷凍式', kw=2.2, n=2, cat='rot')
    S(s, sn, '1', '圧縮空気', '雑用空気槽', spec='5m³', cat='static')
    S(s, sn, '1', '圧縮空気', '計装用空気槽', spec='5m³', cat='static')
    S(s, sn, '1', '圧縮空気', '窒素ガス発生装置', spec='PSA 50m³N/h', kw=11, cat='rot', imp='A')
    s, sn = '11', '電気設備'
    S(s, sn, '1', '受変電', '特高受電設備（GIS）', kind='盤', spec='66kV', cat='hv', imp='A')
    S(s, sn, '1', '受変電', '主変圧器', kind='盤', spec='66/6.6kV 7,500kVA', cat='hv', imp='A')
    for nm in ['高圧受電盤', '高圧発電機連系盤', '高圧き電盤（1号炉）', '高圧き電盤（2号炉）', '高圧き電盤（共通）', '高圧き電盤（PSA）', '高圧進相コンデンサ盤']:
        S(s, sn, '2', '高圧配電', nm, kind='盤', spec='6.6kV', cat='hv', imp='A')
    for nm in ['動力変圧器（1号炉）', '動力変圧器（2号炉）', '動力変圧器（共通）', '動力変圧器（PSA）', '照明変圧器', '建築動力変圧器']:
        S(s, sn, '3', '変圧器', nm, kind='盤', spec='6.6kV/400V モールド', cat='hv')
    for nm in ['低圧配電盤（1号炉）', '低圧配電盤（2号炉）', '低圧配電盤（共通-1）', '低圧配電盤（共通-2）', '低圧配電盤（PSA）', '共通動力制御盤（MCC-1）', '共通動力制御盤（MCC-2）', '共通動力制御盤（MCC-3）']:
        S(s, sn, '4', '低圧配電', nm, kind='盤', spec='400V', cat='panel', imp='A')
    S(s, sn, '5', '非常用電源', '非常用発電機', spec='ディーゼル 6.6kV 750kVA', cat='egen', imp='A', short='非常用発電機')
    S(s, sn, '5', '非常用電源', '非常用発電機始動盤', kind='盤', spec='', cat='panel')
    S(s, sn, '5', '非常用電源', '直流電源装置', kind='盤', spec='100V 鉛蓄電池', cat='panel', imp='A')
    S(s, sn, '5', '非常用電源', '無停電電源装置（UPS-1）', kind='盤', spec='20kVA', cat='panel', imp='A')
    S(s, sn, '5', '非常用電源', '無停電電源装置（UPS-2）', kind='盤', spec='20kVA', cat='panel', imp='A')
    s, sn = '12', '計装制御設備'
    for nm in ['中央操作端末（OPS-1）', '中央操作端末（OPS-2）', '中央操作端末（OPS-3）']:
        S(s, sn, '1', '分散制御装置', nm, kind='盤', spec='', cat='dcs', imp='A')
    S(s, sn, '1', '分散制御装置', '中央操作端末（OPS-4）', kind='盤', spec='', cat='dcs', imp='A', ident='ak_new')
    S(s, sn, '1', '分散制御装置', 'エンジニアリングステーション', kind='盤', spec='', cat='dcs')
    S(s, sn, '1', '分散制御装置', '共通制御ステーション（FCS-C1）', kind='盤', spec='', cat='dcs', imp='A')
    S(s, sn, '1', '分散制御装置', '共通制御ステーション（FCS-C2）', kind='盤', spec='', cat='dcs', imp='A')
    S(s, sn, '1', '分散制御装置', '共通入出力盤', kind='盤', spec='', cat='dcs', imp='A', n=2, duty='')
    S(s, sn, '1', '分散制御装置', 'データロガ・帳票サーバ', kind='盤', spec='', cat='dcs')
    S(s, sn, '1', '分散制御装置', '大型表示装置', kind='盤', spec='', cat='dcs')
    S(s, sn, '1', '分散制御装置', 'ITV監視盤', kind='盤', spec='', cat='panel')
    S(s, sn, '2', '共通計装', '煙突排ガス監視表示盤', kind='盤', spec='公害監視', cat='panel')
    S(s, sn, '2', '共通計装', '気象観測装置', kind='計装', spec='風向風速', cat='instr')
    S(s, sn, '3', '煙突', '煙突外筒', spec='RC H=59m', cat='static')

    s, sn = '14', '付帯設備'
    S(s, sn, '1', '灯油設備', '灯油貯槽', spec='地下 20kL', cat='static')
    S(s, sn, '1', '灯油設備', '灯油移送ポンプ', spec='ギヤ', kw=1.5, n=2, cat='rot')
    S(s, sn, '2', 'ピット汚水', 'ピット汚水槽', spec='RC 30m³', cat='static')
    S(s, sn, '2', 'ピット汚水', 'ピット汚水ポンプ', spec='水中', kw=2.2, n=2, cat='rot')
    S(s, sn, '2', 'ピット汚水', 'ピット消火ポンプ', spec='渦巻 放水銃用', kw=22, cat='rot', imp='A')
    S(s, sn, '2', 'ピット汚水', '洗車ポンプ', spec='', kw=3.7, cat='rot')
    S(s, sn, '3', '保守用クレーン', '炉室天井クレーン', spec='10t 天井走行', cat='crane2', short='炉室天井クレーン')
    S(s, sn, '3', '保守用クレーン', '炉室天井クレーン走行装置', spec='', kw=3.7, cat='rot')
    S(s, sn, '3', '保守用クレーン', '炉室天井クレーン横行装置', spec='', kw=1.5, cat='rot')
    S(s, sn, '3', '保守用クレーン', '炉室天井クレーン巻上装置', spec='', kw=15, cat='rot')
    S(s, sn, '3', '保守用クレーン', 'タービン室天井クレーン', spec='15t 天井走行', cat='crane2', short='タービン室天井クレーン')
    S(s, sn, '3', '保守用クレーン', 'タービン室天井クレーン走行装置', spec='', kw=3.7, cat='rot')
    S(s, sn, '3', '保守用クレーン', 'タービン室天井クレーン横行装置', spec='', kw=1.5, cat='rot')
    S(s, sn, '3', '保守用クレーン', 'タービン室天井クレーン巻上装置', spec='', kw=18.5, cat='rot')
    S(s, sn, '4', 'その他', '純水再生ポンプ', spec='', kw=2.2, n=2, cat='rot')
    S(s, sn, '4', 'その他', 'ACC洗浄ポンプ', spec='高圧 15MPa', kw=11, cat='rot')
    S(s, sn, '4', 'その他', 'コークス篩', spec='振動 10mm', kw=2.2, cat='rot')
    S(s, sn, '4', 'その他', 'スラグヤード散水ポンプ', spec='', kw=3.7, cat='rot')
    S(s, sn, '4', 'その他', '窒素ガス貯槽', spec='5m³', cat='static')
    S(s, sn, '4', 'その他', 'プラットホーム監視ITV', kind='計装', spec='', cat='instr')
    S('11', '電気設備', '6', '発電関係盤', '発電機励磁盤', kind='盤', spec='ブラシレス', cat='panel', imp='A')
    S('11', '電気設備', '6', '発電関係盤', '系統連系保護継電器盤', kind='盤', spec='', cat='panel', imp='A')

COMMON_VALVES = ['主蒸気ヘッダ電動弁', '高圧蒸気だめ圧力調節弁', '低圧蒸気減圧調節弁', 'タービン抽気逆止弁', '脱気器圧力調節弁',
    '脱気器水位調節弁', '給水ヘッダ電動弁（A）', '給水ヘッダ電動弁（B）', '給水再循環調節弁（A）', '給水再循環調節弁（B）',
    '給水再循環調節弁（C）', '復水再循環調節弁', '復水器水位調節弁', '補給水調節弁', 'ACC真空調節弁', '所内蒸気電動弁',
    'プラント用水電動弁', '機器冷却水温度調節弁', '冷却塔補給水調節弁', '排水処理放流電動弁', '消石灰移送切替弁',
    '雑用空気電動弁', '計装用空気圧力調節弁', '窒素ヘッダ電動弁', '酸素供給ヘッダ遮断弁', '酸素供給圧力調節弁（PSA系）',
    '灯油移送電動弁', '灯油タンク元弁（電動）', '復水タンク水位調節弁', '純水補給電動弁', 'グランド蒸気調節弁',
    'タービン排気電動弁', 'ACC入口電動弁（1）', 'ACC入口電動弁（2）', '脱臭切替ダンパ', 'プラットホーム排気ダンパ',
    'ピット排気ダンパ', '副資材集じんダンパ', 'スラグヤード散水電動弁', 'PSA原料空気放風弁', '液酸蒸発器切替弁（A）',
    '液酸蒸発器切替弁（B）', 'アンモニア水移送電動弁', 'キレート注入調節弁', '飛灰貯槽排気ダンパ', '洗車場給水電動弁']

def numbering():
    # 共通の弁の詰め（新台帳時点 395）
    com = [u for u in UNITS if u['furn'] == 0]
    need = 395 - len(com)
    assert 0 <= need <= len(COMMON_VALVES), (len(com), need)
    for nm in COMMON_VALVES[:need]:
        kind = '調節弁' if '調節' in nm else '電動弁'
        add(0, '13', '共通弁類', '1', '電動弁・調節弁', nm, kind=kind, spec='電動' if kind == '電動弁' else '空気作動', cat='valve',
            imp='A' if '酸素' in nm else None)
    # 並べ替え：炉→系統→設備→登録順
    order = {id(u): i for i, u in enumerate(UNITS)}
    UNITS.sort(key=lambda u: (u['furn'] if u['furn'] else 9, u['sys'], u['sub'], order[id(u)]))
    # 系統図番号
    cnt = {}
    for u in UNITS:
        if u['kind'] == '電動機':
            continue
        pu = u.get('_par')
        if pu is None:
            key = (u['furn'], u['sys'], u['sub'])
            cnt[key] = cnt.get(key, 0) + 1
            u['no'] = cnt[key]
        if pu is not None:
            u['parent'] = pu['tag']
            pu['_ch'] = pu.get('_ch', 0) + 1
            u['tag'] = f"{pu['tag']}-{pu['_ch']}"
        else:
            u['tag'] = f"{u['furn']}-{u['sys']}-{u['sub']}-{u['no']:02d}"
    for u in UNITS:
        if u['kind'] == '電動機':
            pu = u['_p']
            pu['_m'] = pu.get('_m', 0) + 1
            u['tag'] = f"{pu['tag']}-M{pu['_m']}"
            u['parent'] = pu['tag']
    # 親子の木の順に並べ直し（親→付属→電動機）
    bytag = {u['tag']: u for u in UNITS}
    kids = {}
    for u in UNITS:
        if u['parent']:
            kids.setdefault(u['parent'], []).append(u)
    out = []
    def walk(u, d):
        u['depth'] = d; out.append(u)
        ks = sorted(kids.get(u['tag'], []), key=lambda k: (k['kind'] != '電動機', k['tag']))
        for k in ks: walk(k, d + 1)
    for u in UNITS:
        if not u['parent']: walk(u, 0)
    UNITS[:] = out
    # 電動機は旧電動機台帳（H18〜H23）の番号 M-0001… で持つ（H24 の統合でも番号は変えなかった）
    mn = 0
    for u in UNITS:
        if u['kind'] == '電動機':
            mn += 1
            u['tag'] = f'M-{mn:04d}'
    for u in UNITS:
        u.pop('_p', None); u.pop('_par', None)
        # 重要度の既定
        if u['imp'] is None:
            if u['kind'] == '電動機':
                pu = bytag[u['parent']]
                u['imp'] = pu['imp'] or 'C'
            elif u['cat'] in ('static', 'instr'):
                u['imp'] = 'C'
            else:
                u['imp'] = 'B' if u['duty'] in ('常用', '予備') else 'C'

# ------------------------------------------------------------------
# 基幹的設備改良工事（令和3〜6年度）の範囲（正解）
# ------------------------------------------------------------------
AK = {
    # tag の一部 / 名前で指す → (施工年度, 内容)
}
def ak_scope(u):
    n = u['base']
    f = u['furn']
    if f in (1, 2):
        y = 2022 if f == 1 else 2023
        if n in ('溶融炉本体', '炉体冷却ジャケット', '主羽口装置', '副羽口装置'): return y, '更新'
        if n in ('羽口冷却水ポンプ', '羽口冷却水ポンプ用電動機'): return y, '更新'
        if n == '出湯口' and f == 2: return 2023, '更新'
        if n in ('炉別制御ステーション（FCS）', '炉別入出力盤（I/O-1）', '炉別入出力盤（I/O-2）'): return 2024, '更新'
        if n in ('炉体温度計',): return y, '更新'
    else:
        if u['sys'] == '03' and u['sub'] == '1': return 2023, '更新'
        if u['sys'] == '12' and u['sub'] == '1' and 'ITV' not in n: return 2024, '更新'
    return None

# ------------------------------------------------------------------
# 契約（正解）。見本に使うものを中心に。
# ------------------------------------------------------------------
CONTRACTS = {
    'H30保守': dict(name=f'平成30年度 {FACN} 年間保守点検整備業務委託', kind='業務委託', vendor=MAKER,
                  start=dt.date(2018, 4, 1), end=dt.date(2019, 3, 31), amount=264_600_000, method='随意契約'),
    'R3保守': dict(name=f'令和3年度 {FACN} 年間保守点検整備業務委託', kind='業務委託', vendor=MAKER,
                 start=dt.date(2021, 4, 1), end=dt.date(2022, 3, 31), amount=286_000_000, method='随意契約',
                 change=dict(date=dt.date(2021, 11, 26), amount=289_520_000, reason='1号ボイラ2次過熱器管肉盛補修の追加')),
    'R3出湯口': dict(name=f'令和3年度 {FACN} 1号溶融炉出湯口補修工事', kind='補修工事（請負）', vendor=MAKER,
                  start=dt.date(2021, 8, 30), end=dt.date(2021, 12, 10), amount=38_610_000, method='随意契約'),
    '基幹改良': dict(name=f'{FACN} 基幹的設備改良工事', kind='更新工事（請負・継続費）', vendor=MAKER,
                 start=dt.date(2021, 7, 5), end=dt.date(2025, 3, 25), amount=7_920_000_000, method='随意契約（公募型）'),
}

# 定修の窓（正解）
OUTAGES = {
    (2018, 2): (dt.date(2018, 6, 4), dt.date(2018, 7, 6)),
    (2018, 1): (dt.date(2018, 10, 1), dt.date(2018, 11, 9)),
    (2021, 2): (dt.date(2021, 6, 7), dt.date(2021, 7, 9)),
    (2021, 1): (dt.date(2021, 9, 27), dt.date(2021, 11, 9)),
}

if __name__ == '__main__':
    us = build()
    from collections import Counter
    print(len(us), Counter(u['furn'] for u in us), Counter(u['kind'] for u in us))
