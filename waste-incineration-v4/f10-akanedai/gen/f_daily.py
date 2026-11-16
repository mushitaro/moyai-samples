# -*- coding: utf-8 -*-
"""作業日報（千草）。工程表（periods の実績）から日ごとの行を作る。
様式：p1 H23〜H24（班別、文章と人員だけ）→ p1 H25 様式（縦型、場所・時間・KY）→ p2 R1 様式（横型、会社別集計が上、会社ごとのブロック）"""
import datetime as dt, random, math
from openpyxl import Workbook
import world, history as H
from xl import *
from rec import Rec

WD = '月火水木金土日'
MH = {}
def acc(ws_, mh):
    for w in ws_: MH[w] = round(MH.get(w, 0) + mh / max(1, len(ws_)), 2)

BYTAG = {}
def U(t):
    if not BYTAG: BYTAG.update({u['tag']: u for u in world.UNITS})
    return BYTAG[t]

def fmt(fy):
    return 'p1a' if fy <= 2012 else 'p1b' if fy <= 2018 else 'p2'

def jdate(d):
    y = f'平成{d.year-1988}' if d < dt.date(2019, 5, 1) else f'令和{d.year-2018}'
    return f'{y}年{d.month}月{d.day}日（{WD[d.weekday()]}）'

def weather(d, r):
    m = d.month
    return r.choices(['晴', '曇', '雨', '晴のち曇', '曇のち雨'], [5, 4, 2 if m not in (6, 7, 9) else 4, 2, 1])[0]

def hours_of(a, b):
    h1, m1 = map(int, a.split(':')); h2, m2 = map(int, b.split(':'))
    t = (h2 * 60 + m2 - h1 * 60 - m1) / 60
    if h1 < 12 and h2 >= 13: t -= 1.0
    return round(t, 1)

VERB = {'◎': ['分解整備（軸受・シール交換）', '分解点検・消耗品交換', '開放・部品点検・組立'], '△': ['補修', '不具合箇所 補修', '補修・調整'],
        '○': ['点検', '外観点検・作動確認', '点検・給油'], '◇': ['計測（肉厚・振動）', '状態計測'], '●': ['更新（撤去・据付・試運転）', '取替え・試運転']}
EL = {'電動機': '絶縁抵抗測定・軸受音確認', '計装': '校正・点検', '盤': '盤内点検・清掃・増締め', '電動弁': '作動試験・開度調整', '調節弁': '作動試験・開度調整'}

def day_rows(p, d, r):
    """その日の作業の行：[(item, 場所, 内容, 会社, 人数, 始, 終, 発見, units)]"""
    out = []
    for it in p['items']:
        if it['co'] == '市運転班' or not it['act']: continue
        a, b = it['act']
        if not (a <= d <= b): continue
        L = (b - a).days + 1; k = (d - a).days
        ph = 'first' if k == 0 else 'last' if k == L - 1 else 'mid'
        n = len(it['units'])
        per = max(1, math.ceil(n / L))
        sub = it['units'][k * per:(k + 1) * per] if it['tmpl'].startswith('ev') else it['units'][:1]
        if it['tmpl'].startswith('ev') and not sub: sub = it['units'][-1:]
        loc, txt, find = text(it, ph, k, sub, p, r)
        ppl = max(1, (it['ppl'] or 2) + r.choice([0, 0, 0, 1, -1]))
        st, en = r.choice([('8:30', '17:00'), ('8:30', '17:00'), ('8:00', '17:00'), ('9:00', '16:30'), ('8:30', '18:30'), ('13:00', '17:00')])
        if it['tmpl'] in ('hydro', 'legal', 'il'): st, en = ('9:00', '15:00') if it['tmpl'] != 'il' else ('13:00', '19:00')
        out.append(dict(it=it, loc=loc, txt=txt, co=it['co'], n=ppl, a=st, b=en, find=find, units=sub, no=it['no']))
    return out

def fn(t): return U(t)['name']

def text(it, ph, k, sub, p, r):
    t = it['tmpl']
    f0 = U(it['units'][0])['furn'] if it['units'] else 0
    fx = f'{f0}号' if f0 in (1, 2) else ''
    loc = fn(it['units'][0]) if it['units'] else ''
    find = ''
    pick = lambda xs: xs[(k + len(p['key'])) % len(xs)]
    if t == 'safety': return f'{fx}溶融炉', '炉内 酸欠・CO 測定、入槽許可（市立会）', ''
    if t == 'scaffold': return f'{fx}炉内・ボイラ', {'first': '足場組立 着手（羽口レベル）', 'last': '足場組立 完了・点検（市確認）'}.get(ph, f'足場組立（ボイラ {k+1}パス側）'), ''
    if t == 'scaffold2': return f'{fx}炉内・ボイラ', '足場解体' + ('・搬出' if ph == 'last' else ''), ''
    if t == 'residue':
        return f'{fx}溶融炉 炉内', {'first': '炉内残留物除去 着手（上段）', 'last': '残留物除去 完了（羽口部まで清掃）'}.get(ph, f'残留物除去 シャフト下部、フレコン {r.randint(4, 7)}袋 搬出'), ''
    if t == 'tuyere':
        return f'{fx}溶融炉 羽口部', '主羽口・副羽口 漏水試験 準備（縁切り）' if ph == 'first' else '主羽口 8本・副羽口 8本 漏水試験（0.5MPa・30分）', ''
    if t == 'refr':
        return f'{fx}溶融炉 炉内', {'first': '耐火物 損耗測定（シャフト部）', 'last': '耐火物補修 仕上げ・確認'}.get(ph, f'耐火物 部分補修（吹付 {r.randint(3, 9) / 10 + 0.5:.1f}m²）'), ''
    if t == 'wash': return f'{fx}ボイラ', f'ボイラ水洗 {min(k+1, 3)}パス' + ('・エコノマイザ' if ph == 'last' else ''), ''
    if t == 'ut':
        h = H.truth(U(it['units'][0]))
        v = h.get(p['fy'], ('', ''))[1]
        if ph == 'last' and v and 'mm' in v: find = f'2次過熱器 {v}'
        return f'{fx}ボイラ', f'過熱器・蒸発管 肉厚測定（UT）{k+1}日目', find
    if t == 'hydro': return f'{fx}ボイラ', 'ボイラ水圧試験 5.0MPa 保持 30分 異常なし', ''
    if t == 'legal': return f'{fx}ボイラ', 'ボイラ定期事業者検査 受検（内部・水圧試験記録）', '良'
    if t == 'nozzle': return f'{fx}減温塔', '噴霧ノズル 6本 交換・噴霧試験' if '交換' in it['name'] and ph != 'first' else '減温塔 内部点検（付着確認）', ''
    if t == 'bf':
        if 'ろ布' in it['name']:
            return f'{fx}BF', {'first': 'ろ布 取外し 1・2室'}.get(ph, f'ろ布 取付 {min(6, 2*k-1)}・{min(6, 2*k)}室' if ph != 'last' else 'ろ布全数交換 完了・差圧確認'), ''
        return f'{fx}BF', f'内部点検 {min(6, 2*k+1)}・{min(6, 2*k+2)}室（ろ布外観・ケージ）', ''
    if t == 'cat':
        s = H.truth(U(it['units'][0])).get(p['fy'], ('', ''))[0]
        return f'{fx}触媒反応塔', '活性測定用 抜取（上層・下層）' if s in ('◇', '○') else '外観点検', ''
    if t == 'idf':
        if '分解' in it['name']:
            return f'{fx}誘引通風機', {'first': '軸受箱 開放・旧軸受取外し', 'last': f'芯出し・試運転 振動 {r.randint(15, 24)/10}mm/s'}.get(ph, '羽根車 点検・バランス確認'), ''
        return f'{fx}誘引通風機', f'振動測定 {r.randint(15, 26)/10}mm/s・外観点検', ''
    if t == 'slag': return f'{fx}水砕槽', '水砕槽 水抜き・堆積スラグ排出' if ph == 'first' else f'掻揚コンベヤ チェーン伸び測定 {r.randint(6, 16)/10}%', ''
    if t == 'fan': return f'{fx}送風機', pick(['羽口送風機（A）（B） 軸受点検・グリス交換', '燃焼用空気送風機 点検']), ''
    if t == 'feed': return f'{fx}給じん装置', pick(['上部シールダンパ シール面点検', '下部シールダンパ シール面点検・研磨', '給じん油圧ユニット 作動確認']), ''
    if t == 'cal':
        ins = [U(x)['name'] for x in []] or ['炉頂温度計・炉内圧力計', '送風流量計・送風酸素濃度計', '排ガス分析計（O2・CO）', '排ガス分析計（HCl・SOx・NOx）', 'ばいじん計・排ガス流量計', 'ボイラ計装（ドラム水位・主蒸気）']
        return f'{fx}炉 計装', pick(ins) + ' 校正', ''
    if t == 'mcc': return f'{fx}電気室' if fx else '電気室', 'MCC・配電盤 停電点検（端子増締め・絶縁抵抗）', ''
    if t == 'close': return f'{fx}溶融炉・ボイラ', 'マンホール閉止' + (' 完了' if ph == 'last' else ''), ''
    if t == 'il': return f'{fx}炉', 'インターロック試験（送風・給じん・出湯口・IDF）', '良'
    if t == 'tap':
        if it.get('sym') == '◎':
            return f'{fx}出湯口', {'first': '出湯口 解体・はつり 着手', 'last': '出湯口 キャスタブル施工・養生'}.get(ph, f'出湯口れんが築造 {k}段目'), ''
        return f'{fx}出湯口', '出湯口 部分補修（はつり・れんが補修）', ''
    if t == 'turbine':
        return 'タービン室', {'first': 'タービン 開放（ケーシング上半 取外し）', 'last': '組立・ターニング・試運転'}.get(ph, pick(['ロータ・動翼 点検', '軸受 点検・隙間計測', 'ガバナ・保安装置 点検', '減速機 歯当たり確認', '発電機 点検（絶縁・軸受）'])), ''
    if t == 'ak':
        z = ['シャフト上部', 'シャフト中部', 'シャフト下部', '朝顔部', '炉床部', '羽口部']
        L = {'A1': ['足場組立 {k}段目', '足場 点検・手すり設置'], 'A2': ['旧耐火物 はつり（{z}）', 'はつり材 搬出 フレコン {n}袋', '旧耐火物 はつり（{z}）・鉄皮清掃'],
             'A3': ['旧ジャケット 撤去（{n}枚目）', '新ジャケット 据付（{k}段）', 'ジャケット 冷却水配管 接続', '新ジャケット 据付（{k}段）・溶接'],
             'A4': ['鉄皮 肉厚測定・マーキング（{z}）', '鉄皮 当て板溶接（{z}）', '鉄皮 補修部 PT 検査'],
             'A5': ['旧ポンプ・電動機 撤去', '新ポンプ 据付・芯出し', '配管接続・試運転 振動 {v}mm/s'],
             'A6': ['旧羽口 取外し（No.{a}〜{b}）', '新羽口装置 取付（No.{a}〜{b}）', '羽口 冷却水 通水・漏水試験'],
             'A7': ['温度計 撤去', '熱電対 取付（{a}〜{b}点）', '配線・ループ確認'],
             'A8': ['炉床 れんが築造 {k}段目', '朝顔部 れんが築造', 'シャフト部 れんが築造 {k}段目', 'キャスタブル 施工（{z}）', '目地確認・寸法測定'],
             'A8-2': ['出湯口 れんが築造', '出湯口 スリーブ据付', '出湯口 キャスタブル施工'], 'A9': ['冷却水系 耐圧試験 1.0MPa 保持'], 'A10': ['足場解体 {k}段目', '足場解体・搬出'],
             'A12': ['性能確認運転 負荷 {pct}%', '排ガス・スラグ性状 確認', '性能確認運転 定格'],
             'P2': ['旧{eq} 撤去', '旧配管 撤去'], 'P3': ['基礎 はつり', '基礎 アンカー打設', '基礎 コンクリート打設・養生'], 'P4': ['{eq} 据付', '{eq} 芯出し・据付'],
             'P5': ['切替弁 据付（{a}〜{b}）', '酸素配管 脱脂・据付'], 'P6': ['制御盤 据付', '電源・信号 配線', 'インバータ盤 据付・配線'], 'P7': ['試運転 製品酸素 {pct}% 確認', '切替サイクル 調整'],
             'D2': ['旧 OPS・EWS 撤去', '新 OPS・EWS・FCS-C 据付', '新盤 電源投入'], 'D3': ['1号 I/O 切替（{a}〜{b} 盤）', '1号 FCS 切替・電源投入'], 'D4': ['2号 I/O 切替（{a}〜{b} 盤）', '2号 FCS 切替・電源投入'],
             'D5': ['共通 I/O 切替'], 'D6': ['ループチェック {n}点', 'ループチェック 不具合点 手直し'], 'D7': ['インターロック試験（{z2}）'],
             'K-1': ['鉄皮 外面 UT（{z}）'], 'K-2': ['FCS・I/O盤 基礎 施工', 'ケーブルラック 取付'], 'K-3': ['ケーブル 先行敷設（{n}条）']}
        ph_ = L.get(it['no'], [it['name']])
        eqs = [U(x)['name'] for x in it['units']]
        tx = ph_[k % len(ph_)].format(k=k + 1, z=z[k % len(z)], n=r.randint(3, 9), v=r.randint(14, 22) / 10, a=1 + (k * 2) % 8, b=2 + (k * 2) % 8,
                                      pct=min(100, 60 + k * 8), eq=eqs[k % len(eqs)] if eqs else '', z2=['送風・給じん', '出湯口・水砕', 'ボイラ・IDF', 'PSA・共通'][k % 4])
        return loc, tx, ''
    if t.startswith('ev'):
        names = [fn(x) for x in sub]
        if not names: return loc, it['name'], ''
        u0 = U(sub[0]); s0 = H.truth(u0).get(p['fy'], ('○', ''))[0]
        if t == 'ev_EL':
            verb = EL.get(u0['kind'], '点検')
            if s0 == '●': verb = '更新（据付・配線・試運転）'
            if u0['cat'] == 'analyzer' and s0 == '◎': verb = 'メーカー点検・校正'
        elif t == 'ev_VL': verb = '分解整備（グランド・弁座）' if s0 == '◎' else '取替え' if s0 == '●' else '作動確認'
        elif t == 'ev_ST': verb = {'△': '補修（肉盛・塗装）', '○': '内部点検', '◇': '肉厚測定', '●': '更新', '◎': '整備'}.get(s0, '点検')
        else: verb = pick(VERB.get(s0, ['点検']))
        if r.random() < 0.05: find = r.choice(['軸受に軽い異音、次回確認', 'シール部 にじみ、増締め', '塗装剥離 一部、タッチアップ', 'ボルト 1本 緩み、増締め'])
        return fx + u0['subname'], '・'.join(names) + ' ' + verb, find
    return loc, it['name'], ''

# ------------------------------------------------------------------ 書き出し
def write(path, p, title, staff, city, hand=None, start=None, total=None):
    """hand：{date: 見本の手書きの日報 dict}（見本の日は手書きの行を使い、出来事の行だけ足す）"""
    fm = p.get('dfmt') or fmt(p['fy'])
    r = random.Random(f'dl-{p["key"]}')
    wb = Workbook(); wb.remove(wb.active)
    props(wb, f'【架空】{title} 作業日報', f'千草環境システム株式会社 {staff}')
    rc = Rec(path.split('/')[-1])
    start = start or p['start']
    total = total or (p['end'] - start).days + 1
    d = start; n = 0
    while d <= min(p['end'], dt.date(2026, 11, 13)):
        if d.weekday() == 6 and not (hand and d in hand): d += dt.timedelta(1); continue
        rows = day_rows(p, d, r)
        if hand and d in hand:
            hd = hand[d]
            hrows = []
            items = {it['no']: it for it in p['items']}
            for row in hd['rows']:
                no, loc, txt, co, nn, a, b = row[:7]; find = row[8] if len(row) > 8 else ''
                it = items.get(no)
                hrows.append(dict(it=it, loc=loc, txt=txt, co=co, n=nn, a=a, b=b, find=find, units=(it['units'][:1] if it else []), no=no))
            rows = hrows + [x for x in rows if x['no'].startswith(('E', 'T'))]
            meta = hd
        else:
            meta = None
        if not rows: d += dt.timedelta(1); continue
        n += 1
        w = meta['w'] if meta else weather(d, r)
        ws = wb.create_sheet(f'{d.month:02d}{d.day:02d}' if fm == 'p2' else f'{d.month}.{d.day}')
        if fm == 'p1a': sheet_p1a(ws, p, d, rows, w, title, staff, city, rc, n)
        elif fm == 'p1b': sheet_p1b(ws, p, d, rows, w, title, staff, city, rc, meta, r)
        else: sheet_p2(ws, p, d, rows, w, title, staff, city, rc, meta, start, total)
        d += dt.timedelta(1)
    if not wb.sheetnames: return None
    wb.save(path)
    return rc

def wl(x): return [w for w in (x['it'] or {}).get('wids', []) if any(U(t)['tag'] == t for t in x['units'])] if x['it'] else []

def link(x):
    if not x['it']: return []
    wids = x['it'].get('wids', []); units = x['it']['units']
    return [w for w, t in zip(wids, units) if t in x['units']]

def sheet_p1a(ws, p, d, rows, w, title, staff, city, rc, n):
    widths(ws, [12, 56, 8])
    merge(ws, 1, 1, 1, 3, '【架空】作業日報（班別）', font(16, True), AC, b=None)
    put(ws, 2, 1, f'{title}　{p["part"]}', font(9), AL)
    put(ws, 3, 1, jdate(d) + f'　天候 {w}', font(10, True), AL); put(ws, 3, 3, f'No.{n}', font(9), AR)
    for c, t in enumerate(['班', '作　業　内　容', '人員'], 1): put(ws, 5, c, t, font(9, True), AC, B, fill('EDEDED'))
    r = 6
    by = {}
    for x in rows: by.setdefault(x['it']['team'] if x['it'] else '機械', []).append(x)
    for team, xs in by.items():
        for i, x in enumerate(xs):
            put(ws, r, 1, team if i == 0 else '', font(9), AC, box(hair, thin, thin, hair))
            txt = f"{x['loc']}　{x['txt']}" + ('' if x['co'] == '千草環境' else f"（{x['co']}）")
            put(ws, r, 2, txt, font(9), ALW, box(hair, thin, thin, hair)); put(ws, r, 3, f"{x['n']}名", font(9), AC, box(hair, thin, thin, hair))
            ws.row_dimensions[r].height = 22
            rc.add(ws.title, r, 'item', dict(date=d, weather=w, blockTeam=team, content=txt, persons=x['n']), link(x)); acc(link(x), x['n'] * 7.5)
            r += 1
    put(ws, r, 2, '計', font(9, True), AR, B); put(ws, r, 3, f"{sum(x['n'] for x in rows)}名", font(9, True), AC, B)
    rc.add(ws.title, r, 'total', dict(date=d, persons=sum(x['n'] for x in rows)))
    fx = [x['find'] for x in rows if x['find']]
    put(ws, r + 2, 1, '特記', font(9, True), AC, B, fill('EDEDED')); merge(ws, r + 2, 2, r + 2, 3, '　'.join(fx) or '特になし', font(9), ALW, B)
    put(ws, r + 4, 2, f'SV {staff}　　市 {city}', font(9), AR)
    page(ws, land=False, fitw=1, fith=1, footer=False)

def sheet_p1b(ws, p, d, rows, w, title, staff, city, rc, meta, r_):
    widths(ws, [9, 16, 42, 7, 7, 7])
    merge(ws, 1, 1, 1, 6, '【架空】作　業　日　報', font(18, True), AC, b=None)
    put(ws, 2, 6, f'No.{(d - p["start"]).days + 1}', font(9), AR)
    put(ws, 3, 1, '工事件名', font(9, True), AC, B, fill('EDEDED')); merge(ws, 3, 2, 3, 6, f"{title}（{p['part']}）", font(9), AL, B)
    put(ws, 4, 1, '日　付', font(9, True), AC, B, fill('EDEDED')); merge(ws, 4, 2, 4, 3, jdate(d), font(10), AL, B)
    t = meta['t'] if meta else f'{r_.randint(12, 28)}℃'
    put(ws, 4, 4, '天候', font(9, True), AC, B, fill('EDEDED')); put(ws, 4, 5, w, font(9), AC, B); put(ws, 4, 6, t, font(9), AC, B)
    for c, tt in enumerate(['班・会社', '作業場所', '作　業　内　容', '人員', '始', '終'], 1): put(ws, 6, c, tt, font(9, True), AC, B, fill('EDEDED'))
    r = 7; prev = None
    for i in range(max(14, len(rows))):
        if i < len(rows):
            x = rows[i]
            loc = '〃' if x['loc'] == prev else x['loc']; prev = x['loc'] if x['loc'] != '〃' else prev
            co = x['co'].replace('足場工業', '足場').replace('工業', '工')
            vals = [co, loc, x['txt'], x['n'], x['a'], x['b']]
            rc.add(ws.title, r, 'item', dict(date=d, weather=w, company=co, place=loc, content=x['txt'], persons=x['n'], start=x['a'], end=x['b']), link(x)); acc(link(x), x['n'] * hours_of(x['a'], x['b']))
        else: vals = [None] * 6
        for c, v in enumerate(vals, 1): put(ws, r, c, v, font(9), AL if c in (2, 3) else AC, box(hair, thin, thin, hair))
        ws.row_dimensions[r].height = 20; r += 1
    tot = {}
    for x in rows: tot[x['co']] = tot.get(x['co'], 0) + x['n']
    put(ws, r, 1, '人員計', font(9, True), AC, B, fill('EDEDED'))
    merge(ws, r, 2, r, 5, '　'.join(f'{k.replace("足場工業","足場").replace("工業","工")} {v}名' for k, v in tot.items()), font(9), AL, B)
    put(ws, r, 6, f'{sum(tot.values())}名', font(9, True), AC, B)
    rc.add(ws.title, r, 'total', dict(date=d, persons=sum(tot.values())))
    r += 2
    ky = meta['ky'] if meta else r_.choice(['上下作業禁止の徹底', '酸欠・CO 測定 2回 異常なし', '重量物の玉掛け確認', '火気使用時の消火器配置', '回転機の試運転時 立入禁止'])
    note = meta['note'] if meta else ('　'.join(x['find'] for x in rows if x['find']) or '特になし')
    nxt = meta['nxt'] if meta else '　'
    for lab, val, h in [('安全事項（KY）', ky, 34), ('特記事項', note, 48), ('翌日予定', nxt, 30)]:
        put(ws, r, 1, lab, font(9, True), AC, B, fill('EDEDED')); merge(ws, r, 2, r, 6, val, font(9), ALW, B); ws.row_dimensions[r].height = h; r += 1
    r += 1
    put(ws, r, 4, 'SV', font(8), AC, B); put(ws, r, 5, '市担当', font(8), AC, B); put(ws, r, 6, '確認', font(8), AC, B)
    put(ws, r + 1, 4, staff, font(10), AC, B); put(ws, r + 1, 5, city, font(10), AC, B); put(ws, r + 1, 6, None, None, None, B)
    put(ws, r + 1, 1, '千草環境システム㈱ あかね台事業所', font(8), AL)
    page(ws, land=False, paper=9, fitw=1, fith=1, margins=(0.6, 0.6, 0.6, 0.6), footer=False)

def sheet_p2(ws, p, d, rows, w, title, staff, city, rc, meta, start, total):
    widths(ws, [7, 16, 40, 11, 7, 8, 6, 8, 26])
    merge(ws, 1, 1, 1, 6, '【架空】作業日報', font(16, True), AL, b=None)
    put(ws, 1, 9, f'{(d - start).days + 1}日目 ／ {total}日', font(10, True), AR)
    put(ws, 2, 1, '工事件名', font(9, True), AC, B, fill('E2EFDA')); merge(ws, 2, 2, 2, 6, title + (f'（{p["part"]}）' if p.get('part') else ''), font(9), AL, B)
    put(ws, 2, 7, '天候', font(9, True), AC, B, fill('E2EFDA')); put(ws, 2, 8, w, font(9), AC, B); put(ws, 2, 9, jdate(d), font(10, True), AC, B)
    r = 4; tot = {}
    for x in rows:
        t = tot.setdefault(x['co'], [0, 0.0, x['a'], x['b']]); t[0] += x['n']; t[1] += round(hours_of(x['a'], x['b']) * x['n'], 1)
        t[2] = min(t[2], x['a'], key=lambda s: int(s.replace(':', ''))); t[3] = max(t[3], x['b'], key=lambda s: int(s.replace(':', '')))
    put(ws, r, 1, '集計', font(9, True), AC, B, fill('E2EFDA'))
    for c, t in enumerate(['会社', '人数（延べ）', '時間', '工数(人・h)'], 2): put(ws, r, c, t, font(9, True), AC, B, fill('E2EFDA'))
    for k, (co, (n, mh, a, b)) in enumerate(tot.items()):
        rr = r + 1 + k
        put(ws, rr, 2, co, font(9), AL, B); put(ws, rr, 3, f'{n} 名', font(9), AC, B); put(ws, rr, 4, f'{a}〜{b}', font(9), AC, B); put(ws, rr, 5, round(mh, 1), font(9), AR, B, nf='0.0')
        rc.add(ws.title, rr, 'subtotal', dict(date=d, weather=w, company=co, persons=n, time=f'{a}〜{b}', manHours=round(mh, 1)))
    rk = r + 1 + len(tot)
    put(ws, rk, 2, '計', font(9, True), AC, B, fill('F2F2F2')); put(ws, rk, 3, f"{sum(v[0] for v in tot.values())} 名", font(9, True), AC, B, fill('F2F2F2'))
    put(ws, rk, 4, None, None, None, B, fill('F2F2F2')); put(ws, rk, 5, round(sum(v[1] for v in tot.values()), 1), font(9, True), AR, B, fill('F2F2F2'), '0.0')
    rc.add(ws.title, rk, 'total', dict(date=d, persons=sum(v[0] for v in tot.values()), manHours=round(sum(v[1] for v in tot.values()), 1)))
    att = (meta or {}).get('att') or ('市 ' + city if any(x['it'] and x['it']['tmpl'] in ('hydro', 'legal', 'il', 'safety') for x in rows) else '—')
    note = (meta or {}).get('note') or ('　'.join(x['find'] for x in rows if x['find']) or '特になし')
    nxt = (meta or {}).get('nxt') or '—'
    put(ws, r, 6, '立会・検査', font(9, True), AC, B, fill('E2EFDA')); merge(ws, r, 7, r, 9, att, font(9), ALW, B)
    merge(ws, r + 1, 6, max(rk, r + 3), 6, '特記', font(9, True), AC, B, fill('E2EFDA')); merge(ws, r + 1, 7, max(rk, r + 3), 9, note, font(9), ALW, B)
    r = max(rk, r + 3) + 1
    put(ws, r, 6, '明日', font(9, True), AC, B, fill('E2EFDA')); merge(ws, r, 7, r, 9, nxt, font(9), ALW, B)
    r += 2
    for c, t in enumerate(['工程No.', '設備・機器', '作　業　内　容', '時間', '人数', '工数', '進捗'], 1): put(ws, r, c, t, font(9, True), AC, B, fill('E2EFDA'))
    merge(ws, r, 8, r, 9, '発見事項・指示', font(9, True), AC, B, fill('E2EFDA')); r += 1
    for co in tot:
        merge(ws, r, 1, r, 9, f'■ {co}', font(9, True), AL, B, fill('F2F2F2')); r += 1
        for x in rows:
            if x['co'] != co: continue
            mh = round(hours_of(x['a'], x['b']) * x['n'], 1)
            prog = ''
            if x['it'] and x['it']['act']:
                s0, e0 = x['it']['act']; prog = f'{min(100, round(((d - s0).days + 1) / ((e0 - s0).days + 1) * 100))}%'
            for c, v in enumerate([x['no'], x['loc'], x['txt'], f"{x['a']}〜{x['b']}", x['n'], mh, prog], 1):
                put(ws, r, c, v, font(9), ALW if c in (2, 3) else AC, box(hair, thin, thin, hair), nf='0.0' if c == 6 else None)
            merge(ws, r, 8, r, 9, x['find'], font(9, color='C00000' if x['find'] else None), ALW, box(hair, thin, thin, hair))
            ws.row_dimensions[r].height = 24
            x['mh'] = mh
            rc.add(ws.title, r, 'item', dict(date=d, weather=w, blockCompany=co, no=x['no'], place=x['loc'], content=x['txt'], time=f"{x['a']}〜{x['b']}",
                   persons=x['n'], manHours=mh, progress=prog, finding=x['find']), link(x)); acc(link(x), mh)
            r += 1
    put(ws, r + 1, 1, '※ 安全事項は KY 記録簿に記入', font(7, color='808080'), AL)
    put(ws, r + 1, 7, '作成', font(8), AC, B); put(ws, r + 1, 8, '市', font(8), AC, B)
    put(ws, r + 2, 7, staff, font(10), AC, B); put(ws, r + 2, 8, city, font(10), AC, B)
    page(ws, land=True, paper=9, fitw=1, fith=1, margins=(0.5, 0.5, 0.5, 0.5), footer=False)
