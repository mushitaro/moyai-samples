# -*- coding: utf-8 -*-
"""日報の行を工程（タスク）から作る。見本の 2 年度の見本日は承認済みの行をそのまま使う。"""
import datetime as dt
from world2 import h, R, work_days, BYTAG, operator_word
from daily_data import H23_DAYS, R7_DAYS

PARTS = {
    'grate': ['火格子片', 'ピン・ボルト', '摺動面'], 'furnace': ['キャスタブル', 'アンカー', '耐火れんが'], 'fan': ['軸受', 'カップリング', 'グランドパッキン'],
    'pump': ['メカニカルシール', '軸受', 'ライナリング'], 'conv': ['チェーン', 'スプロケット', 'フライト'], 'valve': ['シリンダ', 'シールキット', 'ポジショナ'],
    'tube': ['管', 'プロテクタ', 'ヘッダ'], 'panel': ['端子', 'リレー', '冷却ファン'], 'crane': ['ワイヤロープ', 'シーブ', 'ブレーキライニング'],
    'other': ['パッキン', 'ボルト類', 'シール'],
}
def part_group(d):
    p = d['prof']
    for k, keys in (('grate', ('grate',)), ('furnace', ('furnace',)), ('fan', ('fan',)), ('pump', ('pump', 'bfp', 'hyd', 'comp')),
                    ('conv', ('conv', 'rv', 'feeder', 'ashext')), ('valve', ('valve',)), ('tube', ('tube', 'sh', 'drum', 'aph')),
                    ('panel', ('panel', 'transformer', 'itv')), ('crane', ('crane', 'hoist'))):
        if any(p.startswith(x) for x in keys):
            return k
    return 'other'

def _pick(lst, *key):
    return lst[h(*key) % len(lst)]

def phrase(d, k, t, i, n, x):
    """機器 d の作業（区分 k、内容 t）の i 日目／n 日の記事"""
    nm = d['name']
    pg = part_group(d)
    p = PARTS[pg]
    a, b = _pick(p, nm, x, 1), _pick(p, nm, x, 2)
    v = h(d['tag'], x) % 1000
    last = i == n - 1
    rot = pg in ('fan', 'pump', 'conv', 'crane')
    tt = t.split('（')[0]
    if k == '点検':
        if n == 1:
            opts = ['%s 開放点検、%s 状態確認、異常なし。復旧' % (nm, a), '%s 内部点検・清掃、%s に軽微な摩耗（経過観察）' % (nm, a)]
            if rot:
                opts.append('%s 点検（%s・%s）、試運転時振動 %.1fmm/s' % (nm, a, b, 0.8 + v / 500.0))
            if pg == 'panel':
                opts = ['%s 盤内清掃、端子増締め、絶縁抵抗 %dMΩ以上' % (nm, 10 + v % 90), '%s 点検、%s 外観・動作確認 良' % (nm, a)]
            return _pick(opts, d['tag'], x, 'p1')
        if i == 0:
            return '%s 開放、内部清掃、%s 外観点検' % (nm, a)
        return '%s %s 点検後 復旧、作動確認 良' % (nm, b) if last else '%s 点検継続（%s）' % (nm, b)
    if k == '計測':
        pts = 20 + v % 60
        what = '肉厚測定' if pg == 'tube' else ('摩耗計測' if pg in ('grate', 'conv') else '計測')
        if last:
            return '%s %s 完了（累計 %d 点、最小 %.1fmm）' % (nm, what, pts * n, 2.5 + v / 400.0)
        return '%s %s %d 点（%s）' % (nm, what, pts, _pick(['上段', '中段', '下段', '前列', '後列'], nm, x))
    if k in ('整備', '補修'):
        if n == 1:
            return _pick(['%s %s、作動確認 良' % (nm, tt), '%s %s。%s 状態確認、試運転 良' % (nm, tt, a)], d['tag'], x, 'o1')
        if i == 0:
            return '%s 分解・%s 取外し（%s）' % (nm, a, _pick(['摩耗あり', '状態良', '一部腐食'], nm, x)) if k == '整備' else \
                '%s 劣化部確認・マーキング、撤去 %d%%' % (nm, 30 + v % 40)
        if last:
            return '%s 組立・芯出し、単体試運転 良' % nm if k == '整備' else '%s 補修完了、外観・寸法検査 良' % nm
        return '%s %s（%s）' % (nm, tt, _pick(['施工中', '溶接', '取付', '仮組', '%s 含む' % b], nm, x, i))
    # 更新
    if n == 1:
        return '%s 既設撤去・新品据付、%s、試運転 良' % (nm, '配線接続' if pg == 'panel' else '芯出し')
    if i == 0:
        return '%s 既設撤去、据付面 清掃・墨出し' % nm
    if last:
        return '%s 試運転・調整 完了' % nm
    return '%s 新品据付・%s' % (nm, _pick(['芯出し', '配管接続', '配線接続', 'アンカー打設'], nm, x, i))

GENERIC = {
    '準備工': ['資機材搬入（2t車 %d台）', '仮設電源・照明 配置', '新規入場者教育 %d名', '工具・換気設備 配置'],
    '焼却停止': ['給じん停止 → 焼却停止（%s）。立会'],
    '炉冷却': ['IDF 運転継続で冷却。炉出口 %d℃'],
    '安全措置': ['O2 20.9%%・CO 0ppm 確認、電源ロック、マンホール %d箇所開放'],
    '足場組立': ['足場組立 %d%%', '炉内足場 段数追加、ボイラ側 組立'],
    '炉内清掃': ['クリンカ除去 約%.1ft 搬出', '側壁付着物 はつり除去'],
    'ボイラ清掃': ['水管付着ダスト掻落し・エアブロー', 'ボイラ第%d放射室 清掃'],
    '足場解体': ['足場解体 %d%%、残材搬出', '炉内最終清掃'],
    'マンホール閉鎖': ['マンホール閉鎖 %d箇所、回転機 単体試運転', '電動弁・ダンパ 作動試験'],
    '立上げ': ['昇温開始（%s）、耐火物乾燥', '昇温継続、炉出口 %d℃'],
    '試運転': ['負荷運転、排ガス値 確認', '定格負荷 連続運転 良'],
    '施工計画': ['施工計画書 作成・提出'], '資機材搬入': ['資機材搬入・養生'], '出来形': ['出来形・寸法検査（市立会）合格'],
    '復旧・清掃': ['復旧・清掃・搬出 完了'], '仮設': ['仮設足場・養生、既設撤去準備'], '単体試運転': ['単体試運転', '総合試運転 データ採取'],
    '完成検査': ['完成検査（市）受検'],
}
def generic_phrase(t, i, n, x, fy):
    for key, lst in GENERIC.items():
        if t['item'].startswith(key) or key in t['item'][:8]:
            s = lst[(i + h(t['item'], x)) % len(lst)]
            v = h(t['item'], x) % 100
            s = s.replace('%s', operator_word(fy)).replace('%.1f', '%.1f' % (1.0 + v / 40.0))
            if '%d%%' in s:
                s = s.replace('%d%%', '%d%%' % min(100, int(100 * (i + 1) / n)))
            if '%d' in s:
                if '℃' in s:
                    num = max(60, 400 - i * 120 + v)
                elif '台' in s:
                    num = 2 + v % 3
                elif '名' in s:
                    num = 3 + v % 6
                elif '箇所' in s:
                    num = 6 + v % 3
                else:
                    num = 1 + v % 3
                s = s.replace('%d', str(num), 1)
            return s.replace('%%', '%')
    return t['item'] + ' ' + ('完了' if i == n - 1 else '施工')

def place_of(t):
    if t.get('evs'):
        return BYTAG[t['evs'][0][0]]['name']
    for k, v in (('炉内', '炉内'), ('ボイラ', 'ボイラ'), ('足場', '炉内・ボイラ'), ('資機材', '資材置場'), ('中央', '中央制御室'), ('焼却停止', '中央制御室'),
                 ('冷却', '炉・ボイラ'), ('マンホール', '炉・ボイラ'), ('立上げ', '中央制御室'), ('試運転', '中央制御室')):
        if k in t['item']:
            return v
    return '現場'

def day_rows(fy, tasks, x, sample=None):
    """その日の日報の行。(No, 班, 号炉, 箇所, 内容, 社員, 協力, 時間, 進捗)"""
    if sample and x in sample:
        return sample[x]['rows']
    rows = []
    for t in tasks:
        if t['team'] in ('別', '事務') or t['kind'] == 'sep':
            continue
        wd = work_days(t)
        if x not in wd:
            continue
        i, n = wd.index(x), len(wd)
        if t.get('evs'):
            parts = []
            for (tag, idx) in t['evs'][:2]:
                d = BYTAG[tag]
                evs = list(d['act'].get(fy, [])) + [(k, tt, c, 'T') for (k, tt, c) in d.get('pending2026', [])] if fy == 2026 else list(d['act'].get(fy, []))
                if idx < len(evs):
                    k, tt, c, ct = evs[idx]
                else:
                    k, tt, c = '点検', t['item'], 0
                parts.append(phrase(d, k, tt, i, n, x))
            text = '。'.join(parts)
            if len(t['evs']) > 2:
                text += '。ほか %d 台' % (len(t['evs']) - 2)
        else:
            text = generic_phrase(t, i, n, x, fy)
        s, k = t['n']
        r = R('pp', t['no'], x)
        if k > 0 and r.random() < 0.2:
            k = max(0, k + r.choice([-1, 1]))
        hrs_fixed = None
        if t['team'] == '運転':
            s, k = (1, 0) if t['kind'] in ('stop', 'start') else t['n']
            hrs_fixed = 2.0 if t['kind'] in ('stop', 'start') else (8.0 if s else 0.0)
        hrs = 8.0
        if t['grp'] == '全停':
            hrs = 10.0 if i == 0 else 9.0
        elif r.random() < 0.15:
            hrs = r.choice([7.0, 8.5, 9.0])
        prog = 100 if i == n - 1 else int(round(100.0 * (i + 1) / n - r.choice([0, 0, 5, 10])))
        f = t['grp'] if t['grp'] in ('1号炉', '2号炉') else ('共通' if t['grp'] in ('共通', '全停') else (BYTAG[t['evs'][0][0]]['fname'] if t.get('evs') else '—'))
        if hrs_fixed is not None:
            hrs = hrs_fixed
        rows.append((t['no'], '運' if t['team'] == '運転' else t['team'], f, place_of(t), text, s, k, hrs, max(0, prog)))
    return rows

def report_days(tasks, upto=None):
    ds = set()
    for t in tasks:
        if t['team'] in ('別', '事務') or t['kind'] == 'sep':
            continue
        for x in work_days(t):
            if upto is None or x <= upto:
                ds.add(x)
    return sorted(ds)

def dd_for(fy, tasks, x, sample=None):
    rows = day_rows(fy, tasks, x, sample)
    if sample and x in sample:
        return sample[x]
    r = R('dd', fy, x)
    kys = ['足場上の工具落下防止', '開口部の墜落防止', '酸欠・残熱（連続測定器携帯）', '回転体への接近禁止', '重量物の吊り荷下立入禁止', '火気使用時の養生と消火器配置',
           '熱中症予防（WBGT 確認）', '停電範囲の確認と検電', '雨天時の滑り', '粉じん（防じんマスク着用）']
    return dict(no=0, weather=r.choice(['晴', '晴', '曇', '雨', '晴のち曇', '曇時々雨']), temp=str(r.randint(18, 33) if x.month in (6, 7, 8, 9) else r.randint(8, 22)),
                time='8:00〜17:00' if all(rr[7] <= 8.0 for rr in rows) else '8:00〜18:30', rows=rows, ky=r.choice(kys), hiyari='なし' if r.random() < 0.85 else '足元の段差でつまずきかけた → 養生追加',
                witness='なし' if r.random() < 0.6 else '市 担当 立会', note=r.choice(['', '', '別契約の作業区域と朝礼で調整。', '資材の納入あり（検収済み）。', '']),
                tomorrow='工程表どおり', finds=[], permits='—')
