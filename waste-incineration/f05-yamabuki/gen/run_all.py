# -*- coding: utf-8 -*-
"""F05 本作り：python3 run_all.py（full/gen で実行）。
full/files/ に資料、full/answer/ に正解（files/<stem>.json と world.json）を作り、check.py を通す。"""
import os, glob, json, subprocess, sys, datetime as dt
from common import OUT, WORK, ANS, FULL, FAC, SEED, TODAY
os.makedirs(OUT, exist_ok=True); os.makedirs(WORK, exist_ok=True); os.makedirs(os.path.join(ANS, 'files'), exist_ok=True)
for p in glob.glob(os.path.join(OUT, '*')) + glob.glob(os.path.join(ANS, 'files', '*')):
    os.remove(p)

import works as W
import rec
import gen_plan_full, gen_kotei_full, gen_nippo_full, gen_report_full, gen_order_full, gen_ledger_full
from world import MACHINES, old_machines, SYSNAME
import plan_data as PD

files = []
files += [os.path.basename(p) for p in gen_plan_full.build_all()]
files += gen_kotei_full.build_all()
files += gen_nippo_full.build_all()
files += gen_report_full.build_all()
files += gen_order_full.build_all()
files += gen_ledger_full.build_all()
rec.write_all()


def mandays(w):
    if w['contract'] == 'KIKAN':
        return sum(w['persons'] or 0 for s, e in w['act'] for d in W.workdays(s, e))
    tot = 0
    for s, e in w['act']:
        for d in W.workdays(s, e):
            tot += w['persons'] or 0
    return tot


om = old_machines()
machines = []
for m in MACHINES:
    machines.append(dict(code=m['tag'], 台帳名=m['name'], 旧台帳名=m['old'] if not m.get('old_same') else m['name'], 号炉={0: '共通', 1: '1号炉', 2: '2号炉'}[m['furnace']],
                         系統=SYSNAME[m['sys']], 設置=m['inst'], 更新年度=m['renew'], 常用予備=m['standby'] or None, 状態=m['status'],
                         呼び名={'機器台帳R3旧版・y1計画': m['old'] if not m.get('old_same') else m['name'], '機器台帳R5刷新・y2計画': m['name'], '工程表・日報': W.short(m)},
                         重要度_y2=PD.importance(m)))
for o in om:
    if o.get('olddcs'):
        machines.append(dict(code=None, 台帳名=None, 旧台帳名=o['old'], 号炉='共通', 系統='制御設備', 設置='2001-03', 状態='R5基幹改良で撤去（2024-02）',
                             呼び名={'機器台帳R3旧版・y1計画': o['old']}))
works = []
for w in W.WORKS:
    c = W.CONTRACTS.get(w['contract'])
    works.append(dict(id=w['id'], contract=w['contract'], 契約件名=(c or {}).get('name', w['name']), vendor=w['vendor'], 年度=w['year'], 停止=w['event'],
                      作業=w['name'], machines=w['machines'], type=w['wtype'], start=min(s for s, e in w['act']).isoformat(), end=max(e for s, e in w['act']).isoformat(),
                      plan=[[s.isoformat(), e.isoformat()] for s, e in w['plan']], manDays=mandays(w), 発見=w['finding'] or None,
                      部品=[list(p) for p in w['parts']] or None, 追加=bool(w['rev']), unrecorded=w.get('unrecorded', False)))
other_contracts = {w['contract']: dict(key=w['contract'], name=w['name'], vendor=w['vendor']) for w in W.OTHER}
notdone = []
for y, lst in W.TEIKI_SKIP.items():
    for t, s, why in lst:
        notdone.append(dict(年度=y, code=t, 計画記号=s, 理由=why))
world = dict(
    profile=dict(id='f05', 施設=FAC, 構成='3町村の一部事務組合', 竣工='2001-03', 能力='30t/16h×2炉（准連続、並行揺動式ストーカ）', ボイラ発電='なし（水噴射式減温塔）',
                 排ガス処理='消石灰・活性炭吹込＋バグフィルタ', 改良='2022–2023 基幹的設備改良工事（火格子・減温塔・BF・DCS）', 対象期間='2021–2025 年度（2026 年度は計画のみ）',
                 機器台数=dict(合計=len(MACHINES), 炉別=sum(1 for m in MACHINES if m['furnace']), 共通=sum(1 for m in MACHINES if not m['furnace']), 旧台帳=len(om)),
                 記録の時点='2026-11-13', 乱数の種=SEED,
                 担当=dict(組合事務局='y1 2021–2023 → y2 2024–', メーカーSV='n1 2021–2023 → n2 2024–', 基幹改良現場代理人='n3 2022–2023')),
    contracts=list(W.CONTRACTS.values()) + list(other_contracts.values()),
    machines=machines, works=works, links=rec.LINKS,
    unrecordedWorks=W.UNRECORDED,
    planNotDone=notdone,
    planGaps=gen_plan_full.GAPS + [
        dict(file='中長期整備計画 R3〜R5 版（y1）', gap='第二種圧力容器・ホイスト・DXN作業環境測定は運転委託・別契約で実施しているが計画に無い'),
        dict(file='中長期整備計画 R6改訂・R7版（y2）', gap='新DCS の年次点検は毎年「点」だが契約・帳票が無い（瑕疵担保の無償点検）'),
    ],
)


def _j(o):
    if isinstance(o, (dt.date, dt.datetime)):
        return o.isoformat()
    if isinstance(o, tuple):
        return list(o)
    raise TypeError(type(o))


with open(os.path.join(ANS, 'world.json'), 'w', encoding='utf-8') as fp:
    json.dump(world, fp, ensure_ascii=False, indent=1, default=_j)
print(len(files), 'files')
import shutil
v1 = os.path.join(FULL, 'work', 'answer_v1')
shutil.rmtree(v1, ignore_errors=True)
shutil.copytree(ANS, v1)
subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), 'to_v2.py')], check=True)
r = subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), 'check.py')])
sys.exit(r.returncode)
