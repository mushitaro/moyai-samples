# -*- coding: utf-8 -*-
"""F05【架空】やまぶき広域処理組合 第二工場 : 共通定義（施設・人・会社・書式ヘルパ）"""
import os, random, datetime as dt
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill

SEED = 20261113
TODAY = dt.date(2026, 11, 13)
FULL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.path.dirname(FULL)
OUT = os.path.join(FULL, 'files')
WORK = os.path.join(FULL, 'work')
ANS = os.path.join(FULL, 'answer')
NOTES = os.path.join(BASE, 'notes')
LO_PROFILE = 'file:///tmp/lo_F05'

FAC = '【架空】やまぶき広域処理組合 第二工場'
FAC_S = 'やまぶき広域処理組合第二工場'        # 契約件名で使う詰めた書き方
UNION = '【架空】やまぶき広域処理組合'
FN = '【架空】やまぶき広域処理組合第二工場'    # ファイル名の頭

# 会社（すべて架空）
MAKER = '【架空】朝凪炉工株式会社'
MAKER_S = '朝凪炉工'
OPCO = '【架空】やまぶき環境サービス株式会社'      # 運転業務委託（地元）
ELEC = '【架空】一般財団法人 みずなら電気保安協会'  # 電気保安（e1）
CRANE_INSP = '【架空】みさき起重機サービス有限会社'  # クレーン検査の外注（地元）
SCAFF = '【架空】有限会社 ひぐれ足場'
ROT = '【架空】あけぼの回転機株式会社'               # 回転機（動バランス）外注
INSTR = '【架空】東雲計器株式会社'                   # 分析計メーカー
SHOP = '【架空】株式会社 やまぶき機材商会'           # 地元商社
ANALYS = '【架空】株式会社 かやの環境分析センター'
FIRE = '【架空】ささゆり防災設備株式会社'
SCALE = '【架空】とまり計量器株式会社'
CEMENT = '【架空】しらかば資源化株式会社'
PREC = '【架空】一般社団法人 くるみ廃棄物技術協会'  # 精密機能検査

# 人（記号）
Y1, Y2 = 'y1', 'y2'      # 組合事務局 施設係
N1, N2, N3 = 'n1', 'n2', 'n3'  # メーカー SV / 基幹改良 現場代理人

FONT = 'IPAPGothic'
FONT_M = 'IPAGothic'

thin = Side(style='thin', color='000000')
hair = Side(style='hair', color='000000')
med = Side(style='medium', color='000000')
dot = Side(style='dotted', color='000000')
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)


def f(size=10, bold=False, color='000000', name=FONT, italic=False, strike=False, u=None):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic, strike=strike, underline=u)


def fill(hexc):
    return PatternFill('solid', fgColor=hexc)


AC = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL = Alignment(horizontal='left', vertical='center', wrap_text=True)
ALT = Alignment(horizontal='left', vertical='top', wrap_text=True)
AR = Alignment(horizontal='right', vertical='center')


def box(ws, r1, c1, r2, c2, side=thin, inner=True):
    """r1..r2, c1..c2 に罫線"""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            l = side if c == c1 else (thin if inner else None)
            rr = side if c == c2 else (thin if inner else None)
            t = side if r == r1 else (thin if inner else None)
            b = side if r == r2 else (thin if inner else None)
            cell.border = Border(left=l, right=rr, top=t, bottom=b)


def put(ws, r, c, v, font=None, al=None, fl=None, nf=None, border=None):
    cell = ws.cell(r, c, v)
    cell.font = font or f()
    if al is not None:
        cell.alignment = al
    if fl is not None:
        cell.fill = fl
    if nf:
        cell.number_format = nf
    if border is not None:
        cell.border = border
    return cell


def merge(ws, r1, c1, r2, c2, v=None, font=None, al=AC, fl=None, nf=None):
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    return put(ws, r1, c1, v, font, al, fl, nf)


def props(wb, title, creator, subject=''):
    wb.properties.title = title
    wb.properties.creator = creator
    wb.properties.subject = subject
    wb.properties.keywords = '【架空】試験用サンプル'
    wb.properties.created = dt.datetime(2026, 11, 13, 9, 0, 0)
    wb.properties.modified = dt.datetime(2026, 11, 13, 9, 0, 0)


def wareki(y):
    """西暦年度→令和の年（2019=R1）"""
    return y - 2018


def wa_date(d, style='R'):
    if style == 'R':
        return f'R{d.year-2018}.{d.month}.{d.day}'
    if style == 'K':
        return f'令和{d.year-2018}年{d.month}月{d.day}日'
    return d.isoformat()


WD = '月火水木金土日'


def wd(d):
    return WD[d.weekday()]


def daterange(a, b):
    d = a
    while d <= b:
        yield d
        d += dt.timedelta(days=1)


def setup_print(ws, landscape=True, paper=9, fit_w=1, fit_h=0, area=None, title_rows=None, margins=(0.4, 0.4, 0.5, 0.5)):
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_setup.paperSize = paper  # 9=A4, 8=A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_w
    ws.page_setup.fitToHeight = fit_h
    l, r, t, b = margins
    ws.page_margins.left = l; ws.page_margins.right = r
    ws.page_margins.top = t; ws.page_margins.bottom = b
    if area:
        ws.print_area = area
    if title_rows:
        ws.print_title_rows = title_rows
ALN = Alignment(horizontal='left', vertical='center', wrap_text=False)
