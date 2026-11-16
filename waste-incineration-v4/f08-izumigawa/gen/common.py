# -*- coding: utf-8 -*-
"""F08【架空】いずみ川清掃センター 共通の定数と補助関数。"""
import datetime as dt
import os
import random

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))     # f08/full
OUT = os.path.join(BASE, "files")
ANS = os.path.join(BASE, "answer")
NOTES = os.path.join(BASE, "..", "notes")
SEED = 20261113
TODAY = dt.date(2026, 11, 13)          # 物語の「今日」
FAC = "いずみ川清掃センター"
FAC_K = "【架空】いずみ川清掃センター"
CITY = "【架空】いずみ川市"
FONT = "IPAPGothic"                    # LibreOffice で描く日本語フォント
FONT_M = "IPAGothic"

MAKER = "鷹羽エンジニアリング株式会社"     # メーカー（m1 の会社）
LOCAL = "いずみ川設備工業株式会社"        # 地元業者（g1 の会社）
OPER = "いずみ川環境サービス株式会社"     # 運転業務委託


def rng(tag=""):
    """資料ごとに種を分けた乱数（同じ命令で同じものができる）。"""
    return random.Random(f"{SEED}:{tag}")


def wareki_year(fy, style="H"):
    """年度を和暦で書く。style: 'H' → H22 / 'full' → 平成22年度 / 'R' → 令和の略記。"""
    if fy >= 2019:
        n = fy - 2018
        if style == "full":
            return "令和元年度" if n == 1 else f"令和{n}年度"
        return f"R{n}"
    n = fy - 1988
    if style == "full":
        return f"平成{n}年度"
    return f"H{n}"


def wareki_date(d, sep="."):
    """日付を和暦の略記（H22.9.6 / R5.10.16）で書く。改元前は平成。"""
    if d >= dt.date(2019, 5, 1):
        return f"R{d.year - 2018}{sep}{d.month}{sep}{d.day}"
    return f"H{d.year - 1988}{sep}{d.month}{sep}{d.day}"


def fy_of(d):
    return d.year if d.month >= 4 else d.year - 1


def out_path(name):
    os.makedirs(OUT, exist_ok=True)
    return os.path.join(OUT, name)


def set_props(wb, title):
    """Excel の文書のプロパティにも【架空】を書く。"""
    wb.properties.title = title
    wb.properties.creator = f"{CITY}（架空の作成者）"
    wb.properties.subject = f"{FAC_K} 試験用サンプル"
    wb.properties.keywords = "【架空】"
