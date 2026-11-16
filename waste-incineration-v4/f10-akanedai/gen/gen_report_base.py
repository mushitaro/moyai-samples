# -*- coding: utf-8 -*-
"""工事報告書（文字のある PDF、reportlab で作る。メーカーの報告書作成ツールの出力という設定）
旧：p1 様式（〜2019。H30 は p2 が p1 様式で作成）。文章型：表紙→概要→工程→作業内容（章立ての文章）→使用部品→所見。
新：p2 2020 様式（千草の報告書作成ツールの出力）。1 枚目は「系統ごとの件数（作業区分別・判定別）」の集計表と総括、
    2 枚目は追加作業と機器シートの索引、以降は 1 機器 1 枚の定型シート（左に記録、右に写真代替図 2 枚）。原価は載せない。"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether)
from reportlab.lib.styles import ParagraphStyle
from reportlab.graphics.shapes import Drawing, Rect, Line, String, Circle, Polygon
import jobs, world

pdfmetrics.registerFont(TTFont('G', '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'))
pdfmetrics.registerFont(TTFont('GP', '/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf'))
S = lambda sz=9, ld=None, al=0, f='G', **k: ParagraphStyle('x', fontName=f, fontSize=sz, leading=ld or sz * 1.5, alignment=al, **k)
P = lambda t, st=None: Paragraph(t, st or S())

def tbl(data, widths, head=1, sz=8, zebra=False, extra=None):
    d = [[P(str(c), S(sz, al=1 if i < head else 0)) if not hasattr(c, 'wrap') else c for c in row] for i, row in enumerate(data)]
    t = Table(d, colWidths=[w * mm for w in widths], repeatRows=head)
    st = [('GRID', (0, 0), (-1, -1), 0.5, colors.black), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
          ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2)]
    if head: st.append(('BACKGROUND', (0, 0), (-1, head - 1), colors.HexColor('#D9E1F2')))
    if extra: st += extra
    t.setStyle(TableStyle(st))
    return t

def photo(caption, kind='box', w=80, h=52):
    """写真代替図（線画）"""
    d = Drawing(w * mm, h * mm)
    d.add(Rect(0, 0, w * mm, h * mm, strokeColor=colors.black, fillColor=None, strokeWidth=0.6))
    W, H = w * mm, h * mm
    if kind == 'pipe':
        for i in range(5):
            d.add(Line(W * 0.12 + i * W * 0.17, H * 0.15, W * 0.12 + i * W * 0.17, H * 0.85, strokeWidth=4, strokeColor=colors.grey))
        d.add(Circle(W * 0.46, H * 0.3, 9, strokeColor=colors.red, fillColor=None, strokeWidth=1.2))
    elif kind == 'tuyere':
        d.add(Polygon([W * 0.2, H * 0.4, W * 0.7, H * 0.45, W * 0.7, H * 0.55, W * 0.2, H * 0.6], strokeColor=colors.black, fillColor=colors.lightgrey))
        d.add(Circle(W * 0.75, H * 0.5, 10, strokeColor=colors.black, fillColor=None))
    elif kind == 'jacket':
        d.add(Rect(W * 0.25, H * 0.2, W * 0.5, H * 0.6, strokeColor=colors.black, fillColor=colors.whitesmoke))
        d.add(Line(W * 0.4, H * 0.35, W * 0.48, H * 0.6, strokeColor=colors.red, strokeWidth=1.2))
        d.add(Line(W * 0.55, H * 0.3, W * 0.6, H * 0.5, strokeColor=colors.red, strokeWidth=1.2))
    elif kind == 'brick':
        for rr in range(5):
            for cc in range(6):
                d.add(Rect(W * 0.1 + cc * W * 0.13 + (rr % 2) * W * 0.06, H * 0.15 + rr * H * 0.13, W * 0.12, H * 0.11, strokeColor=colors.black, fillColor=None, strokeWidth=0.4))
    else:
        d.add(Rect(W * 0.2, H * 0.2, W * 0.6, H * 0.6, strokeColor=colors.black, fillColor=None))
        d.add(Line(W * 0.2, H * 0.2, W * 0.8, H * 0.8, strokeColor=colors.grey))
    d.add(String(4, 4, '写真代替図（見本用の線画）', fontName='G', fontSize=6))
    return Table([[d], [P(caption, S(8, al=1))]], colWidths=[w * mm])

class Doc(BaseDocTemplate):
    def __init__(self, path, title, author, header, form=None):
        super().__init__(path, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=20 * mm, bottomMargin=16 * mm,
                         title=title, author=author, subject='【架空】見本', creator='千草環境システム 報告書作成ツール')
        self.hdr = header; self.form = form
        fr = Frame(self.leftMargin, self.bottomMargin, self.width, self.height, id='f')
        self.addPageTemplates([PageTemplate(id='p', frames=[fr], onPage=self.deco)])
    def deco(self, c, doc):
        c.saveState(); c.setFont('G', 7)
        c.drawString(16 * mm, A4[1] - 10 * mm, self.hdr)
        if self.form: c.drawRightString(A4[0] - 16 * mm, A4[1] - 10 * mm, self.form)
        c.line(16 * mm, A4[1] - 12 * mm, A4[0] - 16 * mm, A4[1] - 12 * mm)
        c.drawCentredString(A4[0] / 2, 9 * mm, f'- {doc.page} -')
        c.drawRightString(A4[0] - 16 * mm, 9 * mm, '千草環境システム株式会社')
        c.restoreState()

def form_head(st, title, form):
    st.append(Table([[P(title, S(13)), P(form, S(7, al=2))]], colWidths=[128 * mm, 50 * mm],
                    style=[('LINEBELOW', (0, 0), (-1, 0), 1.2, colors.HexColor('#2E7D32'))]))
    st.append(Spacer(1, 2 * mm))

