# -*- coding: utf-8 -*-
"""見た目の確認：LibreOffice（自分専用プロファイル file:///tmp/lo_F08）で PDF → pdftoppm で PNG。"""
import os
import subprocess
import sys
import tempfile

PROFILE = "file:///tmp/lo_F08"


def to_pdf(path):
    tmp = tempfile.mkdtemp(prefix="f08pdf_")
    subprocess.run(["soffice", f"-env:UserInstallation={PROFILE}", "--headless",
                    "--convert-to", "pdf", "--outdir", tmp, path],
                   check=True, capture_output=True, timeout=300)
    return os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + ".pdf")


def pages(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    for ln in out.splitlines():
        if ln.startswith("Pages:"):
            return int(ln.split()[1])
    return 0


def to_png(pdf, base, page_list, dpi=80):
    """page_list のページを base.png（1 枚）または base_p1.png… として書く。"""
    made = []
    for k, pg in enumerate(page_list, 1):
        stem = base if len(page_list) == 1 else f"{base}_p{k}"
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", str(pg), "-l", str(pg),
                        "-singlefile", pdf, stem], check=True)
        made.append(stem + ".png")
    return made


def render(path, page_list=None, dpi=80):
    pdf = path if path.endswith(".pdf") else to_pdf(path)
    n = pages(pdf)
    pl = [p for p in (page_list or [1]) if p <= n] or [1]
    # PNG は「帳票と同名＋.png」（例 xxx.xlsx.png）
    return to_png(pdf, path, pl, dpi), n


if __name__ == "__main__":
    for p in sys.argv[1:]:
        print(render(p))


def render_csv(path, dpi=80):
    """CSV は Excel で開いた見た目を確かめる：cp932 で読み、一時の xlsx（横・幅合わせ）にして描く。"""
    import csv as _csv
    from openpyxl import Workbook
    from openpyxl.styles import Font
    with open(path, encoding="cp932", newline="") as f:
        rows = list(_csv.reader(f))
    wb = Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    for col in ws.columns:
        w = max(len(str(c.value or "")) for c in col[2:]) if len(col) > 2 else 8
        ws.column_dimensions[col[0].column_letter].width = min(max(w * 1.6, 5), 46)
    for row in ws.iter_rows():
        for c in row:
            c.font = Font(name="IPAPGothic", size=8)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    tmp = tempfile.mkdtemp(prefix="f08csv_")
    x = os.path.join(tmp, os.path.basename(path) + ".xlsx")
    wb.save(x)
    pdf = to_pdf(x)
    return to_png(pdf, path, [1], dpi), pages(pdf)
