# -*- coding: utf-8 -*-
from pathlib import Path
from html import escape
from openpyxl import load_workbook
from shutil import copy2

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "заказы-купер" / "реестр-перевыставление-Фурик.xlsx"
DOCS = ROOT / "docs"
copy2(XLSX, DOCS / "заказы-купер" / "реестр-перевыставление-Фурик.xlsx")


def rub(x: float) -> str:
    return f"{x:,.2f}".replace(",", "\u00a0").replace(".", ",")


wb = load_workbook(XLSX)
ws = wb.active
rows = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r[0]:
        continue
    d, no, pos, qty, unit, sm, src = r[:7]
    rows.append(
        (
            str(d)[:10],
            no or "",
            pos or "",
            float(sm),
            (src or "").replace(" / Купер", "").strip(),
        )
    )

total = sum(r[3] for r in rows)
trs = []
for d, no, pos, sm, store in rows:
    trs.append(
        "<tr>"
        f'<td data-label="Дата">{escape(d)}</td>'
        f'<td data-label="№ заказа"><code>{escape(str(no))}</code></td>'
        f'<td data-label="Товар">{escape(pos)}</td>'
        f'<td data-label="Сумма" class="num">{rub(sm)}\u00a0₽</td>'
        f'<td data-label="Магазин">{escape(store)}</td>'
        "</tr>"
    )

fragment = "\n".join(trs)
meta = f"{len(rows)}|{rub(total)}"
(DOCS / "_reg_fragment.html").write_text(fragment, encoding="utf-8")
(DOCS / "_reg_meta.txt").write_text(meta, encoding="utf-8")
print(meta)
