# -*- coding: utf-8 -*-
"""Render key worker PDFs to PNG for OCR / visual read."""
from __future__ import annotations

import sys
from pathlib import Path

import fitz

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(r"D:\projects\Фурик")
docs = ROOT / "документы-работника"
out = ROOT / "_tools" / "_ocr_preview"
out.mkdir(parents=True, exist_ok=True)

for p in sorted(docs.glob("*.pdf")):
    print("PDF:", p.name)

# Map by keywords in filename
rules = [
    ("patent", ["МИНИСТЕРСТВО", "ВНУТРЕНН", "ФЕДЕРАЦ"]),
    ("bank", ["Солидарность", "СОЛИДАРНОСТЬ", "КБ"]),
    ("passport", ["ZBEKISTON", "O'ZBEKISTON", "RESPUBLIKASI"]),
    ("mig", ["Departure", "Выезд"]),
]

picked: dict[str, Path] = {}
for p in docs.glob("*.pdf"):
    name = p.name
    for key, kws in rules:
        if any(k in name for k in kws):
            picked.setdefault(key, p)

print("PICKED:", {k: v.name for k, v in picked.items()})

for key, pdf in picked.items():
    doc = fitz.open(pdf)
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
        dest = out / f"{key}_p{i + 1}.png"
        pix.save(dest)
        print("WROTE", dest.name, pix.width, "x", pix.height)
    doc.close()

print("DONE")
