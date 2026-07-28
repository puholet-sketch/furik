# -*- coding: utf-8 -*-
"""Extract readable text from compressed PDFs for contract placeholders."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import fitz

ROOT = Path(r"D:\projects\Фурик")
OUT = ROOT / "_tools" / "pdf_text_extract.json"
sys.stdout.reconfigure(encoding="utf-8")

results = []
for p in sorted((ROOT / "compressed").glob("*.pdf")):
    doc = fitz.open(p)
    texts = []
    for i, page in enumerate(doc):
        t = page.get_text("text") or ""
        texts.append({"page": i + 1, "text": t[:4000]})
    full = "\n".join(x["text"] for x in texts)
    results.append(
        {
            "name": p.name,
            "pages": len(doc),
            "chars": len(full),
            "preview": full[:2500],
            "has_text": len(full.strip()) > 50,
        }
    )
    doc.close()
    print(f"=== {p.name} pages={results[-1]['pages']} chars={results[-1]['chars']} text={results[-1]['has_text']}")
    print(results[-1]["preview"][:800])
    print()

OUT.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
print("WROTE", OUT)
