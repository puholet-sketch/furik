# -*- coding: utf-8 -*-
from pathlib import Path
import json

ROOT = Path(r"D:\projects\Фурик")
print("=== ROOT PDFs ===")
total = 0
for p in sorted(ROOT.glob("*.pdf")):
    print(f"{p.stat().st_size/1024/1024:8.2f}  {p.name!r}")
    total += p.stat().st_size
print("TOTAL_MB", total/1024/1024)

print("\n=== compressed ===")
total2 = 0
for p in sorted((ROOT/"compressed").glob("*.pdf")):
    print(f"{p.stat().st_size/1024/1024:8.2f}  {p.name!r}")
    total2 += p.stat().st_size
print("TOTAL_MB", total2/1024/1024)

print("\n=== pdf_originals ===")
total3 = 0
for p in sorted((ROOT/"pdf_originals").glob("*.pdf")):
    print(f"{p.stat().st_size/1024/1024:8.2f}  {p.name!r}")
    total3 += p.stat().st_size
print("TOTAL_MB", total3/1024/1024)
print("TARGET", 19.9)
print("OK_ROOT", total <= 19.9*1024*1024)
print("OK_COMPRESSED", total2 <= 19.9*1024*1024)
