# -*- coding: utf-8 -*-
"""Inventory PDFs, backup originals, compress to total <= 19.9 MiB."""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(r"D:\projects\Фурик")
ORIGINALS = ROOT / "pdf_originals"
COMPRESSED = ROOT / "compressed"
TARGET = int(19.9 * 1024 * 1024)
REPORT = ROOT / "PDF_COMPRESSION_REPORT.json"

sys.stdout.reconfigure(encoding="utf-8")


def list_pdfs(base: Path) -> list[Path]:
    skip = {ORIGINALS.name, COMPRESSED.name, "_tools"}
    out: list[Path] = []
    for p in base.rglob("*.pdf"):
        if any(part in skip for part in p.parts):
            continue
        # also skip pdf_originals nested
        if ORIGINALS in p.parents or COMPRESSED in p.parents:
            continue
        out.append(p)
    for p in base.rglob("*.PDF"):
        if ORIGINALS in p.parents or COMPRESSED in p.parents:
            continue
        if p not in out:
            out.append(p)
    return sorted(out, key=lambda x: x.name.lower())


def compress_with_pymupdf(src: Path, dst: Path, jpeg_quality: int, dpi: int) -> int:
    import fitz

    doc = fitz.open(src)
    # Rebuild as image-compressed PDF page by page for scanned docs
    out = fitz.open()
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        # JPEG bytes
        img_bytes = pix.tobytes("jpeg", jpg_quality=jpeg_quality)
        # Create page with image
        rect = page.rect
        # Scale page size to image aspect at chosen dpi
        w = pix.width * 72.0 / dpi
        h = pix.height * 72.0 / dpi
        new_page = out.new_page(width=w, height=h)
        new_page.insert_image(new_page.rect, stream=img_bytes)
    dst.parent.mkdir(parents=True, exist_ok=True)
    out.save(dst, garbage=4, deflate=True, clean=True)
    out.close()
    doc.close()
    return dst.stat().st_size


def try_pikepdf_light(src: Path, dst: Path) -> int:
    import pikepdf

    dst.parent.mkdir(parents=True, exist_ok=True)
    with pikepdf.open(src) as pdf:
        pdf.save(
            dst,
            compress_streams=True,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
            recompress_flate=True,
        )
    return dst.stat().st_size


def main() -> None:
    pdfs = list_pdfs(ROOT)
    inventory = []
    for p in pdfs:
        inventory.append({"name": p.name, "path": str(p), "size": p.stat().st_size})
    (ROOT / "_file_inventory.json").write_text(
        json.dumps(inventory, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    total0 = sum(i["size"] for i in inventory)
    print(f"PDF_COUNT={len(pdfs)}")
    print(f"PDF_TOTAL_MB={total0/1024/1024:.3f}")
    print(f"TARGET_BYTES={TARGET}")
    for i in inventory:
        print(f"  {i['size']/1024/1024:.2f} MB | {i['name']}")

    ORIGINALS.mkdir(exist_ok=True)
    COMPRESSED.mkdir(exist_ok=True)

    # Backup originals once
    for p in pdfs:
        dest = ORIGINALS / p.name
        if not dest.exists():
            shutil.copy2(p, dest)
            print(f"BACKED_UP {p.name}")
        else:
            print(f"BACKUP_EXISTS {p.name}")

    n = len(pdfs) or 1
    # Equal budget per file with small margin
    per_file_budget = int(TARGET * 0.98 / n)

    # Quality ladder: try lighter first, then heavier
    ladder = [
        (85, 120),
        (75, 110),
        (65, 100),
        (55, 90),
        (45, 80),
        (40, 72),
        (35, 70),
    ]

    results = []
    for p in pdfs:
        orig_size = p.stat().st_size
        out_path = COMPRESSED / p.name
        best_size = None
        best_params = None

        # First try pikepdf if already small enough path
        try:
            sz = try_pikepdf_light(p, out_path)
            if sz <= per_file_budget:
                best_size = sz
                best_params = "pikepdf"
        except Exception as e:
            print(f"PIKEPDF_FAIL {p.name}: {e}")

        if best_size is None or best_size > per_file_budget:
            for q, dpi in ladder:
                try:
                    sz = compress_with_pymupdf(p, out_path, q, dpi)
                    print(f"TRY {p.name} q={q} dpi={dpi} -> {sz/1024/1024:.2f} MB")
                    if best_size is None or sz < best_size:
                        best_size = sz
                        best_params = f"pymupdf q={q} dpi={dpi}"
                    if sz <= per_file_budget:
                        break
                except Exception as e:
                    print(f"COMPRESS_FAIL {p.name} q={q}: {e}")

        if best_size is None:
            # fallback copy
            shutil.copy2(p, out_path)
            best_size = out_path.stat().st_size
            best_params = "copy"

        results.append(
            {
                "name": p.name,
                "original_bytes": orig_size,
                "compressed_bytes": best_size,
                "method": best_params,
                "path": str(out_path),
            }
        )

    total_c = sum(r["compressed_bytes"] for r in results)
    print(f"COMPRESSED_TOTAL_MB={total_c/1024/1024:.3f}")

    # If still over target, re-compress largest files more aggressively
    round_n = 0
    while total_c > TARGET and round_n < 6:
        round_n += 1
        results.sort(key=lambda r: r["compressed_bytes"], reverse=True)
        # compress top half more
        for r in results[: max(1, len(results) // 2 + 1)]:
            src = ORIGINALS / r["name"]
            out_path = COMPRESSED / r["name"]
            # more aggressive based on round
            q = max(25, 40 - round_n * 5)
            dpi = max(55, 75 - round_n * 5)
            try:
                sz = compress_with_pymupdf(src, out_path, q, dpi)
                print(f"RECOMPRESS {r['name']} q={q} dpi={dpi} -> {sz/1024/1024:.2f} MB")
                r["compressed_bytes"] = sz
                r["method"] = f"pymupdf-re q={q} dpi={dpi}"
            except Exception as e:
                print(f"RECOMPRESS_FAIL {r['name']}: {e}")
        total_c = sum(r["compressed_bytes"] for r in results)
        print(f"AFTER_ROUND_{round_n}_TOTAL_MB={total_c/1024/1024:.3f}")

    # Also replace root copies with compressed (keep originals in pdf_originals)
    for r in results:
        src = COMPRESSED / r["name"]
        dst = ROOT / r["name"]
        shutil.copy2(src, dst)
        print(f"REPLACED_ROOT {r['name']}")

    report = {
        "target_bytes": TARGET,
        "original_total": total0,
        "compressed_total": total_c,
        "ok": total_c <= TARGET,
        "files": sorted(results, key=lambda x: x["name"]),
    }
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK={report['ok']}")
    print(f"REPORT={REPORT}")


if __name__ == "__main__":
    main()
