# -*- coding: utf-8 -*-
"""Sync для-бухгалтерии → docs/ (GitHub Pages download mirror)."""
from __future__ import annotations

import hashlib
import json
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(r"D:\projects\Фурик")
DOCS = ROOT / "docs"
ACCT = ROOT / "для-бухгалтерии"

LABELS = {
    "O'ZBEKISTON RESPUBLIKASI.pdf": "Паспорт (Узбекистан)",
    "АО КБ Солидарность.pdf": "ДМС / АО КБ «Солидарность»",
    "Б (ВыездDeparture).pdf": "Миграционная карта",
    "Государственное_бюджетное_учреждение_здравоохранения_города.pdf": "Медосмотр",
    "СТРАХОВАЯ КОМПАНИЯ.pdf": "Страховая компания",
}

# Flat accounting package → Pages subfolders
MAP = {
    "трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx": "трудовой-договор",
    "трудовой-договор-Ходжиматов-бессрочный-01.08.2026.pdf": "трудовой-договор",
    "уведомление-МВД-заключение-ТД-Ходжиматов-01.08.2026.docx": "уведомление-мвд",
    "заявление-об-увольнении.docx": "кадры-увольнение",
    "СНИЛС ИНН.txt": ".",
}


def sha(p: Path) -> str:
    h = hashlib.sha256()
    h.update(p.read_bytes())
    return h.hexdigest()


def write_worker_index(folder: Path, back: str) -> None:
    items = []
    for p in sorted(folder.iterdir()):
        if p.suffix.lower() != ".pdf":
            continue
        label = LABELS.get(p.name)
        if not label:
            label = (
                "Патент (МВД)"
                if ("МИНИСТЕРСТВО" in p.name or "ВНУТРЕННИХ" in p.name)
                else p.name
            )
        items.append((label, "./" + quote(p.name)))
    lis_lines = []
    for lab, href in items:
        lis_lines.append(
            f'    <li><span class="name">{lab}</span>\n'
            f'      <a class="dl" href="{href}" download>Скачать PDF</a></li>'
        )
    lis = "\n".join(lis_lines)
    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Документы работника — скачать</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 720px; margin: 2rem auto; padding: 0 1rem; color: #333; }}
    h1 {{ font-size: 1.35rem; color: #1a1a1a; }}
    ul {{ list-style: none; padding: 0; }}
    li {{ margin: 0.6rem 0; padding: 0.75rem 1rem; border: 1px solid #e8e4e1; border-radius: 2px; display: flex; flex-wrap: wrap; gap: 0.5rem 1rem; align-items: center; justify-content: space-between; background: #f3f1ef; }}
    a.dl {{ display: inline-flex; padding: 0.45rem 0.85rem; background: #c00000; color: #fff; text-decoration: none; border-radius: 2px; font-size: 0.9rem; font-weight: 600; margin-left: auto; }}
    a.dl:hover {{ background: #8a0000; }}
    .name {{ font-size: 0.95rem; flex: 1 1 200px; }}
    .muted {{ color: #666; font-size: 0.85rem; }}
  </style>
</head>
<body>
  <p class="muted"><a href="{back}">← к инструкции</a></p>
  <h1>Сканы документов работника</h1>
  <p class="muted">Сжатые PDF. Каждый файл — отдельная кнопка справа.</p>
  <ul>
{lis}
  </ul>
</body>
</html>
"""
    (folder / "index.html").write_text(html, encoding="utf-8")
    print(f"worker index: {len(items)} in {folder}")


def main() -> None:
    if not ACCT.exists():
        raise SystemExit("missing для-бухгалтерии/")

    worker = DOCS / "документы-работника"
    worker.mkdir(parents=True, exist_ok=True)

    for src in ACCT.iterdir():
        if not src.is_file():
            continue
        if src.name in {"README.txt"}:
            continue
        if src.suffix.lower() == ".pdf" and src.name not in MAP:
            shutil.copy2(src, worker / src.name)
            print("worker pdf", src.name)
            continue
        if src.suffix.lower() == ".png":
            shutil.copy2(src, worker / src.name)
            print("worker png", src.name)
            continue
        sub = MAP.get(src.name)
        if sub is None:
            print("skip", src.name)
            continue
        dest_dir = DOCS if sub == "." else DOCS / sub
        dest_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest_dir / src.name)
        print("mirrored", src.name, "->", dest_dir.relative_to(ROOT))

    write_worker_index(worker, "../index.html")

    pages = (DOCS / "index.html").read_text(encoding="utf-8")
    assert "сверните" not in pages
    assert "сверьте" in pages
    assert "check-row__actions" in pages
    print("ok docs/index.html")

    # Sanity: every href download target under docs exists (simple relative paths)
    missing = []
    for token in [
        "трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.pdf",
        "трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx",
        "уведомление-мвд/уведомление-МВД-заключение-ТД-Ходжиматов-01.08.2026.docx",
        "кадры-увольнение/заявление-об-увольнении.docx",
        "СНИЛС ИНН.txt",
        "документы-работника/index.html",
    ]:
        if not (DOCS / token).exists():
            missing.append(token)
    if missing:
        raise SystemExit(f"Pages targets missing: {missing}")
    print("pages targets ok")
    print("done")


if __name__ == "__main__":
    main()
