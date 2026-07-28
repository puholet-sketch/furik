# -*- coding: utf-8 -*-
"""Sync docs canon -> mirrors, worker indexes, minimal .ai update."""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(r"D:\projects\Фурик")
DOCS = ROOT / "docs"
INSTR = ROOT / "инструкция"
WORKER = ROOT / "документы-работника"

LABELS = {
    "O'ZBEKISTON RESPUBLIKASI.pdf": "Паспорт (Узбекистан)",
    "АО КБ Солидарность.pdf": "ДМС / АО КБ «Солидарность»",
    "Б (ВыездDeparture).pdf": "Миграционная карта",
    "Государственное_бюджетное_учреждение_здравоохранения_города.pdf": "Медосмотр",
    "СТРАХОВАЯ КОМПАНИЯ.pdf": "Страховая компания",
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
    h1 {{ font-size: 1.35rem; color: #0b1f35; }}
    ul {{ list-style: none; padding: 0; }}
    li {{ margin: 0.6rem 0; padding: 0.75rem 1rem; border: 1px solid #d7e0ec; border-radius: 8px; display: flex; flex-wrap: wrap; gap: 0.5rem 1rem; align-items: center; justify-content: space-between; background: #fafbfe; }}
    a.dl {{ display: inline-flex; padding: 0.45rem 0.85rem; background: #005bff; color: #fff; text-decoration: none; border-radius: 6px; font-size: 0.9rem; font-weight: 600; margin-left: auto; }}
    a.dl:hover {{ background: #0047cc; }}
    .name {{ font-size: 0.95rem; flex: 1 1 200px; }}
    .muted {{ color: #66758b; font-size: 0.85rem; }}
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
    for name in [
        "трудовой-договор",
        "уведомление-мвд",
        "кадры-увольнение",
        "документы-работника",
    ]:
        src, dst = ROOT / name, DOCS / name
        if dst.exists():
            shutil.rmtree(dst)
        ignore = shutil.ignore_patterns("*.md") if name == "кадры-увольнение" else None
        shutil.copytree(src, dst, ignore=ignore)
        print("mirrored", name)

    shutil.copy2(ROOT / "СНИЛС ИНН.txt", DOCS / "СНИЛС ИНН.txt")

    write_worker_index(WORKER, "../инструкция/index.html")
    write_worker_index(DOCS / "документы-работника", "../index.html")

    pages = (DOCS / "index.html").read_text(encoding="utf-8")
    local = pages
    for folder in [
        "документы-работника/",
        "трудовой-договор/",
        "уведомление-мвд/",
        "кадры-увольнение/",
    ]:
        local = local.replace(f'href="{folder}', f'href="../{folder}')
    local = local.replace('href="СНИЛС%20ИНН.txt"', 'href="../СНИЛС%20ИНН.txt"')
    shutil.copy2(DOCS / "styles.css", INSTR / "styles.css")
    (INSTR / "index.html").write_text(local, encoding="utf-8")
    print("synced инструкция from docs")

    for p in [DOCS / "index.html", INSTR / "index.html"]:
        t = p.read_text(encoding="utf-8")
        assert "сверните" not in t, p
        assert "сверьте" in t, p
        assert "не официальный" not in t, p
        assert "check-row__actions" in t, p
        print("ok", p.relative_to(ROOT))

    ctx = ROOT / ".ai" / "CONTEXT.md"
    text = ctx.read_text(encoding="utf-8")
    text = re.sub(r"context_version:\s*\d+", "context_version: 9", text)
    text = re.sub(r"updated:\s*\d{4}-\d{2}-\d{2}", "updated: 2026-07-28", text)
    if "PDF ТД" not in text and "Word+PDF" not in text:
        text = text.replace(
            "**Следующий шаг:**",
            "**Доп.:** редизайн Pages; ТД Word+PDF; реквизиты 2 колонки; кнопки справа.\n\n**Следующий шаг:**",
        )
    # bump status line lightly
    text = text.replace(
        "HTML-инструкция + GitHub Pages (`/docs`) со скачиванием документов; НДФЛ/отпуск пояснены; заявление DOCX готово.",
        "Pages-редизайн (кнопки справа, Word+PDF ТД, per-file сканы); ТД 2 колонки; НДФЛ/отпуск пояснены.",
    )
    ctx.write_text(text, encoding="utf-8")

    active = [
        "трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx",
        "трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.pdf",
        "docs/index.html",
        "docs/styles.css",
        "инструкция/index.html",
        "инструкция/styles.css",
        "документы-работника/index.html",
        "_tools/fix_td_two_columns.py",
        "_tools/sync_pages_docs.py",
    ]
    hashes = {f: sha(ROOT / f) for f in active if (ROOT / f).exists()}
    state = {
        "context_version": 9,
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S+00:00"),
        "context_hash": sha(ctx),
        "active_files": list(hashes.keys()),
        "active_file_hashes": hashes,
        "decisions": [
            "TD Бармен 01.08.2026 indefinite",
            "TD section 10 two-column A4 margins 1.5cm",
            "Pages Word+PDF TD; download buttons right-aligned",
            "Worker docs per-file compressed PDF on Pages",
        ],
        "blockers": [],
        "last_error": None,
        "retry_count": 0,
        "next_action": "Sign TD; submit MVD notice within 3 business days",
        "publish": {
            "repo": "https://github.com/puholet-sketch/furik",
            "pages": "https://puholet-sketch.github.io/furik/",
            "td_docx": "https://puholet-sketch.github.io/furik/трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx",
            "td_pdf": "https://puholet-sketch.github.io/furik/трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.pdf",
        },
    }
    (ROOT / ".ai" / "state.json").write_text(
        json.dumps(state, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )

    idx = (ROOT / ".ai" / "INDEX.md").read_text(encoding="utf-8")
    if "01.08.2026.pdf" not in idx:
        idx = idx.replace(
            "| `трудовой-договор/…01.08.2026.docx` | Бессрочный ТД |",
            "| `трудовой-договор/…01.08.2026.docx` | Бессрочный ТД (Word) |\n"
            "| `трудовой-договор/…01.08.2026.pdf` | ТД PDF |",
        )
        (ROOT / ".ai" / "INDEX.md").write_text(idx, encoding="utf-8")

    # Update sync script note: docs is canon
    sync = ROOT / "_tools" / "sync_pages_docs.py"
    if sync.exists():
        s = sync.read_text(encoding="utf-8")
        if "docs is canon" not in s:
            s = s.replace(
                '"""Mirror downloadable docs into docs/ and sync Pages HTML from инструкция/."""',
                '"""Mirror downloadable docs into docs/. Canon HTML: docs/ → инструкция/ with ../ prefixes.\n\ndocs is canon for Pages."""',
            )
            sync.write_text(s, encoding="utf-8")

    pdfs = list((ROOT / "трудовой-договор").glob("*.pdf"))
    pdfs_docs = list((DOCS / "трудовой-договор").glob("*.pdf"))
    print("td pdfs root", [p.name for p in pdfs])
    print("td pdfs docs", [p.name for p in pdfs_docs])
    print("done")


if __name__ == "__main__":
    main()
