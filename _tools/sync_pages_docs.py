# -*- coding: utf-8 -*-
"""Mirror downloadable docs into docs/ and sync Pages HTML from инструкция/."""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path
from urllib.parse import quote

ROOT = Path(r"D:\projects\Фурик")
INSTR = ROOT / "инструкция"
DOCS = ROOT / "docs"
WORKER = ROOT / "документы-работника"

sys.stdout.reconfigure(encoding="utf-8")

LABELS = {
    "O'ZBEKISTON RESPUBLIKASI.pdf": "Паспорт (Узбекистан)",
    "АО КБ Солидарность.pdf": "ДМС / АО КБ «Солидарность»",
    "Б (ВыездDeparture).pdf": "Миграционная карта",
    "Государственное_бюджетное_учреждение_здравоохранения_города.pdf": "Медосмотр",
    "СТРАХОВАЯ КОМПАНИЯ.pdf": "Страховая компания",
}


def write_worker_index(folder: Path, back_href: str) -> None:
    items = []
    for p in sorted(folder.iterdir()):
        if p.suffix.lower() != ".pdf":
            continue
        label = LABELS.get(p.name)
        if not label:
            if "МИНИСТЕРСТВО" in p.name or "ВНУТРЕННИХ" in p.name:
                label = "Патент (МВД)"
            else:
                label = p.name
        items.append((label, "./" + quote(p.name)))

    lis = "\n".join(
        f'    <li><span class="name">{label}</span>\n'
        f'      <a class="dl" href="{href}" download>Скачать PDF</a></li>'
        for label, href in items
    )
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
    li {{ margin: 0.6rem 0; padding: 0.75rem 1rem; border: 1px solid #e5e5e5; border-radius: 4px; display: flex; flex-wrap: wrap; gap: 0.5rem 1rem; align-items: center; justify-content: space-between; }}
    a.dl {{ display: inline-block; padding: 0.35rem 0.75rem; background: #0b5fff; color: #fff; text-decoration: none; border-radius: 2px; font-size: 0.9rem; }}
    a.dl:hover {{ background: #0842c0; }}
    .name {{ font-size: 0.95rem; }}
    .muted {{ color: #666; font-size: 0.85rem; }}
  </style>
</head>
<body>
  <p class="muted"><a href="{back_href}">← к инструкции</a></p>
  <h1>Сканы документов работника</h1>
  <p class="muted">Сжатые PDF из пакета. Откройте или сохраните файл.</p>
  <ul>
{lis}
  </ul>
</body>
</html>
"""
    (folder / "index.html").write_text(html, encoding="utf-8")
    print(f"worker index: {len(items)} PDFs in {folder}")


def patch_instruction_html(text: str) -> str:
    """Ensure download CTA wording; remove leftover gosuslugi disclaimer if any."""
    text = re.sub(
        r'\s*<p class="muted">\s*Памятка оформлена в стиле интерфейса Госуслуг.*?</p>\s*',
        "\n",
        text,
        flags=re.S,
    )
    return text


def inject_downloads_block(text: str, prefix: str) -> str:
    """Replace status 'Уже есть' list with download buttons. prefix is '../' or ''."""
    block = f"""          <ul class="status-list">
            <li>Сканы работника —
              <a class="btn btn--ghost" href="{prefix}документы-работника/index.html">Скачать PDF</a></li>
            <li>СНИЛС и ИНН —
              <a class="btn btn--ghost" href="{prefix}СНИЛС%20ИНН.txt" download>Скачать TXT</a></li>
            <li>Бессрочный ТД (бармен, Киевская&nbsp;7к2, Сбер, 01.08.2026) —
              <a class="btn btn--ghost" href="{prefix}трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx" download>Скачать Word</a></li>
            <li>Черновик уведомления МВД о заключении ТД —
              <a class="btn btn--ghost" href="{prefix}уведомление-мвд/уведомление-МВД-заключение-ТД-Ходжиматов-01.08.2026.docx" download>Скачать Word</a>
              · <a href="{prefix}уведомление-мвд/README.md">README</a></li>
            <li>Шаблон заявления об увольнении —
              <a class="btn btn--ghost" href="{prefix}кадры-увольнение/заявление-об-увольнении.docx" download>Скачать Word</a></li>
            <li>Реквизиты ИП, зарплатный счёт Сбера, адрес точки — в ТД</li>
          </ul>"""

    text = re.sub(
        r'<ul class="status-list">\s*<li>Сканы работника[\s\S]*?<li>Реквизиты ИП[\s\S]*?</ul>',
        block,
        text,
        count=1,
    )

    # Contract download button
    text = text.replace(
        f'<a href="{prefix}трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx"><code>трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx</code></a>',
        f'<a class="btn btn--ghost" href="{prefix}трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx" download>Скачать Word</a>',
    )
    # Also handle without prefix already being in href as ../
    text = text.replace(
        '<a href="../трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx"><code>трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx</code></a>',
        f'<a class="btn btn--ghost" href="{prefix}трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-01.08.2026.docx" download>Скачать Word</a>',
    )

    # MVD draft
    text = re.sub(
        r'<a href="(?:\.\./)?уведомление-мвд/уведомление-МВД-заключение-ТД-Ходжиматов-01\.08\.2026\.docx"><code>уведомление-мвд/…01\.08\.2026\.docx</code></a>\s*·\s*инструкция подачи —\s*<a href="(?:\.\./)?уведомление-мвд/README\.md"><code>README\.md</code></a>',
        f'<a class="btn btn--ghost" href="{prefix}уведомление-мвд/уведомление-МВД-заключение-ТД-Ходжиматов-01.08.2026.docx" download>Скачать Word</a>'
        f' · <a href="{prefix}уведомление-мвд/README.md">README</a>',
        text,
        count=1,
    )

    # Prepare checklist link to worker docs
    text = re.sub(
        r'сканы в <a href="(?:\.\./)?документы-работника/"><code>документы-работника/</code></a>',
        f'сканы — <a class="btn btn--ghost" href="{prefix}документы-работника/index.html">Скачать PDF</a>',
        text,
        count=1,
    )
    text = re.sub(
        r'СНИЛС — <a href="(?:\.\./)?СНИЛС%20ИНН\.txt"><code>СНИЛС ИНН\.txt</code></a>',
        f'СНИЛС — <a class="btn btn--ghost" href="{prefix}СНИЛС%20ИНН.txt" download>Скачать TXT</a>',
        text,
        count=1,
    )

    # Fire section: Word only (no .md)
    text = re.sub(
        r'<a[^>]*href="(?:\.\./)?кадры-увольнение/заявление-об-увольнении\.docx"[^>]*>.*?</a>(?:\s*·\s*<a[^>]*заявление-об-увольнении\.md[^>]*>.*?</a>)?',
        f'<a class="btn btn--ghost" href="{prefix}кадры-увольнение/заявление-об-увольнении.docx" download>Скачать Word</a>',
        text,
    )

    return text


def to_pages_html(instr_html: str) -> str:
    """Convert local ../ paths to docs-root relative paths."""
    html = instr_html.replace('href="../', 'href="')
    # Pages note under hero alert
    note = (
        '        Файлы ниже скачиваются прямо с этой страницы (GitHub Pages). '
        'Исходники также в <a href="https://github.com/puholet-sketch/furik" target="_blank" rel="noopener">репозитории</a>.\n'
    )
    if "Файлы ниже скачиваются" not in html and "Файлы проекта открываются" not in html:
        html = html.replace(
            "        Работодатель: <strong>ИП Сорванова А.А.</strong>",
            note + "        Работодатель: <strong>ИП Сорванова А.А.</strong>",
            1,
        )
    else:
        html = re.sub(
            r"Файлы проекта открываются из <a[^>]*>репозитория на GitHub</a>\.",
            "Файлы ниже скачиваются прямо с этой страницы (GitHub Pages).",
            html,
        )
    return html


def mirror() -> None:
    write_worker_index(WORKER, "../index.html")

    mirrors = [
        "трудовой-договор",
        "уведомление-мвд",
        "кадры-увольнение",
        "документы-работника",
    ]
    for name in mirrors:
        src = ROOT / name
        dst = DOCS / name
        if dst.exists():
            shutil.rmtree(dst)
        shutil.copytree(
            src,
            dst,
            ignore=shutil.ignore_patterns("*.md") if name == "кадры-увольнение" else None,
        )
        print("mirrored", name)

    # Worker index back-link for Pages is ../index.html (docs root) — same
    write_worker_index(DOCS / "документы-работника", "../index.html")

    shutil.copy2(ROOT / "СНИЛС ИНН.txt", DOCS / "СНИЛС ИНН.txt")
    shutil.copy2(INSTR / "styles.css", DOCS / "styles.css")

    # Patch local instruction
    local = (INSTR / "index.html").read_text(encoding="utf-8")
    local = patch_instruction_html(local)
    local = inject_downloads_block(local, "../")
    (INSTR / "index.html").write_text(local, encoding="utf-8")
    print("updated инструкция/index.html")

    pages = to_pages_html(local)
    # Re-inject with empty prefix in case any ../ left inconsistently — already replaced
    # Ensure docs version uses empty prefix buttons (already from replace ../)
    (DOCS / "index.html").write_text(pages, encoding="utf-8")
    print("updated docs/index.html")


if __name__ == "__main__":
    mirror()
