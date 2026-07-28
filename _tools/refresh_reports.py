# -*- coding: utf-8 -*-
from pathlib import Path
import json
import hashlib
from datetime import datetime, timezone

ROOT = Path(r"D:\projects\Фурик")
orig_dir = ROOT / "pdf_originals"
comp_dir = ROOT / "compressed"

files = []
for p in sorted(comp_dir.glob("*.pdf")):
    o = orig_dir / p.name
    files.append({
        "name": p.name,
        "original_bytes": o.stat().st_size if o.exists() else None,
        "compressed_bytes": p.stat().st_size,
        "path": str(p),
    })

report = {
    "target_bytes": int(19.9 * 1024 * 1024),
    "original_total": sum(f["original_bytes"] or 0 for f in files),
    "compressed_total": sum(f["compressed_bytes"] for f in files),
    "ok": True,
    "files": files,
}
report["ok"] = report["compressed_total"] <= report["target_bytes"]
(ROOT / "PDF_COMPRESSION_REPORT.json").write_text(
    json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
)

lines = [
    "# Отчёт о сжатии PDF",
    "",
    f"Цель: не более 19.9 МиБ ({report['target_bytes']} байт).",
    "",
    "| Файл | Было | Стало |",
    "|---|---:|---:|",
]
for f in files:
    lines.append(
        f"| {f['name']} | {f['original_bytes']/1024/1024:.2f} МБ | {f['compressed_bytes']/1024/1024:.2f} МБ |"
    )
lines += [
    "",
    f"| **Итого** | **{report['original_total']/1024/1024:.2f} МБ** | "
    f"**{report['compressed_total']/1024/1024:.2f} МБ** |",
    "",
    f"Критерий ≤ 19.9 МБ: **{'выполнен' if report['ok'] else 'НЕ выполнен'}**.",
    "",
    "Оригиналы: `pdf_originals/`. Сжатые: `compressed/` и корень проекта.",
]
(ROOT / "PDF_COMPRESSION_REPORT.md").write_text("\n".join(lines), encoding="utf-8")

# refresh state hashes
ctx = (ROOT / ".ai" / "CONTEXT.md").read_text(encoding="utf-8")
ctx_hash = hashlib.sha256(ctx.encode("utf-8")).hexdigest()
active = [
    "Трудовой_договор_иностранец_Узбекистан.docx",
    "ИНСТРУКЦИЯ_прием_иностранца_Узбекистан.md",
    "PDF_COMPRESSION_REPORT.md",
    "Трудовой_договор_черновик.md",
]
hashes = {n: hashlib.sha256((ROOT / n).read_bytes()).hexdigest() for n in active if (ROOT / n).exists()}
state = {
    "context_version": 1,
    "timestamp": datetime.now(timezone.utc).isoformat(),
    "context_hash": ctx_hash,
    "active_files": active,
    "active_file_hashes": hashes,
    "decisions": [
        f"PDF: {report['original_total']/1024/1024:.2f} → {report['compressed_total']/1024/1024:.2f} МБ",
        "Мин. оплата Москва 2026 = 39730; бессрочный ТД рекомендуется",
        "PII не в .ai; PDF — сканы без OCR",
    ],
    "blockers": ["Реквизиты ИП и ручное заполнение данных работника"],
    "last_error": None,
    "retry_count": 0,
    "next_action": "Заполнить плейсхолдеры ТД; после подписания — уведомление МВД ≤ 3 раб. дн.",
}
(ROOT / ".ai" / "state.json").write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

# update INDEX with draft
idx = (ROOT / ".ai" / "INDEX.md").read_text(encoding="utf-8")
if "Трудовой_договор_черновик.md" not in idx:
    idx = idx.replace(
        "| `_tools/` | Скрипты сжатия/генерации |",
        "| `Трудовой_договор_черновик.md` | Краткое оглавление ТД |\n| `_tools/` | Скрипты сжатия/генерации |",
    )
    (ROOT / ".ai" / "INDEX.md").write_text(idx, encoding="utf-8")

print(json.dumps({
    "ok": report["ok"],
    "was_mb": round(report["original_total"]/1024/1024, 2),
    "now_mb": round(report["compressed_total"]/1024/1024, 2),
}, ensure_ascii=False))
