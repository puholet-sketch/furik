# -*- coding: utf-8 -*-
from pathlib import Path
import json
import hashlib
from datetime import datetime, timezone

ROOT = Path(r"D:\projects\Фурик")

# contract already may be generated; ensure
import runpy
runpy.run_path(str(ROOT / "_tools" / "make_contract_docx.py"), run_name="__main__")

report = json.loads((ROOT / "PDF_COMPRESSION_REPORT.json").read_text(encoding="utf-8"))
lines = [
    "# Отчёт о сжатии PDF",
    "",
    f"Цель: не более 19.9 МиБ ({report['target_bytes']} байт).",
    "",
    "| Файл | Было | Стало | Метод |",
    "|---|---:|---:|---|",
]
for f in report["files"]:
    lines.append(
        f"| {f['name']} | {f['original_bytes']/1024/1024:.2f} МБ | "
        f"{f['compressed_bytes']/1024/1024:.2f} МБ | {f['method']} |"
    )
lines += [
    "",
    f"| **Итого** | **{report['original_total']/1024/1024:.2f} МБ** | "
    f"**{report['compressed_total']/1024/1024:.2f} МБ** | OK={report['ok']} |",
    "",
    "Оригиналы сохранены в `pdf_originals/`.",
    "Сжатые копии: `compressed/` и заменены файлы в корне проекта.",
]
(ROOT / "PDF_COMPRESSION_REPORT.md").write_text("\n".join(lines), encoding="utf-8")

total_root = sum(p.stat().st_size for p in ROOT.glob("*.pdf"))
print("ROOT_PDF_TOTAL_MB", round(total_root / 1024 / 1024, 3))
print("CONTRACT_EXISTS", (ROOT / "Трудовой_договор_иностранец_Узбекистан.docx").exists())

# state.json
ctx = (ROOT / ".ai" / "CONTEXT.md").read_text(encoding="utf-8")
ctx_hash = hashlib.sha256(ctx.encode("utf-8")).hexdigest()
active = [
    "Трудовой_договор_иностранец_Узбекистан.docx",
    "ИНСТРУКЦИЯ_прием_иностранца_Узбекистан.md",
    "PDF_COMPRESSION_REPORT.md",
]
hashes = {}
for name in active:
    p = ROOT / name
    hashes[name] = hashlib.sha256(p.read_bytes()).hexdigest() if p.exists() else None

state = {
    "context_version": 1,
    "timestamp": datetime.now(timezone.utc).isoformat(),
    "context_hash": ctx_hash,
    "active_files": active,
    "active_file_hashes": hashes,
    "decisions": [
        "PDF сжаты до ~13.4 МБ суммарно; оригиналы в pdf_originals/",
        "Рекомендован бессрочный ТД; минимум оплаты Москва 2026 = 39730",
        "PII из PDF не извлечены (сканы) и не пишутся в .ai/",
    ],
    "blockers": [
        "Нужны реквизиты ИП и ручное заполнение данных работника из PDF",
    ],
    "last_error": None,
    "retry_count": 0,
    "next_action": "Заполнить плейсхолдеры ТД и подать уведомление МВД после подписания",
}
(ROOT / ".ai" / "state.json").write_text(
    json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
)
print("STATE_OK", ctx_hash[:12])
