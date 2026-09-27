# -*- coding: utf-8 -*-
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
frag = (DOCS / "_reg_fragment.html").read_text(encoding="utf-8")

html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Расчёты по Фурику — ЗП, налоги, покупки</title>
  <link rel="stylesheet" href="styles.css" />
  <style>
    .table-registry {{ min-width: 720px; }}
    .table-registry col.c-date {{ width: 12%; }}
    .table-registry col.c-order {{ width: 16%; }}
    .table-registry col.c-item {{ width: 42%; }}
    .table-registry col.c-sum {{ width: 14%; }}
    .table-registry col.c-store {{ width: 16%; }}
    .table-registry code {{
      font-size: 0.82em;
      background: var(--blue-soft);
      padding: 0.1em 0.35em;
      border-radius: 4px;
    }}
    .paid-note {{ font-size: 0.92rem; }}
    .split-actions {{ display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }}
  </style>
</head>
<body>
  <header class="topbar" id="topbar">
    <div class="topbar__inner">
      <div class="brand">
        <span class="brand__mark" aria-hidden="true"></span>
        <div>
          <div class="brand__name">Расчёты Фурик</div>
          <div class="brand__sub">0,125 ставки · налоги · покупки 2026</div>
        </div>
      </div>
      <button class="nav-toggle" type="button" aria-label="Меню" aria-expanded="false" id="navToggle">
        <span></span>
      </button>
      <nav class="topnav" aria-label="Разделы" id="topnav">
        <a href="index.html">← Инструкция</a>
        <a href="#zp">ЗП 0,125</a>
        <a href="#paid">Уплачено</a>
        <a href="#contract">Договор</a>
        <a href="#registry">Покупки</a>
      </nav>
    </div>
  </header>

  <main>
    <section class="hero block is-visible">
      <p class="eyebrow">ИП Сорванова · Ходжиматов · с 25.08.2026</p>
      <h1>Зарплата, налоги и покупки Фурика</h1>
      <p class="lead">
        Официальная часть при <strong>0,125 ставки</strong>, что уже ушло в бюджет по выпискам,
        договор для скачивания и реестр личных покупок Купера к перевыставлению.
      </p>
      <div class="hero__actions">
        <a class="btn btn--primary" href="#zp">К зарплате</a>
        <a class="btn btn--ghost" href="#registry">К реестру</a>
        <a class="btn btn--pdf" href="трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-06.08.2026.pdf" download>ТД PDF</a>
      </div>
    </section>

    <div class="facts">
      <div class="fact">
        <span class="fact__label">Оклад 0,125</span>
        <span class="fact__value">4&nbsp;966,25&nbsp;₽</span>
        <span class="fact__hint">1/8 от 39&nbsp;730</span>
      </div>
      <div class="fact">
        <span class="fact__label">На руки</span>
        <span class="fact__value">4&nbsp;320,25&nbsp;₽</span>
        <span class="fact__hint">после НДФЛ 13%</span>
      </div>
      <div class="fact">
        <span class="fact__label">Налоги/взносы</span>
        <span class="fact__value">2&nbsp;145,81&nbsp;₽</span>
        <span class="fact__hint">НДФЛ + 30% + 0,2%</span>
      </div>
      <div class="fact">
        <span class="fact__label">Покупки 2026</span>
        <span class="fact__value">9&nbsp;853,56&nbsp;₽</span>
        <span class="fact__hint">40 позиций</span>
      </div>
    </div>

    <section id="zp" class="block">
      <div class="step-head">
        <span class="step-num">1</span>
        <div>
          <h2>ЗП при 0,125 ставки</h2>
          <p class="muted">0,125 = 1/8 полной ставки. Действует с 25.08.2026 (допсоглашение).</p>
        </div>
      </div>

      <div class="table-wrap">
        <table class="table-stack">
          <thead>
            <tr><th>Параметр</th><th class="num">Полная ставка</th><th class="num">0,125 ставки</th></tr>
          </thead>
          <tbody>
            <tr>
              <td data-label="Параметр">Оклад</td>
              <td data-label="Полная" class="num">39&nbsp;730&nbsp;₽</td>
              <td data-label="0,125" class="num"><strong>4&nbsp;966,25&nbsp;₽</strong></td>
            </tr>
            <tr>
              <td data-label="Параметр">Часы</td>
              <td data-label="Полная" class="num">40 ч/нед.</td>
              <td data-label="0,125" class="num">5 ч/нед. (1 ч/день)</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="table-wrap" style="margin-top:16px">
        <table class="table-stack">
          <thead>
            <tr><th>Статья</th><th class="num">Сумма</th><th>Кто / куда</th></tr>
          </thead>
          <tbody>
            <tr>
              <td data-label="Статья">На руки</td>
              <td data-label="Сумма" class="num"><strong>4&nbsp;320,25&nbsp;₽</strong></td>
              <td data-label="Куда">сотруднику на Сбер</td>
            </tr>
            <tr>
              <td data-label="Статья">НДФЛ 13%</td>
              <td data-label="Сумма" class="num">646&nbsp;₽</td>
              <td data-label="Куда">из зарплаты → ФНС (ЕНП)</td>
            </tr>
            <tr>
              <td data-label="Статья">Взносы 30%</td>
              <td data-label="Сумма" class="num">1&nbsp;489,88&nbsp;₽</td>
              <td data-label="Куда">сверху ИП → ФНС (ЕНП)</td>
            </tr>
            <tr>
              <td data-label="Статья">Травматизм 0,2%</td>
              <td data-label="Сумма" class="num">9,93&nbsp;₽</td>
              <td data-label="Куда">сверху ИП → СФР</td>
            </tr>
            <tr class="row-total">
              <td data-label="Статья">Итого НДФЛ + взносы + травматизм</td>
              <td data-label="Сумма" class="num"><strong>2&nbsp;145,81&nbsp;₽</strong></td>
              <td data-label="Куда">в месяц при 0,125 ставки</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="alert alert--info" style="margin-top:14px">
        Нагрузка ИП с кармана: оклад 4&nbsp;966,25 + взносы 1&nbsp;489,88 + травматизм 9,93 =
        <strong>6&nbsp;466,06&nbsp;₽/мес</strong> (НДФЛ внутри оклада).
        Остальное сверх официальной части — в рамках weekly budget проекта, вне этой таблицы.
      </div>
    </section>

    <section id="paid" class="block">
      <div class="step-head">
        <span class="step-num">2</span>
        <div>
          <h2>Что уже уплачено государству</h2>
          <p class="muted">По сверке с банковскими платежами авг–сен 2026 (только подтверждённые строки).</p>
        </div>
      </div>

      <div class="table-wrap">
        <table class="table-stack">
          <thead>
            <tr><th>Дата</th><th class="num">Сумма</th><th>Назначение / куда</th><th>Период</th></tr>
          </thead>
          <tbody>
            <tr>
              <td data-label="Дата">23.08.2026</td>
              <td data-label="Сумма" class="num">2&nbsp;214&nbsp;₽</td>
              <td data-label="Куда">НДФЛ → ФНС (ЕНП)</td>
              <td data-label="Период">август 2026 (часть)</td>
            </tr>
            <tr>
              <td data-label="Дата">23.09.2026</td>
              <td data-label="Сумма" class="num">1&nbsp;137&nbsp;₽</td>
              <td data-label="Куда">НДФЛ → ФНС (ЕНП)</td>
              <td data-label="Период">август 2026 (доплата)</td>
            </tr>
            <tr>
              <td data-label="Дата">23.09.2026</td>
              <td data-label="Сумма" class="num">323&nbsp;₽</td>
              <td data-label="Куда">НДФЛ → ФНС (ЕНП)</td>
              <td data-label="Период">сентябрь 2026 (аванс)</td>
            </tr>
            <tr>
              <td data-label="Дата">23.09.2026</td>
              <td data-label="Сумма" class="num">7&nbsp;733,16&nbsp;₽</td>
              <td data-label="Куда">Взносы 30% → ФНС (ЕНП)</td>
              <td data-label="Период">август 2026</td>
            </tr>
          </tbody>
        </table>
      </div>

      <p class="paid-note muted" style="margin-top:12px">
        Август: НДФЛ итого <strong>3&nbsp;351&nbsp;₽</strong> (2&nbsp;214 + 1&nbsp;137) и взносы <strong>7&nbsp;733,16&nbsp;₽</strong> —
        сошлись с начислением 25&nbsp;777,20&nbsp;₽.
        Сентябрь: с аванса удержан НДФЛ <strong>323&nbsp;₽</strong>; доплата НДФЛ ≈323&nbsp;₽ и взносы 1&nbsp;489,88&nbsp;₽ —
        по срокам до 28.10.2026.
        Платежи травматизма 0,2% в СФР (авг ≈51,55&nbsp;₽, сен 9,93&nbsp;₽) в этой сверке <strong>не подтверждены</strong> — проверить отдельно.
      </p>
    </section>

    <section id="contract" class="block">
      <div class="step-head">
        <span class="step-num">3</span>
        <div>
          <h2>Договор</h2>
          <p class="muted">Канонический ТД от 06.08.2026 и допсоглашение на 0,125 ставки.</p>
        </div>
      </div>
      <div class="split-actions">
        <a class="btn btn--pdf" href="трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-06.08.2026.pdf" download>Скачать ТД (PDF)</a>
        <a class="btn btn--ghost" href="трудовой-договор/трудовой-договор-Ходжиматов-бессрочный-06.08.2026.docx" download>ТД Word</a>
        <a class="btn btn--ghost" href="трудовой-договор/допсоглашение-0.125-ставки-Ходжиматов.pdf" download>Допсоглашение 0,125 (PDF)</a>
      </div>
    </section>

    <section id="registry" class="block">
      <div class="step-head">
        <span class="step-num">4</span>
        <div>
          <h2>Реестр покупок Фурика за 2026</h2>
          <p class="muted">Личные позиции из заказов Купера к перевыставлению. Итого <strong>9&nbsp;853,56&nbsp;₽</strong> · 40 позиций.</p>
        </div>
      </div>
      <div class="split-actions" style="margin-bottom:12px">
        <a class="btn btn--primary" href="заказы-купер/реестр-перевыставление-Фурик.xlsx" download>Скачать Excel</a>
      </div>
      <div class="table-wrap">
        <table class="table-stack table-registry">
          <colgroup>
            <col class="c-date" /><col class="c-order" /><col class="c-item" /><col class="c-sum" /><col class="c-store" />
          </colgroup>
          <thead>
            <tr>
              <th>Дата</th>
              <th>№ заказа</th>
              <th>Товар</th>
              <th class="num">Сумма</th>
              <th>Магазин</th>
            </tr>
          </thead>
          <tbody>
{frag}
            <tr class="row-total">
              <td data-label="Итого" colspan="3"><strong>Итого к перевыставлению</strong></td>
              <td data-label="Сумма" class="num"><strong>9&nbsp;853,56&nbsp;₽</strong></td>
              <td data-label="Позиции">40 позиций</td>
            </tr>
          </tbody>
        </table>
      </div>
    </section>

    <p class="muted" style="padding: 8px 16px 40px; max-width: var(--max); margin: 0 auto;">
      Страница рядом с <a href="index.html">инструкцией по найму</a>. Цифры ЗП/налогов — ориентир; перед платежами сверьте с бухгалтером.
    </p>
  </main>

  <script>
    (function () {{
      const toggle = document.getElementById("navToggle");
      const nav = document.getElementById("topnav");
      if (!toggle || !nav) return;
      toggle.addEventListener("click", () => {{
        const open = nav.classList.toggle("is-open");
        toggle.setAttribute("aria-expanded", open ? "true" : "false");
      }});
      nav.querySelectorAll("a").forEach((a) => {{
        a.addEventListener("click", () => {{
          nav.classList.remove("is-open");
          toggle.setAttribute("aria-expanded", "false");
        }});
      }});
    }})();
  </script>
</body>
</html>
"""

(DOCS / "raschety.html").write_text(html, encoding="utf-8")
(DOCS / "_reg_fragment.html").unlink(missing_ok=True)
(DOCS / "_reg_meta.txt").unlink(missing_ok=True)
print("ok", len(html), "bytes")
