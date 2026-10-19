#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
make_delta_columns.py — дельты накопительных счётчиков для выгрузки в BI.

Зачем
-----
Форт Монитор отдаёт ТОЛЬКО накопительные счётчики. Чтобы в DataLens показать
«площадь в га», «пробег в м» и «пробег с забитыми семяпроводами в м», нужны
дельты. Считать их «на лету» в BI нельзя: счётчики сбрасываются, и наивная
разность даёт огромный отрицательный скачок, который ломает SUM().

Что делает скрипт
-----------------
1. Идёт по строкам в порядке времени и ведёт «высокую отметку» (high-water mark)
   отдельно по каждому счётчику.
2. Откат вниз, который меньше половины текущей отметки, считается СБРОСОМ:
   начинается новая эпоха, дельта в этой строке = 0.
3. Мелкий откат (джиттер) просто игнорируется — дельта 0, отметка не движется.
4. Масштабы применяются один раз, здесь.

Проверено на объекте 12917 (СКИФ-48 Эл.Доз), файл 09–16.10.2026:
    Σ ΔПлощадь = 408.20 га,  Σ ΔПробег = 680 333.33 м
   инвариант ΔS × 10000 / ΔL = 6.0000 м  == ширине захвата 101_8 (60 → 6.0 м)
Если этот инвариант сходится — дельты посчитаны верно.

Запуск
------
    pip install pandas openpyxl
    python make_delta_columns.py "исходный.xlsx" "результат_с_дельтами.xlsx"
"""

from __future__ import annotations

import sys

import numpy as np
import pandas as pd

SHEET = "Журнал принятых данных"

# колонка-источник → (имя новой колонки, масштаб, подпись для отчёта)
COUNTERS: dict[str, tuple[str, float]] = {
    # 151_22: масштаб 0.1 подтверждён инвариантом на 12917 (ref.param_def area_cum)
    "151_22_ПлощадьСеваНакопительно": ("ΔПлощадь_га", 0.1),
    # 151_23: приходит в метрах, масштаб 1 (ref.param_def distance_cum)
    "151_23_ПробегСеваНакопительно": ("ΔПробег_м", 1.0),
    # 153_28: приходит в метрах, масштаб 1 (ref.param_def blockage_distance).
    # ВНИМАНИЕ: по документации проекта это СУММА метро-сошников, поэтому
    # рядом добавляется ΔЗабитие_на_рядок = ΔЗабитие_м / 101_7_Кол-во_сошников.
    "153_28_Забитие_семяпроводов, м": ("ΔЗабитие_м", 1.0),
}

N_ROWS_COL = "101_7_Кол-во_сошников"

# Сброс — это откат глубже чем в 2 раза. Мелкий джиттер (шум передачи,
# пересчёт округления) сюда не попадает: на 12917 откаты площади были 1…6
# условных единиц (0.1…0.6 га) при настоящем сбросе 1786.
RESET_RATIO = 0.5


def counter_delta(values: pd.Series, reset_ratio: float = RESET_RATIO):
    """Дельта накопительного счётчика с защитой от сброса через ноль.

    Возвращает (delta, epoch, reset_flag). Первая строка файла даёт дельту 0:
    предыдущего значения мы не знаем, а значение счётчика на момент старта
    выгрузки к работе внутри этого файла отношения не имеет.
    """
    v = pd.to_numeric(values, errors="coerce").astype("float64")
    n = len(v)
    delta = np.zeros(n)
    epoch = np.zeros(n, dtype="int32")
    reset = np.zeros(n, dtype="int8")

    run: float | None = None  # текущая высокая отметка внутри эпохи
    e = 1

    for i in range(n):
        x = v[i]
        if np.isnan(x):  # пропуск в выдаче — дельты нет, эпоха сохраняется
            epoch[i] = e
            continue
        if run is None:
            run = x  # первая строка: инициализация, дельта 0
        elif x < reset_ratio * run:
            e += 1  # счётчик уехал в ноль — новая эпоха
            run = x
            reset[i] = 1
        elif x > run:
            delta[i] = x - run  # прирост; откаты вниз игнорируются
            run = x
        epoch[i] = e

    return pd.Series(delta), pd.Series(epoch), pd.Series(reset)


def add_deltas(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    n0 = len(out.columns)

    deltas, epochs, resets, labels = {}, {}, {}, {}
    for src in COUNTERS:
        d, ep, rs = counter_delta(out[src])
        deltas[src], epochs[src], resets[src] = d, ep, rs

    short = {
        "151_22_ПлощадьСеваНакопительно": "Площадь",
        "151_23_ПробегСеваНакопительно": "Пробег",
        "153_28_Забитие_семяпроводов, м": "Забитие",
    }
    lab = np.full(len(out), "", dtype=object)
    for src, name in short.items():
        lab = np.where(resets[src].eq(1), name, lab)
    # площадь и забитие сбрасываются в разные моменты — склеиваем в одну метку
    lab = np.where(resets["151_22_ПлощадьСеваНакопительно"].eq(1)
                   & resets["153_28_Забитие_семяпроводов, м"].eq(1),
                   "Площадь+Забитие", lab)

    pos = n0
    for src, (col, scale) in COUNTERS.items():
        # площадь: квант 0.1 га — 3 знака точны; пробег: квант 1/3 м — 3 знака
        # дают дрейф 1.4 м на 680 км (0.0002 %); забитие: целые метры
        out.insert(pos, col, (deltas[src] * scale).round(3))
        pos += 1

    n_rows = pd.to_numeric(out[N_ROWS_COL], errors="coerce")
    out.insert(pos, "ΔЗабитие_на_рядок", (deltas["153_28_Забитие_семяпроводов, м"] / n_rows).round(3))
    out.insert(pos + 1, "сброс_счётчика", pd.Series(lab))
    out.insert(pos + 2, "эпоха_счётчиков",
               pd.Series(np.maximum.reduce([epochs[s] for s in COUNTERS]).astype("int32")))

    return out


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print(__doc__)
        print("Ошибка: нужно указать исходный и итоговый файл")
        return 2

    src, dst = argv[1], argv[2]
    df = pd.read_excel(src, sheet_name=SHEET, dtype=object)
    missing = [c for c in COUNTERS if c not in df.columns]
    if missing:
        print("Нет колонок:", missing)
        return 1

    out = add_deltas(df)
    out.to_excel(dst, sheet_name=SHEET, index=False)

    # --- самопроверка инвариантом: ΔS × 10000 / ΔL == ширина захвата ---
    s_area = out["ΔПлощадь_га"].sum()
    s_dist = out["ΔПробег_м"].sum()
    s_blk = out["ΔЗабитие_м"].sum()
    width = s_area * 10000 / s_dist if s_dist else float("nan")

    print(f"строк: {len(out)}, колонок: {len(out.columns)} (было {len(df)})")
    print(f"Σ ΔПлощадь_га = {s_area:,.2f}")
    print(f"Σ ΔПробег_м   = {s_dist:,.2f}")
    print(f"Σ ΔЗабитие_м  = {s_blk:,.0f}  ({s_blk / s_dist * 100:.1f}% пробега)" if s_dist else "")
    print(f"контроль ширины = {width:.4f} м")
    marks = out["сброс_счётчика"].fillna("")
    # метка может быть составной («Площадь+Забитие»), поэтому считаем вхождения
    print("сбросов: " + ", ".join(
        f"{name}={int(marks.str.contains(name, regex=False).sum())}"
        for name in ("Площадь", "Пробег", "Забитие")))
    print(f"отрицательных дельт: "
          f"{int((out['ΔПлощадь_га'] < 0).sum())} / "
          f"{int((out['ΔПробег_м'] < 0).sum())} / "
          f"{int((out['ΔЗабитие_м'] < 0).sum())}")
    print("записано:", dst)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))