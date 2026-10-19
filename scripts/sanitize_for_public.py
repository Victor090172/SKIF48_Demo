#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
sanitize_for_public.py — обезличить набор перед публикацией.

Зачем
----
В исходной выгрузке — реальные данные клиента: координаты движущейся техники,
идентификаторы объектов и компаний, привязка к конкретным хозяйствам.
Публичный репозиторий на GitHub и share.streamlit.io доступны всем, поэтому
публиковать как есть нельзя.

Что делает
----------
1. Сдвигает все координаты на постоянное смещение (~2–3 км по широте и
   долготе). Геометрия трека и все производные показатели — площадь, пробег,
   забивание — сохраняются точно: это сдвиг, а не масштаб.
2. Округляет координаты до 4 знаков (~11 м) — убирает лишнюю точность.
3. Заменяет текстовые идентификаторы, если они есть.

Запуск:
    python scripts/sanitize_for_public.py data/telemetry_skif48.csv out.csv

ВАЖНО: скрипт НЕ удаляет даты. По календарю техники можно восстановить, где
и когда она работала. Если это недопустимо, замените колонку «Время» на
сдвиг по постоянному интервалу — но тогда разбивка по дням в дашборде станет
условной.
"""

from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd

# Сдвиг подбирается так, чтобы трек ушёл в сторону и не совпадал с исходным.
LAT_SHIFT = 0.020      # ~2,2 км
LON_SHIFT = 0.030      # ~2,3 км на этой широте
PRECISION = 4          # ~11 м


def sanitize(src: Path, dst: Path) -> pd.DataFrame:
    df = pd.read_csv(src)

    if "Широта" not in df.columns or "Долгота" not in df.columns:
        raise SystemExit("Нет колонок «Широта»/«Долгота» — нечего сдвигать")

    lat = pd.to_numeric(df["Широта"], errors="coerce")
    lon = pd.to_numeric(df["Долгота"], errors="coerce")

    df["Широта"] = (lat + LAT_SHIFT).round(PRECISION).map(repr)
    df["Долгота"] = (lon + LON_SHIFT).round(PRECISION).map(repr)

    # На всякий случай — любые колонки, похожие на идентификаторы.
    for c in df.columns:
        if c and ("ID" in c.upper() or "terminal" in c.lower()):
            df[c] = "removed"

    dst.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(dst, index=False, encoding="utf-8")

    print(f"Записано: {dst}")
    print(f"Сдвиг: {LAT_SHIFT:+.3f} по широте, {LON_SHIFT:+.3f} по долготе, "
          f"точность {PRECISION} знака")
    print("Показатели (площадь, пробег, забивание) не изменились: это сдвиг, не масштаб.")
    return df


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print(__doc__)
        return 2
    sanitize(Path(argv[1]), Path(argv[2]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))