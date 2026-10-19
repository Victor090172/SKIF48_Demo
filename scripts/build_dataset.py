#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
build_dataset.py — собрать рабочий CSV из выгрузки Fort Monitor.

Приложение читает 39 колонок из 81, а xlsx на 49 тыс. строк разбирается
около 20 секунд. CSV с теми же значениями читается за 0,1 с — это и есть
причина, почему в публикацию едет CSV, а не исходная книга.

Запуск:
    python scripts/build_dataset.py <исходный.xlsx> data/telemetry_skif48.csv

Исходный xlsx должен быть с добавленными дельтами (см. make_delta_columns.py)
и с колонками, посчитанными из них:
    ΔПлощадь_га, ΔПробег_м, ΔЗабитие_на_рядок
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

SHEET = "Журнал принятых данных"

BASE_COLS = [
    "Время",
    "Широта",
    "Долгота",
    "101_7_Кол-во_сошников",
    "ΔПлощадь_га",
    "ΔПробег_м",
    "ΔЗабитие_на_рядок",
]
BLOCK_PREFIXES = ("221_", "222_", "223_", "224_")


def build(src: Path, dst: Path) -> pd.DataFrame:
    df = pd.read_excel(src, sheet_name=SHEET)

    flags = [c for c in df.columns if c and c.startswith(BLOCK_PREFIXES)]
    if not flags:
        raise SystemExit("Не найдены флаги забивания 221…224 — набор будет неполным")

    missing = [c for c in BASE_COLS if c not in df.columns]
    if missing:
        raise SystemExit(f"Нет колонок с дельтами: {missing}. "
                         f"Сначала прогоните make_delta_columns.py")

    out = df[BASE_COLS + flags].copy()

    # Время в ISO: pandas разбирает его без указания формата, это быстрее
    # и не зависит от локали.
    out["Время"] = pd.to_datetime(out["Время"], errors="coerce").dt.strftime(
        "%Y-%m-%d %H:%M:%S"
    )
    # Координаты оставляем текстом — так в исходной выгрузке, и это не даёт
    # потерять разрядность.
    for c in ("Широта", "Долгота"):
        out[c] = out[c].astype(str)

    out = out.dropna(subset=["Время"])
    dst.parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(dst, index=False, encoding="utf-8")

    size = dst.stat().st_size / 1024 / 1024
    print(f"Записано: {dst}")
    print(f"Строк: {len(out):,} | колонок: {out.shape[1]} | размер: {size:.1f} МБ")

    # Самопроверка: инвариант ΔS × 10000 / ΔL == ширина захвата.
    area = pd.to_numeric(out["ΔПлощадь_га"], errors="coerce").sum()
    dist = pd.to_numeric(out["ΔПробег_м"], errors="coerce").sum()
    if dist > 0:
        print(f"Контроль ширины: {area * 10000 / dist:.4f} м "
              f"(ожидается ширина из 101_8_ШиринаЗахвата_ПК / 10)")
    return out


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print(__doc__)
        return 2
    build(Path(argv[1]), Path(argv[2]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))