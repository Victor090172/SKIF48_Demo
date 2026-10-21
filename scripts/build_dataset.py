#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
build_dataset.py — собрать рабочий CSV из выгрузки Fort Monitor.

Приложение читает 43 колонки из 81, а xlsx на 49 тыс. строк разбирается
около 20 секунд. CSV с теми же значениями читается за 0,1 с — это и есть
причина, почему в публикацию едет CSV, а не исходная книга.

Запуск:
    python scripts/build_dataset.py <исходный.xlsx> data/telemetry_skif48.csv

Исходный xlsx должен быть с добавленными дельтами (см. make_delta_columns.py)
и с колонками, посчитанными из них:
    ΔПлощадь_га, ΔПробег_м, ΔЗабитие_на_рядок

Помимо отбора колонок скрипт считает НОРМУ СЕВА — см. расчёт нормы внизу.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

SHEET = "Журнал принятых данных"

# Колонки по бункерам: `Норма_номинал_б1`, `Норма_динам_б1`, `Откл_б1`,
# `ΔМасса_б1_кг` — цифра после «_б» это номер дозатора УДПВ.
BUNKER_RE = re.compile(r"_б(\d+)(?:_кг)?$")

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

# --- норма сева ------------------------------------------------------------
# Четыре дозатора УДПВ. Масштабы берутся из ref.param_def проекта:
#   101_29..101_32 НомНорм — 0.01 (кг/га)
#   104_12/14/16/18 производительность (г/оборот) — 0.1
#   104_13/15/17/19 кол-во полюсов — 1
# Формула массы — ref.formula rev_over_poles_times_grams:
#   {delta_counter}/{poles} * {grams_per_rev} / 1000   -> кг
RATE_WINDOW_HA = 0.5  # окно усреднения динамической нормы, га
NOMINAL_COLS = ["101_29_НомНорм_УДПВ1", "101_30_НомНорм_УДПВ2",
                "101_31_НомНорм_УДПВ3", "101_32_НомНорм_УДПВ4"]
DOSER = [  # (счётчик импульсов, полюсов, грамм на оборот)
    ("151_24_УДПВ1_Счетчик", "104_13_УДПВ1_кол-во_полюсов",
     "104_12_УДПВ1_производительность"),
    ("151_25_УДПВ2_Счетчик", "104_15_УДПВ2_кол-во_полюсов",
     "104_14_УДПВ2_производительность"),
    ("151_26_УДПВ3_Счетчик", "104_17_УДПВ3_кол-во_полюсов",
     "104_16_УДПВ3_производительность"),
    ("151_27_УДПВ4_Счетчик", "104_19_УДПВ4_кол-во_полюсов",
     "104_18_УДПВ4_производительность"),
]


RESET_FRACTION = 0.10  # падение больше 10 % от максимума = перезапуск счётчика


def counter_delta(s: pd.Series, reset_fraction: float = RESET_FRACTION) -> pd.Series:
    """Прирост монотонного счётчика, устойчивый к сбросу И к шуму.

    Наивный diff() + отсечение минусов (то, что делало safe_delta) считает
    сброс корректно, но дважды считает НЕМОНОТОННЫЙ ряд: в проблемных
    окнах выгрузки 151_24/151_25 чередуют две последовательности
    (88149, 88018, 88161, 88022, …) — одно и то же значение приходит то
    текущим, то отстающим. Каждое такое чередование давало ложный прирост
    в ~140 импульсов, то есть ~35 кг на строку, а по сумме — 4 группы с
    отклонением 400…2100 % и +2.4 % к общей норме.

    Правило одно: прирост считается от БЕГУЩЕГО МАКСИМУМА. Падение
    игнорируется (ряд возвращается к уже учтённой базе), но если упал
    дальше чем на `reset_fraction` — это настоящий перезапуск, и база
    сдвигается на новое значение.
    """
    v = np.asarray(s, dtype=float)
    out = np.zeros(v.size, dtype=float)
    if v.size == 0:
        return pd.Series(out, index=s.index)
    mx = v[0]
    for i in range(1, v.size):
        x = v[i]
        if x < mx:
            if mx > 0.0 and (mx - x) > reset_fraction * mx:
                mx = x                      # перезапуск: новая база
            # иначе ряд вернулся к уже учтённому максимуму — прироста нет
        else:
            out[i] = x - mx
            mx = x
    return pd.Series(out, index=s.index)


def last_known_config(s: pd.Series) -> pd.Series:
    """Последняя известная конфигурация дозатора вместо нулей в строках.

    Поля 104_13…104_19 и 104_12…104_18 в 154 строках приходят нулём —
    это не «дозатор выключили», а пропавший снимок конфигурации: счётчик
    в эти строки продолжает расти, а масса из-за нуля молча терялась
    (норма первой группы получалась −100 % при фактическом севе).
    Нулём в начале ряда остаётся только то, чего не было никогда.
    """
    raw = pd.to_numeric(s, errors="coerce").replace(0.0, np.nan)
    return raw.ffill().fillna(0.0)


def group_nominal(n_all: np.ndarray, starts: np.ndarray) -> np.ndarray:
    """Номинал группы — среднее ПО НЕНУЛЕВЫМ строкам.

    Обычное среднее занижало его: в первые минуты работы номинал приходит
    нулём (машину ещё не настроили), и вся группа получала 122.78 вместо
    200.28. Нулём остаётся только группа, где номинала не было никогда.
    """
    n_pos = np.where(n_all > 0, n_all, 0.0)
    n_cnt = np.where(n_all > 0, 1.0, 0.0)
    s_pos = np.add.reduceat(n_pos, starts)
    s_cnt = np.add.reduceat(n_cnt, starts)
    return np.where(s_cnt > 0, s_pos / np.where(s_cnt > 0, s_cnt, 1), 0.0)


def add_rates(df: pd.DataFrame) -> pd.DataFrame:
    """Добавляет колонки нормы сева — сводные и ПО КАЖДОМУ БУНКЕРУ.

    Сводные: `ΔМасса_кг`, `Норма_номинал`, `Норма_динам`, `Откл_нормы`.
    По бункеру (индекс N — номер дозатора УДПВ): `ΔМасса_бN_кг`,
    `Норма_номинал_бN`, `Норма_динам_бN`, `Откл_бN`.

    Зачем отдельно: суммарное отклонение скрывает локальный сбой —
    перелив одного дозатора и недолив другого дают ноль в сумме. На
    объекте 12917 бункеры фактически два (УДПВ3/УДПВ4 не настроены),
    и колонки пишутся только для настроенных: ненастроенный превратился
    бы в пустой индикатор в приложении.

    Динамическая норма усредняется по шагам накопленной площади окном
    RATE_WINDOW_HA — это осмысленный шаг для этого прибора: площадь
    обновляется только в 8 % строк, поэтому построчное отношение
    масса/площадь делится на ноль, а скользящее окно по строкам смещено
    на 3 п.п. из-за неравномерности пробега.
    """
    missing = [c for cols in (NOMINAL_COLS, *DOSER) for c in cols
               if c not in df.columns]
    if missing:
        raise SystemExit(f"Нет колонок нормы сева: {sorted(set(missing))}")

    area = pd.to_numeric(df["ΔПлощадь_га"], errors="coerce").fillna(0.0)
    a_all = area.to_numpy()

    # --- по каждому дозатору: масса и номинал ----------------------------
    per = []  # [{"idx": N, "mass": Series, "nom": Series}, ...]
    for n, ((cnt, poles, grams), nom_col) in enumerate(zip(DOSER, NOMINAL_COLS), 1):
        nom = pd.to_numeric(df[nom_col], errors="coerce").fillna(0.0) * 0.01
        p = last_known_config(df[poles])
        g = last_known_config(df[grams]) * 0.1
        # полюса 0 или граммы 0 = дозатор не настроен вовсе: массы нет,
        # и делить на ноль нельзя
        d = counter_delta(pd.to_numeric(df[cnt], errors="coerce").fillna(0.0))
        m = pd.Series(
            np.where((p > 0) & (g > 0),
                     d / p.replace(0, np.nan) * g / 1000, 0.0),
            index=df.index).fillna(0.0)
        # Рабочим считаем дозатор с ненулевым номиналом хотя бы в одной
        # строке либо с накопленной массой: остальные колонок не получают.
        if not ((nom > 0).any() or float(m.sum()) > 0):
            continue
        per.append({"idx": n, "mass": m, "nom": nom})
    if not per:
        raise SystemExit("Ни один дозатор не настроен: нет ни номинала, ни массы")

    # Сводные = сумма по настроенным дозаторам (ненастроенные дают 0,
    # так что это то же самое, что и по всем четырём).
    nominal = sum(b["nom"] for b in per)
    mass = sum(b["mass"] for b in per)

    # Группировка строк по накопленной площади: строка попадает в группу
    # floor(cum_area / W). Группа получает одно значение динамической нормы.
    cum = area.cumsum().to_numpy()
    gid = np.floor(cum / RATE_WINDOW_HA).astype(np.int64)

    m_all = mass.to_numpy()
    n_all = nominal.to_numpy()

    keys, starts, counts = np.unique(gid, return_index=True, return_counts=True)
    g_area = np.add.reduceat(a_all, starts)
    g_mass = np.add.reduceat(m_all, starts)
    # Номинал в группе — среднее ПО НЕНУЛЕВЫМ строкам. Обычное среднее
    # занижало его: в первые минуты работы номинал приходит нулём (машину
    # ещё не настроили), и вся группа получала 122.78 вместо 200.28.
    g_nom = group_nominal(n_all, starts)

    with np.errstate(invalid="ignore", divide="ignore"):
        g_rate = np.where(g_area > 1e-9, g_mass / np.where(g_area > 1e-9, g_area, 1), np.nan)
        g_dev = np.where(g_nom > 1e-9,
                         100.0 * (g_rate - g_nom) / np.where(g_nom > 1e-9, g_nom, 1),
                         np.nan)

    def spread(values: np.ndarray) -> pd.Series:
        """Значения, посчитанные на начало каждой группы, — на все строки.

        Пустоту до первой группы и после последней закрываем соседним
        значением: строка без своего окна не должна выпадать из трека.
        """
        return (pd.Series(values, index=starts)
                .reindex(df.index).ffill().bfill())

    out = df.copy()
    out["ΔМасса_кг"] = mass.round(3)
    out["Норма_номинал"] = spread(g_nom).fillna(0.0).round(2)
    out["Норма_динам"] = spread(g_rate).fillna(0.0).round(2)
    out["Откл_нормы"] = spread(g_dev).fillna(0.0).round(2)

    # --- по каждому бункеру ------------------------------------------------
    for b in per:
        n = b["idx"]
        m = b["mass"].to_numpy()
        nom = b["nom"].to_numpy()
        gb_nom = group_nominal(nom, starts)
        with np.errstate(invalid="ignore", divide="ignore"):
            gb_rate = np.where(g_area > 1e-9,
                               np.add.reduceat(m, starts) /
                               np.where(g_area > 1e-9, g_area, 1), np.nan)
            gb_dev = np.where(gb_nom > 1e-9,
                              100.0 * (gb_rate - gb_nom) /
                              np.where(gb_nom > 1e-9, gb_nom, 1), np.nan)
        out[f"ΔМасса_б{n}_кг"] = b["mass"].round(3)
        out[f"Норма_номинал_б{n}"] = spread(gb_nom).fillna(0.0).round(2)
        out[f"Норма_динам_б{n}"] = spread(gb_rate).fillna(0.0).round(2)
        out[f"Откл_б{n}"] = spread(gb_dev).fillna(0.0).round(2)

    # --- самопроверка: накопительная норма должна совпасть с номиналом -----
    # Номинал — величина на гектар, поэтому сравнивать его с суммой
    # накопленных масс надо через среднее, взвешенное по площади, а не
    # делить сумму построчных номиналов на площадь (это даёт бессмыслицу).
    a_tot = a_all.sum()
    checks = [("Свод", m_all, n_all)] + [
        (f"Бункер {b['idx']}", b["mass"].to_numpy(), b["nom"].to_numpy())
        for b in per
    ]
    for name, m_v, n_v in checks:
        m_tot = m_v.sum()
        w_nom = (n_v * a_all).sum() / a_tot if a_tot > 0 else 0.0
        if a_tot <= 0 or w_nom <= 0:
            continue
        dev = 100.0 * (m_tot / a_tot - w_nom) / w_nom
        print(f"Контроль нормы, {name}: {m_tot / a_tot:.2f} кг/га против номинала "
              f"{w_nom:.2f} (отклонение {dev:+.2f} %)")
        if abs(dev) > 10:
            print("  ВНИМАНИЕ: отклонение больше 10 % — проверьте масштабы "
                  "101_29/104_12 или формулу rev_over_poles_times_grams")
    print(f"Настроено бункеров: {', '.join(str(b['idx']) for b in per)}")
    return out


def build(src: Path, dst: Path) -> pd.DataFrame:
    df = pd.read_excel(src, sheet_name=SHEET)

    flags = [c for c in df.columns if c and c.startswith(BLOCK_PREFIXES)]
    if not flags:
        raise SystemExit("Не найдены флаги забивания 221…224 — набор будет неполным")

    missing = [c for c in BASE_COLS if c not in df.columns]
    if missing:
        raise SystemExit(f"Нет колонок с дельтами: {missing}. "
                         f"Сначала прогоните make_delta_columns.py")

    out = add_rates(df)

    # Время разбираем с ЯВНЫМ форматом: без него pandas выводит формат по
    # первым строкам, на них день месяца ≤ 12, и он принимает первый
    # разряд за месяц — тогда все дни > 12 (13–16 октября) превращаются в
    # NaT и отбрасываются. На этой книге так терялось 24 168 строк из 49 373.
    # Если формат вдруг другой — откатываемся на вывод pandas.
    t = pd.to_datetime(out["Время"], format="%d.%m.%Y %H:%M:%S", errors="coerce")
    if t.isna().any():
        t = pd.to_datetime(out["Время"], errors="coerce")
    out["Время"] = t.dt.strftime("%Y-%m-%d %H:%M:%S")
    # Координаты оставляем текстом — так в исходной выгрузке, и это не даёт
    # потерять разрядность.
    for c in ("Широта", "Долгота"):
        out[c] = out[c].astype(str)

    # Колонки по бункерам — только для настроенных дозаторов (см. add_rates).
    # Сортируем по номеру бункера, чтобы в CSV порядок был 1, 2, а не
    # алфавитный по типу колонки.
    bunker_cols = sorted(
        (c for c in out.columns if BUNKER_RE.search(c)),
        key=lambda c: (int(BUNKER_RE.search(c).group(1)), c),
    )
    cols = (BASE_COLS + ["Норма_номинал", "Норма_динам", "Откл_нормы", "ΔМасса_кг"]
            + bunker_cols + flags)
    out = out[cols]
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
