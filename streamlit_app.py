"""
СКИФ-48 · дашборд телеметрии сева
=================================

Прототип BI-дашборда по выгрузке телеметрии сеялки Fort Monitor (объект 12917,
СКИФ-48 Эл.Доз).

Запуск локально:
    pip install -r requirements.txt
    streamlit run streamlit_app.py

Публикация на https://share.streamlit.io/ — достаточно запушить репозиторий:
площадка сама выполнит `streamlit run streamlit_app.py`.

--------------------------------------------------------------------------------------
ЧТО ПОКАЗЫВАЕТ И КАК СЧИТАЕТСЯ
--------------------------------------------------------------------------------------

1. Карта: трек движения объекта, ДВА СЛОЯ раскраски (переключатель в
   сайдбаре).

   «Забитие сошников»: зелёный сегмент — движение без забития сошников;
   с забитыми — градиент от жёлтого (один забитый сошник) через оранжевый
   к красному (все 32). Сегмент красится по состоянию в его КОНЕЧНОЙ точке:
   забитие фиксируется в момент замера, значит участок, приведший к этому
   замеру, и есть участок с забитыми сошниками. Шкала — 32 сошника: флагов
   221…224 в выгрузке 32 штуки, по одному на сошник машины (block_color).

   «Отклонение нормы»: зелёный, пока модуль отклонения нормы от номинала в
   пределах порога (5 %), дальше градиент до красного (20 %). Цвет берётся
   по МАКСИМАЛЬНОМУ по модулю отклонению среди БУНКЕРОВ — суммарное
   отклонение скрывает локальный сбой. В подсказке указан бункер, а в
   попапе — строка на каждый бункер (см. rate_popup).

2. Показатели за выбранный интервал:

   Общая площадь            = Σ ΔПлощадь_га
   Общий пробег             = Σ (ΔПробег_м × 101_7_Кол-во_сошников)   [м·рядков]
   Пробег с забитием        = Σ (ΔЗабитие_на_рядок × кол-во забитых рядков)
   Доля с забитием          = Пробег с забитием / Общий пробег × 100
   Доля без забития         = 100 − Доля с забитием

   Норма сева — ПО КАЖДОМУ БУНКЕРУ (три карточки на бункер, один ряд):

   Средняя норма · Бункер N   = Σ ΔМасса_бN_кг / Σ ΔПлощадь_га      [кг/га]
   Номинал · Бункер N         = площадно-взвешенное Норма_номинал_бN [кг/га]
   Отклонение · Бункер N      = (средняя − номинал) / номинал × 100      [%]

   Суммарное отклонение НЕ показывается: перелив одного дозатора и недолив
   другого дают в сумме ноль, то есть локальный сбой остаётся незамеченным.
   Каждая величина считается от сумм, а не усреднением построчных значений:
   простое среднее смещено строками простоя (+14 % вместо −0,2 %).

   «Общий пробег» и «Пробег с забитием» оба в метро-рядках — это метраж пути,
   умноженный на число рядков. Единая размерность нужна для того, чтобы их
   отношение было корректной долей.

3. Доли подкрашиваются по трём уровням. Пороги задаются по шкале КАЧЕСТВА
   (доля без забития, где больше = лучше). Доля с забитием оценивается
   зеркально: зелёный, если забитие ≤ 100 − порог зелёного.

--------------------------------------------------------------------------------------
ПОЧЕМУ ТАК, А НЕ ИНАЧЕ
--------------------------------------------------------------------------------------

Исходные данные отдают только НАКОПИТЕЛЬНЫЕ счётчики (151_22 площадь,
151_23 пробег, 153_28 забивание). Показать их в BI можно только через дельты.

Счётчики перезапускаются на ноль. Наивная разность в момент сброса даёт
скачок −1786, а ABS() превращает его в фантомные +178,6 га, которые попадут
в SUM. Поэтому дельты считает scripts/make_delta_columns.py:

  · ведётся «высокая отметка» отдельно по каждому счётчику;
  · откат глубже половины отметки — СБРОС: начинается новая эпоха,
    дельта в этой строке 0;
  · откат мельче половины — шум передачи: дельта 0, отметка не двигается.

Проверка правильности — инвариант, который не должен расходиться:

    ΔПлощадь_га × 10000 / ΔПробег_м  ==  ширина захвата

На объекте 12917 ширина приходит как 101_8_ШиринаЗахвата_ПК = 60, то есть
6,0 м. Инвариант сошёлся в 6,0000 м. Если после правки расчёта он разъезжается
— сломаны дельты, а не данные.

--------------------------------------------------------------------------------------
ДАННЫЕ
--------------------------------------------------------------------------------------

data/telemetry_skif48.csv — производная от выгрузки Fort Monitor:
только те колонки, которые читает приложение. Собрать заново:

    python scripts/build_dataset.py <исходный.xlsx> data/telemetry_skif48.csv

Исходная выгрузка содержит реальные данные клиента (координаты техники,
идентификаторы объектов). Перед публикацией в открытом репозитории прогоните
scripts/sanitize_for_public.py.
"""

from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path

import colorsys
import folium
import numpy as np
import pandas as pd
import re
import streamlit as st
import streamlit_folium as st_folium

# ---------------------------------------------------------------------------
# Конфигурация
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
# CSV читается в ~240 раз быстрее xlsx (0,1 с против 22 с), поэтому он основной.
CSV_PATH = DATA_DIR / "telemetry_skif48.csv"
XLSX_PATH = DATA_DIR / "telemetry_skif48.xlsx"

# Пороги мотивации по умолчанию — задаются по шкале «качество» (доля без забития).
DEFAULT_GREEN_ABOVE = 95.0
DEFAULT_RED_BELOW = 80.0

# Цвета индикаторов и границ легенды порогов. COLOR_OK также красит зелёные
# сегменты трека; COLOR_BAD на самой карте не рисуется — там максимум шкалы
# считает block_color (почти совпадает с ним).
COLOR_OK = "#1a9850"    # зелёный — без забития / уровень «хорошо»
COLOR_BAD = "#d73027"   # красный — уровень «плохо» в порогах

# Шкала одновременно забитых сошников. Флагов 221…224 — 32 штуки, по одному
# на каждый сошник машины, поэтому счётчик идёт от 1 до 32.
BLOCK_SCALE_MAX = 32


def block_color(n: int) -> str:
    """Цвет сегмента трека по числу одновременно забитых сошников.

    Один забитый сошник — жёлтый, дальше через оранжевый к чистому красному
    на максимуме шкалы (там цвет совпадает с прежним COLOR_BAD, так что
    худший участок выглядит как раньше). Цвета интерполируются в HSL по
    каналу оттенка (48° → 3°), а не в RGB: середина шкалы даёт именно
    оранжевый, тогда как линейное смешивание красного с жёлтым в RGB даёт
    грязно-коричневый.
    """
    t = (min(max(n, 1), BLOCK_SCALE_MAX) - 1) / (BLOCK_SCALE_MAX - 1)
    h = (48.0 + (3.0 - 48.0) * t) / 360.0       # жёлтый -> красный
    s = (100.0 + (68.0 - 100.0) * t) / 100.0    # яркость красного чуть ниже
    l = (45.0 + (50.0 - 45.0) * t) / 100.0
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return f"#{round(r * 255):02x}{round(g * 255):02x}{round(b * 255):02x}"


def rate_color(dev: float, green: float, red: float) -> str:
    """Цвет сегмента по отклонению динамической нормы от номинальной.

    До `green` процентов — зелёный, дальше градиент зелёный → жёлтый →
    красный, от `red` — чистый красный. Отклонение берётся по модулю:
    недолив так же плох, как перелив.

    Градиент идёт в HSL по каналу оттенка 120° → 0° (зелёный → красный,
    середина 60° — жёлтый), как и у block_color: в RGB середина вышла бы
    грязно-коричневой.
    """
    a = abs(float(dev))
    if a <= green:
        return COLOR_OK
    if a >= red or red <= green:
        return COLOR_BAD
    t = (a - green) / (red - green)
    h = (120.0 * (1.0 - t)) / 360.0
    s = 1.0
    l = 0.42 + 0.06 * t          # зелёный чуть темнее, красный чуть ярче
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return f"#{round(r * 255):02x}{round(g * 255):02x}{round(b * 255):02x}"


def rate_level(dev: float, green: float, red: float) -> str:
    """Уровень показателя по модулю отклонения нормы — три те же зоны,
    что и в LEVEL_STYLE, только границы задаёт порог отклонения."""
    a = abs(float(dev))
    if a <= green:
        return "хорошо"
    if a < red:
        return "норма"
    return "плохо"


# Три уровня показателя: (цвет рамки, заливка, подпись)
LEVEL_STYLE = {
    "хорошо": ("#1a9850", "#eaf6ee"),
    "норма": ("#e6a700", "#fdf6e3"),
    "плохо": ("#d73027", "#fbeceb"),
}


def rate_bucket_color(bucket: float, green: float, red: float) -> str:
    """Цвет корзины отклонения — ровно тот, которым закрашен сегмент трека.

    Корзины считаются так же, как в build_map: 0 — зелёная зона (<= green),
    1 … RATE_STEPS — ступени внутри градиента, RATE_STEPS+1 — красная
    (>= red), отрицательные — данных о норме нет. Цвет берётся по середине
    диапазона корзины, поэтому полоса в легенде и счёт сегментов на карте
    совпадают один в один.
    """
    if bucket < 0:
        return "#9aa5b1"                       # нет данных о норме
    if bucket <= 0:
        return COLOR_OK                        # <= green
    if bucket >= RATE_STEPS + 1:
        return COLOR_BAD                       # >= red
    span = max(float(red - green), 1e-9)
    return rate_color(green + span * (bucket - 0.5) / RATE_STEPS, green, red)

BLOCK_FLAG_PREFIXES = ("221_", "222_", "223_", "224_")

# Ступеней раскраски по отклонению нормы. Внутри зоны градиента отклонение
# квантуется на RATE_STEPS корзин — так трек из 49 тыс. точек превращается в
# несколько сотен объектов, а цвета в легенде бьются со счётом сегментов.
# Зелёная зона (<= green) и красная (>= red) — своя корзина каждая.
RATE_STEPS = 10

# Сколько карточек показателей в одном ряду. Ряд — это flex-строка
# (см. card_row): на обычном экране шесть карточек укладываются в один
# ряд, дальше пришлось бы листать.
CARDS_IN_ROW = 6

# Колонки нормы сева. Считаются в scripts/build_dataset.py; в CSV, собранном
# старым скриптом, их нет — тогда соответствующие функции отключаются.
RATE_COLS = ("Норма_номинал", "Норма_динам", "Откл_нормы", "ΔМасса_кг")

# Колонки по каждому бункеру (дозатору УДПВ): `Норма_номинал_б1`,
# `Норма_динам_б1`, `Откл_б1`, `ΔМасса_б1_кг`. build_dataset.py пишет их
# только для НАСТРОЕННЫХ дозаторов, поэтому число бункеров приложение
# определяет по именам колонок, а не захардкожено.
BUNKER_RE = re.compile(r"_б(\d+)(?:_кг)?$")
BUNKER_NOM = "Норма_номинал_б{b}"
BUNKER_DYN = "Норма_динам_б{b}"
BUNKER_DEV = "Откл_б{b}"
BUNKER_MASS = "ΔМасса_б{b}_кг"


def bunker_ids(d: pd.DataFrame) -> list[int]:
    """Номера бункеров, по которым в наборе есть колонки (по возрастанию)."""
    return sorted({int(m.group(1)) for c in d.columns if (m := BUNKER_RE.search(c))})

# Пороги раскраски отклонения нормы по умолчанию. Пользователь задаёт их в
# сайдбаре: до green — зелёный, от red — красный, между ними градиент.
RATE_GREEN_DEFAULT = 5.0   # %
RATE_RED_DEFAULT = 20.0    # %

st.set_page_config(
    page_title="СКИФ-48 · телеметрия сева",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# Загрузка
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Читаю телеметрию…")
def load_data() -> pd.DataFrame:
    """Читает набор данных и приводит колонки к рабочим типам."""
    if CSV_PATH.exists():
        df = pd.read_csv(CSV_PATH)
    elif XLSX_PATH.exists():
        st.info("CSV не найден, читаю xlsx — первая загрузка займёт ~20 секунд.")
        df = pd.read_excel(XLSX_PATH, sheet_name="Журнал принятых данных")
    else:
        st.error(f"Файл данных не найден: ни {CSV_PATH.name}, ни {XLSX_PATH.name}")
        st.stop()

    df["ts"] = pd.to_datetime(df["Время"], errors="coerce")

    # Широта и долгота хранятся текстом намеренно: так в исходной выгрузке,
    # и это не даёт потерять разрядность при пересчёте.
    df["lat"] = pd.to_numeric(df["Широта"].astype(str), errors="coerce")
    df["lon"] = pd.to_numeric(df["Долгота"].astype(str), errors="coerce")

    df["d_area"] = pd.to_numeric(df["ΔПлощадь_га"], errors="coerce").fillna(0.0)
    df["d_dist"] = pd.to_numeric(df["ΔПробег_м"], errors="coerce").fillna(0.0)
    df["d_blk_row"] = pd.to_numeric(df["ΔЗабитие_на_рядок"], errors="coerce").fillna(0.0)
    df["n_rows"] = pd.to_numeric(df["101_7_Кол-во_сошников"], errors="coerce").fillna(32.0)

    # Количество забитых рядков считаем из самих пофакторных флагов 221…224,
    # а НЕ из производной колонки «Есть забитие»: в исходной выгрузке она
    # живёт формулой =SUM(...), и всё, что читает файл не из Excel (pandas,
    # BI-коннектор), видит на её месте пустые ячейки. Флаги — единственный
    # источник, не зависящий от пересчёта формул.
    flag_cols = [c for c in df.columns if c and c.startswith(BLOCK_FLAG_PREFIXES)]
    if not flag_cols:
        st.warning("Флаги забивания 221…224 не найдены — карта будет одноцветной.")
        df["n_blocked"] = 0.0
    else:
        flags = df[flag_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
        df["n_blocked"] = flags.sum(axis=1)

    df["blocked"] = df["n_blocked"] > 0

    # Норма сева: считается в scripts/build_dataset.py (см. там же формулу
    # rev_over_poles_times_grams и окно усреднения по шагам площади).
    # Колонок может не быть, если CSV собран старым скриптом — тогда
    # переключатель «норма» прячется, а приложение работает как раньше.
    have_rates = all(c in df.columns for c in RATE_COLS)
    if have_rates:
        for src, dst_ in (("Норма_номинал", "rate_nom"),
                          ("Норма_динам", "rate_dyn"),
                          ("Откл_нормы", "rate_dev")):
            df[dst_] = pd.to_numeric(df[src], errors="coerce")
        df["rate_mass"] = pd.to_numeric(df["ΔМасса_кг"], errors="coerce").fillna(0.0)
        # Отклонение, у которого за спиной есть и площадь, и масса: на строках
        # простоя динамическая норма фиктивно переносится из прошлого шага и
        # в среднем участвовать не должна.
        df["rate_valid"] = (df["rate_dev"].notna()
                            & (df["d_area"] > 0) & df["rate_nom"].gt(0))
    else:
        for c in ("rate_nom", "rate_dyn", "rate_dev"):
            df[c] = np.nan
        df["rate_mass"] = 0.0
        df["rate_valid"] = False

    # Колонки по бункерам. Приводим к числам (в CSV едут как есть, но
    # координаты мы сознательно оставляем текстом — вдруг попадётся битая).
    # Масса без NaN: суммировать по ней надо всегда, даже если часть строк
    # бункера не заполнена.
    for b in bunker_ids(df):
        df[BUNKER_MASS.format(b=b)] = pd.to_numeric(
            df[BUNKER_MASS.format(b=b)], errors="coerce").fillna(0.0)
        for tpl in (BUNKER_NOM, BUNKER_DYN, BUNKER_DEV):
            df[tpl.format(b=b)] = pd.to_numeric(df[tpl.format(b=b)], errors="coerce")

    df = df.dropna(subset=["ts", "lat", "lon"]).sort_values("ts").reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# Показатели
# ---------------------------------------------------------------------------
def kpis(d: pd.DataFrame) -> dict:
    """Показатели за выбранный срез."""
    total = float((d["d_dist"] * d["n_rows"]).sum())
    blocked = float((d["d_blk_row"] * d["n_blocked"]).sum())
    area = float(d["d_area"].sum())
    pct_blocked = 100.0 * blocked / total if total > 0 else 0.0

    # Норма сева. Обе величины считаются от сумм, а не усреднением
    # построчных: средняя норма — это масса на площадь, а отклонение средней
    # нормы от среднего номинала. Площадно-взвешенные номинал и отклонение
    # совпадают с контрольным расчётом в build_dataset; простое среднее
    # построчных отклонений дало бы 14 % — его смещают строки простоя, где
    # динамическая норма перенесена из прошлого шага.
    area_w = float(d["d_area"].sum())
    if "rate_valid" in d.columns and area_w > 0:
        m = float(d["rate_mass"].to_numpy().sum())
        n_w = float((d["rate_nom"].to_numpy() * d["d_area"].to_numpy()).sum()) / area_w
        avg_rate = m / area_w            # кг/га — как в контроле build_dataset
        avg_nom = n_w                    # площадно-взвешенный номинал
        # Отклонение — площадно-взвешенное среднее построчных значений:
        # простое среднее смещено строками простоя (14 % против фактических 1.6 %)
        ok = d["rate_valid"].to_numpy() & np.isfinite(d["rate_dev"].to_numpy())
        if ok.any() and float(d.loc[ok, "d_area"].to_numpy().sum()) > 0:
            w = d.loc[ok, "d_area"].to_numpy()
            pct_rate = float((d.loc[ok, "rate_dev"].to_numpy() * w).sum() / w.sum())
        else:
            pct_rate = float("nan")
    else:
        avg_rate = avg_nom = pct_rate = float("nan")

    # По каждому бункеру — те же суммы, только «своя» масса и «свой» номинал.
    # Суммарное отклонение скрывает локальный сбой: перелив одного дозатора
    # и недолив другого дают ноль в сумме, а здесь каждый виден отдельно.
    bunkers = []
    for b in bunker_ids(d):
        m = float(d[BUNKER_MASS.format(b=b)].to_numpy().sum())
        nom = d[BUNKER_NOM.format(b=b)].to_numpy()
        n_w = (float((nom * d["d_area"].to_numpy()).sum()) / area_w
               if area_w > 0 else float("nan"))
        avg = m / area_w if area_w > 0 else float("nan")
        dev = (100.0 * (avg - n_w) / n_w
               if np.isfinite(n_w) and n_w > 0 else float("nan"))
        bunkers.append({"id": b, "avg": avg, "nom": n_w, "dev": dev})

    return {
        "area_ha": area,
        "total_mrows": total,
        "blocked_mrows": blocked,
        "pct_blocked": pct_blocked,
        "pct_clean": 100.0 - pct_blocked,
        "n_points": int(len(d)),
        "avg_rate": avg_rate,
        "nom_rate": avg_nom,
        "pct_rate_dev": pct_rate,
        "has_rate": bool(pct_rate == pct_rate),  # NaN -> False
        "bunkers": bunkers,
    }


def grade_quality(pct_clean: float, red_below: float, green_above: float) -> str:
    """Уровень по доле пробега БЕЗ забития: чем больше, тем лучше."""
    if pct_clean >= green_above:
        return "хорошо"
    if pct_clean >= red_below:
        return "норма"
    return "плохо"


def grade_blockage(pct_blocked: float, red_below: float, green_above: float) -> str:
    """Уровень по доле пробега С забитием: чем меньше, тем лучше."""
    if pct_blocked <= 100.0 - green_above:
        return "хорошо"
    if pct_blocked <= 100.0 - red_below:
        return "норма"
    return "плохо"


# ---------------------------------------------------------------------------
# Карточки показателей — одна строка
# ---------------------------------------------------------------------------
def status_card(label: str, value: str, level: str, unit: str = "") -> str:
    """Карточка с цветовой рамкой. Уровень не подписывается словом:
    цвет рамки и цифры говорят сами."""
    accent, tint = LEVEL_STYLE[level]
    unit_html = (f" <span style='font-size:.68rem;font-weight:400;"
                 f"color:{accent};opacity:.75'>{unit}</span>" if unit else "")
    return (
        f"<div style='flex:1 1 0;min-width:0;border:2px solid {accent};"
        f"background:{tint};border-radius:8px;padding:7px 10px'>"
        # Подпись намеренно без nowrap: в ряду шесть карточек, и «Отклонение
        # от номинальной · Бункер 1» одной строкой не помещается. min-height
        # держит ряд ровным, когда у соседей подпись в одну строку.
        f"<div style='font-size:.7rem;color:#555;line-height:1.25;"
        f"min-height:2.4em;overflow:hidden'>{label}</div>"
        f"<div style='font-size:1.2rem;font-weight:700;color:{accent};"
        f"line-height:1.25;white-space:nowrap'>{value}{unit_html}</div></div>"
    )


def plain_card(label: str, value: str, unit: str = "") -> str:
    """Карточка нейтрального показателя."""
    return (
        f"<div style='flex:1 1 0;min-width:0;border:1px solid #d7dde3;"
        f"background:#fff;border-radius:8px;padding:7px 10px'>"
        f"<div style='font-size:.7rem;color:#555;line-height:1.25;"
        f"min-height:2.4em;overflow:hidden'>{label}</div>"
        f"<div style='font-size:1.05rem;font-weight:600;color:#1f2933;"
        f"line-height:1.4;white-space:nowrap'>{value} "
        f"<span style='font-size:.68rem;font-weight:400;color:#777'>{unit}</span></div></div>"
    )


def card_row(*cards: str) -> str:
    """Все показатели в одну строку.

    flex вместо st.columns: у колонок Streamlit широкие внутренние отступы, и
    шесть карточек в ряд не помещались по ширине.
    """
    return "<div style='display:flex;gap:7px;align-items:stretch;'>" + "".join(cards) + "</div>"


# ---------------------------------------------------------------------------
# Карта
# ---------------------------------------------------------------------------
def rate_popup(per_b: dict, sdev: np.ndarray, svom: np.ndarray, sdyn: np.ndarray,
               i: int, focus: int) -> str:
    """Всплывающее окно нормы сева на ломаной трека.

    По строке на каждый бункер: своя динамическая норма против своего
    номинала и своё отклонение. Суммарную цифру не показываем — она как
    раз и скрывает локальный сбой. `focus` — бункер, чьим отклонением
    покрашен сегмент: его строка идёт первой, чтобы читать вместе с цветом.

    Если колонок по бункерам нет (CSV собран старым скриптом) — одна
    сводная строка, как в первой версии окна.
    """
    rows = ["<b>Норма сева</b>"]
    if per_b:
        for b in sorted(per_b, key=lambda x: (x != focus, x)):
            s = per_b[b]
            rows.append(
                f"Бункер {b}: <b>{s['dyn'][i]:.2f}</b> из {s['nom'][i]:.2f} кг/га "
                f"— отклонение <b>{s['dev'][i]:+.2f} %</b>"
            )
    else:
        rows.append(f"Номинальная: <b>{svom[i]:.2f}</b> кг/га")
        rows.append(f"Динамическая: <b>{sdyn[i]:.2f}</b> кг/га")
        rows.append(f"Отклонение: <b>{sdev[i]:+.2f} %</b>")
    return (
        "<div style='font-family:sans-serif;font-size:13px;line-height:1.5'>"
        + "<br>".join(rows)
        + "</div>"
    )


def build_map(d: pd.DataFrame, simplify_m: float, mode: str = "blockage",
              green: float = RATE_GREEN_DEFAULT,
              red: float = RATE_RED_DEFAULT) -> folium.Map:
    """Карта с треком в двух режимах раскраски.

    mode="blockage" — зелёный без забития, дальше жёлтый → красный по числу
    одновременно забитых сошников.
    mode="rate" — зелёный в пределах порога отклонения нормы, дальше
    градиент зелёный → жёлтый → красный. Цвет сегмента определяет
    максимальное по модулю отклонение среди БУНКЕРОВ: так виден локальный
    сбой, который в сумме скрывает недолив другого дозатора. Всплывающее
    окно — по строке на каждый бункер (своя норма и своё отклонение).

    Оба режима рисуют ОДИН трек, но разными слоями: переключение — это
    смена слоя, а не наложение двух треков друг на друга.
    """
    center = [float(d["lat"].mean()), float(d["lon"].mean())]
    fmap = folium.Map(location=center, zoom_start=14, tiles="OpenStreetMap",
                      control_scale=True)

    if len(d) < 2:
        return fmap

    lat = d["lat"].to_numpy()
    lon = d["lon"].to_numpy()

    use_rate = mode == "rate" and (
        bool(bunker_ids(d))
        or ("rate_dev" in d.columns and d["rate_dev"].notna().any())
    )

    if use_rate:
        ids = bunker_ids(d)
        # Цвет сегмента — по максимальному по модулю отклонению среди
        # БУНКЕРОВ: так на треке виден локальный сбой, который в сумме
        # скрывает недолив другого дозатора (перелив одного и недолив
        # другого дают ноль). Если бункеров в наборе нет — сводное
        # отклонение, как раньше.
        if ids:
            arr = np.stack([d[BUNKER_DEV.format(b=b)].to_numpy(dtype=float)
                            for b in ids])
            mag = np.where(np.isnan(arr), -np.inf, np.abs(arr))
            has = np.isfinite(mag).any(axis=0)
            pick = np.argmax(mag, axis=0)
            rows = np.arange(arr.shape[1])
            dev = np.where(has, arr[pick, rows], np.nan)
            src = np.where(has, np.asarray(ids, dtype=int)[pick], 0)
        else:
            dev = d["rate_dev"].to_numpy(dtype=float)
            src = np.zeros(len(d), dtype=int)

        # Квантуем отклонение не в абсолютные значения, а в КОРЗИНЫ по
        # порогам: своя корзина для зелёной зоны, RATE_STEPS корзин внутри
        # градиента, своя для красной. Иначе зелёная зона в 0.5 % (в ней
        # подавляющее большинство точек) рассыпалась бы на кучу разных ключей
        # и не склеивалась в ломаные, а цвет перестал бы совпадать с легендой.
        a = np.abs(dev)
        span = max(float(red - green), 1e-9)
        g_ok, g_bad = 0.0, float(RATE_STEPS + 1)
        key = np.where(
            ~np.isfinite(a), -1.0,
            np.where(a <= green, g_ok,
                     np.where(a >= red, g_bad,
                              1.0 + np.floor((a - green) / span * RATE_STEPS))),
        )
    else:
        # Число одновременно забитых сошников: пофакторные флаги 221…224 дают
        # целые 0…32. От него зависит цвет сегмента.
        key = d["n_blocked"].to_numpy().astype(float)

    # Прореживание: оставляем точку, только если она ушла от последней
    # сохранённой дальше чем на simplify_m. Геометрия трека не меняется.
    keep = np.ones(len(lat), dtype=bool)
    if simplify_m > 0:
        dlat = np.diff(lat) * 111_320.0
        dlon = np.diff(lon) * 111_320.0 * np.cos(np.radians(lat[:-1]))
        keep[1:] = np.hypot(dlat, dlon) > simplify_m
        keep[-1] = True

    idx = np.flatnonzero(keep)
    slat, slon = lat[idx], lon[idx]
    sk = key[idx]
    if use_rate:
        sdev, ssrc = dev[idx], src[idx]
        # Для попапа — срезы по каждому бункеру: там показываем свой номинал,
        # свою динамическую норму и своё отклонение, а не сводные цифры.
        per_b = {
            b: {
                "nom": d[BUNKER_NOM.format(b=b)].to_numpy(dtype=float)[idx],
                "dyn": d[BUNKER_DYN.format(b=b)].to_numpy(dtype=float)[idx],
                "dev": d[BUNKER_DEV.format(b=b)].to_numpy(dtype=float)[idx],
            }
            for b in bunker_ids(d)
        }
        # Сводные номинал и динамическая норма — в попап они идут, только
        # когда колонок по бункерам в наборе нет (CSV собран старым
        # скриптом), в остальных случаях строки идут по бункерам.
        svom = (d["rate_nom"].to_numpy(dtype=float)[idx]
                if "rate_nom" in d.columns else np.full(len(idx), np.nan))
        sdyn = (d["rate_dyn"].to_numpy(dtype=float)[idx]
                if "rate_dyn" in d.columns else np.full(len(idx), np.nan))

    seg = sk[1:]  # цвет сегмента определяется его концом
    if len(seg) == 0:
        return fmap

    # Подряд идущие сегменты одного цвета склеиваем в одну ломаную:
    # ~49 тыс. точек превращаются в несколько сотен объектов карты.
    bounds = np.r_[np.flatnonzero(np.r_[True, seg[1:] != seg[:-1]]), len(seg)]
    for start, stop in zip(bounds[:-1], bounds[1:]):
        pts = list(zip(slat[start:stop + 1].tolist(), slon[start:stop + 1].tolist()))
        if len(pts) < 2:
            continue
        k = seg[start]
        if use_rate:
            if k < 0:
                # норма на этом отрезке не определена — рисуем нейтрально,
                # в раскраску по отклонению она не участвует
                color, tip, popup = "#9aa5b1", "нет данных о норме", None
            else:
                # в попапе — значения на середине сегмента, не на его краю:
                # так цифра ближе к тому, что показано цветом
                mid = (start + stop) // 2
                color = rate_bucket_color(k, green, red)
                # Кто дал цвет сегменту: если выбран конкретный бункер,
                # называем его и в подсказке — иначе непонятно, к чему
                # относится выделенный красным участок.
                b = int(ssrc[mid])
                tip = (f"Бункер {b}: отклонение {sdev[mid]:+.1f} %"
                       if b else f"отклонение {sdev[mid]:+.1f} %")
                popup = rate_popup(per_b, sdev, svom, sdyn, mid, b)
        else:
            n = int(k)
            color = COLOR_OK if n == 0 else block_color(n)
            tip = "без забития" if n == 0 else f"забито сошников: {n}"
            popup = None

        pl = folium.PolyLine(
            pts,
            color=color,
            weight=4,
            opacity=0.85,
            tooltip=tip,
        )
        if popup:
            pl.add_child(folium.Popup(popup, max_width=280))
        pl.add_to(fmap)

    folium.Marker(
        [slat[0], slon[0]], tooltip="Старт",
        popup=f"Начало участка<br>{d['ts'].iloc[0]:%d.%m.%Y %H:%M:%S}",
        icon=folium.Icon(color="blue", icon="play"),
    ).add_to(fmap)
    folium.Marker(
        [slat[-1], slon[-1]], tooltip="Финиш",
        popup=f"Конец участка<br>{d['ts'].iloc[-1]:%d.%m.%Y %H:%M:%S}",
        icon=folium.Icon(color="gray", icon="stop"),
    ).add_to(fmap)

    folium.LayerControl().add_to(fmap)

    # Перед словом «Leaflet» в атрибуции рисуется сине-жёлтый прямоугольник.
    # Это фирменный знак самой библиотеки Leaflet: он подставляется по
    # умолчанию в ЛЮБУЮ карту на Leaflet — это не выбранный нами провайдер.
    # Убираем его вместе со ссылкой и разделителем; обязательная подпись
    # «© OpenStreetMap contributors» остаётся (так требует лицензия ODbL).
    #
    # Прячем стилями, а не JS: streamlit-folium пересобирает JS карты сам и
    # теряет всё, что добавлено в root.script, а root.header отдаёт как есть.
    fmap.get_root().header.add_child(folium.Element(
        "<style>"
        ".leaflet-control-attribution"
        " a[title='A JavaScript library for interactive maps'],"
        ".leaflet-control-attribution span[aria-hidden='true']"
        " { display: none; }"
        "</style>"
    ))
    return fmap


# ---------------------------------------------------------------------------
# Календарь
# ---------------------------------------------------------------------------
def dt_picker(label: str, dmin: datetime, dmax: datetime, key: str) -> datetime:
    """Граница периода: календарь + выбор времени двумя полями.

    Отдельного st.datetime_input в этой ветке Streamlit нет, а пара
    date_input + time_input — это ровно тот «традиционный календарь»,
    который и нужен: дата выбирается стрелками, время — списком.

    ВНИМАНИЕ, из-за чего ломалось на хостинге: значение виджета НЕ задаётся
    через `value=`. Оно живёт только в st.session_state, а виджет создаётся
    без value. Одновременная передача value= и ключа, уже присутствующего в
    session_state, на свежих версиях Streamlit поднимает StreamlitValueError.
    Локально на 1.51 это проходило, на хостинге — нет.
    """
    col_d, col_t = st.columns([1, 1], gap="small")
    d = col_d.date_input(
        f"{label} · дата",
        min_value=dmin.date(), max_value=dmax.date(),
        format="DD.MM.YYYY", key=f"{key}_date",
    )
    t = col_t.time_input(
        f"{label} · время",
        step=timedelta(minutes=1), key=f"{key}_time",
    )
    return datetime.combine(d, t)


# ---------------------------------------------------------------------------
# Интерфейс
# ---------------------------------------------------------------------------
def main() -> None:
    # Сайдбар по умолчанию тесноват: обрезались календари, а кириллица в
    # кнопках быстрого выбора распадалась на буквы по вертикали.
    #
    # Именно st.html, а не st.markdown(unsafe_allow_html=True): st.markdown
    # прогоняет содержимое через санитайзер и подмешивает узлы в дерево,
    # которым React управляет сам. На хостинге это давало падение фронтенда
    # «NotFoundError: removeChild ... узел не является дочерним». st.html —
    # штатный способ внедрять HTML, он с этим не сталкивается.
    st.html(
        """
        <style>
          [data-testid="stSidebar"] { min-width: 380px; max-width: 380px; }
          [data-testid="stSidebar"] button { white-space: nowrap; }
          [data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.35rem; }
          [data-testid="stSidebar"] [data-baseweb="input"] { font-size: 0.8rem; }
        </style>
        """
    )

    st.title("СКИФ-48 · телеметрия сева")
    st.caption(
        "Карта в двух слоях: «Забитие сошников» — зелёный без забития, дальше "
        "жёлтый → красный по числу одновременно забитых; «Отклонение нормы» — "
        "зелёный в пределах порога, дальше градиент до красного. Слой выбирается "
        "в сайдбаре. Показатели считаются только за выбранный интервал."
    )

    df = load_data()
    if df.empty:
        st.error("После фильтрации не осталось ни одной точки.")
        st.stop()

    t_min, t_max = df["ts"].min(), df["ts"].max()
    d_lo = t_min.floor("min").to_pydatetime()
    d_hi = t_max.floor("min").to_pydatetime()

    with st.sidebar:
        st.header("Период")

        for k, v in (("from_date", d_lo.date()), ("from_time", d_lo.time()),
                     ("to_date", d_hi.date()), ("to_time", d_hi.time())):
            st.session_state.setdefault(k, v)

        # Начальные значения кладём в session_state, а виджеты создаём без
        # value= — тогда источник значения ровно один, и версия Streamlit
        # ни при чём не мешает.
        sel_from = dt_picker("С", d_lo, d_hi, "from")
        sel_to = dt_picker("По", d_lo, d_hi, "to")

        if sel_from > sel_to:
            st.error("Начало периода позже конца — поменяйте границы.")
            st.stop()

        # Быстрые периоды — под календарями, одной строкой. Переключатель стоит
        # НИЖЕ полей, которыми управляет, поэтому задать session_state обычным
        # порядком нельзя: работает колбэк, он выполняется ДО перезапуска
        # скрипта, когда поля ещё не созданы.
        #
        # Параметр необязательный: Streamlit зовёт on_change у этого виджета без
        # аргументов, и жёсткая сигнатура даёт TypeError.
        def apply_preset(name: str | None = None) -> None:
            name = name or st.session_state.get("preset")
            days = {"День": 1, "Неделя": 7, "Месяц": 30, "Всё": None}.get(name)
            if days is None and name != "Всё":
                return
            start = t_min if days is None else max(
                t_min, t_max - pd.Timedelta(days=days)
            )
            start = start.floor("min").to_pydatetime()
            st.session_state["from_date"] = start.date()
            st.session_state["from_time"] = start.time()
            st.session_state["to_date"] = d_hi.date()
            st.session_state["to_time"] = d_hi.time()

        st.segmented_control(
            "Период", ["День", "Неделя", "Месяц", "Всё"],
            key="preset", on_change=apply_preset,
            label_visibility="collapsed", width="stretch",
        )

        st.divider()
        st.header("Пороги мотивации")
        st.caption(
            "Три уровня. Задаются по шкале «качество» — доле пробега БЕЗ забития. "
            "Доля С забитием оценивается зеркально: 100 % минус порог."
        )
        green_above = st.number_input(
            "Зелёный: качество от (%)", 0.0, 100.0,
            DEFAULT_GREEN_ABOVE, 0.5,
            help="Качество выше этого значения — зелёный уровень.",
        )
        red_below = st.number_input(
            "Жёлтый: качество от (%)", 0.0, 100.0,
            DEFAULT_RED_BELOW, 0.5,
            help="Ниже этого значения — красный.",
        )
        if green_above <= red_below:
            st.warning("Порог зелёного должен быть ВЫШЕ порога жёлтого, иначе жёлтая зона пуста.")
        st.markdown(
            f"<div style='font-size:.78rem;line-height:1.7;color:#555'>"
            f"<span style='color:{COLOR_OK};font-weight:700'>■</span> ≥ {green_above:g} % "
            f"&nbsp;<span style='color:#e6a700;font-weight:700'>■</span> "
            f"{red_below:g}–{green_above - 0.5:g} % "
            f"&nbsp;<span style='color:{COLOR_BAD};font-weight:700'>■</span> &lt; {red_below:g} %"
            f"</div>",
            unsafe_allow_html=True,
        )

        st.divider()
        st.header("Карта")
        simplify = st.slider(
            "Прореживание точек, м", 0, 50, 3, 1,
            help="0 — рисовать каждую точку. Больше — легче карта, форма трека та же.",
        )

        # Переключатель слоя раскраски. «Норма» показывается только когда в
        # наборе есть колонки нормы — иначе пункт вёл бы в пустой трек.
        # Проверка наличия колонки обязательна: load_data кладёт rate_dev и
        # в пустом виде (NaN), но полагаться на это — полагаться на порядок
        # кода в другом месте файла.
        has_rate = bool("rate_dev" in df.columns
                        and df["rate_dev"].notna().any())
        if has_rate:
            layer = st.segmented_control(
                "Слой раскраски", ["Забитие сошников", "Отклонение нормы"],
                key="map_layer", label_visibility="collapsed", width="stretch",
            )
            mode = "rate" if layer == "Отклонение нормы" else "blockage"
        else:
            layer, mode = None, "blockage"
            st.caption("Колонок нормы сева в наборе нет — доступна только раскраска по забитию.")

        if mode == "rate":
            st.caption(
                "Пороги отклонения динамической нормы от номинальной. "
                "Отклонение берётся по модулю: недолив так же плох, как перелив."
            )
            r_green = st.number_input(
                "Зелёный: отклонение до (%)", 0.0, 500.0,
                RATE_GREEN_DEFAULT, 0.5,
                key="rate_green",
                help="Модуль отклонения не выше этого значения — зелёный.",
            )
            r_red = st.number_input(
                "Красный: отклонение от (%)", 0.0, 500.0,
                RATE_RED_DEFAULT, 0.5,
                key="rate_red",
                help="С этого значения — чистый красный, между порогами градиент.",
            )
            if r_red <= r_green:
                st.warning("Красный порог должен быть ВЫШЕ зелёного, иначе градиенту негде развернуться.")
                r_red = r_green + 0.5
            st.markdown(
                f"<div style='font-size:.78rem;line-height:1.7;color:#555'>"
                f"<span style='color:{COLOR_OK};font-weight:700'>■</span> ≤ {r_green:g} % "
                f"&nbsp;<span style='color:#e6a700;font-weight:700'>■</span> "
                f"{r_green:g}–{r_red:g} % "
                f"&nbsp;<span style='color:{COLOR_BAD};font-weight:700'>■</span> &gt; {r_red:g} %"
                f"</div>",
                unsafe_allow_html=True,
            )
        else:
            r_green, r_red = RATE_GREEN_DEFAULT, RATE_RED_DEFAULT

        show_legend = st.checkbox("Показывать легенду", value=True)

    # --- срез ---------------------------------------------------------------
    d = df[(df["ts"] >= sel_from) & (df["ts"] <= sel_to)]
    if d.empty:
        st.warning("В выбранном интервале нет данных. Сдвиньте границы.")
        st.stop()

    k = kpis(d)
    lvl_clean = grade_quality(k["pct_clean"], red_below, green_above)
    lvl_blk = grade_blockage(k["pct_blocked"], red_below, green_above)

    st.markdown(
        card_row(
            status_card("Доля без забития", f"{k['pct_clean']:.2f} %", lvl_clean),
            status_card("Доля с забитием", f"{k['pct_blocked']:.3f} %", lvl_blk),
            plain_card("Площадь", f"{k['area_ha']:,.1f}", "га"),
            plain_card("Общий пробег", f"{k['total_mrows']:,.0f}", "м·ряд."),
            plain_card("Пробег с забитием", f"{k['blocked_mrows']:,.0f}", "м·ряд."),
            plain_card("Пробег без забития",
                       f"{k['total_mrows'] - k['blocked_mrows']:,.0f}", "м·ряд."),
        ),
        unsafe_allow_html=True,
    )

    # Норма сева — ПО КАЖДОМУ БУНКЕРУ отдельно. Суммарное отклонение скрывает
    # локальный сбой: перелив одного дозатора и недолив другого дают ноль
    # в сумме, а по бункерам каждый виден. По три карточки на бункер
    # (средняя норма, номинал, отклонение); при двух бункерах это ровно
    # один ряд. Больше двух бункеров — ряды по CARDS_IN_ROW.
    bunkers = k["bunkers"]
    if bunkers:
        per_row = max(1, CARDS_IN_ROW // 3)
        for start in range(0, len(bunkers), per_row):
            cards = []
            for b in bunkers[start:start + per_row]:
                finite = np.isfinite(b["dev"])
                lvl = rate_level(b["dev"], r_green, r_red) if finite else "норма"
                n = b["id"]
                avg_t = f"{b['avg']:,.2f}" if np.isfinite(b["avg"]) else "—"
                nom_t = f"{b['nom']:,.2f}" if np.isfinite(b["nom"]) else "—"
                dev_t = f"{b['dev']:+.2f} %" if finite else "—"
                cards.append(status_card(f"Средняя норма сева<br>Бункер {n}",
                                         avg_t, lvl, "кг/га"))
                cards.append(plain_card(f"Номинальная норма<br>Бункер {n}",
                                        nom_t, "кг/га"))
                cards.append(status_card(f"Отклонение от номинальной<br>Бункер {n}",
                                         dev_t, lvl))
            st.markdown(card_row(*cards), unsafe_allow_html=True)

        st.caption(
            "Норма: "
            + " · ".join(
                f"Бункер {b['id']} — {b['avg']:,.2f} из {b['nom']:,.2f} кг/га "
                f"({b['dev']:+.2f} %)"
                for b in bunkers if np.isfinite(b["dev"])
            )
        )

    st.caption(
        f"Выбрано: {pd.Timestamp(sel_from):%d.%m.%Y %H:%M} — "
        f"{pd.Timestamp(sel_to):%d.%m.%Y %H:%M} · "
        f"точек в срезе: {k['n_points']:,}"
    )

    st_folium.st_folium(
        build_map(d, float(simplify), mode=mode, green=r_green, red=r_red),
        width=None, height=620, use_container_width=True,
    )

    if show_legend:
        if mode == "rate":
            # Легенда слоя «Отклонение нормы»: те же корзины и те же цвета,
            # которыми закрашен трек (см. rate_bucket_color). Ось полосы
            # идёт от 0 до red + (red − green): зелёная зона укладывается
            # в начало, а красная имеет заметный хвост, а не схлопывается
            # в нулевую ширину на самом краю.
            axis = r_red + (r_red - r_green) if r_red > r_green else max(r_red, 1.0)
            grad_parts = [f"{COLOR_OK} 0%",
                          f"{COLOR_OK} {100.0 * r_green / axis:.0f}%"]
            for bk in range(1, RATE_STEPS + 1):
                pos = r_green + (r_red - r_green) * (bk - 0.5) / RATE_STEPS
                grad_parts.append(
                    f"{rate_bucket_color(bk, r_green, r_red)} {100.0 * pos / axis:.0f}%")
            grad_parts.append(f"{COLOR_BAD} {100.0 * r_red / axis:.0f}%")
            grad_parts.append(f"{COLOR_BAD} 100%")
            grad = ", ".join(grad_parts)
            st.markdown(
                f'<span style="color:{COLOR_OK};font-weight:700">■</span> '
                f'≤ {r_green:g} %'
                f' &nbsp;&nbsp;<span style="font-size:.78rem;color:#666">'
                f'модуль отклонения нормы:</span> '
                f'<span style="display:inline-block;width:130px;height:12px;'
                f'vertical-align:middle;border-radius:3px;'
                f'background:linear-gradient(90deg,{grad})"></span> '
                f'≥ {r_red:g} %',
                unsafe_allow_html=True,
            )
        else:
            # Легенда слоя «Забитие сошников»: зелёный — без забития, дальше
            # жёлтый → красный по числу одновременно забитых сошников.
            # Ступеней 32, поэтому показываем полосу градиента с подписями
            # краёв, а не набор отдельных квадратов.
            stops = [1, 9, 17, 25, BLOCK_SCALE_MAX]
            grad = ", ".join(
                f"{block_color(n)} {100.0 * (n - 1) / (BLOCK_SCALE_MAX - 1):.0f}%"
                for n in stops
            )
            st.markdown(
                f'<span style="color:{COLOR_OK};font-weight:700">■</span> '
                f'движение без забития'
                f' &nbsp;&nbsp;<span style="font-size:.78rem;color:#666">'
                f'забито сошников:</span> '
                f'<span style="display:inline-block;width:130px;height:12px;'
                f'vertical-align:middle;border-radius:3px;'
                f'background:linear-gradient(90deg,{grad})"></span> '
                f'1 … {BLOCK_SCALE_MAX}',
                unsafe_allow_html=True,
            )


if __name__ == "__main__":
    main()