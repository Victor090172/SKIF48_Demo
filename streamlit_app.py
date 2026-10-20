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

1. Карта: трек движения объекта. Зелёный сегмент — движение без забития
   сошников; с забитыми — градиент от жёлтого (один забитый сошник) через
   оранжевый к красному (все 32). Сегмент красится по состоянию в его
   КОНЕЧНОЙ точке: забитие фиксируется в момент замера, значит участок,
   приведший к этому замеру, и есть участок с забитыми сошниками.
   Шкала — 32 сошника: флагов 221…224 в выгрузке 32 штуки, по одному на
   каждый сошник машины (см. block_color).

2. Показатели за выбранный интервал:

   Общая площадь            = Σ ΔПлощадь_га
   Общий пробег             = Σ (ΔПробег_м × 101_7_Кол-во_сошников)   [м·рядков]
   Пробег с забитием        = Σ (ΔЗабитие_на_рядок × кол-во забитых рядков)
   Доля с забитием          = Пробег с забитием / Общий пробег × 100
   Доля без забития         = 100 − Доля с забитием

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

# Три уровня показателя: (цвет рамки, заливка, подпись)
LEVEL_STYLE = {
    "хорошо": ("#1a9850", "#eaf6ee"),
    "норма": ("#e6a700", "#fdf6e3"),
    "плохо": ("#d73027", "#fbeceb"),
}

BLOCK_FLAG_PREFIXES = ("221_", "222_", "223_", "224_")

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
    return {
        "area_ha": area,
        "total_mrows": total,
        "blocked_mrows": blocked,
        "pct_blocked": pct_blocked,
        "pct_clean": 100.0 - pct_blocked,
        "n_points": int(len(d)),
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
def status_card(label: str, value: str, level: str) -> str:
    """Карточка с цветовой рамкой. Уровень не подписывается словом:
    цвет рамки и цифры говорят сами."""
    accent, tint = LEVEL_STYLE[level]
    return (
        f"<div style='flex:1 1 0;min-width:0;border:2px solid {accent};"
        f"background:{tint};border-radius:8px;padding:7px 10px'>"
        f"<div style='font-size:.7rem;color:#555;white-space:nowrap;"
        f"overflow:hidden;text-overflow:ellipsis'>{label}</div>"
        f"<div style='font-size:1.2rem;font-weight:700;color:{accent};"
        f"line-height:1.25;white-space:nowrap'>{value}</div></div>"
    )


def plain_card(label: str, value: str, unit: str = "") -> str:
    """Карточка нейтрального показателя."""
    return (
        f"<div style='flex:1 1 0;min-width:0;border:1px solid #d7dde3;"
        f"background:#fff;border-radius:8px;padding:7px 10px'>"
        f"<div style='font-size:.7rem;color:#555;white-space:nowrap;"
        f"overflow:hidden;text-overflow:ellipsis'>{label}</div>"
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
def build_map(d: pd.DataFrame, simplify_m: float) -> folium.Map:
    """Карта с треком: зелёный — без забития, дальше жёлтый → красный."""
    center = [float(d["lat"].mean()), float(d["lon"].mean())]
    fmap = folium.Map(location=center, zoom_start=14, tiles="OpenStreetMap",
                      control_scale=True)

    if len(d) < 2:
        return fmap

    lat = d["lat"].to_numpy()
    lon = d["lon"].to_numpy()
    # Число одновременно забитых сошников: пофакторные флаги 221…224 дают
    # целые 0…4. От него зависит цвет сегмента.
    n_blk = d["n_blocked"].to_numpy().astype(int)

    # Прореживание: оставляем точку, только если она ушла от последней
    # сохранённой дальше чем на simplify_m. Геометрия трека не меняется.
    keep = np.ones(len(lat), dtype=bool)
    if simplify_m > 0:
        dlat = np.diff(lat) * 111_320.0
        dlon = np.diff(lon) * 111_320.0 * np.cos(np.radians(lat[:-1]))
        keep[1:] = np.hypot(dlat, dlon) > simplify_m
        keep[-1] = True

    idx = np.flatnonzero(keep)
    slat, slon, snb = lat[idx], lon[idx], n_blk[idx]

    seg = snb[1:]  # цвет сегмента определяется его концом
    if len(seg) == 0:
        return fmap

    # Подряд идущие сегменты одного цвета склеиваем в одну ломаную:
    # ~49 тыс. точек превращаются в несколько сотен объектов карты.
    bounds = np.r_[np.flatnonzero(np.r_[True, seg[1:] != seg[:-1]]), len(seg)]
    for start, stop in zip(bounds[:-1], bounds[1:]):
        pts = list(zip(slat[start:stop + 1].tolist(), slon[start:stop + 1].tolist()))
        if len(pts) < 2:
            continue
        n = int(seg[start])
        folium.PolyLine(
            pts,
            color=COLOR_OK if n == 0 else block_color(n),
            weight=4,
            opacity=0.85,
            tooltip="без забития" if n == 0 else f"забито сошников: {n}",
        ).add_to(fmap)

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
        "Карта: зелёный — движение без забития, дальше — жёлтый → красный по числу "
        "одновременно забитых сошников. Показатели считаются только за выбранный интервал."
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

    st.caption(
        f"Выбрано: {pd.Timestamp(sel_from):%d.%m.%Y %H:%M} — "
        f"{pd.Timestamp(sel_to):%d.%m.%Y %H:%M} · "
        f"точек в срезе: {k['n_points']:,}"
    )

    st_folium.st_folium(
        build_map(d, float(simplify)),
        width=None, height=620, use_container_width=True,
    )

    if show_legend:
        # Легенда повторяет шкалу карты: зелёный — без забития, дальше
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