#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================
  Streamlit Financial Analyzer — PRO Edition
  Веб-сервис для анализа себестоимости и продаж
  Дизайн: SaaS Dashboard для eCommerce / Retail
============================================================

Запуск:
    streamlit run app.py

Требования:
    streamlit, pandas, openpyxl, plotly
============================================================
"""

import io
import re
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

# ============================================================
# 1. КОНФИГУРАЦИЯ И CSS
# ============================================================

st.set_page_config(
    page_title="Финансовый анализатор PRO",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Кастомные CSS-стили ──
CUSTOM_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .main-header {
        background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 1.5rem 2rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px rgba(59, 130, 246, 0.3);
    }
    .main-header h1 { margin: 0; font-size: 2rem; font-weight: 700; }
    .main-header p { margin: 0.5rem 0 0; opacity: 0.9; font-size: 1rem; }

    .kpi-card {
        background: white;
        border-radius: 16px;
        padding: 1.2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        border-left: 5px solid;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .kpi-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.12);
    }
    .kpi-blue { border-left-color: #3b82f6; }
    .kpi-amber { border-left-color: #f59e0b; }
    .kpi-green { border-left-color: #10b981; }
    .kpi-red { border-left-color: #ef4444; }
    .kpi-purple { border-left-color: #8b5cf6; }

    .filter-panel {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
    }

    .section-title {
        font-size: 1.4rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 3px solid #3b82f6;
        display: inline-block;
    }

    .styled-table {
        border-collapse: separate;
        border-spacing: 0;
        width: 100%;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        font-size: 0.9rem;
    }
    .styled-table th {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white;
        padding: 14px;
        text-align: left;
        font-weight: 600;
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .styled-table td { padding: 12px 14px; border-bottom: 1px solid #e2e8f0; }
    .styled-table tr:hover { background-color: #f8fafc; }
    .row-high { background-color: #f0fdf4 !important; }
    .row-medium { background-color: #fefce8 !important; }
    .row-low { background-color: #fff7ed !important; }
    .row-loss { background-color: #fef2f2 !important; }

    .status-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    .badge-high { background: #d1fae5; color: #065f46; }
    .badge-medium { background: #fef3c7; color: #92400e; }
    .badge-low { background: #ffedd5; color: #7c2d12; }
    .badge-loss { background: #fee2e2; color: #991b1b; }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ============================================================
# 2. КОНСТАНТЫ И СПРАВОЧНИКИ
# ============================================================

COLUMN_SYNONYMS = {
    "id": [
        "артикул", "id", "код", "sku", "штрихкод", "баркод", "barcode",
        "идентификатор", "номер", "№", "код товара", "артикул товара",
        "sku код", "код sku", "внутренний код", "external_id", "product_id",
        "item_id", "goods_id", "номенклатура", "код номенклатуры",
        "article", "sku code", "product code", "item code", "item number",
        "product id", "sku id", "stock keeping unit", "product number",
        "item id", "goods code", "goods id", "variant sku", "master sku",
    ],
    "name": [
        "наименование", "название", "имя", "товар", "продукт", "описание",
        "наименование товара", "название товара", "продукт название",
        "name", "product name", "item name", "goods name", "title",
        "description", "product title", "item title", "product",
    ],
    "cost_per_unit": [
        "себестоимость", "себестоимость единицы", "цена закупки",
        "закупочная цена", "цена входа", "стоимость единицы",
        "себестоимость товара", "цена себестоимости", "unit cost",
        "cost", "purchase price", "buying price", "cost price",
        "cost per unit", "unit cost price", "cost of goods",
    ],
    "quantity_sold": [
        "продано", "продано шт", "количество", "количество продаж",
        "шт", "штук", "units sold", "quantity", "sales quantity",
        "sold", "sales volume", "volume", "qty", "qty sold",
        "кол-во", "количество шт", "продажи", "quantity sold",
        "sales count", "units", "pieces sold", "total sold",
    ],
    "revenue": [
        "выручка", "доход", "оборот", "сумма продаж", "продажи",
        "выручка от продаж", "общая выручка", "revenue", "total revenue",
        "sales revenue", "gross revenue", "turnover", "income",
        "total sales", "sales amount", "gross sales", "продажи сумма",
    ],
}

STATUS_EMOJI = {
    "Высокомаржинальный": "🟢",
    "Среднемаржинальный": "🟡",
    "Низкомаржинальный": "🟠",
    "Убыточный": "🔴",
}

STATUS_CLASS = {
    "Высокомаржинальный": "badge-high",
    "Среднемаржинальный": "badge-medium",
    "Низкомаржинальный": "badge-low",
    "Убыточный": "badge-loss",
}

STATUS_ROW_CLASS = {
    "Высокомаржинальный": "row-high",
    "Среднемаржинальный": "row-medium",
    "Низкомаржинальный": "row-low",
    "Убыточный": "row-loss",
}


# ============================================================
# 3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def normalize_text(text: str) -> str:
    """Нормализация текста для интеллектуального сопоставления."""
    if pd.isna(text):
        return ""
    text = str(text).lower().strip()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^\w\s]", "", text)
    return text


def find_column(df: pd.DataFrame, semantic_name: str) -> Optional[str]:
    """Ищет колонку по семантическому имени."""
    candidates = COLUMN_SYNONYMS.get(semantic_name, [])
    if not candidates:
        return None

    normalized_cols = {col: normalize_text(col) for col in df.columns}

    for col, norm_col in normalized_cols.items():
        for cand in candidates:
            if norm_col == normalize_text(cand):
                return col

    for col, norm_col in normalized_cols.items():
        for cand in candidates:
            norm_cand = normalize_text(cand)
            if norm_cand in norm_col or norm_col in norm_cand:
                return col

    for col, norm_col in normalized_cols.items():
        for cand in candidates:
            norm_cand = normalize_text(cand)
            if len(norm_cand) >= 3 and norm_col.startswith(norm_cand[:3]):
                return col

    return None


def auto_map_columns(df_cost: pd.DataFrame, df_sales: pd.DataFrame) -> dict:
    """Автоматическое сопоставление колонок."""
    mapping = {}
    mapping["cost_id"] = find_column(df_cost, "id")
    mapping["cost_name"] = find_column(df_cost, "name")
    mapping["cost_per_unit"] = find_column(df_cost, "cost_per_unit")
    mapping["sales_id"] = find_column(df_sales, "id")
    mapping["sales_quantity"] = find_column(df_sales, "quantity_sold")
    mapping["sales_revenue"] = find_column(df_sales, "revenue")
    return mapping


def format_ru_number(x) -> str:
    """Форматирование числа с пробелом как разделителем тысяч."""
    if pd.isna(x):
        return ""
    sign = "-" if x < 0 else ""
    s = f"{abs(x):,.2f}"
    s = s.replace(",", " ")
    return f"{sign}{s}"


def render_kpi_card(icon: str, title: str, value: str, delta: str, color: str) -> str:
    """HTML-шаблон KPI-карточки."""
    return f"""
    <div class="kpi-card kpi-{color}">
        <div style="font-size: 2rem; margin-bottom: 6px;">{icon}</div>
        <div style="font-size: 0.75rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">
            {title}
        </div>
        <div style="font-size: 1.5rem; font-weight: 700; color: #0f172a; margin-bottom: 4px;">
            {value}
        </div>
        <div style="font-size: 0.8rem; color: {color if color != 'green' else '#10b981'}; font-weight: 500;">
            {delta}
        </div>
    </div>
    """


def highlight_margin(val):
    """Градиентная подсветка ячейки Маржинальность %."""
    if pd.isna(val):
        return ''
    if val < 0:
        return 'background-color: #fee2e2; color: #991b1b; font-weight: 600;'
    elif val < 10:
        return 'background-color: #fef3c7; color: #92400e;'
    elif val < 30:
        return 'background-color: #dbeafe; color: #1e40af;'
    else:
        return 'background-color: #d1fae5; color: #065f46; font-weight: 600;'


def color_rows(row):
    """Цветовое кодирование строки по статусу."""
    status = row.get('Статус', '')
    if 'Высокомаржинальный' in status:
        return ['background-color: #f0fdf4'] * len(row)
    elif 'Среднемаржинальный' in status:
        return ['background-color: #fefce8'] * len(row)
    elif 'Низкомаржинальный' in status:
        return ['background-color: #fff7ed'] * len(row)
    elif 'Убыточный' in status:
        return ['background-color: #fef2f2'] * len(row)
    return [''] * len(row)


def build_result_excel(df: pd.DataFrame) -> bytes:
    """Создаёт форматированный Excel-файл."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Результат")
        workbook = writer.book
        worksheet = writer.sheets["Результат"]

        header_fill = PatternFill(start_color="1e3a8a", end_color="1e3a8a", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
            cell.border = thin_border

        for column in worksheet.columns:
            max_length = 0
            col_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                except Exception:
                    pass
            adjusted_width = min(max_length + 4, 50)
            worksheet.column_dimensions[col_letter].width = adjusted_width

        worksheet.freeze_panes = "A2"

    output.seek(0)
    return output.getvalue()


# ============================================================
# 4. SIDEBAR — ЗАГРУЗКА И НАСТРОЙКИ
# ============================================================

with st.sidebar:
    st.markdown("<h2 style='text-align:center; color:#3b82f6;'>📂 Загрузка данных</h2>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("**1. Файл себестоимости**")
    st.caption("Колонки: Артикул/ID, Наименование, Себестоимость ед.")
    cost_file = st.file_uploader(
        "Загрузите файл себестоимости",
        type=["xlsx", "xls", "csv"],
        key="cost_uploader",
        help="Excel (.xlsx, .xls) или CSV",
    )

    st.markdown("**2. Файл продаж**")
    st.caption("Колонки: Артикул/ID, Продано шт, Выручка")
    sales_file = st.file_uploader(
        "Загрузите файл продаж",
        type=["xlsx", "xls", "csv"],
        key="sales_uploader",
        help="Excel (.xlsx, .xls) или CSV",
    )

    st.markdown("---")
    st.markdown("**3. Параметры анализа**")
    high_margin = st.number_input("Высокая маржинальность, %", min_value=0.0, max_value=100.0, value=30.0, step=1.0)
    medium_margin = st.number_input("Средняя маржинальность, %", min_value=0.0, max_value=100.0, value=10.0, step=1.0)
    st.caption("Ниже средней → Низкомаржинальный, <0% → Убыточный")

    st.markdown("---")
    st.info(
        "💡 **Подсказка:** Приложение автоматически распознаёт колонки "
        "по их названиям (Артикул, SKU, ID, Код и т.д.). "
        "Если автоматическое сопоставление не сработает, вы сможете выбрать колонки вручную."
    )


# ============================================================
# 5. ОСНОВНОЙ КОНТЕНТ — ЗАГОЛОВОК
# ============================================================

st.markdown("""
<div class="main-header">
    <h1>📊 Финансовый анализатор PRO</h1>
    <p>Анализ себестоимости, продаж и маржинальности для крупного ассортимента</p>
</div>
""", unsafe_allow_html=True)


# ============================================================
# 6. ЧТЕНИЕ ФАЙЛОВ И ПРЕВЬЮ
# ============================================================

if cost_file is not None and sales_file is not None:

    try:
        if cost_file.name.endswith(".csv"):
            df_cost = pd.read_csv(cost_file, encoding="utf-8")
        else:
            df_cost = pd.read_excel(cost_file, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Ошибка при чтении файла себестоимости: {e}")
        st.stop()

    try:
        if sales_file.name.endswith(".csv"):
            df_sales = pd.read_csv(sales_file, encoding="utf-8")
        else:
            df_sales = pd.read_excel(sales_file, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Ошибка при чтении файла продаж: {e}")
        st.stop()

    # Превью загруженных данных
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("<div class='section-title'>📄 Файл себестоимости</div>", unsafe_allow_html=True)
        st.caption(f"Строк: {len(df_cost):,} | Колонок: {len(df_cost.columns)}".replace(",", " "))
        st.dataframe(df_cost.head(5), use_container_width=True, height=220)
    with col2:
        st.markdown("<div class='section-title'>📄 Файл продаж</div>", unsafe_allow_html=True)
        st.caption(f"Строк: {len(df_sales):,} | Колонок: {len(df_sales.columns)}".replace(",", " "))
        st.dataframe(df_sales.head(5), use_container_width=True, height=220)

    st.markdown("---")

    # ──────────────────────────────────────────────────────
    # 7. СОПОСТАВЛЕНИЕ КОЛОНОК
    # ──────────────────────────────────────────────────────

    auto_map = auto_map_columns(df_cost, df_sales)

    st.markdown("<div class='section-title'>🔗 Сопоставление колонок</div>", unsafe_allow_html=True)
    st.markdown("Приложение автоматически определило соответствие. При необходимости откорректируйте выбор вручную.")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Файл себестоимости:**")
        cost_id_col = st.selectbox(
            "Колонка ID / Артикул",
            options=df_cost.columns.tolist(),
            index=(df_cost.columns.tolist().index(auto_map["cost_id"]) if auto_map["cost_id"] and auto_map["cost_id"] in df_cost.columns else 0),
            help="Колонка для объединения",
        )
        cost_name_col = st.selectbox(
            "Колонка Наименование",
            options=df_cost.columns.tolist(),
            index=(df_cost.columns.tolist().index(auto_map["cost_name"]) if auto_map["cost_name"] and auto_map["cost_name"] in df_cost.columns else 0),
            help="Наименование товара",
        )
        cost_unit_col = st.selectbox(
            "Колонка Себестоимость ед.",
            options=df_cost.columns.tolist(),
            index=(df_cost.columns.tolist().index(auto_map["cost_per_unit"]) if auto_map["cost_per_unit"] and auto_map["cost_per_unit"] in df_cost.columns else 0),
            help="Себестоимость одной единицы",
        )

    with c2:
        st.markdown("**Файл продаж:**")
        sales_id_col = st.selectbox(
            "Колонка ID / Артикул",
            options=df_sales.columns.tolist(),
            index=(df_sales.columns.tolist().index(auto_map["sales_id"]) if auto_map["sales_id"] and auto_map["sales_id"] in df_sales.columns else 0),
            help="Колонка для объединения",
        )
        sales_qty_col = st.selectbox(
            "Колонка Продано шт.",
            options=df_sales.columns.tolist(),
            index=(df_sales.columns.tolist().index(auto_map["sales_quantity"]) if auto_map["sales_quantity"] and auto_map["sales_quantity"] in df_sales.columns else 0),
            help="Количество проданных единиц",
        )
        sales_rev_col = st.selectbox(
            "Колонка Выручка",
            options=df_sales.columns.tolist(),
            index=(df_sales.columns.tolist().index(auto_map["sales_revenue"]) if auto_map["sales_revenue"] and auto_map["sales_revenue"] in df_sales.columns else 0),
            help="Общая выручка",
        )

    with c3:
        st.markdown("&nbsp;")
        st.markdown("&nbsp;")
        st.markdown("&nbsp;")
        st.markdown("&nbsp;")
        st.markdown("&nbsp;")
        st.markdown("&nbsp;")

    # ──────────────────────────────────────────────────────
    # 8. ОБРАБОТКА ДАННЫХ
    # ──────────────────────────────────────────────────────

    if st.button("🚀 Запустить анализ", type="primary", use_container_width=True):
        with st.spinner("Обработка данных..."):

            dfc = df_cost[[cost_id_col, cost_name_col, cost_unit_col]].copy()
            dfs = df_sales[[sales_id_col, sales_qty_col, sales_rev_col]].copy()

            dfc.rename(columns={cost_id_col: "ID", cost_name_col: "Наименование", cost_unit_col: "Себестоимость_ед"}, inplace=True)
            dfs.rename(columns={sales_id_col: "ID", sales_qty_col: "Продано_шт", sales_rev_col: "Выручка"}, inplace=True)

            dfc["ID"] = dfc["ID"].astype(str).str.strip().str.upper()
            dfs["ID"] = dfs["ID"].astype(str).str.strip().str.upper()

            for col in ["Себестоимость_ед", "Продано_шт", "Выручка"]:
                if col in dfc.columns:
                    dfc[col] = pd.to_numeric(dfc[col].astype(str).str.replace(" ", "").str.replace(",", "."), errors="coerce")
                if col in dfs.columns:
                    dfs[col] = pd.to_numeric(dfs[col].astype(str).str.replace(" ", "").str.replace(",", "."), errors="coerce")

            dfc = dfc.drop_duplicates(subset=["ID"], keep="first")
            dfs = dfs.groupby("ID", as_index=False).agg({"Продано_шт": "sum", "Выручка": "sum"})

            merged = dfs.merge(dfc, on="ID", how="left")
            merged["Наименование"] = merged["Наименование"].fillna(merged["ID"])
            merged["Себестоимость_ед"] = merged["Себестоимость_ед"].fillna(0)

            merged["Общая_себестоимость"] = merged["Продано_шт"] * merged["Себестоимость_ед"]
            merged["Валовая_прибыль"] = merged["Выручка"] - merged["Общая_себестоимость"]
            merged["Маржинальность_%"] = merged.apply(
                lambda row: (row["Валовая_прибыль"] / row["Выручка"] * 100) if row["Выручка"] != 0 and pd.notna(row["Выручка"]) else 0, axis=1
            )
            merged["ROI_%"] = merged.apply(
                lambda row: (row["Валовая_прибыль"] / row["Общая_себестоимость"] * 100)
                if row["Общая_себестоимость"] != 0 and pd.notna(row["Общая_себестоимость"])
                else (float("inf") if row["Валовая_прибыль"] > 0 else 0),
                axis=1,
            )

            def classify_margin(margin: float) -> str:
                if margin > high_margin:
                    return "Высокомаржинальный"
                elif margin >= medium_margin:
                    return "Среднемаржинальный"
                elif margin >= 0:
                    return "Низкомаржинальный"
                else:
                    return "Убыточный"

            merged["Статус"] = merged["Маржинальность_%"].apply(classify_margin)
            merged["Маржинальность_%"] = merged["Маржинальность_%"].round(2)
            merged["ROI_%"] = merged["ROI_%"].round(2)
            merged["Валовая_прибыль"] = merged["Валовая_прибыль"].round(2)
            merged["Общая_себестоимость"] = merged["Общая_себестоимость"].round(2)
            merged["Себестоимость_ед"] = merged["Себестоимость_ед"].round(2)
            merged["Выручка"] = merged["Выручка"].round(2)
            merged = merged.sort_values(by="Валовая_прибыль", ascending=False).reset_index(drop=True)

            # ───────────────────────────────────────────────
            # 9. KPI КАРТОЧКИ (4 шт)
            # ───────────────────────────────────────────────
            st.markdown("---")
            st.markdown("<div class='section-title'>📈 Ключевые показатели бизнеса</div>", unsafe_allow_html=True)

            total_revenue = merged["Выручка"].sum()
            total_cost = merged["Общая_себестоимость"].sum()
            total_profit = merged["Валовая_прибыль"].sum()
            overall_margin = (total_profit / total_revenue * 100) if total_revenue else 0
            overall_roi = (total_profit / total_cost * 100) if total_cost else 0
            profitable_count = (merged["Валовая_прибыль"] > 0).sum()
            loss_count = (merged["Валовая_прибыль"] < 0).sum()
            total_items = len(merged)

            profit_color = "green" if total_profit >= 0 else "red"
            profit_delta = "Прибыль положительная ✅" if total_profit >= 0 else "Убыток ❌"

            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(render_kpi_card("💰", "Общая выручка", format_ru_number(total_revenue), f"{total_items} товаров", "blue"), unsafe_allow_html=True)
            with k2:
                st.markdown(render_kpi_card("📉", "Общая себестоимость", format_ru_number(total_cost), f"{profitable_count} прибыльных", "amber"), unsafe_allow_html=True)
            with k3:
                st.markdown(render_kpi_card("📈", "Чистая прибыль", format_ru_number(total_profit), profit_delta, profit_color), unsafe_allow_html=True)
            with k4:
                st.markdown(render_kpi_card("📊", "Средняя маржинальность", f"{overall_margin:.2f}%", f"ROI: {overall_roi:.1f}%", "purple"), unsafe_allow_html=True)

            # ───────────────────────────────────────────────
            # 10. РАСПРЕДЕЛЕНИЕ ПО СТАТУСАМ
            # ───────────────────────────────────────────────
            st.markdown("---")
            status_counts = merged["Статус"].value_counts().reset_index()
            status_counts.columns = ["Статус", "Количество"]

            chart_col1, chart_col2 = st.columns([1, 2])
            with chart_col1:
                st.markdown("<div class='section-title'>📌 Распределение по статусам</div>", unsafe_allow_html=True)
                # Добавляем эмодзи
                status_counts["Статус_экран"] = status_counts["Статус"].map(lambda s: f"{STATUS_EMOJI.get(s, '')} {s}")
                st.dataframe(status_counts[["Статус_экран", "Количество"]].rename(columns={"Статус_экран": "Статус"}), use_container_width=True, hide_index=True)

            with chart_col2:
                try:
                    import plotly.express as px
                    fig_pie = px.pie(
                        status_counts,
                        values="Количество",
                        names="Статус",
                        title="Доля товаров по маржинальности",
                        color="Статус",
                        color_discrete_map={
                            "Высокомаржинальный": "#2ecc71",
                            "Среднемаржинальный": "#f1c40f",
                            "Низкомаржинальный": "#e67e22",
                            "Убыточный": "#e74c3c",
                        },
                    )
                    fig_pie.update_layout(showlegend=True, height=350, title_font_size=16)
                    st.plotly_chart(fig_pie, use_container_width=True)
                except ImportError:
                    st.info("Установите `plotly` для интерактивных графиков: `pip install plotly`")

            # ───────────────────────────────────────────────
            # 11. ТОП-10
            # ───────────────────────────────────────────────
            st.markdown("---")
            top_col1, top_col2 = st.columns(2)
            with top_col1:
                st.markdown("<div class='section-title'>🏆 Топ-10 по прибыли</div>", unsafe_allow_html=True)
                top_profit = merged.nlargest(10, "Валовая_прибыль")[["ID", "Наименование", "Валовая_прибыль", "Маржинальность_%", "Статус"]].copy()
                top_profit["Статус"] = top_profit["Статус"].map(lambda s: f"{STATUS_EMOJI.get(s, '')} {s}")
                st.dataframe(top_profit, use_container_width=True, hide_index=True)

            with top_col2:
                st.markdown("<div class='section-title'>⚠️ Топ-10 убыточных</div>", unsafe_allow_html=True)
                top_loss = merged.nsmallest(10, "Валовая_прибыль")[["ID", "Наименование", "Валовая_прибыль", "Маржинальность_%", "Статус"]].copy()
                top_loss["Статус"] = top_loss["Статус"].map(lambda s: f"{STATUS_EMOJI.get(s, '')} {s}")
                st.dataframe(top_loss, use_container_width=True, hide_index=True)

            # ───────────────────────────────────────────────
            # 12. ФИЛЬТРЫ
            # ───────────────────────────────────────────────
            st.markdown("---")

            # Подготовка данных для отображения
            display_df = merged.copy()
            display_df.rename(
                columns={
                    "Себестоимость_ед": "Себестоимость ед.",
                    "Продано_шт": "Продано шт.",
                    "Общая_себестоимость": "Общая себестоимость",
                    "Валовая_прибыль": "Валовая прибыль",
                    "Маржинальность_%": "Маржинальность %",
                    "ROI_%": "ROI %",
                },
                inplace=True,
            )
            final_cols = [
                "ID", "Наименование", "Себестоимость ед.", "Продано шт.",
                "Выручка", "Общая себестоимость", "Валовая прибыль",
                "Маржинальность %", "ROI %", "Статус",
            ]
            display_df = display_df[[c for c in final_cols if c in display_df.columns]]

            # Добавляем эмодзи к статусам для отображения
            all_statuses = display_df["Статус"].unique().tolist()
            min_rev = float(display_df["Выручка"].min())
            max_rev = float(display_df["Выручка"].max())

            with st.expander("🔍 Фильтры и Поиск данных", expanded=True):
                st.markdown('<div class="filter-panel">', unsafe_allow_html=True)
                f1, f2, f3 = st.columns(3)
                with f1:
                    search_query = st.text_input("🔎 Поиск по Артикулу или Наименованию", placeholder="Введите текст...")
                with f2:
                    status_filter = st.multiselect("📌 Фильтр по статусу", options=all_statuses, default=all_statuses)
                with f3:
                    rev_min, rev_max = st.slider(
                        "💵 Диапазон выручки",
                        min_value=min_rev,
                        max_value=max_rev,
                        value=(min_rev, max_rev),
                        step=1000.0,
                        format="%.0f",
                    )
                st.markdown('</div>', unsafe_allow_html=True)

            # Применение фильтров
            filtered_df = display_df.copy()
            if search_query:
                mask = (
                    filtered_df["ID"].astype(str).str.contains(search_query, case=False, na=False) |
                    filtered_df["Наименование"].astype(str).str.contains(search_query, case=False, na=False)
                )
                filtered_df = filtered_df[mask]
            if status_filter:
                filtered_df = filtered_df[filtered_df["Статус"].isin(status_filter)]
            filtered_df = filtered_df[(filtered_df["Выручка"] >= rev_min) & (filtered_df["Выручка"] <= rev_max)]

            st.caption(f"📋 Показано {len(filtered_df)} из {len(display_df)} записей")

            # ───────────────────────────────────────────────
            # 13. ИНТЕРАКТИВНАЯ ТАБЛИЦА (st.dataframe)
            # ───────────────────────────────────────────────
            st.markdown("<div class='section-title'>📋 Интерактивная таблица результатов</div>", unsafe_allow_html=True)

            # Для st.dataframe оставляем числа числами (сортировка работает)
            interactive_df = filtered_df.copy()
            interactive_df["Статус"] = interactive_df["Статус"].map(lambda s: f"{STATUS_EMOJI.get(s, '')} {s}")

            st.dataframe(
                interactive_df,
                use_container_width=True,
                height=500,
                column_config={
                    "Маржинальность %": st.column_config.NumberColumn(
                        "Маржинальность %", format="%.2f %%", help="Валовая прибыль / Выручка × 100"
                    ),
                    "ROI %": st.column_config.NumberColumn(
                        "ROI %", format="%.2f %%", help="Валовая прибыль / Себестоимость × 100"
                    ),
                    "Валовая прибыль": st.column_config.NumberColumn(
                        "Валовая прибыль", format="%.2f"
                    ),
                    "Выручка": st.column_config.NumberColumn(
                        "Выручка", format="%.2f"
                    ),
                    "Общая себестоимость": st.column_config.NumberColumn(
                        "Общая себестоимость", format="%.2f"
                    ),
                    "Себестоимость ед.": st.column_config.NumberColumn(
                        "Себестоимость ед.", format="%.2f"
                    ),
                    "Статус": st.column_config.TextColumn(
                        "Статус", help="Классификация по маржинальности"
                    ),
                },
                hide_index=True,
            )

            # ───────────────────────────────────────────────
            # 14. СТИЛИЗОВАННАЯ HTML ТАБЛИЦА (топ-50, с градиентом)
            # ───────────────────────────────────────────────
            st.markdown("<div class='section-title'>🎨 Стильная таблица (Топ-50 по прибыли)</div>", unsafe_allow_html=True)
            st.caption("Цветовое кодирование строк: 🟢 высокомаржинальные, 🟡 средние, 🟠 низкие, 🔴 убыточные")

            # Берём топ-50 для HTML-отображения
            html_df = display_df.head(50).copy()
            html_df["Статус_отображение"] = html_df["Статус"].map(
                lambda s: f'<span class="status-badge {STATUS_CLASS.get(s, '')}">{STATUS_EMOJI.get(s, "")} {s}</span>'
            )

            # Форматируем числа для HTML
            num_cols = ["Себестоимость ед.", "Продано шт.", "Выручка", "Общая себестоимость", "Валовая прибыль", "Маржинальность %", "ROI %"]
            for col in num_cols:
                if col in html_df.columns:
                    html_df[col] = html_df[col].apply(format_ru_number)

            # Создаём HTML с помощью pandas Styler
            style_df = display_df.head(50).copy()
            style_df["Статус"] = style_df["Статус"].map(lambda s: f"{STATUS_EMOJI.get(s, '')} {s}")

            styler = style_df.style\
                .apply(color_rows, axis=1)\
                .map(highlight_margin, subset=["Маржинальность %"])\
                .format(lambda x: format_ru_number(x) if pd.notna(x) else "", subset=num_cols)

            st.markdown(styler.to_html(table_attributes='class="styled-table"', index=False), unsafe_allow_html=True)

            # ───────────────────────────────────────────────
            # 15. ГРАФИКИ PLOTLY (ТОП-20)
            # ───────────────────────────────────────────────
            st.markdown("---")
            try:
                import plotly.express as px
                st.markdown("<div class='section-title'>📊 Визуализация (Топ-20 товаров)</div>", unsafe_allow_html=True)

                top_20 = merged.head(20).copy()
                top_20["Статус_иконка"] = top_20["Статус"].map(lambda s: f"{STATUS_EMOJI.get(s, '')} {s}")

                fig_bar = px.bar(
                    top_20,
                    x="Наименование",
                    y=["Выручка", "Общая_себестоимость", "Валовая_прибыль"],
                    title="Выручка / Себестоимость / Прибыль",
                    barmode="group",
                    labels={"value": "Сумма (руб.)", "variable": "Показатель"},
                    height=550,
                )
                fig_bar.update_layout(
                    xaxis_tickangle=-45,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    margin=dict(l=20, r=20, t=80, b=120),
                )
                st.plotly_chart(fig_bar, use_container_width=True)

                fig_scatter = px.scatter(
                    merged,
                    x="Маржинальность_%",
                    y="ROI_%",
                    color="Статус",
                    size="Выручка",
                    hover_data=["ID", "Наименование"],
                    title="Маржинальность vs ROI (размер = Выручка)",
                    color_discrete_map={
                        "Высокомаржинальный": "#2ecc71",
                        "Среднемаржинальный": "#f1c40f",
                        "Низкомаржинальный": "#e67e22",
                        "Убыточный": "#e74c3c",
                    },
                    height=500,
                )
                fig_scatter.update_layout(
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    margin=dict(l=20, r=20, t=80, b=60),
                )
                st.plotly_chart(fig_scatter, use_container_width=True)
            except ImportError:
                st.info("Установите `plotly` для интерактивных графиков: `pip install plotly`")

            # ───────────────────────────────────────────────
            # 16. СКАЧИВАНИЕ
            # ───────────────────────────────────────────────
            st.markdown("---")
            excel_bytes = build_result_excel(display_df)
            st.download_button(
                label="📥 Скачать полный отчёт в Excel",
                data=excel_bytes,
                file_name=f"financial_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

            # ───────────────────────────────────────────────
            # 17. СВОДКА ПО КАТЕГОРИЯМ
            # ───────────────────────────────────────────────
            st.markdown("---")
            st.markdown("<div class='section-title'>📝 Сводка по категориям</div>", unsafe_allow_html=True)
            summary = (
                merged.groupby("Статус")
                .agg(
                    Количество=("ID", "count"),
                    Выручка=("Выручка", "sum"),
                    Себестоимость=("Общая_себестоимость", "sum"),
                    Прибыль=("Валовая_прибыль", "sum"),
                    Средняя_маржа=("Маржинальность_%", "mean"),
                )
                .round(2)
                .reset_index()
            )
            summary["Статус"] = summary["Статус"].map(lambda s: f"{STATUS_EMOJI.get(s, '')} {s}")
            st.dataframe(summary, use_container_width=True, hide_index=True)

            # ───────────────────────────────────────────────
            # 18. НЕСОПОСТАВЛЕННЫЕ ПОЗИЦИИ
            # ───────────────────────────────────────────────
            unmatched = merged[merged["Себестоимость_ед"] == 0]
            if not unmatched.empty:
                st.warning(
                    f"⚠️ **Внимание:** {len(unmatched)} позиций из файла продаж "
                    f"не удалось сопоставить с файлом себестоимости (себестоимость = 0). "
                    f"Проверьте корректность колонки ID/Артикул."
                )
                with st.expander("Показать несопоставленные позиции"):
                    st.dataframe(unmatched[["ID", "Наименование", "Продано_шт", "Выручка"]], use_container_width=True)

else:
    # ──────────────────────────────────────────────────────
    # СТАРТОВЫЙ ЭКРАН
    # ──────────────────────────────────────────────────────
    st.info("⬅️ **Загрузите оба файла в боковой панели**, чтобы начать анализ.")

    st.markdown("""
    ### Как это работает?

    1. **Загрузите файл себестоимости** — должен содержать колонки:
       - `Артикул` / `ID` / `SKU` / `Код` / `Баркод`
       - `Наименование` / `Название` / `Имя` / `Товар`
       - `Себестоимость` / `Цена закупки` / `Cost` / `Unit Cost`

    2. **Загрузите файл продаж** — должен содержать колонки:
       - `Артикул` / `ID` / `SKU` / `Код`
       - `Продано` / `Количество` / `Units Sold` / `Qty`
       - `Выручка` / `Доход` / `Revenue` / `Sales`

    3. **Приложение автоматически** распознаёт колонки, объединяет данные, рассчитывает метрики и выводит отчёт.

    ### Расчитываемые метрики
    | Метрика | Формула |
    |---------|---------|
    | Общая себестоимость | `Продано шт × Себестоимость ед.` |
    | Валовая прибыль | `Выручка − Общая себестоимость` |
    | Маржинальность | `(Прибыль / Выручка) × 100%` |
    | ROI | `(Прибыль / Себестоимость) × 100%` |

    ### Статусы товаров
    - 🟢 **Высокомаржинальный** — > 30%
    - 🟡 **Среднемаржинальный** — 10–30%
    - 🟠 **Низкомаржинальный** — 0–10%
    - 🔴 **Убыточный** — < 0%
    """)

    st.markdown("---")
    st.caption(
        "Создано с использованием **Streamlit**, **Pandas** и **OpenPyXL**. "
        "Для визуализаций используется **Plotly** (опционально)."
    )
