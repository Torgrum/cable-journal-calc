import streamlit as st
import pandas as pd
import io
import json
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

st.set_page_config(page_title="Универсальный КЖ", page_icon="📋", layout="wide")
st.title("📋 Универсальный калькулятор кабельного журнала")

# --- Инициализация состояния ---
if "df" not in st.session_state:
    st.session_state.df = None
if "template" not in st.session_state:
    st.session_state.template = None
if "config" not in st.session_state:
    st.session_state.config = {}  # sheet, header_row, columns_map


# --- Функция: автопоиск строки заголовков ---
def find_header_row(ws, max_search=20):
    """Ищет первую строку, где подряд заполнено ≥3 ячеек (вероятно, заголовки)."""
    for row in range(1, max_search + 1):
        filled = 0
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value and not isinstance(cell, MergedCell):
                filled += 1
            else:
                filled = 0
            if filled >= 3:
                return row
    return 1


# --- Функция: извлечение заголовков и структуры ---
def parse_template(file_bytes, sheet_name=None, header_row=None):
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name] if sheet_name else wb.active
    
    if header_row is None:
        header_row = find_header_row(ws)
    
    # Читаем заголовки
    columns = {}
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=col)
        if cell.value and not isinstance(cell, MergedCell):
            columns[str(cell.value).strip()] = col
    
    # Определяем типы по первой строке данных
    types = {}
    for col_name, col_idx in columns.items():
        sample = ws.cell(row=header_row + 1, column=col_idx).value
        if isinstance(sample, (int, float)):
            types[col_name] = "number"
        else:
            types[col_name] = "text"
    
    return {
        "wb": wb,
        "ws_name": ws.title,
        "header_row": header_row,
        "columns": columns,
        "types": types,
    }


# --- Боковая панель: загрузка и настройки ---
with st.sidebar:
    st.header("⚙️ Настройки")
    
    uploaded = st.file_uploader("📄 Загрузить шаблон .xlsx", type=["xlsx"])
    
    if uploaded is not None:
        st.session_state.template = uploaded.getvalue()
        
        # Получаем список листов
        wb_tmp = load_workbook(io.BytesIO(uploaded.getvalue()), read_only=True)
        sheet_names = wb_tmp.sheetnames
        wb_tmp.close()
        
        selected_sheet = st.selectbox("Лист", sheet_names)
        header_row = st.number_input(
            "Строка заголовков",
            min_value=1, max_value=50, value=1,
            help="Номер строки, где находятся заголовки столбцов"
        )
        
        if st.button("🔄 Применить и распознать"):
            try:
                cfg = parse_template(
                    st.session_state.template,
                    sheet_name=selected_sheet,
                    header_row=int(header_row),
                )
                st.session_state.config = cfg
                # Создаём пустой DataFrame с нужными колонками
                st.session_state.df = pd.DataFrame(columns=list(cfg["columns"].keys()))
                st.success(f"✅ Найдено {len(cfg['columns'])} столбцов")
                st.rerun()
            except Exception as e:
                st.error(f"Ошибка разбора: {e}")
    
    st.divider()
    st.caption("💡 Шаблон может содержать формулы, рамки, стили — всё сохранится")


# --- Основная область ---
if st.session_state.config:
    cfg = st.session_state.config
    
    # --- Информация о шаблоне ---
    with st.expander("📐 Структура шаблона", expanded=False):
        st.write(f"**Лист:** `{cfg['ws_name']}`")
        st.write(f"**Строка заголовков:** {cfg['header_row']}")
        st.write(f"**Найдено столбцов:** {len(cfg['columns'])}")
        
        cols_preview = pd.DataFrame([
            {"Столбец": name, "Колонка Excel": get_column_letter(idx), "Тип": cfg["types"][name]}
            for name, idx in cfg["columns"].items()
        ])
        st.dataframe(cols_preview, use_container_width=True, hide_index=True)
    
    # --- Редактор таблицы ---
    st.subheader("📝 Заполнение данных")
    
    # Динамическая конфигурация столбцов
    column_config = {}
    for col_name in cfg["columns"].keys():
        if cfg["types"].get(col_name) == "number":
            column_config[col_name] = st.column_config.NumberColumn(col_name, step=0.5)
        else:
            column_config[col_name] = st.column_config.TextColumn(col_name)
    
    df = st.data_editor(
        st.session_state.df if st.session_state.df is not None else pd.DataFrame(columns=list(cfg["columns"].keys())),
        column_config=column_config,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
    )
    st.session_state.df = df
    
    # --- Сводка ---
    if df is not None and not df.empty:
        st.subheader("📊 Сводка")
        c1, c2 = st.columns(2)
        c1.metric("Строк заполнено", len(df))
        # Ищем числовые столбцы для суммы
        numeric_cols = [c for c in df.columns if df[c].dtype in ['float64', 'int64']]
        if numeric_cols:
            total = df[numeric_cols].sum().sum()
            c2.metric("Сумма по числам", f"{total:.2f}")
    
    # --- Экспорт ---
    st.divider()
    if st.session_state.df is not None and not st.session_state.df.empty:
        if st.button("📥 Сформировать Excel по шаблону", type="primary"):
            try:
                output = io.BytesIO()
                wb = load_workbook(io.BytesIO(st.session_state.template))
                ws = wb[cfg["ws_name"]]
                
                start_row = cfg["header_row"] + 1
                col_map = cfg["columns"]
                
                # Очистка старых данных (только значения, не формулы в шаблоне)
                # и запись новых
                for i, (_, row) in enumerate(st.session_state.df.iterrows()):
                    excel_row = start_row + i
                    for col_name, value in row.items():
                        if col_name not in col_map:
                            continue
                        excel_col = col_map[col_name]
                        cell = ws.cell(row=excel_row, column=excel_col)
                        
                        # НЕ перезаписываем ячейки с формулами в шаблоне
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            continue
                        
                        # Конвертация NaN/None
                        if pd.isna(value):
                            cell.value = None
                        else:
                            cell.value = value
                
                wb.save(output)
                
                st.download_button(
                    label="⬇️ Скачать заполненный файл",
                    data=output.getvalue(),
                    file_name="кабельный_журнал_заполненный.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.success("✅ Файл готов. Форматирование и формулы шаблона сохранены.")
            except Exception as e:
                st.error(f"Ошибка экспорта: {e}")
                st.exception(e)
    
    # --- Сохранение проекта в JSON ---
    if st.session_state.df is not None and not st.session_state.df.empty:
        json_str = st.session_state.df.to_json(orient="records", force_ascii=False)
        st.download_button(
            "💾 Сохранить проект (.json)",
            data=json_str.encode("utf-8"),
            file_name="project.json",
            mime="application/json",
        )

else:
    st.info("👈 Загрузите шаблон `.xlsx` в боковой панели и нажмите **«Применить и распознать»**")
    
    st.markdown("""
    ### 📋 Как подготовить шаблон
    
    1. Откройте Excel и создайте файл с заголовками в любой строке (например, в строке 6)
    2. Выше можно разместить логотип, штампы, информацию о проекте
    3. Заголовки должны идти подряд (3+ заполненных ячеек в строке)
    4. Формулы, цвета, границы, ширины колонок — всё сохранится
    5. Сохраните как `.xlsx`
    
    ### ✅ Пример минимального шаблона
    
    | A | B | C | D | E |
    |---|---|---|---|---|
    | (логотип) | | | | |
    | Проект: Объект №1 | | | | |
    | | | | | |
    | | | | | |
    | | | | | |
    | № | Марка | Откуда | Куда | Длина, м |  ← строка 6
    | 1 | ВВГнг 3×2.5 | ЩР1 | Роз.1 | 15.5 | ← данные начнутся отсюда
    """)
