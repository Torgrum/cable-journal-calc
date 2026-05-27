import re
import streamlit as st
import pandas as pd
import io
import json
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import MergedCell

st.set_page_config(page_title="Универсальный КЖ", page_icon="📋", layout="wide")
st.title("📋 Универсальный калькулятор кабельного журнала")

# --- Инициализация состояния ---
# Данные таблицы храним отдельно от виджета
if "df_data" not in st.session_state:
    st.session_state.df_data = None
if "template" not in st.session_state:
    st.session_state.template = None
if "config" not in st.session_state:
    st.session_state.config = None


def parse_range(range_str):
    s = range_str.strip().upper().replace(" ", "")
    if not s: raise ValueError("Диапазон пустой")
    if ":" not in s: s = f"{s}:{s}"
    m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", s)
    if not m: raise ValueError("Формат: 'B5:J6'")
    c1, r1, c2, r2 = column_index_from_string(m.group(1)), int(m.group(2)), column_index_from_string(m.group(3)), int(m.group(4))
    return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def get_merged_value(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                return ws.cell(row=rng.min_row, column=rng.min_col).value
    return cell.value


def parse_template(file_bytes, sheet_name, header_range_str):
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name]
    r1, c1, r2, c2 = parse_range(header_range_str)
    
    columns = {}
    for col in range(c1, c2 + 1):
        parts = [str(get_merged_value(ws, r, col)).strip().replace("\n", " ") 
                 for r in range(r1, r2 + 1) if get_merged_value(ws, r, col) not in (None, "")]
        if parts:
            name = " / ".join(parts)
            base, counter = name, 1
            while name in columns:
                name = f"{base} ({counter})"
                counter += 1
            columns[name] = col
            
    if not columns: raise ValueError(f"В диапазоне '{header_range_str}' нет заголовков")
    
    data_start_row = r2 + 1
    types = {}
    for name, idx in columns.items():
        sample = get_merged_value(ws, data_start_row, idx)
        types[name] = "number" if isinstance(sample, (int, float)) and not isinstance(sample, bool) else "text"
        
    return {
        "ws_name": ws.title, "header_range": f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}",
        "data_start_row": data_start_row, "columns": columns, "types": types,
    }


# --- Боковая панель ---
with st.sidebar:
    st.header("⚙️ Настройки")
    uploaded = st.file_uploader("📄 Загрузить шаблон .xlsx", type=["xlsx"])

    if uploaded:
        st.session_state.template = uploaded.getvalue()
        wb_tmp = load_workbook(io.BytesIO(uploaded.getvalue()), read_only=True)
        selected_sheet = st.selectbox("Лист", wb_tmp.sheetnames)
        wb_tmp.close()

        st.markdown("**Диапазон заголовка** (напр. `B5:J6`)")
        header_range = st.text_input("Диапазон", value="A1:E1", placeholder="B5:J6")

        if st.button("🔄 Применить", type="primary"):
            try:
                cfg = parse_template(st.session_state.template, selected_sheet, header_range)
                st.session_state.config = cfg
                # 🔑 СБРОС ДАННЫХ при смене шаблона
                st.session_state.df_data = pd.DataFrame(columns=list(cfg["columns"].keys()))
                st.rerun()
            except Exception as e:
                st.error(f"Ошибка: {e}")

# --- Основная область ---
if st.session_state.config:
    cfg = st.session_state.config
    cols = list(cfg["columns"].keys())
    
    # Защита от рассинхрона колонок
    if st.session_state.df_data is None or list(st.session_state.df_data.columns) != cols:
        st.session_state.df_data = pd.DataFrame(columns=cols)

    with st.expander("📐 Структура", expanded=False):
        st.write(f"**Лист:** `{cfg['ws_name']}` | **Заголовок:** `{cfg['header_range']}` | **Старт данных:** строка {cfg['data_start_row']}")
        st.dataframe(pd.DataFrame([
            {"Имя": n, "Excel": get_column_letter(i), "Тип": "🔢" if cfg["types"][n]=="number" else "🔤"}
            for n, i in cfg["columns"].items()
        ]), hide_index=True, use_container_width=True)

    st.subheader("📝 Заполнение")
    
    # Конфигурация колонок
    col_cfg = {}
    for n in cols:
        if cfg["types"][n] == "number":
            col_cfg[n] = st.column_config.NumberColumn(n, step=0.5)
        else:
            col_cfg[n] = st.column_config.TextColumn(n)

    # 🔑 ГЛАВНОЕ ИСПРАВЛЕНИЕ:
    # 1. НЕ используем key="editor" для хранения данных
    # 2. Передаем текущий DF из session_state
    # 3. Результат сразу сохраняем обратно
    edited_df = st.data_editor(
        st.session_state.df_data,
        column_config=col_cfg,
        column_order=cols,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
    )
    st.session_state.df_data = edited_df

    # Очистка пустых строк для подсчета и экспорта
    df_clean = edited_df.replace('', pd.NA).dropna(how='all')
    
    if not df_clean.empty:
        st.subheader("📊 Сводка")
        c1, c2 = st.columns(2)
        c1.metric("Строк", len(df_clean))
        num_cols = df_clean.select_dtypes(include=['number']).columns
        if len(num_cols) > 0:
            c2.metric("Сумма", f"{df_clean[num_cols].sum().sum():.2f}")

        st.divider()
        
        if st.button("📥 Сформировать Excel", type="primary"):
            try:
                output = io.BytesIO()
                wb = load_workbook(io.BytesIO(st.session_state.template))
                ws = wb[cfg["ws_name"]]
                start_row = cfg["data_start_row"]
                col_map = cfg["columns"]

                # Очистка старых данных (опционально, чтобы не было хвостов)
                for r in range(start_row, start_row + 1000):
                    if all(ws.cell(row=r, column=c).value in (None, "") or 
                           (isinstance(ws.cell(row=r, column=c).value, str) and ws.cell(row=r, column=c).value.startswith("="))
                           for c in col_map.values()):
                        break
                    for c in col_map.values():
                        cell = ws.cell(row=r, column=c)
                        if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                            cell.value = None

                # Запись новых
                for i, (_, row) in enumerate(df_clean.iterrows()):
                    for col_name, val in row.items():
                        if col_name in col_map:
                            cell = ws.cell(row=start_row + i, column=col_map[col_name])
                            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                                cell.value = None if pd.isna(val) else val

                wb.save(output)
                st.download_button("⬇️ Скачать", output.getvalue(), "journal_filled.xlsx", 
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
                st.success(f"✅ Записано {len(df_clean)} строк.")
            except Exception as e:
                st.error(f"Ошибка: {e}")
                
        # JSON экспорт
        st.download_button("💾 JSON", df_clean.to_json(orient="records", force_ascii=False).encode("utf-8"), 
                           "project.json", "application/json")
    else:
        st.info("Добавьте строки в таблицу.")

else:
    st.info("👈 Загрузите шаблон и укажите диапазон заголовков (напр. `B5:J6`).")
