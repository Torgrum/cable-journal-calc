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


# --- Инициализация ---
if "template" not in st.session_state:
    st.session_state.template = None
if "config" not in st.session_state:
    st.session_state.config = None


# --- Парсер диапазона Excel: "B5:J6" → (start_row, start_col, end_row, end_col) ---
def parse_range(range_str):
    """Парсит диапазон вида 'B5:J6' или 'B5' (одна ячейка)."""
    s = range_str.strip().upper().replace(" ", "")
    if not s:
        raise ValueError("Диапазон пустой")
    
    if ":" not in s:
        s = f"{s}:{s}"
    
    m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", s)
    if not m:
        raise ValueError("Формат должен быть как 'B5:J6' (буквы-цифры:буквы-цифры)")
    
    col1 = column_index_from_string(m.group(1))
    row1 = int(m.group(2))
    col2 = column_index_from_string(m.group(3))
    row2 = int(m.group(4))
    
    if row1 > row2: row1, row2 = row2, row1
    if col1 > col2: col1, col2 = col2, col1
    
    return row1, col1, row2, col2


def get_merged_value(ws, row, col):
    """Возвращает значение ячейки, учитывая merged cells."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                return ws.cell(row=rng.min_row, column=rng.min_col).value
    return cell.value


def parse_template(file_bytes, sheet_name, header_range_str):
    """Разбирает шаблон по явно указанному диапазону заголовка."""
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name]
    
    row1, col1, row2, col2 = parse_range(header_range_str)
    
    # Формируем имена колонок: объединяем значения из всех строк диапазона через " / "
    columns = {}  # {имя: номер_колонки}
    for col in range(col1, col2 + 1):
        parts = []
        for r in range(row1, row2 + 1):
            val = get_merged_value(ws, r, col)
            if val not in (None, ""):
                parts.append(str(val).strip().replace("\n", " "))
        if parts:
            name = " / ".join(parts)
            # Уникальность имён
            base, counter = name, 1
            while name in columns:
                name = f"{base} ({counter})"
                counter += 1
            columns[name] = col
    
    if not columns:
        raise ValueError(f"В диапазоне '{header_range_str}' не найдено ни одного заголовка")
    
    # Типы — по первой строке данных (row2 + 1)
    data_start_row = row2 + 1
    types = {}
    for col_name, col_idx in columns.items():
        sample = get_merged_value(ws, data_start_row, col_idx)
        if isinstance(sample, (int, float)) and not isinstance(sample, bool):
            types[col_name] = "number"
        else:
            types[col_name] = "text"
    
    return {
        "ws_name": ws.title,
        "header_range": f"{get_column_letter(col1)}{row1}:{get_column_letter(col2)}{row2}",
        "data_start_row": data_start_row,
        "columns": columns,
        "types": types,
    }


# --- Боковая панель ---
with st.sidebar:
    st.header("⚙️ Настройки")

    uploaded = st.file_uploader("📄 Загрузить шаблон .xlsx", type=["xlsx"])

    if uploaded is not None:
        st.session_state.template = uploaded.getvalue()

        # Список листов
        wb_tmp = load_workbook(io.BytesIO(uploaded.getvalue()), read_only=True)
        sheet_names = wb_tmp.sheetnames
        wb_tmp.close()

        selected_sheet = st.selectbox("Лист", sheet_names)

        # 🔧 ЯВНЫЙ ВВОД ДИАПАЗОНА
        st.markdown("**Диапазон ячеек заголовка**")
        header_range = st.text_input(
            "Диапазон (как в Excel)",
            value="A1:E1",
            placeholder="Например: B5:J6",
            help="Укажите диапазон, где находятся заголовки столбцов. "
                 "Примеры: 'A6:E6' (одна строка), 'B5:J6' (две строки), 'C3:H4' (объединённые)."
        )
        
        # Быстрые примеры
        st.caption("Примеры: `A6:E6`, `B5:J6`, `C3:H4`")

        if st.button("🔄 Применить и распознать", type="primary"):
            try:
                cfg = parse_template(
                    st.session_state.template,
                    sheet_name=selected_sheet,
                    header_range_str=header_range,
                )
                st.session_state.config = cfg
                # Инициализация состояния редактора
                if "editor" not in st.session_state:
                    st.session_state.editor = []
                st.rerun()
            except ValueError as e:
                st.error(f"❌ {e}")
            except Exception as e:
                st.error(f"Ошибка: {e}")

    st.divider()
    st.caption("💡 Формулы, рамки и стили шаблона сохраняются")
    st.caption("💡 Данные записываются со строки, следующей после заголовка")


# --- Основная область ---
if st.session_state.config:
    cfg = st.session_state.config

    # Информация о структуре
    with st.expander("📐 Распознанная структура", expanded=True):
        c1, c2, c3 = st.columns(3)
        c1.write(f"**Лист:** `{cfg['ws_name']}`")
        c2.write(f"**Заголовок:** `{cfg['header_range']}`")
        c3.write(f"**Данные начнутся со строки:** {cfg['data_start_row']}")
        
        st.markdown(f"**Столбцов найдено: {len(cfg['columns'])}**")
        
        cols_preview = pd.DataFrame([
            {
                "№": i + 1,
                "Имя в приложении": name,
                "Колонка Excel": get_column_letter(idx),
                "Тип": "🔢 число" if cfg["types"][name] == "number" else "🔤 текст",
            }
            for i, (name, idx) in enumerate(cfg["columns"].items())
        ])
        st.dataframe(cols_preview, use_container_width=True, hide_index=True)

    # --- Редактор таблицы ---
    st.subheader("📝 Заполнение данных")

    column_config = {}
    for col_name in cfg["columns"].keys():
        if cfg["types"].get(col_name) == "number":
            column_config[col_name] = st.column_config.NumberColumn(col_name, step=0.5)
        else:
            column_config[col_name] = st.column_config.TextColumn(col_name)

    # 🔧 key="editor" — данные живут в st.session_state["editor"] автоматически
    st.data_editor(
        [],  # начальное состояние — пустой список, всё остальное в session_state
        column_config=column_config,
        column_order=list(cfg["columns"].keys()),
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        key="editor",
    )

    # Забираем актуальные данные из состояния редактора
    editor_state = st.session_state.get("editor", [])
    if isinstance(editor_state, dict):
        # Старый формат с added_rows — совместимость
        df_rows = editor_state.get("added_rows", [])
    else:
        df_rows = editor_state if editor_state else []
    
    # Фильтруем пустые строки (которые пользователь мог добавить случайно)
    df_rows = [
        row for row in df_rows
        if any(v not in (None, "", 0, 0.0) and not (isinstance(v, float) and pd.isna(v))
               for v in row.values())
    ]

    # Сводка
    if df_rows:
        st.subheader("📊 Сводка")
        c1, c2 = st.columns(2)
        c1.metric("Строк заполнено", len(df_rows))
        df_tmp = pd.DataFrame(df_rows)
        numeric_cols = [c for c in df_tmp.columns if df_tmp[c].dtype in ['float64', 'int64']]
        if numeric_cols:
            total = df_tmp[numeric_cols].sum().sum()
            c2.metric("Сумма по числам", f"{total:.2f}")

    st.divider()

    # --- Экспорт ---
    if df_rows:
        if st.button("📥 Сформировать Excel по шаблону", type="primary"):
            try:
                output = io.BytesIO()
                wb = load_workbook(io.BytesIO(st.session_state.template))
                ws = wb[cfg["ws_name"]]

                data_start = cfg["data_start_row"]
                col_map = cfg["columns"]

                # Опционально: очищаем старые данные ниже заголовка (по колонкам шаблона)
                # чтобы не было "хвостов" от предыдущих заполнений
                for row_idx in range(data_start, data_start + 500):  # запас
                    all_empty = True
                    for col_idx in col_map.values():
                        cell = ws.cell(row=row_idx, column=col_idx)
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            all_empty = False  # формула — не трогаем
                            continue
                        cell.value = None
                    if all_empty:
                        # Если строка полностью пустая (кроме формул) — можно остановиться
                        pass

                # Записываем новые данные
                for i, row_data in enumerate(df_rows):
                    excel_row = data_start + i
                    for col_name, value in row_data.items():
                        if col_name not in col_map:
                            continue
                        excel_col = col_map[col_name]
                        cell = ws.cell(row=excel_row, column=excel_col)

                        # Не перезаписываем формулы шаблона
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            continue

                        # Пустые значения
                        if value in (None, "") or (isinstance(value, float) and pd.isna(value)):
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
                st.success(f"✅ Готово! Записано {len(df_rows)} строк со строки {data_start}.")
            except Exception as e:
                st.error(f"Ошибка экспорта: {e}")
                st.exception(e)

        # Сохранение проекта
        json_str = json.dumps(df_rows, ensure_ascii=False, default=str)
        st.download_button(
            "💾 Сохранить проект (.json)",
            data=json_str.encode("utf-8"),
            file_name="project.json",
            mime="application/json",
        )
    else:
        st.info("Добавьте строки в таблицу — появится кнопка экспорта")

else:
    st.info("👈 Загрузите шаблон `.xlsx` слева и укажите диапазон заголовков")

    st.markdown("""
    ### 🎯 Как указать диапазон
    
    Откройте ваш файл в Excel и посмотрите координаты ячеек заголовка (в левом верхнем углу Excel).
    
    | Структура шаблона | Что указать |
    |-------------------|-------------|
    | Заголовки в **одну строку**, например A6:E6 | `A6:E6` |
    | Заголовки в **две строки** (например, «Трасса» сверху, «Начало/Конец» снизу) | `B5:C6` |
    | Заголовки в **три строки** | `B4:D6` |
    | Заголовки **только по центру листа** | `C3:H3` |
    
    ### 📐 Пример для типичного КЖ
    
    ```
    Строка 1: [логотип]
    Строка 2: Проект: ...
    Строка 3-5: (пусто)
    Строка 6: № | Марка | Откуда | Куда | Длина   ← заголовки
    Строка 7: 1 | ...   | ...    | ...  | 15.5    ← данные
    ```
    
    → Диапазон: **`A6:E6`**
    → Данные будут записаны со строки **7**
    
    ### 🧩 Многострочные заголовки
    
    ```
    Строка 5: [объединённая ячейка]  Трасса
    Строка 6:                        Начало | Конец
    Строка 7:                        1      | 2     ← данные
    ```
    
    → Диапазон: **`B5:C6`**
    → Получите столбцы: `Трасса / Начало`, `Трасса / Конец`
    """)
