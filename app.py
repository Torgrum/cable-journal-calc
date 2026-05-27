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
if "template" not in st.session_state:
    st.session_state.template = None
if "config" not in st.session_state:
    st.session_state.config = None
# df хранится ВНУТРИ data_editor через key="editor"


def get_merged_cell_value(ws, row, col):
    """Возвращает значение объединённой ячейки (из верхнего-левого угла)."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                return ws.cell(row=rng.min_row, column=rng.min_col).value
    return cell.value


def find_header_row(ws, max_search=20):
    """Ищет первую строку с ≥3 подряд заполненными ячейками."""
    for row in range(1, max_search + 1):
        filled = 0
        for col in range(1, ws.max_column + 1):
            val = get_merged_cell_value(ws, row, col)
            if val not in (None, ""):
                filled += 1
            else:
                filled = 0
            if filled >= 3:
                return row
    return 1


def parse_template(file_bytes, sheet_name=None, header_row=None, header_rows=1):
    """
    Разбирает шаблон Excel.
    header_rows — сколько строк занимает заголовок (1, 2, 3...).
    """
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name] if sheet_name else wb.active

    if header_row is None:
        header_row = find_header_row(ws)

    # Читаем заголовки из нескольких строк
    columns = {}  # {итоговое_имя: номер_колонки}
    max_col = ws.max_column

    for col in range(1, max_col + 1):
        parts = []
        for r in range(header_row, header_row + header_rows):
            val = get_merged_cell_value(ws, r, col)
            if val not in (None, ""):
                parts.append(str(val).strip().replace("\n", " "))
        if parts:
            # Уникальное имя: объединяем через " / "
            name = " / ".join(parts)
            # Делаем имя уникальным, если дублируется
            base, counter = name, 1
            while name in columns:
                name = f"{base} ({counter})"
                counter += 1
            columns[name] = col

    # Определяем типы по первой строке данных
    data_start_row = header_row + header_rows
    types = {}
    for col_name, col_idx in columns.items():
        sample = get_merged_cell_value(ws, data_start_row, col_idx)
        if isinstance(sample, (int, float)) and not isinstance(sample, bool):
            types[col_name] = "number"
        else:
            types[col_name] = "text"

    return {
        "ws_name": ws.title,
        "header_row": header_row,
        "header_rows": header_rows,
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

        wb_tmp = load_workbook(io.BytesIO(uploaded.getvalue()), read_only=True)
        sheet_names = wb_tmp.sheetnames
        wb_tmp.close()

        selected_sheet = st.selectbox("Лист", sheet_names)

        # 🔧 АВТООПРЕДЕЛЕНИЕ строки заголовков
        wb_preview = load_workbook(io.BytesIO(uploaded.getvalue()), read_only=True)
        ws_preview = wb_preview[selected_sheet]
        auto_row = find_header_row(ws_preview)
        wb_preview.close()

        header_row = st.number_input(
            "Строка начала заголовков",
            min_value=1, max_value=50, value=auto_row,
            help="Первая строка, где начинаются заголовки"
        )
        header_rows = st.number_input(
            "Сколько строк занимает заголовок",
            min_value=1, max_value=10, value=1,
            help="Если заголовок в 2-3 строки (например: 'Трасса' сверху, 'Начало/Конец' снизу) — укажите 2 или 3"
        )

        if st.button("🔄 Применить и распознать", type="primary"):
            try:
                cfg = parse_template(
                    st.session_state.template,
                    sheet_name=selected_sheet,
                    header_row=int(header_row),
                    header_rows=int(header_rows),
                )
                st.session_state.config = cfg
                # 🔧 Правильная инициализация editor state
                if "editor_init" not in st.session_state:
                    st.session_state.editor_init = []
                st.success(f"✅ Найдено {len(cfg['columns'])} столбцов. Данные начнутся со строки {cfg['data_start_row']}.")
                st.rerun()
            except Exception as e:
                st.error(f"Ошибка разбора: {e}")

    st.divider()
    st.caption("💡 Формулы, рамки, стили шаблона сохранятся")


# --- Основная область ---
if st.session_state.config:
    cfg = st.session_state.config

    with st.expander("📐 Структура шаблона", expanded=False):
        st.write(f"**Лист:** `{cfg['ws_name']}`")
        st.write(f"**Заголовок:** строки {cfg['header_row']}–{cfg['header_row'] + cfg['header_rows'] - 1}")
        st.write(f"**Данные начнутся со строки:** {cfg['data_start_row']}")
        st.write(f"**Столбцов:** {len(cfg['columns'])}")

        cols_preview = pd.DataFrame([
            {
                "Имя в приложении": name,
                "Колонка Excel": get_column_letter(idx),
                "Тип": "🔢 число" if cfg["types"][name] == "number" else "🔤 текст",
            }
            for name, idx in cfg["columns"].items()
        ])
        st.dataframe(cols_preview, use_container_width=True, hide_index=True)

    # --- Редактор таблицы ---
    st.subheader("📝 Заполнение данных")

    # Динамическая конфигурация
    column_config = {}
    for col_name in cfg["columns"].keys():
        if cfg["types"].get(col_name) == "number":
            column_config[col_name] = st.column_config.NumberColumn(col_name, step=0.5)
        else:
            column_config[col_name] = st.column_config.TextColumn(col_name)

    # 🔧 КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ: используем key="editor"
    # Данные теперь хранятся в st.session_state["editor"] автоматически
    st.data_editor(
        st.session_state.get("editor_init", []),
        column_config=column_config,
        column_order=list(cfg["columns"].keys()),
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        key="editor",
    )

    # Забираем данные из editor state
    editor_state = st.session_state.get("editor", {"added_rows": [], "edited_rows": {}, "deleted_rows": []})
    
    # Формируем DataFrame из состояния редактора
    df_rows = []
    if editor_state.get("added_rows"):
        df_rows.extend(editor_state["added_rows"])
    
    # Применяем правки существующих строк
    if hasattr(st.session_state, "_last_rows"):
        for i, row in enumerate(st.session_state._last_rows):
            if i in editor_state.get("edited_rows", {}):
                row.update(editor_state["edited_rows"][i])
            df_rows.append(row)

    # --- Сводка ---
    if df_rows:
        st.subheader("📊 Сводка")
        c1, c2 = st.columns(2)
        c1.metric("Строк заполнено", len(df_rows))
        # Сохраняем для последующих операций
        st.session_state._last_rows = df_rows

    st.divider()

    # --- Экспорт ---
    if df_rows:
        if st.button("📥 Сформировать Excel по шаблону", type="primary"):
            try:
                output = io.BytesIO()
                wb = load_workbook(io.BytesIO(st.session_state.template))
                ws = wb[cfg["ws_name"]]

                data_start = cfg["data_start_row"]
                col_map = cfg["columns"]  # {имя: номер_колонки}

                for i, row_data in enumerate(df_rows):
                    excel_row = data_start + i
                    for col_name, value in row_data.items():
                        if col_name not in col_map:
                            continue
                        excel_col = col_map[col_name]
                        cell = ws.cell(row=excel_row, column=excel_col)

                        # НЕ перезаписываем формулы шаблона
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
                st.success("✅ Готово! Форматирование и формулы шаблона сохранены.")
            except Exception as e:
                st.error(f"Ошибка экспорта: {e}")

        # Сохранение проекта в JSON
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
    st.info("👈 Загрузите шаблон `.xlsx` слева и нажмите **«Применить и распознать»**")

    st.markdown("""
    ### 📋 Многострочные заголовки
    
    Приложение умеет работать с заголовками в 2–3 строки. Примеры:
    
    **Пример 1: объединённые ячейки**
    | | A | B | C |
    |---|---|---|---|
    | Строка 5 | **Трасса** (merged B5:C5) | | |
    | Строка 6 | | Начало | Конец |
    | Строка 7 | | 1 | 2 ← данные |
    
    → Укажите: строка начала = **5**, строк заголовка = **2**
    → Получите столбцы: `Трасса / Начало`, `Трасса / Конец`
    
    **Пример 2: перенос текста**
    | Строка 4 | № | **Длина кабеля, м** | Примечание |
    | Строка 5 | | (с запасом 15%) | |
    | Строка 6 | 1 | 15.5 | — ← данные |
    
    → Укажите: строка = **4**, строк = **2**
    → Получите: `Длина кабеля, м / (с запасом 15%)`
    """)
