"""Универсальный калькулятор кабельного журнала с объединёнными заголовками и авто-шириной."""
import streamlit as st
import pandas as pd
import io
import json
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, DataReturnMode, GridUpdateMode

from excel_tools import parse_template, export_to_excel

st.set_page_config(page_title="Универсальный КЖ", page_icon="📋", layout="wide")
st.title("📋 Универсальный калькулятор кабельного журнала")

# ==================== Session state ====================
if "template_bytes" not in st.session_state:
    st.session_state.template_bytes = None
if "config" not in st.session_state:
    st.session_state.config = None
if "df_initial" not in st.session_state:
    st.session_state.df_initial = None

def reset_editor():
    if "ag_grid_main" in st.session_state:
        del st.session_state["ag_grid_main"]

def build_column_groups(columns_map, types_map, merged_ranges, header_coords):
    """Строит структуру columnDefs для Ag-Grid с объединениями."""
    r1, c1, r2, c2 = header_coords
    ordered = sorted(columns_map.items(), key=lambda x: x[1])
    groups = []
    processed_cols = set()
    
    for col_name, excel_col in ordered:
        if excel_col in processed_cols:
            continue
            
        top_merge = None
        for mr in merged_ranges:
            if mr["r1"] == r1 and mr["c1"] <= excel_col <= mr["c2"]:
                top_merge = mr
                break
                
        if top_merge and top_merge["colspan"] > 1:
            group_children = []
            group_name = str(top_merge["value"]) if top_merge["value"] else ""
            for child_col in range(top_merge["c1"], top_merge["c2"] + 1):
                child_name = next((n for n, idx in columns_map.items() if idx == child_col), None)
                if child_name:
                    group_children.append({
                        "headerName": child_name,
                        "field": child_name,
                        "editable": True,
                        "type": "numericColumn" if types_map[child_name] == "number" else None,
                        "resizable": True, "sortable": True, "filter": True,
                    })
                    processed_cols.add(child_col)
            if group_children:
                groups.append({"headerName": group_name, "children": group_children})
        else:
            groups.append({
                "headerName": col_name,
                "field": col_name,
                "editable": True,
                "type": "numericColumn" if types_map[col_name] == "number" else None,
                "resizable": True, "sortable": True, "filter": True,
            })
            processed_cols.add(excel_col)
    return groups

# ==================== Сайдбар ====================
with st.sidebar:
    st.header("⚙️ Настройки")
    uploaded = st.file_uploader("📄 Загрузить шаблон .xlsx", type=["xlsx"])
    if uploaded is not None:
        new_bytes = uploaded.getvalue()
        if st.session_state.template_bytes != new_bytes:
            st.session_state.template_bytes = new_bytes
            st.session_state.config = None
            st.session_state.df_initial = None

    if st.session_state.template_bytes is not None:
        wb_tmp = load_workbook(io.BytesIO(st.session_state.template_bytes), read_only=True)
        selected_sheet = st.selectbox("Лист", wb_tmp.sheetnames)
        wb_tmp.close()

        st.markdown("**Диапазон заголовка** (напр. `B5:J6`)")
        header_range = st.text_input("Диапазон", value="A1:E1", placeholder="B5:J6")

        if st.button("🔄 Применить шаблон", type="primary"):
            try:
                cfg = parse_template(st.session_state.template_bytes, selected_sheet, header_range)
                st.session_state.config = cfg
                st.session_state.df_initial = pd.DataFrame(columns=list(cfg["columns"].keys()))
                reset_editor()
                st.success(f"✅ Шаблон распознан. Столбцов: {len(cfg['columns'])}")
                st.rerun()
            except Exception as e:
                st.error(f"Ошибка: {e}")

# ==================== Основная область ====================
if st.session_state.config is None:
    st.info("👈 Загрузите шаблон `.xlsx` слева и укажите диапазон заголовков.")
    st.stop()

cfg = st.session_state.config
cols = list(cfg["columns"].keys())

if st.session_state.df_initial is None or list(st.session_state.df_initial.columns) != cols:
    st.session_state.df_initial = pd.DataFrame(columns=cols)

with st.expander("📐 Структура шаблона", expanded=False):
    st.write(f"**Лист:** `{cfg['ws_name']}` | **Диапазон:** `{cfg['header_range']}` | **Старт данных:** строка {cfg['data_start_row']}")

st.subheader("📝 Заполнение данных")

# --- Кнопки управления ---
col_add, col_clear, _ = st.columns([1, 1, 4])
with col_add:
    if st.button("➕ Добавить строку"):
        new_row = pd.DataFrame([{c: None for c in cols}])
        st.session_state.df_initial = pd.concat([st.session_state.df_initial, new_row], ignore_index=True)
        reset_editor()
        st.rerun()
with col_clear:
    if st.button("🗑️ Очистить всё"):
        st.session_state.df_initial = pd.DataFrame(columns=cols)
        reset_editor()
        st.rerun()

# ============================================================
# 🔑 AG-GRID С ОБЪЕДИНЕНИЯМИ И АВТО-ШИРИНОЙ
# ============================================================
column_groups = build_column_groups(cfg["columns"], cfg["types"], cfg["merged_ranges"], cfg["header_coords"])

gb = GridOptionsBuilder.from_dataframe(st.session_state.df_initial)
gb.configure_default_column(editable=True, resizable=True, sortable=True, filter=True, wrapText=True, autoHeight=True)

grid_options = gb.build()
grid_options["columnDefs"] = column_groups
grid_options["rowSelection"] = "multiple"
grid_options["animateRows"] = True
grid_options["stopEditingWhenCellsLoseFocus"] = True
grid_options["singleClickEdit"] = True

# 🔑 БЛОК АВТО-ШИРИНЫ (растягивает таблицу под экран)
grid_options["autoSizeStrategy"] = {
    "type": "fitGridWidth",
    "skipHeader": False,
    "defaultMinWidth": 100,
    "defaultMaxWidth": 400,
}
grid_options["domLayout"] = "autoHeight"

# Рендер таблицы
grid_response = AgGrid(
    st.session_state.df_initial,
    gridOptions=grid_options,
    height=500,
    width="100%",
    data_return_mode=DataReturnMode.AS_INPUT,
    update_mode=GridUpdateMode.VALUE_CHANGED,
    fit_columns_on_grid_load=True,
    allow_unsafe_jscode=True,
    theme="streamlit",
    key="ag_grid_main",
)

# Получаем данные из таблицы
edited_df = pd.DataFrame(grid_response["data"]) if grid_response["data"] is not None else st.session_state.df_initial
st.session_state.df_initial = edited_df

df_clean = edited_df.copy().replace('', pd.NA).dropna(how='all').reset_index(drop=True)

# ==================== Сводка и Экспорт ====================
if not df_clean.empty:
    st.subheader("📊 Сводка")
    c1, c2 = st.columns(2)
    c1.metric("Строк заполнено", len(df_clean))
    num_cols = df_clean.select_dtypes(include=['number']).columns
    if len(num_cols) > 0:
        c2.metric("Сумма по числам", f"{df_clean[num_cols].sum().sum():.2f}")

    st.divider()

    if st.button("📥 Сформировать Excel", type="primary"):
        try:
            result_bytes = export_to_excel(
                template_bytes=st.session_state.template_bytes,
                ws_name=cfg["ws_name"],
                data_start_row=cfg["data_start_row"],
                columns_map=cfg["columns"],
                df=df_clean,
            )
            st.download_button("⬇️ Скачать файл", result_bytes, "journal_filled.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
            st.success("✅ Файл готов! Объединения и стили шаблона сохранены.")
        except Exception as e:
            st.error(f"Ошибка экспорта: {e}")

    # JSON Сохранение/Загрузка
    try:
        json_str = df_clean.to_json(orient="records", force_ascii=False, default_handler=str)
        st.download_button("💾 Сохранить проект (.json)", json_str.encode("utf-8"), "project.json", "application/json")
    except Exception:
        pass

    uploaded_project = st.file_uploader("📂 Загрузить сохранённый проект", type=["json"], key="proj_loader")
    if uploaded_project is not None:
        try:
            data = json.loads(uploaded_project.read().decode("utf-8"))
            loaded_df = pd.DataFrame(data)
            for c in cols:
                if c not in loaded_df.columns: loaded_df[c] = None
            st.session_state.df_initial = loaded_df[cols]
            reset_editor()
            st.success(f"✅ Загружено {len(loaded_df)} строк")
            st.rerun()
        except Exception as e:
            st.error(f"Ошибка: {e}")
else:
    st.info("Таблица пуста. Нажмите **➕ Добавить строку** выше, чтобы начать.")
