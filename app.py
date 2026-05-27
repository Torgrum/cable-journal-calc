"""Универсальный калькулятор кабельного журнала на основе Excel-шаблона."""
import streamlit as st
import pandas as pd
import io
import json
from openpyxl.utils import get_column_letter

from excel_tools import parse_template, export_to_excel

st.set_page_config(page_title="Универсальный КЖ", page_icon="📋", layout="wide")
st.title("📋 Универсальный калькулятор кабельного журнала")


# ==================== Инициализация session_state ====================

if "template_bytes" not in st.session_state:
    st.session_state.template_bytes = None
if "config" not in st.session_state:
    st.session_state.config = None
if "df_initial" not in st.session_state:
    st.session_state.df_initial = None
if "col_width_mode" not in st.session_state:
    st.session_state.col_width_mode = "Средний"


# ==================== Функции ====================

def get_width_for_mode(mode, col_name, col_type):
    if mode == "Компактный":
        return 80
    if mode == "Средний":
        return 150
    if mode == "Широкий":
        return 250
    # Авто
    estimated = max(80, min(400, len(str(col_name)) * 8 + 30))
    if col_type == "number":
        estimated = max(80, estimated - 30)
    return estimated


def reset_editor():
    """Сбрасывает состояние редактора при смене шаблона."""
    if "data_editor" in st.session_state:
        del st.session_state["data_editor"]


# ==================== Сайдбар ====================

with st.sidebar:
    st.header("⚙️ Настройки")

    uploaded = st.file_uploader("📄 Загрузить шаблон .xlsx", type=["xlsx"])
    if uploaded is not None:
        # Сохраняем байты шаблона
        new_bytes = uploaded.getvalue()
        if st.session_state.template_bytes != new_bytes:
            st.session_state.template_bytes = new_bytes
            # Сбрасываем старую конфигурацию при загрузке нового файла
            st.session_state.config = None
            st.session_state.df_initial = None

    if st.session_state.template_bytes is not None:
        # Список листов
        from openpyxl import load_workbook
        wb_tmp = load_workbook(io.BytesIO(st.session_state.template_bytes), read_only=True)
        sheet_names = wb_tmp.sheetnames
        wb_tmp.close()

        selected_sheet = st.selectbox("Лист", sheet_names)

        st.markdown("**Диапазон заголовка** (напр. `B5:J6`)")
        header_range = st.text_input("Диапазон", value="A1:E1", placeholder="B5:J6")

        if st.button("🔄 Применить шаблон", type="primary"):
            try:
                cfg = parse_template(st.session_state.template_bytes, selected_sheet, header_range)
                st.session_state.config = cfg
                # Создаём каркас DataFrame с правильными колонками (один раз!)
                st.session_state.df_initial = pd.DataFrame(columns=list(cfg["columns"].keys()))
                # Сбрасываем состояние редактора, чтобы он подхватил новые колонки
                reset_editor()
                st.success(f"✅ Шаблон распознан. Столбцов: {len(cfg['columns'])}")
                st.rerun()
            except Exception as e:
                st.error(f"Ошибка: {e}")

    st.divider()
    st.header("📏 Отображение")
    st.session_state.col_width_mode = st.radio(
        "Размер столбцов",
        options=["Компактный", "Средний", "Широкий", "Авто"],
        index=["Компактный", "Средний", "Широкий", "Авто"].index(st.session_state.col_width_mode),
    )
    st.info("💡 Ширину столбцов можно менять мышкой — перетаскивайте границу заголовка.")


# ==================== Основная область ====================

if st.session_state.config is None:
    st.info("👈 Загрузите шаблон `.xlsx` слева и укажите диапазон заголовков.")
    st.markdown("""
    ### 📋 Как пользоваться
    
    1. Загрузите Excel-шаблон
    2. Укажите диапазон ячеек заголовка (например, `A6:E6` или `B5:J6` для многострочного)
    3. Нажмите «Применить шаблон»
    4. Заполните данные в таблице
    5. Нажмите «Сформировать Excel» — данные запишутся в шаблон с сохранением формул и стиля
    """)
    st.stop()

cfg = st.session_state.config
cols = list(cfg["columns"].keys())

# Гарантируем, что каркас DataFrame существует
if st.session_state.df_initial is None or list(st.session_state.df_initial.columns) != cols:
    st.session_state.df_initial = pd.DataFrame(columns=cols)
    reset_editor()

# --- Превью структуры ---
with st.expander("📐 Структура шаблона", expanded=False):
    c1, c2, c3 = st.columns(3)
    c1.write(f"**Лист:** `{cfg['ws_name']}`")
    c2.write(f"**Диапазон:** `{cfg['header_range']}`")
    c3.write(f"**Старт данных:** строка {cfg['data_start_row']}")

    st.markdown("**📋 Визуализация заголовка:**")
    st.markdown(cfg["header_html"], unsafe_allow_html=True)

    if cfg["merged_ranges"]:
        st.info(f"🔗 Объединённых диапазонов: **{len(cfg['merged_ranges'])}**")

    st.dataframe(pd.DataFrame([
        {"Имя": n, "Excel": get_column_letter(i),
         "Тип": "🔢" if cfg["types"][n] == "number" else "🔤"}
        for n, i in cfg["columns"].items()
    ]), hide_index=True, use_container_width=True)

st.subheader("📝 Заполнение данных")
st.caption("💡 Данные сохраняются автоматически. Перезагрузка страницы сбросит их — используйте 💾 Сохранить проект перед закрытием.")

# Конфигурация колонок
col_cfg = {}
for n in cols:
    w = get_width_for_mode(st.session_state.col_width_mode, n, cfg["types"][n])
    if cfg["types"][n] == "number":
        col_cfg[n] = st.column_config.NumberColumn(n, step=0.5, width=w)
    else:
        col_cfg[n] = st.column_config.TextColumn(n, width=w)

# 🔑 КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ:
# - data=st.session_state.df_initial (НЕИЗМЕНЯЕМЫЙ каркас)
# - key="data_editor" (виджет сам хранит состояние)
# - НЕ перезаписываем df_initial на каждом рендере
edited_df = st.data_editor(
    st.session_state.df_initial,
    column_config=col_cfg,
    column_order=cols,
    hide_index=True,
    use_container_width=True,
    num_rows="dynamic",
    key="data_editor",
)

# Фильтруем пустые строки для сводки и экспорта
df_clean = edited_df.copy()
# Убираем полностью пустые строки
df_clean = df_clean.replace('', pd.NA).dropna(how='all').reset_index(drop=True)

# --- Сводка ---
if not df_clean.empty:
    st.subheader("📊 Сводка")
    c1, c2 = st.columns(2)
    c1.metric("Строк", len(df_clean))
    num_cols = df_clean.select_dtypes(include=['number']).columns
    if len(num_cols) > 0:
        c2.metric("Сумма", f"{df_clean[num_cols].sum().sum():.2f}")

    st.divider()

    # --- Экспорт ---
    if st.button("📥 Сформировать Excel", type="primary"):
        try:
            result_bytes = export_to_excel(
                template_bytes=st.session_state.template_bytes,
                ws_name=cfg["ws_name"],
                data_start_row=cfg["data_start_row"],
                columns_map=cfg["columns"],
                df=df_clean,
            )
            st.download_button(
                "⬇️ Скачать файл",
                result_bytes,
                "journal_filled.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.success(f"✅ Записано {len(df_clean)} строк. Формулы и стиль шаблона сохранены.")
        except Exception as e:
            st.error(f"Ошибка экспорта: {e}")
            st.exception(e)

    # --- Сохранение проекта в JSON ---
    try:
        json_str = df_clean.to_json(orient="records", force_ascii=False, default_handler=str)
        st.download_button(
            "💾 Сохранить проект (.json)",
            json_str.encode("utf-8"),
            "project.json",
            "application/json",
        )
    except Exception:
        pass

    # --- Загрузка проекта ---
    uploaded_project = st.file_uploader("📂 Загрузить сохранённый проект", type=["json"], key="proj_loader")
    if uploaded_project is not None:
        try:
            data = json.loads(uploaded_project.read().decode("utf-8"))
            loaded_df = pd.DataFrame(data)
            # Приводим к нужным колонкам
            for c in cols:
                if c not in loaded_df.columns:
                    loaded_df[c] = None
            loaded_df = loaded_df[cols]
            # Обновляем каркас и сбрасываем виджет
            st.session_state.df_initial = loaded_df
            reset_editor()
            st.success(f"✅ Загружено {len(loaded_df)} строк")
            st.rerun()
        except Exception as e:
            st.error(f"Ошибка загрузки проекта: {e}")

else:
    st.info("Добавьте строки в таблицу (кнопка + внизу таблицы).")
