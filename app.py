import streamlit as st
import pandas as pd
import io
import json
from openpyxl import load_workbook

st.set_page_config(page_title="КЖ — СКС", page_icon="🔌", layout="wide")
st.title("🔌 Калькулятор кабельного журнала (СКС)")

# --- Инициализация ---
if "df" not in st.session_state:
    # Колонки соответствуют входным данным шаблона (без формул)
    st.session_state.df = pd.DataFrame(columns=[
        "Обозначение", "Шкаф", "Панель", "Порт", "Помещение", "Этаж",
        "Розетка", "Открыто_м", "Гофра_м", "ТЖГ_м", "Подъём_розетка",
        "Вертикаль_кросс", "Марка", "Жилы", "Длина_линии"
    ])
if "template" not in st.session_state:
    st.session_state.template = None

# --- Загрузка шаблона ---
with st.expander("⚙️ Шаблон и проект", expanded=True):
    col1, col2, col3 = st.columns(3)
    
    tpl = col1.file_uploader("📄 Загрузить КЖ.xlsm", type=["xlsm"])
    if tpl:
        st.session_state.template = tpl.read()
        st.success("✅ Шаблон загружен")
    
    if not st.session_state.df.empty:
        json_str = st.session_state.df.to_json(orient="records", force_ascii=False)
        col2.download_button("💾 Проект (.json)", json_str.encode("utf-8"), 
                           "project.json", "application/json", use_container_width=True)

# --- Экспорт в XLSM с сохранением формул ---
def export_to_xlsm(df, template_bytes):
    output = io.BytesIO()
    wb = load_workbook(filename=io.BytesIO(template_bytes), keep_vba=True)
    ws = wb["Кабельный"]  # работаем только с листом данных
    
    # Данные начинаем писать с 8-й строки (после шапки)
    START_ROW = 8
    
    # Маппинг: колонка DataFrame → колонка Excel (1-based)
    COL_MAP = {
        "Обозначение": 2,    # B
        "Шкаф": 3,           # C
        "Панель": 4,         # D
        "Порт": 5,           # E
        "Помещение": 20,     # T → формула F=T сработает автоматически
        "Этаж": 7,           # G
        "Розетка": 8,        # H
        "Открыто_м": 9,      # I
        # J, K, N — формулы, не трогаем
        "Гофра_м": 22,       # V
        "ТЖГ_м": 23,         # W
        "Подъём_розетка": 24,# X
        "Вертикаль_кросс": 26,# Z
        "Марка": 12,         # L
        "Жилы": 13,          # M
        "Длина_линии": 21,   # U — база для формул J,K
    }
    
    for i, (_, row) in enumerate(df.iterrows()):
        excel_row = START_ROW + i
        for col_name, excel_col in COL_MAP.items():
            if col_name in row and pd.notna(row[col_name]):
                val = row[col_name]
                # Конвертация типов для openpyxl
                if isinstance(val, (pd.Int64Dtype, pd.Float64Dtype)):
                    val = float(val) if pd.notna(val) else None
                ws.cell(row=excel_row, column=excel_col, value=val)
    
    wb.save(output)
    return output.getvalue()

# --- Кнопка экспорта ---
if not st.session_state.df.empty and st.session_state.template:
    if st.button("📥 Выгрузить в КЖ.xlsm", type="primary"):
        try:
            xlsm_data = export_to_xlsm(st.session_state.df, st.session_state.template)
            st.download_button(
                label="⬇️ Скачать готовый файл",
                data=xlsm_data,
                file_name="КЖ_заполненный.xlsm",
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                use_container_width=True
            )
            st.success("✅ Файл сформирован! Формулы и макросы сохранены.")
        except Exception as e:
            st.error(f"Ошибка экспорта: {e}")

# --- Редактор данных ---
st.subheader("📝 Ввод данных по линиям")
df = st.data_editor(
    st.session_state.df,
    column_config={
        "Обозначение": st.column_config.TextColumn("№ линии", width="small"),
        "Шкаф": st.column_config.TextColumn("Шкаф", width="small"),
        "Панель": st.column_config.TextColumn("Панель", width="tiny"),
        "Порт": st.column_config.TextColumn("Порт", width="tiny"),
        "Помещение": st.column_config.TextColumn("Пом. №", width="tiny"),
        "Этаж": st.column_config.NumberColumn("Этаж", min_value=0, width="tiny"),
        "Розетка": st.column_config.TextColumn("Розетка", width="medium"),
        "Открыто_м": st.column_config.NumberColumn("Открыто, м", min_value=0, step=0.5),
        "Гофра_м": st.column_config.NumberColumn("Гофра, м", min_value=0, step=0.5),
        "ТЖГ_м": st.column_config.NumberColumn("ТЖГ, м", min_value=0, step=0.5),
        "Подъём_розетка": st.column_config.NumberColumn("↑ к розетке, м", min_value=0, step=0.1),
        "Вертикаль_кросс": st.column_config.NumberColumn("↑ в кросс, м", min_value=0, step=0.1),
        "Марка": st.column_config.TextColumn("Марка", width="medium"),
        "Жилы": st.column_config.TextColumn("Жилы", width="tiny"),
        "Длина_линии": st.column_config.NumberColumn("База, м", min_value=0, step=0.5,
            help="Базовая длина для формул запаса")
    },
    hide_index=True,
    use_container_width=True,
    num_rows="dynamic"
)
st.session_state.df = df

# --- Сводка ---
if not df.empty:
    st.subheader("📊 Сводка")
    c1, c2, c3 = st.columns(3)
    c1.metric("Линий", len(df))
    c2.metric("Кабель (итого)", f"{df['Длина_линии'].sum():.1f} м")
    c3.metric("Средний запас", "15% (в формулах)")
