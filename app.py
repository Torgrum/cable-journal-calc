import streamlit as st
import pandas as pd
import io
import json
import base64

st.set_page_config(page_title="Кабельный журнал", layout="wide", page_icon="📋")
st.title("📋 Калькулятор кабельного журнала")

# Инициализация таблицы
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(columns=[
        "Марка", "Откуда", "Куда", "Длина трассы, м", "Запас, %", "Итого, м", "Примечание"
    ])

# Функция авторасчёта
def recalculate(df):
    if df.empty:
        return df
    df["Длина трассы, м"] = pd.to_numeric(df["Длина трассы, м"], errors="coerce").fillna(0.0)
    df["Запас, %"] = pd.to_numeric(df["Запас, %"], errors="coerce").fillna(0.0)
    df["Итого, м"] = (df["Длина трассы, м"] * (1 + df["Запас, %"] / 100)).round(2)
    return df

# 🔹 Сайдбар: Сохранение / Загрузка проекта
with st.sidebar:
    st.header("💾 Управление проектом")
    
    # Сохранить
    if not st.session_state.df.empty:
        json_data = st.session_state.df.to_dict(orient="records")
        json_str = json.dumps(json_data, ensure_ascii=False, indent=2)
        b64 = base64.b64encode(json_str.encode()).decode()
        st.download_button(
            "💾 Скачать проект (.json)",
            data=b64,
            file_name="project.json",
            mime="application/json",
            use_container_width=True
        )
    
    # Загрузить
    uploaded = st.file_uploader("📂 Загрузить проект", type=["json"])
    if uploaded:
        try:
            data = json.loads(uploaded.read().decode("utf-8"))
            st.session_state.df = recalculate(pd.DataFrame(data))
            st.success("✅ Проект загружен!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Ошибка файла: {e}")

# 🔹 Основная таблица
df = st.data_editor(
    st.session_state.df,
    column_config={
        "Марка": st.column_config.TextColumn("Марка кабеля"),
        "Откуда": st.column_config.TextColumn("Откуда"),
        "Куда": st.column_config.TextColumn("Куда"),
        "Длина трассы, м": st.column_config.NumberColumn("Длина трассы, м", min_value=0.0, step=0.5, format="%.2f"),
        "Запас, %": st.column_config.NumberColumn("Запас, %", min_value=0, max_value=100, step=1, default=5),
        "Итого, м": st.column_config.NumberColumn("Итого, м", format="%.2f", disabled=True),
        "Примечание": st.column_config.TextColumn("Примечание")
    },
    hide_index=True,
    num_rows="dynamic",
    use_container_width=True
)

# Обновляем данные и пересчитываем
st.session_state.df = recalculate(df)

# 🔹 Сводка и экспорт
if not st.session_state.df.empty:
    st.divider()
    st.subheader("📊 Сводка")
    c1, c2, c3 = st.columns(3)
    c1.metric("Кабелей", len(st.session_state.df))
    c2.metric("Общая длина", f"{st.session_state.df['Итого, м'].sum():.2f} м")
    c3.metric("Средний запас", f"{st.session_state.df['Запас, %'].mean():.1f} %")

    st.markdown("🔹 **Расход по маркам**")
    grouped = st.session_state.df.groupby("Марка").agg(
        Количество=("Марка", "size"),
        Общая_длина=("Итого, м", "sum")
    ).reset_index()
    st.dataframe(grouped, use_container_width=True, hide_index=True)

    # Подготовка Excel (добавляем нумерацию только при экспорте)
    export_df = st.session_state.df.copy()
    export_df.insert(0, "№ п/п", range(1, len(export_df) + 1))
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="Кабели", index=False)
        grouped.to_excel(writer, sheet_name="Сводка", index=False)
    
    st.download_button(
        "📥 Скачать кабельный журнал (.xlsx)",
        data=output.getvalue(),
        file_name="кабельный_журнал.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.info("👆 Нажмите `+ Add row` в таблице, чтобы начать работу.")