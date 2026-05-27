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

# --- Состояние ---
if "df_data" not in st.session_state:
    st.session_state.df_data = None
if "template" not in st.session_state:
    st.session_state.template = None
if "config" not in st.session_state:
    st.session_state.config = None
if "col_width_mode" not in st.session_state:
    st.session_state.col_width_mode = "Средний"


def parse_range(range_str):
    s = range_str.strip().upper().replace(" ", "")
    if not s: raise ValueError("Диапазон пустой")
    if ":" not in s: s = f"{s}:{s}"
    m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", s)
    if not m: raise ValueError("Формат: 'B5:J6'")
    c1, r1, c2, r2 = column_index_from_string(m.group(1)), int(m.group(2)), \
                      column_index_from_string(m.group(3)), int(m.group(4))
    return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def get_merged_value(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                return ws.cell(row=rng.min_row, column=rng.min_col).value
    return cell.value


def get_merged_ranges_in_range(ws, r1, c1, r2, c2):
    result = []
    for rng in ws.merged_cells.ranges:
        if (rng.min_row <= r2 and rng.max_row >= r1 and
            rng.min_col <= c2 and rng.max_col >= c1):
            local_r1 = max(rng.min_row, r1)
            local_c1 = max(rng.min_col, c1)
            local_r2 = min(rng.max_row, r2)
            local_c2 = min(rng.max_col, c2)
            value = ws.cell(row=rng.min_row, column=rng.min_col).value
            result.append({
                "r1": local_r1, "c1": local_c1,
                "r2": local_r2, "c2": local_c2,
                "rowspan": local_r2 - local_r1 + 1,
                "colspan": local_c2 - local_c1 + 1,
                "value": value,
            })
    return result


def build_header_html(r1, c1, r2, c2, merged_ranges):
    nrows = r2 - r1 + 1
    ncols = c2 - c1 + 1
    covered = [[False] * ncols for _ in range(nrows)]
    
    html = '<table style="border-collapse: collapse; margin: 10px 0; font-family: monospace;">'
    for ri in range(nrows):
        html += "<tr>"
        for ci in range(ncols):
            if covered[ri][ci]:
                continue
            abs_r = r1 + ri
            abs_c = c1 + ci
            
            found = None
            for mr in merged_ranges:
                if mr["r1"] == abs_r and mr["c1"] == abs_c:
                    found = mr
                    break
            
            if found:
                for dr in range(found["rowspan"]):
                    for dc in range(found["colspan"]):
                        if ri + dr < nrows and ci + dc < ncols:
                            covered[ri + dr][ci + dc] = True
                val = str(found["value"]) if found["value"] not in (None, "") else ""
                html += (f'<td rowspan="{found["rowspan"]}" colspan="{found["colspan"]}" '
                         f'style="border: 1px solid #888; padding: 6px 10px; '
                         f'background: #e7f3ff; font-weight: bold; text-align: center; '
                         f'vertical-align: middle; min-width: 60px;">{val}</td>')
            else:
                html += (f'<td style="border: 1px solid #888; padding: 6px 10px; '
                         f'text-align: center; min-width: 60px;"></td>')
        html += "</tr>"
    html += "</table>"
    return html


def parse_template(file_bytes, sheet_name, header_range_str):
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name]
    r1, c1, r2, c2 = parse_range(header_range_str)

    merged_ranges = get_merged_ranges_in_range(ws, r1, c1, r2, c2)

    # 🔗 НОВОЕ: собираем информацию об объединениях для каждой колонки
    # col_merge_info[col_idx] = {"group_name": "...", "is_merged": True/False, "merge_span": 2}
    col_merge_info = {}
    for mr in merged_ranges:
        if mr["rowspan"] == 1 and mr["colspan"] > 1:
            # Горизонтальное объединение в одной строке
            group_name = str(mr["value"]) if mr["value"] else ""
            for c in range(mr["c1"], mr["c2"] + 1):
                col_merge_info[c] = {
                    "group_name": group_name,
                    "is_merged": True,
                    "merge_span": mr["colspan"],
                    "merge_start": c == mr["c1"],  # первая колонка в группе
                }
        elif mr["colspan"] == 1 and mr["rowspan"] > 1:
            # Вертикальное объединение — одна колонка
            col_merge_info[mr["c1"]] = {
                "group_name": str(mr["value"]) if mr["value"] else "",
                "is_merged": True,
                "merge_span": 1,
                "merge_start": True,
            }

    def get_top_left_for_cell(row, col):
        for mr in merged_ranges:
            if mr["r1"] <= row <= mr["r2"] and mr["c1"] <= col <= mr["c2"]:
                return (mr["r1"], mr["c1"])
        return (row, col)

    columns = {}
    for col in range(c1, c2 + 1):
        parts_by_group = {}
        for r in range(r1, r2 + 1):
            tl = get_top_left_for_cell(r, col)
            val = get_merged_value(ws, r, col)
            if val not in (None, ""):
                key = str(val).strip().replace("\n", " ")
                parts_by_group.setdefault(tl, []).append(key)
        
        seen = []
        for tl in sorted(parts_by_group.keys()):
            if parts_by_group[tl]:
                v = parts_by_group[tl][0]
                if v not in seen:
                    seen.append(v)
        
        if seen:
            name = " / ".join(seen)
            base, counter = name, 1
            final_name = name
            while final_name in columns:
                final_name = f"{base} [{get_column_letter(col)}]"
                counter += 1
                if counter > 10:
                    break
            
            # 🔗 НОВОЕ: добавляем маркер объединения в имя
            if col in col_merge_info and col_merge_info[col]["is_merged"]:
                info = col_merge_info[col]
                if info["merge_start"]:
                    final_name = f"🔗 {final_name}"
            
            columns[final_name] = col

    if not columns:
        raise ValueError(f"В диапазоне '{header_range_str}' нет заголовков")

    data_start_row = r2 + 1
    types = {}
    for name, idx in columns.items():
        sample = get_merged_value(ws, data_start_row, idx)
        types[name] = "number" if isinstance(sample, (int, float)) and not isinstance(sample, bool) else "text"

    header_html = build_header_html(r1, c1, r2, c2, merged_ranges)

    return {
        "ws_name": ws.title,
        "header_range": f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}",
        "data_start_row": data_start_row,
        "columns": columns,
        "types": types,
        "merged_ranges": merged_ranges,
        "col_merge_info": col_merge_info,
        "header_html": header_html,
        "header_coords": (r1, c1, r2, c2),
    }


# --- Сайдбар ---
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
                st.session_state.df_data = pd.DataFrame(columns=list(cfg["columns"].keys()))
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


def get_width_for_mode(mode, col_name, col_type):
    if mode == "Компактный": return 80
    elif mode == "Средний": return 150
    elif mode == "Широкий": return 250
    else:
        estimated = max(80, min(400, len(str(col_name)) * 8 + 30))
        if col_type == "number": estimated = max(80, estimated - 30)
        return estimated


# --- Основная область ---
if st.session_state.config:
    cfg = st.session_state.config
    cols = list(cfg["columns"].keys())

    if st.session_state.df_data is None or list(st.session_state.df_data.columns) != cols:
        st.session_state.df_data = pd.DataFrame(columns=cols)

    with st.expander("📐 Структура шаблона (с объединениями)", expanded=False):
        c1, c2, c3 = st.columns(3)
        c1.write(f"**Лист:** `{cfg['ws_name']}`")
        c2.write(f"**Диапазон:** `{cfg['header_range']}`")
        c3.write(f"**Старт данных:** строка {cfg['data_start_row']}")
        
        st.markdown("**📋 Визуализация заголовка:**")
        st.markdown(cfg["header_html"], unsafe_allow_html=True)
        
        if cfg["merged_ranges"]:
            st.info(f"🔗 Найдено **{len(cfg['merged_ranges'])} объединённых диапазонов**.")
        
        st.markdown("**Сопоставление столбцов:**")
        st.dataframe(pd.DataFrame([
            {
                "Имя в калькуляторе": n,
                "Excel": get_column_letter(i),
                "Тип": "🔢 число" if cfg["types"][n] == "number" else "🔤 текст",
                "Объединено": "✅ Да" if n.startswith("🔗") else "—",
            }
            for n, i in cfg["columns"].items()
        ]), hide_index=True, use_container_width=True)

    st.subheader("📝 Заполнение данных")
    
    # 🔗 НОВОЕ: Легенда объединений над редактором
    merged_cols = [n for n in cols if n.startswith("🔗")]
    if merged_cols:
        with st.expander("🔗 Карта объединённых колонок", expanded=True):
            st.caption("Колонки с эмодзи 🔗 объединены в шаблоне. Они сгруппированы вместе для удобства.")
            
            # Группируем по именам (убираем 🔗 и берём первую часть до " / ")
            groups = {}
            for col_name in merged_cols:
                clean_name = col_name.replace("🔗 ", "")
                group_key = clean_name.split(" / ")[0] if " / " in clean_name else clean_name
                groups.setdefault(group_key, []).append(col_name)
            
            for group_name, group_cols in groups.items():
                if len(group_cols) > 1:
                    st.markdown(f"**{group_name}** объединяет: {', '.join([f'`{c}`' for c in group_cols])}")
    
    st.caption("💡 **Совет:** Объединённые колонки помечены 🔗 и идут рядом друг с другом.")

    col_cfg = {}
    for n in cols:
        w = get_width_for_mode(st.session_state.col_width_mode, n, cfg["types"][n])
        if cfg["types"][n] == "number":
            col_cfg[n] = st.column_config.NumberColumn(n, step=0.5, width=w)
        else:
            col_cfg[n] = st.column_config.TextColumn(n, width=w)

    edited_df = st.data_editor(
        st.session_state.df_data,
        column_config=col_cfg,
        column_order=cols,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
    )
    st.session_state.df_data = edited_df

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
                
                # 🔗 Убираем эмодзи 🔗 из имён колонок для маппинга
                col_map = {}
                for name, idx in cfg["columns"].items():
                    clean_name = name.replace("🔗 ", "")
                    col_map[clean_name] = idx

                # Очистка старых данных
                for r in range(start_row, start_row + 1000):
                    all_empty = True
                    for c in col_map.values():
                        cell = ws.cell(row=r, column=c)
                        if cell.value not in (None, "") and not (isinstance(cell.value, str) and cell.value.startswith("=")):
                            all_empty = False
                    if all_empty:
                        break
                    for c in col_map.values():
                        cell = ws.cell(row=r, column=c)
                        if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                            cell.value = None

                # Запись новых
                for i, (_, row) in enumerate(df_clean.iterrows()):
                    for col_name, val in row.items():
                        # Убираем 🔗 из имени
                        clean_col = col_name.replace("🔗 ", "")
                        if clean_col in col_map:
                            cell = ws.cell(row=start_row + i, column=col_map[clean_col])
                            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                                cell.value = None if pd.isna(val) else val

                wb.save(output)
                st.download_button("⬇️ Скачать", output.getvalue(), "journal_filled.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
                st.success(f"✅ Записано {len(df_clean)} строк. Объединения сохранены.")
            except Exception as e:
                st.error(f"Ошибка: {e}")

        st.download_button("💾 JSON", df_clean.to_json(orient="records", force_ascii=False).encode("utf-8"),
                           "project.json", "application/json")
    else:
        st.info("Добавьте строки в таблицу.")

else:
    st.info("👈 Загрузите шаблон и укажите диапазон заголовков (напр. `B5:J6`).")
    st.markdown("""
    ### 🔗 Объединённые ячейки
    
    Приложение автоматически:
    - ✅ Распознаёт объединения в заголовке
    - ✅ Показывает их в превью (голубые ячейки)
    - ✅ Маркирует в редакторе эмодзи 🔗
    - ✅ Сохраняет при экспорте в Excel
    """)
