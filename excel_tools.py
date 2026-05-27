"""Вспомогательные инструменты для работы с Excel-шаблонами."""
import re
import io
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import MergedCell


# ==================== Парсинг диапазонов ====================

def parse_range(range_str):
    """Парсит диапазон вида 'B5:J6' или 'B5'."""
    s = range_str.strip().upper().replace(" ", "")
    if not s:
        raise ValueError("Диапазон пустой")
    if ":" not in s:
        s = f"{s}:{s}"
    m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", s)
    if not m:
        raise ValueError("Формат: 'B5:J6' (например)")
    c1 = column_index_from_string(m.group(1))
    r1 = int(m.group(2))
    c2 = column_index_from_string(m.group(3))
    r2 = int(m.group(4))
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


# ==================== Работа с merged cells ====================

def get_merged_value(ws, row, col):
    """Возвращает значение ячейки, учитывая объединения."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                return ws.cell(row=rng.min_row, column=rng.min_col).value
    return cell.value


def get_merged_ranges_in_range(ws, r1, c1, r2, c2):
    """Находит все merged-диапазоны, пересекающиеся с указанным."""
    result = []
    for rng in ws.merged_cells.ranges:
        if (rng.min_row <= r2 and rng.max_row >= r1 and
                rng.min_col <= c2 and rng.max_col >= c1):
            local_r1, local_c1 = max(rng.min_row, r1), max(rng.min_col, c1)
            local_r2, local_c2 = min(rng.max_row, r2), min(rng.max_col, c2)
            value = ws.cell(row=rng.min_row, column=rng.min_col).value
            result.append({
                "r1": local_r1, "c1": local_c1,
                "r2": local_r2, "c2": local_c2,
                "rowspan": local_r2 - local_r1 + 1,
                "colspan": local_c2 - local_c1 + 1,
                "value": value,
            })
    return result


# ==================== HTML-превью заголовка ====================

def build_header_html(r1, c1, r2, c2, merged_ranges):
    """Строит HTML-таблицу с визуализацией объединений."""
    nrows, ncols = r2 - r1 + 1, c2 - c1 + 1
    covered = [[False] * ncols for _ in range(nrows)]

    html = '<table style="border-collapse: collapse; margin: 10px 0; font-family: monospace;">'
    for ri in range(nrows):
        html += "<tr>"
        for ci in range(ncols):
            if covered[ri][ci]:
                continue
            abs_r, abs_c = r1 + ri, c1 + ci
            found = next((mr for mr in merged_ranges if mr["r1"] == abs_r and mr["c1"] == abs_c), None)

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


# ==================== Парсинг шаблона ====================

def parse_template(file_bytes, sheet_name, header_range_str):
    """Разбирает Excel-шаблон и возвращает конфигурацию."""
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name]
    r1, c1, r2, c2 = parse_range(header_range_str)

    merged_ranges = get_merged_ranges_in_range(ws, r1, c1, r2, c2)

    def top_left_for(row, col):
        for mr in merged_ranges:
            if mr["r1"] <= row <= mr["r2"] and mr["c1"] <= col <= mr["c2"]:
                return mr["r1"], mr["c1"]
        return row, col

    columns = {}
    for col in range(c1, c2 + 1):
        parts_by_group = {}
        for r in range(r1, r2 + 1):
            tl = top_left_for(r, col)
            val = get_merged_value(ws, r, col)
            if val not in (None, ""):
                parts_by_group.setdefault(tl, []).append(str(val).strip().replace("\n", " "))

        seen = []
        for tl in sorted(parts_by_group.keys()):
            if parts_by_group[tl]:
                v = parts_by_group[tl][0]
                if v not in seen:
                    seen.append(v)

        if seen:
            name = " / ".join(seen)
            final_name, counter = name, 1
            while final_name in columns:
                final_name = f"{name} [{get_column_letter(col)}]"
                counter += 1
                if counter > 10:
                    break
            columns[final_name] = col

    if not columns:
        raise ValueError(f"В диапазоне '{header_range_str}' нет заголовков")

    data_start_row = r2 + 1
    types = {}
    for name, idx in columns.items():
        sample = get_merged_value(ws, data_start_row, idx)
        types[name] = "number" if isinstance(sample, (int, float)) and not isinstance(sample, bool) else "text"

    return {
        "ws_name": ws.title,
        "header_range": f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}",
        "data_start_row": data_start_row,
        "columns": columns,
        "types": types,
        "merged_ranges": merged_ranges,
        "header_html": build_header_html(r1, c1, r2, c2, merged_ranges),
    }


# ==================== Экспорт в Excel ====================

def export_to_excel(template_bytes, ws_name, data_start_row, columns_map, df):
    """Записывает DataFrame в Excel-шаблон, сохраняя формулы и стили."""
    output = io.BytesIO()
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb[ws_name]

    # Очистка старых данных
    for r in range(data_start_row, data_start_row + 2000):
        all_empty = True
        for c in columns_map.values():
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v not in (None, "") and not (isinstance(v, str) and v.startswith("=")):
                all_empty = False
        if all_empty:
            break
        for c in columns_map.values():
            cell = ws.cell(row=r, column=c)
            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                cell.value = None

    # Запись новых данных
    for i, (_, row) in enumerate(df.iterrows()):
        for col_name, val in row.items():
            if col_name in columns_map:
                cell = ws.cell(row=data_start_row + i, column=columns_map[col_name])
                if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                    try:
                        cell.value = None if _is_empty(val) else _to_python(val)
                    except Exception:
                        cell.value = None

    wb.save(output)
    return output.getvalue()


def _is_empty(val):
    """Проверяет, пустое ли значение."""
    if val is None or val == "":
        return True
    try:
        import pandas as pd
        return pd.isna(val)
    except Exception:
        return False


def _to_python(val):
    """Конвертирует numpy-типы в Python для openpyxl."""
    import numpy as np
    if isinstance(val, (np.integer,)):
        return int(val)
    if isinstance(val, (np.floating,)):
        return float(val)
    if isinstance(val, (np.bool_,)):
        return bool(val)
    return val
