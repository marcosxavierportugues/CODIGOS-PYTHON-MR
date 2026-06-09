# -*- coding: utf-8 -*-
"""
TABELAS DXF.dxf → Excel (8 abas)
=================================
Lê TABELAS DXF.dxf, extrai as 2 tabelas de aço,
divide cada uma em 4 partes iguais por N° de linhas.
Saída: TABELAS_DXF_Split.xlsx  (8 planilhas)
"""

import re
import sys
import math
import unicodedata
from pathlib import Path
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERRO: pip install openpyxl")
    sys.exit(1)

DXF_FILE = r"C:\Users\PC2\Desktop\TABELAS DXF.dxf"
OUT_FILE  = r"C:\Users\PC2\Desktop\TABELAS_DXF_Split.xlsx"

COLS = ["AÇO", "N", "DIAM", "QUANT", "C.UNIT", "C.TOTAL"]
COL_WIDTHS = [10, 7, 8, 8, 10, 11]

# ─── helpers ──────────────────────────────────────────────────────────────────

def _ascii_norm(text: str) -> str:
    return (
        unicodedata.normalize("NFD", text)
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
        .strip()
    )

HEADER_KW = {"n", "aco", "diam", "qtd", "quant", "c.unit", "c.total",
             "c.unit(cm)", "c.total(cm)"}

def _is_header(t: str) -> bool:
    return _ascii_norm(t) in HEADER_KW

def _clean(raw: str) -> str:
    """Limpa códigos RTF do MTEXT."""
    # Formato: {\fCalibri|b1|i0;TEXT}  → extrair após ';'
    if ";" in raw:
        raw = raw.split(";", 1)[1]
    raw = raw.rstrip("}")
    # Remover outros escapes
    raw = re.sub(r'\{\\[^}]*\}', '', raw)
    raw = re.sub(r'\\[A-Za-z]+[0-9]*;?', '', raw)
    return raw.strip()

# ─── parser ───────────────────────────────────────────────────────────────────

def parse_entities(path: str) -> list[dict]:
    try:
        with open(path, encoding="utf-8-sig") as f:
            lines = f.readlines()
    except UnicodeDecodeError:
        with open(path, encoding="cp1252", errors="replace") as f:
            lines = f.readlines()

    entities = []
    i = 0
    while i < len(lines) - 1:
        etype = lines[i + 1].strip() if lines[i].strip() == "0" else None
        if etype in ("TEXT", "MTEXT"):
            e = {}
            text_parts = []
            i += 2
            while i < len(lines) - 1:
                code = lines[i].strip()
                val  = lines[i + 1].strip()
                if code == "0":
                    break
                try:
                    if code == "10":
                        e["x"] = float(val)
                    elif code == "20":
                        e["y"] = float(val)
                except ValueError:
                    pass
                if code in ("1", "3"):
                    cleaned = _clean(val) if etype == "MTEXT" else val
                    if cleaned:
                        text_parts.append(cleaned)
                i += 2
            if text_parts:
                e["text"] = "".join(text_parts).strip()
            if e.get("text") and "x" in e and "y" in e:
                entities.append(e)
        else:
            i += 1
    return entities

# ─── group by Y ───────────────────────────────────────────────────────────────

def group_by_y(entities: list[dict], tol: float = 0.10) -> dict:
    rows = defaultdict(dict)
    for e in entities:
        y_key = round(e["y"] / tol) * tol
        rows[y_key][round(e["x"], 3)] = e["text"]
    return rows

# ─── detect columns from header ───────────────────────────────────────────────

def find_header(rows: dict):
    for y in sorted(rows.keys(), reverse=True):
        row = rows[y]
        if sum(1 for t in row.values() if _is_header(t)) >= 2:
            merged = dict(row)
            for y2 in rows:
                if y2 == y or abs(y2 - y) > 0.6:
                    continue
                row2 = rows[y2]
                if any(_is_header(t) for t in row2.values()):
                    merged.update(row2)
            return y, merged
    return None, None

def col_ranges(header_row: dict):
    items = sorted(header_row.items())
    xs    = [x for x, _ in items]
    names = [t for _, t in items]
    bounds = []
    for i, x in enumerate(xs):
        lo = x - 1.0      if i == 0              else (xs[i-1] + x) / 2.0
        hi = x + 2.0      if i == len(xs) - 1    else (x + xs[i+1]) / 2.0
        bounds.append((lo, hi))
    return names, bounds

def classify(x: float, names, bounds) -> str | None:
    for col, (lo, hi) in zip(names, bounds):
        if lo <= x < hi:
            return col
    return None

# ─── build one table ──────────────────────────────────────────────────────────

def build_table(entities: list[dict]) -> dict:
    rows = group_by_y(entities)
    header_y, hrow = find_header(rows)
    if header_y is None:
        raise ValueError("Cabeçalho não encontrado.")

    names, bounds = col_ranges(hrow)

    all_above = [y for y in rows if y > header_y]
    max_y = max(all_above, default=None)

    title    = ""
    sections = []
    data_raw = defaultdict(dict)  # y → {col: val}

    for y in sorted(rows.keys(), reverse=True):
        row = rows[y]
        if y > header_y:
            for _, t in sorted(row.items()):
                if _is_header(t):
                    continue
                if max_y and abs(y - max_y) < 0.02:
                    title = title or t
                else:
                    sections.append((y, t))
        elif abs(y - header_y) < 0.1:
            continue
        else:
            vals = list(row.values())
            if all(v in ("(mm)", "(cm)", "mm", "cm", "") for v in vals):
                continue
            for x, t in row.items():
                col = classify(x, names, bounds)
                if col:
                    data_raw[y][col] = t

    # Remove empty rows
    data_raw = {y: d for y, d in data_raw.items()
                if any(v not in ("", None) for v in d.values())}

    # ── Merge split rows ───────────────────────────────────────────────────────
    # Algumas linhas do DXF têm as colunas em 2 Y ligeiramente diferentes.
    # Para cada linha incompleta (sem QUANT ou sem AÇO/N), procura a linha
    # mais próxima dentro de 0.25 unidades e mescla as colunas.
    LEFT_COLS  = set(names[:3]) if len(names) >= 3 else set(names)
    RIGHT_COLS = set(names[3:]) if len(names) > 3  else set()

    all_ys = sorted(data_raw.keys(), reverse=True)

    merged_into: set = set()
    for yk in list(all_ys):
        if yk in merged_into:
            continue
        d = data_raw[yk]
        has_left  = any(d.get(c) for c in LEFT_COLS)
        has_right = any(d.get(c) for c in RIGHT_COLS)

        if has_left and not has_right:
            # procurar parceiro próximo que tenha colunas direitas
            for yk2 in all_ys:
                if yk2 == yk or yk2 in merged_into:
                    continue
                if abs(yk2 - yk) > 0.25:
                    continue
                d2 = data_raw[yk2]
                has_right2 = any(d2.get(c) for c in RIGHT_COLS)
                has_left2  = any(d2.get(c) for c in LEFT_COLS)
                if has_right2 and not has_left2:
                    d.update(d2)
                    merged_into.add(yk2)
                    break

        elif has_right and not has_left:
            # procurar parceiro próximo que tenha colunas esquerdas
            for yk2 in all_ys:
                if yk2 == yk or yk2 in merged_into:
                    continue
                if abs(yk2 - yk) > 0.25:
                    continue
                d2 = data_raw[yk2]
                has_left2  = any(d2.get(c) for c in LEFT_COLS)
                has_right2 = any(d2.get(c) for c in RIGHT_COLS)
                if has_left2 and not has_right2:
                    d.update(d2)
                    merged_into.add(yk2)
                    break

    # Remover as linhas que foram absorvidas
    for yk in merged_into:
        data_raw.pop(yk, None)

    sorted_y = sorted(data_raw.keys(), reverse=True)

    # Deduplicate sections
    seen, uniq = set(), []
    for y, t in sorted(sections, key=lambda s: -s[0]):
        if t not in seen and not _is_header(t):
            seen.add(t); uniq.append((y, t))

    return {
        "title":    title or "RELAÇÃO DO AÇO",
        "sections": uniq,
        "names":    names,
        "sorted_y": sorted_y,
        "data":     data_raw,
    }

# ─── to_number ────────────────────────────────────────────────────────────────

def to_number(val):
    if not isinstance(val, str) or val.strip() in ("VAR", ""):
        return val
    try:
        f = float(val.replace(",", "."))
        return int(f) if f == int(f) else f
    except (ValueError, OverflowError):
        return val

# ─── Excel writer ─────────────────────────────────────────────────────────────

thin   = None  # init in write_all
medium = None

def _styles():
    t = Side(style="thin")
    m = Side(style="medium")
    bdr   = Border(left=t, right=t, top=t, bottom=t)
    bdr_m = Border(left=m, right=m, top=m, bottom=m)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    f_title   = (Font(bold=True, name="Calibri", size=12, color="FFFFFF"),
                 PatternFill("solid", fgColor="1F3864"))
    f_section = (Font(bold=True, name="Calibri", size=10, italic=True),
                 PatternFill("solid", fgColor="BDD7EE"))
    f_header  = (Font(bold=True, name="Calibri", size=10),
                 PatternFill("solid", fgColor="D9D9D9"))
    f_normal  = Font(name="Calibri", size=10)
    f_alt     = PatternFill("solid", fgColor="EBF3FB")
    return bdr, bdr_m, center, f_title, f_section, f_header, f_normal, f_alt


def write_sheet(ws, table: dict, data_slice: list, part_num: int, total_parts: int):
    """Escreve uma parte da tabela em ws."""
    bdr, bdr_m, center, f_title, f_section, f_header, f_normal, f_alt = _styles()

    names   = table["names"]
    n_cols  = len(names)
    data    = table["data"]

    # Column widths
    for i in range(n_cols):
        w = COL_WIDTHS[i] if i < len(COL_WIDTHS) else 10
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    def merge_row(r, text, font, fill, height=18, border=None):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
        c = ws.cell(row=r, column=1, value=text)
        c.font = font; c.alignment = center; c.fill = fill
        b = border or bdr
        for col in range(1, n_cols + 1):
            ws.cell(row=r, column=col).border = b
        ws.row_dimensions[r].height = height

    r = 1
    title_text = f"{table['title']}  —  Parte {part_num}/{total_parts}"
    merge_row(r, title_text, f_title[0], f_title[1], height=22, border=bdr_m)
    r += 1

    for _, label in table["sections"]:
        merge_row(r, label, f_section[0], f_section[1])
        r += 1

    # Column headers
    for i, col in enumerate(names):
        c = ws.cell(row=r, column=i + 1, value=col)
        c.font = f_header[0]; c.alignment = center
        c.fill = f_header[1]; c.border = bdr
    ws.row_dimensions[r].height = 16
    r += 1

    # Data
    first_col  = names[0]
    prev_first = ""

    for idx, y_key in enumerate(data_slice):
        row_data = data[y_key]
        fill = f_alt if idx % 2 == 1 else None

        val0 = row_data.get(first_col, "")
        if val0:
            prev_first = val0
        elif _ascii_norm(first_col) == "aco":
            val0 = prev_first

        for i, col in enumerate(names):
            raw = val0 if i == 0 else row_data.get(col, "")
            val = to_number(raw)
            c = ws.cell(row=r, column=i + 1, value=val)
            c.font = f_normal; c.alignment = center; c.border = bdr
            if fill:
                c.fill = fill
        r += 1


def write_all(tables: list[dict], out_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    for t_idx, table in enumerate(tables, start=1):
        sorted_y = table["sorted_y"]
        n = len(sorted_y)
        n_parts = 4
        part_size = math.ceil(n / n_parts)

        for p in range(n_parts):
            start = p * part_size
            end   = min(start + part_size, n)
            if start >= n:
                break
            data_slice = sorted_y[start:end]

            sheet_name = f"Tab{t_idx}_Parte{p+1}"
            ws = wb.create_sheet(title=sheet_name)
            write_sheet(ws, table, data_slice, p + 1, n_parts)
            print(f"  {sheet_name}: {len(data_slice)} linhas")

    wb.save(out_path)
    print(f"\nSalvo: {out_path}")


# ─── main ─────────────────────────────────────────────────────────────────────

def main():
    print(f"Lendo: {DXF_FILE}")
    entities = parse_entities(DXF_FILE)
    print(f"  Entidades: {len(entities)}")

    # Split by X coordinate: Table 1 (X < 100), Table 2 (X > 100)
    ent1 = [e for e in entities if e["x"] < 100]
    ent2 = [e for e in entities if e["x"] > 100]
    print(f"  Tabela 1: {len(ent1)} entidades")
    print(f"  Tabela 2: {len(ent2)} entidades")

    table1 = build_table(ent1)
    table2 = build_table(ent2)

    print(f"  Tabela 1: {table1['title']!r}  |  {len(table1['sorted_y'])} linhas")
    print(f"  Tabela 2: {table2['title']!r}  |  {len(table2['sorted_y'])} linhas")

    write_all([table1, table2], OUT_FILE)


if __name__ == "__main__":
    main()
