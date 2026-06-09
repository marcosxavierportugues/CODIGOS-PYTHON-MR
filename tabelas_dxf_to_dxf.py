# -*- coding: utf-8 -*-
"""
TABELAS DXF.dxf  →  TABELAS_Split.dxf  (8 tabelas)
=====================================================
Extrai as 2 tabelas de aço, divide cada uma em 4 partes,
desenha as 8 tabelas em grid no DXF de saída.
"""

import re
import sys
import math
import unicodedata
from pathlib import Path
from collections import defaultdict

try:
    import ezdxf
    from ezdxf import colors
    from ezdxf.enums import TextEntityAlignment
except ImportError:
    print("ERRO: pip install ezdxf")
    sys.exit(1)

DXF_FILE = r"C:\Users\PC2\Desktop\TABELAS DXF.dxf"
OUT_FILE  = r"C:\Users\PC2\Desktop\TABELAS_Split_v2.dxf"

# ── Dimensões (unidades DXF) ─────────────────────────────────────────────────
COL_NAMES  = ["AÇO",  "N",   "DIAM", "QUANT", "C.UNIT", "C.TOTAL"]

ROW_H      = 0.60
TITLE_H    = 1.00
SECTION_H  = 0.75
HEADER_H   = 0.75
TEXT_H     = ROW_H * 0.58   # ~60 % da altura da linha
CHAR_W     = TEXT_H * 0.70  # fator empírico Calibri
PADDING    = TEXT_H * 0.80  # margem lateral por célula

GAP_X      = 3.0    # espaço entre tabelas na horizontal
GAP_Y      = 8.0    # espaço entre linhas de tabelas

# COL_WIDTHS e TOTAL_W serão calculados dinamicamente em main()
COL_WIDTHS: list = []
TOTAL_W: float   = 0.0


def compute_col_widths(tables: list, names: list) -> list:
    """Largura de cada coluna = maior conteúdo real, mínimo = cabeçalho."""
    max_len = {col: len(col) for col in names}
    for tbl in tables:
        for row_data in tbl["data"].values():
            for col, raw in row_data.items():
                if col not in max_len:
                    continue
                val = to_number(raw)
                s = str(val) if val is not None else ""
                if len(s) > max_len[col]:
                    max_len[col] = len(s)
    return [max_len.get(col, 4) * CHAR_W + 2 * PADDING for col in names]

# ── Helpers de parsing ────────────────────────────────────────────────────────

def _ascii_norm(t: str) -> str:
    return (unicodedata.normalize("NFD", t)
            .encode("ascii", "ignore").decode("ascii")
            .lower().strip())

HEADER_KW = {"n","aco","diam","qtd","quant","c.unit","c.total",
             "c.unit(cm)","c.total(cm)"}

def _is_header(t: str) -> bool:
    return _ascii_norm(t) in HEADER_KW

def _clean(raw: str) -> str:
    if ";" in raw:
        raw = raw.split(";", 1)[1]
    raw = raw.rstrip("}")
    raw = re.sub(r'\{\\[^}]*\}', '', raw)
    raw = re.sub(r'\\[A-Za-z]+[0-9]*;?', '', raw)
    return raw.strip()

def parse_entities(path: str) -> list:
    try:
        with open(path, encoding="utf-8-sig") as f:
            lines = f.readlines()
    except UnicodeDecodeError:
        with open(path, encoding="cp1252", errors="replace") as f:
            lines = f.readlines()

    out = []
    i = 0
    while i < len(lines) - 1:
        etype = lines[i + 1].strip() if lines[i].strip() == "0" else None
        if etype in ("TEXT", "MTEXT"):
            e = {}
            parts = []
            i += 2
            while i < len(lines) - 1:
                code = lines[i].strip()
                val  = lines[i + 1].strip()
                if code == "0":
                    break
                try:
                    if code == "10": e["x"] = float(val)
                    elif code == "20": e["y"] = float(val)
                except ValueError:
                    pass
                if code in ("1", "3"):
                    c = _clean(val) if etype == "MTEXT" else val
                    if c: parts.append(c)
                i += 2
            if parts:
                e["text"] = "".join(parts).strip()
            if e.get("text") and "x" in e and "y" in e:
                out.append(e)
        else:
            i += 1
    return out

def group_by_y(entities, tol=0.10):
    rows = defaultdict(dict)
    for e in entities:
        yk = round(e["y"] / tol) * tol
        rows[yk][round(e["x"], 3)] = e["text"]
    return rows

def find_header(rows):
    for y in sorted(rows.keys(), reverse=True):
        row = rows[y]
        if sum(1 for t in row.values() if _is_header(t)) >= 2:
            merged = dict(row)
            for y2 in rows:
                if y2 == y or abs(y2 - y) > 0.6: continue
                row2 = rows[y2]
                if any(_is_header(t) for t in row2.values()):
                    merged.update(row2)
            return y, merged
    return None, None

def col_ranges(hrow):
    items = sorted(hrow.items())
    xs = [x for x, _ in items]
    ns = [t for _, t in items]
    bds = []
    for i, x in enumerate(xs):
        lo = x - 1.0  if i == 0           else (xs[i-1]+x)/2.0
        hi = x + 2.0  if i == len(xs)-1   else (x+xs[i+1])/2.0
        bds.append((lo, hi))
    return ns, bds

def classify(x, names, bounds):
    for col, (lo, hi) in zip(names, bounds):
        if lo <= x < hi: return col
    return None

def build_table(entities):
    rows = group_by_y(entities)
    hy, hrow = find_header(rows)
    if hy is None:
        raise ValueError("Cabecalho nao encontrado")
    names, bounds = col_ranges(hrow)
    all_above = [y for y in rows if y > hy]
    max_y = max(all_above, default=None)
    title = ""
    sections = []
    data_raw = defaultdict(dict)

    for y in sorted(rows.keys(), reverse=True):
        row = rows[y]
        if y > hy:
            for _, t in sorted(row.items()):
                if _is_header(t): continue
                if max_y and abs(y - max_y) < 0.02:
                    title = title or t
                else:
                    sections.append((y, t))
        elif abs(y - hy) < 0.1:
            continue
        else:
            vals = list(row.values())
            if all(v in ("(mm)","(cm)","mm","cm","") for v in vals): continue
            for x, t in row.items():
                col = classify(x, names, bounds)
                if col: data_raw[y][col] = t

    data_raw = {y: d for y, d in data_raw.items()
                if any(v not in ("", None) for v in d.values())}

    # Merge split rows (colunas em Y ligeiramente diferentes)
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
            for yk2 in all_ys:
                if yk2 == yk or yk2 in merged_into or abs(yk2 - yk) > 0.25:
                    continue
                d2 = data_raw[yk2]
                if any(d2.get(c) for c in RIGHT_COLS) and not any(d2.get(c) for c in LEFT_COLS):
                    d.update(d2); merged_into.add(yk2); break
        elif has_right and not has_left:
            for yk2 in all_ys:
                if yk2 == yk or yk2 in merged_into or abs(yk2 - yk) > 0.25:
                    continue
                d2 = data_raw[yk2]
                if any(d2.get(c) for c in LEFT_COLS) and not any(d2.get(c) for c in RIGHT_COLS):
                    d.update(d2); merged_into.add(yk2); break
    for yk in merged_into:
        data_raw.pop(yk, None)

    sorted_y = sorted(data_raw.keys(), reverse=True)

    seen, uniq = set(), []
    for y, t in sorted(sections, key=lambda s: -s[0]):
        if t not in seen and not _is_header(t):
            seen.add(t); uniq.append((y, t))

    return {"title": title or "RELAÇÃO DO AÇO",
            "sections": uniq,
            "names": names,
            "sorted_y": sorted_y,
            "data": data_raw}

def to_number(val):
    if not isinstance(val, str) or val.strip() in ("VAR", ""):
        return val
    try:
        f = float(val.replace(",", "."))
        return int(f) if f == int(f) else f
    except (ValueError, OverflowError):
        return val

# ── DXF drawing helpers ───────────────────────────────────────────────────────

def _rect(msp, x, y, w, h, layer="0"):
    """Desenha retângulo (y = topo, cresce para baixo)."""
    pts = [(x, y), (x+w, y), (x+w, y-h), (x, y-h), (x, y)]
    for i in range(4):
        msp.add_line(pts[i], pts[i+1], dxfattribs={"layer": layer})

def _hline(msp, x, y, w, layer="0"):
    msp.add_line((x, y), (x+w, y), dxfattribs={"layer": layer})

def _vline(msp, x, y1, y2, layer="0"):
    msp.add_line((x, y1), (x, y2), dxfattribs={"layer": layer})

def _txt(msp, text, x, y, h, layer="0", bold=False):
    """Adiciona TEXT centralizado na célula (x, y = centro da célula)."""
    style = "BOLD" if bold else "Standard"
    t = msp.add_text(
        str(text),
        dxfattribs={"height": h, "layer": layer, "style": style},
    )
    t.set_placement((x, y), align=TextEntityAlignment.MIDDLE_CENTER)

# ── Draw one table part ───────────────────────────────────────────────────────

def draw_part(msp, table, data_slice, part_num, total_parts, ox, oy):
    """
    Desenha uma parte da tabela.
    ox, oy = canto superior-esquerdo.
    Retorna Y do fundo da tabela.
    """
    data  = table["data"]
    names = table["names"]

    # ── Título ────────────────────────────────────────────────────────────────
    title_text = "{} - Parte {}/{}".format(table["title"], part_num, total_parts)
    _rect(msp, ox, oy, TOTAL_W, TITLE_H, "TITULO")
    _txt(msp, title_text, ox + TOTAL_W/2, oy - TITLE_H/2, TEXT_H*1.1, "TITULO", bold=True)
    oy -= TITLE_H

    # ── Seções ────────────────────────────────────────────────────────────────
    for _, sec in table["sections"]:
        _rect(msp, ox, oy, TOTAL_W, SECTION_H, "SECAO")
        _txt(msp, sec, ox + TOTAL_W/2, oy - SECTION_H/2, TEXT_H, "SECAO", bold=True)
        oy -= SECTION_H

    # ── Cabeçalho ────────────────────────────────────────────────────────────
    x = ox
    for cname, cw in zip(COL_NAMES, COL_WIDTHS):
        _rect(msp, x, oy, cw, HEADER_H, "CABECALHO")
        _txt(msp, cname, x + cw/2, oy - HEADER_H/2, TEXT_H, "CABECALHO", bold=True)
        x += cw
    # Linha superior e inferior do cabeçalho já estão; adicionar bordas internas verticais
    oy -= HEADER_H

    # ── Dados ─────────────────────────────────────────────────────────────────
    first_col  = names[0]
    prev_first = ""

    for idx, y_key in enumerate(data_slice):
        row_data = data[y_key]
        layer = "DADOS_ALT" if idx % 2 == 1 else "DADOS"

        val0 = row_data.get(first_col, "")
        if val0:
            prev_first = val0
        else:
            val0 = prev_first

        x = ox
        for i, (cname, cw) in enumerate(zip(COL_NAMES, COL_WIDTHS)):
            raw = val0 if i == 0 else row_data.get(cname, "")
            val = to_number(raw)
            val_str = "" if val is None else str(val)

            _rect(msp, x, oy, cw, ROW_H, layer)
            _txt(msp, val_str, x + cw/2, oy - ROW_H/2, TEXT_H*0.9, layer)
            x += cw

        oy -= ROW_H

    return oy  # Y do fundo

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("Lendo:", DXF_FILE)
    entities = parse_entities(DXF_FILE)
    print(f"  Entidades: {len(entities)}")

    ent1 = [e for e in entities if e["x"] < 100]
    ent2 = [e for e in entities if e["x"] > 100]

    table1 = build_table(ent1)
    table2 = build_table(ent2)

    tables = [table1, table2]

    # Calcular larguras de coluna baseadas no conteúdo real
    global COL_WIDTHS, TOTAL_W
    names = table1["names"]  # mesmas colunas nas 2 tabelas
    COL_WIDTHS = compute_col_widths(tables, names)
    TOTAL_W    = sum(COL_WIDTHS)
    print(f"  Colunas: {[f'{n}={w:.2f}' for n, w in zip(names, COL_WIDTHS)]}")
    print(f"  Largura total: {TOTAL_W:.2f}")

    # Criar DXF
    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 6   # milímetros

    # Estilos de texto
    styles = doc.styles
    try:
        styles.add("BOLD", font="calibrib.ttf")
    except Exception:
        pass

    # Layers
    layer_defs = {
        "TITULO":      colors.WHITE,
        "SECAO":       colors.CYAN,
        "CABECALHO":   colors.YELLOW,
        "DADOS":       colors.WHITE,
        "DADOS_ALT":   7,   # cinza
        "0":           colors.WHITE,
    }
    for lname, col in layer_defs.items():
        try:
            doc.layers.add(lname, color=col)
        except Exception:
            pass

    msp = doc.modelspace()

    # Layout: 4 colunas × 2 linhas
    # Linha 1: Tab1_Parte1..4   (Tabela 1)
    # Linha 2: Tab2_Parte1..4   (Tabela 2)

    n_parts = 4

    for t_idx, table in enumerate(tables, start=1):
        sorted_y = table["sorted_y"]
        n = len(sorted_y)
        part_size = math.ceil(n / n_parts)

        # Calcular altura máxima desta tabela para o offset da linha
        max_rows = part_size
        table_h = (TITLE_H
                   + len(table["sections"]) * SECTION_H
                   + HEADER_H
                   + max_rows * ROW_H)

        row_oy = 0 if t_idx == 1 else -(table_h + GAP_Y)
        # Para tabela 2, usamos base Y de linha 1 menos altura - gap
        # Calculado dinamicamente:
        if t_idx == 2:
            # pegar altura da tabela 1
            n1 = len(table1["sorted_y"])
            ps1 = math.ceil(n1 / n_parts)
            h1 = (TITLE_H
                  + len(table1["sections"]) * SECTION_H
                  + HEADER_H
                  + ps1 * ROW_H)
            row_oy = -(h1 + GAP_Y)

        for p in range(n_parts):
            start = p * part_size
            end   = min(start + part_size, n)
            if start >= n: break
            data_slice = sorted_y[start:end]

            ox = p * (TOTAL_W + GAP_X)
            oy = row_oy

            bottom = draw_part(msp, table, data_slice,
                               p + 1, n_parts, ox, oy)

            label = "Tab{}_Parte{}".format(t_idx, p + 1)
            print(f"  {label}: {len(data_slice)} linhas  ({ox:.1f}, {oy:.1f})")

    doc.saveas(OUT_FILE)
    print(f"\nSalvo: {OUT_FILE}")


if __name__ == "__main__":
    main()
