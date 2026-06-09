# -*- coding: utf-8 -*-
"""
TABELAS DXF -> DXF  |  Relação do Aço  |  Marcos Roldão
=========================================================
Lê um arquivo DXF com 2 tabelas de armação, divide cada uma
em 4 partes e desenha as 8 tabelas em grid no DXF de saída.

Uso:
    python tabelas_dxf_para_dxf.py                        # abre janela
    python tabelas_dxf_para_dxf.py entrada.dxf            # saída automática
    python tabelas_dxf_para_dxf.py entrada.dxf saida.dxf  # explícito

Requisito:
    pip install ezdxf
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

# ─── Proporções do desenho ────────────────────────────────────────────────────

ROW_H     = 0.60                 # altura de cada linha de dado
TITLE_H   = 1.00                 # altura da linha de título
SECTION_H = 0.75                 # altura da linha de seção
HEADER_H  = 0.75                 # altura do cabeçalho
TEXT_H    = ROW_H * 0.58         # altura do texto (~60 % do ROW_H)
CHAR_W    = TEXT_H * 0.70        # largura média por caractere (Calibri)
PADDING   = TEXT_H * 0.80        # margem lateral por célula
GAP_X     = 3.0                  # espaço horizontal entre tabelas
GAP_Y     = 8.0                  # espaço vertical entre grupos de tabelas
N_PARTS   = 4                    # número de partes por tabela

# ─── Keywords de cabeçalho ────────────────────────────────────────────────────

HEADER_KW = {
    "n", "aco", "diam", "qtd", "quant",
    "c.unit", "c.total", "c.unit(cm)", "c.total(cm)",
}

# ─────────────────────────────────────────────────────────────────────────────
#  PARSING DXF
# ─────────────────────────────────────────────────────────────────────────────

def _ascii_norm(text: str) -> str:
    """Remove acentos, minúscula, strip — para comparação robusta."""
    return (
        unicodedata.normalize("NFD", text)
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
        .strip()
    )


def _is_header(text: str) -> bool:
    return _ascii_norm(text) in HEADER_KW


def _clean_mtext(raw: str) -> str:
    """
    Limpa códigos RTF do MTEXT.
    Formato do AutoCAD: {\\fCalibri|b1|i0;TEXTO}
    Extrai o texto após o ';' e remove sobras.
    """
    if ";" in raw:
        raw = raw.split(";", 1)[1]
    raw = raw.rstrip("}")
    raw = re.sub(r'\{\\[^}]*\}', '', raw)   # blocos restantes
    raw = re.sub(r'\\[A-Za-z]+[0-9]*;?', '', raw)  # escapes avulsos
    return raw.strip()


def parse_entities(path: str) -> list[dict]:
    """
    Extrai entidades TEXT e MTEXT do DXF.
    Retorna lista de {"x": float, "y": float, "text": str}.
    Tenta UTF-8-BOM primeiro; fallback para cp1252.
    """
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
            e: dict = {}
            parts: list[str] = []
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
                    cleaned = _clean_mtext(val) if etype == "MTEXT" else val
                    if cleaned:
                        parts.append(cleaned)

                i += 2

            if parts:
                e["text"] = "".join(parts).strip()

            if e.get("text") and "x" in e and "y" in e:
                out.append(e)
        else:
            i += 1

    return out


# ─────────────────────────────────────────────────────────────────────────────
#  ESTRUTURA DA TABELA
# ─────────────────────────────────────────────────────────────────────────────

def group_by_y(entities: list[dict], tol: float = 0.10) -> dict:
    """
    Agrupa entidades por Y com tolerância.
    tol=0.10 une as "meias linhas" que o AutoCAD coloca 0.05–0.10 apart.
    """
    rows: dict = defaultdict(dict)
    for e in entities:
        yk = round(e["y"] / tol) * tol
        rows[yk][round(e["x"], 3)] = e["text"]
    return rows


def find_header(rows: dict) -> tuple:
    """Localiza a linha de cabeçalho (≥2 keywords). Mescla linhas próximas."""
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


def col_ranges(hrow: dict) -> tuple[list, list]:
    """Mapa de colunas: nomes + intervalos X baseados nos X do cabeçalho."""
    items = sorted(hrow.items())
    xs    = [x for x, _ in items]
    names = [t for _, t in items]
    bounds = []
    for i, x in enumerate(xs):
        lo = x - 1.0  if i == 0            else (xs[i - 1] + x) / 2.0
        hi = x + 2.0  if i == len(xs) - 1  else (x + xs[i + 1]) / 2.0
        bounds.append((lo, hi))
    return names, bounds


def classify(x: float, names: list, bounds: list) -> str | None:
    for col, (lo, hi) in zip(names, bounds):
        if lo <= x < hi:
            return col
    return None


def merge_split_rows(data_raw: dict, names: list) -> dict:
    """
    Algumas linhas do DXF têm colunas esquerda (AÇO/N/DIAM) e direita
    (QUANT/C.UNIT/C.TOTAL) em Y ligeiramente diferentes (gap ~0.10).
    Une cada "meia linha" com sua parceira mais próxima (≤0.25 unidades).
    """
    left_cols  = set(names[:3]) if len(names) >= 3 else set(names)
    right_cols = set(names[3:]) if len(names) > 3  else set()

    all_ys      = sorted(data_raw.keys(), reverse=True)
    merged_into: set = set()

    for yk in list(all_ys):
        if yk in merged_into:
            continue
        d = data_raw[yk]
        has_left  = any(d.get(c) for c in left_cols)
        has_right = any(d.get(c) for c in right_cols)

        search_left  = has_right and not has_left
        search_right = has_left  and not has_right

        if not search_left and not search_right:
            continue

        for yk2 in all_ys:
            if yk2 == yk or yk2 in merged_into or abs(yk2 - yk) > 0.25:
                continue
            d2 = data_raw[yk2]
            partner_has_left  = any(d2.get(c) for c in left_cols)
            partner_has_right = any(d2.get(c) for c in right_cols)

            if search_right and partner_has_right and not partner_has_left:
                d.update(d2)
                merged_into.add(yk2)
                break
            if search_left and partner_has_left and not partner_has_right:
                d.update(d2)
                merged_into.add(yk2)
                break

    for yk in merged_into:
        data_raw.pop(yk, None)

    return data_raw


def build_table(entities: list[dict]) -> dict:
    rows = group_by_y(entities)
    hy, hrow = find_header(rows)

    if hy is None:
        raise ValueError("Cabeçalho não encontrado.")

    names, bounds = col_ranges(hrow)
    all_above = [y for y in rows if y > hy]
    max_y     = max(all_above, default=None)

    title     = ""
    sections: list = []
    data_raw: dict = defaultdict(dict)

    for y in sorted(rows.keys(), reverse=True):
        row = rows[y]

        if y > hy:
            for _, t in sorted(row.items()):
                if _is_header(t):
                    continue
                if max_y and abs(y - max_y) < 0.02:
                    title = title or t
                else:
                    sections.append((y, t))

        elif abs(y - hy) < 0.1:
            continue  # linha de cabeçalho em si

        else:
            vals = list(row.values())
            if all(v in ("(mm)", "(cm)", "mm", "cm", "") for v in vals):
                continue  # linha de unidades
            for x, t in row.items():
                col = classify(x, names, bounds)
                if col:
                    data_raw[y][col] = t

    # Remove linhas vazias
    data_raw = {
        y: d for y, d in data_raw.items()
        if any(v not in ("", None) for v in d.values())
    }

    # Mescla linhas partidas
    data_raw = merge_split_rows(data_raw, names)

    sorted_y = sorted(data_raw.keys(), reverse=True)

    # Seções únicas
    seen: set = set()
    uniq_sections: list = []
    for y, t in sorted(sections, key=lambda s: -s[0]):
        if t not in seen and not _is_header(t):
            seen.add(t)
            uniq_sections.append((y, t))

    return {
        "title":    title or "RELAÇÃO DO AÇO",
        "sections": uniq_sections,
        "names":    names,
        "sorted_y": sorted_y,
        "data":     data_raw,
    }


# ─────────────────────────────────────────────────────────────────────────────
#  CONVERSÃO DE VALORES
# ─────────────────────────────────────────────────────────────────────────────

def to_number(val):
    """Converte string para int/float quando possível. Mantém 'VAR' e ''."""
    if not isinstance(val, str) or val.strip() in ("VAR", ""):
        return val
    try:
        f = float(val.replace(",", "."))
        return int(f) if f == int(f) else f
    except (ValueError, OverflowError):
        return val


# ─────────────────────────────────────────────────────────────────────────────
#  LARGURAS DE COLUNA DINÂMICAS
# ─────────────────────────────────────────────────────────────────────────────

def compute_col_widths(tables: list, names: list) -> list[float]:
    """
    Calcula a largura de cada coluna pelo conteúdo real.
    Mínimo = tamanho do nome do cabeçalho.
    """
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


# ─────────────────────────────────────────────────────────────────────────────
#  DESENHO NO DXF
# ─────────────────────────────────────────────────────────────────────────────

def _rect(msp, x: float, y: float, w: float, h: float, layer: str = "0") -> None:
    """Retângulo: y = topo, cresce para baixo."""
    pts = [(x, y), (x + w, y), (x + w, y - h), (x, y - h), (x, y)]
    for i in range(4):
        msp.add_line(pts[i], pts[i + 1], dxfattribs={"layer": layer})


def _txt(
    msp,
    text: str,
    x: float,
    y: float,
    h: float,
    layer: str = "0",
    bold: bool = False,
) -> None:
    """Texto centralizado em (x, y) com altura h."""
    style = "BOLD" if bold else "Standard"
    t = msp.add_text(
        str(text),
        dxfattribs={"height": h, "layer": layer, "style": style},
    )
    t.set_placement((x, y), align=TextEntityAlignment.MIDDLE_CENTER)


def draw_table_part(
    msp,
    table: dict,
    col_names: list,
    col_widths: list,
    total_w: float,
    data_slice: list,
    part_num: int,
    total_parts: int,
    ox: float,
    oy: float,
) -> float:
    """
    Desenha uma parte da tabela a partir de (ox, oy).
    Retorna Y do fundo da tabela desenhada.
    """
    data      = table["data"]
    first_col = col_names[0]
    prev_first = ""

    # Título
    title_text = "{} - Parte {}/{}".format(table["title"], part_num, total_parts)
    _rect(msp, ox, oy, total_w, TITLE_H, "TITULO")
    _txt(msp, title_text, ox + total_w / 2, oy - TITLE_H / 2,
         TEXT_H * 1.1, "TITULO", bold=True)
    oy -= TITLE_H

    # Seções
    for _, sec in table["sections"]:
        _rect(msp, ox, oy, total_w, SECTION_H, "SECAO")
        _txt(msp, sec, ox + total_w / 2, oy - SECTION_H / 2,
             TEXT_H, "SECAO", bold=True)
        oy -= SECTION_H

    # Cabeçalho de colunas
    x = ox
    for cname, cw in zip(col_names, col_widths):
        _rect(msp, x, oy, cw, HEADER_H, "CABECALHO")
        _txt(msp, cname, x + cw / 2, oy - HEADER_H / 2,
             TEXT_H, "CABECALHO", bold=True)
        x += cw
    oy -= HEADER_H

    # Linhas de dados
    for idx, y_key in enumerate(data_slice):
        row_data = data[y_key]
        layer    = "DADOS_ALT" if idx % 2 == 1 else "DADOS"

        val0 = row_data.get(first_col, "")
        if val0:
            prev_first = val0
        else:
            val0 = prev_first  # repete AÇO quando ausente na linha

        x = ox
        for i, (cname, cw) in enumerate(zip(col_names, col_widths)):
            raw     = val0 if i == 0 else row_data.get(cname, "")
            val     = to_number(raw)
            val_str = "" if val is None else str(val)

            _rect(msp, x, oy, cw, ROW_H, layer)
            _txt(msp, val_str, x + cw / 2, oy - ROW_H / 2,
                 TEXT_H * 0.92, layer)
            x += cw

        oy -= ROW_H

    return oy


# ─────────────────────────────────────────────────────────────────────────────
#  GERAR DXF
# ─────────────────────────────────────────────────────────────────────────────

def write_dxf(tables: list, out_path: str) -> None:
    # Colunas detectadas da primeira tabela
    col_names  = tables[0]["names"]
    col_widths = compute_col_widths(tables, col_names)
    total_w    = sum(col_widths)

    print(f"  Colunas: {dict(zip(col_names, [f'{w:.2f}' for w in col_widths]))}")
    print(f"  Largura total/tabela: {total_w:.2f}")

    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 6  # milímetros

    # Estilo negrito
    try:
        doc.styles.add("BOLD", font="calibrib.ttf")
    except Exception:
        pass

    # Layers
    layer_defs = {
        "TITULO":    colors.WHITE,
        "SECAO":     colors.CYAN,
        "CABECALHO": colors.YELLOW,
        "DADOS":     colors.WHITE,
        "DADOS_ALT": 7,  # cinza
    }
    for lname, col in layer_defs.items():
        try:
            doc.layers.add(lname, color=col)
        except Exception:
            pass

    msp = doc.modelspace()

    # Altura máxima de qualquer tabela (para calcular offset da 2ª linha)
    n_max    = math.ceil(max(len(t["sorted_y"]) for t in tables) / N_PARTS)
    sec_max  = max(len(t["sections"]) for t in tables)
    row_oy_2 = -(TITLE_H + sec_max * SECTION_H + HEADER_H + n_max * ROW_H + GAP_Y)

    for t_idx, table in enumerate(tables, start=1):
        sorted_y  = table["sorted_y"]
        n         = len(sorted_y)
        part_size = math.ceil(n / N_PARTS)
        row_oy    = 0.0 if t_idx == 1 else row_oy_2

        for p in range(N_PARTS):
            start = p * part_size
            end   = min(start + part_size, n)
            if start >= n:
                break
            data_slice = sorted_y[start:end]
            ox = p * (total_w + GAP_X)

            draw_table_part(
                msp, table, col_names, col_widths, total_w,
                data_slice, p + 1, N_PARTS, ox, row_oy,
            )

            print(f"  Tab{t_idx}_Parte{p+1}: {len(data_slice)} linhas"
                  f"  pos=({ox:.1f}, {row_oy:.1f})")

    doc.saveas(out_path)
    print(f"\nSalvo: {out_path}")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def process(dxf_in: str, dxf_out: str | None = None) -> None:
    if dxf_out is None:
        dxf_out = str(Path(dxf_in).with_stem(Path(dxf_in).stem + "_Split"))

    print(f"\nEntrada : {dxf_in}")
    print(f"Saida   : {dxf_out}")

    entities = parse_entities(dxf_in)
    print(f"  Entidades de texto: {len(entities)}")

    # Separar as 2 tabelas pelo X (Tabela 1: X<100, Tabela 2: X>100)
    ent1 = [e for e in entities if e["x"] < 100]
    ent2 = [e for e in entities if e["x"] > 100]
    print(f"  Tabela 1: {len(ent1)} entidades")
    print(f"  Tabela 2: {len(ent2)} entidades")

    table1 = build_table(ent1)
    table2 = build_table(ent2)
    print(f"  Tabela 1: {len(table1['sorted_y'])} linhas")
    print(f"  Tabela 2: {len(table2['sorted_y'])} linhas")

    write_dxf([table1, table2], dxf_out)


def main() -> None:
    args = sys.argv[1:]

    if len(args) == 0:
        # Seleção via janela
        try:
            import tkinter as tk
            from tkinter import filedialog, messagebox
        except ImportError:
            print("Uso: python tabelas_dxf_para_dxf.py arquivo.dxf")
            sys.exit(1)

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        dxf_in = filedialog.askopenfilename(
            title="Selecione o arquivo DXF",
            initialdir=str(Path.home() / "Downloads"),
            filetypes=[("DXF", "*.dxf"), ("Todos", "*.*")],
        )
        root.destroy()

        if not dxf_in:
            print("Nenhum arquivo selecionado.")
            sys.exit(0)

        try:
            process(dxf_in)
            messagebox.showinfo("Concluído", "DXF gerado com sucesso.")
        except Exception as ex:
            print(f"ERRO: {ex}")
            messagebox.showerror("Erro", str(ex))
            sys.exit(1)

    elif len(args) == 1:
        process(args[0])

    elif len(args) == 2:
        process(args[0], args[1])

    else:
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
