"""
Gerador de planilha de PILARES a partir de exportação Navisworks/AltoQi Eberick.
Abas: DETALHADO | RESUMO TOTAL | RESUMO POR CAMADA | PRORRATA
Colunas sem dados são suprimidas automaticamente.
"""
import re, sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side)
from openpyxl.utils import get_column_letter

# ── CONFIG ──────────────────────────────────────────────────────────────────
FILE_IN  = Path(r"D:\00-Obras - SLN\IFMT - SLN\0000- NAVISWORKS IFMT\IFMT TODOS OS PILARES TERREO.xlsx")
FILE_OUT = FILE_IN.with_name(FILE_IN.stem + "_PROCESSADO.xlsx")
TITULO   = "IFMT - PILARES TÉRREO"
SHEET_IN = 0          # índice da aba de origem (0 = primeira)

# ── ESTILOS ──────────────────────────────────────────────────────────────────
hdr_blue   = PatternFill('solid', start_color='1F4E79')
hdr_green  = PatternFill('solid', start_color='1E6B3C')
hdr_purple = PatternFill('solid', start_color='4A235A')
fill_total = PatternFill('solid', start_color='1F4E79')
fill_gold  = PatternFill('solid', start_color='D4AC0D')
fill_conc  = PatternFill('solid', start_color='D6E4F0')
fill_forma = PatternFill('solid', start_color='D5E8D4')
fill_aco   = PatternFill('solid', start_color='EDE7F6')
fill_prorata_hdr = PatternFill('solid', start_color='4A235A')
fill_prorata_sub = PatternFill('solid', start_color='F3EFF8')

wf = Font(name='Arial', bold=True, color='FFFFFF', size=10)
nf = Font(name='Arial', size=10)
bf = Font(name='Arial', bold=True, size=10)
center   = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_al  = Alignment(horizontal='left',   vertical='center')
right_al = Alignment(horizontal='right',  vertical='center')
thin = Side(style='thin', color='BFBFBF')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# ── PARSE ────────────────────────────────────────────────────────────────────
def parse_qty(val: str) -> float:
    """Extrai número de strings como '0.872 m³' ou '40.44 kg'."""
    if pd.isna(val):
        return 0.0
    m = re.search(r'[\d,.]+', str(val).replace(',', '.'))
    return float(m.group()) if m else 0.0

def extract_bitola(desc: str) -> float | None:
    """Retorna diâmetro em mm a partir da descrição, ex: 'Armadura - Aço CA50 - Ø 12.5 mm' → 12.5"""
    m = re.search(r'[Øø⌀-]?\s*([\d.]+)\s*mm', str(desc), re.IGNORECASE)
    return float(m.group(1)) if m else None

def load_source(path: Path, sheet) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet)
    col_name  = [c for c in df.columns if c.startswith('Item') and 'Name'  in c][0]
    col_layer = [c for c in df.columns if c.startswith('Item') and 'Layer' in c and 'Layer Id' not in c][0]
    desc_cols = sorted([c for c in df.columns if 'Descri' in c])
    qty_cols  = sorted([c for c in df.columns if 'Quantidade' in c])

    records = []
    for _, row in df.iterrows():
        rec = {'PILAR': str(row[col_name]).strip(),
               'CAMADA': str(row[col_layer]).strip(),
               'CONCRETO_m3': 0.0, 'FORMA_m2': 0.0}
        for dc, qc in zip(desc_cols, qty_cols):
            desc = str(row[dc]) if not pd.isna(row[dc]) else ''
            qty  = parse_qty(row[qc])
            if not desc or qty == 0:
                continue
            if 'Concreto' in desc:
                rec['CONCRETO_m3'] += qty
            elif 'Forma' in desc:
                rec['FORMA_m2'] += qty
            elif 'Armadura' in desc or 'Aço' in desc or 'Aco' in desc or 'CA50' in desc or 'CA60' in desc or 'Armadura' in desc.encode('latin-1', errors='replace').decode('latin-1', errors='replace'):
                b = extract_bitola(desc)
                if b:
                    key = f'ACO_{b}mm_kg'
                    rec[key] = rec.get(key, 0.0) + qty
        records.append(rec)

    result = pd.DataFrame(records)

    # bitolas presentes (sem zeros)
    aco_cols_all = sorted({c for c in result.columns if c.startswith('ACO_')},
                          key=lambda x: float(re.search(r'ACO_([\d.]+)mm', x).group(1)))
    active_aco   = [c for c in aco_cols_all if result[c].sum() > 0]
    bitolas      = [float(re.search(r'ACO_([\d.]+)mm', c).group(1)) for c in active_aco]

    # garante colunas
    for c in active_aco:
        if c not in result.columns:
            result[c] = 0.0
    result[active_aco] = result[active_aco].fillna(0.0)

    return result, bitolas

# ── HELPERS ──────────────────────────────────────────────────────────────────
def set_hdr(ws, row, col, val, font, fill):
    c = ws.cell(row=row, column=col, value=val)
    c.font = font; c.fill = fill; c.alignment = center; c.border = border

def set_cell(ws, row, col, val, font=None, fill=None, align=None, fmt=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font    = font  or nf
    c.fill    = fill  or PatternFill()
    c.alignment = align or right_al
    c.border  = border
    if fmt: c.number_format = fmt

# ── ABA 1: DETALHADO ─────────────────────────────────────────────────────────
def build_detalhado(wb, result, bitolas):
    ws = wb.active
    ws.title = 'DETALHADO'

    hdr_lbls  = ['PILAR', 'CAMADA', 'CONCRETO', 'FORMA'] + [f'ACO O{b}mm' for b in bitolas]
    unit_lbls = ['', '',    'm3',      'm2'     ] + ['kg'] * len(bitolas)
    fills_h   = [hdr_blue]*2 + [hdr_green] + [hdr_purple] + [hdr_purple] * len(bitolas)
    # corrige: CONCRETO=blue, FORMA=green, ACO=purple
    fills_h   = [hdr_blue, hdr_blue, hdr_blue, hdr_green] + [hdr_purple] * len(bitolas)

    for ci, (h, u, f) in enumerate(zip(hdr_lbls, unit_lbls, fills_h), 1):
        set_hdr(ws, 1, ci, h, wf, f)
        set_hdr(ws, 2, ci, u, wf, f)

    for ri, (_, rec) in enumerate(result.iterrows(), 3):
        rf = PatternFill('solid', start_color='F5F5F5') if ri % 2 == 1 else PatternFill('solid', start_color='FFFFFF')
        vals = [rec['PILAR'], rec['CAMADA'], rec['CONCRETO_m3'], rec['FORMA_m2']] + \
               [rec[f'ACO_{b}mm_kg'] for b in bitolas]
        for ci, val in enumerate(vals, 1):
            fmt = '#,##0.00' if ci > 2 else None
            set_cell(ws, ri, ci, val, fill=rf,
                     align=left_al if ci <= 2 else right_al, fmt=fmt)

    ri_tot = 3 + len(result)
    tvs = ['TOTAL', '',
           result['CONCRETO_m3'].sum(),
           result['FORMA_m2'].sum()] + \
          [result[f'ACO_{b}mm_kg'].sum() for b in bitolas]
    for ci, v in enumerate(tvs, 1):
        c = ws.cell(row=ri_tot, column=ci, value=v)
        c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
        c.fill = fill_total; c.border = border
        c.alignment = left_al if ci <= 2 else right_al
        if ci > 2 and isinstance(v, float): c.number_format = '#,##0.00'

    for ci, w in enumerate([14, 20, 14, 14] + [13] * len(bitolas), 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = 'A3'

# ── ABA 2: RESUMO TOTAL ───────────────────────────────────────────────────────
def build_resumo_total(wb, result, bitolas):
    ws2 = wb.create_sheet('RESUMO TOTAL')
    ws2['A1'] = f'RESUMO GERAL - {TITULO}'
    ws2['A1'].font = Font(name='Arial', bold=True, size=14, color='1F4E79')
    ws2.merge_cells('A1:D1'); ws2['A1'].alignment = center

    for ci, h in enumerate(['CATEGORIA', 'DESCRICAO', 'TOTAL', 'UNIDADE'], 1):
        set_hdr(ws2, 3, ci, h, wf, hdr_blue)

    rows_r = [('CONCRETO', 'Concreto C-30 - Abatimento 12 cm', result['CONCRETO_m3'].sum(), 'm3'),
              ('FORMA',    'Forma - Estrutura - Concreto',      result['FORMA_m2'].sum(),    'm2')]
    for b in bitolas:
        t = result[f'ACO_{b}mm_kg'].sum()
        if t > 0:
            rows_r.append(('ACO', f'Armadura CA50 - O {b} mm', t, 'kg'))

    fc = {'CONCRETO': fill_conc, 'FORMA': fill_forma, 'ACO': fill_aco}
    for ri, (cat, desc, val, unit) in enumerate(rows_r, 4):
        f = fc.get(cat, fill_conc)
        for ci, v in enumerate([cat, desc, val, unit], 1):
            set_cell(ws2, ri, ci, v, fill=f,
                     align=left_al if ci == 2 else center,
                     fmt='#,##0.00' if ci == 3 else None)

    r_gt = 4 + len(rows_r)
    ag = sum(result[f'ACO_{b}mm_kg'].sum() for b in bitolas)
    for ci, v in enumerate(['ACO TOTAL', 'Total Geral ACO CA50 (todas as bitolas)', ag, 'kg'], 1):
        c = ws2.cell(row=r_gt, column=ci, value=v)
        c.font = Font(name='Arial', bold=True, size=10)
        c.fill = fill_gold; c.border = border
        c.alignment = left_al if ci == 2 else center
        if ci == 3: c.number_format = '#,##0.00'

    ws2.column_dimensions['A'].width = 14
    ws2.column_dimensions['B'].width = 44
    ws2.column_dimensions['C'].width = 16
    ws2.column_dimensions['D'].width = 12

# ── ABA 3: RESUMO POR CAMADA ──────────────────────────────────────────────────
def build_resumo_camada(wb, result, bitolas):
    ws3 = wb.create_sheet('RESUMO POR CAMADA')
    nc3 = 3 + len(bitolas) + 1
    ws3['A1'] = f'RESUMO POR CAMADA - {TITULO}'
    ws3['A1'].font = Font(name='Arial', bold=True, size=14, color='1F4E79')
    ws3.merge_cells(f'A1:{get_column_letter(nc3)}1')
    ws3['A1'].alignment = center

    h3 = ['CAMADA', 'CONCRETO (m3)', 'FORMA (m2)'] + \
         [f'ACO O{b}mm (kg)' for b in bitolas] + ['ACO TOTAL (kg)']
    for ci, h in enumerate(h3, 1):
        set_hdr(ws3, 3, ci, h, wf, hdr_blue)

    for ri, cam in enumerate(result['CAMADA'].unique(), 4):
        sub = result[result['CAMADA'] == cam]
        at  = sum(sub[f'ACO_{b}mm_kg'].sum() for b in bitolas)
        v3  = [cam, sub['CONCRETO_m3'].sum(), sub['FORMA_m2'].sum()] + \
              [sub[f'ACO_{b}mm_kg'].sum() for b in bitolas] + [at]
        rf  = PatternFill('solid', start_color='EBF3FB') if ri % 2 == 0 \
              else PatternFill('solid', start_color='FFFFFF')
        for ci, v in enumerate(v3, 1):
            set_cell(ws3, ri, ci, v, fill=rf,
                     align=left_al if ci == 1 else right_al,
                     fmt='#,##0.00' if ci > 1 else None)

    rtt = 4 + len(result['CAMADA'].unique())
    aa  = sum(result[f'ACO_{b}mm_kg'].sum() for b in bitolas)
    tt  = ['TOTAL GERAL', result['CONCRETO_m3'].sum(), result['FORMA_m2'].sum()] + \
          [result[f'ACO_{b}mm_kg'].sum() for b in bitolas] + [aa]
    for ci, v in enumerate(tt, 1):
        c = ws3.cell(row=rtt, column=ci, value=v)
        c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
        c.fill = fill_total; c.border = border
        c.alignment = left_al if ci == 1 else right_al
        if ci > 1: c.number_format = '#,##0.00'

    for ci, w in enumerate([22, 16, 16] + [14] * len(bitolas) + [16], 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w

# ── ABA 4: PRORRATA ───────────────────────────────────────────────────────────
def build_prorrata(wb, result, bitolas):
    """
    Para cada sapata: valor absoluto + % do total de cada material.
    Útil para ver quais sapatas dominam o consumo de material.
    """
    ws4 = wb.create_sheet('PRORRATA')

    # totais
    tot_conc  = result['CONCRETO_m3'].sum()
    tot_forma = result['FORMA_m2'].sum()
    tot_acos  = {b: result[f'ACO_{b}mm_kg'].sum() for b in bitolas}
    tot_aco_g = sum(tot_acos.values())

    # cabeçalho título
    n_cols = 3 + len(bitolas) * 2 + 2   # SAPATA + CAMADA + conc(val+%) + forma(val+%) + aco bitolas(val+%) + total_aco(val+%)
    ws4['A1'] = f'PRORRATA (% PARTICIPAÇÃO) - {TITULO}'
    ws4['A1'].font = Font(name='Arial', bold=True, size=14, color='1F4E79')
    ws4.merge_cells(f'A1:{get_column_letter(n_cols)}1')
    ws4['A1'].alignment = center

    # cabeçalhos linha 2 (grupo) e linha 3 (sub)
    # estrutura: SAPATA | CAMADA | CONCRETO(m3) | % | FORMA(m2) | % | ACO_b1(kg) | % | ... | ACO TOTAL(kg) | %
    # montar via loop
    col_map = []  # (label_l2, label_l3, fill, is_pct, key_abs, tot)
    col_map.append(('PILAR',   '',   hdr_blue,   False, 'PILAR',        None))
    col_map.append(('CAMADA',  '',   hdr_blue,   False, 'CAMADA',       None))
    col_map.append(('CONCRETO', 'm3', hdr_blue,  False, 'CONCRETO_m3',  tot_conc))
    col_map.append(('%',        '%',  hdr_blue,  True,  'CONCRETO_m3',  tot_conc))
    col_map.append(('FORMA',   'm2', hdr_green,  False, 'FORMA_m2',     tot_forma))
    col_map.append(('%',       '%',  hdr_green,  True,  'FORMA_m2',     tot_forma))
    for b in bitolas:
        col_map.append((f'ACO O{b}mm', 'kg', hdr_purple, False, f'ACO_{b}mm_kg', tot_acos[b]))
        col_map.append(('%',           '%',  hdr_purple, True,  f'ACO_{b}mm_kg', tot_acos[b]))
    col_map.append(('ACO TOTAL', 'kg', fill_gold, False, '__ACO_TOTAL__', tot_aco_g))
    col_map.append(('%',         '%',  fill_gold, True,  '__ACO_TOTAL__', tot_aco_g))

    # escreve cabeçalhos linha 2-3
    for ci, (l2, l3, fill, _, __, ___) in enumerate(col_map, 1):
        set_hdr(ws4, 2, ci, l2, wf, fill)
        set_hdr(ws4, 3, ci, l3, wf, fill)

    # dados
    for ri, (_, rec) in enumerate(result.iterrows(), 4):
        rf = PatternFill('solid', start_color='F3EFF8') if ri % 2 == 0 \
             else PatternFill('solid', start_color='FFFFFF')
        aco_total_row = sum(rec[f'ACO_{b}mm_kg'] for b in bitolas)

        for ci, (_, _, fill, is_pct, key, tot) in enumerate(col_map, 1):
            if key == 'PILAR':
                set_cell(ws4, ri, ci, rec['PILAR'], fill=rf, align=left_al)
            elif key == 'CAMADA':
                set_cell(ws4, ri, ci, rec['CAMADA'], fill=rf, align=left_al)
            elif key == '__ACO_TOTAL__':
                val = (aco_total_row / tot * 100) if (is_pct and tot) else aco_total_row
                fmt = '0.00"%"' if is_pct else '#,##0.00'
                set_cell(ws4, ri, ci, val, fill=rf, align=right_al, fmt=fmt)
            else:
                abs_val = rec.get(key, 0.0) if hasattr(rec, 'get') else rec[key]
                val = (abs_val / tot * 100) if (is_pct and tot) else abs_val
                fmt = '0.00"%"' if is_pct else '#,##0.00'
                set_cell(ws4, ri, ci, val, fill=rf, align=right_al, fmt=fmt)

    # linha total
    ri_tot = 4 + len(result)
    for ci, (_, _, fill, is_pct, key, tot) in enumerate(col_map, 1):
        c = ws4.cell(row=ri_tot, column=ci)
        c.border = border
        c.font   = Font(name='Arial', bold=True, color='FFFFFF', size=10)
        c.fill   = fill_total
        if key == 'PILAR':
            c.value = 'TOTAL'; c.alignment = left_al
        elif key == 'CAMADA':
            c.value = ''; c.alignment = left_al
        elif is_pct:
            c.value = 100.0; c.number_format = '0.00"%"'; c.alignment = right_al
        elif key == '__ACO_TOTAL__':
            c.value = tot_aco_g; c.number_format = '#,##0.00'; c.alignment = right_al
        else:
            c.value = tot; c.number_format = '#,##0.00'; c.alignment = right_al

    # larguras
    for ci, (l2, _, __, is_pct, ___, ____) in enumerate(col_map, 1):
        if ci <= 2:
            ws4.column_dimensions[get_column_letter(ci)].width = 14 if ci == 1 else 22
        elif is_pct:
            ws4.column_dimensions[get_column_letter(ci)].width = 8
        else:
            ws4.column_dimensions[get_column_letter(ci)].width = 13

    ws4.row_dimensions[2].height = 30
    ws4.row_dimensions[3].height = 18
    ws4.freeze_panes = 'A4'

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print(f"Lendo: {FILE_IN}")
    result, bitolas = load_source(FILE_IN, SHEET_IN)
    print(f"  {len(result)} pilares | bitolas ativas: {bitolas}")

    wb = Workbook()
    build_detalhado(wb, result, bitolas)
    build_resumo_total(wb, result, bitolas)
    build_resumo_camada(wb, result, bitolas)
    build_prorrata(wb, result, bitolas)

    wb.save(FILE_OUT)
    print(f"\nArquivo salvo: {FILE_OUT}")

if __name__ == '__main__':
    main()
