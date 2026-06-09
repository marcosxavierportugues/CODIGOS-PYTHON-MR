import pandas as pd
import re
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── INPUTS ────────────────────────────────────────────────────────────────────
FILES = [
    r'D:\00-Obras - SLN\IFMT - SLN\002- NAVISWORKS EXECUTIVO SLN\2 PAVIMENTO - EXECUTIVO\LAJES TRELIÇADAS 2 PAV EXECUTIVO\LAES 2PAV - TRELIÇADAS EXECUTIVO - SEG 1.xlsx',
    r'D:\00-Obras - SLN\IFMT - SLN\002- NAVISWORKS EXECUTIVO SLN\2 PAVIMENTO - EXECUTIVO\LAJES TRELIÇADAS 2 PAV EXECUTIVO\lajes - 2 pav- exec - seg 2- TRELIÇADA.xlsx',
]
OUT = r'D:\00-Obras - SLN\IFMT - SLN\002- NAVISWORKS EXECUTIVO SLN\2 PAVIMENTO - EXECUTIVO\LAJES TRELIÇADAS 2 PAV EXECUTIVO\RESUMO\LAJE TRELICADA 2 PAV - TOTAL.xlsx'
ELEM_LABEL = 'LAJE'

# ── HELPERS ───────────────────────────────────────────────────────────────────
def parse_qty(v):
    if pd.isna(v): return 0.0
    m = re.search(r'[\d]+[.,]?[\d]*', str(v))
    return float(m.group().replace(',', '.')) if m else 0.0

def parse_tipo(desc):
    if pd.isna(desc): return None
    d = str(desc).upper()
    if 'CONCRETO' in d:
        return 'CONCRETO'
    if 'ENCHIMENTO' in d or 'EPS' in d:
        m = re.search(r'B\d+/\d+/\d+', str(desc), re.I)
        return ('EPS_' + m.group().upper()) if m else 'EPS_GERAL'
    if 'TRELI' in d:
        m = re.search(r'TR\s*(\d+)', str(desc), re.I)
        return ('TRELICA_TR' + m.group(1)) if m else 'TRELICA_GERAL'
    if 'ARMADURA' in d and 'PUN' not in d and 'CHAPA' not in d:
        m = re.search(r'(\d+[.,]\d+)\s*mm', str(desc), re.I)
        grade = 'CA60' if 'CA60' in d else 'CA25' if 'CA25' in d else 'CA50'
        return (m.group(1).replace(',', '.') + 'mm_' + grade) if m else None
    return None

# ── PARSE ALL FILES ────────────────────────────────────────────────────────────
records = []
for f in FILES:
    df = pd.read_excel(f, header=None)
    hdr = [str(c).upper().replace('\n', ' ').replace('\r', '') for c in df.iloc[0]]
    n = len(hdr)

    nivel_col = seg_col = None
    for i in range(n - 1, -1, -1):
        h = hdr[i]
        if nivel_col is None and 'NIVEL' in h:
            nivel_col = i
        if seg_col is None and ('SEG' in h or 'SEGMENTO' in h) and 'DESCR' not in h:
            seg_col = i
        if nivel_col is not None and seg_col is not None:
            break

    end_items = min(nivel_col, seg_col)

    for _, row in df.iloc[1:].iterrows():
        nome = str(row.iloc[0])
        if pd.isna(row.iloc[0]) or nome == 'nan':
            continue
        nivel = str(row.iloc[nivel_col]) if nivel_col is not None else ''
        seg   = str(row.iloc[seg_col])   if seg_col   is not None else ''

        rec = {'NOME': nome, 'NIVEL': nivel, 'SEG': seg, 'CONCRETO': 0.0}

        i = 1
        while i + 1 < end_items:
            tipo = parse_tipo(row.iloc[i])
            val  = parse_qty(row.iloc[i + 1])
            if tipo == 'CONCRETO':
                rec['CONCRETO'] += val
            elif tipo and tipo != 'CONCRETO':
                rec[tipo] = rec.get(tipo, 0.0) + val
            i += 2

        records.append(rec)

df_all = pd.DataFrame(records)
# fill NaN for optional cols
for col in df_all.columns:
    if col not in ('NOME', 'NIVEL', 'SEG'):
        df_all[col] = pd.to_numeric(df_all[col], errors='coerce').fillna(0.0)

print('Total lajes:', len(df_all))

# ── DETERMINE COLUMN ORDER ─────────────────────────────────────────────────────
eps_cols     = sorted([c for c in df_all.columns if c.startswith('EPS_')])
trelica_cols = sorted([c for c in df_all.columns if c.startswith('TRELICA_')],
                      key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)
aco_cols_raw = [c for c in df_all.columns if re.match(r'^\d+\.\d+mm_CA', c)]
aco_cols     = sorted(aco_cols_raw, key=lambda x: float(x.split('mm')[0]))
all_qty_cols = ['CONCRETO'] + eps_cols + trelica_cols + aco_cols

print('Qty cols:', all_qty_cols)

# ── VALIDATION ────────────────────────────────────────────────────────────────
print('--- VALIDATION ---')
for col in all_qty_cols:
    if col in df_all.columns:
        print(f'{col}: {df_all[col].sum():.4f}')
segs = sorted(df_all['SEG'].unique())
print('Segments:', segs)
for s in segs:
    sub = df_all[df_all['SEG'] == s]
    print(f'  {s}: {len(sub)} lajes')

# ── STYLES ────────────────────────────────────────────────────────────────────
CI = '1F4E79'; CC = '1565C0'; CE = 'C47000'; CT = '1E6B3C'; CA = '4A235A'
ROW_ODD = 'F0F4F8'; ROW_EVEN = 'FFFFFF'; ROW_TOT = 'D6E4F0'

def col_color(col):
    if col == 'CONCRETO':              return CC
    if col.startswith('EPS_'):         return CE
    if col.startswith('TRELICA_'):     return CT
    if re.match(r'^\d+', col):         return CA
    return CI

def col_display(col):
    if col == 'CONCRETO': return 'CONCRETO'
    if col.startswith('EPS_'):
        return 'EPS ' + col[4:]
    if col.startswith('TRELICA_'):
        return col[8:].replace('TR', 'TR ')
    if re.match(r'^\d+\.\d+mm_CA', col):
        return chr(216) + col.split('mm')[0] + 'mm'
    return col

def col_unit(col):
    if col == 'CONCRETO': return 'm³'
    if col.startswith('EPS_'): return 'un'
    if col.startswith('TRELICA_'): return 'kg'
    if re.match(r'^\d+', col): return 'kg'
    return ''

thin   = Side(style='thin', color='CCCCCC')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
font_base = Font(name='Arial', size=10)
font_hdr  = Font(name='Arial', size=10, bold=True, color='FFFFFF')
font_tot  = Font(name='Arial', size=10, bold=True)

def fill(hex_color):
    return PatternFill('solid', fgColor=hex_color)

def write_sheet(ws, rows_df, show_seg=False):
    out_cols = [ELEM_LABEL]
    if show_seg: out_cols.append('SEG')
    out_cols.append('NIVEL')
    out_cols += all_qty_cols

    # Row 1 - header names
    for ci, col in enumerate(out_cols, 1):
        hx = CI if col in (ELEM_LABEL, 'SEG', 'NIVEL') else col_color(col)
        label = col if col in (ELEM_LABEL, 'SEG', 'NIVEL') else col_display(col)
        c = ws.cell(row=1, column=ci)
        c.value = label; c.font = font_hdr; c.fill = fill(hx)
        c.border = border
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Row 2 - units
    for ci, col in enumerate(out_cols, 1):
        hx = CI if col in (ELEM_LABEL, 'SEG', 'NIVEL') else col_color(col)
        c = ws.cell(row=2, column=ci)
        c.value = '' if col in (ELEM_LABEL, 'SEG', 'NIVEL') else col_unit(col)
        c.font = font_hdr; c.fill = fill(hx); c.border = border
        c.alignment = Alignment(horizontal='center', vertical='center')

    # Data rows
    rows_df = rows_df.reset_index(drop=True)
    for ri_idx, (_, row) in enumerate(rows_df.iterrows()):
        ri = ri_idx + 3
        bg = ROW_ODD if (ri_idx % 2 == 0) else ROW_EVEN
        for ci, col in enumerate(out_cols, 1):
            c = ws.cell(row=ri, column=ci)
            c.border = border; c.fill = fill(bg)
            if col == ELEM_LABEL:
                c.value = row['NOME']; c.font = font_base
                c.alignment = Alignment(horizontal='left', vertical='center')
            elif col == 'SEG':
                c.value = row['SEG']; c.font = font_base
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif col == 'NIVEL':
                c.value = row['NIVEL']; c.font = font_base
                c.alignment = Alignment(horizontal='center', vertical='center')
            else:
                val = float(row.get(col, 0.0)) if not pd.isna(row.get(col, 0.0)) else 0.0
                c.value = val; c.font = font_base
                c.number_format = '#,##0.0000'
                c.alignment = Alignment(horizontal='right', vertical='center')

    last_data = 2 + len(rows_df)
    tot_row   = last_data + 1
    for ci, col in enumerate(out_cols, 1):
        c = ws.cell(row=tot_row, column=ci)
        c.fill = fill(ROW_TOT); c.border = border; c.font = font_tot
        if ci == 1:
            c.value = 'TOTAL GERAL'
            c.alignment = Alignment(horizontal='left', vertical='center')
        elif col in ('SEG', 'NIVEL'):
            c.alignment = Alignment(horizontal='center', vertical='center')
        else:
            cl = get_column_letter(ci)
            c.value = f'=SUM({cl}3:{cl}{last_data})'
            c.number_format = '#,##0.0000'
            c.alignment = Alignment(horizontal='right', vertical='center')

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20
    for r in range(3, tot_row + 1):
        ws.row_dimensions[r].height = 18

    for ci, col in enumerate(out_cols, 1):
        if col == ELEM_LABEL:           w = 18
        elif col == 'SEG':              w = 8
        elif col == 'NIVEL':            w = 16
        elif col.startswith('EPS_'):    w = 15
        elif col.startswith('TRELICA_'): w = 13
        else:                           w = 13
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = 'A3'

# ── BUILD WORKBOOK ─────────────────────────────────────────────────────────────
wb = Workbook()
wb.remove(wb.active)

ws_all = wb.create_sheet('TODOS OS LAJES')
write_sheet(ws_all, df_all, show_seg=True)

for seg in segs:
    sub = df_all[df_all['SEG'] == seg].copy()
    ws_seg = wb.create_sheet(seg)
    write_sheet(ws_seg, sub, show_seg=False)

# ── LISTA DE MATERIAIS ─────────────────────────────────────────────────────────
ws_lista = wb.create_sheet('LISTA DE MATERIAIS')
n_seg_cols = len(segs)
total_header_cols = 4 + n_seg_cols  # #, ITEM, DESC, TOTAL, segs..., UN
ws_lista.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_header_cols + 1)
tc = ws_lista.cell(row=1, column=1)
tc.value = 'LISTA DE MATERIAIS - LAJE TRELICADA 2 PAV'
tc.font  = Font(name='Arial', size=13, bold=True, color='FFFFFF')
tc.fill  = fill('1F4E79')
tc.alignment = Alignment(horizontal='center', vertical='center')
ws_lista.row_dimensions[1].height = 30

hdr_cols = ['#', 'ITEM', 'DESCRICAO', 'TOTAL'] + list(segs) + ['UN']
seg_palette = ['2E4057', '1B4332', '4A235A', '7B3F00']
hdr_colors_lista = {
    '#': '1F4E79', 'ITEM': '1F4E79', 'DESCRICAO': '1F4E79',
    'TOTAL': '1F4E79', 'UN': '1F4E79'
}
for si, s in enumerate(segs):
    hdr_colors_lista[s] = seg_palette[si % len(seg_palette)]

for ci, col in enumerate(hdr_cols, 1):
    c = ws_lista.cell(row=3, column=ci)
    c.value = 'DESCRIÇÃO' if col == 'DESCRICAO' else col
    c.font  = font_hdr
    c.fill  = fill(hdr_colors_lista.get(col, '1F4E79'))
    c.border = border
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws_lista.row_dimensions[3].height = 25

mat_rows = []
mat_rows.append({
    'ITEM': 'CONCRETO', 'DESC': 'Concreto C-30', 'UN': 'm³',
    'TOTAL': df_all['CONCRETO'].sum(),
    **{s: df_all[df_all['SEG'] == s]['CONCRETO'].sum() for s in segs}
})
for col in eps_cols:
    mat_rows.append({
        'ITEM': 'ENCHIMENTO', 'DESC': 'EPS Unidirecional - ' + col[4:], 'UN': 'un',
        'TOTAL': df_all[col].sum() if col in df_all.columns else 0.0,
        **{s: df_all[df_all['SEG'] == s][col].sum() if col in df_all.columns else 0.0 for s in segs}
    })
for col in trelica_cols:
    model = col[8:].replace('TR', 'TR ')
    mat_rows.append({
        'ITEM': 'TRELIÇA', 'DESC': 'Aço CA60 - ' + model, 'UN': 'kg',
        'TOTAL': df_all[col].sum() if col in df_all.columns else 0.0,
        **{s: df_all[df_all['SEG'] == s][col].sum() if col in df_all.columns else 0.0 for s in segs}
    })
for col in aco_cols:
    bitola = chr(216) + col.split('mm')[0] + 'mm'
    grade  = col.split('_')[1] if '_' in col else 'CA50'
    mat_rows.append({
        'ITEM': 'AÇO', 'DESC': 'Armadura ' + grade + ' - ' + bitola, 'UN': 'kg',
        'TOTAL': df_all[col].sum() if col in df_all.columns else 0.0,
        **{s: df_all[df_all['SEG'] == s][col].sum() if col in df_all.columns else 0.0 for s in segs}
    })

for ri, mrow in enumerate(mat_rows, 4):
    bg = 'D6E4F0' if ri % 2 == 0 else 'FFFFFF'
    for ci, col in enumerate(hdr_cols, 1):
        c = ws_lista.cell(row=ri, column=ci)
        c.border = border; c.fill = fill(bg)
        if col == '#':
            c.value = ri - 3
            c.font = Font(name='Arial', size=10, bold=True)
            c.alignment = Alignment(horizontal='center', vertical='center')
        elif col == 'ITEM':
            c.value = mrow['ITEM']
            c.font = Font(name='Arial', size=10, bold=True)
            c.alignment = Alignment(horizontal='left', vertical='center')
        elif col == 'DESCRICAO':
            c.value = mrow['DESC']; c.font = font_base
            c.alignment = Alignment(horizontal='left', vertical='center')
        elif col == 'UN':
            c.value = mrow['UN']; c.font = font_base
            c.alignment = Alignment(horizontal='center', vertical='center')
        elif col == 'TOTAL':
            c.value = mrow['TOTAL']; c.number_format = '#,##0.0000'
            c.font = Font(name='Arial', size=10, bold=True)
            c.alignment = Alignment(horizontal='right', vertical='center')
        else:
            c.value = mrow.get(col, 0.0); c.number_format = '#,##0.0000'
            c.font = font_base
            c.alignment = Alignment(horizontal='right', vertical='center')

ws_lista.column_dimensions['A'].width = 5
ws_lista.column_dimensions['B'].width = 14
ws_lista.column_dimensions['C'].width = 38
ws_lista.column_dimensions['D'].width = 16
for i, s in enumerate(segs, 5):
    ws_lista.column_dimensions[get_column_letter(i)].width = 16
ws_lista.column_dimensions[get_column_letter(len(hdr_cols))].width = 6
ws_lista.freeze_panes = 'A4'

# ── SAVE ──────────────────────────────────────────────────────────────────────
wb.save(OUT)
print('Saved:', OUT)
