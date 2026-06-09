import pandas as pd
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

file_in = r'D:\00-Obras - SLN\IFMT - SLN\00- ENVIADO IFMT\IFMT - VIGAS 1PAV  - ENVIADO - 16.4.26-V0.xlsx'
file_out = r'D:\00-Obras - SLN\IFMT - SLN\00- ENVIADO IFMT\IFMT - VIGAS 1PAV  - ENVIADO - 16.4.26-V0_RESUMO.xlsx'

df = pd.read_excel(file_in, sheet_name='Sheet1', header=0)
df.columns = [str(c) for c in df.columns]

item_pairs = [(3, 4), (5, 6), (7, 8), (9, 10)]


def extract_bitola(desc):
    m = re.search(r'[O\u00d8\u00f8]\s*([\d,\.]+)\s*mm', str(desc))
    if not m:
        # fallback: look for digit pattern after space
        m = re.search(r'\xd8\s*([\d,\.]+)\s*mm', str(desc))
    if not m:
        m = re.search(r'\xf8\s*([\d,\.]+)\s*mm', str(desc))
    if m:
        val = m.group(1).replace(',', '.')
        return float(val)
    # Try raw bytes pattern for corrupted Ø
    m2 = re.search(r'(\d+[,\.]\d+)\s*mm', str(desc))
    if m2:
        val = m2.group(1).replace(',', '.')
        return float(val)
    return None


def classify(desc):
    d = str(desc)
    if 'Forma' in d:
        return 'FORMA'
    if 'Concreto' in d:
        return 'CONCRETO'
    if 'Armadura' in d or 'rmadura' in d or 'A\u00e7o' in d or 'A\ufffd' in d:
        return 'ACO'
    return None


# Collect all bitolas
all_bitolas = set()
for _, row in df.iterrows():
    for dc, qc in item_pairs:
        desc = row.iloc[dc]
        if pd.notna(desc):
            cat = classify(desc)
            if cat == 'ACO':
                b = extract_bitola(desc)
                if b:
                    all_bitolas.add(b)

bitolas = sorted(all_bitolas)
print('Bitolas encontradas:', bitolas)

# Build structured data
records = []
for _, row in df.iterrows():
    pilar = str(row.iloc[0])
    layer = str(row.iloc[1])
    ref = str(row.iloc[2])

    concreto = 0.0
    forma = 0.0
    aco = {b: 0.0 for b in bitolas}

    for dc, qc in item_pairs:
        desc = row.iloc[dc]
        qty = row.iloc[qc]
        if pd.isna(desc) or pd.isna(qty):
            continue
        cat = classify(desc)
        if cat == 'CONCRETO':
            concreto += float(qty)
        elif cat == 'FORMA':
            forma += float(qty)
        elif cat == 'ACO':
            b = extract_bitola(desc)
            if b and b in aco:
                aco[b] += float(qty)

    rec = {'PILAR': pilar, 'CAMADA': layer, 'SECAO': ref,
           'CONCRETO_m3': concreto, 'FORMA_m2': forma}
    for b in bitolas:
        rec[f'ACO_{b}mm_kg'] = aco[b]
    records.append(rec)

result = pd.DataFrame(records)

# VALIDATION
print('\n=== VALIDACAO - TOTAIS GERAIS ===')
print(f'Total CONCRETO: {result["CONCRETO_m3"].sum():.4f} m3')
print(f'Total FORMA:    {result["FORMA_m2"].sum():.4f} m2')
for b in bitolas:
    col = f'ACO_{b}mm_kg'
    t = result[col].sum()
    if t > 0:
        print(f'Total ACO O{b}mm: {t:.4f} kg')
print(f'Linhas (pilares): {len(result)}')

# ========================
# BUILD EXCEL
# ========================
wb = openpyxl.Workbook()

thin = Side(style='thin', color='BBBBBB')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_al = Alignment(horizontal='left', vertical='center')
right_al = Alignment(horizontal='right', vertical='center')

white_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
bold_font = Font(name='Arial', bold=True, size=10)
normal_font = Font(name='Arial', size=10)

hdr_blue = PatternFill('solid', start_color='1F4E79')
hdr_green = PatternFill('solid', start_color='375623')
hdr_purple = PatternFill('solid', start_color='4A235A')
hdr_dark = PatternFill('solid', start_color='1F4E79')
fill_gold = PatternFill('solid', start_color='D4AC0D')
fill_conc = PatternFill('solid', start_color='D6E4F0')
fill_forma = PatternFill('solid', start_color='D5E8D4')
fill_aco = PatternFill('solid', start_color='EDE7F6')
fill_total = PatternFill('solid', start_color='1F4E79')

# =============================================
# ABA 1 - DETALHADO
# =============================================
ws = wb.active
ws.title = 'DETALHADO'

header_labels = ['PILAR', 'CAMADA', 'SECAO', 'CONCRETO', 'FORMA'] + [f'ACO O{b}mm' for b in bitolas]
unit_labels = ['', '', '', 'm3', 'm2'] + ['kg'] * len(bitolas)

fills_hdr = [hdr_blue, hdr_blue, hdr_blue, hdr_blue, hdr_green] + [hdr_purple] * len(bitolas)

for ci, (h, u, f) in enumerate(zip(header_labels, unit_labels, fills_hdr), 1):
    c1 = ws.cell(row=1, column=ci, value=h)
    c1.font = white_font
    c1.fill = f
    c1.alignment = center
    c1.border = border
    c2 = ws.cell(row=2, column=ci, value=u)
    c2.font = white_font
    c2.fill = f
    c2.alignment = center
    c2.border = border

for ri, rec in enumerate(records, 3):
    row_fill = PatternFill('solid', start_color='F5F5F5') if ri % 2 == 1 else PatternFill('solid', start_color='FFFFFF')
    vals = [rec['PILAR'], rec['CAMADA'], rec['SECAO'],
            rec['CONCRETO_m3'], rec['FORMA_m2']] + [rec[f'ACO_{b}mm_kg'] for b in bitolas]
    for ci, val in enumerate(vals, 1):
        c = ws.cell(row=ri, column=ci, value=val)
        c.font = normal_font
        c.border = border
        c.fill = row_fill
        c.alignment = left_al if ci <= 3 else right_al
        if ci > 3:
            c.number_format = '#,##0.00'

# TOTAL row
ri_tot = 3 + len(records)
tot_vals = ['TOTAL', '', ''] + \
    [result['CONCRETO_m3'].sum(), result['FORMA_m2'].sum()] + \
    [result[f'ACO_{b}mm_kg'].sum() for b in bitolas]
for ci, v in enumerate(tot_vals, 1):
    c = ws.cell(row=ri_tot, column=ci, value=v)
    c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    c.fill = fill_total
    c.border = border
    c.alignment = left_al if ci <= 3 else right_al
    if ci > 3 and isinstance(v, float):
        c.number_format = '#,##0.00'

col_widths = [10, 18, 34, 14, 14] + [13] * len(bitolas)
for ci, w in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(ci)].width = w
ws.row_dimensions[1].height = 30
ws.row_dimensions[2].height = 18
ws.freeze_panes = 'A3'

# =============================================
# ABA 2 - RESUMO TOTAL
# =============================================
ws2 = wb.create_sheet('RESUMO TOTAL')
ws2['A1'] = 'RESUMO GERAL - PILARES TERREO'
ws2['A1'].font = Font(name='Arial', bold=True, size=14, color='1F4E79')
ws2.merge_cells('A1:D1')
ws2['A1'].alignment = center

for ci, h in enumerate(['CATEGORIA', 'DESCRICAO', 'TOTAL', 'UNIDADE'], 1):
    c = ws2.cell(row=3, column=ci, value=h)
    c.font = white_font
    c.fill = hdr_dark
    c.alignment = center
    c.border = border

rows_resumo = []
rows_resumo.append(('CONCRETO', 'Concreto C-30 - Abatimento 12 cm', result['CONCRETO_m3'].sum(), 'm3'))
rows_resumo.append(('FORMA', 'Forma - Estrutura - Concreto', result['FORMA_m2'].sum(), 'm2'))
for b in bitolas:
    col = f'ACO_{b}mm_kg'
    t = result[col].sum()
    if t > 0:
        steel = 'CA60' if b == 5.0 else 'CA50'
        rows_resumo.append(('ACO', f'Armadura {steel} - O {b} mm', t, 'kg'))

fills_cat = {'CONCRETO': fill_conc, 'FORMA': fill_forma, 'ACO': fill_aco}

for ri, (cat, desc, val, unit) in enumerate(rows_resumo, 4):
    f = fills_cat.get(cat, fill_conc)
    for ci, v in enumerate([cat, desc, val, unit], 1):
        c = ws2.cell(row=ri, column=ci, value=v)
        c.font = normal_font
        c.border = border
        c.fill = f
        c.alignment = left_al if ci == 2 else center
        if ci == 3:
            c.number_format = '#,##0.00'

# Grand total ACO
r_gt = 4 + len(rows_resumo)
aco_grand = sum(result[f'ACO_{b}mm_kg'].sum() for b in bitolas)
for ci, v in enumerate(['ACO TOTAL', 'Total Geral de Aco (todas as bitolas)', aco_grand, 'kg'], 1):
    c = ws2.cell(row=r_gt, column=ci, value=v)
    c.font = Font(name='Arial', bold=True, size=10)
    c.fill = fill_gold
    c.border = border
    c.alignment = left_al if ci == 2 else center
    if ci == 3:
        c.number_format = '#,##0.00'

ws2.column_dimensions['A'].width = 14
ws2.column_dimensions['B'].width = 44
ws2.column_dimensions['C'].width = 16
ws2.column_dimensions['D'].width = 12

# =============================================
# ABA 3 - RESUMO POR CAMADA
# =============================================
ws3 = wb.create_sheet('RESUMO POR CAMADA')
n_cols_3 = 3 + len(bitolas) + 1
ws3['A1'] = 'RESUMO POR CAMADA - PILARES TERREO'
ws3['A1'].font = Font(name='Arial', bold=True, size=14, color='1F4E79')
ws3.merge_cells(f'A1:{get_column_letter(n_cols_3)}1')
ws3['A1'].alignment = center

hdrs3 = ['CAMADA', 'CONCRETO (m3)', 'FORMA (m2)'] + [f'ACO O{b}mm (kg)' for b in bitolas] + ['ACO TOTAL (kg)']
for ci, h in enumerate(hdrs3, 1):
    c = ws3.cell(row=3, column=ci, value=h)
    c.font = white_font
    c.fill = hdr_dark
    c.alignment = center
    c.border = border

camadas = result['CAMADA'].unique()
for ri, cam in enumerate(camadas, 4):
    sub = result[result['CAMADA'] == cam]
    aco_tot = sum(sub[f'ACO_{b}mm_kg'].sum() for b in bitolas)
    vals3 = [cam, sub['CONCRETO_m3'].sum(), sub['FORMA_m2'].sum()] + \
            [sub[f'ACO_{b}mm_kg'].sum() for b in bitolas] + [aco_tot]
    row_fill = PatternFill('solid', start_color='EBF3FB') if ri % 2 == 0 else PatternFill('solid', start_color='FFFFFF')
    for ci, v in enumerate(vals3, 1):
        c = ws3.cell(row=ri, column=ci, value=v)
        c.font = normal_font
        c.border = border
        c.fill = row_fill
        c.alignment = left_al if ci == 1 else right_al
        if ci > 1:
            c.number_format = '#,##0.00'

ri_tot3 = 4 + len(camadas)
aco_all = sum(result[f'ACO_{b}mm_kg'].sum() for b in bitolas)
totals3 = ['TOTAL GERAL', result['CONCRETO_m3'].sum(), result['FORMA_m2'].sum()] + \
          [result[f'ACO_{b}mm_kg'].sum() for b in bitolas] + [aco_all]
for ci, v in enumerate(totals3, 1):
    c = ws3.cell(row=ri_tot3, column=ci, value=v)
    c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    c.fill = fill_total
    c.border = border
    c.alignment = left_al if ci == 1 else right_al
    if ci > 1:
        c.number_format = '#,##0.00'

for ci, w in enumerate([22, 16, 16] + [14] * len(bitolas) + [16], 1):
    ws3.column_dimensions[get_column_letter(ci)].width = w

wb.save(file_out)
print(f'\nArquivo salvo: {file_out}')
