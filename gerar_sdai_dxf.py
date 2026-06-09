# -*- coding: utf-8 -*-
"""
SDAI Hospital  |  Diagrama Multifilar Típico  |  Marcos Roldão
==============================================================
Gera DXF com 4 pranchas (A1 horizontal) em model space:
  Prancha 01 – Diagrama Multifilar Geral
  Prancha 02 – Diagrama Típico de Zona
  Prancha 03 – Tabela de Cabos
  Prancha 04 – Memória de Cálculo

Requisito: pip install ezdxf
"""

import sys
import math
from datetime import date

try:
    import ezdxf
    from ezdxf import colors
    from ezdxf.enums import TextEntityAlignment
except ImportError:
    print("ERRO: pip install ezdxf")
    sys.exit(1)

OUT_FILE = r"C:\Users\PC2\Desktop\SDAI_Hospital.dxf"

# ─── HIPÓTESES ADOTADAS ───────────────────────────────────────────────────────
#  H1 : Tensão do sistema: 24 Vcc (padrão central convencional)
#  H2 : Área de cada enfermaria: 20 m² (estimada)
#  H3 : 1 detector de fumaça óptico + 1 acionador manual por zona
#  H4 : Indicador visual de alarme (IVA) incluso em 1 zona típica (Folha 02)
#  H5 : Comprimento estimado do cabo por zona: 30 m (pior caso)
#  H6 : Resistor de fim de linha (RFL): 10 kΩ
#  H7 : Corrente de supervisão por dispositivo: 50 µA
#  H8 : Corrente de alarme por dispositivo: 50 mA
#  H9 : Corrente quiescente da central: 100 mA
#  H10: Corrente da central em alarme: 500 mA
#  H11: Autonomia exigida: 24 h repouso + 30 min alarme (NBR 17240:2022 §10.4)
#  H12: Fator de segurança da bateria: 1,25
#  H13: Resistividade do cobre: 0,0175 Ω·mm²/m

# ─── DADOS DO PROJETO ─────────────────────────────────────────────────────────
N_ZONES   = 16
DET_ZONE  = 1           # detectores por zona
AM_ZONE   = 1           # acionadores manuais por zona
CABLE_T   = "Cabo Alarme Inc. 4x0,50mm² 600V Verm."
L_ZONE    = 30.0        # m  (H5)
V_SYS     = 24.0        # Vcc (H1)
I_SUP     = 50e-6       # A/dispositivo em supervisão (H7)
I_ALM_D   = 50e-3       # A/dispositivo em alarme (H8)
I_CEN_Q   = 100e-3      # A central em repouso (H9)
I_CEN_A   = 500e-3      # A central em alarme (H10)
RHO_CU    = 0.0175      # Ω·mm²/m (H13)
A_CABLE   = 0.50        # mm²
RFL_VAL   = 10000       # Ω (H6)
BAT_STBY  = 24.0        # h autonomia repouso (H11)
BAT_ALM   = 0.5         # h autonomia alarme (H11)
BAT_SF    = 1.25        # fator segurança (H12)
TODAY     = date.today().strftime("%d/%m/%Y")

# ─── CÁLCULOS ─────────────────────────────────────────────────────────────────

# 1. Dispositivos
devs_zone  = DET_ZONE + AM_ZONE
total_devs = N_ZONES * devs_zone
total_pts  = total_devs  # 1 ponto monitorado = 1 dispositivo (sistema conv.)

# 2. Cabos
L_total_raw   = N_ZONES * L_ZONE
L_total       = L_total_raw * 1.20   # +20% reserva técnica
L_total_round = math.ceil(L_total / 10) * 10  # arredonda p/ cima em 10 m

# 3. Correntes
I_sup_zona  = devs_zone  * I_SUP          # A por zona
I_sup_total = N_ZONES * I_sup_zona + I_CEN_Q
I_alm_zona  = devs_zone  * I_ALM_D        # pior caso: 1 zona em alarme
I_alm_total = I_alm_zona + I_CEN_A        # central + 1 zona (conf. NBR)

# 4. Queda de tensão
R_cond  = RHO_CU * L_ZONE / A_CABLE      # 1 condutor Ω
R_loop  = 2 * R_cond                     # ida+volta
DV      = R_loop * I_alm_zona            # Vcc
DV_PCT  = DV / V_SYS * 100

# 5. Baterias
Q_stby  = I_sup_total * BAT_STBY         # Ah repouso
Q_alm   = I_alm_total * BAT_ALM         # Ah alarme
Q_min   = (Q_stby + Q_alm) * BAT_SF    # Ah com fator segurança
# Bateria de 12V em série → 24V  — escolher capacidade >= Q_min
BAT_CAP = 7.0    # Ah padrão comercial (2 × 12V/7Ah)

# ─── DIMENSÕES DAS PRANCHAS ───────────────────────────────────────────────────
SW = 840.0    # largura A1 (mm)
SH = 594.0    # altura A1 (mm)
MG = 10.0     # margem
TB_H = 45.0   # altura do carimbo inferior

# Pranchas no model space (espaçamento de 20 mm)
P1x, P1y = 0.0,   0.0
P2x, P2y = 860.0, 0.0
P3x, P3y = 0.0,  -620.0
P4x, P4y = 860.0, -620.0


# ─── HELPERS DE DESENHO ──────────────────────────────────────────────────────

def txt(msp, text: str, x: float, y: float, h: float = 3.0,
        layer: str = "TEXTO", bold: bool = False,
        align=TextEntityAlignment.MIDDLE_CENTER):
    style = "BOLD" if bold else "Standard"
    t = msp.add_text(str(text), dxfattribs={"height": h, "layer": layer,
                                            "style": style})
    t.set_placement((x, y), align=align)


def rect(msp, x: float, y: float, w: float, h: float, layer: str = "CONTORNO"):
    pts = [(x, y), (x+w, y), (x+w, y+h), (x, y+h)]
    msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": layer})


def line(msp, x1, y1, x2, y2, layer="CONTORNO"):
    msp.add_line((x1, y1), (x2, y2), dxfattribs={"layer": layer})


def hline(msp, x, y, w, layer="FIO"):
    msp.add_line((x, y), (x+w, y), dxfattribs={"layer": layer})


def vline(msp, x, y, h, layer="FIO"):
    msp.add_line((x, y), (x, y+h), dxfattribs={"layer": layer})


# ─── SÍMBOLOS ────────────────────────────────────────────────────────────────

def sym_detector(msp, cx, cy, r=5.0, label="DF"):
    """Detector de fumaça: círculo com X"""
    lyr = "DISPOSITIVOS"
    msp.add_circle((cx, cy), r, dxfattribs={"layer": lyr})
    d = r * 0.65
    msp.add_line((cx-d, cy-d), (cx+d, cy+d), dxfattribs={"layer": lyr})
    msp.add_line((cx-d, cy+d), (cx+d, cy-d), dxfattribs={"layer": lyr})
    txt(msp, label, cx, cy+r+2, 2.5, "TEXTO")


def sym_call_point(msp, cx, cy, s=10.0, label="AM"):
    """Acionador manual: quadrado com triângulo / rótulo AM"""
    lyr = "DISPOSITIVOS"
    x, y = cx - s/2, cy - s/2
    rect(msp, x, y, s, s, lyr)
    # Triângulo interno
    msp.add_lwpolyline(
        [(cx, cy+s*0.35), (cx-s*0.30, cy-s*0.25), (cx+s*0.30, cy-s*0.25)],
        close=True, dxfattribs={"layer": lyr}
    )
    txt(msp, label, cx, cy-s/2-3, 2.5, "TEXTO")


def sym_iva(msp, cx, cy, s=8.0, label="IVA"):
    """Indicador visual de alarme: losango"""
    lyr = "DISPOSITIVOS"
    msp.add_lwpolyline(
        [(cx, cy+s/2), (cx+s/2, cy), (cx, cy-s/2), (cx-s/2, cy)],
        close=True, dxfattribs={"layer": lyr}
    )
    txt(msp, label, cx, cy-s/2-3, 2.5, "TEXTO")


def sym_eol(msp, x, y, w=15.0, h=7.0, val="10kΩ"):
    """Resistor de fim de linha: retângulo com R"""
    lyr = "DISPOSITIVOS"
    rect(msp, x, y-h/2, w, h, lyr)
    txt(msp, "RFL", x+w/2, y, 2.5, "TEXTO")
    txt(msp, val, x+w/2, y-h/2-3, 2.2, "TEXTO")


def sym_battery(msp, x, y, w=18.0, h=12.0, label="BAT"):
    lyr = "DISPOSITIVOS"
    rect(msp, x, y, w, h, lyr)
    txt(msp, label, x+w/2, y+h/2, 3.0, "TEXTO", bold=True)
    txt(msp, "12V/7Ah", x+w/2, y+2, 2.2, "TEXTO")


def sym_central(msp, x, y, w=80.0, h=60.0):
    lyr = "CENTRAL"
    rect(msp, x, y, w, h, lyr)
    txt(msp, "CENTRAL CONVENCIONAL", x+w/2, y+h-8, 3.5, "TEXTO", bold=True)
    txt(msp, "16 ZONAS", x+w/2, y+h-15, 3.0, "TEXTO")
    txt(msp, "24 Vcc", x+w/2, y+8, 2.5, "TEXTO")


# ─── CARIMBO ────────────────────────────────────────────────────────────────

def draw_stamp(msp, ox, oy, sheet_no, sheet_title):
    """Carimbo padrão ABNT na parte inferior da prancha."""
    bx, by = ox + MG, oy + MG
    bw, bh = SW - 2*MG, TB_H

    rect(msp, bx, by, bw, bh, "CARIMBO")
    # Divisórias
    line(msp, bx+bw*0.55, by, bx+bw*0.55, by+bh, "CARIMBO")
    line(msp, bx+bw*0.55, by+bh*0.55, bx+bw, by+bh*0.55, "CARIMBO")
    line(msp, bx+bw*0.55, by+bh*0.28, bx+bw, by+bh*0.28, "CARIMBO")

    txt(msp, "SISTEMA DE DETECÇÃO E ALARME DE INCÊNDIO - SDAI",
        bx+bw*0.275, by+bh*0.70, 4.5, "CARIMBO", bold=True)
    txt(msp, "HOSPITAL - ENFERMARIAS  |  SISTEMA CONVENCIONAL  |  16 ZONAS",
        bx+bw*0.275, by+bh*0.38, 3.5, "CARIMBO")
    txt(msp, sheet_title,
        bx+bw*0.275, by+bh*0.12, 3.0, "CARIMBO")

    txt(msp, f"PRANCHA {sheet_no:02d}", bx+bw*0.775, by+bh*0.82, 4.0,
        "CARIMBO", bold=True)
    txt(msp, f"REVISÃO: 00  DATA: {TODAY}", bx+bw*0.775, by+bh*0.52, 2.8,
        "CARIMBO")
    txt(msp, "RESP. TÉCNICO: ENG. ___________",
        bx+bw*0.775, by+bh*0.17, 2.5, "CARIMBO")
    # Borda principal
    rect(msp, ox+5, oy+5, SW-10, SH-10, "BORDA")


# ─── PRANCHA 01 – DIAGRAMA MULTIFILAR GERAL ──────────────────────────────────

def draw_sheet1(msp, ox, oy):
    draw_stamp(msp, ox, oy, 1, "PRANCHA 01 - DIAGRAMA MULTIFILAR GERAL")

    # Área de desenho
    dx, dy = ox + MG + 5, oy + TB_H + MG + 5
    dw, dh = SW - 2*MG - 10, SH - TB_H - 2*MG - 10 - 12

    # Título da prancha
    txt(msp, "DIAGRAMA MULTIFILAR GERAL  –  16 ZONAS",
        dx + dw/2, dy + dh + 5, 5.0, "TITULO", bold=True)

    # ── PAINEL CENTRAL ─────────────────────────────────────────────────
    pan_x = dx + 5
    pan_y = dy + 10
    pan_w = 115.0
    pan_h = dh - 15

    rect(msp, pan_x, pan_y, pan_w, pan_h, "CENTRAL")

    # Título painel
    txt(msp, "CENTRAL", pan_x+pan_w/2, pan_y+pan_h-7, 4.0,
        "TEXTO", bold=True)
    txt(msp, "CONVENCIONAL", pan_x+pan_w/2, pan_y+pan_h-14, 3.5, "TEXTO")
    txt(msp, "16 ZONAS / 24Vcc", pan_x+pan_w/2, pan_y+pan_h-21, 3.0, "TEXTO")
    line(msp, pan_x, pan_y+pan_h-26, pan_x+pan_w, pan_y+pan_h-26, "CENTRAL")

    # Seção de alimentação (parte superior do painel)
    pow_y = pan_y + pan_h - 70
    line(msp, pan_x, pow_y, pan_x+pan_w, pow_y, "CENTRAL")
    txt(msp, "ALIMENTAÇÃO", pan_x+pan_w/2, pow_y+30, 3.0, "TEXTO")
    txt(msp, "220 Vca / 60 Hz", pan_x+pan_w/2, pow_y+22, 2.8, "TEXTO")
    txt(msp, "Fonte Chaveada", pan_x+pan_w/2, pow_y+15, 2.8, "TEXTO")
    txt(msp, "24 Vcc / 3,0 A", pan_x+pan_w/2, pow_y+8, 2.8, "TEXTO")
    # Bateria dentro do painel
    sym_battery(msp, pan_x+8, pow_y-30, 40, 22, "BATERIAS")
    txt(msp, "2×12V/7Ah série", pan_x+28, pow_y-35, 2.2, "TEXTO")
    txt(msp, "BATERIAS SELADAS", pan_x+28, pow_y-40, 2.5, "TEXTO", bold=True)

    # Zona de terminais no painel
    line(msp, pan_x, pow_y-45, pan_x+pan_w, pow_y-45, "CENTRAL")
    txt(msp, "BORNES DE ZONAS", pan_x+pan_w/2, pow_y-50, 3.0,
        "TEXTO", bold=True)

    # ── ZONAS ─────────────────────────────────────────────────────────
    # Área disponível para zonas (abaixo da linha de bornes)
    zone_area_top = pan_y + pan_h - 30   # Y do topo da zona mais alta
    zone_area_bot = pan_y + 8            # Y do fundo da zona mais baixa
    zone_h_total  = zone_area_top - zone_area_bot
    zone_pitch    = zone_h_total / N_ZONES  # altura por zona

    # Largura disponível para o loop de campo
    field_x_start = pan_x + pan_w
    field_x_end   = dx + dw - 5
    dev_area_w    = field_x_end - field_x_start   # ~600 mm

    # Espaçamento dos dispositivos no campo
    dev_spacing = 45.0   # mm entre dispositivos
    wire_offset = 4.0    # mm entre os 4 condutores

    for i in range(N_ZONES):
        z_num  = i + 1
        z_cy   = zone_area_top - (i + 0.5) * zone_pitch  # centro Y da zona

        # Terminal no painel
        borne_x = pan_x + pan_w - 5
        txt(msp, f"Z{z_num:02d}+", borne_x - 12, z_cy + 3, 2.0, "BORNE")
        txt(msp, f"Z{z_num:02d}-", borne_x - 12, z_cy - 3, 2.0, "BORNE")

        # Linhas de borne para exterior do painel
        hline(msp, borne_x, z_cy + 3, 5, "FIO")
        hline(msp, borne_x, z_cy - 3, 5, "FIO")

        # Label zona à esquerda do campo
        lbl_x = field_x_start + 3
        txt(msp, f"ZONA {z_num:02d}", lbl_x + 8, z_cy + zone_pitch*0.38,
            2.5, "ZONA")
        txt(msp, f"Enf. {z_num:02d}", lbl_x + 8, z_cy + zone_pitch*0.20,
            2.2, "ZONA")

        # Quatro condutores horizontais
        wires_y = [z_cy + 1.5*wire_offset,
                   z_cy + 0.5*wire_offset,
                   z_cy - 0.5*wire_offset,
                   z_cy - 1.5*wire_offset]
        wire_labels = ["+SV", "-SV", "+AL", "-AL"]
        wire_colors  = ["", "", "", ""]  # labels apenas na folha 02

        # Início dos fios
        x0 = field_x_start + 20

        # Posições X dos dispositivos
        x_det = x0 + 60   # detector
        x_am  = x_det + dev_spacing    # acionador manual
        x_eol = x_am  + dev_spacing    # fim de linha

        # Fios da central até o detector
        for wy in wires_y:
            hline(msp, field_x_start, wy, x_det - 3 - field_x_start, "FIO")

        # Conexão vertical até detector (nó)
        vline(msp, x_det, wires_y[-1], wires_y[0] - wires_y[-1], "FIO")

        # Detector de fumaça
        sym_detector(msp, x_det, z_cy, r=min(5.0, zone_pitch*0.38))

        # Fios do detector até acionador manual
        for wy in wires_y:
            hline(msp, x_det+6, wy, x_am - x_det - 6, "FIO")

        # Nó vertical acionador
        vline(msp, x_am, wires_y[-1], wires_y[0] - wires_y[-1], "FIO")

        # Acionador manual
        sym_size = min(8.0, zone_pitch * 0.35)
        sym_call_point(msp, x_am, z_cy, s=sym_size)

        # Fios do acionador até EOL
        for wy in wires_y:
            hline(msp, x_am + sym_size/2, wy, x_eol - x_am - sym_size/2, "FIO")

        # EOL Resistor
        sym_eol(msp, x_eol, z_cy, w=14, h=6)

        # Fio de retorno após EOL
        hline(msp, x_eol + 14, z_cy, 8, "FIO")
        # Terminação (stub vertical)
        vline(msp, x_eol + 22, z_cy - 4, 8, "FIO")

        # Número do cabo acima da linha de fios
        if i == 0:   # legenda apenas na primeira zona
            y_leg = wires_y[0] + 3
            txt(msp, "+SV  -SV  +AL  -AL", x0 + 20, y_leg + 2, 2.0, "LEGENDA")
            txt(msp, f"[{CABLE_T}]", x0 + 60, y_leg + 5, 2.0, "LEGENDA")

        # Linha horizontal separadora de zona
        if i < N_ZONES - 1:
            y_sep = z_cy - zone_pitch/2
            hline(msp, field_x_start + 20, y_sep,
                  x_eol + 22 - field_x_start - 20, "SEPARADOR")

    # Legenda de símbolos
    leg_x = dx + 5
    leg_y = dy + 5
    txt(msp, "LEGENDA:", leg_x, leg_y + 12, 3.0, "LEGENDA", bold=True)
    sym_detector(msp, leg_x+10, leg_y+4, r=4)
    txt(msp, "= Det. Fumaça Óptico", leg_x+18, leg_y+4, 2.5, "LEGENDA",
        align=TextEntityAlignment.MIDDLE_LEFT)
    sym_call_point(msp, leg_x+85, leg_y+4, s=7)
    txt(msp, "= Acion. Manual", leg_x+93, leg_y+4, 2.5, "LEGENDA",
        align=TextEntityAlignment.MIDDLE_LEFT)
    sym_eol(msp, leg_x+165, leg_y+1, w=12, h=5, val="")
    txt(msp, "= Res. Fim de Linha 10kΩ", leg_x+180, leg_y+4, 2.5, "LEGENDA",
        align=TextEntityAlignment.MIDDLE_LEFT)


# ─── PRANCHA 02 – DIAGRAMA TÍPICO DE ZONA ────────────────────────────────────

def draw_sheet2(msp, ox, oy):
    draw_stamp(msp, ox, oy, 2, "PRANCHA 02 - DIAGRAMA TÍPICO DE ZONA")

    dx, dy = ox + MG + 5, oy + TB_H + MG + 5
    dw, dh = SW - 2*MG - 10, SH - TB_H - 2*MG - 10 - 15

    txt(msp, "DIAGRAMA TÍPICO DE ZONA  –  ZONA 01 / ENFERMARIA 01",
        dx+dw/2, dy+dh+5, 5.0, "TITULO", bold=True)

    # Eixo central da zona
    cy = dy + dh * 0.55

    # Posições X
    x_cen   = dx + 30    # saída da central
    x_det   = dx + 180   # detector
    x_am    = dx + 350   # acionador manual
    x_iva   = dx + 500   # indicador visual
    x_eol   = dx + 640   # EOL
    x_end   = dx + 680   # fim

    # Condutores (4 fios)
    wire_gap = 6.0
    wires = {
        "+SV": cy + 1.5*wire_gap,
        "-SV": cy + 0.5*wire_gap,
        "+AL": cy - 0.5*wire_gap,
        "-AL": cy - 1.5*wire_gap,
    }
    wire_colors_map = {
        "+SV": "PRETO",
        "-SV": "VERMELHO",
        "+AL": "VERDE",
        "-AL": "AMARELO",
    }

    # Caixa da central (fonte)
    rect(msp, x_cen-25, cy-25, 50, 50, "CENTRAL")
    txt(msp, "CENTRAL", x_cen, cy+8, 3.5, "TEXTO", bold=True)
    txt(msp, "Bornes Z01", x_cen, cy, 2.5, "TEXTO")
    txt(msp, "Zona 01", x_cen, cy-8, 2.5, "TEXTO")

    # Todos os fios do painel até o fim
    for name, wy in wires.items():
        hline(msp, x_cen+25, wy, x_eol-x_cen-25, "FIO")
        # Label na central
        txt(msp, f"B{list(wires.keys()).index(name)+1}", x_cen-5,
            wy, 2.2, "BORNE", align=TextEntityAlignment.MIDDLE_RIGHT)
        # Label cor do condutor
        txt(msp, f"{name} [{wire_colors_map[name]}]",
            x_cen+30, wy+2, 2.0, "LEGENDA")

    # ── DETECTOR DE FUMAÇA ─────────────────────────────────────────────
    # Nó nos fios +SV e -SV
    for name in ("+SV", "-SV"):
        wy = wires[name]
        msp.add_circle((x_det, wy), 1.2,
                       dxfattribs={"layer": "FIO"})  # ponto de nó
    vline(msp, x_det, wires["-SV"], wires["+SV"] - wires["-SV"], "FIO")
    sym_detector(msp, x_det, cy + 30, r=10)

    # Fio de ligação (vertical) do detector aos condutores
    vline(msp, x_det, wires["+SV"], 30 - 10 - wires["+SV"] + cy+wires["+SV"], "FIO")
    # Linha vertical do nó ao símbolo
    line(msp, x_det, wires["+SV"], x_det, cy+30-10, "FIO")

    # Caixa de terminais do detector
    term_x, term_y = x_det - 12, cy + 50
    rect(msp, term_x, term_y, 24, 18, "DISPOSITIVOS")
    txt(msp, "DFO-01", term_x+12, term_y+13, 2.5, "TEXTO", bold=True)
    txt(msp, "Detector Fumaça Óptico", term_x+12, term_y+7, 2.2, "TEXTO")
    txt(msp, "T1: +SV    T2: -SV", term_x+12, term_y+2, 2.0, "TEXTO")
    vline(msp, x_det, cy+30+10, term_y - cy - 30 - 10, "FIO")

    # ── ACIONADOR MANUAL ───────────────────────────────────────────────
    for name in ("+SV", "-SV"):
        wy = wires[name]
        msp.add_circle((x_am, wy), 1.2, dxfattribs={"layer": "FIO"})
    vline(msp, x_am, wires["-SV"], wires["+SV"] - wires["-SV"], "FIO")
    sym_call_point(msp, x_am, cy + 35, s=14)

    line(msp, x_am, wires["+SV"], x_am, cy+35-7, "FIO")

    term_x2, term_y2 = x_am - 15, cy + 60
    rect(msp, term_x2, term_y2, 30, 18, "DISPOSITIVOS")
    txt(msp, "AM-01", term_x2+15, term_y2+13, 2.5, "TEXTO", bold=True)
    txt(msp, "Acionador Manual", term_x2+15, term_y2+7, 2.2, "TEXTO")
    txt(msp, "T1: +SV    T2: -SV", term_x2+15, term_y2+2, 2.0, "TEXTO")
    vline(msp, x_am, cy+35+7, term_y2 - cy - 35 - 7, "FIO")

    # ── INDICADOR VISUAL ───────────────────────────────────────────────
    for name in ("+AL", "-AL"):
        wy = wires[name]
        msp.add_circle((x_iva, wy), 1.2, dxfattribs={"layer": "FIO"})
    vline(msp, x_iva, wires["-AL"], wires["+AL"] - wires["-AL"], "FIO")
    sym_iva(msp, x_iva, cy - 35, s=14)

    line(msp, x_iva, wires["-AL"], x_iva, cy-35+7, "FIO")

    term_x3, term_y3 = x_iva - 20, cy - 75
    rect(msp, term_x3, term_y3, 40, 18, "DISPOSITIVOS")
    txt(msp, "IVA-01", term_x3+20, term_y3+13, 2.5, "TEXTO", bold=True)
    txt(msp, "Indicador Visual Alarme", term_x3+20, term_y3+7, 2.2, "TEXTO")
    txt(msp, "T3: +AL   T4: -AL", term_x3+20, term_y3+2, 2.0, "TEXTO")
    vline(msp, x_iva, cy-35-7, term_y3+18 - (cy-35-7), "FIO")

    # ── RFL (resistor fim de linha) ─────────────────────────────────────
    sym_eol(msp, x_eol, cy, w=20, h=8, val="10 kΩ ±5%")
    hline(msp, x_eol+20, cy, 15, "FIO")
    vline(msp, x_eol+35, cy-5, 10, "FIO")  # terminação

    # Setas e notas de dimensionamento do cabo
    note_y = dy + 15
    txt(msp, f"CABO: {CABLE_T}", dx+dw/2, note_y+10, 3.0,
        "NOTA", bold=True)
    txt(msp, f"Comprimento estimado por zona: {L_ZONE:.0f} m  "
             f"|  R_total = {R_loop:.3f} Ω  "
             f"|  ΔV = {DV*1000:.1f} mV  ({DV_PCT:.2f}%  < 5% ✓)",
        dx+dw/2, note_y, 2.8, "NOTA")
    txt(msp, "CONDUTORES:  T1=+SV (supervisão+)   T2=-SV (supervisão-)   "
             "T3=+AL (alarme+)   T4=-AL (alarme-)",
        dx+dw/2, note_y - 8, 2.8, "NOTA")
    txt(msp, "Todos os condutores com seção de 0,50 mm²  "
             "|  Sistema 24 Vcc  |  RFL = 10 kΩ no último dispositivo do loop",
        dx+dw/2, note_y - 16, 2.8, "NOTA")


# ─── PRANCHA 03 – TABELA DE CABOS ────────────────────────────────────────────

def draw_sheet3(msp, ox, oy):
    draw_stamp(msp, ox, oy, 3, "PRANCHA 03 - TABELA DE CABOS")

    dx, dy = ox + MG + 5, oy + TB_H + MG + 5
    dw, dh = SW - 2*MG - 10, SH - TB_H - 2*MG - 10 - 15

    txt(msp, "TABELA DE CABOS  –  SISTEMA DE DETECÇÃO E ALARME DE INCÊNDIO",
        dx+dw/2, dy+dh+5, 5.0, "TITULO", bold=True)

    # Cabeçalho da tabela
    cols = [
        ("Nº",         18),
        ("ZONA",       25),
        ("ORIGEM",     90),
        ("DESTINO",    95),
        ("TIPO DO CABO",             190),
        ("VIAS", 20),
        ("BITOLA (mm²)", 35),
        ("COMP. (m)", 30),
        ("TAG",     35),
    ]
    total_col_w = sum(c[1] for c in cols)

    t_x  = dx + (dw - total_col_w) / 2
    t_y  = dy + dh - 5
    r_h  = 8.0      # altura de cada linha
    hdr_h = 12.0    # altura do cabeçalho

    # Cabeçalho
    cx = t_x
    rect(msp, t_x, t_y - hdr_h, total_col_w, hdr_h, "TABELA")
    for cname, cw in cols:
        txt(msp, cname, cx + cw/2, t_y - hdr_h/2, 2.5, "TABELA", bold=True)
        vline(msp, cx, t_y - hdr_h, hdr_h, "TABELA")
        cx += cw
    vline(msp, t_x + total_col_w, t_y - hdr_h, hdr_h, "TABELA")

    # Linhas de dados
    row_y = t_y - hdr_h

    cables = []
    for i in range(1, N_ZONES + 1):
        cables.append({
            "no":      str(i),
            "zona":    f"Z{i:02d}",
            "origem":  f"Central – Borne Z{i:02d}",
            "destino": f"Enf.{i:02d} – DFO+AM",
            "tipo":    CABLE_T,
            "vias":    "4",
            "bitola":  "0,50",
            "comp":    f"{L_ZONE:.0f}",
            "tag":     f"CAB-Z{i:02d}",
        })
    # Linha de alimentação da central
    cables.append({
        "no":      "17",
        "zona":    "–",
        "origem":  "QD-ILF – Disjuntor 10A",
        "destino": "Central – Alim. 220Vca",
        "tipo":    "Cabo Flexível 2x1,5mm² 750V",
        "vias":    "2+T",
        "bitola":  "1,50",
        "comp":    "20",
        "tag":     "CAB-ALI",
    })

    for row in cables:
        row_y -= r_h
        cx = t_x
        vals = [row["no"], row["zona"], row["origem"], row["destino"],
                row["tipo"], row["vias"], row["bitola"], row["comp"],
                row["tag"]]
        for (cname, cw), val in zip(cols, vals):
            txt(msp, val, cx + cw/2, row_y + r_h/2, 2.0, "TABELA")
            vline(msp, cx, row_y, r_h, "TABELA")
            cx += cw
        vline(msp, t_x + total_col_w, row_y, r_h, "TABELA")
        hline(msp, t_x, row_y, total_col_w, "TABELA")

    hline(msp, t_x, row_y, total_col_w, "TABELA")

    # Totais
    sum_y = row_y - 2*r_h
    rect(msp, t_x, sum_y, total_col_w, r_h*2, "TABELA")
    txt(msp, f"TOTAL CABO Z01-Z16 (com 20% reserva): "
             f"{L_total:.1f} m  →  Fornecer: {L_total_round:.0f} m",
        t_x + total_col_w/2, sum_y + r_h/2 + r_h*0.5, 3.0, "TABELA",
        bold=True)

    # Nota
    txt(msp,
        "NOTA: Comprimentos estimados para fins de projeto executivo. "
        "Medições definitivas deverão ser feitas em obra.",
        dx + dw/2, sum_y - 10, 2.5, "NOTA")


# ─── PRANCHA 04 – MEMÓRIA DE CÁLCULO ─────────────────────────────────────────

def draw_sheet4(msp, ox, oy):
    draw_stamp(msp, ox, oy, 4, "PRANCHA 04 - MEMÓRIA DE CÁLCULO")

    dx, dy = ox + MG + 5, oy + TB_H + MG + 5
    dw, dh = SW - 2*MG - 10, SH - TB_H - 2*MG - 10 - 15

    txt(msp, "MEMÓRIA DE CÁLCULO  –  SDAI CONVENCIONAL  –  16 ZONAS",
        dx+dw/2, dy+dh+5, 5.0, "TITULO", bold=True)

    # Duas colunas
    col1_x = dx + 5
    col2_x = dx + dw/2 + 5
    col_w  = dw/2 - 15

    def block_title(cx, cy, title):
        rect(msp, cx, cy-1, col_w, 8, "CALC_TITULO")
        txt(msp, title, cx+col_w/2, cy+3, 3.5, "CALC_TITULO", bold=True)
        return cy - 12

    def row_txt(cx, cy, label, value, unit=""):
        txt(msp, label, cx+2, cy, 2.5, "CALC",
            align=TextEntityAlignment.MIDDLE_LEFT)
        txt(msp, f"{value}  {unit}", cx+col_w-2, cy, 2.5, "CALC",
            align=TextEntityAlignment.MIDDLE_RIGHT)
        return cy - 6

    def formula(cx, cy, text):
        txt(msp, text, cx+4, cy, 2.2, "FORMULA",
            align=TextEntityAlignment.MIDDLE_LEFT)
        return cy - 5

    def note_row(cx, cy, text):
        txt(msp, text, cx+2, cy, 2.2, "NOTA",
            align=TextEntityAlignment.MIDDLE_LEFT)
        return cy - 5

    # ════════ COLUNA 1 ════════════════════════════════════════════════
    cy = dy + dh - 5

    # --- HIPÓTESES ---
    cy = block_title(col1_x, cy, "HIPÓTESES ADOTADAS")
    hyps = [
        "H1:  Tensão do sistema: 24 Vcc",
        "H2:  Área de cada enfermaria: 20 m² (estimada)",
        "H3:  1 detector de fumaça óptico + 1 acionador manual / zona",
        "H4:  Cabo de alarme de incêndio: 4x0,50mm² 600V vermelho",
        "H5:  Comprimento estimado do cabo por zona: 30 m (pior caso)",
        "H6:  Resistor de fim de linha (RFL): 10 kΩ",
        "H7:  Corrente de supervisão por dispositivo: 50 µA",
        "H8:  Corrente de alarme por dispositivo: 50 mA",
        "H9:  Corrente quiescente da central: 100 mA",
        "H10: Corrente da central em alarme: 500 mA",
        "H11: Autonomia: 24 h repouso + 30 min alarme (NBR 17240:2022 §10.4)",
        "H12: Fator de segurança da bateria: 1,25",
        "H13: Resistividade do cobre: 0,0175 Ω·mm²/m",
    ]
    for h in hyps:
        cy = note_row(col1_x, cy, h)
    cy -= 3

    # --- DIMENSIONAMENTO DAS ZONAS ---
    cy = block_title(col1_x, cy, "1. DIMENSIONAMENTO DAS ZONAS")
    cy = row_txt(col1_x, cy, "Número de zonas",           N_ZONES,       "zonas")
    cy = row_txt(col1_x, cy, "Detectores por zona",        DET_ZONE,      "und")
    cy = row_txt(col1_x, cy, "Acionadores por zona",       AM_ZONE,       "und")
    cy = row_txt(col1_x, cy, "Dispositivos por zona",      devs_zone,     "und")
    cy = row_txt(col1_x, cy, "Total de dispositivos",      total_devs,    "und")
    cy = row_txt(col1_x, cy, "Total pontos monitorados",   total_pts,     "pontos")
    cy -= 3

    # --- DIMENSIONAMENTO DOS CABOS ---
    cy = block_title(col1_x, cy, "2. DIMENSIONAMENTO DOS CABOS")
    cy = formula(col1_x, cy, f"L_zona = {L_ZONE:.0f} m  (estimado, H5)")
    cy = formula(col1_x, cy, f"L_bruto = {N_ZONES} × {L_ZONE:.0f} = {L_total_raw:.0f} m")
    cy = formula(col1_x, cy, f"L_total = {L_total_raw:.0f} × 1,20 = {L_total:.1f} m  (+20% reserva)")
    cy = row_txt(col1_x, cy, "Comprimento bruto",          f"{L_total_raw:.0f}", "m")
    cy = row_txt(col1_x, cy, "Comprimento c/ reserva",     f"{L_total:.1f}",     "m")
    cy = row_txt(col1_x, cy, "→ FORNECER",                 f"{L_total_round:.0f}", "m")
    cy = row_txt(col1_x, cy, "Tipo",                       "4x0,50mm² 600V Verm.", "")
    cy -= 3

    # --- CORRENTES ---
    cy = block_title(col1_x, cy, "3. CORRENTES DO SISTEMA")
    cy = formula(col1_x, cy, f"I_sup/zona = {devs_zone} × 50µA = {I_sup_zona*1e6:.0f} µA")
    cy = formula(col1_x, cy, f"I_sup_total = 16×{I_sup_zona*1e6:.0f}µA + 100mA = {I_sup_total*1e3:.2f} mA")
    cy = formula(col1_x, cy, f"I_alm/zona  = {devs_zone} × 50mA = {I_alm_zona*1e3:.0f} mA")
    cy = formula(col1_x, cy, f"I_alm_total = {I_alm_zona*1e3:.0f}mA + 500mA = {I_alm_total*1e3:.0f} mA")
    cy = row_txt(col1_x, cy, "Corrente em supervisão",    f"{I_sup_total*1e3:.2f}", "mA")
    cy = row_txt(col1_x, cy, "Corrente em alarme (pior caso)", f"{I_alm_total*1e3:.0f}", "mA")

    # ════════ COLUNA 2 ════════════════════════════════════════════════
    cy = dy + dh - 5

    # --- QUEDA DE TENSÃO ---
    cy = block_title(col2_x, cy, "4. QUEDA DE TENSÃO")
    cy = formula(col2_x, cy, f"ρ = {RHO_CU} Ω·mm²/m  |  L = {L_ZONE} m  |  A = {A_CABLE} mm²")
    cy = formula(col2_x, cy, f"R_cond = ρ × L / A = {RHO_CU} × {L_ZONE} / {A_CABLE}")
    cy = formula(col2_x, cy, f"       = {R_cond:.3f} Ω  (um condutor)")
    cy = formula(col2_x, cy, f"R_loop = 2 × {R_cond:.3f} = {R_loop:.3f} Ω  (ida+volta)")
    cy = formula(col2_x, cy, f"ΔV = R_loop × I_alm = {R_loop:.3f} × {I_alm_zona:.3f}")
    cy = formula(col2_x, cy, f"   = {DV*1000:.2f} mV   = {DV:.4f} Vcc")
    cy = formula(col2_x, cy, f"ΔV% = ΔV / V_sys × 100 = {DV:.4f} / {V_SYS} × 100")
    cy = formula(col2_x, cy, f"    = {DV_PCT:.3f} %")
    cy = row_txt(col2_x, cy, "ΔV absoluta",    f"{DV*1000:.2f}", "mV")
    cy = row_txt(col2_x, cy, "ΔV percentual",  f"{DV_PCT:.3f}", "%")
    cy = row_txt(col2_x, cy, "Limite NBR 17240", "≤ 5,0", "%  ✓  APROVADO")
    cy -= 3

    # --- BATERIAS ---
    cy = block_title(col2_x, cy, "5. DIMENSIONAMENTO DAS BATERIAS")
    cy = formula(col2_x, cy, "AUTONOMIA EXIGIDA (NBR 17240:2022 §10.4):")
    cy = formula(col2_x, cy, f"  Repouso: {BAT_STBY:.0f} horas")
    cy = formula(col2_x, cy, f"  Alarme:  {BAT_ALM*60:.0f} minutos")
    cy = formula(col2_x, cy, f"Q_repouso = I_sup × t = {I_sup_total*1e3:.2f}mA × {BAT_STBY:.0f}h")
    cy = formula(col2_x, cy, f"          = {Q_stby*1e3:.1f} mAh  = {Q_stby:.4f} Ah")
    cy = formula(col2_x, cy, f"Q_alarme  = I_alm × t = {I_alm_total*1e3:.0f}mA × {BAT_ALM*60:.0f}min")
    cy = formula(col2_x, cy, f"          = {Q_alm*1e3:.1f} mAh  = {Q_alm:.4f} Ah")
    cy = formula(col2_x, cy, f"Q_total   = {Q_stby:.4f} + {Q_alm:.4f} = {Q_stby+Q_alm:.4f} Ah")
    cy = formula(col2_x, cy, f"Q_min     = {Q_stby+Q_alm:.4f} × {BAT_SF} (FS) = {Q_min:.3f} Ah")
    cy = row_txt(col2_x, cy, "Capacidade mínima calculada", f"{Q_min:.2f}", "Ah")
    cy = row_txt(col2_x, cy, "→ BATERIA RECOMENDADA",
                 f"2 × 12V / {BAT_CAP:.0f} Ah  (série = 24V)", "")
    cy = row_txt(col2_x, cy, "Capacidade adotada",
                 f"2 × {BAT_CAP:.0f} = {2*BAT_CAP:.0f}", "Ah  (> {:.2f} Ah ✓)".format(Q_min))
    cy = row_txt(col2_x, cy, "Tipo",           "VLRA / AGM", "selada isenta manutenção")
    cy -= 3

    # --- QUANTITATIVOS ---
    cy = block_title(col2_x, cy, "6. QUANTITATIVO DE MATERIAIS")
    items = [
        ("Central convencional 16 zonas 24Vcc",  "1",  "und"),
        ("Detector de fumaça óptico",             str(N_ZONES * DET_ZONE), "und"),
        ("Acionador manual simples (quebra-vidro)", str(N_ZONES * AM_ZONE), "und"),
        (f"Cabo alarme 4×0,50mm² 600V vermelho",  f"{L_total_round:.0f}",   "m"),
        ("Bateria selada 12V / 7Ah VLRA",         "2",  "und"),
        ("Resistor fim de linha 10kΩ ±5% 1W",     str(N_ZONES), "und"),
        ("Quadro/rack para central",              "1",  "und"),
        ("Sirene/buzina de alarme",               "1",  "und"),
    ]
    for name, qty, unit in items:
        cy = row_txt(col2_x, cy, name, qty, unit)
    cy -= 3

    # Referência normativa
    cy = block_title(col2_x, cy, "REFERÊNCIA NORMATIVA")
    for ref in [
        "ABNT NBR 17240:2022 – Sistemas de detecção e alarme de incêndio",
        "ABNT NBR 5410:2004  – Instalações elétricas de baixa tensão",
        "ABNT NBR 13714:2000 – Hidrantes e mangotinhos para combate a incêndio",
        "IT-16 CBMESP (ou equivalente estadual) – SDAI",
    ]:
        cy = note_row(col2_x, cy, ref)


# ─── SETUP LAYERS ─────────────────────────────────────────────────────────────

def setup_layers(doc):
    layer_defs = {
        "BORDA":        (colors.WHITE,    "Continuous",  0.35),
        "CARIMBO":      (colors.WHITE,    "Continuous",  0.25),
        "TITULO":       (colors.CYAN,     "Continuous",  0.35),
        "CENTRAL":      (colors.YELLOW,   "Continuous",  0.35),
        "DISPOSITIVOS": (colors.GREEN,    "Continuous",  0.25),
        "FIO":          (colors.RED,      "Continuous",  0.18),
        "BORNE":        (colors.WHITE,    "Continuous",  0.18),
        "ZONA":         (colors.CYAN,     "Continuous",  0.18),
        "TABELA":       (colors.WHITE,    "Continuous",  0.18),
        "TEXTO":        (colors.WHITE,    "Continuous",  0.18),
        "LEGENDA":      (7,               "Continuous",  0.18),
        "NOTA":         (7,               "Continuous",  0.13),
        "CALC":         (colors.WHITE,    "Continuous",  0.18),
        "CALC_TITULO":  (colors.CYAN,     "Continuous",  0.25),
        "FORMULA":      (colors.GREEN,    "Continuous",  0.18),
        "CONTORNO":     (colors.WHITE,    "Continuous",  0.25),
        "SEPARADOR":    (7,               "DASHED",      0.13),
    }
    for lname, (col, lt, lw) in layer_defs.items():
        try:
            ly = doc.layers.add(lname, color=col)
            ly.linetype = lt
            ly.lineweight = int(lw * 100)
        except Exception:
            pass


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("Gerando SDAI_Hospital.dxf ...")

    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 4   # mm

    # Adicionar linetype DASHED
    try:
        doc.linetypes.add("DASHED", pattern=[0.5, -0.25])
    except Exception:
        pass

    # Estilo negrito
    try:
        doc.styles.add("BOLD", font="calibrib.ttf")
    except Exception:
        pass

    setup_layers(doc)
    msp = doc.modelspace()

    print("  Prancha 01 - Diagrama Multifilar Geral ...")
    draw_sheet1(msp, P1x, P1y)

    print("  Prancha 02 - Diagrama Típico de Zona ...")
    draw_sheet2(msp, P2x, P2y)

    print("  Prancha 03 - Tabela de Cabos ...")
    draw_sheet3(msp, P3x, P3y)

    print("  Prancha 04 - Memória de Cálculo ...")
    draw_sheet4(msp, P4x, P4y)

    doc.saveas(OUT_FILE)
    print(f"\nSalvo: {OUT_FILE}")

    # Resumo no console
    print("\n" + "="*60)
    print("RESUMO DOS CÁLCULOS")
    print("="*60)
    print(f"Zonas:              {N_ZONES}")
    print(f"Dispositivos/zona:  {devs_zone}  (1 DFO + 1 AM)")
    print(f"Total dispositivos: {total_devs}")
    print(f"Cabo total:         {L_total_round:.0f} m  (c/ 20% reserva)")
    print(f"I supervisão:       {I_sup_total*1e3:.2f} mA")
    print(f"I alarme (pior):    {I_alm_total*1e3:.0f} mA")
    print(f"Queda tensão:       {DV_PCT:.3f}%  (limite 5%) ✓")
    print(f"Q_bateria:          {Q_min:.2f} Ah  → 2 × 12V/7Ah ✓")
    print("="*60)


if __name__ == "__main__":
    main()
