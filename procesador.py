"""
Procesador de Reposición de Inventario
Genera un .xlsx con 6 hojas siguiendo exactamente el formato de referencia.
"""
import math
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Paleta exacta del archivo de referencia ───────────────────────────────────
C_HDR_DARK  = "1F4E79"   # azul oscuro — encabezados
C_HDR_LIGHT = "BDD7EE"   # azul claro  — total saldo / total reponer
C_ROW_A     = "EBF3FA"   # fila par
C_ROW_B     = "FFFFFF"   # fila impar
C_REPONER   = "D6E4F0"   # fila Reponer en Rep+Saldo
C_SALDO_ROW = "DEE4F1"   # fila Saldo   en Rep+Saldo
C_YELLOW    = "FFFF00"   # selector días
C_MAPA2_HD  = "FFF2CC"   # cabecera y celdas Mapa2
C_MAPA2_DIN = "FFE699"   # columna dinámica sugerido Mapa2
C_OK        = "E2EFDA"
C_BAJO      = "FCE4D6"
C_SIN_VENTA = "D9D9D9"

def _fill(c):      return PatternFill("solid", fgColor=c)
def _font(bold=False, color="000000", size=11):
    return Font(bold=bold, color=color, size=size, name="Arial")
def _center():     return Alignment(horizontal="center", vertical="center", wrap_text=True)
def _left():       return Alignment(horizontal="left",   vertical="center", wrap_text=False)
def _border():
    s = Side(border_style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr(ws, coord, val, size=10):
    ws[coord].value = val
    ws[coord].font  = _font(bold=True, color="FFFFFF", size=size)
    ws[coord].fill  = _fill(C_HDR_DARK)
    ws[coord].alignment = _center()
    ws[coord].border    = _border()

def _row_fill(i):  return C_ROW_A if i % 2 == 0 else C_ROW_B

def _estado_color(e):
    return {"OK": C_OK, "BAJO": C_BAJO, "BAJA ROTACIÓN": C_BAJO,
            "SUGERIDO": C_MAPA2_HD, "SIN VENTA": C_SIN_VENTA}.get(e, C_SIN_VENTA)

# ── Parser ────────────────────────────────────────────────────────────────────
def parse_xls(filepath):
    engine = "openpyxl" if filepath.endswith(".xlsx") else "xlrd"
    df = pd.read_excel(filepath, engine=engine, header=None)
    branch_starts = {}
    for col_idx, val in enumerate(df.iloc[1, :]):
        if pd.notna(val) and str(val).strip() and col_idx > 1:
            branch_starts[str(val).strip()] = col_idx
    products = []
    for row_idx in range(3, len(df)):
        item = df.iloc[row_idx, 0]
        desc = df.iloc[row_idx, 1]
        if pd.isna(item) or str(item).startswith('R(s)'):
            continue
        prod = {'item': str(item).strip(), 'desc': str(desc).strip() if pd.notna(desc) else ''}
        for bname, bcol in branch_starts.items():
            pvtas = df.iloc[row_idx, bcol + 1]
            saldo = df.iloc[row_idx, bcol + 4]
            prod[f'{bname}__pvtas'] = float(pvtas) if pd.notna(pvtas) else 0.0
            prod[f'{bname}__saldo'] = float(saldo) if pd.notna(saldo) else 0.0
        products.append(prod)
    return products, list(branch_starts.keys())

# ── Helpers de cálculo ────────────────────────────────────────────────────────
def _avg_pvtas(p, branches):
    vals = [p[f'{b}__pvtas'] for b in branches if p[f'{b}__pvtas'] > 0]
    return round(np.mean(vals), 2) if vals else 0.0

def _locales_venta(p, branches):
    return sum(1 for b in branches if p[f'{b}__pvtas'] > 0)

def _estado_prod(pvtas_total, saldo_total, dias):
    if pvtas_total == 0: return "SIN VENTA"
    cob = saldo_total / (pvtas_total / 30)
    if cob >= dias: return "OK"
    if cob >= dias * 0.5: return "BAJA ROTACIÓN"
    return "BAJO"

def _estado_suc(pvt, sal, dias):
    if pvt == 0: return "SIN VENTA"
    cob = sal / (pvt / 30)
    if cob >= dias: return "OK"
    if cob >= dias * 0.5: return "SUGERIDO"
    return "BAJO"

# ── HOJA 1: Consolidado Reposición ───────────────────────────────────────────
def sheet_consolidado(wb, products, branches, dias):
    ws = wb.create_sheet("Consolidado Reposición")
    N = len(branches)

    ws["A1"].value = "Días de cálculo:"; ws["A1"].font = _font(bold=True, size=11)
    ws["B1"].value = dias
    ws["B1"].font  = _font(bold=True, color="0000FF", size=11)
    ws["B1"].fill  = _fill(C_YELLOW); ws["B1"].alignment = _center()
    ws["C1"].value = "← Cambia este valor para recalcular (30, 45 o 60 días)"
    ws["C1"].font  = _font(color="7F7F7F", size=10)

    _hdr(ws, "A3", "Código"); _hdr(ws, "B3", "Descripción")
    ws.row_dimensions[3].height = 34.5

    for bi, b in enumerate(branches):
        _hdr(ws, f"{get_column_letter(3+bi)}3", b)

    mapa2_col  = 3 + N
    total_col  = mapa2_col + 1
    vprom_col  = total_col + 1
    estado_col = vprom_col + 1

    ws[f"{get_column_letter(mapa2_col)}3"].value = "11 Mapa 2 (Sugerido)"
    ws[f"{get_column_letter(mapa2_col)}3"].font  = _font(bold=True, color="7F4F00", size=10)
    ws[f"{get_column_letter(mapa2_col)}3"].fill  = _fill(C_MAPA2_HD)
    ws[f"{get_column_letter(mapa2_col)}3"].alignment = _center()
    ws[f"{get_column_letter(mapa2_col)}3"].border    = _border()

    _hdr(ws, f"{get_column_letter(total_col)}3",  "Total Reponer")
    _hdr(ws, f"{get_column_letter(vprom_col)}3",  "V.Prom/Mes")
    _hdr(ws, f"{get_column_letter(estado_col)}3", "Estado")

    ws.column_dimensions["A"].width = 11
    ws.column_dimensions["B"].width = 48
    for ci in range(3, estado_col+1):
        ws.column_dimensions[get_column_letter(ci)].width = 14

    for pi, p in enumerate(products):
        row = pi + 4; rf = _row_fill(pi)
        pvt_tot = sum(p[f'{b}__pvtas'] for b in branches)
        sal_tot = sum(p[f'{b}__saldo'] for b in branches)

        def dc(col, val, bold=False, fill=None, align="center", left=False):
            c = ws.cell(row, col, val)
            c.font = _font(bold=bold, size=11)
            c.fill = _fill(fill or rf)
            c.alignment = _left() if left else _center()
            c.border = _border()

        dc(1, p['item'], bold=True)
        dc(2, p['desc'], left=True)

        for bi, b in enumerate(branches):
            pvt = p[f'{b}__pvtas']; sal = p[f'{b}__saldo']
            val = f"=MAX(0,ROUNDUP({pvt}*$B$1/30,0)-{int(sal)})" if pvt > 0 else 0
            dc(3+bi, val)

        avg = _avg_pvtas(p, branches)
        mc_val = f"=MAX(0,ROUNDUP({avg}*$B$1/30,0)-0)" if avg > 0 else "-"
        c = ws.cell(row, mapa2_col, mc_val)
        c.font = _font(bold=True, color="7F4F00", size=11)
        c.fill = _fill(C_MAPA2_HD); c.alignment = _center(); c.border = _border()

        sc = get_column_letter(3); ec = get_column_letter(2+N); mcl = get_column_letter(mapa2_col)
        tc = ws.cell(row, total_col, f"=SUM({sc}{row}:{ec}{row})+IF(ISNUMBER({mcl}{row}),{mcl}{row},0)")
        tc.font = _font(bold=True, size=11); tc.fill = _fill(C_HDR_LIGHT)
        tc.alignment = _center(); tc.border = _border()

        dc(vprom_col, int(pvt_tot) if pvt_tot == int(pvt_tot) else pvt_tot)

        est = _estado_prod(pvt_tot, sal_tot, dias)
        ec2 = ws.cell(row, estado_col, est)
        ec2.font = _font(bold=True, size=11); ec2.fill = _fill(_estado_color(est))
        ec2.alignment = _center(); ec2.border = _border()

    ws.freeze_panes = "A4"

# ── HOJA 2: Rep + Saldo + Cob + PVtas ────────────────────────────────────────
def sheet_rep_saldo(wb, products, branches, dias):
    ws = wb.create_sheet("Rep + Saldo + Cob + PVtas")
    N = len(branches)

    ws["A1"].value = "Días de cálculo:"; ws["A1"].font = _font(bold=True, size=11)
    ws["B1"].value = "='Consolidado Reposición'!B1"
    ws["B1"].font  = _font(bold=True, color="0000FF", size=11)
    ws["B1"].fill  = _fill(C_YELLOW); ws["B1"].alignment = _center()

    _hdr(ws, "A3", "Código"); _hdr(ws, "B3", "Descripción"); _hdr(ws, "C3", "Fila")
    ws.row_dimensions[3].height = 34.5

    for bi, b in enumerate(branches):
        _hdr(ws, f"{get_column_letter(4+bi)}3", b)

    mapa2_col  = 4 + N
    total_col  = mapa2_col + 1
    estado_col = total_col + 1

    ws[f"{get_column_letter(mapa2_col)}3"].value = "11 Mapa 2"
    ws[f"{get_column_letter(mapa2_col)}3"].font  = _font(bold=True, color="7F4F00", size=10)
    ws[f"{get_column_letter(mapa2_col)}3"].fill  = _fill(C_MAPA2_HD)
    ws[f"{get_column_letter(mapa2_col)}3"].alignment = _center()
    ws[f"{get_column_letter(mapa2_col)}3"].border    = _border()
    _hdr(ws, f"{get_column_letter(total_col)}3",  "Total")
    _hdr(ws, f"{get_column_letter(estado_col)}3", "Estado")

    ws.column_dimensions["A"].width = 11; ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 18
    for ci in range(4, estado_col+1):
        ws.column_dimensions[get_column_letter(ci)].width = 14

    REF = "'Consolidado Reposición'!$B$1"
    s_col = get_column_letter(4); e_col = get_column_letter(3+N); mcl = get_column_letter(mapa2_col)

    for pi, p in enumerate(products):
        base = pi * 3 + 4; rf = _row_fill(pi)
        pvt_tot = sum(p[f'{b}__pvtas'] for b in branches)
        sal_tot = sum(p[f'{b}__saldo'] for b in branches)
        avg = _avg_pvtas(p, branches)
        est = _estado_prod(pvt_tot, sal_tot
