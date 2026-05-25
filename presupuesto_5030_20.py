"""Genera una replica funcional del 'Presupuesto Mensual 50/30/20' en Excel.

Crea un .xlsx con 12 hojas (Enero..Diciembre). Cada hoja tiene:
  - Titulo del mes + Vista General (Año / Mes / Moneda)
  - Tarjetas: Ingreso Total, Total Gastado, Presupuesto a Asignar, Total Ahorrado
  - Tablas: Ingresos, Presupuesto 50/30/20, Presupuesto vs Actual,
            Facturas (gastos fijos), Gastos Variables, Seguimiento de Gastos,
            Pago de Deuda, Ahorros, Inversiones/Donaciones
  - Formulas (totales, %, reparto 50/30/20)
  - Graficos: dona 50/30/20, dona Gastos por categoria, dona Dinero Restante,
              barras Ingresos vs Gastos, barras Presupuesto vs Actual

Uso:
    pip install openpyxl
    python presupuesto_5030_20.py
Genera 'Presupuesto_Mensual_5030_20.xlsx' en la carpeta actual.
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------- Paleta de colores (pastel, aprox al original) ----------
AZUL      = "B4C7E7"   # encabezados generales / vista general
AZUL_CL   = "D9E4F2"
ROSA      = "E6A8A8"   # facturas / gastos fijos
ROSA_CL   = "F4CCCC"
NARANJA   = "F6B26B"   # gastos variables
NARANJA_CL= "FCE5CD"
MORADO    = "B4A7D6"   # pago de deuda
MORADO_CL = "D9D2E9"
VERDE     = "93C47D"   # ahorros / inversiones
VERDE_CL  = "D9EAD3"
AMARILLO  = "FFD966"   # presupuesto 50/30/20
AMARILLO_CL = "FFF2CC"
GRIS_CL   = "F3F3F3"
BLANCO    = "FFFFFF"

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

MONEY = '"$"#,##0.00'
PCT = '0%'


def fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def style_cell(ws, ref, *, value=None, bold=False, color=None, bg=None,
               align=None, fmt=None, size=None, border=True):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = Font(bold=bold, color=color or "000000", size=size or 10)
    if bg:
        c.fill = fill(bg)
    if align:
        c.alignment = align
    if fmt:
        c.number_format = fmt
    if border:
        c.border = BORDER
    return c


def section_header(ws, row, col_start, col_end, text, bg):
    """Encabezado de seccion fusionado."""
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row, end_column=col_end)
    cl = get_column_letter(col_start)
    style_cell(ws, f"{cl}{row}", value=text, bold=True, color="FFFFFF",
               bg=bg, align=CENTER, size=11)
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).border = BORDER


def col_headers(ws, row, col_start, headers, bg):
    for i, h in enumerate(headers):
        cl = get_column_letter(col_start + i)
        style_cell(ws, f"{cl}{row}", value=h, bold=True, bg=bg,
                   align=CENTER, size=9)


def money_rows(ws, r0, r1, col_start, ncols, with_first_text=True):
    """Filas vacias con formato moneda. Primera col puede ser texto (concepto)."""
    for r in range(r0, r1 + 1):
        for i in range(ncols):
            cl = get_column_letter(col_start + i)
            c = style_cell(ws, f"{cl}{r}", bg=BLANCO, align=LEFT)
            if not (with_first_text and i == 0):
                c.number_format = MONEY
                c.alignment = RIGHT


def build_month(ws, mes: str) -> None:
    # ----- Ancho de columnas -----
    widths = {
        "A": 2, "B": 16, "C": 12, "D": 12, "E": 2,
        "F": 16, "G": 12, "H": 12, "I": 11, "J": 2,
        "K": 16, "L": 12, "M": 12, "N": 11, "O": 2,
        "P": 16, "Q": 12, "R": 12, "S": 11, "T": 2,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False

    # ====== TITULO ======
    ws.merge_cells("B2:D3")
    style_cell(ws, "B2", value=mes, bold=True, size=26, align=Alignment(
        horizontal="left", vertical="center"), border=False)
    style_cell(ws, "B4", value="-Dashboard-", align=LEFT, border=False, size=9)

    # ====== VISTA GENERAL ======
    section_header(ws, 6, 2, 4, "VISTA GENERAL", AZUL)
    datos_vista = [("Año", 2026), ("Mes", mes), ("Moneda", "$")]
    for i, (k, v) in enumerate(datos_vista):
        r = 7 + i
        style_cell(ws, f"B{r}", value=k, bold=True, bg=AZUL_CL, align=LEFT)
        ws.merge_cells(f"C{r}:D{r}")
        style_cell(ws, f"C{r}", value=v, bg=BLANCO, align=CENTER)
        ws[f"D{r}"].border = BORDER

    # ====== TARJETAS (KPIs) ======
    # Definimos refs de formulas mas abajo; usamos celdas fijas.
    # Ingreso Total -> suma actual ingresos ; Total Gastado -> suma gastos
    # Las pondremos luego de conocer rangos. Placeholders por ahora.
    kpi_row = 11
    style_cell(ws, "B11", value="INGRESO TOTAL", bold=True, bg=AZUL_CL,
               align=CENTER, size=9)
    style_cell(ws, "C11", value="TOTAL GASTADO", bold=True, bg=ROSA_CL,
               align=CENTER, size=9)
    ws.merge_cells("B11:B11")

    # Reorganizamos: 2 tarjetas arriba (B/C) y 2 abajo (B/C) fila 12 y 13/14
    style_cell(ws, "D11", value="", bg=BLANCO)  # relleno
    # Valores (fila 12)
    # INGRESOS table se define en filas 16+, GASTOS en otras hojas-secciones.

    # --- Definimos las secciones primero para tener los rangos ---

    # ====== INGRESOS (col B-D) ======
    ing_hdr = 16
    section_header(ws, ing_hdr, 2, 4, "INGRESOS", AZUL)
    col_headers(ws, ing_hdr + 1, 2, ["Concepto", "Presupuesto", "Actual"], AZUL_CL)
    ing_r0, ing_r1 = ing_hdr + 2, ing_hdr + 7
    money_rows(ws, ing_r0, ing_r1, 2, 3)
    ws[f"B{ing_r0}"].value = "Sueldo"
    ws[f"B{ing_r0+1}"].value = "Alquiler"
    ing_total = ing_r1 + 1
    style_cell(ws, f"B{ing_total}", value="TOTAL", bold=True, bg=AZUL_CL)
    style_cell(ws, f"C{ing_total}", bold=True, bg=AZUL_CL, fmt=MONEY,
               value=f"=SUM(C{ing_r0}:C{ing_r1})")
    style_cell(ws, f"D{ing_total}", bold=True, bg=AZUL_CL, fmt=MONEY,
               value=f"=SUM(D{ing_r0}:D{ing_r1})")

    ingreso_actual = f"D{ing_total}"   # ingreso total real

    # ====== PRESUPUESTO 50/30/20 (col B-D) ======
    p_hdr = ing_total + 2
    section_header(ws, p_hdr, 2, 4, "PRESUPUESTO 50/30/20", AMARILLO)
    col_headers(ws, p_hdr + 1, 2, ["Categoria", "%", "Presupuesto"], AMARILLO_CL)
    cats = [("Necesidades", 0.50), ("Deseos", 0.30), ("Ahorros", 0.20)]
    p_r0 = p_hdr + 2
    for i, (cat, pct) in enumerate(cats):
        r = p_r0 + i
        style_cell(ws, f"B{r}", value=cat, bg=BLANCO, align=LEFT)
        style_cell(ws, f"C{r}", value=pct, bg=BLANCO, align=CENTER, fmt=PCT)
        style_cell(ws, f"D{r}", bg=BLANCO, align=RIGHT, fmt=MONEY,
                   value=f"={ingreso_actual}*C{r}")
    p_total = p_r0 + len(cats)
    style_cell(ws, f"B{p_total}", value="TOTAL", bold=True, bg=AMARILLO_CL)
    style_cell(ws, f"C{p_total}", value=f"=SUM(C{p_r0}:C{p_total-1})",
               bold=True, bg=AMARILLO_CL, fmt=PCT, align=CENTER)
    style_cell(ws, f"D{p_total}", value=f"=SUM(D{p_r0}:D{p_total-1})",
               bold=True, bg=AMARILLO_CL, fmt=MONEY)
    # Rango para dona 50/30/20
    p5030_cats = (p_r0, p_total - 1)

    # ====== FACTURAS / GASTOS FIJOS (col F-I) ======
    f_hdr = 16
    section_header(ws, f_hdr, 6, 9, "FACTURAS (GASTOS FIJOS)", ROSA)
    col_headers(ws, f_hdr + 1, 6, ["Concepto", "Presupuesto", "Actual", "Tipo"], ROSA_CL)
    f_r0, f_r1 = f_hdr + 2, f_hdr + 13
    for r in range(f_r0, f_r1 + 1):
        style_cell(ws, f"F{r}", bg=BLANCO, align=LEFT)
        style_cell(ws, f"G{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"H{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"I{r}", bg=BLANCO, align=CENTER)
    for r, nombre in zip(range(f_r0, f_r0 + 4), ["Renta", "Agua", "Internet", "Netflix"]):
        ws[f"F{r}"].value = nombre
    f_total = f_r1 + 1
    style_cell(ws, f"F{f_total}", value="TOTAL", bold=True, bg=ROSA_CL)
    style_cell(ws, f"G{f_total}", value=f"=SUM(G{f_r0}:G{f_r1})", bold=True, bg=ROSA_CL, fmt=MONEY)
    style_cell(ws, f"H{f_total}", value=f"=SUM(H{f_r0}:H{f_r1})", bold=True, bg=ROSA_CL, fmt=MONEY)
    style_cell(ws, f"I{f_total}", bold=True, bg=ROSA_CL)

    # ====== AHORROS (col F-I, debajo de facturas) ======
    a_hdr = f_total + 2
    section_header(ws, a_hdr, 6, 9, "AHORROS", VERDE)
    col_headers(ws, a_hdr + 1, 6, ["Concepto", "Presupuesto", "Actual", "Tipo"], VERDE_CL)
    a_r0, a_r1 = a_hdr + 2, a_hdr + 6
    for r in range(a_r0, a_r1 + 1):
        style_cell(ws, f"F{r}", bg=BLANCO, align=LEFT)
        style_cell(ws, f"G{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"H{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"I{r}", bg=BLANCO, align=CENTER)
    a_total = a_r1 + 1
    style_cell(ws, f"F{a_total}", value="TOTAL", bold=True, bg=VERDE_CL)
    style_cell(ws, f"G{a_total}", value=f"=SUM(G{a_r0}:G{a_r1})", bold=True, bg=VERDE_CL, fmt=MONEY)
    style_cell(ws, f"H{a_total}", value=f"=SUM(H{a_r0}:H{a_r1})", bold=True, bg=VERDE_CL, fmt=MONEY)
    style_cell(ws, f"I{a_total}", bold=True, bg=VERDE_CL)
    ahorro_actual = f"H{a_total}"

    # ====== GASTOS VARIABLES (col K-N) ======
    g_hdr = 16
    section_header(ws, g_hdr, 11, 14, "GASTOS VARIABLES", NARANJA)
    col_headers(ws, g_hdr + 1, 11, ["Concepto", "Presupuesto", "Actual", "Tipo"], NARANJA_CL)
    g_r0, g_r1 = g_hdr + 2, g_hdr + 15
    for r in range(g_r0, g_r1 + 1):
        style_cell(ws, f"K{r}", bg=BLANCO, align=LEFT)
        style_cell(ws, f"L{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"M{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"N{r}", bg=BLANCO, align=CENTER)
    for r, nombre in zip(range(g_r0, g_r0 + 8),
                         ["Mercado", "Transporte", "Restaurante", "Entretenimiento",
                          "Salud", "Cuidado personal", "Hogar", "Ropa"]):
        ws[f"K{r}"].value = nombre
    g_total = g_r1 + 1
    style_cell(ws, f"K{g_total}", value="TOTAL", bold=True, bg=NARANJA_CL)
    style_cell(ws, f"L{g_total}", value=f"=SUM(L{g_r0}:L{g_r1})", bold=True, bg=NARANJA_CL, fmt=MONEY)
    style_cell(ws, f"M{g_total}", value=f"=SUM(M{g_r0}:M{g_r1})", bold=True, bg=NARANJA_CL, fmt=MONEY)
    style_cell(ws, f"N{g_total}", bold=True, bg=NARANJA_CL)

    # ====== SEGUIMIENTO DE GASTOS (col P-S) ======
    s_hdr = 16
    section_header(ws, s_hdr, 16, 19, "SEGUIMIENTO DE GASTOS", AZUL)
    col_headers(ws, s_hdr + 1, 16, ["Concepto", "Monto", "Categoria", "Fecha"], AZUL_CL)
    s_r0, s_r1 = s_hdr + 2, s_hdr + 9
    for r in range(s_r0, s_r1 + 1):
        style_cell(ws, f"P{r}", bg=BLANCO, align=LEFT)
        style_cell(ws, f"Q{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"R{r}", bg=BLANCO, align=CENTER)
        style_cell(ws, f"S{r}", bg=BLANCO, align=CENTER)
    s_total = s_r1 + 1
    style_cell(ws, f"P{s_total}", value="TOTAL", bold=True, bg=AZUL_CL)
    style_cell(ws, f"Q{s_total}", value=f"=SUM(Q{s_r0}:Q{s_r1})", bold=True, bg=AZUL_CL, fmt=MONEY)
    style_cell(ws, f"R{s_total}", bold=True, bg=AZUL_CL)
    style_cell(ws, f"S{s_total}", bold=True, bg=AZUL_CL)

    # ====== PAGO DE DEUDA (col P-S, debajo) ======
    d_hdr = s_total + 2
    section_header(ws, d_hdr, 16, 19, "PAGO DE DEUDA", MORADO)
    col_headers(ws, d_hdr + 1, 16, ["Concepto", "Presupuesto", "Actual", "Tipo"], MORADO_CL)
    d_r0, d_r1 = d_hdr + 2, d_hdr + 6
    for r in range(d_r0, d_r1 + 1):
        style_cell(ws, f"P{r}", bg=BLANCO, align=LEFT)
        style_cell(ws, f"Q{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"R{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"S{r}", bg=BLANCO, align=CENTER)
    d_total = d_r1 + 1
    style_cell(ws, f"P{d_total}", value="TOTAL", bold=True, bg=MORADO_CL)
    style_cell(ws, f"Q{d_total}", value=f"=SUM(Q{d_r0}:Q{d_r1})", bold=True, bg=MORADO_CL, fmt=MONEY)
    style_cell(ws, f"R{d_total}", value=f"=SUM(R{d_r0}:R{d_r1})", bold=True, bg=MORADO_CL, fmt=MONEY)
    style_cell(ws, f"S{d_total}", bold=True, bg=MORADO_CL)

    # ====== INVERSIONES / DONACIONES (col P-S, debajo) ======
    iv_hdr = d_total + 2
    section_header(ws, iv_hdr, 16, 19, "INVERSIONES / DONACIONES", VERDE)
    col_headers(ws, iv_hdr + 1, 16, ["Concepto", "Presupuesto", "Actual", "Tipo"], VERDE_CL)
    iv_r0, iv_r1 = iv_hdr + 2, iv_hdr + 6
    for r in range(iv_r0, iv_r1 + 1):
        style_cell(ws, f"P{r}", bg=BLANCO, align=LEFT)
        style_cell(ws, f"Q{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"R{r}", bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"S{r}", bg=BLANCO, align=CENTER)
    iv_total = iv_r1 + 1
    style_cell(ws, f"P{iv_total}", value="TOTAL", bold=True, bg=VERDE_CL)
    style_cell(ws, f"Q{iv_total}", value=f"=SUM(Q{iv_r0}:Q{iv_r1})", bold=True, bg=VERDE_CL, fmt=MONEY)
    style_cell(ws, f"R{iv_total}", value=f"=SUM(R{iv_r0}:R{iv_r1})", bold=True, bg=VERDE_CL, fmt=MONEY)
    style_cell(ws, f"S{iv_total}", bold=True, bg=VERDE_CL)

    # ====== TOTALES GLOBALES PARA KPIs ======
    # Total gastado = facturas + gastos variables + deuda + inversiones (actual)
    total_gastado = (f"=H{f_total}+M{g_total}+R{d_total}+R{iv_total}")
    # Tarjetas (fila 12 valores)
    style_cell(ws, "B12", value=f"={ingreso_actual}", bold=True, bg=BLANCO,
               fmt=MONEY, align=CENTER, size=12)
    style_cell(ws, "C12", value=total_gastado, bold=True, bg=BLANCO,
               fmt=MONEY, align=CENTER, size=12)
    style_cell(ws, "B13", value="PRESUPUESTO A ASIGNAR", bold=True, bg=AMARILLO_CL,
               align=CENTER, size=9)
    style_cell(ws, "C13", value="TOTAL AHORRADO", bold=True, bg=VERDE_CL,
               align=CENTER, size=9)
    style_cell(ws, "B14", value=f"={ingreso_actual}-C12-{ahorro_actual}",
               bold=True, bg=BLANCO, fmt=MONEY, align=CENTER, size=12)
    style_cell(ws, "C14", value=f"={ahorro_actual}", bold=True, bg=BLANCO,
               fmt=MONEY, align=CENTER, size=12)
    style_cell(ws, "D12", bg=BLANCO)
    style_cell(ws, "D13", bg=BLANCO)
    style_cell(ws, "D14", bg=BLANCO)

    # ====== PRESUPUESTO VS ACTUAL (col B-D) ======
    pa_hdr = p_total + 2
    section_header(ws, pa_hdr, 2, 4, "PRESUPUESTO VS ACTUAL", AZUL)
    col_headers(ws, pa_hdr + 1, 2, ["Categoria", "Presupuesto", "Actual"], AZUL_CL)
    pa_r0 = pa_hdr + 2
    filas_pa = [
        ("Ingreso",   f"=C{ing_total}",  f"=D{ing_total}"),
        ("Gastos",    f"=L{g_total}",    f"=M{g_total}"),
        ("Facturas",  f"=G{f_total}",    f"=H{f_total}"),
        ("Ahorros",   f"=G{a_total}",    f"=H{a_total}"),
        ("Inversion", f"=Q{iv_total}",   f"=R{iv_total}"),
        ("Deuda",     f"=Q{d_total}",    f"=R{d_total}"),
    ]
    for i, (cat, pres, act) in enumerate(filas_pa):
        r = pa_r0 + i
        style_cell(ws, f"B{r}", value=cat, bg=BLANCO, align=LEFT)
        style_cell(ws, f"C{r}", value=pres, bg=BLANCO, align=RIGHT, fmt=MONEY)
        style_cell(ws, f"D{r}", value=act, bg=BLANCO, align=RIGHT, fmt=MONEY)
    pa_r1 = pa_r0 + len(filas_pa) - 1

    # ====== GRAFICOS ======
    # Dona 50/30/20 (categorias B / presupuesto D)
    d1 = DoughnutChart()
    d1.title = "50/30/20 Presupuesto"
    labels = Reference(ws, min_col=2, min_row=p5030_cats[0], max_row=p5030_cats[1])
    data = Reference(ws, min_col=4, min_row=p_hdr + 1, max_row=p5030_cats[1])
    d1.add_data(data, titles_from_data=True)
    d1.set_categories(labels)
    d1.height, d1.width = 6.5, 8
    ws.add_chart(d1, "F2")

    # Barras Ingresos vs Gastos (usa tarjetas)
    # Construimos mini tabla auxiliar oculta para el grafico
    aux_row = iv_total + 3
    style_cell(ws, f"P{aux_row}", value="Ingreso", border=False)
    style_cell(ws, f"Q{aux_row}", value=f"={ingreso_actual}", fmt=MONEY, border=False)
    style_cell(ws, f"P{aux_row+1}", value="Gastos", border=False)
    style_cell(ws, f"Q{aux_row+1}", value="=C12", fmt=MONEY, border=False)
    b1 = BarChart()
    b1.title = "Ingresos vs Gastos"
    b1.type = "col"
    bdata = Reference(ws, min_col=17, min_row=aux_row, max_row=aux_row + 1)
    blab = Reference(ws, min_col=16, min_row=aux_row, max_row=aux_row + 1)
    b1.add_data(bdata, titles_from_data=False)
    b1.set_categories(blab)
    b1.legend = None
    b1.height, b1.width = 6.5, 8
    ws.add_chart(b1, "K2")

    # Dona Gastos por categoria (gastos variables: K conceptos / M actual)
    d2 = DoughnutChart()
    d2.title = "Gastos"
    glab = Reference(ws, min_col=11, min_row=g_r0, max_row=g_r0 + 7)
    gdat = Reference(ws, min_col=13, min_row=g_hdr + 1, max_row=g_r0 + 7)
    d2.add_data(gdat, titles_from_data=True)
    d2.set_categories(glab)
    d2.height, d2.width = 6.5, 8
    ws.add_chart(d2, "P2")

    # Barras Presupuesto vs Actual
    b2 = BarChart()
    b2.title = "Presupuesto vs Actual"
    b2.type = "bar"
    b2data = Reference(ws, min_col=3, min_row=pa_hdr + 1, max_col=4, max_row=pa_r1)
    b2lab = Reference(ws, min_col=2, min_row=pa_r0, max_row=pa_r1)
    b2.add_data(b2data, titles_from_data=True)
    b2.set_categories(b2lab)
    b2.height, b2.width = 7, 9
    ws.add_chart(b2, "F18")


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    for mes in MESES:
        ws = wb.create_sheet(title=mes)
        build_month(ws, mes)
    out = "Presupuesto_Mensual_5030_20.xlsx"
    wb.save(out)
    print(f"Listo: {out} ({len(MESES)} hojas)")


if __name__ == "__main__":
    main()
