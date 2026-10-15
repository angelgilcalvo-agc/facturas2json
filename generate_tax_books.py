import os, sys, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure output directory exists
output_dir = 'data/processed'
os.makedirs(output_dir, exist_ok=True)

# Data sets

# 1. ISSUED INVOICES (FACTURAS EMITIDAS)
issued_data = [
    # (Fecha, NumFactura, NIF, NombreCliente, Concepto, BaseImponible, TipoIVA, CuotaIVA, Total)
    ("08/07/2026", "F0726018", "03909382A", "MARTA ENCINAS LORENZO", "Primer pago reforma vivienda Avda. Comuneros, 39, 3ºI", 5241.00, 21.0, 6341.61),
    ("08/07/2026", "F0726019", "70769671R", "MARÍA ELISA PÉREZ RAMOS", "Primer pago pintura y varios C/ Álvaro Gil, 27, 1ºM", 745.86, 21.0, 902.49),
    ("10/07/2026", "F0726020", "07969159G", "MARÍA VICTORIA GÓMEZ RAMOS", "Primer pago reforma vivienda Avda. Comuneros, 39, 6ºO", 1589.25, 21.0, 1922.99),
    ("30/07/2026", "F0726021", "H37204567", "CP CARRETERA DE LEDESMA, 103", "Pago final suministro y rotulación puerta Ctra. Ledesma, 103", 4919.94, 10.0, 5411.93),
    ("30/07/2026", "F0726022", "E37065117", "CP PASEO DE CANALEJAS, 67-73", "Segundo pago reparación fachada y balcones P.º Canalejas 67-73", 7000.00, 10.0, 7700.00),
    ("07/08/2026", "F0826023", "03909382A", "MARTA ENCINAS LORENZO", "Segundo pago reforma vivienda Avda. Comuneros, 39, 3ºI", 3000.00, 21.0, 3630.00),
    ("11/08/2026", "F0826024", "H37257706", "CP CARRETERA DE LEDESMA, 164-170", "Pago 25% arreglos garaje Ctra. Ledesma, 164-170", 493.56, 10.0, 542.92),
    ("11/08/2026", "F0826025", "07422547X", "JUAN JOSÉ MATEOS MORENO", "Segundo pago reforma inmueble Plz. Trujillo, 17, 4ºA", 8349.90, 21.0, 10103.38),
    ("14/08/2026", "F0826026", "03909382A", "MARTA ENCINAS LORENZO", "Tercer pago reforma vivienda Avda. Comuneros, 39, 3ºI", 1945.87, 21.0, 2354.50),
    ("27/08/2026", "F0826027", "B56323140", "CALAMO CURRENTE, S.L.", "Quinto pago (2 de 2) reforma inmueble Plz. Escritores, 1-3, 1ºD", 655.25, 21.0, 792.85),
    ("27/08/2026", "F0826028", "07422547X", "JUAN JOSÉ MATEOS MORENO", "Sustitución lavabo Plz. Trujillo, 17, 4ºA", 717.85, 21.0, 868.60),
    ("01/09/2026", "F0926029", "70769671R", "MARÍA ELISA PÉREZ RAMOS", "Segundo pago pintura y varios C/ Álvaro Gil, 27, 1ºM", 908.89, 21.0, 1099.76),
    ("01/09/2026", "F0926030", "70872265L", "CONRADO ALBERTO PELEGRÍN HERNANZ", "Primer pago reforma C/ Maestro Soler, 6, 2º3", 880.00, 21.0, 1064.80),
    ("09/09/2026", "F0926031", "70769671R", "MARÍA ELISA PÉREZ RAMOS", "Reparaciones C/ Álvaro Gil, 27, 1ºM", 675.00, 21.0, 816.75)
]

# 2. RECEIVED INVOICES (FACTURAS RECIBIDAS)
received_data = [
    # (NumRegistro, Fecha, NumFactura, NIF, NombreProveedor, Concepto, BaseImponible, TipoIVA, CuotaIVA, TipoIRPF, CuotaIRPF, Total)
    (86, "02/07/2026", "14", "07976736Z", "JUAN CARLOS GONZÁLEZ CABRERA", "Pintura vivienda Plz. Trujillo, 17, 4º1ºF", 2880.00, 21.0, 604.80, 0.0, 0.0, 3484.80),
    (87, "03/07/2026", "2610114032", "A82505660", "IDEALISTA, S.A.", "Servicios portal inmobiliario Julio 2026", 195.00, 21.0, 40.95, 0.0, 0.0, 235.95),
    (88, "09/07/2026", "20.261.066/1", "70886824Y", "MÓNICA SÁNCHEZ BARBERO (EIMARSAN)", "Ventana Terradillos", 363.22, 21.0, 76.28, 0.0, 0.0, 439.50),
    (89, "10/07/2026", "022/0002/011910", "B83720753", "BRICOLAJE BRICOMAN, S.L.U. (OBRAMAT)", "Materiales yeso Terradillos", 871.30, 21.0, 182.97, 0.0, 0.0, 1054.27),
    (90, "14/07/2026", "022/0002/011977", "B83720753", "BRICOLAJE BRICOMAN, S.L.U. (OBRAMAT)", "Materiales yeso 2º Terradillos", 258.83, 21.0, 54.35, 0.0, 0.0, 313.18),
    (91, "14/07/2026", "022/0002/011979", "B83720753", "BRICOLAJE BRICOMAN, S.L.U. (OBRAMAT)", "Materiales perfiles Terradillos", 368.10, 21.0, 77.30, 0.0, 0.0, 445.40),
    (92, "14/07/2026", "022/0002/011980", "B83720753", "BRICOLAJE BRICOMAN, S.L.U. (OBRAMAT)", "Materiales hormigón Terradillos", 283.80, 21.0, 59.60, 0.0, 0.0, 343.40),
    (93, "15/07/2026", "2459", "B37427787", "MARTÍN GARCÍA 2005, S.L. (BIGMAT)", "Suministro viga Morille", 78.43, 21.0, 16.47, 0.0, 0.0, 94.90),
    (94, "16/07/2026", "20.261.069/1", "70886824Y", "MÓNICA SÁNCHEZ BARBERO (EIMARSAN)", "Ventana sin persiana Terradillos", 363.22, 21.0, 76.28, 0.0, 0.0, 439.50),
    (95, "17/07/2026", "126334", "B37000857", "HIJOS DE ISAAC LÓPEZ, S.L.", "Suministro e instalación puerta Ctra. Ledesma", 1018.00, 21.0, 213.78, 0.0, 0.0, 1231.78),
    (96, "17/07/2026", "P-067/26", "B37324531", "BLOQUE SALAMANCA, S.L.", "Piedras y granito CP Canalejas", 765.00, 21.0, 160.65, 0.0, 0.0, 925.65),
    (97, "24/07/2026", "126339", "B37000857", "HIJOS DE ISAAC LÓPEZ, S.L.", "Rotulación puerta Ctra. Ledesma", 220.00, 21.0, 46.20, 0.0, 0.0, 266.20),
    (98, "27/07/2026", "177", "52413636V", "JOSÉ RAÚL SÁNCHEZ SÁNCHEZ", "Trabajos fontanería Terradillos", 590.00, 21.0, 123.90, 0.0, 0.0, 713.90),
    (99, "27/07/2026", "178", "52413636V", "JOSÉ RAÚL SÁNCHEZ SÁNCHEZ", "Instalación manguetas inodoro Avda. Comuneros", 180.00, 21.0, 37.80, 0.0, 0.0, 217.80),
    (100, "28/07/2026", "16/26", "07850061A", "JOSÉ ÁNGEL GONZÁLEZ CABRERA", "Pladur Victoria y agujeros Marta Comuneros", 1850.00, 21.0, 388.50, 0.0, 0.0, 2238.50),
    (101, "30/07/2026", "djm26-079", "71060584Z", "LUIS HERNÁN MAIGUA YÁÑEZ (MULTISERVICIOS DJM)", "Trabajos albañilería Comuneros Marta", 3058.00, 21.0, 642.18, 0.0, 0.0, 3700.18),
    (102, "31/07/2026", "26/0209", "B37320753", "ALCASA CASTILLA, S.L.", "Mamparas Trujillo y Terradillos", 1198.35, 21.0, 251.65, 0.0, 0.0, 1450.00),
    (103, "31/07/2026", "F2026/07010", "52411407L", "JOSÉ JOAQUÍN PELEGRÍN HERNANZ", "Honorarios asesoramiento jurídico Julio 2026", 2305.00, 21.0, 484.05, 15.0, 345.75, 2443.30),
    (104, "03/08/2026", "2610141670", "A82505660", "IDEALISTA, S.A.", "Servicios portal inmobiliario Agosto 2026", 195.00, 21.0, 40.95, 0.0, 0.0, 235.95),
    (105, "07/08/2026", "F/26-00977", "A31188283", "INNOVACIONES EN COMERCIO Y MANUFACTURAS, S.A.", "Suministro pletina Avda. Comuneros 39 6ºO", 40.23, 21.0, 8.45, 0.0, 0.0, 48.68),
    (106, "17/08/2026", "08/2026", "7539079R", "MICHAEL HENRI JOSEPH PAUL", "Instalación ventanas C/ Maestro Soler 6 2º3", 300.00, 0.0, 0.0, 0.0, 0.0, 300.00),
    (107, "17/08/2026", "101", "70796823L", "MARÍA LUISA HERNÁNDEZ RIVAS (SAMA)", "Pago primera fase puerta entrada Comuneros", 1018.00, 21.0, 213.78, 0.0, 0.0, 1231.78),
    (108, "20/08/2026", "126378", "B37000857", "HIJOS DE ISAAC LÓPEZ, S.L.", "Tubos acero C/ Álvaro Gil", 225.00, 21.0, 47.25, 0.0, 0.0, 272.25),
    (109, "21/08/2026", "102", "70796823L", "MARÍA LUISA HERNÁNDEZ RIVAS (SAMA)", "Puertas y parquet Comuneros Victoria", 2100.00, 21.0, 441.00, 0.0, 0.0, 2541.00),
    (110, "21/08/2026", "022/0002/013093", "B83720753", "BRICOLAJE BRICOMAN, S.L.U. (OBRAMAT)", "Materiales vinilo Plz. Trujillo 17 4º1ºF", 109.80, 21.0, 23.06, 0.0, 0.0, 132.86),
    (111, "27/08/2026", "105-C01-013083", "A80298839", "REPSOL SOLUCIONES ENERGÉTICAS, S.A.", "Combustible vehículo Rubén", 31.74, 21.0, 6.66, 0.0, 0.0, 38.40),
    (112, "31/08/2026", "20.261.084/1", "70886824Y", "MÓNICA SÁNCHEZ BARBERO (EIMARSAN)", "Recogedores persiana Plz. Trujillo 17", 85.00, 21.0, 17.85, 0.0, 0.0, 102.85),
    (113, "31/08/2026", "F2026/08011", "52411407L", "JOSÉ JOAQUÍN PELEGRÍN HERNANZ", "Honorarios asesoramiento jurídico Agosto 2026", 1400.00, 21.0, 294.00, 15.0, 210.00, 1484.00),
    (114, "01/09/2026", "19", "07976736Z", "JUAN CARLOS GONZÁLEZ CABRERA", "Pintura escaparate DONPISO", 450.00, 21.0, 94.50, 0.0, 0.0, 544.50),
    (115, "01/09/2026", "20", "07976736Z", "JUAN CARLOS GONZÁLEZ CABRERA", "Pintura vivienda C/ Álvaro Gil", 800.00, 21.0, 168.00, 0.0, 0.0, 968.00),
    (116, "03/09/2026", "2610173620", "A82505660", "IDEALISTA, S.A.", "Servicios portal inmobiliario Septiembre 2026", 195.00, 21.0, 40.95, 0.0, 0.0, 235.95),
    (117, "15/09/2026", "S-20269-119691", "US852286309", "WATERMARKLY (AQUA SOFT)", "Licencia software marca de agua fotos", 36.32, 0.0, 0.0, 0.0, 0.0, 36.32),
    (118, "15/09/2026", "022/0002/013896", "B83720753", "BRICOLAJE BRICOMAN, S.L.U. (OBRAMAT)", "Materiales yeso cocina Plz. Escritores", 325.32, 21.0, 68.32, 0.0, 0.0, 393.64),
    (119, "15/09/2026", "3372", "B37427787", "MARTÍN GARCÍA 2005, S.L. (BIGMAT)", "Azulejos y materiales cocina Plz. Escritores", 501.90, 21.0, 105.40, 0.0, 0.0, 607.30),
    (120, "18/09/2026", "201", "52413636V", "JOSÉ RAÚL SÁNCHEZ SÁNCHEZ", "Instalación fregadero C/ Álvaro Gil", 275.00, 21.0, 57.75, 0.0, 0.0, 332.75),
    (121, "19/09/2026", "205", "52413636V", "JOSÉ RAÚL SÁNCHEZ SÁNCHEZ", "Mueble baño, espejo y grifería Plz. Trujillo", 250.00, 21.0, 52.50, 0.0, 0.0, 302.50),
    (122, "19/09/2026", "206", "52413636V", "JOSÉ RAÚL SÁNCHEZ SÁNCHEZ", "Modificación tuberías radiador Avda. Comuneros 6ºO", 162.00, 21.0, 34.02, 0.0, 0.0, 196.02),
    (123, "21/09/2026", "022/0002/014167", "B83720753", "BRICOLAJE BRICOMAN, S.L.U. (OBRAMAT)", "Extracción cocina Plz. Escritores", 109.80, 21.0, 23.06, 0.0, 0.0, 132.86),
    (124, "28/09/2026", "22/26", "07850061A", "JOSÉ ÁNGEL GONZÁLEZ CABRERA", "Pladur cocina Escritores y agujero Borrow", 1250.00, 21.0, 262.50, 0.0, 0.0, 1512.50),
    (125, "28/09/2026", "1-261319", "B37548988", "CENTROLED, S.L.U.", "Materiales electricidad cocina Plz. Escritores", 258.98, 21.0, 54.39, 0.0, 0.0, 313.37),
    (126, "29/09/2026", "djm26-087", "71060584Z", "LUIS HERNÁN MAIGUA YÁÑEZ (MULTISERVICIOS DJM)", "Trabajos albañilería cocina Escritores, Comuneros, Gil, Terradillos", 2905.00, 21.0, 610.05, 0.0, 0.0, 3515.05),
    (127, "30/09/2026", "F2026/09013", "52411407L", "JOSÉ JOAQUÍN PELEGRÍN HERNANZ", "Honorarios asesoramiento jurídico Septiembre 2026", 1400.00, 21.0, 294.00, 15.0, 210.00, 1484.00),
    (128, "19/08/2026", "022/0003/001090", "B83720753", "BRICOLAJE BRICOMAN, S.L.U. (OBRAMAT)", "Factura rectificativa / Devolución reparación Gomecello", -45.45, 21.0, -9.55, 0.0, 0.0, -55.00),
    (129, "21/09/2026", "022/0003/001198", "B83720753", "BRICOLAJE BRICOMAN, S.L.U. (OBRAMAT)", "Factura rectificativa / Devolución vinilo Obramat", -109.80, 21.0, -23.06, 0.0, 0.0, -132.86)
]

# 3. NON-INVOICE MUNICIPAL TAXES / FEES (TRIBUTOS Y TASAS MUNICIPALES)
municipal_fees = [
    # (Fecha, Tipo, Referencia, Concepto, Inmueble, Importe)
    ("06/07/2026", "ICIO", "2640450627", "Impuesto sobre Construcciones, Instalaciones y Obras", "Avda. Comuneros, 39, 6ºO", 37.10),
    ("06/07/2026", "Licencia", "2630450630", "Tasa por Licencia Urbanística", "Avda. Comuneros, 39, 6ºO", 61.83),
    ("06/07/2026", "ICIO", "2640450616", "Impuesto sobre Construcciones, Instalaciones y Obras", "C/ Álvaro Gil, 27, 1ºM", 20.00),
    ("06/07/2026", "Licencia", "2630450629", "Tasa por Licencia Urbanística", "C/ Álvaro Gil, 27, 1ºM", 61.83),
    ("28/08/2026", "ICIO", "2640578613", "Impuesto sobre Construcciones, Instalaciones y Obras", "Plz. Escritores, 1-3, 3ºB", 20.00),
    ("28/08/2026", "Licencia", "2630578601", "Tasa por Licencia Urbanística", "Plz. Escritores, 1-3, 3ºB", 61.83)
]

# Styles setup
header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

sub_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")

total_font = Font(name="Arial", size=11, bold=True, color="000000")
total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

data_font = Font(name="Arial", size=10)
thin_border_side = Side(border_style="thin", color="D9D9D9")
thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

double_bottom_side = Side(border_style="double", color="000000")
top_thin_side = Side(border_style="thin", color="000000")
total_border = Border(top=top_thin_side, bottom=double_bottom_side, left=thin_border_side, right=thin_border_side)

# Helper function to style ranges
def apply_table_styles(ws, start_row, end_row, num_cols, number_formats):
    for r in range(start_row, end_row + 1):
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = data_font
            cell.border = thin_border
            if c in number_formats:
                cell.number_format = number_formats[c]

def auto_fit_columns(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.number_format and ('#' in cell.number_format or '0' in cell.number_format):
                if isinstance(cell.value, (int, float)):
                    val_str = f"{cell.value:,.2f}"
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

# ==========================================
# 1. GENERATE LIBRO REGISTRO FACTURAS EMITIDAS
# ==========================================
wb_emitidas = openpyxl.Workbook()
ws_e = wb_emitidas.active
ws_e.title = "Facturas Emitidas 3T 2026"
ws_e.views.sheetView[0].showGridLines = True

# Title block
ws_e['A1'] = "LIBRO REGISTRO DE FACTURAS EMITIDAS - TERCER TRIMESTRE 2026"
ws_e['A1'].font = Font(name="Arial", size=14, bold=True, color="1F4E78")
ws_e['A2'] = "Sujeto Pasivo: GRUPO EXITER ESPAÑA S.L. | NIF: B-13658109"
ws_e['A2'].font = Font(name="Arial", size=11, italic=True)

headers_e = [
    "Fecha Exp.", "Nº Factura", "NIF Cliente", "Nombre / Razón Social", "Concepto / Inmueble",
    "Base Imponible (€)", "% IVA", "Cuota IVA (€)", "Total Factura (€)"
]

ws_e.append([]) # Row 3 empty
ws_e.append(headers_e) # Row 4 headers

for c_idx, h in enumerate(headers_e, start=1):
    cell = ws_e.cell(row=4, column=c_idx)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

start_r = 5
for row_data in issued_data:
    fecha, num_fact, nif, cliente, concepto, base, tipo_iva, total = row_data
    # Append row with formula for cuota_iva = Base * TipoIVA / 100
    r_idx = ws_e.max_row + 1
    ws_e.cell(row=r_idx, column=1, value=fecha).alignment = Alignment(horizontal="center")
    ws_e.cell(row=r_idx, column=2, value=str(num_fact)).number_format = '@'
    ws_e.cell(row=r_idx, column=2).alignment = Alignment(horizontal="center")
    ws_e.cell(row=r_idx, column=3, value=str(nif)).number_format = '@'
    ws_e.cell(row=r_idx, column=3).alignment = Alignment(horizontal="center")
    ws_e.cell(row=r_idx, column=4, value=cliente)
    ws_e.cell(row=r_idx, column=5, value=concepto)
    ws_e.cell(row=r_idx, column=6, value=base)
    ws_e.cell(row=r_idx, column=7, value=tipo_iva / 100.0)
    ws_e.cell(row=r_idx, column=8, value=f"=F{r_idx}*G{r_idx}")
    ws_e.cell(row=r_idx, column=9, value=total)

end_r = ws_e.max_row

# Apply styles to data rows
formats_e = {
    6: "#,##0.00 €",
    7: "0.0%",
    8: "#,##0.00 €",
    9: "#,##0.00 €"
}
apply_table_styles(ws_e, start_r, end_r, len(headers_e), formats_e)

# Total Row
tot_r = end_r + 1
ws_e.cell(row=tot_r, column=1, value="TOTALES").font = total_font
ws_e.cell(row=tot_r, column=6, value=f"=SUM(F{start_r}:F{end_r})")
ws_e.cell(row=tot_r, column=8, value=f"=SUM(H{start_r}:H{end_r})")
ws_e.cell(row=tot_r, column=9, value=f"=SUM(I{start_r}:I{end_r})")

for c_idx in range(1, len(headers_e) + 1):
    cell = ws_e.cell(row=tot_r, column=c_idx)
    cell.font = total_font
    cell.fill = total_fill
    cell.border = total_border
    if c_idx in [6, 8, 9]:
        cell.number_format = "#,##0.00 €"

# Freeze panes & autofilter
ws_e.freeze_panes = "A5"
ws_e.auto_filter.ref = f"A4:I{end_r}"
auto_fit_columns(ws_e)

wb_emitidas.save(os.path.join(output_dir, "Libro_Registro_Facturas_Emitidas_3T_2026.xlsx"))
print("Saved Libro_Registro_Facturas_Emitidas_3T_2026.xlsx")


# ==========================================
# 2. GENERATE LIBRO REGISTRO FACTURAS RECIBIDAS
# ==========================================
wb_recibidas = openpyxl.Workbook()
ws_r = wb_recibidas.active
ws_r.title = "Facturas Recibidas 3T 2026"
ws_r.views.sheetView[0].showGridLines = True

ws_r['A1'] = "LIBRO REGISTRO DE FACTURAS RECIBIDAS - TERCER TRIMESTRE 2026"
ws_r['A1'].font = Font(name="Arial", size=14, bold=True, color="1F4E78")
ws_r['A2'] = "Sujeto Pasivo: GRUPO EXITER ESPAÑA S.L. | NIF: B-13658109"
ws_r['A2'].font = Font(name="Arial", size=11, italic=True)

headers_r = [
    "Nº Reg.", "Fecha Reg.", "Nº Factura Provedor", "NIF Proveedor", "Nombre / Razón Social", "Concepto / Destino",
    "Base Imponible (€)", "% IVA", "Cuota IVA (€)", "% IRPF", "Retención IRPF (€)", "Total Factura (€)"
]

ws_r.append([])
ws_r.append(headers_r)

for c_idx, h in enumerate(headers_r, start=1):
    cell = ws_r.cell(row=4, column=c_idx)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

start_r = 5
for row_data in received_data:
    n_reg, fecha, num_fact, nif, prov, concepto, base, tipo_iva, cuota_iva, tipo_irpf, cuota_irpf, total = row_data
    r_idx = ws_r.max_row + 1
    ws_r.cell(row=r_idx, column=1, value=n_reg).alignment = Alignment(horizontal="center")
    ws_r.cell(row=r_idx, column=2, value=fecha).alignment = Alignment(horizontal="center")
    ws_r.cell(row=r_idx, column=3, value=str(num_fact)).number_format = '@'
    ws_r.cell(row=r_idx, column=3).alignment = Alignment(horizontal="center")
    ws_r.cell(row=r_idx, column=4, value=str(nif)).number_format = '@'
    ws_r.cell(row=r_idx, column=4).alignment = Alignment(horizontal="center")
    ws_r.cell(row=r_idx, column=5, value=prov)
    ws_r.cell(row=r_idx, column=6, value=concepto)
    ws_r.cell(row=r_idx, column=7, value=base)
    ws_r.cell(row=r_idx, column=8, value=tipo_iva / 100.0)
    ws_r.cell(row=r_idx, column=9, value=f"=G{r_idx}*H{r_idx}")
    ws_r.cell(row=r_idx, column=10, value=tipo_irpf / 100.0)
    ws_r.cell(row=r_idx, column=11, value=f"=G{r_idx}*J{r_idx}")
    ws_r.cell(row=r_idx, column=12, value=total)

end_r = ws_r.max_row

formats_r = {
    7: "#,##0.00 €",
    8: "0.0%",
    9: "#,##0.00 €",
    10: "0.0%",
    11: "#,##0.00 €",
    12: "#,##0.00 €"
}
apply_table_styles(ws_r, start_r, end_r, len(headers_r), formats_r)

# Total Row
tot_r = end_r + 1
ws_r.cell(row=tot_r, column=1, value="TOTALES").font = total_font
ws_r.cell(row=tot_r, column=7, value=f"=SUM(G{start_r}:G{end_r})")
ws_r.cell(row=tot_r, column=9, value=f"=SUM(I{start_r}:I{end_r})")
ws_r.cell(row=tot_r, column=11, value=f"=SUM(K{start_r}:K{end_r})")
ws_r.cell(row=tot_r, column=12, value=f"=SUM(L{start_r}:L{end_r})")

for c_idx in range(1, len(headers_r) + 1):
    cell = ws_r.cell(row=tot_r, column=c_idx)
    cell.font = total_font
    cell.fill = total_fill
    cell.border = total_border
    if c_idx in [7, 9, 11, 12]:
        cell.number_format = "#,##0.00 €"

ws_r.freeze_panes = "A5"
ws_r.auto_filter.ref = f"A4:L{end_r}"
auto_fit_columns(ws_r)

wb_recibidas.save(os.path.join(output_dir, "Libro_Registro_Facturas_Recibidas_3T_2026.xlsx"))
print("Saved Libro_Registro_Facturas_Recibidas_3T_2026.xlsx")


# ==========================================
# 3. GENERATE RESUMEN FISCAL MODELO 303 & MODELO 111
# ==========================================
wb_summary = openpyxl.Workbook()

# Sheet 1: Liquidacion Modelo 303 (IVA)
ws_303 = wb_summary.active
ws_303.title = "Modelo 303 (IVA 3T 2026)"
ws_303.views.sheetView[0].showGridLines = True

ws_303['A1'] = "RESUMEN FISCAL MODELO 303 - LIQUIDACIÓN DE IVA (3T 2026)"
ws_303['A1'].font = Font(name="Arial", size=14, bold=True, color="1F4E78")
ws_303['A2'] = "Sujeto Pasivo: GRUPO EXITER ESPAÑA S.L. | NIF: B-13658109"
ws_303['A2'].font = Font(name="Arial", size=11, italic=True)

# Section 1: IVA DEVENGADO
ws_303['A4'] = "1. IVA DEVENGADO (Facturas Emitidas)"
ws_303['A4'].font = Font(name="Arial", size=12, bold=True, color="1F4E78")

headers_303_dev = ["Concepto / Tipo Impuesto", "Base Imponible (€)", "Tipo IVA", "Cuota Devengada (€)"]
ws_303.cell(row=5, column=1, value=headers_303_dev[0])
ws_303.cell(row=5, column=2, value=headers_303_dev[1])
ws_303.cell(row=5, column=3, value=headers_303_dev[2])
ws_303.cell(row=5, column=4, value=headers_303_dev[3])

for col in range(1, 5):
    cell = ws_303.cell(row=5, column=col)
    cell.font = header_font
    cell.fill = sub_header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Rows for 21% and 10%
# Base 21%: Sum of 21% issued base
# Base 10%: Sum of 10% issued base
dev_21_base = sum(x[5] for x in issued_data if x[6] == 21.0)
dev_10_base = sum(x[5] for x in issued_data if x[6] == 10.0)

ws_303.cell(row=6, column=1, value="Operaciones Régimen General (21%)")
ws_303.cell(row=6, column=2, value=dev_21_base)
ws_303.cell(row=6, column=3, value=0.21)
ws_303.cell(row=6, column=4, value="=B6*C6")

ws_303.cell(row=7, column=1, value="Operaciones Régimen General (10%)")
ws_303.cell(row=7, column=2, value=dev_10_base)
ws_303.cell(row=7, column=3, value=0.10)
ws_303.cell(row=7, column=4, value="=B7*C7")

ws_303.cell(row=8, column=1, value="TOTAL IVA DEVENGADO").font = total_font
ws_303.cell(row=8, column=2, value="=SUM(B6:B7)").font = total_font
ws_303.cell(row=8, column=4, value="=SUM(D6:D7)").font = total_font

apply_table_styles(ws_303, 6, 7, 4, {2: "#,##0.00 €", 3: "0.0%", 4: "#,##0.00 €"})
for c in range(1, 5):
    cell = ws_303.cell(row=8, column=c)
    cell.fill = total_fill
    cell.border = total_border
    if c in [2, 4]:
        cell.number_format = "#,##0.00 €"

# Section 2: IVA DEDUCIBLE
ws_303['A10'] = "2. IVA DEDUCIBLE (Facturas Recibidas)"
ws_303['A10'].font = Font(name="Arial", size=12, bold=True, color="1F4E78")

ws_303.cell(row=11, column=1, value="Concepto / Tipo Impuesto")
ws_303.cell(row=11, column=2, value="Base Imponible (€)")
ws_303.cell(row=11, column=3, value="Tipo IVA")
ws_303.cell(row=11, column=4, value="Cuota Deducible (€)")

for col in range(1, 5):
    cell = ws_303.cell(row=11, column=col)
    cell.font = header_font
    cell.fill = sub_header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

ded_21_base = sum(x[6] for x in received_data if x[7] == 21.0)
ded_0_base = sum(x[6] for x in received_data if x[7] == 0.0)

ws_303.cell(row=12, column=1, value="Operaciones Interiores Corrientes (21%)")
ws_303.cell(row=12, column=2, value=ded_21_base)
ws_303.cell(row=12, column=3, value=0.21)
ws_303.cell(row=12, column=4, value="=B12*C12")

ws_303.cell(row=13, column=1, value="Operaciones Exentas / Sin IVA (0%)")
ws_303.cell(row=13, column=2, value=ded_0_base)
ws_303.cell(row=13, column=3, value=0.00)
ws_303.cell(row=13, column=4, value="=B13*C13")

ws_303.cell(row=14, column=1, value="TOTAL IVA DEDUCIBLE").font = total_font
ws_303.cell(row=14, column=2, value="=SUM(B12:B13)").font = total_font
ws_303.cell(row=14, column=4, value="=SUM(D12:D13)").font = total_font

apply_table_styles(ws_303, 12, 13, 4, {2: "#,##0.00 €", 3: "0.0%", 4: "#,##0.00 €"})
for c in range(1, 5):
    cell = ws_303.cell(row=14, column=c)
    cell.fill = total_fill
    cell.border = total_border
    if c in [2, 4]:
        cell.number_format = "#,##0.00 €"

# Section 3: RESULTADO LIQUIDACION MODELO 303
ws_303['A16'] = "3. RESULTADO DE LA LIQUIDACIÓN (MODELO 303)"
ws_303['A16'].font = Font(name="Arial", size=12, bold=True, color="1F4E78")

ws_303.cell(row=17, column=1, value="Total IVA Devengado")
ws_303.cell(row=17, column=4, value="=D8")
ws_303.cell(row=18, column=1, value="Total IVA Deducible")
ws_303.cell(row=18, column=4, value="=D14")

ws_303.cell(row=19, column=1, value="RESULTADO A INGRESAR EN HACIENDA (€)").font = Font(name="Arial", size=11, bold=True, color="006100")
ws_303.cell(row=19, column=4, value="=D17-D18").font = Font(name="Arial", size=11, bold=True, color="006100")

apply_table_styles(ws_303, 17, 18, 4, {4: "#,##0.00 €"})
cell_res = ws_303.cell(row=19, column=4)
cell_res.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
cell_res.border = total_border
cell_res.number_format = "#,##0.00 €"
ws_303.cell(row=19, column=1).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
ws_303.cell(row=19, column=1).border = total_border

auto_fit_columns(ws_303)

# Sheet 2: Retenciones Modelo 111 (IRPF)
ws_111 = wb_summary.create_sheet(title="Modelo 111 (IRPF 3T 2026)")
ws_111.views.sheetView[0].showGridLines = True

ws_111['A1'] = "RESUMEN FISCAL MODELO 111 - RETENCIONES IRPF ACTIVIDADES PROFESIONALES (3T 2026)"
ws_111['A1'].font = Font(name="Arial", size=14, bold=True, color="1F4E78")
ws_111['A2'] = "Sujeto Pasivo: GRUPO EXITER ESPAÑA S.L. | NIF: B-13658109"
ws_111['A2'].font = Font(name="Arial", size=11, italic=True)

headers_111 = ["Nº Reg.", "Fecha", "Nº Factura", "NIF Perceptor", "Nombre Profesional / Perceptor", "Base Retención (€)", "% Retención", "Importe Retención (€)"]

ws_111.append([])
ws_111.append(headers_111)

for col, h in enumerate(headers_111, start=1):
    cell = ws_111.cell(row=4, column=col)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Professional retention records (José Joaquín Pelegrín)
irpf_data = [x for x in received_data if x[9] > 0]

start_r = 5
for row_data in irpf_data:
    n_reg, fecha, num_fact, nif, prov, concepto, base, tipo_iva, cuota_iva, tipo_irpf, cuota_irpf, total = row_data
    r_idx = ws_111.max_row + 1
    ws_111.cell(row=r_idx, column=1, value=n_reg).alignment = Alignment(horizontal="center")
    ws_111.cell(row=r_idx, column=2, value=fecha).alignment = Alignment(horizontal="center")
    ws_111.cell(row=r_idx, column=3, value=str(num_fact)).number_format = '@'
    ws_111.cell(row=r_idx, column=3).alignment = Alignment(horizontal="center")
    ws_111.cell(row=r_idx, column=4, value=str(nif)).number_format = '@'
    ws_111.cell(row=r_idx, column=4).alignment = Alignment(horizontal="center")
    ws_111.cell(row=r_idx, column=5, value=prov)
    ws_111.cell(row=r_idx, column=6, value=base)
    ws_111.cell(row=r_idx, column=7, value=tipo_irpf / 100.0)
    ws_111.cell(row=r_idx, column=8, value=f"=F{r_idx}*G{r_idx}")

end_r = ws_111.max_row
apply_table_styles(ws_111, start_r, end_r, len(headers_111), {6: "#,##0.00 €", 7: "0.0%", 8: "#,##0.00 €"})

tot_r = end_r + 1
ws_111.cell(row=tot_r, column=1, value="TOTAL A INGRESAR (MODELO 111)").font = total_font
ws_111.cell(row=tot_r, column=6, value=f"=SUM(F{start_r}:F{end_r})").font = total_font
ws_111.cell(row=tot_r, column=8, value=f"=SUM(H{start_r}:H{end_r})").font = total_font

for c in range(1, len(headers_111) + 1):
    cell = ws_111.cell(row=tot_r, column=c)
    cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    cell.border = total_border
    if c in [6, 8]:
        cell.number_format = "#,##0.00 €"

auto_fit_columns(ws_111)

# Sheet 3: Tributos y Tasas Municipales (Sin IVA)
ws_tasas = wb_summary.create_sheet(title="Tributos Municipales 3T 2026")
ws_tasas.views.sheetView[0].showGridLines = True

ws_tasas['A1'] = "REGISTRO DE TRIBUTOS Y TASAS MUNICIPALES (ICIO / LICENCIAS) - 3T 2026"
ws_tasas['A1'].font = Font(name="Arial", size=14, bold=True, color="1F4E78")
ws_tasas['A2'] = "Sujeto Pasivo: GRUPO EXITER ESPAÑA S.L. | NIF: B-13658109 (Gastos no sujetosas / tributarios sin IVA)"
ws_tasas['A2'].font = Font(name="Arial", size=11, italic=True)

headers_tasas = ["Fecha Pago", "Tipo de Tributo", "Referencia / Justificante", "Concepto Tributario", "Inmueble / Emplazamiento", "Importe Pagado (€)"]

ws_tasas.append([])
ws_tasas.append(headers_tasas)

for col, h in enumerate(headers_tasas, start=1):
    cell = ws_tasas.cell(row=4, column=col)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

start_r = 5
for row_data in municipal_fees:
    fecha, tipo, ref, concepto, inmueble, importe = row_data
    r_idx = ws_tasas.max_row + 1
    ws_tasas.cell(row=r_idx, column=1, value=fecha).alignment = Alignment(horizontal="center")
    ws_tasas.cell(row=r_idx, column=2, value=tipo).alignment = Alignment(horizontal="center")
    ws_tasas.cell(row=r_idx, column=3, value=str(ref)).number_format = '@'
    ws_tasas.cell(row=r_idx, column=3).alignment = Alignment(horizontal="center")
    ws_tasas.cell(row=r_idx, column=4, value=concepto)
    ws_tasas.cell(row=r_idx, column=5, value=inmueble)
    ws_tasas.cell(row=r_idx, column=6, value=importe)

end_r = ws_tasas.max_row
apply_table_styles(ws_tasas, start_r, end_r, len(headers_tasas), {6: "#,##0.00 €"})

tot_r = end_r + 1
ws_tasas.cell(row=tot_r, column=1, value="TOTAL TRIBUTOS Y TASAS MUNICIPALES").font = total_font
ws_tasas.cell(row=tot_r, column=6, value=f"=SUM(F{start_r}:F{end_r})").font = total_font

for c in range(1, len(headers_tasas) + 1):
    cell = ws_tasas.cell(row=tot_r, column=c)
    cell.fill = total_fill
    cell.border = total_border
    if c == 6:
        cell.number_format = "#,##0.00 €"

auto_fit_columns(ws_tasas)

wb_summary.save(os.path.join(output_dir, "Resumen_Fiscal_Modelo303_111_3T_2026.xlsx"))
print("Saved Resumen_Fiscal_Modelo303_111_3T_2026.xlsx")
