import openpyxl

wb = openpyxl.load_workbook("data/interim/ANA_FEIJOO_CALCULO_2T_2026.xlsx")

# 1. Update Resumen ejecutivo sheet
sheet_exec = wb["Resumen ejecutivo"]
sheet_exec["B6"] = 122
sheet_exec["B7"] = 117
sheet_exec["B8"] = 6
sheet_exec["B9"] = 36
sheet_exec["B10"] = "=SUM('Facturas recibidas'!S2:S150)"
sheet_exec["B11"] = "=SUMIFS('Facturas recibidas'!K2:K150, 'Facturas recibidas'!X2:X150, \"Sí\")"
sheet_exec["C11"] = "=SUMIFS('Facturas recibidas'!M2:M150, 'Facturas recibidas'!X2:X150, \"Sí\")"
sheet_exec["E11"] = "=SUMIFS('Facturas recibidas'!N2:N150, 'Facturas recibidas'!X2:X150, \"Sí\")"

for r in range(1, sheet_exec.max_row + 1):
    for c in range(1, sheet_exec.max_column + 1):
        val = sheet_exec.cell(row=r, column=c).value
        if isinstance(val, str) and val.startswith("="):
            new_val = val.replace("102", "150").replace("16", "150").replace("105", "150")
            sheet_exec.cell(row=r, column=c, value=new_val)

# 2. Let us look at "IVA por tipos" sheet
sheet_iva = wb["IVA por tipos"]
sheet_iva.cell(row=2, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí")')
sheet_iva.cell(row=2, column=5, value='=SUMIFS(\'Facturas recibidas\'!M2:M150, \'Facturas recibidas\'!X2:X150, "Sí")')
sheet_iva.cell(row=2, column=7, value='=SUMIFS(\'Facturas recibidas\'!N2:N150, \'Facturas recibidas\'!X2:X150, "Sí")')

sheet_iva.cell(row=3, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!L2:L150, 0.21, \'Facturas recibidas\'!T2:T150, "<>Alquiler de local comercial sujeto a retención")')
sheet_iva.cell(row=3, column=5, value='=SUMIFS(\'Facturas recibidas\'!M2:M150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!L2:L150, 0.21, \'Facturas recibidas\'!T2:T150, "<>Alquiler de local comercial sujeto a retención")')
sheet_iva.cell(row=3, column=7, value='=SUMIFS(\'Facturas recibidas\'!N2:N150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!L2:L150, 0.21, \'Facturas recibidas\'!T2:T150, "<>Alquiler de local comercial sujeto a retención")')

sheet_iva.cell(row=4, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!L2:L150, 0.0)')
sheet_iva.cell(row=4, column=5, value='=SUMIFS(\'Facturas recibidas\'!M2:M150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!L2:L150, 0.0)')
sheet_iva.cell(row=4, column=7, value='=SUMIFS(\'Facturas recibidas\'!N2:N150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!L2:L150, 0.0)')

sheet_iva.cell(row=5, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Operación intracomunitaria con inversión sujeto pasivo")')
sheet_iva.cell(row=5, column=6, value='=C5*0.21')

sheet_iva.cell(row=6, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Servicio intracomunitario con inversión sujeto pasivo")')
sheet_iva.cell(row=6, column=6, value='=C6*0.21')

sheet_iva.cell(row=7, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!L2:L150, 0.10)')
sheet_iva.cell(row=7, column=5, value='=SUMIFS(\'Facturas recibidas\'!M2:M150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!L2:L150, 0.10)')
sheet_iva.cell(row=7, column=7, value='=SUMIFS(\'Facturas recibidas\'!N2:N150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!L2:L150, 0.10)')

sheet_iva.cell(row=8, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Servicio de proveedor no UE sin IVA español")')
sheet_iva.cell(row=8, column=6, value='=C8*0.21')

sheet_iva.cell(row=9, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Compra exterior sin IVA español")')

sheet_iva.cell(row=10, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "No")')
sheet_iva.cell(row=10, column=5, value='=SUMIFS(\'Facturas recibidas\'!M2:M150, \'Facturas recibidas\'!X2:X150, "No")')
sheet_iva.cell(row=10, column=7, value='=SUMIFS(\'Facturas recibidas\'!N2:N150, \'Facturas recibidas\'!X2:X150, "No")')

# 3. Preparación de modelos
sheet_mod = wb["Preparación de modelos"]
sheet_mod.cell(row=4, column=3, value='=\'IVA por tipos\'!C3 + 280.00')
sheet_mod.cell(row=4, column=4, value='=\'IVA por tipos\'!E3 + 58.80')

sheet_mod.cell(row=5, column=3, value='=\'IVA por tipos\'!C7')
sheet_mod.cell(row=5, column=4, value='=\'IVA por tipos\'!E7')

sheet_mod.cell(row=6, column=3, value='=\'IVA por tipos\'!C5')
sheet_mod.cell(row=6, column=4, value='=\'IVA por tipos\'!F5')

sheet_mod.cell(row=7, column=3, value='=\'IVA por tipos\'!C5')
sheet_mod.cell(row=7, column=4, value='=\'IVA por tipos\'!F5')

sheet_mod.cell(row=8, column=3, value='=\'IVA por tipos\'!C6')
sheet_mod.cell(row=8, column=4, value='=\'IVA por tipos\'!F6')

sheet_mod.cell(row=9, column=3, value='=\'IVA por tipos\'!C8')
sheet_mod.cell(row=9, column=4, value='=\'IVA por tipos\'!F8')

sheet_mod.cell(row=10, column=3, value='=\'IVA por tipos\'!C9')

sheet_mod.cell(row=12, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Retenciones profesionales")')
sheet_mod.cell(row=12, column=4, value='=SUMIFS(\'Facturas recibidas\'!S2:S150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Retenciones profesionales")')

sheet_mod.cell(row=13, column=3, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Alquiler de local comercial sujeto a retención")')
sheet_mod.cell(row=13, column=4, value='=SUMIFS(\'Facturas recibidas\'!S2:S150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Alquiler de local comercial sujeto a retención")')
sheet_mod.cell(row=13, column=6, value='Preparado')

# 4. Retenciones sheet
sheet_ret = wb["Retenciones"]
sheet_ret.cell(row=4, column=1, value="Alquiler")
sheet_ret.cell(row=4, column=2, value='=SUMIFS(\'Facturas recibidas\'!K2:K150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Alquiler de local comercial sujeto a retención")')
sheet_ret.cell(row=4, column=3, value=0.19)
sheet_ret.cell(row=4, column=4, value='=SUMIFS(\'Facturas recibidas\'!S2:S150, \'Facturas recibidas\'!X2:X150, "Sí", \'Facturas recibidas\'!T2:T150, "Alquiler de local comercial sujeto a retención")')
sheet_ret.cell(row=4, column=5, value="115")
sheet_ret.cell(row=4, column=6, value="Alquiler JUNIO 03-26 ANA BELEN FEIJOO.pdf")
sheet_ret.cell(row=4, column=7, value="Retención de alquiler de Rosa Fernández Da Bouza")

wb.save("data/interim/ANA_FEIJOO_CALCULO_2T_2026.xlsx")
print("Updated summaries.")
