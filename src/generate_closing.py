import os
import pandas as pd

def generate_year_end_closing():
    # Define summary values from earlier steps
    # Issued
    ventas_10_base = 1641.00
    ventas_10_cuota = 164.10
    ventas_21_base = 30261.12
    ventas_21_cuota = 6354.84
    total_ingresos_base = ventas_10_base + ventas_21_base # 31902.12
    total_ingresos_iva = ventas_10_cuota + ventas_21_cuota # 6518.94

    # Received
    gastos_21_base = 21375.82
    gastos_21_cuota = 4485.41
    axa_seguro = 470.45

    # Local Municipal Fees (extracted and summed up)
    icio_maestro_soler = 186.12
    icio_italia = 74.19
    icio_comuneros = 303.34
    licencia_maestro_soler = 12.90
    licencia_italia = 11.42
    licencia_comuneros = 21.03
    total_tasas_municipales = 609.00 # Sum of above

    total_gastos = gastos_21_base + axa_seguro + total_tasas_municipales # 22455.27
    resultado_antes_impuestos = total_ingresos_base - total_gastos # 9446.85
    impuesto_sociedades_est = round(resultado_antes_impuestos * 0.25, 2) # 2361.71
    resultado_neto = round(resultado_antes_impuestos - impuesto_sociedades_est, 2) # 7085.14

    # 1. P&L Sheet Data
    pl_data = [
        ["CUENTA DE PÉRDIDAS Y GANANCIAS ESTIMADA (2T 2026)", ""],
        ["GRUPO EXITER ESPAÑA, S.L. (B13658109)", ""],
        ["", ""],
        ["Concepto / Partida de Ingresos y Gastos", "Importe (€)"],
        ["(+) INGRESOS DE EXPLOTACIÓN", ""],
        ["  Ventas y Prestación de Servicios (Obras y Reformas)", round(total_ingresos_base, 2)],
        ["  (-) Bonificaciones y Descuentos", "0.00"],
        ["TOTAL INGRESOS NETOS (A)", round(total_ingresos_base, 2)],
        ["", ""],
        ["(-) GASTOS DE EXPLOTACIÓN", ""],
        ["  Compras y suministros (Leroy Merlin, Obramat, Pladur, Fontanería, etc.)", round(gastos_21_base, 2)],
        ["  Primas de Seguros (AXA Seguro RC Empresa)", round(axa_seguro, 2)],
        ["  Tributos y Tasas Locales (ICIO e Impuesto Licencia de Obras)", round(total_tasas_municipales, 2)],
        ["TOTAL GASTOS DE EXPLOTACIÓN (B)", round(total_gastos, 2)],
        ["", ""],
        ["(=) RESULTADO DE EXPLOTACIÓN / BENEFICIO ANTES DE IMPUESTOS (A - B)", round(resultado_antes_impuestos, 2)],
        ["  (-) Impuesto sobre Sociedades Estimado (25%)", round(impuesto_sociedades_est, 2)],
        ["(=) RESULTADO NETO DEL PERIODO (BENEFICIO NETO)", round(resultado_neto, 2)]
    ]
    df_pl = pd.DataFrame(pl_data)

    # 2. Year-End Tax Summary Draft (Modelo 390 & 190 preparation)
    annual_tax_data = [
        ["INFORME DE CONSOLIDACIÓN Y PREPARACIÓN DE IMPUESTOS ANUALES (2026)", ""],
        ["GRUPO EXITER ESPAÑA, S.L. (B13658109)", ""],
        ["", ""],
        ["1. PREPARACIÓN DEL RESUMEN ANUAL DE IVA (MODELO 390)", ""],
        ["Concepto", "Base Imponible (€)", "Cuota IVA (€)"],
        ["Operaciones en Régimen General (10%)", round(ventas_10_base, 2), round(ventas_10_cuota, 2)],
        ["Operaciones en Régimen General (21%)", round(ventas_21_base, 2), round(ventas_21_cuota, 2)],
        ["TOTAL IVA DEVENGADO ANUAL (EST.)", round(total_ingresos_base, 2), round(total_ingresos_iva, 2)],
        ["", ""],
        ["IVA Deducible en Operaciones Interiores Corrientes (21%)", round(gastos_21_base, 2), round(gastos_21_cuota, 2)],
        ["TOTAL IVA DEDUCIBLE ANUAL (EST.)", round(gastos_21_base, 2), round(gastos_21_cuota, 2)],
        ["Resultado Acumulado de la Liquidación Anual", "", round(total_ingresos_iva - gastos_21_cuota, 2)],
        ["", ""],
        ["2. PREPARACIÓN DEL RESUMEN ANUAL DE RETENCIONES IRPF (MODELO 190)", ""],
        ["Perceptor / Razón Social", "NIF", "Clave", "Percepciones (€)", "Retención (€)"],
        ["JOSÉ JOAQUÍN PELEGRÍN HERNANZ", "52411407-L", "G (Act. Prof.)", 3600.00, 540.00],
        ["TOTAL RETENCIONES ANUALES (EST.)", "", "", 3600.00, 540.00]
    ]
    df_annual_tax = pd.DataFrame(annual_tax_data)

    # Export to Excel
    output_path = "data/processed/Preparacion_Cierre_Anual_2026.xlsx"
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_pl.to_excel(writer, sheet_name='Pérdidas y Ganancias', index=False, header=False)
        df_annual_tax.to_excel(writer, sheet_name='Modelos Anuales (Borrador)', index=False, header=False)

    print(f"Year-end closing workbook successfully written to {output_path}")

if __name__ == "__main__":
    generate_year_end_closing()