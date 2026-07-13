import os
import pandas as pd
from clean_data import clean_and_process_data

def generate_quarterly_models():
    # Load clean data
    emitidas, recibidas = clean_and_process_data()

    # Calculate Modelo 303 (IVA) Values
    # 1. Devengado (Accrued) - Grouped by rate
    emitidas_grouped = emitidas.groupby('IVA REPERCUTIDO')[['BASE IMPONIBLE', 'RETENCIÓN APLICADA']].sum().reset_index()

    # 2. Deducible (Deductible) - Grouped by rate
    recibidas_taxable = recibidas[recibidas['BASE IMPONIBLE'].notna() & (recibidas['BASE IMPONIBLE'] > 0)]
    recibidas_grouped = recibidas_taxable.groupby('IVA SPORTADO')[['BASE IMPONIBLE', 'IMPORTE IVA']].sum().reset_index()

    # Total Gasto Exento (AXA Insurance)
    gasto_exento = recibidas['GASTOS SIN IVA'].sum()

    # Prepare Modelo 303 structure
    m303_data = []

    m303_data.append(["MODELO 303 - LIQUIDACIÓN MENSUAL/TRIMESTRAL DE IVA", ""])
    m303_data.append(["GRUPO EXITER ESPAÑA, S.L. (B13658109)", ""])
    m303_data.append(["EJERCICIO: 2026", "PERIODO: 2T (Segundo Trimestre)"])
    m303_data.append(["", ""])

    m303_data.append(["I. IVA DEVENGADO (REPERCUTIDO)", ""])
    m303_data.append(["Concepto", "Base Imponible (€)", "Tipo (%)", "Cuota (€)"])

    total_dev_base = 0.0
    total_dev_cuota = 0.0
    for _, row in emitidas_grouped.iterrows():
        base = row['BASE IMPONIBLE']
        rate = row['IVA REPERCUTIDO'] * 100
        quota = row['RETENCIÓN APLICADA']
        m303_data.append([f"Régimen General {int(rate)}%", round(base, 2), f"{int(rate)}%", round(quota, 2)])
        total_dev_base += base
        total_dev_cuota += quota

    m303_data.append(["Total IVA Devengado", round(total_dev_base, 2), "", round(total_dev_cuota, 2)])
    m303_data.append(["", ""])

    m303_data.append(["II. IVA DEDUCIBLE (SOPORTADO)", ""])
    m303_data.append(["Concepto", "Base Imponible (€)", "Tipo (%)", "Cuota (€)"])

    total_ded_base = 0.0
    total_ded_cuota = 0.0
    for _, row in recibidas_grouped.iterrows():
        base = row['BASE IMPONIBLE']
        rate = row['IVA SPORTADO'] * 100
        quota = row['IMPORTE IVA']
        m303_data.append([f"Operaciones Interiores Corrientes {int(rate)}%", round(base, 2), f"{int(rate)}%", round(quota, 2)])
        total_ded_base += base
        total_ded_cuota += quota

    # Gasto exento is not reported in Modelo 303 as deductible VAT, but can be noted
    m303_data.append(["Gasto Exento / Sin IVA", round(gasto_exento, 2), "0%", "0.00"])
    m303_data.append(["Total IVA Deducible", round(total_ded_base, 2), "", round(total_ded_cuota, 2)])
    m303_data.append(["", ""])

    # Liquidación (Result)
    result = round(total_dev_cuota - total_ded_cuota, 2)
    m303_data.append(["III. RESULTADO DE LA LIQUIDACIÓN", ""])
    m303_data.append(["[Casilla 27] Total Cuota IVA Devengado", round(total_dev_cuota, 2)])
    m303_data.append(["[Casilla 45] Total IVA Deducible", round(total_ded_cuota, 2)])
    m303_data.append(["[Casilla 46] Diferencia (IVA Devengado - IVA Deducible)", result])
    m303_data.append(["[Casilla 71] RESULTADO DE LA LIQUIDACIÓN (€)", result])
    m303_data.append(["Resultado final", "A INGRESAR" if result > 0 else "A COMPENSAR"])

    df_m303 = pd.DataFrame(m303_data)

    # Calculate Modelo 111 (Retenciones IRPF) Values
    # Compile the lawyer minutas (abogado Pelegrín)
    abogado_rows = recibidas[recibidas['PROVEEDOR'].str.contains('PELEGRÍN') | recibidas['NIF'].str.contains('52411407')]
    num_perceptores = len(abogado_rows)
    total_bi_irpf = abogado_rows['BASE IMPONIBLE'].sum()
    total_ret_irpf = round(total_bi_irpf * 0.15, 2)

    m111_data = []
    m111_data.append(["MODELO 111 - RETENCIONES E INGRESOS A CUENTA DEL IRPF", ""])
    m111_data.append(["GRUPO EXITER ESPAÑA, S.L. (B13658109)", ""])
    m111_data.append(["EJERCICIO: 2026", "PERIODO: 2T (Segundo Trimestre)"])
    m111_data.append(["", ""])
    m111_data.append(["RENDIMIENTOS DE ACTIVIDADES ECONÓMICAS (Profesionales)", ""])
    m111_data.append(["Concepto / Casilla", "Valor"])
    m111_data.append(["[Casilla 07] Número de Perceptores", num_perceptores])
    m111_data.append(["[Casilla 08] Base de las Retenciones (€)", round(total_bi_irpf, 2)])
    m111_data.append(["[Casilla 09] Importe de las Retenciones (15%) (€)", round(total_ret_irpf, 2)])
    m111_data.append(["", ""])
    m111_data.append(["LIQUIDACIÓN", ""])
    m111_data.append(["[Casilla 28] Suma de Retenciones (€)", round(total_ret_irpf, 2)])
    m111_data.append(["[Casilla 30] RESULTADO A INGRESAR (€)", round(total_ret_irpf, 2)])

    df_m111 = pd.DataFrame(m111_data)

    # Export to Excel
    output_path = "data/processed/Modelos_Trimestrales_2T_2026.xlsx"
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_m303.to_excel(writer, sheet_name='Modelo 303', index=False, header=False)
        df_m111.to_excel(writer, sheet_name='Modelo 111', index=False, header=False)

    print(f"Quarterly tax models successfully written to {output_path}")

if __name__ == "__main__":
    generate_quarterly_models()