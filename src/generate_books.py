import os
import pandas as pd
from clean_data import clean_and_process_data

def generate_registration_books():
    # Load clean data
    emitidas, recibidas = clean_and_process_data()

    # Reset index to avoid alignment issues when assigning to a new DataFrame
    emitidas = emitidas.reset_index(drop=True)
    recibidas = recibidas.reset_index(drop=True)

    # 1. Format Facturas Expedidas (Emitidas) official registration book
    lib_exp = pd.DataFrame()
    lib_exp['Reg. Número'] = range(1, len(emitidas) + 1)
    lib_exp['Fecha de Expedición'] = emitidas['FECHA'].dt.strftime('%d/%m/%Y')
    lib_exp['N.º Factura'] = emitidas['N.º ORDEN']
    lib_exp['Cliente / Razón Social'] = emitidas['CLIENTE']
    lib_exp['NIF / CIF Destinatario'] = emitidas['NIF']
    lib_exp['Base Imponible (€)'] = emitidas['BASE IMPONIBLE']
    lib_exp['Tipo de IVA (%)'] = (emitidas['IVA REPERCUTIDO'] * 100).astype(int)
    lib_exp['Cuota de IVA (€)'] = emitidas['RETENCIÓN APLICADA']
    lib_exp['Total Factura (€)'] = emitidas['TOTAL FACTURA']
    lib_exp['Estado de Cobro'] = emitidas['ESTADO'].fillna('PENDIENTE')

    # Add summary row at the end
    totals_exp = pd.Series({
        'Reg. Número': 'TOTALES',
        'Base Imponible (€)': lib_exp['Base Imponible (€)'].sum(),
        'Cuota de IVA (€)': lib_exp['Cuota de IVA (€)'].sum(),
        'Total Factura (€)': lib_exp['Total Factura (€)'].sum()
    })
    lib_exp = pd.concat([lib_exp, pd.DataFrame([totals_exp])], ignore_index=True)

    # 2. Format Facturas Recibidas official registration book
    lib_rec = pd.DataFrame()
    lib_rec['Reg. Número'] = range(1, len(recibidas) + 1)
    lib_rec['Fecha de Recepción'] = recibidas['FECHA'].dt.strftime('%d/%m/%Y')
    lib_rec['N.º Factura Original'] = recibidas['N.º ORDEN']
    lib_rec['Proveedor / Razón Social'] = recibidas['PROVEEDOR']
    lib_rec['NIF / CIF Expedidor'] = recibidas['NIF']

    lib_rec['Base Imponible (€)'] = recibidas['BASE IMPONIBLE'].fillna(0.0)
    lib_rec['Tipo de IVA (%)'] = (recibidas['IVA SPORTADO'].fillna(0.0) * 100).astype(int)
    lib_rec['Cuota de IVA (€)'] = recibidas['IMPORTE IVA'].fillna(0.0)
    lib_rec['Gasto Exento / Sin IVA (€)'] = recibidas['GASTOS SIN IVA'].fillna(0.0)

    # Withholding tax (15% applied on abogado fees, orders 57, 67, 85)
    withholding_list = []
    for idx, row in recibidas.iterrows():
        if "PELEGRÍN" in str(row['PROVEEDOR']).upper() or "52411407" in str(row['NIF']):
            withholding_list.append(round(row['BASE IMPONIBLE'] * 0.15, 2))
        else:
            withholding_list.append(0.0)
    lib_rec['Retención IRPF (€)'] = withholding_list

    lib_rec['Total Factura (€)'] = recibidas['TOTAL FACTURA']

    # Add summary row at the end
    totals_rec = pd.Series({
        'Reg. Número': 'TOTALES',
        'Base Imponible (€)': lib_rec['Base Imponible (€)'].sum(),
        'Cuota de IVA (€)': lib_rec['Cuota de IVA (€)'].sum(),
        'Gasto Exento / Sin IVA (€)': lib_rec['Gasto Exento / Sin IVA (€)'].sum(),
        'Retención IRPF (€)': lib_rec['Retención IRPF (€)'].sum(),
        'Total Factura (€)': lib_rec['Total Factura (€)'].sum()
    })
    lib_rec = pd.concat([lib_rec, pd.DataFrame([totals_rec])], ignore_index=True)

    # Write both to an Excel workbook with styling
    output_path = "data/processed/Libros_Registro_Oficiales_2T_2026.xlsx"
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        lib_exp.to_excel(writer, sheet_name='Facturas Expedidas', index=False)
        lib_rec.to_excel(writer, sheet_name='Facturas Recibidas', index=False)

    print(f"Official VAT Books successfully written to {output_path}")

if __name__ == "__main__":
    generate_registration_books()