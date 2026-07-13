import os
import pandas as pd
import pytest
from clean_data import clean_and_process_data

def test_clean_and_process_data():
    emitidas, recibidas = clean_and_process_data()

    # Verify counts
    assert len(emitidas) == 11
    assert len(recibidas) == 35

    # Verify exact math for Emitidas
    assert emitidas['BASE IMPONIBLE'].sum() == pytest.approx(31902.12)
    assert emitidas['TOTAL FACTURA'].sum() == pytest.approx(38421.06)

    # Verify exact math for Recibidas
    assert recibidas['BASE IMPONIBLE'].sum() == pytest.approx(21375.82)
    assert recibidas['IMPORTE IVA'].sum() == pytest.approx(4485.41)
    assert recibidas['GASTOS SIN IVA'].sum() == pytest.approx(470.45)
    assert recibidas['TOTAL FACTURA'].sum() == pytest.approx(25791.06)

def test_generated_excel_files():
    # Verify files exist
    assert os.path.exists("data/processed/Libros_Registro_Oficiales_2T_2026.xlsx")
    assert os.path.exists("data/processed/Modelos_Trimestrales_2T_2026.xlsx")
    assert os.path.exists("data/processed/Preparacion_Cierre_Anual_2026.xlsx")

    # Read and inspect sheets
    df_exp = pd.read_excel("data/processed/Libros_Registro_Oficiales_2T_2026.xlsx", sheet_name="Facturas Expedidas")
    df_rec = pd.read_excel("data/processed/Libros_Registro_Oficiales_2T_2026.xlsx", sheet_name="Facturas Recibidas")

    # Verify total rows match
    assert df_exp.iloc[-1]['Reg. Número'] == 'TOTALES'
    assert df_exp.iloc[-1]['Base Imponible (€)'] == pytest.approx(31902.12)
    assert df_exp.iloc[-1]['Total Factura (€)'] == pytest.approx(38421.06)

    assert df_rec.iloc[-1]['Reg. Número'] == 'TOTALES'
    assert df_rec.iloc[-1]['Base Imponible (€)'] == pytest.approx(21375.82)
    assert df_rec.iloc[-1]['Cuota de IVA (€)'] == pytest.approx(4485.41)
    assert df_rec.iloc[-1]['Gasto Exento / Sin IVA (€)'] == pytest.approx(470.45)
    assert df_rec.iloc[-1]['Total Factura (€)'] == pytest.approx(25791.06)
