import os
import openpyxl
import pandas as pd
import pytest

def test_ledger_exists():
    path = 'data/processed/Libro_Registro_Facturas_Recibidas_2T_2026.xlsx'
    assert os.path.exists(path), f"File {path} does not exist"

def test_ledger_structure():
    path = 'data/processed/Libro_Registro_Facturas_Recibidas_2T_2026.xlsx'
    wb = openpyxl.load_workbook(path, data_only=False)
    assert "Facturas recibidas" in wb.sheetnames

    ws = wb["Facturas recibidas"]
    assert ws.max_row == 105
    assert ws.max_column == 24

    # Test formula in totals row
    assert ws.cell(105, 11).value == "=SUM(K2:K104)"
    assert ws.cell(105, 13).value == "=SUM(M2:M104)"
    assert ws.cell(105, 14).value == "=SUM(N2:N104)"

def test_ledger_data_types():
    path = 'data/processed/Libro_Registro_Facturas_Recibidas_2T_2026.xlsx'
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb["Facturas recibidas"]

    # Check that NIF in row 2 is written as text format
    nif_cell = ws.cell(2, 5) # Nº factura/documento
    assert nif_cell.number_format == '@'
