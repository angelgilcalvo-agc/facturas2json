import os
import openpyxl
from datetime import datetime
from generate_excel import create_tax_ledger_2026

def test_excel_generation():
    # Execute the generation function
    create_tax_ledger_2026()

    # Assert output file exists
    filepath = "data/processed/libro_registro_impuestos_2026.xlsx"
    assert os.path.exists(filepath), "Excel ledger file was not created"

    # Load workbook and check sheet names
    wb = openpyxl.load_workbook(filepath, data_only=False)
    expected_sheets = ["Facturas Emitidas", "Facturas Recibidas", "Resumen de Impuestos"]
    for sheet in expected_sheets:
        assert sheet in wb.sheetnames, f"Missing expected sheet: {sheet}"

    # Check Facturas Emitidas content and formulas
    ws_emitidas = wb["Facturas Emitidas"]
    # Header check
    assert ws_emitidas["A4"].value == "Trimestre"
    assert ws_emitidas["E4"].value == "Cliente / Razón Social"
    # Row 5 data check
    assert ws_emitidas["A5"].value == "2T"
    assert ws_emitidas["B5"].value == "F0526016"
    assert ws_emitidas["F5"].value == 1500.00
    assert ws_emitidas["G5"].value == 0.21
    assert ws_emitidas["H5"].value == 315.00
    assert ws_emitidas["I5"].value == 1815.00
    # Row 101 total formula check
    assert ws_emitidas["E101"].value == "TOTAL"
    assert ws_emitidas["F101"].value == "=SUM(F5:F100)"
    assert ws_emitidas["H101"].value == "=SUM(H5:H100)"
    assert ws_emitidas["I101"].value == "=SUM(I5:I100)"

    # Check Facturas Recibidas content
    ws_recibidas = wb["Facturas Recibidas"]
    assert ws_recibidas["A4"].value == "Trimestre"
    assert ws_recibidas["E101"].value == "TOTAL"
    assert ws_recibidas["F101"].value == "=SUM(F5:F100)"

    # Check Resumen de Impuestos formulas
    ws_resumen = wb["Resumen de Impuestos"]
    # Table 1: IVA Repercutido
    assert ws_resumen["A5"].value == "Periodo"
    assert ws_resumen["A6"].value == "1T"
    assert ws_resumen["B6"].value == "=SUMIFS('Facturas Emitidas'!F:F, 'Facturas Emitidas'!A:A, \"1T\")"
    assert ws_resumen["C6"].value == "=SUMIFS('Facturas Emitidas'!H:H, 'Facturas Emitidas'!A:A, \"1T\")"
    assert ws_resumen["A7"].value == "2T"
    assert ws_resumen["B7"].value == "=SUMIFS('Facturas Emitidas'!F:F, 'Facturas Emitidas'!A:A, \"2T\")"
    assert ws_resumen["C7"].value == "=SUMIFS('Facturas Emitidas'!H:H, 'Facturas Emitidas'!A:A, \"2T\")"

    # Table 3: Liquidacion
    assert ws_resumen["A22"].value == "1T"
    assert ws_resumen["B22"].value == "=C6"
    assert ws_resumen["C22"].value == "=C14"
    assert ws_resumen["D22"].value == "=B22-C22"

    assert ws_resumen["A23"].value == "2T"
    assert ws_resumen["B23"].value == "=C7"
    assert ws_resumen["C23"].value == "=C15"
    assert ws_resumen["D23"].value == "=B23-C23"

    # TOTAL ANUAL row 26
    assert ws_resumen["A26"].value == "TOTAL ANUAL"
    assert ws_resumen["B26"].value == "=C10"
    assert ws_resumen["C26"].value == "=C18"
    assert ws_resumen["D26"].value == "=B26-C26"

if __name__ == "__main__":
    test_excel_generation()
    print("All tests passed!")
