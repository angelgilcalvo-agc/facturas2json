import json
import os
import pandas as pd

def test_exported_ledger_recibidas():
    assert os.path.exists("data/processed/Libro_Registro_Facturas_Recibidas_2T_2026.xlsx")
    df = pd.read_excel("data/processed/Libro_Registro_Facturas_Recibidas_2T_2026.xlsx")
    assert len(df) == 117
    assert round(df["Base Imponible (EUR)"].sum(), 2) == 34192.44
    assert round(df["Cuota Soportada (EUR)"].sum(), 2) == 6409.86
    assert round(df["Retención IRPF (EUR)"].sum(), 2) == 53.20

def test_exported_ledger_emitidas():
    assert os.path.exists("data/processed/Libro_Registro_Facturas_Emitidas_2T_2026.xlsx")
    df = pd.read_excel("data/processed/Libro_Registro_Facturas_Emitidas_2T_2026.xlsx")
    assert len(df) == 33
    assert round(df["Base Imponible (EUR)"].sum(), 2) == 48586.69
    assert round(df["Cuota Repercutida (EUR)"].sum(), 2) == 9611.95

def test_tax_json_structure():
    assert os.path.exists("data/processed/Modelos_Fiscales_2T_2026.json")
    with open("data/processed/Modelos_Fiscales_2T_2026.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    assert "Modelo_303" in data
    assert "Modelo_115" in data
    assert "Modelo_111" in data

    assert data["Modelo_303"]["Resultado_Fin"]["Diferencia"] == 2463.53
    assert data["Modelo_115"]["Retenciones_Alquiler"]["Retenciones_Ingresadas"] == 53.20
    assert data["Modelo_111"]["Resultado_Fin"]["Resultado"] == 0.00
