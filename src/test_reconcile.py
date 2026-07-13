import os
import pandas as pd
import pytest
from src.reconcile_expenses import categorize_concepto, parse_card_movements, PDF_METADATA_MAP


def test_categorize_concepto():
    # Test non-deductible personal medical expense
    cat, ded, note = categorize_concepto("Clinica Hernade", -488.00)
    assert "No Deducible" in cat
    assert ded == "No"

    # Test cabify
    cat, ded, note = categorize_concepto("Cabify Es 2620ffqntwmz", -26.21)
    assert "Cabify" in note or "VTC" in cat
    assert ded == "Sí"

    # Test fuel
    cat, ded, note = categorize_concepto("Cepsa Carbajal", -50.00)
    assert "Combustible" in cat
    assert "50%" in ded

    # Test restaurant
    cat, ded, note = categorize_concepto("Restaurante Sal", -16.90)
    assert "Manutención" in cat
    assert ded == "Sí (Dieta)"


def test_pdf_metadata():
    # Verify that we have defined metadata for O2, Renfe, Cabify, Uber, Taffijn, etc.
    assert "O2 (1).pdf" in PDF_METADATA_MAP
    assert "Herramientas.pdf" in PDF_METADATA_MAP
    assert "cabify-sale-ESPV26ES107295.pdf" in PDF_METADATA_MAP

    # Verify values are reasonable
    for k, v in PDF_METADATA_MAP.items():
        assert "Fecha_Factura" in v
        assert "Numero_Factura" in v
        assert "Total" in v
        assert isinstance(v["Total"], float)


def test_parse_card_movements():
    # Test parsing card movements of Q2
    df = parse_card_movements()
    assert isinstance(df, pd.DataFrame)
    assert len(df) > 0

    # Check that all dates are indeed in Q2 2026 (April, May, June)
    assert df["Fecha_Clean"].min() >= pd.to_datetime("2026-04-01")
    assert df["Fecha_Clean"].max() <= pd.to_datetime("2026-06-30")

    # Check Clinica Hernade is classified correctly in card movements
    clinica_row = df[df["Concepto"].str.contains("Clinica")]
    assert len(clinica_row) == 1
    assert clinica_row.iloc[0]["Deducible_IRPF"] == "No"
