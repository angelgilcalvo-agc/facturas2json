import openpyxl
import datetime

wb = openpyxl.load_workbook("data/interim/ANA_FEIJOO_CALCULO_2T_2026.xlsx")
sheet = wb["Facturas emitidas"]

# Clear the placeholder row at row 4
for col in range(1, 10):
    sheet.cell(row=4, column=col, value=None)

sales_records = [
    # ABRIL 2026
    {
        "Fecha": datetime.date(2026, 4, 30),
        "Factura": "RESUMEN_AMAZON_ES_202604_21",
        "Destinatario": "Ventas Península 21% - Clientes Varios (España)",
        "Base": 13795.05,
        "IVA": 2899.11,
        "Total": 16694.16,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado)"
    },
    {
        "Fecha": datetime.date(2026, 4, 30),
        "Factura": "RESUMEN_AMAZON_ES_202604_00",
        "Destinatario": "Ventas Canarias 0% (Exentas) - Clientes Varios (España)",
        "Base": 53.23,
        "IVA": 0.00,
        "Total": 53.23,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado)"
    },
    {
        "Fecha": datetime.date(2026, 4, 30),
        "Factura": "RESUMEN_AMAZON_PT_202604_23",
        "Destinatario": "Ventas Particulares 23% - Clientes Varios (Portugal)",
        "Base": 444.09,
        "IVA": 102.19,
        "Total": 546.28,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 4, 30),
        "Factura": "RESUMEN_AMAZON_PT_202604_00",
        "Destinatario": "Ventas Empresas Intracomunitarias - Clientes Varios (Portugal)",
        "Base": 37.73,
        "IVA": 0.00,
        "Total": 37.73,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - exento ROI)"
    },
    {
        "Fecha": datetime.date(2026, 4, 30),
        "Factura": "RESUMEN_AMAZON_FR_202604_20",
        "Destinatario": "Ventas Particulares 20% - Clientes Varios (Francia)",
        "Base": 473.31,
        "IVA": 94.66,
        "Total": 567.97,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 4, 30),
        "Factura": "RESUMEN_AMAZON_FR_202604_00",
        "Destinatario": "Ventas Empresas Intracomunitarias - Clientes Varios (Francia)",
        "Base": 0.43,
        "IVA": 0.00,
        "Total": 0.43,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - exento ROI)"
    },
    {
        "Fecha": datetime.date(2026, 4, 30),
        "Factura": "RESUMEN_AMAZON_IT_202604_22",
        "Destinatario": "Ventas Particulares 22% - Clientes Varios (Italia)",
        "Base": 69.24,
        "IVA": 15.23,
        "Total": 84.47,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 4, 30),
        "Factura": "RESUMEN_AMAZON_LU_202604_17",
        "Destinatario": "Ventas Particulares 17% - Clientes Varios (Luxemburgo)",
        "Base": 20.84,
        "IVA": 3.54,
        "Total": 24.38,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },

    # MAYO 2026
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_ES_202605_21",
        "Destinatario": "Ventas Península 21% - Clientes Varios (España)",
        "Base": 14512.34,
        "IVA": 3050.26,
        "Total": 17562.60,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_ES_202605_10",
        "Destinatario": "Ventas Península IVA Reducido 10% - Clientes Varios (España)",
        "Base": 1557.65,
        "IVA": 155.72,
        "Total": 1713.37,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_ES_202605_00",
        "Destinatario": "Ventas Canarias 0% (Exentas) - Clientes Varios (España)",
        "Base": 9.11,
        "IVA": 0.00,
        "Total": 9.11,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_PT_202605_23",
        "Destinatario": "Ventas Particulares 23% - Clientes Varios (Portugal)",
        "Base": 487.43,
        "IVA": 112.17,
        "Total": 599.60,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_PT_202605_06",
        "Destinatario": "Ventas Particulares IVA Reducido 6% - Clientes Varios (Portugal)",
        "Base": 168.67,
        "IVA": 10.11,
        "Total": 178.78,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_PT_202605_00",
        "Destinatario": "Ventas Empresas Intracomunitarias - Clientes Varios (Portugal)",
        "Base": 95.63,
        "IVA": 0.00,
        "Total": 95.63,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - exento ROI)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_FR_202605_20",
        "Destinatario": "Ventas Particulares 20% - Clientes Varios (Francia)",
        "Base": 809.86,
        "IVA": 162.20,
        "Total": 972.06,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_FR_202605_00",
        "Destinatario": "Ventas Empresas Intracomunitarias - Clientes Varios (Francia)",
        "Base": 57.44,
        "IVA": 0.00,
        "Total": 57.44,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - exento ROI)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_IT_202605_22",
        "Destinatario": "Ventas Particulares 22% - Clientes Varios (Italia)",
        "Base": 115.76,
        "IVA": 25.36,
        "Total": 141.12,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_NL_202605_21",
        "Destinatario": "Ventas Particulares 21% - Clientes Varios (Países Bajos)",
        "Base": 11.47,
        "IVA": 2.41,
        "Total": 13.88,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 5, 31),
        "Factura": "RESUMEN_AMAZON_BE_202605_21",
        "Destinatario": "Ventas Particulares 21% - Clientes Varios (Bélgica)",
        "Base": 8.83,
        "IVA": 1.86,
        "Total": 10.69,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },

    # JUNIO 2026
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_ES_202606_21",
        "Destinatario": "Ventas Península 21% - Clientes Varios (España)",
        "Base": 11736.88,
        "IVA": 2466.09,
        "Total": 14202.97,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_ES_202606_10",
        "Destinatario": "Ventas Península IVA Reducido 10% - Clientes Varios (España)",
        "Base": 2479.16,
        "IVA": 247.74,
        "Total": 2726.90,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_ES_202606_00",
        "Destinatario": "Ventas Canarias 0% (Exentas) - Clientes Varios (España)",
        "Base": 57.77,
        "IVA": 0.00,
        "Total": 57.77,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_PT_202606_23",
        "Destinatario": "Ventas Particulares 23% - Clientes Varios (Portugal)",
        "Base": 218.74,
        "IVA": 50.33,
        "Total": 269.07,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_PT_202606_06",
        "Destinatario": "Ventas Particulares IVA Reducido 6% - Clientes Varios (Portugal)",
        "Base": 276.51,
        "IVA": 16.56,
        "Total": 293.07,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_PT_202606_00",
        "Destinatario": "Ventas Empresas Intracomunitarias - Clientes Varios (Portugal)",
        "Base": 24.90,
        "IVA": 0.00,
        "Total": 24.90,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - exento ROI)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_FR_202606_20",
        "Destinatario": "Ventas Particulares 20% - Clientes Varios (Francia)",
        "Base": 763.96,
        "IVA": 153.01,
        "Total": 916.97,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_FR_202606_00",
        "Destinatario": "Ventas Empresas Intracomunitarias - Clientes Varios (Francia)",
        "Base": 42.50,
        "IVA": 0.00,
        "Total": 42.50,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - exento ROI)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_IT_202606_22",
        "Destinatario": "Ventas Particulares 22% - Clientes Varios (Italia)",
        "Base": 160.11,
        "IVA": 34.89,
        "Total": 195.00,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_IT_202606_00",
        "Destinatario": "Ventas Empresas Intracomunitarias - Clientes Varios (Italia)",
        "Base": 53.77,
        "IVA": 0.00,
        "Total": 53.77,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - exento ROI)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_DK_202606_19",
        "Destinatario": "Ventas Particulares 19% - Clientes Varios (Dinamarca)",
        "Base": 10.95,
        "IVA": 2.08,
        "Total": 13.03,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_NL_202606_19",
        "Destinatario": "Ventas Particulares 19% - Clientes Varios (Países Bajos)",
        "Base": 10.95,
        "IVA": 2.08,
        "Total": 13.03,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_LU_202606_17",
        "Destinatario": "Ventas Particulares 17% - Clientes Varios (Luxemburgo)",
        "Base": 8.79,
        "IVA": 1.50,
        "Total": 10.29,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    },
    {
        "Fecha": datetime.date(2026, 6, 30),
        "Factura": "RESUMEN_AMAZON_BE_202606_21",
        "Destinatario": "Ventas Particulares 21% - Clientes Varios (Bélgica)",
        "Base": 13.59,
        "IVA": 2.85,
        "Total": 16.44,
        "Retencion": 0.0,
        "Estado": "Revisado (Agregado - OSS)"
    }
]

row_idx = 4
for r in sales_records:
    sheet.cell(row=row_idx, column=1, value=r["Fecha"])
    sheet.cell(row=row_idx, column=2, value=r["Factura"])
    sheet.cell(row=row_idx, column=3, value=r["Destinatario"])
    sheet.cell(row=row_idx, column=4, value=r["Base"])
    sheet.cell(row=row_idx, column=5, value=r["IVA"])
    sheet.cell(row=row_idx, column=6, value=r["Total"])
    sheet.cell(row=row_idx, column=7, value=r["Retencion"])
    sheet.cell(row=row_idx, column=8, value=r["Estado"])
    row_idx += 1

wb.save("data/interim/ANA_FEIJOO_CALCULO_2T_2026.xlsx")
print("Appended sales.")
