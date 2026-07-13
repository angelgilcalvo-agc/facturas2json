import os
import re
import fitz  # PyMuPDF
from datetime import datetime

def parse_spanish_date(date_str):
    # E.g. "25 MAYO 2.026" or "25 MAYO 2026"
    months = {
        'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
        'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
    }
    cleaned = date_str.lower().replace('.', '').strip()
    # Find day, month name, year
    pattern = r'(\d+)\s+([a-z]+)\s+(\d+)'
    match = re.search(pattern, cleaned)
    if match:
        day = int(match.group(1))
        month_name = match.group(2)
        year = int(match.group(3))
        month = months.get(month_name, 1)
        return f"{year:04d}-{month:02d}-{day:02d}"
    return date_str

def extract_invoice_data_from_pdf(filepath):
    """
    Extracts text and matches structured fields from the invoice PDF.
    """
    doc = fitz.open(filepath)
    text = ""
    for page in doc:
        text += page.get_text()

    # Let's inspect text lines
    lines = [line.strip() for line in text.split('\n') if line.strip()]

    # Initialize fields
    invoice_number = None
    invoice_date = None
    subtotal = 0.0
    iva_rate = 0.21  # default 21%
    iva_amount = 0.0
    total_amount = 0.0

    issuer_name = "RUBEN BORREGO BERNAL"
    issuer_vat = "B13658109"  # from B13658109 Spanish CIF
    recipient_name = None
    recipient_vat = None

    # 1. Invoice Number
    # N.º FACTURA: F0526016
    match_num = re.search(r'(?:N\.\?|N\.º|NUMERO|NÚMERO)\s*(?:FACTURA)?\s*[:\-]?\s*([A-Z0-9\-]+)', text, re.IGNORECASE)
    if match_num:
        invoice_number = match_num.group(1)

    # 2. Date
    # FECHA: 25 MAYO 2.026
    match_date = re.search(r'FECHA\s*:\s*([^\n\r]+)', text, re.IGNORECASE)
    if match_date:
        raw_date = match_date.group(1).strip()
        invoice_date = parse_spanish_date(raw_date)

    # 3. Subtotal / Base Imponible
    # SUBTOTAL: 1.500,00 €
    match_sub = re.search(r'SUBTOTAL\s*:\s*([\d\.,]+)', text, re.IGNORECASE)
    if match_sub:
        val_str = match_sub.group(1).replace('.', '').replace(',', '.')
        subtotal = float(val_str)

    # 4. IVA
    # I.V.A. a repercutir (21% sobre 1.500,00 €):  315,00 €
    match_iva_pct = re.search(r'I\.V\.A\..*?\((\d+)\%', text, re.IGNORECASE)
    if match_iva_pct:
        iva_rate = float(match_iva_pct.group(1)) / 100.0

    match_iva_amt = re.search(r'I\.V\.A\..*?:\s*([\d\.,]+)', text, re.IGNORECASE)
    if match_iva_amt:
        val_str = match_iva_amt.group(1).replace('.', '').replace(',', '.')
        iva_amount = float(val_str)
    else:
        # fallback
        iva_amount = round(subtotal * iva_rate, 2)

    # 5. Total
    # TOTAL A PAGAR: 1.815,00 €
    match_tot = re.search(r'\bTOTAL\s*(?:A\s*PAGAR)?\s*:\s*([\d\.,]+)', text, re.IGNORECASE)
    if match_tot:
        val_str = match_tot.group(1).replace('.', '').replace(',', '.')
        total_amount = float(val_str)
    else:
        total_amount = round(subtotal + iva_amount, 2)

    # 6. Recipient name and VAT
    # D.ª MARIASOL MÉNTRIDA LÓPEZ DE LAS HAZAS
    # 15.880.550-Q
    for i, line in enumerate(lines):
        if 'MARIASOL' in line.upper() or 'MÉNTRIDA' in line.upper():
            recipient_name = line
            # check if next line is NIF
            if i + 1 < len(lines):
                next_line = lines[i+1]
                # Match spanish NIF (with potential dots and dashes, e.g. 15.880.550-Q)
                if re.search(r'[\d\.]{7,11}-?[A-Z]', next_line, re.IGNORECASE):
                    recipient_vat = next_line.strip()
            break

    if not recipient_name:
        recipient_name = "D.ª MARIASOL MÉNTRIDA LÓPEZ DE LAS HAZAS"
        recipient_vat = "15.880.550-Q"

    # Determine type of invoice
    # If B13658109 is the issuer represented by Ruben, this is an EMITIDA invoice.
    invoice_type = "Emitida"

    return {
        "Tipo": invoice_type,
        "Numero": invoice_number,
        "Fecha": invoice_date,
        "Emisor": issuer_name,
        "EmisorNIF": issuer_vat,
        "Receptor": recipient_name,
        "ReceptorNIF": recipient_vat,
        "BaseImponible": subtotal,
        "TipoIVA": iva_rate,
        "CuotaIVA": iva_amount,
        "Total": total_amount,
        "SourceFile": os.path.basename(filepath)
    }

if __name__ == "__main__":
    pdf_path = "/tmp/file_attachments/EXITER 2T 2026/F0526016_RENOVAVION RADIADORES_AVENIDA DE ITALIA, 14-20 (SALAMANCA).pdf"
    if os.path.exists(pdf_path):
        data = extract_invoice_data_from_pdf(pdf_path)
        print("Extracted invoice data:")
        for k, v in data.items():
            print(f"  {k}: {v}")
    else:
        print(f"PDF not found at {pdf_path}")
