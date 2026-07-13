import json
import os
import datetime

tax_data = {
    "Modelo_303": {
        "Periodo": "2T 2026",
        "NIF_Declarante": "34975084L",
        "Nombre_Declarante": "ANA BELEN FEIJOO VIDUEIRA",
        "IVA_Devengado": {
            "Regimen_General_21": {
                "Base": 40044.27,
                "Tipo": 0.21,
                "Cuota": 8415.46,
                "Casillas": ["01", "02", "03"]
            },
            "Regimen_General_10": {
                "Base": 4036.81,
                "Tipo": 0.10,
                "Cuota": 403.46,
                "Casillas": ["04", "05", "06"]
            },
            "Adquisiciones_Intracomunitarias_Bienes": {
                "Base": 851.43,
                "Tipo": 0.21,
                "Cuota": 178.80,
                "Casillas": ["10", "11"]
            },
            "Otras_Inversiones_Sujeto_Pasivo_Servicios": {
                "Base": 69.82,
                "Tipo": 0.21,
                "Cuota": 14.66,
                "Casillas": ["12", "13"]
            },
            "Total_Devengado": {
                "Cuota": 9012.38,
                "Casilla": "27"
            }
        },
        "IVA_Deducible": {
            "Operaciones_Interiores_Corrientes": {
                "Base": 30618.68,
                "Cuota": 6355.39,
                "Casillas": ["28", "29"]
            },
            "Adquisiciones_Intracomunitarias_Bienes": {
                "Base": 851.43,
                "Cuota": 178.80,
                "Casillas": ["36", "37"]
            },
            "Adquisiciones_Intracomunitarias_Servicios": {
                "Base": 69.82,
                "Cuota": 14.66,
                "Casillas": ["40", "41"]
            },
            "Total_Deducible": {
                "Cuota": 6548.85,
                "Casilla": "46"
            }
        },
        "Resultado_Fin": {
            "Diferencia": 2463.53,
            "Casilla_71": 2463.53,
            "Concepto": "A ingresar (A pagar)"
        }
    },
    "Modelo_115": {
        "Periodo": "2T 2026",
        "NIF_Declarante": "34975084L",
        "Nombre_Declarante": "ANA BELEN FEIJOO VIDUEIRA",
        "Retenciones_Alquiler": {
            "Numero_Perceptores": 1,
            "Base_Retenciones": 280.00,
            "Porcentaje": 0.19,
            "Retenciones_Ingresadas": 53.20,
            "Casillas": {
                "Perceptores": "01",
                "Base_Imponible": "02",
                "Retenciones": "03"
            }
        },
        "Resultado_Fin": {
            "Resultado": 53.20,
            "Concepto": "A ingresar (A pagar)"
        }
    },
    "Modelo_111": {
        "Periodo": "2T 2026",
        "NIF_Declarante": "34975084L",
        "Nombre_Declarante": "ANA BELEN FEIJOO VIDUEIRA",
        "Retenciones_Trabajadores_Profesionales": {
            "Rendimientos_Trabajo": {
                "Perceptores": 0,
                "Base": 0.00,
                "Retencion": 0.00
            },
            "Rendimientos_Actividades_Profesionales": {
                "Perceptores": 0,
                "Base": 0.00,
                "Retencion": 0.00
            }
        },
        "Resultado_Fin": {
            "Resultado": 0.00,
            "Concepto": "Sin actividad/importes"
        }
    }
}

os.makedirs("data/processed", exist_ok=True)
with open("data/processed/Modelos_Fiscales_2T_2026.json", "w", encoding="utf-8") as f:
    json.dump(tax_data, f, indent=4, ensure_ascii=False)

report = f"""# INFORME DE MODELOS FISCALES - 2T 2026
**Cliente:** ANA BELEN FEIJOO VIDUEIRA
**NIF:** 34975084L
**Periodo:** Segundo Trimestre (2T) de 2026
**Fecha del Informe:** {datetime.date.today().strftime('%d/%m/%Y')}

---

## 1. MODELO 303 - IMPUESTO SOBRE EL VALOR AÑADIDO (IVA)

### IVA DEVENGADO (Ventas e Inversión de Sujeto Pasivo)
* **Régimen General (21%) [Ventas España 21%]:**
  - **Base Imponible (Casilla 01):** {tax_data['Modelo_303']['IVA_Devengado']['Regimen_General_21']['Base']:,.2f} €
  - **Tipo aplicable (Casilla 02):** {tax_data['Modelo_303']['IVA_Devengado']['Regimen_General_21']['Tipo']*100:.1f}%
  - **Cuota Devengada (Casilla 03):** {tax_data['Modelo_303']['IVA_Devengado']['Regimen_General_21']['Cuota']:,.2f} €
* **Régimen General (10%) [Ventas España 10%]:**
  - **Base Imponible (Casilla 04):** {tax_data['Modelo_303']['IVA_Devengado']['Regimen_General_10']['Base']:,.2f} €
  - **Tipo aplicable (Casilla 05):** {tax_data['Modelo_303']['IVA_Devengado']['Regimen_General_10']['Tipo']*100:.1f}%
  - **Cuota Devengada (Casilla 06):** {tax_data['Modelo_303']['IVA_Devengado']['Regimen_General_10']['Cuota']:,.2f} €
* **Adquisiciones Intracomunitarias de Bienes (Auto-repercutido):**
  - **Base Imponible (Casilla 10):** {tax_data['Modelo_303']['IVA_Devengado']['Adquisiciones_Intracomunitarias_Bienes']['Base']:,.2f} €
  - **Cuota Devengada (Casilla 11):** {tax_data['Modelo_303']['IVA_Devengado']['Adquisiciones_Intracomunitarias_Bienes']['Cuota']:,.2f} €
* **Otras Inversiones del Sujeto Pasivo (Servicios UE auto-repercutido):**
  - **Base Imponible (Casilla 12):** {tax_data['Modelo_303']['IVA_Devengado']['Otras_Inversiones_Sujeto_Pasivo_Servicios']['Base']:,.2f} €
  - **Cuota Devengada (Casilla 13):** {tax_data['Modelo_303']['IVA_Devengado']['Otras_Inversiones_Sujeto_Pasivo_Servicios']['Cuota']:,.2f} €

* **TOTAL IVA DEVENGADO (Casilla 27):** **{tax_data['Modelo_303']['IVA_Devengado']['Total_Devengado']['Cuota']:,.2f} €**

### IVA DEDUCIBLE (Gastos y Compras)
* **Operaciones Interiores Corrientes (Bienes y Servicios España):**
  - **Base Imponible (Casilla 28):** {tax_data['Modelo_303']['IVA_Deducible']['Operaciones_Interiores_Corrientes']['Base']:,.2f} €
  - **Cuota Deducible (Casilla 29):** {tax_data['Modelo_303']['IVA_Deducible']['Operaciones_Interiores_Corrientes']['Cuota']:,.2f} €
* **Adquisiciones Intracomunitarias de Bienes (Simétrico):**
  - **Base Imponible (Casilla 36):** {tax_data['Modelo_303']['IVA_Deducible']['Adquisiciones_Intracomunitarias_Bienes']['Base']:,.2f} €
  - **Cuota Deducible (Casilla 37):** {tax_data['Modelo_303']['IVA_Deducible']['Adquisiciones_Intracomunitarias_Bienes']['Cuota']:,.2f} €
* **Adquisiciones Intracomunitarias de Servicios (Simétrico):**
  - **Base Imponible (Casilla 40):** {tax_data['Modelo_303']['IVA_Deducible']['Adquisiciones_Intracomunitarias_Servicios']['Base']:,.2f} €
  - **Cuota Deducible (Casilla 41):** {tax_data['Modelo_303']['IVA_Deducible']['Adquisiciones_Intracomunitarias_Servicios']['Cuota']:,.2f} €

* **TOTAL IVA DEDUCIBLE (Casilla 46):** **{tax_data['Modelo_303']['IVA_Deducible']['Total_Deducible']['Cuota']:,.2f} €**

### RESULTADO DEL PERIODO (Modelo 303)
* **Resultado Neto (Diferencia Casilla 71):** **{tax_data['Modelo_303']['Resultado_Fin']['Diferencia']:,.2f} €**
* **Estado:** **{tax_data['Modelo_303']['Resultado_Fin']['Concepto']}**

---

## 2. MODELO 115 - RETENCIONES DE ARRENDAMIENTOS (Alquiler de Local)
* **Número de perceptores (Casilla 01):** {tax_data['Modelo_115']['Retenciones_Alquiler']['Numero_Perceptores']}
* **Base de las retenciones (Casilla 02):** {tax_data['Modelo_115']['Retenciones_Alquiler']['Base_Retenciones']:,.2f} €
* **Retenciones e ingresos a cuenta (Casilla 03):** **{tax_data['Modelo_115']['Retenciones_Alquiler']['Retenciones_Ingresadas']:,.2f} €**
* **Resultado Neto (Casilla 05):** **{tax_data['Modelo_115']['Resultado_Fin']['Resultado']:,.2f} €**
* **Estado:** **{tax_data['Modelo_115']['Resultado_Fin']['Concepto']}**
*(Arrendador: Rosa Fernández Da Bouza, NIF: 34934884T)*

---

## 3. MODELO 111 - RETENCIONES DE TRABAJO Y ACTIVIDADES PROFESIONALES
* **Rendimientos del trabajo / profesionales:**
  - **Número de perceptores:** 0
  - **Base de retenciones:** 0,00 €
  - **Retenciones practicadas:** 0,00 €
* **Resultado Neto (Casilla 28/30):** **0,00 €**
* **Estado:** **Sin importes / Sin actividad sujeta en este trimestre**

---

## NOTA IMPORTANTE SOBRE OPERACIONES FUERA DE MODELO 303 (Régimen de Ventas en la UE - Ventanilla Única OSS)
Las ventas a particulares en otros países de la Unión Europea (B2C) se declaran trimestralmente bajo el régimen de la **Ventanilla Única (OSS - One Stop Shop)** en lugar del Modelo 303.
* **Base imponible total facturada en OSS (Resto de Europa):** 3,923.63 €
* **Cuota de IVA total recaudada para otros estados miembros:** 793.03 €
*(Se detallará el desglose por países en los libros registro oficiales)*.
"""

with open("data/processed/Informe_Modelos_Fiscales_2T_2026.md", "w", encoding="utf-8") as f:
    f.write(report)

print("Tax report generated.")
