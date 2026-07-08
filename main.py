# -*- coding: utf-8 -*-
"""
Microservicio Contaser - Módulo 1: COMPRAS FC ELEC
Lee un .xlsm (preservando macros y todas las hojas), extrae y calcula la hoja
"COMPRAS FC ELEC", escribe los totales y devuelve resultados + archivo modificado.
"""
import base64
import io
import os
from datetime import datetime, date

from fastapi import FastAPI, UploadFile, File, Header, HTTPException
import openpyxl

app = FastAPI(title="Contaser - COMPRAS FC ELEC", version="1.0")

SERVICE_TOKEN = os.getenv("SERVICE_TOKEN", "")  # si está vacío, no exige token
SHEET_NAME = "COMPRAS FC ELEC"
END_MARKER = "facturas procesadas disponibles"

# Título de encabezado de datos -> clave interna
COL_LABELS = {
    "identificación emisor factura": "nit_emisor",
    "nombre emisor factura": "nombre_emisor",
    "fecha emisión": "fecha_emision",
    "valor facturado": "valor_facturado",
    "valor notas crédito": "notas_credito",
    "valor notas débito": "notas_debito",
    "valor factura / afectada con notas débito - crédito": "valor_neto",
    "valor susceptible beneficio": "valor_beneficio",
    "medios de pago": "medio_pago",
    "num_factura_venta": "num_factura",
}


def _norm(s):
    return str(s).strip().lower() if s is not None else ""


def to_number(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_iso(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip() if v is not None else ""


def classify_medio(v):
    s = _norm(v)
    if s.startswith("efectivo y"):
        return "efectivo_otros"
    if "lectr" in s:
        return "electronico"
    if "fectivo" in s:
        return "efectivo"
    return "otro"


def find_sheet(wb):
    for ws in wb.worksheets:
        if _norm(ws.title) == _norm(SHEET_NAME):
            return ws
    for ws in wb.worksheets:
        if _norm(SHEET_NAME) in _norm(ws.title):
            return ws
    raise HTTPException(400, f"No se encontró la hoja '{SHEET_NAME}'")


def find_header_row(ws):
    for r in range(1, 60):
        for c in range(1, 20):
            if _norm(ws.cell(row=r, column=c).value) == "identificación emisor factura":
                return r
    raise HTTPException(400, "No se encontró la fila de encabezados de datos (COMPRAS FC ELEC)")


def map_columns(ws, header_row):
    colmap = {}
    for c in range(1, 25):
        label = _norm(ws.cell(row=header_row, column=c).value)
        if label.startswith("cufe"):
            colmap["cufe"] = c
        elif label in COL_LABELS:
            colmap[COL_LABELS[label]] = c
    requeridas = ["nit_emisor", "valor_facturado", "valor_neto", "valor_beneficio", "medio_pago", "cufe"]
    faltan = [k for k in requeridas if k not in colmap]
    if faltan:
        raise HTTPException(400, f"Faltan columnas en el encabezado: {faltan}")
    return colmap


def parse_header(ws, header_row):
    anio = nit = nombre = None
    nit_row = None
    for r in range(1, header_row):
        label = valor = None
        for c in range(1, 6):
            cell = ws.cell(row=r, column=c).value
            if cell is not None and str(cell).strip() != "":
                if label is None:
                    label = _norm(cell)
                else:
                    valor = cell
                    break
        if label == "año gravable":
            anio = valor
        elif label == "nit":
            nit = valor
            nit_row = r
        elif label == "nombre":
            if nit_row is not None and r > nit_row:
                nombre = valor
            elif nombre is None and valor is not None and not str(valor).lower().startswith("informe"):
                nombre = valor
    return anio, nit, nombre


@app.get("/health")
def health():
    return {"status": "ok", "service": "contaser-xlsx", "version": "1.0"}


@app.post("/process")
async def process(file: UploadFile = File(...), x_service_token: str = Header(default="")):
    if SERVICE_TOKEN and x_service_token != SERVICE_TOKEN:
        raise HTTPException(401, "Token de servicio inválido")

    content = await file.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), keep_vba=True, data_only=False)
    except Exception as e:
        raise HTTPException(400, f"No se pudo abrir el archivo .xlsm: {e}")

    ws = find_sheet(wb)
    header_row = find_header_row(ws)
    colmap = map_columns(ws, header_row)
    anio, nit, nombre = parse_header(ws, header_row)

    facturas = []
    last_row = header_row
    r = header_row + 1
    while r < header_row + 100000:
        nit_em = ws.cell(row=r, column=colmap["nit_emisor"]).value
        if END_MARKER in _norm(nit_em):
            break
        claves = [ws.cell(row=r, column=colmap[k]).value for k in ("nit_emisor", "valor_facturado", "cufe")]
        if all(v is None or str(v).strip() == "" for v in claves):
            break
        facturas.append({
            "nit_emisor": str(nit_em).strip() if nit_em is not None else "",
            "nombre_emisor": str(ws.cell(row=r, column=colmap["nombre_emisor"]).value or "").strip(),
            "fecha_emision": to_iso(ws.cell(row=r, column=colmap["fecha_emision"]).value),
            "valor_facturado": to_number(ws.cell(row=r, column=colmap["valor_facturado"]).value),
            "notas_credito": to_number(ws.cell(row=r, column=colmap["notas_credito"]).value),
            "notas_debito": to_number(ws.cell(row=r, column=colmap["notas_debito"]).value),
            "valor_neto": to_number(ws.cell(row=r, column=colmap["valor_neto"]).value),
            "valor_beneficio": to_number(ws.cell(row=r, column=colmap["valor_beneficio"]).value),
            "medio_pago": str(ws.cell(row=r, column=colmap["medio_pago"]).value or "").strip(),
            "num_factura": str(ws.cell(row=r, column=colmap["num_factura"]).value or "").strip(),
            "cufe": str(ws.cell(row=r, column=colmap["cufe"]).value or "").strip(),
        })
        last_row = r
        r += 1

    if not facturas:
        raise HTTPException(400, "No se encontraron facturas en la hoja COMPRAS FC ELEC")

    # ---- Totales ----
    t_fact = sum(f["valor_facturado"] for f in facturas)
    t_nc = sum(f["notas_credito"] for f in facturas)
    t_nd = sum(f["notas_debito"] for f in facturas)
    t_neto = sum(f["valor_neto"] for f in facturas)
    t_benef = sum(f["valor_beneficio"] for f in facturas)
    benef_1pct = round(t_benef * 0.01)

    # ---- Conteos por medio ----
    medios = {"electronico": {"count": 0, "valor": 0.0},
              "efectivo": {"count": 0, "valor": 0.0},
              "efectivo_otros": {"count": 0, "valor": 0.0},
              "otro": {"count": 0, "valor": 0.0}}
    for f in facturas:
        m = classify_medio(f["medio_pago"])
        medios[m]["count"] += 1
        medios[m]["valor"] += f["valor_facturado"]

    pct_elec = round((medios["electronico"]["valor"] / t_fact * 100), 1) if t_fact else 0
    pct_efec = round(100 - pct_elec, 1)

    # ---- Alertas ----
    vistos = {}
    duplicados = []
    for f in facturas:
        k = (f["nit_emisor"], f["num_factura"])
        if k in vistos:
            duplicados.append(f"{f['nit_emisor']} - {f['num_factura']}")
        vistos[k] = True
    alertas = {
        "facturas_en_cero": [f for f in facturas if f["valor_facturado"] == 0],
        "duplicadas": duplicados,
        "mayores_10m": [f for f in facturas if f["valor_facturado"] > 10_000_000],
        "efectivo_mayor_5m": [f for f in facturas
                              if classify_medio(f["medio_pago"]) in ("efectivo", "efectivo_otros")
                              and f["valor_facturado"] > 5_000_000],
        "desviacion_60_40": abs(pct_elec - 60) > 15,
    }

    # ---- Beneficio susceptible SOLO de pagos electrónicos ----
    benef_elec = sum(f["valor_beneficio"] for f in facturas
                     if classify_medio(f["medio_pago"]) == "electronico")
    base_fact_nc = t_fact - t_nc  # Facturado - Notas Crédito (base de porcentajes)
    pct_benef = round(benef_elec / base_fact_nc * 100, 1) if base_fact_nc else 0.0

    # ---- Escribir tabla de resultados al final del Excel ----
    # El marcador DIAN ("...Facturas procesadas disponibles") queda en last_row+1;
    # la tabla inicia en last_row+3 dejando una fila en blanco.
    col_label = colmap["nit_emisor"]
    col_valor = colmap["valor_facturado"]
    start = last_row + 3
    tabla = [
        ("Facturado - Notas Crédito", base_fact_nc, "#,##0"),
        ("Valor Susceptible Beneficio (Electrónicos)", benef_elec, "#,##0"),
        ("60% (Facturado - NC)", round(0.6 * base_fact_nc), "#,##0"),
        ("40% (Facturado - NC)", round(0.4 * base_fact_nc), "#,##0"),
        ("% Beneficio / (Facturado - NC)",
         (benef_elec / base_fact_nc) if base_fact_nc else 0, "0.0%"),
    ]
    for i, (label, value, fmt) in enumerate(tabla):
        ws.cell(row=start + i, column=col_label).value = label
        cell = ws.cell(row=start + i, column=col_valor)
        cell.value = value
        cell.number_format = fmt

    out = io.BytesIO()
    wb.save(out)
    archivo_b64 = base64.b64encode(out.getvalue()).decode()

    return {
        "cliente": {"nit": str(nit).strip() if nit else "", "nombre": str(nombre).strip() if nombre else "",
                    "anio_gravable": int(to_number(anio)) if anio else None},
        "fecha_procesamiento": datetime.now().isoformat(timespec="seconds"),
        "total_facturas": len(facturas),
        "totales": {"valor_facturado": t_fact, "notas_credito": t_nc, "notas_debito": t_nd,
                    "valor_neto": t_neto, "valor_beneficio": t_benef},
        "beneficio_1pct": benef_1pct,
        "tabla_final": {
            "facturado_menos_nc": base_fact_nc,
            "beneficio_electronicos": benef_elec,
            "estimado_60": round(0.6 * base_fact_nc),
            "estimado_40": round(0.4 * base_fact_nc),
            "pct_beneficio_sobre_base": pct_benef,
        },
        "medios": medios,
        "pct_real": {"electronico": pct_elec, "efectivo": pct_efec},
        "alertas": alertas,
        "facturas": facturas,
        "archivo_modificado_b64": archivo_b64,
    }
