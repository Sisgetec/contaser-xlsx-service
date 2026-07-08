# -*- coding: utf-8 -*-
"""
Microservicio Contaser - Módulo 1: COMPRAS FC ELEC (v1.1)
Lee un .xlsm, calcula la hoja "COMPRAS FC ELEC" y escribe la tabla de
resultados mediante EDICIÓN QUIRÚRGICA del XML interno: solo se modifica
el XML de esa hoja; el resto del archivo (58 hojas, macros, dibujos,
vínculos externos) se copia byte por byte, evitando corrupción.
"""
import base64
import io
import os
import re
import zipfile
from datetime import datetime, date
from xml.sax.saxutils import escape, unescape

from fastapi import FastAPI, UploadFile, File, Header, HTTPException
import openpyxl
from openpyxl.utils import get_column_letter

app = FastAPI(title="Contaser - COMPRAS FC ELEC", version="1.1")

SERVICE_TOKEN = os.getenv("SERVICE_TOKEN", "")  # si está vacío, no exige token
SHEET_NAME = "COMPRAS FC ELEC"
END_MARKER = "facturas procesadas disponibles"

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


# ---------------------------------------------------------------------------
# Escritura quirúrgica del XML de la hoja (sin reescribir el resto del libro)
# ---------------------------------------------------------------------------

def _sheet_xml_path(z: zipfile.ZipFile, sheet_title: str) -> str:
    wb_xml = z.read("xl/workbook.xml").decode("utf-8")
    rid = None
    for sm in re.finditer(r"<sheet\b[^>]*>", wb_xml):
        tag = sm.group(0)
        nm = re.search(r'name="([^"]*)"', tag)
        if nm and unescape(nm.group(1)) == sheet_title:
            rm = re.search(r'r:id="([^"]*)"', tag)
            if rm:
                rid = rm.group(1)
            break
    if not rid:
        raise HTTPException(500, f"No se encontró la hoja '{sheet_title}' en workbook.xml")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rel = re.search(r'<Relationship\b[^>]*\bId="%s"[^>]*/?>' % re.escape(rid), rels)
    if not rel:
        raise HTTPException(500, f"No se encontró la relación {rid} de la hoja")
    tm = re.search(r'Target="([^"]*)"', rel.group(0))
    target = tm.group(1)
    return target[1:] if target.startswith("/") else "xl/" + target


def _cell_style(sheet_xml: str, ref_row: int, col_letter: str) -> str:
    """Devuelve el atributo s=".." de la celda de referencia, o cadena vacía."""
    rm = re.search(r'<row r="%d"\b.*?(?:</row>|/>)' % ref_row, sheet_xml, re.DOTALL)
    if not rm:
        return ""
    cm = re.search(r'<c r="%s%d"([^>]*)>' % (col_letter, ref_row), rm.group(0))
    if not cm:
        return ""
    sm = re.search(r'\bs="(\d+)"', cm.group(1))
    return ' s="%s"' % sm.group(1) if sm else ""


def write_results_table(content: bytes, sheet_title: str, last_row: int,
                        col_label_idx: int, col_valor_idx: int,
                        tabla: list) -> bytes:
    """tabla: lista de (etiqueta, valor) donde valor int -> numérico, str -> texto."""
    zin = zipfile.ZipFile(io.BytesIO(content))
    sheet_path = _sheet_xml_path(zin, sheet_title)
    sxml = zin.read(sheet_path).decode("utf-8")

    col_l = get_column_letter(col_label_idx)
    col_v = get_column_letter(col_valor_idx)

    # Inicio de la tabla: después del marcador DIAN y de cualquier fila ya
    # existente en el XML (garantiza orden ascendente de filas = sin merge).
    existing = [int(x) for x in re.findall(r'<row r="(\d+)"', sxml)]
    max_existing = max(existing) if existing else last_row
    start = max(last_row + 3, max_existing + 2)

    # Reutilizar estilos de una factura real para que el formato coincida
    style_label = _cell_style(sxml, last_row, col_l)
    style_valor = _cell_style(sxml, last_row, col_v)

    rows_xml = []
    for i, (label, value) in enumerate(tabla):
        r = start + i
        lbl = ('<c r="%s%d"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
               % (col_l, r, style_label, escape(str(label))))
        if isinstance(value, (int, float)):
            val = '<c r="%s%d"%s><v>%s</v></c>' % (col_v, r, style_valor, value)
        else:
            val = ('<c r="%s%d"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                   % (col_v, r, style_valor, escape(str(value))))
        cells = (lbl + val) if col_label_idx < col_valor_idx else (val + lbl)
        rows_xml.append('<row r="%d">%s</row>' % (r, cells))

    idx = sxml.rfind("</sheetData>")
    if idx == -1:
        raise HTTPException(500, "XML de hoja sin </sheetData>; estructura inesperada")
    new_sxml = sxml[:idx] + "".join(rows_xml) + sxml[idx:]

    # Reempaquetar: todo idéntico excepto el XML de esta hoja
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = new_sxml.encode("utf-8")
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zout.writestr(zi, data)
    zin.close()
    return out.getvalue()


@app.get("/health")
def health():
    return {"status": "ok", "service": "contaser-xlsx", "version": "1.1"}


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

    # ---- Tabla final: base = Facturado - Notas Crédito ----
    benef_elec = sum(f["valor_beneficio"] for f in facturas
                     if classify_medio(f["medio_pago"]) == "electronico")
    base_fact_nc = t_fact - t_nc
    pct_benef = round(benef_elec / base_fact_nc * 100, 1) if base_fact_nc else 0.0

    def money(v):
        return int(round(v))

    tabla = [
        ("Facturado - Notas Crédito", money(base_fact_nc)),
        ("Valor Susceptible Beneficio (Electrónicos)", money(benef_elec)),
        ("60% (Facturado - NC)", money(0.6 * base_fact_nc)),
        ("40% (Facturado - NC)", money(0.4 * base_fact_nc)),
        ("% Beneficio / (Facturado - NC)", ("%.1f%%" % pct_benef).replace(".", ",")),
    ]

    # ---- Escritura quirúrgica sobre los bytes ORIGINALES ----
    nuevo = write_results_table(
        content, ws.title, last_row,
        colmap["nit_emisor"], colmap["valor_facturado"], tabla,
    )
    archivo_b64 = base64.b64encode(nuevo).decode()

    return {
        "cliente": {"nit": str(nit).strip() if nit else "", "nombre": str(nombre).strip() if nombre else "",
                    "anio_gravable": int(to_number(anio)) if anio else None},
        "fecha_procesamiento": datetime.now().isoformat(timespec="seconds"),
        "total_facturas": len(facturas),
        "totales": {"valor_facturado": t_fact, "notas_credito": t_nc, "notas_debito": t_nd,
                    "valor_neto": t_neto, "valor_beneficio": t_benef},
        "beneficio_1pct": benef_1pct,
        "tabla_final": {
            "facturado_menos_nc": money(base_fact_nc),
            "beneficio_electronicos": money(benef_elec),
            "estimado_60": money(0.6 * base_fact_nc),
            "estimado_40": money(0.4 * base_fact_nc),
            "pct_beneficio_sobre_base": pct_benef,
        },
        "medios": medios,
        "pct_real": {"electronico": pct_elec, "efectivo": pct_efec},
        "alertas": alertas,
        "facturas": facturas,
        "archivo_modificado_b64": archivo_b64,
    }
