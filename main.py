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

app = FastAPI(title="Contaser - COMPRAS FC ELEC", version="1.3")

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


def _bump_collection(xml: str, tag: str, items_xml: str, n_items: int):
    """Suma n_items al count de <tag> e inserta items antes de </tag>.
    Devuelve (xml_modificado, índice_inicial_de_los_nuevos_items)."""
    m = re.search(r'<%s count="(\d+)"' % tag, xml)
    if not m:
        raise HTTPException(500, f"styles.xml sin colección <{tag}>")
    old = int(m.group(1))
    xml = xml[:m.start(1)] + str(old + n_items) + xml[m.end(1):]
    close = xml.find("</%s>" % tag)
    if close == -1:
        raise HTTPException(500, f"styles.xml sin cierre </{tag}>")
    xml = xml[:close] + items_xml + xml[close:]
    return xml, old


# Paleta corporativa de la tabla de resultados (gama verde)
C_TITULO_FONDO = "FF375623"   # verde oscuro
C_HEADER_FONDO = "FF70AD47"   # verde medio
C_LABEL_FONDO = "FFE2EFDA"    # verde claro
C_TEXTO_OSCURO = "FF375623"


def _augment_styles(styles_xml: str) -> tuple:
    """Agrega fuentes/rellenos/bordes/formatos y 5 estilos de celda nuevos.
    Devuelve (styles_xml_nuevo, dict_estilos)."""
    # Formato de porcentaje 0.0% (id personalizado libre)
    ids = [int(x) for x in re.findall(r'numFmtId="(\d+)"', styles_xml)]
    pct_id = max([163] + ids) + 1
    numfmt = '<numFmt numFmtId="%d" formatCode="0.0%%"/>' % pct_id
    m = re.search(r'<numFmts count="(\d+)">', styles_xml)
    if m:
        styles_xml = styles_xml[:m.start(1)] + str(int(m.group(1)) + 1) + styles_xml[m.end(1):]
        close = styles_xml.find("</numFmts>")
        styles_xml = styles_xml[:close] + numfmt + styles_xml[close:]
    else:
        m2 = re.search(r"<styleSheet[^>]*>", styles_xml)
        styles_xml = (styles_xml[:m2.end()]
                      + '<numFmts count="1">%s</numFmts>' % numfmt
                      + styles_xml[m2.end():])

    fonts = (
        '<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="%s"/><name val="Calibri"/></font>' % C_TEXTO_OSCURO
    )
    styles_xml, f0 = _bump_collection(styles_xml, "fonts", fonts, 3)

    fills = (
        '<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % C_TITULO_FONDO
        + '<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % C_HEADER_FONDO
        + '<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % C_LABEL_FONDO
    )
    styles_xml, l0 = _bump_collection(styles_xml, "fills", fills, 3)

    border = ('<border><left style="thin"><color indexed="64"/></left>'
              '<right style="thin"><color indexed="64"/></right>'
              '<top style="thin"><color indexed="64"/></top>'
              '<bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>')
    styles_xml, b0 = _bump_collection(styles_xml, "borders", border, 1)

    xfs = (
        # título (banda azul marino, texto blanco centrado)
        '<xf numFmtId="0" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' % (f0, l0, b0)
        # encabezados Concepto/Valor (azul medio, texto blanco)
        + '<xf numFmtId="0" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' % (f0 + 1, l0 + 1, b0)
        # etiquetas (azul claro, texto azul oscuro en negrilla)
        + '<xf numFmtId="0" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>' % (f0 + 2, l0 + 2, b0)
        # valores numéricos con miles (#,##0 = formato incorporado 3)
        + '<xf numFmtId="3" fontId="0" fillId="0" borderId="%d" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>' % b0
        # porcentaje 0.0%%
        + '<xf numFmtId="%d" fontId="0" fillId="0" borderId="%d" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>' % (pct_id, b0)
    )
    styles_xml, x0 = _bump_collection(styles_xml, "cellXfs", xfs, 5)

    estilos = {"titulo": x0, "header": x0 + 1, "label": x0 + 2,
               "money": x0 + 3, "pct": x0 + 4}
    return styles_xml, estilos


def write_results_table(content: bytes, sheet_title: str, last_row: int,
                        tabla: list) -> bytes:
    """Escribe la tabla de resultados en columnas B y C con título,
    encabezados, colores y formatos numéricos. tabla: [(label, value, kind)]
    con kind en {money, pct}."""
    zin = zipfile.ZipFile(io.BytesIO(content))
    sheet_path = _sheet_xml_path(zin, sheet_title)
    sxml = zin.read(sheet_path).decode("utf-8")
    styles_xml, st = _augment_styles(zin.read("xl/styles.xml").decode("utf-8"))

    col_l, col_v = "C", "D"

    existing = [int(x) for x in re.findall(r'<row r="(\d+)"', sxml)]
    max_existing = max(existing) if existing else last_row
    start = max(last_row + 3, max_existing + 2)

    def txt(col, r, s, text):
        return ('<c r="%s%d" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                % (col, r, s, escape(str(text))))

    rows_xml = []
    # Fila título (banda en B y C)
    rows_xml.append('<row r="%d" ht="20" customHeight="1">%s%s</row>' % (
        start,
        txt(col_l, start, st["titulo"], "RESULTADOS COMPRAS FC ELEC"),
        txt(col_v, start, st["titulo"], ""),
    ))
    # Fila encabezados
    rows_xml.append("<row r=\"%d\">%s%s</row>" % (
        start + 1,
        txt(col_l, start + 1, st["header"], "Concepto"),
        txt(col_v, start + 1, st["header"], "Valor"),
    ))
    # Filas de datos
    for i, (label, value, kind) in enumerate(tabla):
        r = start + 2 + i
        s_val = st["pct"] if kind == "pct" else st["money"]
        val = '<c r="%s%d" s="%d"><v>%s</v></c>' % (col_v, r, s_val, value)
        rows_xml.append('<row r="%d">%s%s</row>' % (
            r, txt(col_l, r, st["label"], label), val))

    idx = sxml.rfind("</sheetData>")
    if idx == -1:
        raise HTTPException(500, "XML de hoja sin </sheetData>; estructura inesperada")
    new_sxml = sxml[:idx] + "".join(rows_xml) + sxml[idx:]

    # Reempaquetar: todo idéntico excepto la hoja y styles.xml
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = new_sxml.encode("utf-8")
            elif item.filename == "xl/styles.xml":
                data = styles_xml.encode("utf-8")
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
        ("Facturado - Notas Crédito", money(base_fact_nc), "money"),
        ("Valor Susceptible Beneficio (Electrónicos)", money(benef_elec), "money"),
        ("60% (Facturado - NC)", money(0.6 * base_fact_nc), "money"),
        ("40% (Facturado - NC)", money(0.4 * base_fact_nc), "money"),
        # como fracción: Excel lo muestra 91,8% gracias al formato 0.0%
        ("% Beneficio / (Facturado - NC)",
         round(benef_elec / base_fact_nc, 4) if base_fact_nc else 0, "pct"),
    ]

    # ---- Escritura quirúrgica sobre los bytes ORIGINALES ----
    nuevo = write_results_table(content, ws.title, last_row, tabla)
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
