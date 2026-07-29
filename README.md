# Contaser - Microservicio COMPRAS FC ELEC

Microservicio (FastAPI + openpyxl) para el flujo n8n `CONTASER - COMPRAS FC ELEC (Módulo 1)`.

Lee un archivo `.xlsm` preservando macros y todas las hojas, procesa la hoja
"COMPRAS FC ELEC", calcula totales/beneficio/alertas y devuelve el JSON de
resultados junto con el `.xlsm` modificado (en base64).

## Endpoints
- `GET /health` → estado del servicio.
- `POST /process` → `multipart/form-data` con campo `file` (el .xlsm).
  Header opcional `X-Service-Token` si se define la variable `SERVICE_TOKEN`.
- `POST /audit` → **solo lectura**: audita el .xlsm sin modificarlo ni devolverlo.
  Mismo formato de entrada y mismo token que `/process`.
- `POST /inspect` → **solo lectura**: estructura de una hoja cualquiera.
  Campos: `file`, `hoja` (opcional), `filas` (opcional). Sin `hoja` lista todas.
- `POST /dian-preview` → **solo lectura**: resultado del Módulo 2 sin escribir.

### `/audit` — para qué sirve
Responde con datos (no con teoría) dos preguntas abiertas del proyecto:

1. **¿La tabla de resultados puede contaminar fórmulas existentes?**
   `write_results_table` escribe en las columnas **C y D**. Si alguna fórmula
   del libro suma un rango que las incluya (p. ej. `SUMA('COMPRAS FC ELEC'!D:D)`),
   los valores de la tabla se sumarían al total del cliente **en silencio**.
   El endpoint barre las fórmulas de todas las hojas + los nombres definidos y
   reporta cuáles tocan C o D.
2. **¿Por qué sale vacío el nombre del cliente?**
   `header_dump` devuelve las celdas reales de las filas previas al encabezado,
   así se ve la etiqueta exacta que `parse_header` no está macheando.

Campos clave de la respuesta:

| Campo | Qué dice |
|---|---|
| `resumen.veredicto` | `RIESGO DETECTADO` / `SIN RIESGO DETECTADO` |
| `formulas_que_tocan_C_o_D` | Hoja, celda y fórmula de cada caso peligroso |
| `columnas_datos` | Qué letra de columna ocupa cada campo de datos |
| `columnas_datos_en_C_o_D` | Si el área de datos invade la zona de escritura |
| `header_dump` / `header_parseado` | Diagnóstico del nombre vacío |
| `zona_escritura` | Fila exacta donde caería la tabla |
| `tabla_previa_detectada` | Si ya hay una tabla de una corrida anterior |
| `marcador_en_sharedStrings` | Si Excel movió el marcador a `sharedStrings.xml`, la idempotencia dejaría de detectarlo y se escribiría una **tabla duplicada** |

Limitaciones conocidas: la detección de rangos es heurística (regex sobre el XML),
no evalúa fórmulas indirectas (`INDIRECTO`, `DESREF`) ni referencias generadas por
macros VBA. Las listas se truncan a 200 elementos (`resumen.listas_truncadas`).

## Cambios de la v1.9 — Módulo 2: comparación REPORTE DIAN

`/process` ahora corre **dos módulos en una sola pasada** y escribe el archivo
**una sola vez**. El flujo de n8n no cambia.

```
descarga → /process ─┬─ Módulo 1: COMPRAS FC ELEC   (siempre)
                     └─ Módulo 2: REPORTE DIAN      (si hay 2+ hojas)
                     → un archivo de salida → una escritura en Drive
```

**Cómo identifica las hojas.** Toma todas las hojas cuyo nombre empiece por
`REPORTE DIAN` y que tengan estructura de reporte, y las ordena por la **fecha
de reporte interna** (fila 2 de cada hoja), no por el nombre:

- **base** = la fecha más antigua → es la que Contaser ya trabajó
- **nueva** = la fecha más reciente → el export recién llegado

Así funciona con `REPORTE DIAN ULTIMO`, `REPORTE DIAN 19 AGOST`, `SEPT 3` o
cualquier convención que use el equipo.

**Por qué no se comparan los textos.** La hoja base está **anotada por
Contaser**: trae los cinco `Tope N`, textos de `Uso declaración` más largos y
dos columnas extra del consultante que la hoja nueva no tiene. Comparar textos
marcaría esas anotaciones como diferencias. El emparejamiento va en tres pasadas:

| Pasada | Criterio | Resultado |
|---|---|---|
| 1 | `NIT + Valor` exacto | sin cambio |
| 2 | `NIT + prefijo del Detalle` normalizado | **valor cambiado** |
| 3 | lo que sobra | solo en la nueva → **nuevo**; solo en la base → **desaparecido** |

**Mapeo de columnas: gana la primera aparición.** La hoja base repite `NIT` y
`Nombre / Razón Social` (una vez para el tercero que reporta y otra para el
consultante). Sin esta regla se compararía el NIT del cliente contra sí mismo.
Las filas sin NIT se descartan: ahí caen los `Tope N`.

**El cuadro que escribe**, al final de la hoja nueva, en las columnas B a F:

1. Resumen — conteos y valores de cada categoría, más totales de ambas hojas
2. **Impacto en la declaración** — agrupa lo nuevo y lo cambiado por código de
   renglón (`R29`, `R74`, `R99`…) y por `Tope N`, extraídos de la columna
   `Uso declaración Sugerida`. Renglones y topes se listan **separados**: un
   registro puede afectar a ambos y sumarlos daría un total engañoso
3. Detalle de los registros nuevos, los cambiados (con su valor anterior) y los
   desaparecidos

**Idempotente**, con la misma detección de `sharedStrings` de la v1.7. ⚠️ Al
reescribir se borra **todo lo que haya desde la fila del marcador hacia abajo**;
el cuadro va siempre al final, así que no escribas notas manuales debajo.

**Nuevo `POST /dian-preview`** (solo lectura): corre el Módulo 2 y devuelve el
resultado **sin escribir nada**. Sirve para validar la comparación antes de
dejarla tocar un archivo.

**Limitación conocida:** si un mismo tercero reporta dos conceptos con el
**mismo valor exacto**, el emparejamiento es ambiguo. Esos casos se emparejan
igual pero se listan en `ambiguos` para revisión manual, en vez de adivinar en
silencio.

## Cambios de la v1.8

Blindaje ante columnas corridas y un endpoint de inspección.

1. **Barrido más amplio.** Los nombres de las columnas que entrega la DIAN no
   cambian, pero su **posición sí** puede correrse si se agregan columnas.
   `find_header_row` y `map_columns` ahora recorren hasta la fila 80 y la
   columna 40 (antes 60 y 24). Con el CUFE en la columna L, el margen pasa de
   12 a 28 columnas nuevas.
2. **Error de columnas faltantes con diagnóstico.** El 400 ahora incluye los
   encabezados **reales** que encontró en la fila, con su letra de columna. Antes
   solo decía qué faltaba, sin pistas de qué había en su lugar.
3. **Columnas opcionales sin `KeyError`.** Solo 6 de las 11 columnas son
   obligatorias, pero el código leía las 11. Si faltaba `Valor Notas Crédito`
   o `num_factura_venta`, reventaba con un 500 sin explicación. Ahora el helper
   `_celda` devuelve `None` y los valores caen a 0 o cadena vacía.
4. **Nuevo `POST /inspect`** (solo lectura). Devuelve la estructura de cualquier
   hoja: encabezados reales, filas de muestra y dimensiones. Sin `hoja`, lista
   todas las hojas del libro. Pensado para diseñar los módulos siguientes
   (REPORTE DIAN, PATRIMONIO, CED.1 GENERAL) sobre datos reales.

   Campos del formulario: `file` (el .xlsm), `hoja` (opcional), `filas`
   (opcional, por defecto 25, máximo 200).

## Cambios de la v1.7

Tres correcciones sobre problemas confirmados con archivos reales:

1. **Nombre de cliente vacío.** `parse_header` comparaba las etiquetas por
   igualdad exacta contra `"nombre"`. El export de la DIAN cambia de formato
   entre versiones: RAMIREZ trae `"Nombre"` pero CUADROS, PAREDES y ALVARO
   traen `"Nombre o razón social"` — **3 de cada 4 archivos fallaban**. Ahora
   se normaliza sin tildes ni dos puntos y se compara **por prefijo**. Si aun
   así queda vacío, cae al nombre del archivo sin extensión ni sufijo ` - AAAA`.
2. **Tabla duplicada.** La idempotencia buscaba el texto del marcador dentro
   del XML de la hoja. Si alguien abre el archivo en Excel y lo guarda, Excel
   mueve ese texto a `xl/sharedStrings.xml` y la celda pasa a `<c t="s">`; el
   marcador ya no se encontraba y se escribía una **segunda** tabla. Ahora
   también se resuelven los índices de `sharedStrings`.
3. **Validación de salida** (`validar_salida`). Antes de devolver el archivo se
   comprueba que no falte ninguna entrada del ZIP, que `xl/vbaProject.bin` sea
   **idéntico byte a byte**, que los CRC sean válidos, que abra con openpyxl y
   que conserve el mismo número de hojas. Si algo falla se lanza HTTP 500 y la
   ejecución de n8n se detiene **antes** del nodo que sobrescribe el original.

## Variables de entorno
- `SERVICE_TOKEN` (opcional): si se define, las peticiones a `/process` deben
  enviar el header `X-Service-Token` con ese valor.

## Despliegue en EasyPanel
1. Subir estos archivos a un repositorio GitHub (raíz).
2. EasyPanel → mismo proyecto de n8n → **+ Service → App**.
3. Nombre: `contaser-xlsx`. Source: GitHub → repo + rama `main`.
4. Build type: Dockerfile (autodetectado).
5. Environment: `SERVICE_TOKEN` = <secreto>.
6. Puerto interno: 8000. No requiere dominio público.
7. Deploy. Verificar logs: "Uvicorn running on http://0.0.0.0:8000".

Desde n8n se invoca por red interna: `http://contaser-xlsx:8000/process`.
