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
