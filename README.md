# Contaser - Microservicio COMPRAS FC ELEC

Microservicio (FastAPI + openpyxl) para el flujo n8n `CONTASER - COMPRAS FC ELEC (Módulo 1)`.

Lee un archivo `.xlsm` preservando macros y todas las hojas, procesa la hoja
"COMPRAS FC ELEC", calcula totales/beneficio/alertas y devuelve el JSON de
resultados junto con el `.xlsm` modificado (en base64).

## Endpoints
- `GET /health` → estado del servicio.
- `POST /process` → `multipart/form-data` con campo `file` (el .xlsm).
  Header opcional `X-Service-Token` si se define la variable `SERVICE_TOKEN`.

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
