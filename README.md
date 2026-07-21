# POC Generador de Reportes XOC

Este directorio contiene un POC para generar reportes DOCX a partir de texto del analista e imágenes/capturas usando Azure AI Foundry multimodal.

El objetivo del POC es validar el flujo técnico antes de integrarlo al API principal del backend.

Importante:

- La IA no genera el archivo Word.
- Azure solo analiza texto/imágenes y devuelve JSON.
- Python toma ese JSON y genera el DOCX usando plantillas.
- Este POC no modifica el backend productivo `XOC-API-AWS`.
- No se deben subir credenciales, API keys, `.env`, imágenes sensibles ni reportes generados.

## Estado actual del flujo

El flujo base ya está validado:

1. El usuario ingresa texto del analista.
2. El usuario adjunta una o varias imágenes.
3. Opcionalmente agrega una descripción breve por imagen.
4. Azure AI Foundry multimodal analiza la evidencia.
5. Azure devuelve JSON estructurado.
6. Python genera un DOCX usando la plantilla correspondiente.
7. El DOCX se guarda en `output/`.

Tipos de reporte soportados actualmente:

- `small_report`
- `informe_soporte`

Nota de negocio: estos reportes son internos. Para entrega al cliente se planteó un formato futuro llamado `minority_report`, que todavía no está implementado en este POC.

## Estructura principal

```text
prueba-script-xoc/
├── azure_foundry_multimodal_client.py
├── generate_report.py
├── generate_small_report.py
├── generate_support_report.py
├── test_foundry_multimodal_connection.py
├── requirements.txt
├── README.md
├── plantillas/
│   ├── Plantilla Small Report.docx
│   └── Plantilla Informe Soporte.docx
├── data/
└── output/
```

Además existe un mini front de prueba fuera de esta carpeta:

```text
../mini-front/
├── server.py
├── index.html
├── app.js
└── styles.css
```

El mini front sirve solo para pruebas locales. La integración productiva debería hacerse desde el frontend/backend oficial.

## Responsabilidad de cada archivo

### `azure_foundry_multimodal_client.py`

Cliente encargado de hablar con Azure AI Foundry / Azure OpenAI Responses API.

Responsabilidades:

- Leer imágenes locales.
- Convertir imágenes a `data:image/...;base64`.
- Enviar texto + imágenes al modelo multimodal.
- Solicitar que Azure devuelva solo JSON.
- Validar el JSON recibido.
- Reparar errores comunes de JSON cuando Azure responde con formato casi válido.
- Guardar el payload final en `output/generated_report_payload.json`.
- Usar fallback local si Azure falla y el flujo lo permite.

### `generate_report.py`

Motor principal de generación DOCX.

Responsabilidades:

- Cargar las plantillas DOCX.
- Reemplazar placeholders de portada.
- Mantener el diseño base de la portada.
- Generar el cuerpo del reporte dinámicamente.
- Aplicar diseño visual:
  - ficha inicial del caso;
  - resumen ejecutivo;
  - títulos con color corporativo;
  - separadores;
  - bullets;
  - cajas de limitaciones;
  - captions de imágenes.
- Validar que el DOCX final sea válido.

### `generate_small_report.py`

Wrapper CLI para generar un `small_report` desde consola.

### `generate_support_report.py`

Wrapper CLI para generar un `informe_soporte` desde consola.

### `test_foundry_multimodal_connection.py`

Script mínimo para validar conexión con Azure y verificar que el JSON devuelto tenga el esquema esperado.

### `plantillas/`

Contiene las plantillas Word base.

Actualmente las plantillas son principalmente portadas con placeholders tipo `Change` o `Change Tenant`.

El contenido del cuerpo no viene fijo desde la plantilla. Se genera dinámicamente según el JSON que devuelve Azure.

## Requisitos

Python recomendado:

```text
Python 3.11+
```

Instalar dependencias:

```powershell
cd "C:\Users\saalc\OneDrive\Escritorio\XOC\backend y pruebas\prueba-script-xoc"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Dependencias principales:

- `openai`
- `azure-identity`
- `azure-ai-projects`
- `azure-ai-agents`
- `python-docx`
- `python-dotenv`
- `lxml`
- `Pillow`

## Variables de entorno

Crear un archivo `.env` dentro de `prueba-script-xoc`.

No subir `.env` al repositorio.

Ejemplo seguro, sin credenciales reales:

```env
USE_AZURE_FOUNDRY=true
USE_AZURE_FOUNDRY_AGENT=false

AZURE_FOUNDRY_PROJECT_ENDPOINT=<endpoint-del-proyecto-foundry>
AZURE_FOUNDRY_OPENAI_ENDPOINT=<endpoint-compatible-openai-v1>
AZURE_FOUNDRY_MODEL_DEPLOYMENT=gpt-5-mini

AZURE_FOUNDRY_AGENT_NAME=<nombre-del-agente-si-se-usa-agent>
AZURE_FOUNDRY_AGENT_VERSION=<version-opcional>

AZURE_FOUNDRY_API_KEY=<no-commitear>
AZURE_OPENAI_API_KEY=<no-commitear>

REPORT_TYPE=small_report
REPORT_MAX_IMAGE_MB=10
AZURE_FOUNDRY_MAX_OUTPUT_TOKENS=6000
AZURE_FOUNDRY_JSON_SCHEMA=true
```

Notas:

- Si se usa API key, debe venir desde secrets/variables de entorno.
- Si no se usa API key, el cliente puede intentar usar `DefaultAzureCredential`.
- Para producción, usar el sistema de secretos del entorno, no archivos `.env` versionados.

## Contrato de Azure

Azure debe devolver solo JSON válido. No debe devolver Markdown, bloques de código ni texto adicional.

Esquema esperado:

```json
{
  "title": "",
  "summary": "",
  "sections": [
    {
      "title": "",
      "content": "",
      "bullets": [],
      "subsections": [
        {
          "title": "",
          "content": "",
          "bullets": []
        }
      ]
    }
  ],
  "limitations": [],
  "image_citations": [
    {
      "label": "Figura 1",
      "description": "",
      "used_in_sections": []
    }
  ]
}
```

### Regla importante sobre secciones

El reporte no debe forzar secciones.

Azure solo debe devolver secciones sustentadas por la evidencia recibida.

Por ejemplo, si el usuario solo reporta un problema, pero no indica solución ni acciones ejecutadas, Azure no debería devolver secciones como:

- `Acciones Realizadas`
- `Resultados`
- `Recomendaciones`

Si algo no se puede confirmar, debe ir en `limitations`.

## Prompt base esperado

El cliente en `azure_foundry_multimodal_client.py` usa un prompt con estas reglas:

- Actuar como analista senior XOC.
- Usar solo la evidencia entregada.
- No inventar datos, fechas, IPs, activos, severidades, causas, acciones ni resultados.
- Citar imágenes como `Figura 1`, `Figura 2`, etc.
- No agregar secciones vacías.
- No redundar.
- Devolver solo JSON válido.

## Manejo de imágenes

El flujo soporta una o varias imágenes.

Formatos soportados:

- PNG
- JPG/JPEG
- WEBP

Cada imagen se etiqueta como:

```text
Figura 1
Figura 2
Figura 3
```

Desde el mini front se puede agregar una descripción breve por imagen. Esa descripción se usa como caption:

```text
Figura 1. Captura del cliente mostrando login fallido.
```

Las imágenes se insertan cerca del párrafo donde Azure cite su figura. Si no se encuentra una referencia clara, se insertan en una sección fallback de evidencia.

## Cómo probar desde consola

### Probar conexión con Azure

```powershell
cd "C:\Users\saalc\OneDrive\Escritorio\XOC\backend y pruebas\prueba-script-xoc"
python test_foundry_multimodal_connection.py --report-type small_report
```

Con imagen:

```powershell
python test_foundry_multimodal_connection.py --report-type small_report --image ".\ruta\a\captura.png"
```

### Generar Small Report

```powershell
python generate_small_report.py `
  --text "El cliente reporta error de login. No se confirma resolución." `
  --image ".\ruta\a\captura-login.png"
```

### Generar Informe de Soporte

```powershell
python generate_support_report.py `
  --text "Se solicita revisión de conectividad. No se adjuntan logs." `
  --image ".\ruta\a\captura-dashboard.png"
```

### Modo local sin Azure

Sirve solo para validar generación del DOCX sin llamar al modelo.

```powershell
python generate_small_report.py --text "Prueba local" --no-azure
python generate_support_report.py --text "Prueba local" --no-azure
```

## Cómo probar con el mini front

El mini front está fuera de esta carpeta, a la misma altura de `prueba-script-xoc`.

Desde la raíz del workspace:

```powershell
cd "C:\Users\saalc\OneDrive\Escritorio\XOC\backend y pruebas"
python mini-front/server.py
```

Abrir en navegador:

```text
http://127.0.0.1:8765
```

En el mini front se puede:

- Elegir tipo de reporte.
- Ingresar nombre del cliente o compañía.
- Ingresar texto del analista.
- Adjuntar una o varias imágenes.
- Agregar descripción por imagen.
- Usar Azure real.
- Permitir fallback local si se desea.

Si se hacen cambios al código, reiniciar el servidor:

```powershell
Ctrl + C
python mini-front/server.py
```

Si el navegador mantiene caché, refrescar con:

```text
Ctrl + F5
```

## Salidas generadas

Los reportes se generan en:

```text
prueba-script-xoc/output/
```

Ejemplo de nombre:

```text
SMALL-REPORT_SR-POC-YYYYMMDD-HHMMSS_CLIENTE_YYYY-MM-DD.docx
INFORME-SOPORTE_IF-POC-YYYYMMDD-HHMMSS_CLIENTE_YYYY-MM-DD.docx
```

También se guarda el último JSON generado en:

```text
prueba-script-xoc/output/generated_report_payload.json
```

No se recomienda versionar `output/`, ya que puede contener información sensible.

## Consideraciones de seguridad

No subir:

- `.env`
- API keys
- capturas sensibles
- documentos generados
- payloads con información de clientes
- logs con contenido de imágenes/base64

El POC evita imprimir imágenes o base64 en consola.

Para producción se recomienda:

- Validar usuario/tenant antes de generar reportes.
- Verificar permisos sobre el caso.
- Limitar tamaño y cantidad de imágenes.
- Limpiar archivos temporales.
- Evitar guardar evidencia sensible más tiempo del necesario.
- Enviar logs solo técnicos, sin contenido sensible.
- Usar secretos del entorno/cloud y no archivos `.env`.

## Consideraciones para migrar al backend principal

Este POC no debería copiarse tal cual a producción. La recomendación es integrarlo como módulo o servicio dentro del backend principal.

### Flujo sugerido para integración

Endpoint sugerido:

```text
POST /reports/generate
```

Payload sugerido:

```json
{
  "report_type": "small_report",
  "client_name": "Cliente Demo",
  "analyst_text": "Texto ingresado por el analista",
  "images": [
    {
      "file": "<archivo multipart o referencia interna>",
      "description": "Captura del dashboard con alerta"
    }
  ]
}
```

Respuesta sugerida:

```json
{
  "ok": true,
  "report_id": "uuid",
  "file_name": "SMALL-REPORT_....docx",
  "download_url": "/reports/{id}/download"
}
```

### Módulos reutilizables

Se pueden reutilizar principalmente:

- `azure_foundry_multimodal_client.py`
- funciones de generación dentro de `generate_report.py`

El mini front no debería migrarse a producción. Solo sirve como herramienta local de prueba.

### Pendientes antes de producción

- Definir endpoint final.
- Definir autenticación/autorización por usuario y tenant.
- Definir almacenamiento final del DOCX:
  - disco temporal;
  - S3;
  - Azure Blob;
  - storage interno.
- Definir limpieza de archivos temporales.
- Definir límites:
  - número máximo de imágenes;
  - tamaño máximo por imagen;
  - tiempo máximo de generación.
- Agregar manejo de timeouts/reintentos con Azure.
- Agregar pruebas unitarias.
- Agregar pruebas de integración.
- Agregar monitoreo de costos.
- Definir si se guarda o no el JSON intermedio.
- Preparar formato futuro `minority_report` para cliente.

## Troubleshooting

### Error: Azure devuelve JSON inválido

El cliente ya intenta reparar errores comunes, como comas sobrantes o texto alrededor del JSON.

Si sigue fallando:

- revisar que `AZURE_FOUNDRY_JSON_SCHEMA=true`;
- revisar que `AZURE_FOUNDRY_MAX_OUTPUT_TOKENS=6000` o mayor;
- reducir cantidad/tamaño de imágenes;
- revisar el prompt;
- revisar si el modelo/deployment soporta salida estructurada.

### Error: el mini front no responde

Levantar desde consola:

```powershell
cd "C:\Users\saalc\OneDrive\Escritorio\XOC\backend y pruebas"
python mini-front/server.py
```

No usar Live Server de VS Code para probar generación, porque Live Server solo sirve archivos estáticos y no expone el endpoint `/api/generate`.

### El botón no cambia a "Generando"

Posibles causas:

- JS cacheado en el navegador;
- servidor viejo corriendo;
- se abrió con Live Server en vez de `python mini-front/server.py`.

Solución:

```text
Ctrl + F5
```

o reiniciar el mini front.

### Se genera el DOCX pero no cambia el diseño

Reiniciar el mini front. Si el servidor Python quedó abierto antes de los cambios, seguirá usando código viejo.

## Validaciones rápidas antes de push

Desde la raíz del workspace:

```powershell
python -m compileall -q .\prueba-script-xoc\azure_foundry_multimodal_client.py .\prueba-script-xoc\generate_report.py .\prueba-script-xoc\generate_small_report.py .\prueba-script-xoc\generate_support_report.py .\mini-front\server.py
node --check .\mini-front\app.js
```

Verificar que no se suban secretos:

```powershell
git status --short
```

Revisar especialmente que no aparezcan:

```text
.env
output/
uploads/
*.docx generado con datos reales
```

## Resumen para integración

El POC demuestra que el flujo funciona:

```text
Texto + imágenes
      ↓
Azure AI Foundry multimodal
      ↓
JSON dinámico validado
      ↓
Python DOCX generator
      ↓
Reporte Word con plantilla, diseño, portada, secciones e imágenes
```

El siguiente paso recomendado es adaptar el motor de Azure + generación DOCX al API principal, reemplazando el mini front por el frontend/backend real y agregando controles productivos de seguridad, almacenamiento, permisos y limpieza.
