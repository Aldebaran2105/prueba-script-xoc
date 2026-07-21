# Generador de reportes XOC

POC para generar reportes DOCX desde las plantillas locales usando:

- texto del analista
- una o varias imagenes/capturas locales
- datos base generados por el mini front, como cliente, caso y fecha
- Azure AI Foundry multimodal para producir JSON
- `python-docx` para generar el Word final

La IA no genera el DOCX. La IA devuelve solo JSON valido y Python arma el documento.
El cuerpo del reporte es dinamico: Azure solo debe devolver secciones sustentadas por el texto o las imagenes. Si no hay evidencia de acciones, resultados o recomendaciones, esas secciones no se agregan.

## Tipos soportados

- `small_report`
- `informe_soporte`

## Instalacion

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
```

Configura `.env`:

```env
USE_AZURE_FOUNDRY=true
USE_AZURE_FOUNDRY_AGENT=false
AZURE_FOUNDRY_PROJECT_ENDPOINT=https://aoai-sophia-xoc-eus2.services.ai.azure.com/api/projects/sophia-project
AZURE_FOUNDRY_OPENAI_ENDPOINT=https://aoai-sophia-xoc-eus2.services.ai.azure.com/openai/v1
AZURE_FOUNDRY_MODEL_DEPLOYMENT=gpt-5-mini
REPORT_TYPE=small_report
```

Si no usas API key, ejecuta antes:

```powershell
az login
```

El cliente usa `DefaultAzureCredential` cuando no hay `AZURE_FOUNDRY_API_KEY` ni `AZURE_OPENAI_API_KEY`.

## Flujo

```text
texto + imagenes + datos base
        ↓
Azure AI Foundry multimodal
        ↓
output/generated_report_payload.json
        ↓
python-docx
        ↓
output/*.docx
```

No se imprimen imagenes ni base64 en consola. Las imagenes se leen desde ruta local y se envian como `data:image/...;base64` al modelo/agente.

## Small Report

```powershell
python generate_small_report.py `
  --text "Validar evidencia de asignacion DHCP para radioenlace." `
  --data data/small_report_example.json `
  --image evidencias/captura_1.png
```

Modo local sin Azure:

```powershell
python generate_small_report.py --text "Prueba local" --data data/small_report_example.json --no-azure
```

JSON esperado desde Azure:

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

## Informe de Soporte

```powershell
python generate_support_report.py `
  --text "Renovacion de certificado local Cisco ISE." `
  --data data/informe_soporte_example.json `
  --image evidencias/certificado_anterior.png `
  --image evidencias/validacion_radius.png
```

Modo local sin Azure:

```powershell
python generate_support_report.py --text "Prueba local" --data data/informe_soporte_example.json --no-azure
```

JSON esperado desde Azure:

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

## Probar Foundry

```powershell
python test_foundry_multimodal_connection.py --report-type small_report
```

Con imagen:

```powershell
python test_foundry_multimodal_connection.py --report-type informe_soporte --image evidencias/captura.png
```

El test solo imprime estado, tipo de reporte, cantidad de imagenes y claves del JSON.

## Variables

- `USE_AZURE_FOUNDRY=true|false`: activa o desactiva Azure.
- `USE_AZURE_FOUNDRY_AGENT=true|false`: usa Agent de Foundry si esta en `true`; si esta en `false`, usa Responses API.
- `AZURE_FOUNDRY_PROJECT_ENDPOINT`: endpoint del proyecto Foundry.
- `AZURE_FOUNDRY_OPENAI_ENDPOINT`: endpoint compatible con OpenAI Responses API, por ejemplo `https://...services.ai.azure.com/openai/v1`.
- `AZURE_FOUNDRY_MODEL_DEPLOYMENT`: deployment multimodal, por ejemplo `gpt-5-mini`.
- `AZURE_FOUNDRY_AGENT_NAME`: nombre/id del agente si se usa Agent.
- `AZURE_FOUNDRY_AGENT_VERSION`: version opcional del agente.
- `REPORT_TYPE=small_report|informe_soporte`: tipo por defecto para pruebas.
- `REPORT_MAX_IMAGE_MB`: limite local por imagen.

## Notas de seguridad

- No hardcodear credenciales.
- No subir `.env`.
- No loguear imagenes, base64 ni capturas sensibles.
- La IA debe decir limitaciones cuando algo no se confirma desde texto o imagen.
- El backend productivo no se modifica desde este POC.

## Mini front local

Para probar texto + imagenes desde navegador sin tocar el backend productivo:

```powershell
cd "C:\Users\saalc\OneDrive\Escritorio\XOC\backend y pruebas"
python mini-front/server.py
```

Abre:

```text
http://127.0.0.1:8765
```

El mini front usa Azure real por defecto y no permite fallback local salvo que marques
la opcion en pantalla. Las imagenes se envian al server local, se guardan temporalmente
en `mini-front/uploads/` y no se imprimen en consola. Por cada imagen puedes escribir
una descripcion breve; el reporte la usa como caption tipo `Figura 1. Revision de
metricas generales` y coloca la imagen cerca del parrafo que cite esa figura.

## Comando anterior

El comando `generate_report.py` se mantiene para generar reportes desde JSON estructurado sin la capa multimodal:

```powershell
python generate_report.py small data/small_report_example.json --no-azure
python generate_report.py soporte data/informe_soporte_example.json --no-azure
```
