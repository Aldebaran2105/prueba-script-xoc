"""Cliente POC para Azure AI Foundry multimodal.

La IA solo devuelve JSON. El DOCX lo genera Python en una etapa posterior.
"""

from __future__ import annotations

import base64
import json
import mimetypes
import os
import re
from pathlib import Path
from typing import Any

from azure.identity import DefaultAzureCredential
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "output"
DEFAULT_PAYLOAD_PATH = OUTPUT_DIR / "generated_report_payload.json"
SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}

DYNAMIC_REPORT_KEYS = {
    "title",
    "summary",
    "sections",
    "limitations",
    "image_citations",
}

LEGACY_SMALL_REPORT_KEYS = {
    "title",
    "summary",
    "observations",
    "risk_level",
    "affected_component",
    "technical_evidence",
    "recommended_actions",
    "conclusion",
}

LEGACY_SUPPORT_REPORT_KEYS = {
    "title",
    "case_context",
    "visual_evidence_summary",
    "technical_analysis",
    "impact",
    "actions_performed",
    "results_obtained",
    "pending_items",
    "recommended_actions",
    "conclusion",
    "analyst_notes",
}

DYNAMIC_JSON_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["title", "summary", "sections", "limitations", "image_citations"],
    "properties": {
        "title": {"type": "string"},
        "summary": {"type": "string"},
        "sections": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["title", "content", "bullets", "subsections"],
                "properties": {
                    "title": {"type": "string"},
                    "content": {"type": "string"},
                    "bullets": {"type": "array", "items": {"type": "string"}},
                    "subsections": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "required": ["title", "content", "bullets"],
                            "properties": {
                                "title": {"type": "string"},
                                "content": {"type": "string"},
                                "bullets": {"type": "array", "items": {"type": "string"}},
                            },
                        },
                    },
                },
            },
        },
        "limitations": {"type": "array", "items": {"type": "string"}},
        "image_citations": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["label", "description", "used_in_sections"],
                "properties": {
                    "label": {"type": "string"},
                    "description": {"type": "string"},
                    "used_in_sections": {"type": "array", "items": {"type": "string"}},
                },
            },
        },
    },
}

PROMPT_BASE = """Eres un analista senior XOC.
Debes analizar el texto y las imagenes proporcionadas para generar contenido de reporte.
Usa unicamente la evidencia entregada.
No inventes datos, fechas, IPs, activos, severidades, causas, acciones, resultados ni recomendaciones.
Si algo no se puede confirmar desde la imagen o texto, agregalo en limitations.
Cuando uses evidencia de una imagen, cita su etiqueta exacta dentro del texto, por ejemplo: Figura 1.
Si una imagen no aporta evidencia relacionada con el caso, indicalo como limitacion y no fuerces conclusiones.
No agregues secciones vacias, redundantes o con frases genericas como "no especificado".
No incluyas "Acciones Realizadas", "Resultados" o "Recomendaciones" si el input no confirma acciones, resultados o recomendaciones.
Devuelve SOLO JSON valido, sin markdown ni bloques de codigo.

Para small_report puedes usar, solo si aplica, esta estructura ideal:
1. Descripcion del Soporte
  1.1 Caso
  1.2 Fecha
  1.3 Forma de Conexion de Soporte
  1.4 Datos Base
2. Pasos realizados
  2.1 Relevamiento de Informacion
  2.2 Investigacion de posibles causas
  2.3 Revision de Procedimiento Onboarding
3. Acciones Realizadas
4. Resultados
5. Recomendaciones

Para informe_soporte puedes usar, solo si aplica, esta estructura ideal:
1. Descripcion del Soporte
  1.1 Tipo de Soporte
  1.2 Forma de Conexion de Soporte
  1.3 Objetivo
2. Antecedentes y Escenario Inicial
3. Desarrollo de la Actividad
  3.1 Acciones Realizadas
4. Resultados Obtenidos
5. Recomendaciones

Devuelve exactamente este JSON:
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
"""


def normalize_report_type(report_type: str) -> str:
    value = report_type.strip().lower().replace("-", "_")
    aliases = {
        "small": "small_report",
        "small_report": "small_report",
        "soporte": "informe_soporte",
        "support": "informe_soporte",
        "support_report": "informe_soporte",
        "informe": "informe_soporte",
        "informe_soporte": "informe_soporte",
    }
    if value not in aliases:
        raise ValueError("REPORT_TYPE debe ser small_report o informe_soporte")
    return aliases[value]


def expected_keys(report_type: str) -> set[str]:
    normalize_report_type(report_type)
    return DYNAMIC_REPORT_KEYS


def _strip_json_fence(raw: str) -> str:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.IGNORECASE)
    return raw.strip()


def _extract_json_object(raw: str) -> str:
    start = raw.find("{")
    end = raw.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return raw
    return raw[start : end + 1]


def _repair_common_json_issues(raw: str) -> str:
    repaired = _extract_json_object(raw)
    repaired = repaired.replace("\ufeff", "").replace("“", '"').replace("”", '"')
    repaired = re.sub(r"//.*?$", "", repaired, flags=re.MULTILINE)
    repaired = re.sub(r"/\*.*?\*/", "", repaired, flags=re.DOTALL)
    repaired = re.sub(r",\s*([}\]])", r"\1", repaired)
    return repaired.strip()


def _loads_json_with_repair(raw: str) -> dict[str, Any]:
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as first_exc:
        repaired = _repair_common_json_issues(raw)
        if repaired != raw:
            try:
                parsed = json.loads(repaired)
            except json.JSONDecodeError:
                raise first_exc
        else:
            raise
    if not isinstance(parsed, dict):
        raise RuntimeError("Azure no devolvio un objeto JSON")
    return parsed


def parse_and_validate_json(raw: str, report_type: str) -> dict[str, Any]:
    clean = _strip_json_fence(raw or "")
    if not clean:
        raise RuntimeError(
            "Azure no devolvio texto JSON visible. Posibles causas: respuesta vacia del modelo, "
            "limite de tokens insuficiente o el modelo no siguio el formato solicitado."
        )
    try:
        parsed = _loads_json_with_repair(clean)
    except json.JSONDecodeError as exc:
        sample = re.sub(r"\s+", " ", clean[:220]).strip()
        raise RuntimeError(
            "Azure devolvio una respuesta que no es JSON valido "
            f"({exc.msg}, linea {exc.lineno}, columna {exc.colno}). "
            f"Inicio seguro de la respuesta: {sample!r}"
        ) from exc
    if not isinstance(parsed, dict):
        raise RuntimeError("Azure no devolvio un objeto JSON")
    parsed = _normalize_legacy_payload(parsed, report_type)
    required = expected_keys(report_type)
    received = set(parsed)
    if received != required:
        missing = ", ".join(sorted(required - received)) or "ninguna"
        extra = ", ".join(sorted(received - required)) or "ninguna"
        raise RuntimeError(f"JSON inesperado. Faltan: {missing}. Sobran: {extra}.")
    parsed["title"] = str(parsed["title"]).strip()
    parsed["summary"] = str(parsed["summary"]).strip()
    if not isinstance(parsed["sections"], list):
        raise RuntimeError("El campo sections debe ser una lista")
    if not isinstance(parsed["limitations"], list):
        raise RuntimeError("El campo limitations debe ser una lista")
    if not isinstance(parsed["image_citations"], list):
        raise RuntimeError("El campo image_citations debe ser una lista")
    parsed["sections"] = [_normalize_section(item) for item in parsed["sections"] if _normalize_section(item)]
    parsed["limitations"] = [str(item).strip() for item in parsed["limitations"] if str(item).strip()]
    parsed["image_citations"] = [_normalize_image_citation(item) for item in parsed["image_citations"] if _normalize_image_citation(item)]
    return parsed


def _normalize_legacy_payload(parsed: dict[str, Any], report_type: str) -> dict[str, Any]:
    normalized = normalize_report_type(report_type)
    received = set(parsed)
    if received == DYNAMIC_REPORT_KEYS:
        return parsed
    if normalized == "small_report" and received == LEGACY_SMALL_REPORT_KEYS:
        sections: list[dict[str, Any]] = []
        if str(parsed.get("summary") or "").strip() or str(parsed.get("technical_evidence") or "").strip():
            sections.append(
                {
                    "title": "Descripcion del Soporte",
                    "content": "\n\n".join(
                        value
                        for value in (
                            str(parsed.get("summary") or "").strip(),
                            str(parsed.get("technical_evidence") or "").strip(),
                        )
                        if value
                    ),
                    "bullets": [],
                    "subsections": [],
                }
            )
        observations = [str(value).strip() for value in parsed.get("observations", []) if str(value).strip()]
        if observations:
            sections.append({"title": "Observaciones", "content": "", "bullets": observations, "subsections": []})
        recommendations = [str(value).strip() for value in parsed.get("recommended_actions", []) if str(value).strip()]
        if recommendations:
            sections.append({"title": "Recomendaciones", "content": "", "bullets": recommendations, "subsections": []})
        conclusion = str(parsed.get("conclusion") or "").strip()
        if conclusion:
            sections.append({"title": "Conclusion", "content": conclusion, "bullets": [], "subsections": []})
        return {
            "title": str(parsed.get("title") or "").strip(),
            "summary": str(parsed.get("summary") or "").strip(),
            "sections": sections,
            "limitations": [],
            "image_citations": [],
        }
    if normalized == "informe_soporte" and received == LEGACY_SUPPORT_REPORT_KEYS:
        pairs = [
            ("Descripcion del Soporte", parsed.get("case_context")),
            ("Antecedentes y Escenario Inicial", parsed.get("visual_evidence_summary")),
            ("Analisis Tecnico", parsed.get("technical_analysis")),
            ("Impacto", parsed.get("impact")),
            ("Acciones Realizadas", parsed.get("actions_performed")),
            ("Resultados Obtenidos", parsed.get("results_obtained")),
            ("Recomendaciones", parsed.get("recommended_actions")),
            ("Pendientes", parsed.get("pending_items")),
            ("Conclusion", parsed.get("conclusion")),
            ("Notas del Analista", parsed.get("analyst_notes")),
        ]
        sections = []
        for title, value in pairs:
            if isinstance(value, list):
                bullets = [str(item).strip() for item in value if str(item).strip()]
                if bullets:
                    sections.append({"title": title, "content": "", "bullets": bullets, "subsections": []})
            else:
                content = str(value or "").strip()
                if content:
                    sections.append({"title": title, "content": content, "bullets": [], "subsections": []})
        return {
            "title": str(parsed.get("title") or "").strip(),
            "summary": str(parsed.get("case_context") or parsed.get("visual_evidence_summary") or "").strip(),
            "sections": sections,
            "limitations": [],
            "image_citations": [],
        }
    return parsed


def _normalize_section(item: Any) -> dict[str, Any] | None:
    if not isinstance(item, dict):
        return None
    title = str(item.get("title") or "").strip()
    content = str(item.get("content") or "").strip()
    bullets = [str(value).strip() for value in item.get("bullets", []) if str(value).strip()] if isinstance(item.get("bullets"), list) else []
    subsections = [_normalize_section(value) for value in item.get("subsections", [])] if isinstance(item.get("subsections"), list) else []
    subsections = [value for value in subsections if value]
    if not title or (not content and not bullets and not subsections):
        return None
    return {"title": title, "content": content, "bullets": bullets, "subsections": subsections}


def _normalize_image_citation(item: Any) -> dict[str, Any] | None:
    if not isinstance(item, dict):
        return None
    label = str(item.get("label") or "").strip()
    description = str(item.get("description") or "").strip()
    used_in = item.get("used_in_sections", [])
    used_in_sections = [str(value).strip() for value in used_in if str(value).strip()] if isinstance(used_in, list) else []
    if not label and not description:
        return None
    return {"label": label, "description": description, "used_in_sections": used_in_sections}


def save_payload(payload: dict[str, Any], output_path: Path = DEFAULT_PAYLOAD_PATH) -> Path:
    output_path.parent.mkdir(exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


def image_to_data_url(image_path: Path) -> str:
    path = image_path.resolve()
    if not path.is_file():
        raise FileNotFoundError(f"No existe la imagen: {image_path}")
    if path.suffix.lower() not in SUPPORTED_IMAGE_EXTENSIONS:
        raise ValueError(f"Formato no soportado: {image_path.suffix}. Use PNG, JPG, JPEG o WEBP.")
    max_mb = float(os.environ.get("REPORT_MAX_IMAGE_MB", "10"))
    size_mb = path.stat().st_size / (1024 * 1024)
    if size_mb > max_mb:
        raise ValueError(f"La imagen excede REPORT_MAX_IMAGE_MB={max_mb:g}: {image_path.name}")
    mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def build_prompt(report_type: str, text: str, structured_data: dict[str, Any] | None, image_count: int) -> str:
    normalized = normalize_report_type(report_type)
    structured = structured_data or {}
    return (
        f"{PROMPT_BASE}\n\n"
        f"REPORT_TYPE={normalized}\n"
        f"Cantidad de imagenes recibidas: {image_count}\n\n"
        "Texto del analista:\n"
        f"{text.strip() or 'No se proporciono texto libre.'}\n\n"
        "Datos estructurados opcionales:\n"
        f"{json.dumps(structured, ensure_ascii=False, indent=2)}"
    )


def local_fallback_payload(report_type: str, text: str, image_paths: list[Path], structured_data: dict[str, Any] | None) -> dict[str, Any]:
    normalized = normalize_report_type(report_type)
    structured = structured_data or {}
    title = str(structured.get("title") or structured.get("subject") or "Reporte XOC")
    limitation = (
        f"Azure Foundry no fue utilizado. Se recibieron {len(image_paths)} imagen(es), "
        "pero no fueron interpretadas por un modelo multimodal."
    )
    if normalized == "small_report":
        return {
            "title": title,
            "summary": text.strip() or str(structured.get("summary") or "Borrador local generado sin analisis multimodal."),
            "sections": [
                {
                    "title": "Descripcion del Soporte",
                    "content": text.strip() or "Se recibio una solicitud de reporte sin texto descriptivo confirmado.",
                    "bullets": [],
                    "subsections": [],
                }
            ],
            "limitations": [limitation],
            "image_citations": [
                {
                    "label": f"Figura {index}",
                    "description": str(item.get("analyst_description") or item.get("file_name") or "Evidencia visual proporcionada."),
                    "used_in_sections": [],
                }
                for index, item in enumerate(structured.get("evidence_images", []), start=1)
            ],
        }
    return {
        "title": title,
        "summary": text.strip() or str(structured.get("summary") or "Borrador local generado sin analisis multimodal."),
        "sections": [
            {
                "title": "Descripcion del Soporte",
                "content": text.strip() or "Se recibio una solicitud de informe sin texto descriptivo confirmado.",
                "bullets": [],
                "subsections": [],
            }
        ],
        "limitations": [limitation],
        "image_citations": [
            {
                "label": f"Figura {index}",
                "description": str(item.get("analyst_description") or item.get("file_name") or "Evidencia visual proporcionada."),
                "used_in_sections": [],
            }
            for index, item in enumerate(structured.get("evidence_images", []), start=1)
        ],
    }


def _content_blocks(prompt: str, image_paths: list[Path]) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = [{"type": "input_text", "text": prompt}]
    for path in image_paths:
        blocks.append({"type": "input_image", "image_url": image_to_data_url(path)})
    return blocks


def _credential_token() -> str:
    credential = DefaultAzureCredential()
    return credential.get_token("https://ai.azure.com/.default").token


def _foundry_openai_client() -> Any:
    from openai import OpenAI

    endpoint = (
        os.environ.get("AZURE_FOUNDRY_OPENAI_ENDPOINT", "").strip()
        or os.environ.get("AZURE_OPENAI_ENDPOINT", "").strip()
        or os.environ["AZURE_FOUNDRY_PROJECT_ENDPOINT"].strip()
    ).rstrip("/")
    if endpoint.endswith("/openai/v1"):
        base_url = f"{endpoint}/"
    elif "/api/projects/" in endpoint:
        resource_root = endpoint.split("/api/projects/", 1)[0].rstrip("/")
        base_url = f"{resource_root}/openai/v1/"
    else:
        base_url = f"{endpoint}/openai/v1/"
    api_key = (
        os.environ.get("AZURE_FOUNDRY_API_KEY", "").strip()
        or os.environ.get("AZURE_OPENAI_API_KEY", "").strip()
        or _credential_token()
    )
    return OpenAI(base_url=base_url, api_key=api_key)


def _generate_with_responses_api(report_type: str, prompt: str, image_paths: list[Path]) -> dict[str, Any]:
    client = _foundry_openai_client()
    deployment = (
        os.environ.get("AZURE_FOUNDRY_MODEL_DEPLOYMENT", "").strip()
        or os.environ.get("AZURE_OPENAI_DEPLOYMENT", "").strip()
        or "gpt-5-mini"
    )
    request: dict[str, Any] = {
        "model": deployment,
        "input": [{"role": "user", "content": _content_blocks(prompt, image_paths)}],
        "max_output_tokens": int(os.environ.get("AZURE_FOUNDRY_MAX_OUTPUT_TOKENS", "6000")),
    }
    if os.environ.get("AZURE_FOUNDRY_JSON_SCHEMA", "true").strip().lower() not in {"0", "false", "no"}:
        request["text"] = {
            "format": {
                "type": "json_schema",
                "name": "xoc_dynamic_report_payload",
                "schema": DYNAMIC_JSON_SCHEMA,
                "strict": True,
            }
        }
    try:
        response = client.responses.create(**request)
    except Exception:
        if "text" not in request:
            raise
        request.pop("text", None)
        response = client.responses.create(**request)
    return parse_and_validate_json(_extract_response_text(response), report_type)


def _extract_response_text(response: Any) -> str:
    output_text = str(getattr(response, "output_text", "") or "").strip()
    if output_text:
        return output_text

    pieces: list[str] = []
    for item in getattr(response, "output", []) or []:
        item_text = getattr(item, "text", None)
        if item_text:
            pieces.append(str(item_text))
        for content in getattr(item, "content", []) or []:
            text = getattr(content, "text", None)
            if text:
                pieces.append(str(text))
            if isinstance(content, dict):
                value = content.get("text") or content.get("value")
                if value:
                    pieces.append(str(value))
        if isinstance(item, dict):
            value = item.get("text") or item.get("value")
            if value:
                pieces.append(str(value))
            for content in item.get("content", []) or []:
                if isinstance(content, dict):
                    value = content.get("text") or content.get("value")
                    if value:
                        pieces.append(str(value))
    text = "\n".join(piece.strip() for piece in pieces if piece and piece.strip())
    if text:
        return text
    status = getattr(response, "status", "")
    incomplete = getattr(response, "incomplete_details", None)
    raise RuntimeError(
        "Azure respondio, pero no devolvio texto visible para convertir a JSON. "
        f"status={status or 'desconocido'} incomplete_details={incomplete or 'sin_detalle'}"
    )


def _find_agent(agents: Any, agent_name: str, agent_version: str | None) -> Any:
    for method_name in ("get", "get_agent"):
        method = getattr(agents, method_name, None)
        if not method:
            continue
        try:
            return method(agent_name)
        except Exception:
            pass
    list_method = getattr(agents, "list_agents", None) or getattr(agents, "list", None)
    if list_method:
        for agent in list_method():
            name = str(getattr(agent, "name", ""))
            version = str(getattr(agent, "version", ""))
            if name == agent_name and (not agent_version or version in {"", agent_version}):
                return agent
    raise RuntimeError(f"No se encontro el agente de Foundry: {agent_name}")


def _agent_content_blocks(prompt: str, image_paths: list[Path]) -> list[Any]:
    try:
        from azure.ai.agents.models import MessageImageUrlParam, MessageInputImageUrlBlock, MessageInputTextBlock
    except Exception as exc:
        raise RuntimeError(
            "El SDK instalado no expone bloques de imagen para Agents. "
            "Use USE_AZURE_FOUNDRY_AGENT=false para probar Responses API."
        ) from exc
    blocks: list[Any] = [MessageInputTextBlock(text=prompt)]
    for path in image_paths:
        blocks.append(MessageInputImageUrlBlock(image_url=MessageImageUrlParam(url=image_to_data_url(path), detail="high")))
    return blocks


def _message_text(message: Any) -> str:
    content = getattr(message, "content", []) or []
    pieces: list[str] = []
    for item in content:
        text = getattr(item, "text", None)
        if text and hasattr(text, "value"):
            pieces.append(str(text.value))
        elif isinstance(item, dict):
            value = item.get("text") or item.get("value")
            if value:
                pieces.append(str(value))
    return "\n".join(pieces).strip()


def _generate_with_foundry_agent(report_type: str, prompt: str, image_paths: list[Path]) -> dict[str, Any]:
    from azure.ai.projects import AIProjectClient

    endpoint = os.environ["AZURE_FOUNDRY_PROJECT_ENDPOINT"].strip()
    agent_name = os.environ["AZURE_FOUNDRY_AGENT_NAME"].strip()
    agent_version = os.environ.get("AZURE_FOUNDRY_AGENT_VERSION", "").strip() or None
    client = AIProjectClient(endpoint=endpoint, credential=DefaultAzureCredential())
    agents = client.agents
    agent = _find_agent(agents, agent_name, agent_version)
    thread = agents.threads.create()
    agents.messages.create(thread_id=thread.id, role="user", content=_agent_content_blocks(prompt, image_paths))
    run = agents.runs.create_and_process(thread_id=thread.id, agent_id=agent.id)
    if str(getattr(run, "status", "")).lower() != "completed":
        raise RuntimeError(f"Foundry Agent no completo la ejecucion: {getattr(run, 'last_error', None)}")
    for message in agents.messages.list(thread_id=thread.id):
        if getattr(message, "role", "") != "assistant":
            continue
        text = _message_text(message)
        if text:
            return parse_and_validate_json(text, report_type)
    raise RuntimeError("Foundry Agent no devolvio contenido")


def generate_report_payload(
    report_type: str,
    *,
    text: str = "",
    image_paths: list[str | Path] | None = None,
    structured_data: dict[str, Any] | None = None,
    use_azure: bool | None = None,
    allow_local_fallback: bool = True,
    output_path: Path = DEFAULT_PAYLOAD_PATH,
) -> dict[str, Any]:
    load_dotenv(ROOT / ".env")
    normalized = normalize_report_type(report_type)
    images = [Path(path) for path in (image_paths or [])]
    if use_azure is None:
        use_azure = os.environ.get("USE_AZURE_FOUNDRY", "true").strip().lower() not in {"0", "false", "no"}

    if not use_azure:
        payload = local_fallback_payload(normalized, text, images, structured_data)
        save_payload(payload, output_path)
        return payload

    prompt = build_prompt(normalized, text, structured_data, len(images))
    try:
        use_agent = os.environ.get("USE_AZURE_FOUNDRY_AGENT", "false").strip().lower() in {"1", "true", "yes"}
        payload = (
            _generate_with_foundry_agent(normalized, prompt, images)
            if use_agent
            else _generate_with_responses_api(normalized, prompt, images)
        )
    except Exception:
        if not allow_local_fallback:
            raise
        payload = local_fallback_payload(normalized, text, images, structured_data)

    save_payload(payload, output_path)
    return payload
