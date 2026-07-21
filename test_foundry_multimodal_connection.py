"""Prueba minima de conexion a Azure AI Foundry multimodal.

No imprime imagenes, base64 ni contenido completo del reporte.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path

from dotenv import load_dotenv

from azure_foundry_multimodal_client import generate_report_payload, normalize_report_type

ROOT = Path(__file__).resolve().parent


def main() -> None:
    parser = argparse.ArgumentParser(description="Valida conexion Foundry multimodal y JSON estricto.")
    parser.add_argument("--report-type", default=os.environ.get("REPORT_TYPE", "small_report"))
    parser.add_argument("--text", default="Prueba de conexion. Responde un JSON valido usando solo esta evidencia.")
    parser.add_argument("--image", action="append", default=[], type=Path, help="Imagen local opcional. Se puede repetir.")
    parser.add_argument("--allow-local-fallback", action="store_true", help="Permite validar el flujo local si Azure falla.")
    args = parser.parse_args()
    load_dotenv(ROOT / ".env")

    report_type = normalize_report_type(args.report_type)
    payload = generate_report_payload(
        report_type,
        text=args.text,
        image_paths=args.image,
        structured_data={},
        use_azure=True,
        allow_local_fallback=args.allow_local_fallback,
    )
    print("Conexion/flujo OK")
    print(f"REPORT_TYPE: {report_type}")
    print(f"Imagenes enviadas: {len(args.image)}")
    print("Claves JSON:", ", ".join(sorted(payload)))


if __name__ == "__main__":
    main()
