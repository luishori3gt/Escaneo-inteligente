"""
utils/openai_client.py
Cliente de extracción de datos de folios desde imágenes.
Soporta Google Gemini 1.5 Pro (prioridad) y OpenAI GPT-4o (fallback).
"""

import json
import base64
import time
from typing import Any

import config


# Esquema JSON para Structured Outputs
FOLIO_JSON_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "proveedor": {
            "type": "string",
            "description": "Razón social del proveedor. Si no aparece, usar 'No identificado'.",
        },
        "sucursal_receptora": {
            "type": "string",
            "description": "Nombre o código de la sucursal que recibe la mercancía. Si no aparece, usar 'No identificado'.",
        },
        "numero_factura_folio": {
            "type": "string",
            "description": "Número de factura o remisión. Si no aparece, usar 'No identificado'.",
        },
        "folio_puerta": {
            "type": "string",
            "description": "Folio de puerta asignado. Si no aparece, usar 'No identificado'.",
        },
        "acuse_de_recibo": {
            "type": "string",
            "description": "Fecha del acuse de recibo en formato YYYY-MM-DD si es posible. Si no aparece, usar 'No identificado'.",
        },
        "fecha": {
            "type": "string",
            "description": "Fecha del folio en formato YYYY-MM-DD si es posible. Si no aparece, usar 'No identificado'.",
        },
        "partidas": {
            "type": "array",
            "description": "Lista de partidas/artículos del folio",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "descripcion": {
                        "type": "string",
                        "description": "Descripción del artículo/producto",
                    },
                    "cajas": {
                        "type": "integer",
                        "description": "Número de cajas del artículo",
                    },
                    "piezas": {
                        "type": "integer",
                        "description": "Número de piezas del artículo. Si no aparece, usar 0.",
                    },
                },
                "required": ["descripcion", "cajas", "piezas"],
            },
        },
    },
    "required": [
        "proveedor",
        "sucursal_receptora",
        "numero_factura_folio",
        "folio_puerta",
        "acuse_de_recibo",
        "fecha",
        "partidas",
    ],
}

# Prompt de extracción (compartido entre Gemini y OpenAI)
PROMPT_EXTRACCION = (
    "Analiza esta imagen de un Acuse de Recibo logístico escaneado. "
    "Extrae TODA la información visible. Lee con mucho cuidado los textos "
    "escritos a mano y los números pequeños. Incluye TODAS las "
    "partidas/artículos de la tabla con su descripción, número de cajas y piezas.\n\n"
    "Devuelve EXACTAMENTE este formato JSON:\n"
    "{\n"
    '  "proveedor": "razón social del proveedor",\n'
    '  "sucursal_receptora": "sucursal que recibe",\n'
    '  "numero_factura_folio": "número de factura",\n'
    '  "folio_puerta": "folio de puerta",\n'
    '  "acuse_de_recibo": "fecha YYYY-MM-DD",\n'
    '  "fecha": "fecha YYYY-MM-DD",\n'
    '  "partidas": [\n'
    '    {"descripcion": "producto", "cajas": 0, "piezas": 0}\n'
    "  ]\n"
    "}\n\n"
    "Si un campo no aparece o no es legible, usa 'No identificado'.\n"
    "Si no hay partidas visibles, devuelve partidas como array vacío []."
)


def extraer_datos_folio(imagen_bytes: bytes) -> dict[str, Any]:
    """
    Envía una imagen de folio y extrae datos estructurados.
    Usa Gemini 1.5 Pro si hay API key, sino OpenAI GPT-4o, sino simulación.
    """
    if config.EXTRACTION_MODEL == "gemini":
        return _extraer_con_gemini(imagen_bytes)
    elif config.EXTRACTION_MODEL == "openai":
        return _extraer_con_openai(imagen_bytes)
    else:
        return _datos_simulacion()


def _extraer_con_gemini(imagen_bytes: bytes) -> dict[str, Any]:
    """Extrae datos usando Google Gemini. Si se excede cuota, falla a OpenAI."""
    try:
        import google.generativeai as genai

        genai.configure(api_key=config.GEMINI_API_KEY)

        model = genai.GenerativeModel(
            config.GEMINI_MODEL,
            system_instruction=config.OPENAI_SYSTEM_PROMPT,
        )

        img_part = {
            "mime_type": "image/jpeg",
            "data": imagen_bytes,
        }

        max_reintentos = 2
        for intento in range(max_reintentos + 1):
            try:
                response = model.generate_content(
                    [PROMPT_EXTRACCION, img_part],
                    generation_config={
                        "temperature": 0,
                        "max_output_tokens": 8192,
                        "response_mime_type": "application/json",
                    },
                    request_options={"timeout": 120},
                )

                contenido = response.text
                datos = json.loads(contenido)

                # Asegurar campos requeridos
                for campo in ["proveedor", "sucursal_receptora", "numero_factura_folio",
                              "folio_puerta", "acuse_de_recibo", "fecha", "partidas"]:
                    if campo not in datos:
                        datos[campo] = "No identificado" if campo != "partidas" else []
                if not isinstance(datos["partidas"], list):
                    datos["partidas"] = []

                return datos

            except Exception as e:
                error_str = str(e)
                # Si es error de cuota 429, hacer fallback a OpenAI
                if "429" in error_str and config.OPENAI_API_KEY:
                    print("[Gemini] Cuota excedida, fallback a OpenAI...")
                    return _extraer_con_openai(imagen_bytes)
                if intento < max_reintentos:
                    time.sleep(3)
                    continue
                return _respuesta_error(str(e))

    except Exception as e:
        error_str = str(e)
        if "429" in error_str and config.OPENAI_API_KEY:
            print("[Gemini] Cuota excedida, fallback a OpenAI...")
            return _extraer_con_openai(imagen_bytes)
        return _respuesta_error(f"Gemini init error: {e}")


def _extraer_con_openai(imagen_bytes: bytes) -> dict[str, Any]:
    """Extrae datos usando OpenAI GPT-4o."""
    from openai import OpenAI

    cliente = OpenAI(api_key=config.OPENAI_API_KEY)
    data_url = _imagen_a_base64(imagen_bytes)

    max_reintentos = 2
    for intento in range(max_reintentos + 1):
        try:
            response = cliente.chat.completions.create(
                model=config.OPENAI_MODEL,
                messages=[
                    {
                        "role": "system",
                        "content": config.OPENAI_SYSTEM_PROMPT,
                    },
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": PROMPT_EXTRACCION},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": data_url,
                                    "detail": "high",
                                },
                            },
                        ],
                    },
                ],
                response_format={
                    "type": "json_schema",
                    "json_schema": {
                        "name": "folio_recibo",
                        "strict": True,
                        "schema": FOLIO_JSON_SCHEMA,
                    },
                },
                max_tokens=4000,
                temperature=0,
                timeout=120,
            )

            contenido = response.choices[0].message.content
            return json.loads(contenido)

        except Exception as e:
            if intento < max_reintentos:
                time.sleep(3)
                continue
            return _respuesta_error(str(e))


def _imagen_a_base64(imagen_bytes: bytes) -> str:
    """Convierte bytes de imagen a base64 data URL (JPEG)."""
    b64 = base64.b64encode(imagen_bytes).decode("utf-8")
    return f"data:image/jpeg;base64,{b64}"


def _respuesta_error(error: str) -> dict[str, Any]:
    """Retorna dict de error estandarizado."""
    return {
        "proveedor": "No identificado",
        "sucursal_receptora": "No identificado",
        "numero_factura_folio": "No identificado",
        "folio_puerta": "No identificado",
        "acuse_de_recibo": "No identificado",
        "fecha": "No identificado",
        "partidas": [],
        "_error": error,
    }


def _datos_simulacion() -> dict[str, Any]:
    """Genera datos de ejemplo para modo simulación (sin API key)."""
    return {
        "proveedor": "PROVEEDOR EJEMPLO (SIMULACIÓN)",
        "sucursal_receptora": "BODEGA CENTRAL",
        "numero_factura_folio": "FAC-000001",
        "folio_puerta": "FP-001",
        "acuse_de_recibo": "2026-06-20",
        "fecha": "2026-06-20",
        "partidas": [
            {"descripcion": "Producto Ejemplo A", "cajas": 10, "piezas": 120},
            {"descripcion": "Producto Ejemplo B", "cajas": 5, "piezas": 60},
        ],
        "_simulacion": True,
    }
