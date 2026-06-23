"""
config.py
Configuración central de variables de entorno, credenciales y parámetros.
Lee desde archivo .env si existe; de lo contrario usa valores por defecto (vacíos)
que activan el "modo simulación" para pruebas locales.
"""

import os
from dotenv import load_dotenv

# Cargar .env si existe (desarrollo local)
load_dotenv()

# En Streamlit Cloud, los secrets se configuran en el panel
# Intentamos leerlos si están disponibles
_st_secrets = {}
try:
    import streamlit as st
    _st_secrets = dict(st.secrets) if hasattr(st, "secrets") else {}
except Exception:
    pass


def _get(key: str, default: str = "") -> str:
    """Lee variable de Streamlit secrets primero, luego .env, luego default."""
    if key in _st_secrets:
        return str(_st_secrets[key])
    return os.getenv(key, default)


# ──────────────────────────────────────────────
# OpenAI
# ──────────────────────────────────────────────
OPENAI_API_KEY: str = _get("OPENAI_API_KEY")
OPENAI_MODEL: str = "gpt-4o"

# ──────────────────────────────────────────────
# Google Gemini
# ──────────────────────────────────────────────
GEMINI_API_KEY: str = _get("GEMINI_API_KEY")
GEMINI_MODEL: str = "gemini-2.5-pro"

# Modelo activo para extracción: "gemini" o "openai"
# Gemini tiene prioridad si hay API key configurada
EXTRACTION_MODEL: str = "gemini" if GEMINI_API_KEY else ("openai" if OPENAI_API_KEY else "none")

# Prompt del sistema para extracción estructurada
OPENAI_SYSTEM_PROMPT: str = (
    "Eres un asistente experto en extracción de datos de documentos logísticos "
    "escaneados en español. Analiza la imagen del folio de recibo (Acuse de Recibo) "
    "y extrae la información en formato JSON estructurado según el esquema indicado.\n\n"
    "INSTRUCCIONES CRÍTICAS:\n"
    "1. PROVEEDOR: Es la RAZÓN SOCIAL DEL PROVEEDOR que aparece en el campo "
    "'Razón Social del Proveedor' del formulario. Es la empresa que provee/vende "
    "la mercancía (ej: FRUTAS DEL TRÓPICO, HORTALIZAS DEL VALLE, etc.). "
    "NO es 'COMERCIAL CITY FRESKO' — esa es la empresa receptora que aparece "
    "en el encabezado. Extrae el nombre exactamente como aparece escrito.\n"
    "2. SUCURSAL RECEPTORA: Es la bodega o sucursal de City Fresko que recibe. "
    "Puede ser un nombre, número o código de bodega.\n"
    "3. NUMERO_FACTURA_FOLIO: Es el número de factura o remisión del proveedor.\n"
    "4. FOLIO_PUERTA: Es el folio de puerta asignado por la bodega/sucursal.\n"
    "5. ACUSE_DE_RECIBO: Es la fecha del acuse de recibo (cuando se recibió).\n"
    "6. FECHA: La fecha principal del documento.\n"
    "7. PARTIDAS: Lista cada artículo/producto de la tabla con su descripción, "
    "número de cajas y piezas. Lee con cuidado los números escritos a mano.\n\n"
    "Si un campo no es legible o no aparece, usa 'No identificado'."
)


# ──────────────────────────────────────────────
# Google Drive
# ──────────────────────────────────────────────
GOOGLE_SERVICE_ACCOUNT_FILE: str = _get("GOOGLE_SERVICE_ACCOUNT_FILE")
GOOGLE_DRIVE_MASTER_FOLDER_ID: str = _get("GOOGLE_DRIVE_MASTER_FOLDER_ID")
GOOGLE_API_SCOPES: list[str] = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
]


# ──────────────────────────────────────────────
# SMTP Email
# ──────────────────────────────────────────────
SMTP_SERVER: str = _get("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT: int = int(_get("SMTP_PORT", "587"))
SMTP_USER: str = _get("SMTP_USER")
SMTP_PASSWORD: str = _get("SMTP_PASSWORD")
SMTP_FROM_NAME: str = "Sistema de Folios - Operación City y Fresko"


# ──────────────────────────────────────────────
# Archivos y directorios locales
# ──────────────────────────────────────────────
DIRECTORIO_PROVEEDORES_FILE: str = "directorio_proveedores.xlsx"
TEMP_DIR: str = "temp_pdfs"


# ──────────────────────────────────────────────
# Parámetros de la aplicación
# ──────────────────────────────────────────────
MAX_EXCEL_SIZE_MB: int = 50
MAX_PDF_SIZE_MB: int = 200
FUZZY_MATCH_THRESHOLD: float = 0.60


# ──────────────────────────────────────────────
# Modo simulación
# ──────────────────────────────────────────────
def is_simulation_mode() -> bool:
    """Retorna True si no hay credenciales configuradas (modo simulación)."""
    return (
        not OPENAI_API_KEY
        and not GEMINI_API_KEY
        and not GOOGLE_SERVICE_ACCOUNT_FILE
        and not SMTP_USER
    )
