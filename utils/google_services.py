"""
utils/google_services.py
Lógica de conexión con Google Drive API.
Soporta jerarquía dinámica de 3 niveles:
  [Proveedor] / [Fecha (YYYY-MM-DD)] / Folio_[Numero]_[Proveedor].pdf
"""

import io
from typing import Any

import config


def _obtener_credenciales():
    """
    Obtiene credenciales de Service Account de Google.
    Prioridad: string JSON (Streamlit secrets) > archivo JSON (local).
    Retorna None si no hay configuración (modo simulación).
    """
    try:
        from google.oauth2 import service_account
        from google.auth.transport.requests import Request

        # Prioridad 1: String JSON desde Streamlit secrets
        if config.GOOGLE_SERVICE_ACCOUNT_JSON:
            import json as _json
            info = _json.loads(config.GOOGLE_SERVICE_ACCOUNT_JSON)
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=config.GOOGLE_API_SCOPES
            )
            creds.refresh(Request())
            return creds

        # Prioridad 2: Archivo JSON local
        if config.GOOGLE_SERVICE_ACCOUNT_FILE:
            creds = service_account.Credentials.from_service_account_file(
                config.GOOGLE_SERVICE_ACCOUNT_FILE,
                scopes=config.GOOGLE_API_SCOPES,
            )
            creds.refresh(Request())
            return creds

        return None
    except Exception as e:
        print(f"[Google Drive] Error de credenciales: {e}")
        return None


def _obtener_servicio():
    """
    Construye el servicio de Google Drive.
    Retorna None si no hay credenciales.
    """
    creds = _obtener_credenciales()
    if creds is None:
        return None
    try:
        from googleapiclient.discovery import build

        return build("drive", "v3", credentials=creds, static_discovery=False)
    except Exception as e:
        print(f"[Google Drive] Error al construir servicio: {e}")
        return None


def _buscar_carpeta(
    servicio, nombre: str, carpeta_padre_id: str | None = None
) -> str | None:
    """Busca una carpeta por nombre dentro de una carpeta padre."""
    query = f"mimeType='application/vnd.google-apps.folder' and name='{nombre}'"
    if carpeta_padre_id:
        query += f" and '{carpeta_padre_id}' in parents"
    query += " and trashed=false"

    resultado = (
        servicio.files()
        .list(q=query, fields="files(id, name)", pageSize=1)
        .execute()
    )
    archivos = resultado.get("files", [])
    return archivos[0]["id"] if archivos else None


def _crear_carpeta(
    servicio, nombre: str, carpeta_padre_id: str | None = None
) -> str:
    """Crea una carpeta en Google Drive y retorna su ID."""
    metadata: dict[str, Any] = {
        "name": nombre,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if carpeta_padre_id:
        metadata["parents"] = [carpeta_padre_id]

    carpeta = servicio.files().create(body=metadata, fields="id").execute()
    return carpeta["id"]


def _obtener_o_crear_carpeta(
    servicio, nombre: str, carpeta_padre_id: str | None = None
) -> str:
    """Busca o crea una carpeta y retorna su ID."""
    carpeta_id = _buscar_carpeta(servicio, nombre, carpeta_padre_id)
    if carpeta_id:
        return carpeta_id
    return _crear_carpeta(servicio, nombre, carpeta_padre_id)


def subir_micro_pdf(
    pdf_bytes: bytes,
    nombre_archivo: str,
    proveedor: str,
    fecha: str,
) -> dict[str, Any]:
    """
    Sube un micro-PDF a Google Drive siguiendo la jerarquía:
    [Proveedor] / [Fecha YYYY-MM-DD] / [nombre_archivo]

    Args:
        pdf_bytes: Bytes del micro-PDF.
        nombre_archivo: Nombre del archivo PDF.
        proveedor: Nombre del proveedor (nivel 1).
        fecha: Fecha del folio en formato YYYY-MM-DD (nivel 2).

    Returns:
        Dict con 'success', 'file_id', 'url' o 'error'.
    """
    servicio = _obtener_servicio()

    # Modo simulación
    if servicio is None:
        return {
            "success": True,
            "file_id": "SIMULADO",
            "url": f"https://drive.google.com/file/d/SIMULADO/view",
            "simulacion": True,
            "ruta": f"{proveedor}/{fecha}/{nombre_archivo}",
        }

    try:
        from googleapiclient.http import MediaIoBaseUpload

        # Nivel 1: Carpeta del proveedor
        carpeta_proveedor = _obtener_o_crear_carpeta(
            servicio, proveedor, config.GOOGLE_DRIVE_MASTER_FOLDER_ID
        )

        # Nivel 2: Carpeta de fecha
        carpeta_fecha = _obtener_o_crear_carpeta(
            servicio, fecha, carpeta_proveedor
        )

        # Nivel 3: Subir el archivo PDF
        media = MediaIoBaseUpload(
            io.BytesIO(pdf_bytes), mimetype="application/pdf", resumable=True
        )
        metadata = {
            "name": nombre_archivo,
            "parents": [carpeta_fecha],
        }

        archivo = (
            servicio.files()
            .create(body=metadata, media_body=media, fields="id, webViewLink")
            .execute()
        )

        file_id = archivo.get("id", "")
        url = archivo.get("webViewLink", "")

        return {
            "success": True,
            "file_id": file_id,
            "url": url,
            "ruta": f"{proveedor}/{fecha}/{nombre_archivo}",
        }

    except Exception as e:
        return {
            "success": False,
            "file_id": None,
            "url": None,
            "error": str(e),
        }


def es_simulacion() -> bool:
    """Retorna True si Google Drive no está configurado."""
    return not config.GOOGLE_SERVICE_ACCOUNT_FILE and not config.GOOGLE_SERVICE_ACCOUNT_JSON
