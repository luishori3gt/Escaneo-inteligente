"""
utils/email_sender.py
Módulo para envío de correos SMTP con adjuntos dinámicos por proveedor.
Incluye búsqueda local de directorio de proveedores y matching fuzzy.
"""

import io
import os
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from difflib import SequenceMatcher
from typing import Any

import pandas as pd

import config


def _separar_emails(texto_emails: str) -> list[str]:
    """
    Separa un string que puede contener múltiples emails
    separados por espacios, comas, punto y coma, o combinaciones.
    Filtra strings que no son emails válidos.
    """
    if not texto_emails:
        return []
    # Reemplazar separadores comunes por un separador único
    texto = texto_emails.replace(",", " ").replace(";", " ")
    # Dividir por espacios
    partes = texto.split()
    # Filtrar solo lo que parece email (contiene @ y .)
    emails = [p.strip() for p in partes if "@" in p and "." in p]
    return emails


# ──────────────────────────────────────────────
# Directorio de proveedores
# ──────────────────────────────────────────────

def cargar_emails_cc() -> list[str]:
    """
    Carga los emails para ir en copia (CC) desde la hoja
    'Para ir en todos los correos' del directorio_proveedores.xlsx.
    """
    ruta = config.DIRECTORIO_PROVEEDORES_FILE
    if not os.path.exists(ruta):
        return []
    try:
        xl = pd.ExcelFile(ruta)
        nombre_hoja = "Para ir en todos los correos"
        if nombre_hoja not in xl.sheet_names:
            return []
        df = pd.read_excel(ruta, sheet_name=nombre_hoja)
        # Los emails están en formato "Nombre" <email> en una sola celda
        texto = ""
        for col in df.columns:
            valores = df[col].astype(str).tolist()
            texto += " ".join(valores) + " "
        # Extraer emails con regex: buscar patrón <email> o email directo
        emails_encontrados = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', texto)
        # Filtrar duplicados manteniendo orden
        vistos = set()
        emails = []
        for e in emails_encontrados:
            e = e.lower().strip()
            if e not in vistos and "." in e:
                vistos.add(e)
                emails.append(e)
        return emails
    except Exception as e:
        print(f"[CC] Error al cargar: {e}")
        return []

def cargar_directorio_proveedores() -> pd.DataFrame | None:
    """
    Carga el archivo local directorio_proveedores.xlsx.
    Lee las hojas 'Proveedores City' y 'Proveedores Fresko' y las combina.
    Retorna None si el archivo no existe.
    """
    ruta = config.DIRECTORIO_PROVEEDORES_FILE
    if not os.path.exists(ruta):
        return None
    try:
        xl = pd.ExcelFile(ruta)
        hojas = []
        for nombre_hoja in ["Proveedores City", "Proveedores Fresko"]:
            if nombre_hoja in xl.sheet_names:
                df_hoja = pd.read_excel(ruta, sheet_name=nombre_hoja)
                df_hoja.columns = df_hoja.columns.str.strip().str.lower()
                # Renombrar columnas a un formato estándar
                rename_map = {}
                for col in df_hoja.columns:
                    if "nombre" in col and "corto" not in col:
                        rename_map[col] = "nombre"
                    elif "correo" in col or "email" in col:
                        rename_map[col] = "correo"
                df_hoja = df_hoja.rename(columns=rename_map)
                if "nombre" in df_hoja.columns and "correo" in df_hoja.columns:
                    df_hoja = df_hoja[["nombre", "correo"]].copy()
                    df_hoja["origen"] = nombre_hoja
                    hojas.append(df_hoja)

        if not hojas:
            return None

        df = pd.concat(hojas, ignore_index=True)
        # Filtrar filas sin email válido
        df["correo"] = df["correo"].astype(str).str.strip()
        df = df[
            (df["correo"] != "nan")
            & (df["correo"] != "0")
            & (df["correo"] != "Sin Correo")
            & (df["correo"] != "")
        ]
        return df
    except Exception as e:
        print(f"[Directorio] Error al cargar: {e}")
        return None


def buscar_email_proveedor(nombre_proveedor: str, df_directorio: pd.DataFrame) -> str | None:
    """
    Busca el email de un proveedor usando matching fuzzy.
    Compara el nombre extraído contra la columna 'nombre' del directorio.

    Retorna el email si la confianza >= 60%, o None si no hay match.
    """
    if df_directorio is None or df_directorio.empty:
        return None

    col_nombre = "nombre"
    col_email = "correo"

    if col_nombre not in df_directorio.columns or col_email not in df_directorio.columns:
        return None

    nombre_normalizado = _normalizar_texto(nombre_proveedor)
    mejor_score = 0.0
    mejor_email = None

    for _, row in df_directorio.iterrows():
        nombre_dir = str(row[col_nombre]).strip()
        if not nombre_dir or nombre_dir == "nan":
            continue

        nombre_dir_norm = _normalizar_texto(nombre_dir)
        score = _calcular_similitud(nombre_normalizado, nombre_dir_norm)

        if score > mejor_score:
            email = str(row[col_email]).strip()
            if email and email.lower() != "nan":
                mejor_score = score
                mejor_email = email

    if mejor_score >= config.FUZZY_MATCH_THRESHOLD:
        return mejor_email
    return None


def _normalizar_texto(texto: str) -> str:
    """Normaliza texto para comparación: minúsculas, sin acentos, sin caracteres especiales."""
    if not texto:
        return ""
    texto = texto.lower().strip()
    # Quitar acentos
    reemplazos = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "ñ": "n", "ü": "u",
    }
    for acento, sin_acento in reemplazos.items():
        texto = texto.replace(acento, sin_acento)
    # Quitar sufijos comunes de empresas
    texto = re.sub(
        r"\b(s\.a\.? de c\.v\.?|s\.a\.?|s\.r\.l\.?|inc\.?|corp\.?)\b",
        "",
        texto,
    )
    # Quitar caracteres no alfanuméricos
    texto = re.sub(r"[^a-z0-9\s]", "", texto)
    # Quitar espacios extra
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def _calcular_similitud(texto1: str, texto2: str) -> float:
    """
    Calcula similitud entre dos textos usando:
    1. SequenceMatcher (ratio de secuencia)
    2. Overlap de tokens/palabras clave
    3. Coincidencia de palabras parciales
    Toma el mayor de los scores.
    """
    if not texto1 or not texto2:
        return 0.0

    # 1. SequenceMatcher
    ratio = SequenceMatcher(None, texto1, texto2).ratio()

    # 2. Overlap de tokens
    tokens1 = set(texto1.split())
    tokens2 = set(texto2.split())
    if tokens1 and tokens2:
        interseccion = tokens1 & tokens2
        union = tokens1 | tokens2
        overlap = len(interseccion) / len(union)
    else:
        overlap = 0.0

    # 3. Coincidencia de tokens parciales (subcadenas)
    tokens_match = 0.0
    if tokens1 and tokens2:
        matches = 0
        for t1 in tokens1:
            for t2 in tokens2:
                if t1 in t2 or t2 in t1:
                    matches += 1
                    break
        tokens_match = matches / max(len(tokens1), len(tokens2))

    return max(ratio, overlap, tokens_match)


# ──────────────────────────────────────────────
# Generación de Excel adjunto por proveedor
# ──────────────────────────────────────────────

def generar_excel_proveedor(
    datos_proveedor: list[dict[str, Any]],
    proveedor: str,
) -> bytes:
    """
    Genera un Excel personalizado para un proveedor específico.
    Agrupa por sucursal/folio/factura sumando el total de cajas.
    Columnas: Proveedor, Sucursal, Fecha, Factura, Folio, Total Cajas.

    Args:
        datos_proveedor: Lista de dicts con los datos de las partidas del proveedor.
        proveedor: Nombre del proveedor.

    Returns:
        Bytes del archivo Excel generado.
    """
    # Agrupar por sucursal + folio + factura + fecha, sumando cajas
    grupos = {}
    for item in datos_proveedor:
        sucursal = item.get("sucursal_receptora", "No identificado")
        folio = item.get("folio_puerta", "No identificado")
        factura = item.get("numero_factura_folio", "No identificado")
        fecha = item.get("fecha", "No identificado")
        cajas = item.get("cajas", 0) or 0

        clave = (sucursal, folio, factura, fecha)
        if clave not in grupos:
            grupos[clave] = 0
        grupos[clave] += cajas

    filas = []
    for (sucursal, folio, factura, fecha), total_cajas in grupos.items():
        filas.append(
            {
                "Proveedor": proveedor,
                "Sucursal": sucursal,
                "Fecha": fecha,
                "Factura": factura,
                "Folio": folio,
                "Total Cajas": total_cajas,
            }
        )

    df = pd.DataFrame(filas)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resumen")
    buf.seek(0)
    return buf.getvalue()


# ──────────────────────────────────────────────
# Envío de correo SMTP
# ──────────────────────────────────────────────

def enviar_correo_proveedor(
    email_destino: str,
    proveedor: str,
    fecha_folio: str,
    pdf_bytes: bytes | None,
    pdf_nombre: str,
    excel_bytes: bytes,
    excel_nombre: str,
) -> dict[str, Any]:
    """
    Envía un correo SMTP con adjuntos a un proveedor.

    Asunto: "Folio de recibo operacion City y Fresko ({Proveedor}), ({Fecha})"

    Returns:
        Dict con 'success' (bool) y 'error' (str o None).
    """
    # Modo simulación
    if not config.SMTP_USER or not config.SMTP_PASSWORD:
        return {
            "success": True,
            "error": None,
            "simulacion": True,
            "mensaje": f"Correo simulado enviado a {email_destino}",
        }

    try:
        asunto = f"Folio de recibo operacion City y Fresko ({proveedor}), ({fecha_folio})"

        cuerpo = _generar_cuerpo_correo(proveedor, fecha_folio)

        msg = MIMEMultipart()
        msg["From"] = f"{config.SMTP_FROM_NAME} <{config.SMTP_USER}>"
        # Modo de prueba: enviar solo a lmartinezh@vpcom.com
        if os.getenv("EMAIL_TEST_MODE", "").lower() in ("1", "true", "yes"):
            emails = ["lmartinezh@vpcom.com"]
            emails_cc = []
            msg["To"] = ", ".join(emails)
            msg["Subject"] = f"[PRUEBA] {asunto}"
        else:
            emails = _separar_emails(email_destino)
            msg["To"] = ", ".join(emails)
            emails_cc = cargar_emails_cc()
            if emails_cc:
                msg["Cc"] = ", ".join(emails_cc)
            msg["Subject"] = asunto

        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

        # Adjuntar PDF
        if pdf_bytes:
            part_pdf = MIMEBase("application", "octet-stream")
            part_pdf.set_payload(pdf_bytes)
            encoders.encode_base64(part_pdf)
            part_pdf.add_header(
                "Content-Disposition",
                f'attachment; filename="{pdf_nombre}"',
            )
            msg.attach(part_pdf)

        # Adjuntar Excel
        part_excel = MIMEBase(
            "application",
            "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        part_excel.set_payload(excel_bytes)
        encoders.encode_base64(part_excel)
        part_excel.add_header(
            "Content-Disposition",
            f'attachment; filename="{excel_nombre}"',
        )
        msg.attach(part_excel)

        # Enviar (To + CC como destinatarios SMTP)
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.SMTP_USER, config.SMTP_PASSWORD)
        destinatarios = emails + emails_cc
        server.sendmail(config.SMTP_USER, destinatarios, msg.as_string())
        server.quit()

        return {"success": True, "error": None}

    except Exception as e:
        return {"success": False, "error": str(e)}


def _generar_cuerpo_correo(proveedor: str, fecha_folio: str) -> str:
    """Genera el cuerpo del correo en español."""
    return f"""Hola:

Se adjunta el folio de recibo correspondiente a la operación City y Fresko.

Proveedor: {proveedor}
Fecha del folio: {fecha_folio}

Se incluyen los siguientes documentos:
1. PDF del folio escaneado.
2. Resumen en Excel con el detalle de cajas y folio.

Favor de revisar y confirmar la recepción.

Saludos cordiales,
Equipo de VPC CEDA
"""


# ──────────────────────────────────────────────
# Generación de reportes Excel globales
# ──────────────────────────────────────────────

def generar_reporte_detallado(datos: list[dict[str, Any]]) -> bytes:
    """
    Genera el Reporte Detallado (una fila por partida).
    Columnas: Proveedor, Fecha, Número de Factura/Folio, Descripción, Cajas, Estatus de Validación.
    """
    filas = []
    for item in datos:
        for partida in item.get("partidas", []):
            filas.append(
                {
                    "Proveedor": item.get("proveedor", ""),
                    "Fecha": item.get("fecha", ""),
                    "Número de Factura/Folio": item.get("numero_factura_folio", ""),
                    "Descripción": partida.get("descripcion", ""),
                    "Cajas": partida.get("cajas", 0),
                    "Estatus de Validación": item.get("estatus_validacion", "Pendiente de Validación"),
                }
            )

    df = pd.DataFrame(filas)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reporte Detallado")
    buf.seek(0)
    return buf.getvalue()


def generar_resumen_totales(datos: list[dict[str, Any]]) -> bytes:
    """
    Genera el Resumen de Totales (una fila por factura).
    Columnas: Proveedor, Número de Factura/Folio, Fecha, Cajas Totales.
    """
    filas = []
    for item in datos:
        cajas_totales = sum(
            p.get("cajas", 0) for p in item.get("partidas", [])
        )
        filas.append(
            {
                "Proveedor": item.get("proveedor", ""),
                "Número de Factura/Folio": item.get("numero_factura_folio", ""),
                "Fecha": item.get("fecha", ""),
                "Cajas Totales": cajas_totales,
            }
        )

    df = pd.DataFrame(filas)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resumen de Totales")
    buf.seek(0)
    return buf.getvalue()
