"""
utils/pdf_processor.py
Lógica para conteo, división y conversión de PDFs a texto/imagenes.
Usa PyMuPDF (fitz) para conversión a imagen — no requiere Poppler.
"""

import io
import os
from typing import Any

import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageEnhance, ImageFilter

import config


def _mejorar_imagen(img: Image.Image) -> Image.Image:
    """
    Preprocesa la imagen para mejorar la legibilidad de escaneos de baja calidad.
    - Convierte a escala de grises
    - Aumenta contraste
    - Aumenta nitidez
    - Optimiza tamaño
    """
    img = img.convert("L")  # Escala de grises
    img = ImageEnhance.Contrast(img).enhance(1.8)  # Más contraste
    img = ImageEnhance.Sharpness(img).enhance(2.5)  # Más nitidez
    img = img.filter(ImageFilter.MedianFilter(size=3))  # Reducir ruido
    return img


def contar_paginas(pdf_bytes: bytes) -> int:
    """Cuenta el número de páginas de un PDF."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    return len(reader.pages)


def extraer_texto_pagina(pdf_bytes: bytes, pagina_idx: int) -> str:
    """
    Intenta extraer texto nativo de una página específica.
    Retorna string vacío si no hay texto (PDF escaneado).
    """
    reader = PdfReader(io.BytesIO(pdf_bytes))
    if pagina_idx < len(reader.pages):
        return reader.pages[pagina_idx].extract_text() or ""
    return ""


def convertir_pagina_a_imagen(
    pdf_bytes: bytes, pagina_idx: int, dpi: int = 250
) -> bytes:
    """
    Convierte una página específica del PDF a imagen (bytes).
    Usa PyMuPDF (fitz) + preprocesamiento con Pillow para mejorar legibilidad.
    Retorna JPEG comprimido para reducir tamaño de envío a Gemini.
    DPI 200 es suficiente para Gemini 2.5 Flash (más rápido).
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if pagina_idx < len(doc):
        page = doc[pagina_idx]
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)

        # Convertir pixmap a imagen Pillow
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Aplicar preprocesamiento para mejorar legibilidad
        img = _mejorar_imagen(img)

        # Comprimir como JPEG (mucho más pequeño que PNG)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85, optimize=True)
        return buf.getvalue()
    return b""


def dividir_imagen_en_zonas(imagen_bytes: bytes) -> dict[str, bytes]:
    """
    Divide la imagen del folio en zonas para enviar a Gemini por separado.
    Aprovecha que todos los folios tienen la misma estructura.

    Retorna dict con:
    - 'campos': zona superior con datos del proveedor, factura, fechas
    - 'tabla': zona media con las partidas/artículos
    """
    img = Image.open(io.BytesIO(imagen_bytes))
    ancho, alto = img.size

    # Zona 1: Campos superiores (0% a 55% de la altura)
    # Incluye: Razón Social, Sucursal, Factura, Folio Puerta, Fechas, Acuse
    campos = img.crop((0, 0, ancho, int(alto * 0.55)))
    buf_campos = io.BytesIO()
    campos.save(buf_campos, format="JPEG", quality=85, optimize=True)

    # Zona 2: Tabla de partidas (40% a 95% de la altura)
    # Solapamiento del 15% para no perder datos entre zonas
    tabla = img.crop((0, int(alto * 0.40), ancho, int(alto * 0.95)))
    buf_tabla = io.BytesIO()
    tabla.save(buf_tabla, format="JPEG", quality=85, optimize=True)

    return {
        "campos": buf_campos.getvalue(),
        "tabla": buf_tabla.getvalue(),
    }


def obtener_imagen_pagina(pdf_bytes: bytes, pagina_idx: int) -> bytes:
    """
    Obtiene la representación de una página como imagen.
    Primero intenta texto nativo; si está vacío, convierte a imagen.
    Siempre retorna bytes de imagen para enviar a OpenAI Vision.
    """
    return convertir_pagina_a_imagen(pdf_bytes, pagina_idx)


def dividir_pdf_por_paginas(
    pdf_bytes: bytes, paginas: list[int]
) -> bytes:
    """
    Divide el PDF extrayendo solo las páginas indicadas (0-indexed).
    Retorna los bytes del micro-PDF resultante.
    """
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for idx in paginas:
        if idx < len(reader.pages):
            writer.add_page(reader.pages[idx])
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def dividir_pdf_en_micro_pdfs(
    pdf_bytes: bytes, grupos: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """
    Divide el PDF en micro-PDFs según grupos de páginas contiguas.

    Args:
        pdf_bytes: Bytes del PDF original.
        grupos: Lista de dicts con claves:
            - 'paginas': lista de índices (0-indexed)
            - 'proveedor': nombre del proveedor
            - 'numero_factura_folio': número de factura/folio
            - 'fecha': fecha del folio

    Returns:
        Lista de dicts con los mismos datos + 'pdf_bytes' y 'nombre_archivo'.
    """
    resultados = []
    for grupo in grupos:
        micro_pdf = dividir_pdf_por_paginas(pdf_bytes, grupo["paginas"])
        proveedor_safe = _sanitizar_nombre(grupo.get("proveedor", "Desconocido"))
        folio_safe = _sanitizar_nombre(
            str(grupo.get("numero_factura_folio", "SIN_FOLIO"))
        )
        nombre_archivo = f"Folio_{folio_safe}_{proveedor_safe}.pdf"
        resultados.append(
            {
                **grupo,
                "pdf_bytes": micro_pdf,
                "nombre_archivo": nombre_archivo,
            }
        )
    return resultados


def _sanitizar_nombre(nombre: str) -> str:
    """Sanitiza un string para usarlo como nombre de archivo."""
    caracteres_invalidos = '<>:"/\\|?*'
    for char in caracteres_invalidos:
        nombre = nombre.replace(char, "_")
    return nombre.strip()[:100]


def asegurar_directorio(path: str) -> None:
    """Crea un directorio si no existe."""
    os.makedirs(path, exist_ok=True)
