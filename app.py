"""
app.py
Aplicación principal de Streamlit para automatización de folios logísticos.
Interfaz en español para procesamiento, extracción, división, validación
y envío de correos de folios de recibo.
"""

import os
import io
import time
import tempfile
from datetime import datetime
from typing import Any

import streamlit as st
import pandas as pd

APP_VERSION = "v2.5.0"
APP_BUILD_DATE = "2026-06-25"

import config
from utils.pdf_processor import (
    contar_paginas,
    obtener_imagen_pagina,
    dividir_pdf_en_micro_pdfs,
    asegurar_directorio,
)
from utils.openai_client import extraer_datos_folio
from utils.google_services import subir_micro_pdf, es_simulacion as gdrive_simulacion
from utils.email_sender import (
    cargar_directorio_proveedores,
    buscar_email_proveedor,
    generar_excel_proveedor,
    generar_reporte_detallado,
    generar_resumen_totales,
    enviar_correo_proveedor,
)


# ──────────────────────────────────────────────
# Configuración de la página
# ──────────────────────────────────────────────

st.set_page_config(
    page_title="Folios Logísticos - City y Fresko",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ──────────────────────────────────────────────
# Inicialización de estado de sesión
# ──────────────────────────────────────────────

def init_session_state():
    """Inicializa variables de estado de sesión."""
    if "procesando" not in st.session_state:
        st.session_state.procesando = False
    if "resultados" not in st.session_state:
        st.session_state.resultados = []
    if "micro_pdfs" not in st.session_state:
        st.session_state.micro_pdfs = []
    if "excel_maestro" not in st.session_state:
        st.session_state.excel_maestro = None
    if "errores" not in st.session_state:
        st.session_state.errores = 0
    if "exitos" not in st.session_state:
        st.session_state.exitos = 0


init_session_state()


# ──────────────────────────────────────────────
# Cargar directorio de proveedores al inicio
# ──────────────────────────────────────────────

@st.cache_data
def cargar_directorio():
    """Carga el directorio de proveedores (cacheado)."""
    return cargar_directorio_proveedores()


df_directorio = cargar_directorio()


# ──────────────────────────────────────────────
# Sidebar
# ──────────────────────────────────────────────

with st.sidebar:
    st.markdown("## ⚙️ Configuración")

    # Badge de versión
    st.markdown(
        f"`{APP_VERSION}` | Build: {APP_BUILD_DATE} | "
        f"Actualizado: {datetime.now().strftime('%H:%M')}"
    )

    modo_sim = config.is_simulation_mode()
    if modo_sim:
        st.warning(
            "🔴 **Modo Simulación Activo**\n\n"
            "No se detectaron credenciales en `.env`. "
            "La aplicación funcionará con datos de prueba.\n\n"
            "Configura `.env` con:\n"
            "- `OPENAI_API_KEY`\n"
            "- `GOOGLE_SERVICE_ACCOUNT_FILE`\n"
            "- `SMTP_USER` / `SMTP_PASSWORD`"
        )
    else:
        st.success("🟢 Credenciales detectadas. Modo producción activo.")

    st.markdown("---")

    # Estado del directorio de proveedores
    st.markdown("### 📇 Directorio de Proveedores")
    if df_directorio is not None:
        st.success(
            f"✅ `directorio_proveedores.xlsx` cargado.\n\n"
            f"**{len(df_directorio)}** proveedores encontrados."
        )
        with st.expander("Ver directorio"):
            st.dataframe(df_directorio, use_container_width=True)
    else:
        st.warning(
            "⚠️ No se encontró `directorio_proveedores.xlsx`.\n\n"
            "El envío automático de correos está deshabilitado.\n"
            "Coloca el archivo en el directorio raíz para activarlo."
        )

    st.markdown("---")
    st.markdown("### ℹ️ Ayuda")
    with st.expander("Instrucciones de uso"):
        st.markdown(
            """
            1. **Opcional**: Carga el Excel Maestro del Día (control matutino).
            2. **Obligatorio**: Carga el PDF con los folios escaneados.
            3. Click en **Procesar Folios**.
            4. Descarga los reportes generados.
            5. Click en **✉ Enviar Correo** por cada proveedor.
            """
        )

    st.caption("Motor de PDF: PyMuPDF (sin dependencias externas)")


# ──────────────────────────────────────────────
# Título principal
# ──────────────────────────────────────────────

st.title("📦 Folios Logísticos - Operación City y Fresko")
st.markdown(
    "Automatización de procesamiento, extracción, división, "
    "validación y envío de correos de folios de recibo."
)
st.markdown("---")


# ──────────────────────────────────────────────
# Carga de archivos
# ──────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📊 Excel Maestro del Día (Opcional)")
    excel_file = st.file_uploader(
        "Cargar Excel Maestro del Día",
        type=["xlsx", "xls"],
        key="excel_maestro_upload",
        help="Archivo Excel de control matutino. Si no se carga, "
        "los folios se marcarán como 'Pendiente de Validación'.",
    )

    if excel_file is not None:
        size_mb = len(excel_file.getvalue()) / (1024 * 1024)
        if size_mb > config.MAX_EXCEL_SIZE_MB:
            st.error(
                f"El archivo excede el tamaño máximo de {config.MAX_EXCEL_SIZE_MB} MB."
            )
        else:
            st.success(f"✅ Excel cargado: {excel_file.name} ({size_mb:.1f} MB)")
            try:
                st.session_state.excel_maestro = pd.read_excel(excel_file)
            except Exception as e:
                st.error(f"Error al leer el Excel: {e}")
                st.session_state.excel_maestro = None
    else:
        st.session_state.excel_maestro = None
        st.info("Sin cargar - Los folios se marcarán como pendientes.")

with col2:
    st.markdown("### 📄 PDF con Folios Escaneados (Obligatorio)")
    pdf_file = st.file_uploader(
        "Cargar PDF con Folios Escaneados",
        type=["pdf"],
        key="pdf_folios_upload",
        help="PDF multi-página que contiene los folios de recibo escaneados.",
    )

    if pdf_file is not None:
        size_mb = len(pdf_file.getvalue()) / (1024 * 1024)
        if size_mb > config.MAX_PDF_SIZE_MB:
            st.error(
                f"El archivo excede el tamaño máximo de {config.MAX_PDF_SIZE_MB} MB."
            )
        else:
            st.success(f"✅ PDF cargado: {pdf_file.name} ({size_mb:.1f} MB)")
            try:
                num_paginas = contar_paginas(pdf_file.getvalue())
                st.info(f"📄 **{num_paginas}** páginas detectadas.")
            except Exception as e:
                st.error(f"Error al leer el PDF: {e}")


# ──────────────────────────────────────────────
# Validación contra Excel maestro
# ──────────────────────────────────────────────

def validar_con_excel_maestro(
    datos_extraidos: list[dict[str, Any]],
    df_maestro: pd.DataFrame,
) -> list[dict[str, Any]]:
    """
    Valida las cajas extraídas contra el Excel maestro.
    Marca cada item con 'Validado' o 'Discrepancia'.
    """
    df_maestro_norm = df_maestro.copy()
    df_maestro_norm.columns = df_maestro_norm.columns.str.strip().str.lower()

    # Identificar columnas del Excel maestro
    col_factura = None
    col_cajas = None
    for col in df_maestro_norm.columns:
        if "factura" in col or "folio" in col or "remision" in col:
            col_factura = col
        if "cajas" in col:
            col_cajas = col

    for item in datos_extraidos:
        factura = str(item.get("numero_factura_folio", "")).strip()
        cajas_extraidas = sum(p.get("cajas", 0) for p in item.get("partidas", []))

        if col_factura and col_cajas and factura:
            fila = df_maestro_norm[
                df_maestro_norm[col_factura].astype(str).str.strip() == factura
            ]
            if not fila.empty:
                cajas_esperadas = int(fila.iloc[0][col_cajas])
                if cajas_extraidas == cajas_esperadas:
                    item["estatus_validacion"] = "Validado"
                else:
                    item["estatus_validacion"] = (
                        f"Discrepancia (Esperado: {cajas_esperadas}, "
                        f"Extraído: {cajas_extraidas})"
                    )
            else:
                item["estatus_validacion"] = "No encontrado en Excel maestro"
        else:
            item["estatus_validacion"] = "Pendiente de Validación"

    return datos_extraidos


# ──────────────────────────────────────────────
# Agrupar páginas contiguas por proveedor/folio
# ──────────────────────────────────────────────

def agrupar_paginas(
    datos_por_pagina: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Agrupa páginas contiguas que comparten el mismo proveedor y folio.
    """
    grupos = []
    grupo_actual: dict[str, Any] | None = None

    for idx, datos in enumerate(datos_por_pagina):
        proveedor = datos.get("proveedor") or "Desconocido"
        folio = datos.get("numero_factura_folio") or f"SIN_FOLIO_{idx}"

        if (
            grupo_actual
            and grupo_actual["proveedor"] == proveedor
            and grupo_actual["numero_factura_folio"] == folio
        ):
            grupo_actual["paginas"].append(idx)
            # Acumular partidas
            grupo_actual["partidas"].extend(datos.get("partidas", []))
        else:
            if grupo_actual:
                grupos.append(grupo_actual)
            grupo_actual = {
                "proveedor": proveedor,
                "numero_factura_folio": folio,
                "fecha": datos.get("fecha", ""),
                "sucursal_receptora": datos.get("sucursal_receptora"),
                "folio_puerta": datos.get("folio_puerta"),
                "tipo_documento": datos.get("tipo_documento", "No identificado"),
                "acuse_de_recibo": datos.get("acuse_de_recibo"),
                "paginas": [idx],
                "partidas": list(datos.get("partidas", [])),
            }

    if grupo_actual:
        grupos.append(grupo_actual)

    return grupos


# ──────────────────────────────────────────────
# Procesamiento principal
# ──────────────────────────────────────────────

st.markdown("---")

if st.button(
    "🚀 Procesar Folios",
    disabled=(pdf_file is None),
    use_container_width=True,
    type="primary",
):
    if pdf_file is None:
        st.error("Debe cargar un PDF con folios escaneados.")
    else:
        st.session_state.procesando = True
        st.session_state.resultados = []
        st.session_state.micro_pdfs = []
        st.session_state.errores = 0
        st.session_state.exitos = 0

        pdf_bytes = pdf_file.getvalue()
        total_paginas = contar_paginas(pdf_bytes)

        # ── Barra de progreso y métricas ──
        st.markdown("### 📊 Panel de Procesamiento")
        progress_bar = st.progress(0.0, text="Iniciando procesamiento...")

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        metric_exitos = col_m1.empty()
        metric_errores = col_m2.empty()
        metric_paginas = col_m3.empty()
        metric_proveedores = col_m4.empty()

        metric_exitos.metric("✅ Éxitos", 0)
        metric_errores.metric("❌ Errores", 0)
        metric_paginas.metric("📄 Páginas", total_paginas)
        metric_proveedores.metric("🏢 Proveedores", 0)

        # ── Paso 1: Extraer datos de cada página (paralelo) ──
        st.markdown("#### Paso 1: Extracción de datos con IA (procesamiento paralelo)")
        datos_por_pagina = [None] * total_paginas

        import concurrent.futures

        def procesar_pagina(idx):
            """Procesa una página y retorna (idx, datos)."""
            try:
                imagen = obtener_imagen_pagina(pdf_bytes, idx)
                datos = extraer_datos_folio(imagen)
                return (idx, datos)
            except Exception as e:
                return (idx, {"_error": str(e)})

        MAX_HILOS = 4
        completadas = 0

        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_HILOS) as executor:
            futures = {
                executor.submit(procesar_pagina, i): i
                for i in range(total_paginas)
            }

            for future in concurrent.futures.as_completed(futures):
                idx, datos = future.result()
                datos_por_pagina[idx] = datos
                completadas += 1

                progress_bar.progress(
                    completadas / (total_paginas * 2),
                    text=f"Extrayendo datos de página {completadas} de {total_paginas}...",
                )

                if datos.get("_error"):
                    st.session_state.errores += 1
                    st.warning(
                        f"Página {idx + 1}: Error de extracción - {datos['_error']}"
                    )
                else:
                    st.session_state.exitos += 1

                metric_exitos.metric("✅ Éxitos", st.session_state.exitos)
                metric_errores.metric("❌ Errores", st.session_state.errores)

        # ── Paso 2: Agrupar páginas contiguas ──
        st.markdown("#### Paso 2: Agrupación de folios por proveedor")
        grupos = agrupar_paginas(datos_por_pagina)
        metric_proveedores.metric("🏢 Proveedores", len(grupos))

        # ── Paso 3: Dividir PDF en micro-PDFs ──
        st.markdown("#### Paso 3: División de PDF en micro-PDFs")
        micro_pdfs = dividir_pdf_en_micro_pdfs(pdf_bytes, grupos)
        st.session_state.micro_pdfs = micro_pdfs

        # ── Paso 4: Subir a Google Drive ──
        st.markdown("#### Paso 4: Subida a Google Drive")
        if gdrive_simulacion():
            st.info("🔌 Google Drive en modo simulación - Los PDFs no se subirán.")

        resultados_finales = []
        for j, micro in enumerate(micro_pdfs):
            progress_bar.progress(
                0.5 + (j + 1) / (len(micro_pdfs) * 2),
                text=f"Subiendo folio {j + 1} de {len(micro_pdfs)} a Google Drive...",
            )

            proveedor = micro.get("proveedor", "Desconocido")
            fecha = micro.get("fecha", "SIN_FECHA")
            nombre_arch = micro.get("nombre_archivo", "folio.pdf")

            resultado_upload = subir_micro_pdf(
                micro["pdf_bytes"],
                nombre_arch,
                proveedor,
                fecha,
            )

            # Construir resultado final
            resultado = {
                "proveedor": proveedor,
                "sucursal_receptora": micro.get("sucursal_receptora"),
                "numero_factura_folio": micro.get("numero_factura_folio"),
                "folio_puerta": micro.get("folio_puerta"),
                "tipo_documento": micro.get("tipo_documento", "No identificado"),
                "acuse_de_recibo": micro.get("acuse_de_recibo"),
                "fecha": fecha,
                "partidas": micro.get("partidas", []),
                "estatus_validacion": "Pendiente de Validación",
                "drive_url": resultado_upload.get("url", ""),
                "drive_success": resultado_upload.get("success", False),
                "nombre_archivo": nombre_arch,
                "pdf_bytes": micro["pdf_bytes"],
            }
            resultados_finales.append(resultado)

        # ── Paso 5: Validación con Excel maestro ──
        if st.session_state.excel_maestro is not None:
            st.markdown("#### Paso 5: Validación con Excel Maestro")
            resultados_finales = validar_con_excel_maestro(
                resultados_finales, st.session_state.excel_maestro
            )
        else:
            st.info(
                "ℹ️ No se cargó Excel maestro. "
                "Todos los folios marcados como 'Pendiente de Validación'."
            )

        st.session_state.resultados = resultados_finales
        progress_bar.progress(1.0, text="¡Procesamiento completado!")
        st.success("✅ Procesamiento completado correctamente.")
        st.session_state.procesando = False


# ──────────────────────────────────────────────
# Tabla de resultados interactiva
# ──────────────────────────────────────────────

if st.session_state.resultados:
    st.markdown("---")
    st.markdown("## 📋 Resultados de Procesamiento")

    # Separar Acuse de Recibo de documentos no identificados
    resultados_acuse = []
    resultados_no_identificados = []
    for r in st.session_state.resultados:
        tipo = r.get("tipo_documento", "No identificado")
        if "acuse" in tipo.lower() and "recibo" in tipo.lower():
            resultados_acuse.append(r)
        else:
            resultados_no_identificados.append(r)

    # ── Sección 1: Acuses de Recibo ──
    st.markdown(f"### ✅ Acuses de Recibo ({len(resultados_acuse)})")

    if resultados_acuse:
        filas_resumen = []
        for r in resultados_acuse:
            cajas_totales = sum(p.get("cajas", 0) for p in r.get("partidas", []))
            filas_resumen.append(
                {
                    "Proveedor": r.get("proveedor", ""),
                    "Fecha": r.get("fecha", ""),
                    "Factura/Folio": r.get("numero_factura_folio", ""),
                    "Folio Puerta": r.get("folio_puerta", ""),
                    "Sucursal": r.get("sucursal_receptora", ""),
                    "Cajas Totales": cajas_totales,
                    "Partidas": len(r.get("partidas", [])),
                    "Estatus": r.get("estatus_validacion", ""),
                    "Drive": "✅" if r.get("drive_success") else "❌",
                }
            )

        df_resumen = pd.DataFrame(filas_resumen)
        st.dataframe(df_resumen, use_container_width=True, hide_index=True)
    else:
        st.info("No se encontraron Acuses de Recibo.")

    # ── Sección 2: Documentos no identificados ──
    if resultados_no_identificados:
        st.markdown(f"### ⚠️ Documentos No Identificados ({len(resultados_no_identificados)})")
        filas_no_id = []
        for r in resultados_no_identificados:
            cajas_totales = sum(p.get("cajas", 0) for p in r.get("partidas", []))
            filas_no_id.append(
                {
                    "Tipo Detectado": r.get("tipo_documento", "No identificado"),
                    "Proveedor": r.get("proveedor", ""),
                    "Fecha": r.get("fecha", ""),
                    "Cajas": cajas_totales,
                    "Drive": "✅" if r.get("drive_success") else "❌",
                }
            )
        df_no_id = pd.DataFrame(filas_no_id)
        st.dataframe(df_no_id, use_container_width=True, hide_index=True)
        st.caption("Estos documentos no fueron identificados como 'Acuse de Recibo'.")

    # ── Botones de descarga global ──
    st.markdown("### 📥 Descargar Reportes")

    col_d1, col_d2 = st.columns(2)

    with col_d1:
        reporte_detallado_bytes = generar_reporte_detallado(
            st.session_state.resultados
        )
        st.download_button(
            label="📊 Descargar Reporte Detallado (.xlsx)",
            data=reporte_detallado_bytes,
            file_name="reporte_detallado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_d2:
        resumen_totales_bytes = generar_resumen_totales(
            st.session_state.resultados
        )
        st.download_button(
            label="📈 Descargar Resumen de Totales (.xlsx)",
            data=resumen_totales_bytes,
            file_name="resumen_totales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ── Tabla detallada con botón de envío por proveedor ──
    st.markdown("---")
    st.markdown("## ✉ Envío de Correos por Proveedor")

    # Botón para enviar todos los correos (solo Acuses de Recibo)
    df_directorio_global = cargar_directorio_proveedores()
    proveedores_con_email = []
    for r in resultados_acuse:
        proveedor = r.get("proveedor", "Desconocido")
        if df_directorio_global is not None:
            email = buscar_email_proveedor(proveedor, df_directorio_global)
            if email:
                proveedores_con_email.append((r, email))

    if proveedores_con_email:
        if st.button(
            f"📧 Enviar TODOS los correos ({len(proveedores_con_email)} proveedores)",
            key="btn_send_all",
            use_container_width=True,
            type="primary",
        ):
            enviados = 0
            errores = 0
            with st.spinner(f"Enviando {len(proveedores_con_email)} correos..."):
                for r, email_destino in proveedores_con_email:
                    proveedor = r.get("proveedor", "Desconocido")
                    fecha = r.get("fecha", "")
                    folio = r.get("numero_factura_folio", "")

                    datos_proveedor = []
                    for p in r.get("partidas", []):
                        datos_proveedor.append(
                            {
                                "fecha": fecha,
                                "numero_factura_folio": folio,
                                "cajas": p.get("cajas", 0),
                                "folio_puerta": r.get("folio_puerta", ""),
                                "sucursal_receptora": r.get("sucursal_receptora", ""),
                            }
                        )

                    excel_bytes = generar_excel_proveedor(
                        datos_proveedor, proveedor
                    )
                    excel_nombre = f"Resumen_{proveedor}_{folio}.xlsx"
                    pdf_bytes = r.get("pdf_bytes")
                    pdf_nombre = r.get("nombre_archivo", "folio.pdf")

                    resultado_email = enviar_correo_proveedor(
                        email_destino=email_destino,
                        proveedor=proveedor,
                        fecha_folio=fecha,
                        pdf_bytes=pdf_bytes,
                        pdf_nombre=pdf_nombre,
                        excel_bytes=excel_bytes,
                        excel_nombre=excel_nombre,
                    )

                    if resultado_email.get("success"):
                        enviados += 1
                    else:
                        errores += 1
                        st.warning(
                            f"Error {proveedor}: {resultado_email.get('error', 'Desconocido')}"
                        )

            if enviados > 0:
                st.success(f"✅ {enviados} correo(s) enviado(s) correctamente.")
            if errores > 0:
                st.error(f"❌ {errores} correo(s) con error.")
            if errores == 0 and enviados > 0:
                st.balloons()

    for idx, r in enumerate(resultados_acuse):
        proveedor = r.get("proveedor", "Desconocido")
        fecha = r.get("fecha", "")
        folio = r.get("numero_factura_folio", "")
        cajas = sum(p.get("cajas", 0) for p in r.get("partidas", []))

        col_p1, col_p2, col_p3, col_p4, col_p5 = st.columns(
            [4, 2, 1.5, 1, 1]
        )

        with col_p1:
            st.markdown(f"**{proveedor}**")

        with col_p2:
            st.caption(f"Fact: {folio} | {fecha}")

        with col_p3:
            st.caption(f"Cajas: {cajas}")

        with col_p4:
            pdf_bytes_prov = r.get("pdf_bytes")
            pdf_nombre_prov = r.get("nombre_archivo", "folio.pdf")
            if pdf_bytes_prov:
                st.download_button(
                    "📥",
                    data=pdf_bytes_prov,
                    file_name=pdf_nombre_prov,
                    mime="application/pdf",
                    key=f"dl_pdf_{idx}",
                )

        with col_p5:
            email_destino = None
            if df_directorio is not None:
                email_destino = buscar_email_proveedor(
                    proveedor, df_directorio
                )

            if email_destino:
                if st.button(
                    "✉",
                    key=f"btn_email_{idx}",
                ):
                    with st.spinner(f"Enviando correo a {email_destino}..."):
                        datos_proveedor = []
                        for p in r.get("partidas", []):
                            datos_proveedor.append(
                                {
                                    "fecha": fecha,
                                    "numero_factura_folio": folio,
                                    "cajas": p.get("cajas", 0),
                                    "folio_puerta": r.get("folio_puerta", ""),
                                    "sucursal_receptora": r.get("sucursal_receptora", ""),
                                }
                            )

                        excel_bytes = generar_excel_proveedor(
                            datos_proveedor, proveedor
                        )
                        excel_nombre = f"Resumen_{proveedor}_{folio}.xlsx"
                        pdf_bytes = r.get("pdf_bytes")
                        pdf_nombre = r.get("nombre_archivo", "folio.pdf")

                        resultado_email = enviar_correo_proveedor(
                            email_destino=email_destino,
                            proveedor=proveedor,
                            fecha_folio=fecha,
                            pdf_bytes=pdf_bytes,
                            pdf_nombre=pdf_nombre,
                            excel_bytes=excel_bytes,
                            excel_nombre=excel_nombre,
                        )

                        if resultado_email.get("success"):
                            if resultado_email.get("simulacion"):
                                st.toast(f"📧 Simulado a {email_destino}", icon="📧")
                            else:
                                st.toast(f"✅ Enviado a {email_destino}", icon="✅")
                        else:
                            st.error(f"Error: {resultado_email.get('error', 'Desconocido')}")
            else:
                st.caption("🚫 Sin email", help="No se encontró email en el directorio")


# ──────────────────────────────────────────────
# Footer
# ──────────────────────────────────────────────

st.markdown("---")
st.caption(
    f"Sistema de Folios Logísticos - Operación City y Fresko | "
    f"Modelo: {config.GEMINI_MODEL if config.EXTRACTION_MODEL == 'gemini' else config.OPENAI_MODEL} | "
    f"{'Modo Simulación' if config.is_simulation_mode() else 'Modo Producción'} | "
    f"`{APP_VERSION}` ({APP_BUILD_DATE})"
)
