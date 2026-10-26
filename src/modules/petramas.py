# ====================================================================
# --- MÓDULO STREAMLIT: GESTIÓN DE EXPEDIENTES DE DISPOSICIÓN FINAL ---
# ====================================================================
import streamlit as st
import pandas as pd
import io
import os
import gc
import json
from datetime import datetime
from typing import Dict, Any, List

from src.config.settings import (
    CARPETAS_PETRAMAS,
    EMPRESAS_EORS,
    SHEET_CERTIFICADOS_PETRAMAS,
    SHEET_SUSTENTOS_CAMPO
)
from src.services.petramas_service import (
    detectar_tipo_documento,
    analizar_expediente_completo_inversa,
    desglosar_y_registrar_expediente_inversa,
    analizar_constancia_petramas,
    guardar_constancia_petramas_en_sheet,
    analizar_sustento_campo,
    guardar_sustento_campo_en_sheet,
    ensamblar_expediente_disposicion_final,
    obtener_resumen_sustentos_pendientes,
    obtener_constancias_petramas_registradas,
    normalizar_empresa_eors,
    subir_pdf_petramas_drive
)

def render_petramas():
    st.title("📑 Gestión de Expedientes de Disposición Final (Petramás)")
    st.caption("Control integral de constancias Petramás, escaneos de campo, ingeniería inversa de expedientes y ensamblado secuencial en 5 partes.")

    # Manejo de query params para Endpoint REST: /api/petramas/pendientes
    params = st.query_params
    if params.get("endpoint") == "api/petramas/pendientes" or params.get("api") == "petramas_pendientes":
        empresa_param = params.get("empresa")
        data_api = obtener_resumen_sustentos_pendientes(empresa=empresa_param)
        st.subheader("🌐 Endpoint API: Sustentos Pendientes / Completos")
        st.json(data_api)
        st.stop()

    tab_inversa, tab_sustentos, tab_ensamblador, tab_constancias = st.tabs([
        "🔄 Desglose por Ingeniería Inversa",
        "📊 Control de Sustentos (Pendientes / Completos)",
        "📑 Ensamblador de Expedientes (5 Partes)",
        "🎫 Constancias Petramás"
    ])

    usuario_actual = st.session_state.get('usuario_nombre', 'Operador')

    # ====================================================================
    # --- TAB 1: DESGLOSE POR INGENIERÍA INVERSA ---
    # ====================================================================
    with tab_inversa:
        st.markdown("### 1. Ingesta de Expediente Completo en PDF")
        st.markdown("""
        Carga un expediente final completo (ej: 20 páginas). El motor cognitivo **Google Gemini Multimodal** 
        analizará la estructura visual y de texto para desglosar:
        1. **Certificado Emitido** (Carátula de la EO-RS)
        2. **Constancia de Petramás** (Relleno Huaycoloro)
        3. **Sustentos de Campo** (Manifiestos MINAM + Boletas de balanza + Guías de remisión)
        """)

        col_up1, col_up2 = st.columns([3, 1])
        with col_up1:
            exp_file = st.file_uploader(
                "Selecciona o arrastra el expediente PDF completo:",
                type=["pdf"],
                key="uploader_exp_inversa"
            )
        with col_up2:
            st.write("")
            st.write("")
            btn_limpiar_inv = st.button("🧹 Limpiar Análisis", use_container_width=True, key="btn_clear_inversa")

        if btn_limpiar_inv:
            for k in ['inversa_datos_extraidos', 'inversa_pdf_bytes', 'inversa_resultado_desglose']:
                st.session_state.pop(k, None)
            st.rerun()

        if exp_file is not None and 'inversa_pdf_bytes' not in st.session_state:
            st.session_state['inversa_pdf_bytes'] = exp_file.getvalue()
            st.session_state.pop('inversa_datos_extraidos', None)
            st.session_state.pop('inversa_resultado_desglose', None)

        pdf_bytes = st.session_state.get('inversa_pdf_bytes')

        if pdf_bytes:
            # Inspección inicial de tipo
            tipo_detectado = detectar_tipo_documento(pdf_bytes)
            col_info_doc1, col_info_doc2 = st.columns([2, 2])
            with col_info_doc1:
                st.info(f"📄 **Tipo detectado:** `{tipo_detectado}` (Tamaño: {len(pdf_bytes) / 1024:.1f} KB)")
            with col_info_doc2:
                btn_analizar_inv = st.button(
                    "🧠 Analizar Estructura con IA Gemini",
                    type="primary",
                    use_container_width=True,
                    key="btn_analizar_gemini_inversa"
                )

            if btn_analizar_inv or 'inversa_datos_extraidos' in st.session_state:
                if 'inversa_datos_extraidos' not in st.session_state:
                    with st.spinner("Analizando páginas y segmentando secciones del expediente con Gemini..."):
                        try:
                            datos_inv = analizar_expediente_completo_inversa(pdf_bytes)
                            st.session_state['inversa_datos_extraidos'] = datos_inv
                            st.toast("✅ Expediente analizado con éxito.")
                        except Exception as e_inv:
                            st.error(f"❌ Error al analizar expediente con IA: {e_inv}")
                            st.stop()

                datos_inv = st.session_state.get('inversa_datos_extraidos')

                if datos_inv:
                    st.divider()
                    st.markdown("### 2. Revisión y Corrección de Datos Extraídos")
                    st.caption("Verifica o ajusta los valores antes de confirmar el despiece y la publicación en Google Drive y Sheets.")

                    c_form1, c_form2, c_form3 = st.columns(3)
                    with c_form1:
                        opciones_empresas = list(EMPRESAS_EORS.keys())
                        emp_detectada = datos_inv.get("empresa_eors", "EPMI S.A.C.")
                        idx_emp = opciones_empresas.index(emp_detectada) if emp_detectada in opciones_empresas else 0
                        edit_empresa = st.selectbox("Empresa EO-RS:", opciones_empresas, index=idx_emp, key="edit_inv_empresa")
                        edit_correlativo = st.text_input("N° Correlativo Certificado:", value=str(datos_inv.get("correlativo_cert", "")), key="edit_inv_corr")
                        edit_num_cert = st.text_input("Número Certificado Completo:", value=str(datos_inv.get("numero_cert_completo", "")), key="edit_inv_num_cert")

                    with c_form2:
                        edit_cliente = st.text_input("Cliente Generador:", value=str(datos_inv.get("cliente_generador", "")), key="edit_inv_cliente")
                        edit_sede = st.text_input("Sede / Fundo:", value=str(datos_inv.get("sede", "")), key="edit_inv_sede")
                        edit_fecha = st.text_input("Fecha Servicio (DD/MM/YYYY):", value=str(datos_inv.get("fecha_servicio", "")), key="edit_inv_fecha")

                    with c_form3:
                        edit_placa = st.text_input("Placa Vehículo:", value=str(datos_inv.get("placa", "")), key="edit_inv_placa")
                        edit_guia_rem = st.text_input("Guía Remisión Remitente:", value=str(datos_inv.get("guia_remision", "")), key="edit_inv_guia_rem")
                        edit_guia_trans = st.text_input("Guía Remisión Transportista:", value=str(datos_inv.get("guia_transporte", "")), key="edit_inv_guia_trans")

                    st.markdown("#### 📑 Distribución de Páginas por Sección")
                    col_p1, col_p2, col_p3 = st.columns(3)

                    pet_info = datos_inv.get("petramas", {})
                    sus_info = datos_inv.get("sustentos", {})

                    with col_p1:
                        st.markdown("**🎫 Constancia Petramás**")
                        edit_pet_existe = st.checkbox("Incluye Petramás", value=bool(pet_info.get("existe", True)), key="chk_inv_pet_existe")
                        edit_pet_const = st.text_input("N° Constancia Petramás:", value=str(pet_info.get("numero_constancia", "")), key="edit_inv_pet_const")
                        pags_pet_str = ",".join(str(p) for p in pet_info.get("paginas", []))
                        edit_pet_pags = st.text_input("Páginas Petramás (ej: 2,3):", value=pags_pet_str, key="edit_inv_pet_pags")
                        edit_pet_peso = st.number_input("Peso Total (kg):", value=float(pet_info.get("peso_total_kg", 0.0)), step=10.0, key="edit_inv_pet_peso")

                    with col_p2:
                        st.markdown("**📄 Manifiestos de Manejo**")
                        pags_man_str = ",".join(str(p) for p in sus_info.get("paginas_manifiestos", []))
                        edit_man_pags = st.text_input("Páginas Manifiestos (ej: 4,5,6,7):", value=pags_man_str, key="edit_inv_man_pags")

                        st.markdown("**⚖️ Boletas de Pesaje (Balanza)**")
                        pags_bol_str = ",".join(str(p) for p in sus_info.get("paginas_boletas", []))
                        edit_bol_pags = st.text_input("Páginas Boletas (ej: 8,9,10):", value=pags_bol_str, key="edit_inv_bol_pags")

                    with col_p3:
                        st.markdown("**🚛 Guías de Remisión / Transporte**")
                        pags_gui_str = ",".join(str(p) for p in sus_info.get("paginas_guias", []))
                        edit_gui_pags = st.text_input("Páginas Guías (ej: 11,12):", value=pags_gui_str, key="edit_inv_gui_pags")

                        boletas_detectadas = sus_info.get("boletas_pesaje", pet_info.get("boletas", []))
                        boletas_txt = ", ".join(str(b) for b in boletas_detectadas)
                        edit_boletas_pesaje = st.text_area("Números de Boletas de Pesaje (Tickets):", value=boletas_txt, key="edit_inv_boletas_txt")

                    st.divider()
                    st.markdown("### 3. Ejecutar Despiece y Sincronización")

                    def _parse_lista_pags(txt: str) -> List[int]:
                        res = []
                        for item in txt.replace(" ", "").split(","):
                            if item.isdigit():
                                res.append(int(item))
                        return res

                    btn_desglosar = st.button(
                        "🚀 Desglosar PDF, Subir a Drive y Registrar en Google Sheets",
                        type="primary",
                        use_container_width=True,
                        key="btn_ejecutar_desglose_inversa"
                    )

                    if btn_desglosar:
                        with st.spinner("Desglosando sub-PDFs con PyPDF, subiendo a Google Drive y sincronizando con Google Sheets..."):
                            try:
                                # Reconstruir payload validado
                                p_pet = _parse_lista_pags(edit_pet_pags)
                                p_man = _parse_lista_pags(edit_man_pags)
                                p_bol = _parse_lista_pags(edit_bol_pags)
                                p_gui = _parse_lista_pags(edit_gui_pags)
                                bol_list = [b.strip() for b in edit_boletas_pesaje.replace("\n", ",").split(",") if b.strip()]

                                payload_actualizado = {
                                    "empresa_eors": edit_empresa,
                                    "correlativo_cert": edit_correlativo,
                                    "numero_cert_completo": edit_num_cert,
                                    "cliente_generador": edit_cliente,
                                    "sede": edit_sede,
                                    "fecha_servicio": edit_fecha,
                                    "placa": edit_placa,
                                    "guia_remision": edit_guia_rem,
                                    "guia_transporte": edit_guia_trans,
                                    "petramas": {
                                        "existe": edit_pet_existe,
                                        "numero_constancia": edit_pet_const,
                                        "paginas": p_pet,
                                        "peso_total_kg": edit_pet_peso,
                                        "boletas": bol_list
                                    },
                                    "sustentos": {
                                        "paginas_manifiestos": p_man,
                                        "paginas_boletas": p_bol,
                                        "paginas_guias": p_gui,
                                        "boletas_pesaje": bol_list,
                                        "rangos_texto": f"M:{len(p_man)}p | B:{len(p_bol)}p | G:{len(p_gui)}p"
                                    }
                                }

                                resultado = desglosar_y_registrar_expediente_inversa(
                                    pdf_bytes,
                                    payload_actualizado,
                                    usuario=usuario_actual
                                )
                                st.session_state['inversa_resultado_desglose'] = resultado
                                st.toast("✅ ¡Desglose completado y registrado exitosamente!")
                            except Exception as e_proc:
                                st.error(f"❌ Error durante el desglose: {e_proc}")

                    res_inv = st.session_state.get('inversa_resultado_desglose')
                    if res_inv:
                        st.success(f"🎉 **¡Expediente de {res_inv.get('empresa')} procesado con éxito!**")
                        c_res1, c_res2, c_res3 = st.columns(3)
                        with c_res1:
                            if res_inv.get('link_expediente'):
                                st.link_button("📑 Ver Expediente Final (Drive)", res_inv.get('link_expediente'), use_container_width=True)
                        with c_res2:
                            if res_inv.get('link_petramas'):
                                st.link_button("🎫 Ver Constancia Petramás", res_inv.get('link_petramas'), use_container_width=True)
                        with c_res3:
                            if res_inv.get('link_sustento'):
                                st.link_button("📊 Ver Sustento de Campo", res_inv.get('link_sustento'), use_container_width=True)

                        st.info(f"📊 Registrado en la pestaña **`{res_inv.get('hoja_historial')}`** (Fila {res_inv.get('fila_historial')}), y vinculado en `Certificados_Petramas` y `Sustentos_Campo`.")

    # ====================================================================
    # --- TAB 2: CONTROL DE SUSTENTOS (PENDIENTES / COMPLETOS) ---
    # ====================================================================
    with tab_sustentos:
        st.markdown("### 📊 Matriz de Control de Sustentos de Campo")
        st.caption("Supervisa los manifiestos, tickets de balanza y su cruce con las constancias de Petramás.")

        col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
        with col_f1:
            emp_filtro_sus = st.selectbox(
                "Filtrar por Empresa EO-RS:",
                ["Todas"] + list(EMPRESAS_EORS.keys()),
                key="filtro_empresa_sustentos"
            )
        with col_f2:
            txt_buscar_sus = st.text_input("🔍 Buscar por Placa, Boleta o Cliente:", key="busq_txt_sustentos")
        with col_f3:
            st.write("")
            st.write("")
            btn_recargar_sus = st.button("🔄 Recargar", use_container_width=True, key="btn_refresh_sustentos")

        # Cargar datos de sustentos
        with st.spinner("Consultando registros en hoja Sustentos_Campo..."):
            datos_sus = obtener_resumen_sustentos_pendientes(empresa=emp_filtro_sus)

        todos_sus = datos_sus.get("todos", [])
        if txt_buscar_sus:
            b_low = txt_buscar_sus.lower()
            todos_sus = [
                s for s in todos_sus
                if b_low in s.get("placa", "").lower()
                or b_low in s.get("boletas", "").lower()
                or b_low in s.get("generador", "").lower()
                or b_low in s.get("guia", "").lower()
            ]

        # Métricas principales
        cnt_pend_bol = sum(1 for s in todos_sus if "PENDIENTE BOLETAS" in s.get("estado", "").upper())
        cnt_pend_man = sum(1 for s in todos_sus if "PENDIENTE MANIFIESTOS" in s.get("estado", "").upper())
        cnt_completos = sum(1 for s in todos_sus if "SUSTENTOS COMPLETOS" in s.get("estado", "").upper())
        cnt_finalizados = sum(1 for s in todos_sus if "EXPEDIENTE" in s.get("estado", "").upper())

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("🟡 Pendiente Boletas", cnt_pend_bol)
        m2.metric("🟡 Pendiente Manifiestos", cnt_pend_man)
        m3.metric("🟢 Sustentos Completos", cnt_completos)
        m4.metric("✅ Expedientes Emitidos", cnt_finalizados)

        st.divider()

        # Ingesta rápida de sustento aislado
        with st.expander("➕ Cargar Nuevo Escaneo de Sustentos de Campo (PDF)", expanded=False):
            st.markdown("Carga un PDF que contenga manifiestos, boletas de pesaje o guías. La IA identificará automáticamente las boletas, placa y generador.")
            f_nuevo_sus = st.file_uploader("Archivo PDF de sustento:", type=["pdf"], key="up_nuevo_sustento_aislado")
            if f_nuevo_sus:
                c_ns1, c_ns2 = st.columns([2, 1])
                with c_ns1:
                    emp_nuevo_sus = st.selectbox("Empresa EO-RS:", list(EMPRESAS_EORS.keys()), key="sel_emp_nuevo_sus")
                with c_ns2:
                    st.write("")
                    st.write("")
                    btn_analizar_ns = st.button("🧠 Procesar y Guardar Sustento", type="primary", use_container_width=True, key="btn_save_ns")

                if btn_analizar_ns:
                    with st.spinner("Analizando con Gemini y subiendo a Google Drive..."):
                        try:
                            sus_bytes = f_nuevo_sus.getvalue()
                            ns_data = analizar_sustento_campo(sus_bytes)
                            ns_data["empresa_eors"] = emp_nuevo_sus
                            
                            fec_clean = str(ns_data.get("fecha_servicio", "")).replace("/", "")
                            placa_c = str(ns_data.get("placa", "CAMION"))
                            nom_drive = f"SUSTENTO-{placa_c}-{fec_clean}.pdf"
                            
                            res_drv_sus = subir_pdf_petramas_drive(sus_bytes, "sustentos", emp_nuevo_sus, nom_drive)
                            link_drv_sus = res_drv_sus.get("link", "")

                            guardar_sustento_campo_en_sheet(ns_data, link_drv_sus)
                            st.success(f"✅ Sustento registrado con éxito. Estado: `{ns_data.get('estado_documentacion')}`")
                            st.rerun()
                        except Exception as e_ns:
                            st.error(f"Error procesando sustento: {e_ns}")

        # Listado interactivo de sustentos
        st.markdown(f"#### Registros de Sustentos ({len(todos_sus)})")
        if not todos_sus:
            st.info("No se encontraron registros de sustentos con los filtros seleccionados.")
        else:
            for s in todos_sus:
                est = s.get("estado", "")
                badge = "🟢" if "COMPLETO" in est.upper() else ("✅" if "EXPEDIENTE" in est.upper() else "🟡")
                with st.expander(f"{badge} Fila {s.get('fila')} | {s.get('fecha_servicio')} | {s.get('empresa')} | Placa: {s.get('placa')} | {s.get('generador')}", expanded=False):
                    col_det1, col_det2, col_det3 = st.columns(3)
                    with col_det1:
                        st.markdown(f"🏢 **Empresa:** {s.get('empresa')}")
                        st.markdown(f"🏭 **Generador:** {s.get('generador')}")
                        st.markdown(f"📍 **Sede:** {s.get('sede')}")
                        st.markdown(f"📅 **Fecha Servicio:** {s.get('fecha_servicio')}")
                    with col_det2:
                        st.markdown(f"🚛 **Placa:** `{s.get('placa')}`")
                        st.markdown(f"📄 **Guía:** `{s.get('guia')}`")
                        st.markdown(f"⚖️ **Boletas Pesaje:** `{s.get('boletas')}`")
                        st.markdown(f"📑 **Rangos:** `{s.get('rangos')}`")
                    with col_det3:
                        st.markdown(f"🏷️ **Estado:** **{est}**")
                        st.markdown(f"🎫 **Petramás Vinculado:** `{s.get('cert_petramas') or 'Sin vincular'}`")
                        st.markdown(f"📜 **Certificado Ligado:** `{s.get('cert_epmi') or 'Sin vincular'}`")

                        if s.get("enlace_drive") and str(s.get("enlace_drive")).startswith("http"):
                            st.link_button("📂 Ver Escaneo Drive ↗", s.get("enlace_drive"), use_container_width=True)
                        if s.get("enlace_expediente") and str(s.get("enlace_expediente")).startswith("http"):
                            st.link_button("📑 Ver Expediente Final ↗", s.get("enlace_expediente"), use_container_width=True)

    # ====================================================================
    # --- TAB 3: ENSAMBLADOR DE EXPEDIENTES (5 PARTES) ---
    # ====================================================================
    with tab_ensamblador:
        st.markdown("### 📑 Ensamblador de Expediente de Disposición Final")
        st.markdown("""
        Ensambla un expediente completo conforme a la secuencia oficial de **5 partes**:
        1. **Certificado Emitido** por la EO-RS (EPMI, INECOVE o ALHUAY)
        2. **Constancia de Petramás** (Relleno Huaycoloro)
        3. **Manifiestos de Manejo de Residuos**
        4. **Boletas de Pesaje de Balanza**
        5. **Guías de Remisión y Transporte**
        """)

        c_ens1, c_ens2 = st.columns(2)
        with c_ens1:
            emp_ensamblar = st.selectbox("Empresa EO-RS:", list(EMPRESAS_EORS.keys()), key="sel_emp_ensamblar")
        with c_ens2:
            corr_ensamblar = st.text_input("N° Correlativo del Certificado:", placeholder="Ej: 039", key="input_corr_ensamblar")

        st.divider()
        st.markdown("#### Configuración de Componentes para el Expediente")

        col_c1, col_c2 = st.columns(2)

        with col_c1:
            st.markdown("##### 1. Certificado Emitido (Carátula PDF)")
            up_cert_caratula = st.file_uploader("Cargar Certificado Emitido (PDF):", type=["pdf"], key="up_ens_caratula")

            st.markdown("##### 2. Constancia de Petramás (PDF)")
            up_petramas_pdf = st.file_uploader("Cargar Constancia Petramás (PDF):", type=["pdf"], key="up_ens_petramas")

        with col_c2:
            st.markdown("##### 3, 4 y 5. Sustentos de Campo (Manifiestos, Boletas y Guías)")
            up_sustento_pdf = st.file_uploader("Cargar PDF con Sustentos de Campo:", type=["pdf"], key="up_ens_sustento")

            st.caption("Distribución de páginas dentro del archivo de sustentos (1-based, separadas por coma):")
            ens_pags_man = st.text_input("Páginas de Manifiestos (Parte 3):", placeholder="Ej: 1, 2, 3", key="ens_pags_man")
            ens_pags_bol = st.text_input("Páginas de Boletas de Pesaje (Parte 4):", placeholder="Ej: 4, 5", key="ens_pags_bol")
            ens_pags_gui = st.text_input("Páginas de Guías de Remisión (Parte 5):", placeholder="Ej: 6", key="ens_pags_gui")

        st.divider()

        def _pags_to_list(txt: str) -> List[int]:
            return [int(x) for x in txt.replace(" ", "").split(",") if x.isdigit()]

        ya_ensamblado = bool(st.session_state.get('resultado_ensamblado'))
        btn_ensamblar = st.button(
            "✅ Expediente Ya Ensamblado y Publicado" if ya_ensamblado else "⚡ Ensamblar y Publicar Expediente Final",
            type="primary",
            disabled=ya_ensamblado,
            use_container_width=True,
            key="btn_run_ensamblar"
        )

        if btn_ensamblar:
            if not up_cert_caratula:
                st.error("⚠️ Debes proporcionar obligatoriamente la Carátula del Certificado Emitido (Parte 1).")
            elif not corr_ensamblar:
                st.error("⚠️ Debes indicar el número correlativo.")
            else:
                with st.spinner("Fusionando documentos en orden estricto de 5 partes y subiendo a Google Drive..."):
                    try:
                        caratula_bytes = up_cert_caratula.getvalue()
                        pet_bytes = up_petramas_pdf.getvalue() if up_petramas_pdf else None
                        sus_bytes = up_sustento_pdf.getvalue() if up_sustento_pdf else None

                        p_man = _pags_to_list(ens_pags_man)
                        p_bol = _pags_to_list(ens_pags_bol)
                        p_gui = _pags_to_list(ens_pags_gui)

                        res_ens = ensamblar_expediente_disposicion_final(
                            empresa=emp_ensamblar,
                            correlativo=corr_ensamblar,
                            cert_caratula_pdf_bytes=caratula_bytes,
                            petramas_pdf_bytes=pet_bytes,
                            sustento_pdf_bytes=sus_bytes,
                            paginas_manifiestos=p_man,
                            paginas_boletas=p_bol,
                            paginas_guias=p_gui,
                            usuario_editor=usuario_actual
                        )

                        st.session_state['resultado_ensamblado'] = res_ens
                        st.toast("✅ Expediente ensamblado y publicado con éxito.")
                    except Exception as e_ens:
                        st.error(f"❌ Error ensamblando expediente: {e_ens}")

        res_ens = st.session_state.get('resultado_ensamblado')
        if res_ens:
            st.success(f"🎉 **¡Expediente '{res_ens.get('nombre_archivo')}' creado con éxito!** ({res_ens.get('total_paginas')} páginas)")
            c_e1, c_e2 = st.columns(2)
            with c_e1:
                if res_ens.get('enlace_drive'):
                    st.link_button("📂 Ver Expediente en Google Drive ↗", res_ens.get('enlace_drive'), use_container_width=True)
            with c_e2:
                if res_ens.get('pdf_bytes'):
                    st.download_button(
                        "⬇️ Descargar PDF Ensamblado",
                        data=res_ens.get('pdf_bytes'),
                        file_name=res_ens.get('nombre_archivo'),
                        mime="application/pdf",
                        use_container_width=True
                    )
            if st.button("🔄 Ensamblar Otro Expediente", key="btn_reset_ensamblar", use_container_width=True):
                del st.session_state['resultado_ensamblado']
                st.rerun()

    # ====================================================================
    # --- TAB 4: CONSTANCIAS PETRAMÁS ---
    # ====================================================================
    with tab_constancias:
        st.markdown("### 🎫 Ingesta y Catálogo de Constancias Petramás")
        st.caption("Registra y consulta las constancias de disposición final emitidas por el relleno Huaycoloro.")

        with st.expander("➕ Cargar y Analizar Nueva Constancia Petramás (PDF)", expanded=False):
            f_const_pdf = st.file_uploader("PDF de Constancia Petramás:", type=["pdf"], key="up_constancia_individual")
            if f_const_pdf:
                c_cp1, c_cp2 = st.columns([2, 1])
                with c_cp1:
                    emp_const = st.selectbox("Empresa EO-RS receptora:", list(EMPRESAS_EORS.keys()), key="sel_emp_const_ind")
                with c_cp2:
                    st.write("")
                    st.write("")
                    btn_analizar_cp = st.button("🧠 Analizar con Gemini", type="primary", use_container_width=True, key="btn_ana_cp")

                if btn_analizar_cp or 'constancia_analizada_data' in st.session_state:
                    if 'constancia_analizada_data' not in st.session_state:
                        with st.spinner("Extrayendo datos de la constancia Petramás con IA..."):
                            try:
                                d_cp = analizar_constancia_petramas(f_const_pdf.getvalue())
                                d_cp["empresa_eors"] = emp_const
                                st.session_state['constancia_analizada_data'] = d_cp
                                st.session_state['constancia_analizada_bytes'] = f_const_pdf.getvalue()
                            except Exception as e_cp:
                                st.error(f"Error analizando constancia: {e_cp}")
                                st.stop()

                    d_cp = st.session_state.get('constancia_analizada_data')
                    if d_cp:
                        st.markdown("##### Revisión de Datos Detectados")
                        c_ed1, c_ed2, c_ed3 = st.columns(3)
                        with c_ed1:
                            e_num_const = st.text_input("N° Constancia:", value=str(d_cp.get("numero_constancia", "")), key="edit_cp_num")
                            e_fec_disp = st.text_input("Fecha Disposición:", value=str(d_cp.get("fecha_disposicion", "")), key="edit_cp_fec")
                        with c_ed2:
                            e_gen = st.text_input("Generador:", value=str(d_cp.get("generador", "")), key="edit_cp_gen")
                            e_sede = st.text_input("Sede:", value=str(d_cp.get("sede", "")), key="edit_cp_sede")
                        with c_ed3:
                            e_peso = st.number_input("Peso Total (kg):", value=float(d_cp.get("peso_total_kg", 0.0)), step=10.0, key="edit_cp_peso")
                            e_boletas_str = st.text_area("Boletas de Pesaje:", value=", ".join(str(b) for b in d_cp.get("boletas", [])), key="edit_cp_bol")

                        btn_guardar_cp = st.button("🚀 Subir a Drive y Registrar en Certificados_Petramas", type="primary", use_container_width=True, key="btn_save_cp")

                        if btn_guardar_cp:
                            with st.spinner("Subiendo a Google Drive y registrando en Google Sheets..."):
                                try:
                                    cp_bytes = st.session_state['constancia_analizada_bytes']
                                    nombre_dest = f"CONSTANCIA-PETRAMAS-{e_num_const}.pdf"
                                    res_drv_cp = subir_pdf_petramas_drive(cp_bytes, "certificados", emp_const, nombre_dest)
                                    link_cp = res_drv_cp.get("link", "")

                                    payload_cp = {
                                        "empresa_eors": emp_const,
                                        "numero_constancia": e_num_const,
                                        "fecha_disposicion": e_fec_disp,
                                        "generador": e_gen,
                                        "sede": e_sede,
                                        "boletas": [b.strip() for b in e_boletas_str.split(",") if b.strip()],
                                        "peso_total_kg": e_peso,
                                        "certificado_ligado": ""
                                    }

                                    guardar_constancia_petramas_en_sheet(payload_cp, link_cp)
                                    st.success(f"✅ Constancia {e_num_const} registrada exitosamente.")
                                    st.session_state.pop('constancia_analizada_data', None)
                                    st.session_state.pop('constancia_analizada_bytes', None)
                                    st.rerun()
                                except Exception as e_gcp:
                                    st.error(f"Error guardando constancia: {e_gcp}")

        st.divider()
        st.markdown("#### Catálogo de Constancias Registradas en Google Sheets")

        col_fc1, col_fc2 = st.columns([3, 1])
        with col_fc1:
            emp_filtro_const = st.selectbox("Filtrar constancias por EO-RS:", ["Todas"] + list(EMPRESAS_EORS.keys()), key="filtro_cat_const")
        with col_fc2:
            st.write("")
            st.write("")
            btn_ref_cat = st.button("🔄 Recargar Catálogo", use_container_width=True, key="btn_ref_cat_const")

        constancias_list = obtener_constancias_petramas_registradas(empresa=emp_filtro_const)

        if not constancias_list:
            st.info("No se encontraron constancias registradas en 'Certificados_Petramas'.")
        else:
            for c in constancias_list:
                with st.container(border=True):
                    col_ci1, col_ci2, col_ci3 = st.columns([2, 2, 1])
                    with col_ci1:
                        st.markdown(f"🎫 **Constancia:** `{c.get('numero_constancia')}`")
                        st.markdown(f"🏢 **EO-RS:** {c.get('empresa')}")
                        st.markdown(f"📅 **Fecha:** {c.get('fecha_disposicion')}")
                    with col_ci2:
                        st.markdown(f"🏭 **Generador:** {c.get('generador')} ({c.get('sede')})")
                        st.markdown(f"⚖️ **Boletas:** `{c.get('boletas')}`")
                        st.markdown(f"📊 **Peso:** `{c.get('peso_total_kg')} kg`")
                    with col_ci3:
                        st.markdown(f"🏷️ **Estado:** `{c.get('estado')}`")
                        if c.get("enlace_drive") and str(c.get("enlace_drive")).startswith("http"):
                            st.link_button("📂 Ver PDF ↗", c.get("enlace_drive"), use_container_width=True)
