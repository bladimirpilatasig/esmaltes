import streamlit as st
import pandas as pd
import datetime
import os
import io
import time  # <--- Nueva importación para manejar la pausa del mensaje
from b_prep_esm_past import cargar_datos1 
from c_prep_esm_fuer import cargar_datos2




# ********************************************************************************************************
# ********************************************************************************************************
# 1. CONFIGURACIÓN GLOBAL (Layout Wide para usar toda la pantalla)
# ********************************************************************************************************
# ********************************************************************************************************
# layout="wide" permite que la app use el 100% del ancho disponible.
st.set_page_config(page_title="EGRESO DE ESMALTES", layout="wide")
#Nombre del Archivo
DB_FILE = "0.3_verificacion_bath.csv"
# ESto es para el selectbox
OPCIONES_TIPO_ESMALTE_VERIFICACION = ["FERRUM", "EB16", "EB10 ACCESORIOS", "PRUEBAS"," "]
OPCIONES_COLOR_PREPARADO = ["BLANCO","BONE", "VERDE PRIMAVERA", "CARIBBEAN SHELL", "GRIS", "CELESTE","NEGRO", "CIPRÉS", "AZUL MEDITERRANEO", "MERLOT", "VERDE", "NARANJA","ROJO", "MORA","AZUL SAFIRO"]
# ---------------------------------------------------------
# CAMPOS DE LA BASE DE DATOS: 
# Si quieres agregar campos de cálculo (ej. 'DIFERENCIA_HORAS'), añádelos a esta lista.
COLUMNAS_BASE = ["ID", "BATCH_VERI", "COLOR_VERI", "TIPO_ESMALTE_VERI", "VOLUMEN_VERI", "DENSIDAD(KG/L)_VERI", "KG_SECOS_VERI"]
# ----------------------------------------------------------
# --- ESTILOS CSS ---
st.markdown("""
    <style>
    /* Botones normales */
    .stButton>button {
        background-color: #E65100;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        border: none;
    }
    .stButton>button:hover {
        background-color: #BF360C;
        color: white;
    }

    /* Botón dentro de formularios */
    div[data-testid="stForm"] button {
        background-color: #E65100;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        border: none;
    }
    div[data-testid="stForm"] button:hover {
        background-color: #BF360C;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)





# ********************************************************************************************************
# ********************************************************************************************************
# 2. LÓGICA DE PERSISTENCIA (Carga y Guardado)
# ********************************************************************************************************
# ********************************************************************************************************
@st.cache_data(show_spinner=False)
def cargar_datos3():
    """Lee el CSV local. Si no existe, crea uno nuevo con las columnas base."""
    if not os.path.exists(DB_FILE):
        return pd.DataFrame(columns=COLUMNAS_BASE)
    df = pd.read_csv(DB_FILE)
    # Aseguramos que el ID sea entero para evitar decimales innecesarios
    df["ID"] = pd.to_numeric(df["ID"], errors='coerce').fillna(0).astype(int)
    return df
def guardar_datos3(df):
    """Guarda el DataFrame y limpia el caché para refrescar la vista."""
    df.to_csv(DB_FILE, index=False)
    st.cache_data.clear()
def generar_excel3(df):
    """Convierte el DataFrame actual a un archivo binario de Excel."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Novedades')
    return output.getvalue()





# ********************************************************************************************************
# ********************************************************************************************************
# 3. INTERFAZ DE USUARIO PRINCIPAL
# ********************************************************************************************************
# ********************************************************************************************************
st.markdown("<h1 style='text-align: center; color:darkblue ;'>📊 VERIFICACION DE BATCH </h1>", unsafe_allow_html=True)
#st.divider()
# Pestañas de navegación central
tab_add, tab_view, tab_edit, tab_del = st.tabs([
    "      📥 INGRESAR REGISTRO", "      📋 VISUALIZAR REGISTRO", "       ✏️ EDITAR REGISTRO", "       🗑️ ELIMINAR REGISTRO"
])





# ********************************************************************************************************
# ********************************************************************************************************
# --- PESTAÑA: INGRESAR ---
# ********************************************************************************************************
# ********************************************************************************************************
with tab_add:
    df = cargar_datos3()
    # GENERADOR DEL ID
    id_actual = 1 if df.empty else int(df["ID"].max()) + 1
    st.subheader(f"Registro No. {id_actual}")
    
    # --- FORMULARIO DE INGRESO DE DATOS ---
    # BATCH - CORREGIDO: key añadido
    batch_verificacion_texto = st.text_input('BATCH (No.)', value="", placeholder="Ingresa el número de BATCH", key="txt_batch_veri")
    try:
        batch_verificacion_in = float(batch_verificacion_texto.replace(',', '.')) if batch_verificacion_texto else 0.000
    except ValueError:
        batch_verificacion_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en el BATCH.")
    # COLOR - CORREGIDO: key añadido
    color_verificacion_in = st.selectbox("COLOR", OPCIONES_COLOR_PREPARADO, key="sb_color_veri")
    
    # TIPO DE ESMALTE - CORREGIDO: key añadido
    tipo_esmalte_verificacion_in = st.selectbox("TIPO DE ESMALTE", OPCIONES_TIPO_ESMALTE_VERIFICACION, key="sb_tipo_esm_veri")
    # VOLUMEN DE LA MEDICIÓN - CORREGIDO: key añadido
    volumen_verificacion_texto = st.text_input('VOLUMEN (lt)', value="", placeholder="Ingresa el volumen en litros", key="txt_volumen_veri")
    try:
        volumen_verificacion_in = float(volumen_verificacion_texto.replace(',', '.')) if volumen_verificacion_texto else 0.000
    except ValueError:
        volumen_verificacion_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en el VOLUMEN.")
    # DENSIDAD DE LA MEDICIÓN - CORREGIDO: key añadido
    densidad_texto = st.text_input('DENSIDAD (Kg/lt)', value="", placeholder="Ingresa la densidad en Kg/lt", key="txt_densidad_veri")
    try:
        densidad_verificacion_in = float(densidad_texto.replace(',', '.')) if densidad_texto else 0.000
    except ValueError:
        densidad_verificacion_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en la DENSIDAD.")
    # --- BLOQUE DE CÁLCULOS EN TIEMPO REAL ---
    pendiente = 1.5419 # ecuacion de esmaltes: pendiente = 1.5419
    ordenada_origen = -1.5435 # ecuacion de esmaltes ordenada_origen = -1.5435
    kg_verificacion_secos = volumen_verificacion_in * ((pendiente * densidad_verificacion_in) + ordenada_origen)  # Peso seco en Kg
    # --- MOSTRAR RESULTADOS EN PANTALLA ---
    st.info(f"📐 **Kilogramos Secos de verificación (Kg):** {kg_verificacion_secos:.2f}")



    # Inicializar la variable de estado para la confirmación si no existe
    if "confirmar_guardado_veri" not in st.session_state:
        st.session_state.confirmar_guardado_veri = False
    # Botón principal para iniciar el proceso de guardado - CORREGIDO: key añadido
    btn_save = st.button("Guardar Registro", disabled=st.session_state.confirmar_guardado_veri, key="btn_save_veri")
    if btn_save:
        # FILTRO DE VALIDACION DE DATOS PREVIO
        if not batch_verificacion_in or not tipo_esmalte_verificacion_in or not volumen_verificacion_in or not densidad_verificacion_in:
            st.error("⚠️ Error: Por favor completa todos los campos.")
        else:
            # Activamos el modo de confirmación
            st.session_state.confirmar_guardado_veri = True
            st.rerun()



    # --- BLOQUE DE CONFIRMACIÓN (SÍ / NO) ---
    if st.session_state.confirmar_guardado_veri:
        st.warning(f"❓ ¿Está seguro de que desea guardar el Registro No. {id_actual}?")
        
        # Colocamos los dos botones uno al lado del otro
        col_si, col_no = st.columns(2)
        
        with col_si:
            btn_si = st.button("✔️ SÍ, guardar", use_container_width=True, key="btn_si_veri")
        with col_no:
            btn_no = st.button("❌ NO, cancelar", use_container_width=True, key="btn_no_veri")
        if btn_si:
            # Construcción del registro
            nuevo_item = {
                "ID": id_actual,
                "BATCH_VERI": batch_verificacion_in,
                "COLOR_VERI": color_verificacion_in,
                "TIPO_ESMALTE_VERI": tipo_esmalte_verificacion_in,
                "VOLUMEN_VERI": volumen_verificacion_in,
                "DENSIDAD(KG/L)_VERI": densidad_verificacion_in,
                "KG_SECOS_VERI": kg_verificacion_secos
            }
            # Guardado
            df_final = pd.concat([df, pd.DataFrame([nuevo_item])], ignore_index=True)
            guardar_datos3(df_final)
            # Éxito y limpieza de estado
            st.success(f"✅ Registro #{id_actual} guardado exitosamente.")
            st.session_state.confirmar_guardado_veri = False
            time.sleep(2)
            st.rerun()
        if btn_no:
            # Cancelar acción y reestablecer la vista normal
            st.session_state.confirmar_guardado_veri = False
            st.rerun()





# ********************************************************************************************************
# ********************************************************************************************************
# --- PESTAÑA: VISUALIZAR Y FILTRAR ---
# ********************************************************************************************************
# ********************************************************************************************************
with tab_view:
    st.write("")
    st.write("")
    st.markdown("<h3 style='text-align: center; color:darkgreen ;'>TABLA DE DATOS DE VERIFICACIÓN DE ESMALTES </h3>", unsafe_allow_html=True)
    # Cargar datos actualizados
    df_visualizar = cargar_datos3()
    # Columnas que se mostrarán en la pestaña Visualizar
    columnas_mostrar = [
        "ID",
        "BATCH_VERI",
        "COLOR_VERI",
        "TIPO_ESMALTE_VERI",
        "VOLUMEN_VERI",
        "DENSIDAD(KG/L)_VERI",
        "KG_SECOS_VERI",
    ]
    df_visualizar = df_visualizar[columnas_mostrar]
    if df_visualizar.empty:
        st.warning("📭 No hay registros guardados todavía.")
    else:
        # 1. MOSTRAR TABLA COMPLETA
        st.dataframe(
            df_visualizar.style.format(precision=0),
            use_container_width=True, 
            hide_index=True
        )
        # --- BOTÓN DE DESCARGA: TODOS LOS DATOS (EXCEL) ---
        buffer_todos = io.BytesIO()
        with pd.ExcelWriter(buffer_todos, engine='openpyxl') as writer:
            df_visualizar.to_excel(writer, index=False, sheet_name='Todos los Registros')
        st.download_button(
            label="📥 Descargar TODOS los datos en Excel",
            data=buffer_todos.getvalue(),
            file_name='todos_los_registros_esmaltes.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type="primary",
            key="btn_descarga_todos_veri"
        )



    # TABLA DE BALANCES DE ESMALTES PREPARADOS, FUERTES VS VERIFICADOS
    # --- FUNCIÓN COMPACTADA (PROCESAMIENTO EN UNA LÍNEA) ---
    def calcular_balance(df_p_ind, df_v_comp):
        # Agrupa preparación
        p_g = df_p_ind.groupby(["BATCH_PREP", "COLOR_PREP", "TIPO_ESMALTE_PREP"], as_index=False)["KG_SECOS_PREP"].sum().rename(columns={"BATCH_PREP": "BATCH", "COLOR_PREP": "COLOR", "TIPO_ESMALTE_PREP": "TIPO_ESMALTE", "KG_SECOS_PREP": "TOTAL_PREP"})
        # Agrupa, filtra por las 3 columnas usando inner merge, cruza (outer) y calcula la diferencia en una sola línea continua
        return pd.merge(p_g, pd.merge(df_v_comp.groupby(["BATCH_VERI", "COLOR_VERI", "TIPO_ESMALTE_VERI"], as_index=False)["KG_SECOS_VERI"].sum().rename(columns={"BATCH_VERI": "BATCH", "COLOR_VERI": "COLOR", "TIPO_ESMALTE_VERI": "TIPO_ESMALTE", "KG_SECOS_VERI": "TOTAL_VERI"}), p_g[["BATCH", "COLOR", "TIPO_ESMALTE"]], on=["BATCH", "COLOR", "TIPO_ESMALTE"], how="inner"), on=["BATCH", "COLOR", "TIPO_ESMALTE"], how="outer").fillna(0).assign(DIFERENCIA_KG=lambda df: df["TOTAL_PREP"] - df["TOTAL_VERI"])
    # --- 1. CARGA DE DATOS Y CÁLCULOS ---
    df_prep, df_fuer, df_veri = cargar_datos1(), cargar_datos2(), cargar_datos3()
    df_b1 = calcular_balance(df_prep, df_veri)
    df_b2 = calcular_balance(df_fuer, df_veri)
    # Tabla 3 (Consolidado general en una sola línea de procesamiento)
    df_balance_final = pd.merge(pd.concat([df_prep, df_fuer], ignore_index=True).groupby(["BATCH_PREP", "COLOR_PREP", "TIPO_ESMALTE_PREP"], as_index=False)["KG_SECOS_PREP"].sum().rename(columns={"BATCH_PREP": "BATCH", "COLOR_PREP": "COLOR", "TIPO_ESMALTE_PREP": "TIPO_ESMALTE", "KG_SECOS_PREP": "TOTAL_PREP"}), df_veri.groupby(["BATCH_VERI", "COLOR_VERI", "TIPO_ESMALTE_VERI"], as_index=False)["KG_SECOS_VERI"].sum().rename(columns={"BATCH_VERI": "BATCH", "COLOR_VERI": "COLOR", "TIPO_ESMALTE_VERI": "TIPO_ESMALTE", "KG_SECOS_VERI": "TOTAL_VERI"}), on=["BATCH", "COLOR", "TIPO_ESMALTE"], how="outer").fillna(0).assign(DIFERENCIA_KG=lambda df: df["TOTAL_PREP"] - df["TOTAL_VERI"])
    # --- 2. MOSTRAR TABLAS EN STREAMLIT (SIN DECIMALES) ---
    formato = {"BATCH": "{:.0f}","TOTAL_PREP": "{:.0f}", "TOTAL_VERI": "{:.0f}", "DIFERENCIA_KG": "{:.0f}"}
    st.markdown("<h4 style='text-align: left; color:black ;'>Tabla Agrupada de Preparacion Esmalte Pasteles vs Verificación </h4>", unsafe_allow_html=True)
    st.dataframe(df_b1.style.format(formato))
    st.markdown("<h4 style='text-align: left; color:black ;'>Tabla Agrupada de Preparacion Esmalte Fuertes vs Verificación </h4>", unsafe_allow_html=True)
    st.dataframe(df_b2.style.format(formato))
    st.markdown("<h3 style='text-align: center; color:purple ;'>✅ TABLA DE DATOS PREPARACION DE ESMALTES VS VERIFICACION </h3>", unsafe_allow_html=True)
    st.dataframe(df_balance_final.style.format(formato))
    # --- 3. EXCEL MULTIPESTAÑA Y DESCARGA ---
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_b1.to_excel(w, index=False, sheet_name="Balance_Tabla_1")
        df_b2.to_excel(w, index=False, sheet_name="Balance_Tabla_2")
        df_balance_final.to_excel(w, index=False, sheet_name="Balance_Total")
    st.download_button("📥 Descargar Todos los Balances en Excel", buf.getvalue(), "balances_esmaltes.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")




    # --- TABLA DE FILTRO DINAMICO ---
    st.write("")
    st.write("")
    st.markdown("<h3 style='text-align: center; color:darkorange ;'>🔍 TABLA DE DATOS FILTRADA </h3>", unsafe_allow_html=True)
    # 2. SELECTBOX PARA ELEGIR LA COLUMNA
    columnas_disponibles = df_balance_final.columns.tolist()
    columna_seleccionada = st.selectbox(
        "Selecciona la columna por la que deseas filtrar:", 
        options=columnas_disponibles
    )
    # --- NUEVO: OBTENER VALORES ÚNICOS Y ORDENADOS SIN REPETIR ---
    # Eliminamos valores nulos (dropna) y obtenemos los registros únicos (.unique())
    valores_unicos = df_balance_final[columna_seleccionada].dropna().unique().tolist()
    valores_unicos.sort()  # Los ordena alfabéticamente o numéricamente para que sea más fácil buscar
    # 3. SEGUNDO SELECTBOX EN LUGAR DE TEXT_INPUT
    busqueda_seleccionada = st.selectbox(
        f"Selecciona el valor que deseas buscar en '{columna_seleccionada}':",
        options=valores_unicos,
        placeholder="Selecciona una opción...",
        index=0  # Por defecto selecciona el primer valor único
    )
    # 4. LÓGICA DE FILTRADO (COMPROBACIÓN ROBUSTA)
    df_filtrado = df_balance_final.copy()
    # Nos aseguramos de que haya algo seleccionado en el segundo selectbox
    if busqueda_seleccionada is not None and busqueda_seleccionada != "":
        # Convertimos temporalmente a string ambos lados para evitar conflictos de tipo (int vs float vs str)
        # y usamos una comparación exacta limpia.
        df_filtrado = df_balance_final[
            df_balance_final[columna_seleccionada].astype(str).str.strip() == str(busqueda_seleccionada).strip()
        ]
    # 5. MOSTRAR RESULTADOS DEL FILTRO (Esto debe ejecutarse SIEMPRE)
    st.write(f"📋 **Resultados del filtro:** ({len(df_filtrado)} registros encontrados)")
    if not df_filtrado.empty:
        st.dataframe(df_filtrado.style.format(precision=0), use_container_width=True, hide_index=True)
        
        # --- BOTÓN DE DESCARGA: DATOS FILTRADOS (EXCEL) ---
        buffer_filtrado = io.BytesIO()
        with pd.ExcelWriter(buffer_filtrado, engine='openpyxl') as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name='Datos Filtrados')
        st.download_button(
            label="📥 Descargar tabla FILTRADA en Excel",
            data=buffer_filtrado.getvalue(),
            file_name='datos_filtrados_esmaltes.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type="primary",
            key="btn_descarga_filtrado"
        )
    else:
        st.info("No hay datos que coincidan con la búsqueda.")





# ********************************************************************************************************
# ********************************************************************************************************
# --- PESTAÑA: EDITAR REGISTRO ---
# ********************************************************************************************************
# ********************************************************************************************************
with tab_edit:
    df_editar = cargar_datos3()
    st.markdown("<h3 style='text-align: center; color:darkblue ;'>✏️ EDITAR REGISTRO EXISTENTE</h3>", unsafe_allow_html=True)
    
    if df_editar.empty:
        st.warning("📭 No hay registros para editar.")
    else:
        
        
        # Selector para elegir el ID del registro a editar
        lista_ids = df_editar["ID"].tolist()
        id_seleccionar = st.selectbox("Selecciona el ID del registro a EDITAR:", options=lista_ids, key="sb_id_editar")
        
        # Obtener los datos actuales del registro seleccionado
        fila_actual = df_editar[df_editar["ID"] == id_seleccionar].iloc[0]
        
        # --- FORMULARIO DE EDICIÓN DE DATOS ---
        batch_edit_texto = st.text_input('BATCH (No.)', value=str(fila_actual["BATCH_VERI"]), key="txt_batch_edit")
        try:
            batch_edit_in = float(batch_edit_texto.replace(',', '.')) if batch_edit_texto else 0.000
        except ValueError:
            batch_edit_in = 0.000
            st.error("⚠️ Por favor, ingresa un número válido en el BATCH.")
            
        # Buscar índice por defecto para el color
        try:
            idx_color = OPCIONES_COLOR_PREPARADO.index(fila_actual["COLOR_VERI"])
        except ValueError:
            idx_color = 0
        color_edit_in = st.selectbox("COLOR", OPCIONES_COLOR_PREPARADO, index=idx_color, key="sb_color_edit")
        
        # Buscar índice por defecto para el tipo de esmalte
        try:
            idx_tipo = OPCIONES_TIPO_ESMALTE_VERIFICACION.index(fila_actual["TIPO_ESMALTE_VERI"])
        except ValueError:
            idx_tipo = 0
        tipo_esmalte_edit_in = st.selectbox("TIPO DE ESMALTE", OPCIONES_TIPO_ESMALTE_VERIFICACION, index=idx_tipo, key="sb_tipo_esm_edit")
        
        volumen_edit_texto = st.text_input('VOLUMEN (lt)', value=str(fila_actual["VOLUMEN_VERI"]), key="txt_volumen_edit")
        try:
            volumen_edit_in = float(volumen_edit_texto.replace(',', '.')) if volumen_edit_texto else 0.000
        except ValueError:
            volumen_edit_in = 0.000
            st.error("⚠️ Por favor, ingresa un número válido en el VOLUMEN.")
            
        densidad_edit_texto = st.text_input('DENSIDAD (Kg/lt)', value=str(fila_actual["DENSIDAD(KG/L)_VERI"]), key="txt_densidad_edit")
        try:
            densidad_edit_in = float(densidad_edit_texto.replace(',', '.')) if densidad_edit_texto else 0.000
        except ValueError:
            densidad_edit_in = 0.000
            st.error("⚠️ Por favor, ingresa un número válido en la DENSIDAD.")
            
        # --- BLOQUE DE CÁLCULOS EN TIEMPO REAL ---
        pendiente = 1.5419
        ordenada_origen = -1.5435
        kg_edit_secos = volumen_edit_in * ((pendiente * densidad_edit_in) + ordenada_origen)
        
        st.info(f"📐 **Nuevos Kilogramos Secos de verificación (Kg):** {kg_edit_secos:.2f}")
        
        # Inicializar estado de confirmación para edición
        if "confirmar_edit_veri" not in st.session_state:
            st.session_state.confirmar_edit_veri = False
            
        btn_update = st.button("Actualizar Registro", disabled=st.session_state.confirmar_edit_veri, key="btn_update_veri")
        if btn_update:
            if not batch_edit_in or not tipo_esmalte_edit_in or not volumen_edit_in or not densidad_edit_in:
                st.error("⚠️ Error: Por favor completa todos los campos con valores válidos.")
            else:
                st.session_state.confirmar_edit_veri = True
                st.rerun()
                
        # --- BLOQUE DE CONFIRMACIÓN (SÍ / NO) ---
        if st.session_state.confirmar_edit_veri:
            st.warning(f"❓ ¿Está seguro de que desea guardar los cambios en el Registro No. {id_seleccionar}?")
            col_si_ed, col_no_ed = st.columns(2)
            
            with col_si_ed:
                btn_si_ed = st.button("✔️ SÍ, actualizar", use_container_width=True, key="btn_si_edit_veri")
            with col_no_ed:
                btn_no_ed = st.button("❌ NO, cancelar", use_container_width=True, key="btn_no_edit_veri")
                
            if btn_si_ed:
                # Modificar fila en el DataFrame original
                idx_target = df_editar[df_editar["ID"] == id_seleccionar].index[0]
                df_editar.at[idx_target, "BATCH_VERI"] = batch_edit_in
                df_editar.at[idx_target, "COLOR_VERI"] = color_edit_in
                df_editar.at[idx_target, "TIPO_ESMALTE_VERI"] = tipo_esmalte_edit_in
                df_editar.at[idx_target, "VOLUMEN_VERI"] = volumen_edit_in
                df_editar.at[idx_target, "DENSIDAD(KG/L)_VERI"] = densidad_edit_in
                df_editar.at[idx_target, "KG_SECOS_VERI"] = kg_edit_secos
                
                guardar_datos3(df_editar)
                st.success(f"✅ Registro #{id_seleccionar} actualizado exitosamente.")
                st.session_state.confirmar_edit_veri = False
                time.sleep(2)
                st.rerun()
                
            if btn_no_ed:
                st.session_state.confirmar_edit_veri = False
                st.rerun()





# ********************************************************************************************************
# ********************************************************************************************************
# --- PESTAÑA: ELIMINAR REGISTRO ---
# ********************************************************************************************************
# ********************************************************************************************************
with tab_del:
    df_eliminar = cargar_datos3()
    st.markdown("<h3 style='text-align: center; color:darkred ;'>🗑️ ELIMINAR REGISTRO</h3>", unsafe_allow_html=True)
    
    if df_eliminar.empty:
        st.warning("📭 No hay registros para eliminar.")
    else:
        
        
        # Selector para elegir el ID del registro a borrar
        lista_ids_del = df_eliminar["ID"].tolist()
        id_eliminar_sel = st.selectbox("Selecciona el ID del registro que deseas ELIMINAR:", options=lista_ids_del, key="sb_id_eliminar")
        
        # Extraer y estructurar la fila como un DataFrame individual para mostrarlo en formato tabla
        fila_tabla_del = df_eliminar[df_eliminar["ID"] == id_eliminar_sel]
        
        st.write("📋 **Vista previa del registro seleccionado para eliminación:**")
        # Mostrar en formato tabla exacta sin índices antes de realizar la acción
        st.dataframe(
            fila_tabla_del.style.format(precision=0),
            use_container_width=True,
            hide_index=True
        )
        
        # Inicializar estado de confirmación para eliminación
        if "confirmar_del_veri" not in st.session_state:
            st.session_state.confirmar_del_veri = False
            
        btn_delete = st.button("Eliminar Registro", disabled=st.session_state.confirmar_del_veri, key="btn_delete_veri")
        if btn_delete:
            st.session_state.confirmar_del_veri = True
            st.rerun()
            
        # --- BLOQUE DE CONFIRMACIÓN (SÍ / NO) ---
        if st.session_state.confirmar_del_veri:
            st.error(f"🚨 ALERTAL: ¿Está completamente seguro de que desea ELIMINAR permanentemente el Registro No. {id_eliminar_sel}?")
            col_si_del, col_no_del = st.columns(2)
            
            with col_si_del:
                btn_si_del = st.button("💥 SÍ, eliminar permanentemente", use_container_width=True, key="btn_si_del_veri")
            with col_no_del:
                btn_no_del = st.button("❌ NO, cancelar", use_container_width=True, key="btn_no_del_veri")
                
            if btn_si_del:
                # Filtrar el DataFrame para remover la fila correspondiente
                df_final_del = df_eliminar[df_eliminar["ID"] != id_eliminar_sel]
                guardar_datos3(df_final_del)
                
                st.success(f"🗑️ Registro #{id_eliminar_sel} eliminado exitosamente.")
                st.session_state.confirmar_del_veri = False
                time.sleep(2)
                st.rerun()
                
            if btn_no_del:
                st.session_state.confirmar_del_veri = False
                st.rerun()