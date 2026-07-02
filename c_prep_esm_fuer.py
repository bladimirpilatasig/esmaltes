import streamlit as st
import pandas as pd
import datetime
import os
import io
import time  # <--- Nueva importación para manejar la pausa del mensaje



# ********************************************************************************************************
# ********************************************************************************************************
# 1. CONFIGURACIÓN GLOBAL (Layout Wide para usar toda la pantalla)
# ********************************************************************************************************
# ********************************************************************************************************
# layout="wide" permite que la app use el 100% del ancho disponible.
st.set_page_config(page_title="PREP. ESMALTES FUERTES", layout="wide")
#Nombre del Archivo
DB_FILE = "0.2_prep_esm_fuer.csv"
# ESto es para el selectbox
OPCIONES_COLOR_PREPARADO_FUERTE = ["NEGRO", "CIPRÉS", "AZUL MEDITERRANEO", "MERLOT", "VERDE", "NARANJA","ROJO", "MORA","AZUL SAFIRO" ]
OPCIONES_TIPO_ESMALTE_PREPARADO_FUERTE = ["FERRUM", "EB16", "EB10 ACCESORIOS", "PRUEBAS"," "]
# ---------------------------------------------------------
# CAMPOS DE LA BASE DE DATOS: 
# Si quieres agregar campos de cálculo (ej. 'DIFERENCIA_HORAS'), añádelos a esta lista.
COLUMNAS_BASE_PREPARADO_FUERTE = ["ID", "BATCH_PREP_FUER","COLOR_PREP_FUER","TIPO_ESMALTE_PREP_FUER","VOLUMEN(L)_PREP_FUER","KG_HUMEDOS_PREP_FUER", "DENSIDAD(KG/L)_PREP_FUER","KG_SECOS_PREP_FUER","KG_VIRGEN_PREP_FUER","%_RECUPERADOS_PREP_FUER","KG_RECUPERADOS_PREP_FUER" ]
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
def cargar_datos2():
    """Lee el CSV local. Si no existe, crea uno nuevo con las columnas base."""
    if not os.path.exists(DB_FILE):
        return pd.DataFrame(columns=COLUMNAS_BASE_PREPARADO_FUERTE)
    df = pd.read_csv(DB_FILE)
    # Aseguramos que el ID sea entero para evitar decimales innecesarios
    df["ID"] = pd.to_numeric(df["ID"], errors='coerce').fillna(0).astype(int)
    return df
def guardar_datos2(df):
    """Guarda el DataFrame y limpia el caché para refrescar la vista."""
    df.to_csv(DB_FILE, index=False)
    st.cache_data.clear()
def generar_excel2(df):
    """Convierte el DataFrame actual a un archivo binario de Excel."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Novedades')
    return output.getvalue()




# ********************************************************************************************************
# ********************************************************************************************************
# 3. INTERFAZ DE USUARIO PRINCIPASL
# ********************************************************************************************************
# ********************************************************************************************************
st.markdown("<h1 style='text-align: center; color:darkblue ;'>📊 PREPARACIÓN DE ESMALTES FUERTES</h1>", unsafe_allow_html=True)
#st.divider()
# Pestañas de navegación central
tab_add, tab_view, tab_edit, tab_del = st.tabs([
    "       📥 INGRESAR REGISTRO", "      📋 VISUALIZAR REGISTRO", "       ✏️ EDITAR REGISTRO", "       🗑️ ELIMINAR REGISTRO"
])




# ********************************************************************************************************
# ********************************************************************************************************
# --- PESTAÑA: INGRESAR ---
# ********************************************************************************************************
# ********************************************************************************************************
with tab_add:
    df = cargar_datos2()
    # GENERADOR DEL ID
    id_actual = 1 if df.empty else int(df["ID"].max()) + 1
    st.subheader(f"Registro No. {id_actual}")


    # --- FORMULARIO DE INGRESO DE DATOS ---
    # BATCH
    batch_texto = st.text_input('BATCH (No.)', value="", placeholder="Ingresa el número de BATCH")
    try:
        batch_preparadof_in = float(batch_texto.replace(',', '.')) if batch_texto else 0.000
    except ValueError:
        batch_preparadof_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en el BATCH.")
    # COLOR
    color_preparadof_in = st.selectbox("COLOR", OPCIONES_COLOR_PREPARADO_FUERTE)
    # TIPO DE ESMALTE
    tipo_esmalte_preparadof_in = st.selectbox("TIPO DE ESMALTE", OPCIONES_TIPO_ESMALTE_PREPARADO_FUERTE)
    # VOLUMEN DE LA MEDICIÓN
    volumen_texto = st.text_input('VOLUMEN (L)', value="", placeholder="Ingresa el volumen en litros")
    try:
        volumen_preparadof_in = float(volumen_texto.replace(',', '.')) if volumen_texto else 0.000
    except ValueError:
        volumen_preparadof_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en el VOLUMEN.")
    # DENSIDAD DE LA MEDICIÓN
    densidad_texto = st.text_input('DENSIDAD (Kg/lt³)', value="", placeholder="Ingresa la densidad en Kg/lt³")
    try:
        densidad_preparadof_in = float(densidad_texto.replace(',', '.')) if densidad_texto else 0.000
    except ValueError:
        densidad_preparadof_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en la DENSIDAD.")
    # PARTE RECUPERADA EN %
    porcentaje_recuperada_texto = st.text_input('PARTE RECUPERADA (%)', value="", placeholder="Ingresa la parte recuperada en porcentaje")
    porcentaje_recuperadaf_in = float(porcentaje_recuperada_texto.replace(',', '.')) if porcentaje_recuperada_texto else 0.000
    # --- BLOQUE DE CÁLCULOS EN TIEMPO REAL ---
    kg_humedos_preparadof = volumen_preparadof_in * densidad_preparadof_in  # Peso final en Kg
    pendiente = 1.5816 # ecuacion de esmaltes: pendiente = 1.5419
    ordenada_origen = -1.5791 # ecuacion de esmaltes ordenada_origen = -1.5435
    kg_preparadof_secos = volumen_preparadof_in * ((pendiente * densidad_preparadof_in) + ordenada_origen)  # Peso seco en Kg, redondeado a 2 decimales
    kg_virgenf = kg_preparadof_secos * (100 - porcentaje_recuperadaf_in) / 100
    kg_recuperadosf = kg_preparadof_secos - kg_virgenf  # Parte recuperada en Kg, redondeado a 2 decimales



    # --- MOSTRAR RESULTADOS EN PANTALLA ---
    st.info(f"📐 **Kilogramos Húmedos (Kg):** {kg_humedos_preparadof:.2f}")
    st.info(f"📐 **Kilogramos Secos (Kg):** {kg_preparadof_secos:.2f}")
    st.info(f"📐 **Kilogramos Virgen (Kg):** {kg_virgenf:.2f}")
    st.info(f"📐 **Kilogramos Recuperados (Kg):** {kg_recuperadosf:.2f}")



    # Inicializar la variable de estado para la confirmación si no existe
    if "confirmar_guardado" not in st.session_state:
        st.session_state.confirmar_guardado = False
    # Botón principal para iniciar el proceso de guardado
    btn_save = st.button("Guardar Registro", disabled=st.session_state.confirmar_guardado)
    if btn_save:
        # FILTRO DE VALIDACION DE DATOS PREVIO
        if not batch_preparadof_in or not color_preparadof_in or not volumen_preparadof_in or not densidad_preparadof_in:
            st.error("⚠️ Error: Por favor completa los campos de Batch, Color, Volumen y Densidad.")
        else:
            # Activamos el modo de confirmación
            st.session_state.confirmar_guardado = True
            st.rerun()



    # --- BLOQUE DE CONFIRMACIÓN (SÍ / NO) ---
    if st.session_state.confirmar_guardado:
        st.warning(f"❓ ¿Está seguro de que desea guardar el Registro No. {id_actual}?")
        
        # Colocamos los dos botones uno al lado del otro
        col_si, col_no = st.columns(2)
        
        with col_si:
            btn_si = st.button("✔️ SÍ, guardar", use_container_width=True)
        with col_no:
            btn_no = st.button("❌ NO, cancelar", use_container_width=True)
        if btn_si:
            # --- CONSTRUCCIÓN DEL REGISTRO CORREGIDO ---
            nuevo_item = {
                "ID": id_actual,
                "BATCH_PREP": batch_preparadof_in,
                "COLOR_PREP": color_preparadof_in,
                "TIPO_ESMALTE_PREP": tipo_esmalte_preparadof_in,
                "VOLUMEN(L)_PREP_FUER": volumen_preparadof_in,
                "KG_HUMEDOS_PREP_FUER": kg_humedos_preparadof,
                "DENSIDAD(KG/L)_PREP_FUER": densidad_preparadof_in,
                "KG_SECOS_PREP": kg_preparadof_secos,
                "KG_VIRGEN_PREP_FUER": kg_virgenf,
                "%_RECUPERADOS_PREP_FUER": porcentaje_recuperadaf_in,
                "KG_RECUPERADOS_PREP_FUER": kg_recuperadosf,
            }
            # Guardado
            df_final = pd.concat([df, pd.DataFrame([nuevo_item])], ignore_index=True)
            guardar_datos2(df_final)
            # Éxito y limpieza de estado
            st.success(f"✅ Registro #{id_actual} guardado exitosamente.")
            st.session_state.confirmar_guardado = False
            time.sleep(2)
            st.rerun()
        if btn_no:
            # Cancelar acción y reestablecer la vista normal
            st.session_state.confirmar_guardado = False
            st.rerun()





# ********************************************************************************************************
# ********************************************************************************************************
# --- PESTAÑA: VISUALIZAR Y FILTRAR ---
# ********************************************************************************************************
# ********************************************************************************************************
with tab_view: 
    st.write("")
    st.markdown("<h3 style='text-align: center; color:darkgreen ;'>TABLA DE DATOS DE PREPARACIÓN DE ESMALTES FUERTES </h3>", unsafe_allow_html=True)
    # Cargar datos actualizados
    df_visualizar = cargar_datos2()  # <--- Asegúrate de usar el nombre correcto de tu función (cargar_datos2)
    # Columnas que se mostrarán en la pestaña Visualizar
    columnas_mostrar = [
        "ID",
        "BATCH_PREP",
        "COLOR_PREP",
        "TIPO_ESMALTE_PREP",
        #"ALTURA",
        "VOLUMEN(L)_PREP_FUER",
        "KG_HUMEDOS_PREP_FUER",
        "DENSIDAD(KG/L)_PREP_FUER",
        "KG_SECOS_PREP",
        "KG_VIRGEN_PREP_FUER",
        "%_RECUPERADOS_PREP_FUER",
        "KG_RECUPERADOS_PREP_FUER",
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
            key="btn_descarga_todos",
            type="primary"
        )



        # --- TABLA ACUMULADO POR BATCH Y COLOR ---
        st.write("")
        st.markdown("<h3 style='text-align: center; color:darkgreen ;'>TABLA PREPARACIÓN COLORES FUERTES POR BATCH Y COLOR </h3>", unsafe_allow_html=True)
        # 1. Agrupamos por BATCH y por COLOR pasándolos como una lista en el groupby
        df_acumulado_total = df_visualizar.groupby(["BATCH_PREP", "COLOR_PREP"])[["KG_SECOS_PREP", "KG_VIRGEN_PREP_FUER", "KG_RECUPERADOS_PREP_FUER"]].sum().reset_index()
        # 2. Renombramos las columnas en un solo paso para que se vean profesionales
        df_acumulado_total = df_acumulado_total.rename(columns={
            "BATCH_PREP": "BATCH_PREP",
            "COLOR_PREP": "COLOR_PREP",
            "KG_SECOS_PREP": "TOTAL KG SECOS PREP FUERTES",
            "KG_VIRGEN_PREP_FUER": "TOTAL KG VÍRGEN PREP FUERTES",
            "KG_RECUPERADOS_PREP_FUER": "TOTAL KG RECUPERADOS PREP FUERTES"
        })
        # 3. Mostramos la tabla resumen final
        st.dataframe(
            df_acumulado_total.style.format(precision=0),
            use_container_width=True,
            hide_index=True
        )



        # --- TABLA DE TOTALES GENERALES ACUMULADOS --
        st.write("")
        st.markdown("<h3 style='text-align: center; color:purple ;'>✅ TOTALES GENERALES ACUMULADOS </h3>", unsafe_allow_html=True)
        # 1. Sumamos directamente todo el contenido de cada columna cuidando que existan en el CSV
        total_kg_secos = df_visualizar["KG_SECOS_PREP"].sum() if "KG_SECOS_PREP" in df_visualizar.columns else 0.0
        total_kg_virgen = df_visualizar["KG_VIRGEN_PREP_FUER"].sum() if "KG_VIRGEN_PREP_FUER" in df_visualizar.columns else 0.0
        total_kg_recuperados = df_visualizar["KG_RECUPERADOS_PREP_FUER"].sum() if "KG_RECUPERADOS_PREP_FUER" in df_visualizar.columns else 0.0  # Calculamos la parte recuperada como diferencia
        # 2. Creamos el nuevo DataFrame con una sola fila para mostrarlo en formato de tabla
        df_totales_generales = pd.DataFrame([{
            "TOTAL KG_SECOS": total_kg_secos,
            "TOTAL KG_VIRGEN": total_kg_virgen,
            "TOTAL KG_RECUPERADOS": total_kg_recuperados
        }])
        # 3. Mostramos la tabla en Streamlit
        st.dataframe(
            df_totales_generales.style.format(precision=0),
            use_container_width=True,
            hide_index=True
        )
        # --- BOTÓN DE DESCARGA: TOTALES GENERALES ---
        buffer_totales = io.BytesIO()
        with pd.ExcelWriter(buffer_totales, engine='openpyxl') as writer:
            df_totales_generales.to_excel(writer, index=False, sheet_name='Totales Generales')
            
        st.download_button(
            label="📥 Descargar Totales Generales en Excel",
            data=buffer_totales.getvalue(),
            file_name='totales_generales_esmaltes.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key="btn_descarga_totales",
            type="primary"
        )



        # --- TABLA DE FILTRO DINAMICO ---
        st.write("")
        st.write("")
        st.markdown("<h3 style='text-align: center; color:darkorange ;'>🔍 TABLA DE DATOS FILTRADA </h3>", unsafe_allow_html=True)
        # 2. SELECTBOX PARA ELEGIR LA COLUMNA
        columnas_disponibles = df_visualizar.columns.tolist()
        columna_seleccionada = st.selectbox(
            "Selecciona la columna por la que deseas filtrar:", 
            options=columnas_disponibles
        )
        # --- NUEVO: OBTENER VALORES ÚNICOS Y ORDENADOS SIN REPETIR ---
        # Eliminamos valores nulos (dropna) y obtenemos los registros únicos (.unique())
        valores_unicos = df_visualizar[columna_seleccionada].dropna().unique().tolist()
        valores_unicos.sort()  # Los ordena alfabéticamente o numéricamente para que sea más fácil buscar
        # 3. SEGUNDO SELECTBOX EN LUGAR DE TEXT_INPUT
        busqueda_seleccionada = st.selectbox(
            f"Selecciona el valor que deseas buscar en '{columna_seleccionada}':",
            options=valores_unicos,
            placeholder="Selecciona una opción...",
            index=0  # Por defecto selecciona el primer valor único
        )
        # 4. LÓGICA DE FILTRADO (COMPROBACIÓN ROBUSTA)
        df_filtrado = df_visualizar.copy()
        # Nos aseguramos de que haya algo seleccionado en el segundo selectbox
        if busqueda_seleccionada is not None and busqueda_seleccionada != "":
            # Convertimos temporalmente a string ambos lados para evitar conflictos de tipo (int vs float vs str)
            # y usamos una comparación exacta limpia.
            df_filtrado = df_visualizar[
                df_visualizar[columna_seleccionada].astype(str).str.strip() == str(busqueda_seleccionada).strip()
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