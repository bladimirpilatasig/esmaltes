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
st.set_page_config(page_title="PREP. ESMALTES PASTELES", layout="wide")
#Nombre del Archivo
DB_FILE = "0.1_prep_esm_past.csv"
# ESto es para el selectbox
OPCIONES_COLOR_PREPARADO = ["BLANCO", "BONE", "VERDE PRIMAVERA", "CARIBBEAN SHELL", "GRIS", "CELESTE"]
OPCIONES_TIPO_ESMALTE_PREPARADO = ["FERRUM", "EB16", "EB10 ACCESORIOS", "PRUEBAS"," "]
# ---------------------------------------------------------
# CAMPOS DE LA BASE DE DATOS: 
# Si quieres agregar campos de cálculo (ej. 'DIFERENCIA_HORAS'), añádelos a esta lista.
COLUMNAS_BASE_PREPARADO = ["BATCH_PREP","TIPO_ESMALTE_PREP","COLOR_PREP","ALTURA_PREP","DENSIDAD(KG/L)_PREP","VOLUMEN(L)_PREP","KG_HUMEDOS_PREP","KG_SECOS_PREP"]
#COLUMNAS_BASE_PREPARADO = ["ID", "BATCH","COLOR","TIPO_DE_ESMALTE", "ALTURA", "DENSIDAD", "VOLUMEN(L)", "KG HUMEDOS", "KG SECOS"]
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
def cargar_datos1():
    """Lee el CSV local. Si no existe, crea uno nuevo con las columnas base."""
    if not os.path.exists(DB_FILE):
        return pd.DataFrame(columns=COLUMNAS_BASE_PREPARADO)
    df = pd.read_csv(DB_FILE)
    # Aseguramos que el ID sea entero para evitar decimales innecesarios
    df["ID"] = pd.to_numeric(df["ID"], errors='coerce').fillna(0).astype(int)
    return df
def guardar_datos1(df):
    """Guarda el DataFrame y limpia el caché para refrescar la vista."""
    df.to_csv(DB_FILE, index=False)
    st.cache_data.clear()
def generar_excel1(df):
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
st.markdown("<h1 style='text-align: center; color:darkblue ;'>📊 PREPARACIÓN DE ESMALTES PASTELES</h1>", unsafe_allow_html=True)
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
    df = cargar_datos1()
    # GENERADOR DEL ID
    id_actual = 1 if df.empty else int(df["ID"].max()) + 1
    st.subheader(f"Registro No. {id_actual}")
    # --- FORMULARIO DE INGRESO DE DATOS ---
    # BATCH
    batch_texto = st.text_input('BATCH (No.)', value="", placeholder="Ingresa el número de BATCH")
    try:
        batch_preparadop_in = float(batch_texto.replace(',', '.')) if batch_texto else 0.000
    except ValueError:
        batch_preparadop_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en el BATCH.")
    # COLOR
    color_preparadop_in = st.selectbox("COLOR", OPCIONES_COLOR_PREPARADO)
    # TIPO DE ESMALTE
    tipo_esmalte_preparadop_in = st.selectbox("TIPO DE ESMALTE", OPCIONES_TIPO_ESMALTE_PREPARADO)
    # ALTURA DE LA MEDICIÓN
    altura_texto = st.text_input('ALTURA (cm)', value="", placeholder="Ingresa la altura en centímetros")
    try:
        altura_preparadop_in = float(altura_texto.replace(',', '.')) if altura_texto else 0.000
    except ValueError:
        altura_preparadop_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en la ALTURA.")
    # DENSIDAD DE LA MEDICIÓN
    densidad_texto = st.text_input('DENSIDAD (Kg/lt)', value="", placeholder="Ingresa la densidad en Kg/lt")
    try:
        densidad_preparadop_in = float(densidad_texto.replace(',', '.')) if densidad_texto else 0.000
    except ValueError:
        densidad_preparadop_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en la DENSIDAD.")
    # --- BLOQUE DE CÁLCULOS EN TIEMPO REAL ---
    pi = 3.141592653589793
    diametro = 1.925
    altura_calc = 1.44 - (altura_preparadop_in / 100)      # Altura total - altura medida en metros
    area_calc = (pi * (diametro ** 2)) / 4       # Área de la base del cilindro
    volumen_calc = area_calc * altura_calc * 1000 # Volumen del cilindro (convertido a litros)
    volumen_fij0 = 240
    correccion = -32
    volumen_preparadop_final = volumen_calc + volumen_fij0 + correccion # Ajuste de volumen según la corrección
    kg_humedos_preparadop_cal = volumen_preparadop_final * densidad_preparadop_in  # Peso final en Kg
    pendiente = 1.5419 # ecuacion de esmaltes: pendiente = 1.5419
    ordenada_origen = -1.5435 # ecuacion de esmaltes ordenada_origen = -1.5435
    kg_preparadop_secos = volumen_preparadop_final * ((pendiente * densidad_preparadop_in) + ordenada_origen)  # Peso seco en Kg, redondeado a 2 decimales



    # --- MOSTRAR RESULTADOS EN PANTALLA ---
    st.info(f"📐 **Volumen final:** {volumen_preparadop_final:.2f} litros")
    st.info(f"📐 **Kilogramos Húmedos (Kg):** {kg_humedos_preparadop_cal:.2f}")
    st.info(f"📐 **Kilogramos Secos (Kg):** {kg_preparadop_secos:.2f}")



    # Inicializar la variable de estado para la confirmación si no existe
    if "confirmar_guardado" not in st.session_state:
        st.session_state.confirmar_guardado = False
    # Botón principal para iniciar el proceso de guardado
    btn_save = st.button("Guardar Registro", disabled=st.session_state.confirmar_guardado)
    if btn_save:
        # FILTRO DE VALIDACION DE DATOS PREVIO
        if not batch_preparadop_in or not color_preparadop_in or not altura_preparadop_in or not densidad_preparadop_in:
            st.error("⚠️ Error: Por favor completa los campos de Batch, Color, Altura y Densidad.")
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
            # Construcción del registro
            nuevo_item = {
                "ID": id_actual,
                "BATCH_PREP": batch_preparadop_in,
                "TIPO_ESMALTE_PREP": tipo_esmalte_preparadop_in,
                "COLOR_PREP": color_preparadop_in,
                "ALTURA_PREP": altura_preparadop_in,
                "DENSIDAD(KG/L)_PREP": densidad_preparadop_in,
                "VOLUMEN(L)_PREP": volumen_preparadop_final,
                "KG_HUMEDOS_PREP": kg_humedos_preparadop_cal,
                "KG_SECOS_PREP": kg_preparadop_secos
            }
            # Guardado
            df_final = pd.concat([df, pd.DataFrame([nuevo_item])], ignore_index=True)
            guardar_datos1(df_final)
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
    st.markdown("<h3 style='text-align: center; color:darkgreen ;'>TABLA DE DATOS DE PREPARACION DE ESMALTES PASTELES </h3>", unsafe_allow_html=True)
    # Cargar datos actualizados
    df_visualizar = cargar_datos1()
    # Columnas que se mostrarán en la pestaña Visualizar
    columnas_mostrar = [
        "ID",
        "TIPO_ESMALTE_PREP",
        "COLOR_PREP",
        "BATCH_PREP",
        "ALTURA_PREP",
        "VOLUMEN(L)_PREP",
        "KG_HUMEDOS_PREP",
        "DENSIDAD(KG/L)_PREP",
        "KG_SECOS_PREP"
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
            key="btn_descarga_todos"
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