import streamlit as st
import pandas as pd
import datetime
import os
import io
import time  # <--- Nueva importación para manejar la pausa del mensaje
from b_prep_esm_past import cargar_datos1
from e_egreso_esm import cargar_datos4


# ********************************************************************************************************
# ********************************************************************************************************
# 1. CONFIGURACIÓN GLOBAL (Layout Wide para usar toda la pantalla)
# ********************************************************************************************************
# ********************************************************************************************************
# layout="wide" permite que la app use el 100% del ancho disponible.
st.set_page_config(page_title="EGRESO DE ESMALTES COLOR", layout="wide")
#Nombre del Archivo
DB_FILE = "0.4_egreso_esm.csv"
# ESto es para el selectbox
OPCIONES_TIPO_ESMALTE_EGRESO = ["FERRUM", "EB16", "EB10 ACCESORIOS", "PRUEBAS"," "]
OPCIONES_COLOR_EGRESO = ["BLANCO", "BONE", "VERDE PRIMAVERA", "CARIBBEAN SHELL", "GRIS", "CELESTE","NEGRO", "CIPRÉS", "AZUL MEDITERRANEO", "MERLOT", "VERDE", "NARANJA","ROJO", "MORA","AZUL SAFIRO"]
# ---------------------------------------------------------
# CAMPOS DE LA BASE DE DATOS: 
# Si quieres agregar campos de cálculo (ej. 'DIFERENCIA_HORAS'), añádelos a esta lista.
COLUMNAS_BASE_EGRESO = ["ID", "FECHA_EGRESO", "BATCH_EGRESO","COLOR_EGRESO", "TIPO_ESMALTE_EGRESO", "NUMERO_CONSUMO_EGRESO", "CANTIDAD_CONSUMO_EGRESO"]
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
# 3. INTERFAZ DE USUARIO PRINCIPAL
# ********************************************************************************************************
# ********************************************************************************************************
st.markdown("<h1 style='text-align: center; color:darkblue ;'>📊 EGRESO DE ESMALTES POR COLORES </h1>", unsafe_allow_html=True)
#st.divider()




# ********************************************************************************************************
# ********************************************************************************************************
# --- PESTAÑA: VISUALIZAR Y FILTRAR ---
# ********************************************************************************************************
# ********************************************************************************************************
# --- TOTAL DE EGRESOS POR COLOR ---
df_egreso_colores = cargar_datos4()
# Agrupación por color, suma del consumo y renombrado de columnas
df_balance_color = df_egreso_colores.groupby(['COLOR_EGRESO'], as_index=False)['CANTIDAD_CONSUMO_EGRESO'].sum()[["COLOR_EGRESO", "CANTIDAD_CONSUMO_EGRESO"]].rename(columns={"COLOR_EGRESO": "COLOR", "CANTIDAD_CONSUMO_EGRESO": "TOTAL CONSUMIDO"})
if not df_balance_color.empty:
    # Mostrar la tabla original (idéntica a image_96ff5b.png)
    st.dataframe(df_balance_color.style.format(precision=0), use_container_width=True, hide_index=True)
    
    # Calcular el total fuera de la tabla
    total_general = df_balance_color["TOTAL CONSUMIDO"].sum()
    
    st.markdown(
    f"""
    <div style="text-align: center; margin-top: 15px; margin-bottom: 15px;">
        <p style="font-size: 16px; color: gray; margin-bottom: 5px;">✅ TOTAL CONSUMO GENERAL (TODOS LOS COLORES)</p>
        <p style="font-size: 32px; font-weight: bold; color: black; margin-top: 0px;">{total_general:.0f}</p>
    </div>
    """, 
    unsafe_allow_html=True
    )
    
    # Generación del Excel y botón de descarga optimizados
    buffer_color = io.BytesIO()
    with pd.ExcelWriter(buffer_color, engine="openpyxl") as writer: 
        df_balance_color.to_excel(writer, index=False, sheet_name="Consumo por Color")
    buffer_color.seek(0)
    
    st.download_button(
        label="📥 Descargar Tabla de Consumos por Color (Excel)", 
        data=buffer_color.getvalue(), 
        file_name="resumen_consumo_por_color.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        type="primary", 
        key="btn_descarga_balance_color"
    )
else:
    st.warning("No hay datos suficientes para calcular los consumos por color.")






# --- TOTAL DE EGRESOS POR FECHA Y COLOR ---
df_egreso_colores = cargar_datos4()
# Agrupación, selección y renombrado en una sola línea limpia
df_balance_final = df_egreso_colores.groupby(['FECHA_EGRESO', 'COLOR_EGRESO'], as_index=False)['CANTIDAD_CONSUMO_EGRESO'].sum()[["FECHA_EGRESO", "COLOR_EGRESO", "CANTIDAD_CONSUMO_EGRESO"]].rename(columns={"FECHA_EGRESO": "FECHA DE EGRESO", "COLOR_EGRESO": "COLOR", "CANTIDAD_CONSUMO_EGRESO": "TOTAL CONSUMIDO"})
# --- NUEVA SECCIÓN: TABLA DE FILTRO DINAMICO ---
st.write("")
st.write("")
st.markdown("<h3 style='text-align: center; color:darkorange ;'>🔍 TABLA DE DATOS FILTRADA</h3>", unsafe_allow_html=True)
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
