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
# 2. LÓGICA DE PERSISTENCIA (Carga y Guardado)
# ********************************************************************************************************
# ********************************************************************************************************
@st.cache_data(show_spinner=False)
def cargar_datos4():
    """Lee el CSV local. Si no existe, crea uno nuevo con las columnas base."""
    if not os.path.exists(DB_FILE):
        return pd.DataFrame(columns=COLUMNAS_BASE_EGRESO)
    df = pd.read_csv(DB_FILE)
    # Aseguramos que el ID sea entero para evitar decimales innecesarios
    df["ID"] = pd.to_numeric(df["ID"], errors='coerce').fillna(0).astype(int)
    return df
def guardar_datos4(df):
    """Guarda el DataFrame y limpia el caché para refrescar la vista."""
    df.to_csv(DB_FILE, index=False)
    st.cache_data.clear()
def generar_excel4(df):
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
st.markdown("<h1 style='text-align: center; color:darkblue ;'>📊 EGRESO DE ESMALTES </h1>", unsafe_allow_html=True)
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
    df = cargar_datos4()
    # GENERADOR DEL ID
    id_actual = 1 if df.empty else int(df["ID"].max()) + 1
    st.subheader(f"Registro No. {id_actual}")



    # --- FORMULARIO DE INGRESO DE DATOS ---
    # FECHA
    fecha_egreso_in = st.date_input("FECHA", value=datetime.date.today(), format="DD/MM/YYYY") # Por defecto toma la fecha del día de hoy
    # BATCH
    batch_egreso_texto = st.text_input('BATCH (No.)', value="", placeholder="Ingresa el número de BATCH")
    try:
        batch_egreso_in = float(batch_egreso_texto.replace(',', '.')) if batch_egreso_texto else 0.000
    except ValueError:
        batch_egreso_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en el BATCH.")
    # COLOR
    color_egreso_in = st.selectbox("COLOR", OPCIONES_COLOR_EGRESO)
    # TIPO DE ESMALTE
    tipo_esmalte_egreso_in = st.selectbox("TIPO DE ESMALTE", OPCIONES_TIPO_ESMALTE_EGRESO)
    # NUMERO DE CONSUMO
    consumo_texto = st.text_input('NÚMERO DE CONSUMO', value="", placeholder="Ingresa el número de consumo")
    try:
        consumo_egreso_in = float(consumo_texto.replace(',', '.')) if consumo_texto else 0.000
    except ValueError:
        consumo_egreso_in = 0
        st.error("⚠️ Por favor, ingresa un número válido en el NÚMERO DE CONSUMO.")
    # CANTIDAD DE CONSUMO
    cantidad_egreso_texto = st.text_input('CANTIDAD DE CONSUMO (litros)', value="", placeholder="Ingresa la cantidad consumida en litros")
    try:
        cantidad_egreso_in = float(cantidad_egreso_texto.replace(',', '.')) if cantidad_egreso_texto else 0.000
    except ValueError:
        cantidad_egreso_in = 0.000
        st.error("⚠️ Por favor, ingresa un número válido en la CANTIDAD DE CONSUMO.")




    # Inicializar la variable de estado para la confirmación si no existe
    if "confirmar_guardado" not in st.session_state:
        st.session_state.confirmar_guardado = False
    # Botón principal para iniciar el proceso de guardado
    btn_save = st.button("Guardar Registro", disabled=st.session_state.confirmar_guardado)
    if btn_save:
        # FILTRO DE VALIDACION DE DATOS PREVIO
        # Validamos que los textos no estén vacíos en vez de evaluar si el número es 0
        if not fecha_egreso_in or not tipo_esmalte_egreso_in or consumo_texto.strip() == "" or cantidad_egreso_texto.strip() == "":
            st.error("⚠️ Error: Por favor completa todos los campos.")
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
                "FECHA_EGRESO": fecha_egreso_in.strftime("%d/%m/%Y"),
                "BATCH_EGRESO": batch_egreso_in,
                "COLOR_EGRESO": color_egreso_in,
                "TIPO_ESMALTE_EGRESO": tipo_esmalte_egreso_in,
                "NUMERO_CONSUMO_EGRESO": consumo_egreso_in,
                "CANTIDAD_CONSUMO_EGRESO": cantidad_egreso_in,
            }
            # Guardado
            df_final = pd.concat([df, pd.DataFrame([nuevo_item])], ignore_index=True)
            guardar_datos4(df_final)
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
    st.write("")
    st.markdown("<h3 style='text-align: center; color:darkgreen ;'>TABLA DE DATOS DE EGRESOS </h3>", unsafe_allow_html=True)
    # Cargar datos actualizados
    df_visualizar = cargar_datos4()
    # Columnas que se mostrarán en la pestaña Visualizar
    columnas_mostrar = [
        "ID",
        "FECHA_EGRESO",
        "BATCH_EGRESO",
        "COLOR_EGRESO",
        "TIPO_ESMALTE_EGRESO",
        "NUMERO_CONSUMO_EGRESO",
        "CANTIDAD_CONSUMO_EGRESO",
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


# --- NUEVA SECCIÓN: TABLA DE TOTALES ACUMULADOS POR BATCH, COLOR Y TIPO ---
        st.write("")
        st.markdown("<h3 style='text-align: center; color:darkgreen ;'>TABLA DE CANTIDAD DE EGRESOS POR BATCH, COLOR Y TIPO </h3>", unsafe_allow_html=True)
        # 1. Agrupamos por BATCH_EGRESO, COLOR_EGRESO y TIPO_ESMALTE_EGRESO pasándolos como lista
        df_acumulado_batch = df_visualizar.groupby(["BATCH_EGRESO", "COLOR_EGRESO", "TIPO_ESMALTE_EGRESO"])["CANTIDAD_CONSUMO_EGRESO"].sum().reset_index()
        # 2. Renombramos las columnas para que la tabla luzca limpia y profesional
        df_acumulado_batch = df_acumulado_batch.rename(columns={
            "BATCH_EGRESO": "BATCH",
            "COLOR_EGRESO": "COLOR",
            "TIPO_ESMALTE_EGRESO": "TIPO DE ESMALTE",
            "CANTIDAD_CONSUMO_EGRESO": "CANTIDAD TOTAL CONSUMIDA (litros)"
        })
        # 3. Mostramos la nueva tabla resumen (Sin decimales)
        st.dataframe(
            df_acumulado_batch.style.format(precision=0),
            use_container_width=True,
            hide_index=True
        )
        # Botón opcional de descarga para la tabla acumulada
        buffer_acumulado = io.BytesIO()
        with pd.ExcelWriter(buffer_acumulado, engine='openpyxl') as writer:
            df_acumulado_batch.to_excel(writer, index=False, sheet_name='Resumen por Batch')
        
        st.download_button(
            label="📥 Descargar Resumen por BATCH en Excel",
            data=buffer_acumulado.getvalue(),
            file_name='resumen_acumulado_batch.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type="primary",
            key="btn_descarga_acumulado"
        )


        # --- COMPARACION DE PREPARADO VS CANTIDAD DE EGRESO (CON TABLA 1 + TABLA 2) ---
        df_prep, df_fuer, df_egresos = cargar_datos1(), cargar_datos2(), cargar_datos4()
        
        # 1. Combinamos Tabla 1 y Tabla 2 en una sola línea continua y agrupamos
        df_prep_grouped = pd.concat([df_prep, df_fuer], ignore_index=True).groupby(["BATCH_PREP", "COLOR_PREP", "TIPO_ESMALTE_PREP"], as_index=False)["KG_SECOS_PREP"].sum()
        df_egresos_grouped = df_egresos.groupby(["BATCH_EGRESO", "COLOR_EGRESO", "TIPO_ESMALTE_EGRESO"], as_index=False)["CANTIDAD_CONSUMO_EGRESO"].sum()
        
        # 2. Cruce Outer Join y unificación/limpieza de nulos (NaN) en un bloque continuo
        df_comparativa = pd.merge(df_prep_grouped, df_egresos_grouped, left_on=["BATCH_PREP", "COLOR_PREP", "TIPO_ESMALTE_PREP"], right_on=["BATCH_EGRESO", "COLOR_EGRESO", "TIPO_ESMALTE_EGRESO"], how="outer")
        df_comparativa["BATCH_FINAL"] = df_comparativa["BATCH_PREP"].fillna(df_comparativa["BATCH_EGRESO"])
        df_comparativa["COLOR_FINAL"] = df_comparativa["COLOR_PREP"].fillna(df_comparativa["COLOR_EGRESO"])
        df_comparativa["TIPO_FINAL"] = df_comparativa["TIPO_ESMALTE_PREP"].fillna(df_comparativa["TIPO_ESMALTE_EGRESO"])
        
        # Unimos la asignación de ceros a las columnas numéricas en una sola línea
        df_comparativa["KG_SECOS_PREP"], df_comparativa["CANTIDAD_CONSUMO_EGRESO"] = df_comparativa["KG_SECOS_PREP"].fillna(0.0), df_comparativa["CANTIDAD_CONSUMO_EGRESO"].fillna(0.0)
        df_comparativa["DIFERENCIA"] = df_comparativa["KG_SECOS_PREP"] - df_comparativa["CANTIDAD_CONSUMO_EGRESO"]
        
        # 3. Creación y filtrado de la tercera tabla resumida en un solo paso
        df_balance_final = df_comparativa[["BATCH_FINAL", "COLOR_FINAL", "TIPO_FINAL", "KG_SECOS_PREP", "CANTIDAD_CONSUMO_EGRESO", "DIFERENCIA"]].rename(
            columns={"BATCH_FINAL": "BATCH", "COLOR_FINAL": "COLOR", "TIPO_FINAL": "TIPO DE ESMALTE", "KG_SECOS_PREP": "KG SECOS (PREPARACIÓN)", "CANTIDAD_CONSUMO_EGRESO": "CANTIDAD CONSUMIDA (EGRESO)", "DIFERENCIA": "DIFERENCIA TOTAL"}
        )
        st.markdown("<h3 style='text-align: center; color:purple ;'>✅ TABLA DE CANTIDAD DE PREPARACION DE ESMALTE VS EGRESOS </h3>", unsafe_allow_html=True)
        if not df_balance_final.empty:
            # Mostramos la tabla formateada estrictamente sin decimales (precision=0)
            st.dataframe(df_balance_final.style.format(precision=0), use_container_width=True, hide_index=True)
            # Generación del Excel y botón de descarga optimizados en espacio
            buffer_balance = io.BytesIO()
            with pd.ExcelWriter(buffer_balance, engine="openpyxl") as writer: df_balance_final.to_excel(writer, index=False, sheet_name="Balance General")
            buffer_balance.seek(0)
            st.download_button(label="📥 Descargar Tabla de Balances (Excel)", data=buffer_balance.getvalue(), file_name="balance_preparacion_vs_egreso.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="btn_descarga_balance_egr")
        else:
            st.warning("No hay datos suficientes para calcular balances.")


        # --- NUEVA SECCIÓN: TABLA DE FILTRO DINAMICO ---
        st.write("")
        st.write("")
        st.markdown("<h3 style='text-align: center; color:darkorange ;'>🔍 TABLA DE DATOS FILTRADA DE PREPARACION VS EGRESOS</h3>", unsafe_allow_html=True)
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
# --- PESTAÑA: EDITAR ---
# ********************************************************************************************************
# ********************************************************************************************************
with tab_edit:
    df_editar = cargar_datos4()
    st.markdown("<h3 style='text-align: center; color:darkblue ;'>✏️ EDITAR REGISTRO EXISTENTE</h3>", unsafe_allow_html=True)
    
    
    if df_editar.empty:
        st.warning("📭 No hay registros guardados para editar.")
    else:
        
        # Selectbox para elegir el ID a editar
        lista_ids_edit = df_editar["ID"].tolist()
        id_seleccionar_edit = st.selectbox("Seleccione el ID del registro que desea editar:", options=lista_ids_edit, key="sb_id_editar")
        
        # Obtener la fila actual del ID seleccionado
        fila_editar = df_editar[df_editar["ID"] == id_seleccionar_edit].iloc[0]
        
        # --- FORMULARIO DE EDICIÓN (Cargado con los datos actuales) ---
        # FECHA: Convertimos de str a objeto datetime.date para el st.date_input
        try:
            fecha_actual_dt = datetime.datetime.strptime(str(fila_editar["FECHA_EGRESO"]), "%d/%m/%Y").date()
        except ValueError:
            fecha_actual_dt = datetime.date.today()
            
        fecha_egreso_edit = st.date_input("FECHA (EDICIÓN)", value=fecha_actual_dt, format="DD/MM/YYYY", key="date_edit")
        
        # BATCH
        batch_egreso_edit_text = st.text_input('BATCH (No.) (EDICIÓN)', value=str(fila_editar["BATCH_EGRESO"]), key="txt_batch_edit")
        try:
            batch_egreso_edit_val = float(batch_egreso_edit_text.replace(',', '.')) if batch_egreso_edit_text else 0.000
        except ValueError:
            batch_egreso_edit_val = 0.000
            st.error("⚠️ Por favor, ingresa un número válido en el BATCH.")
            
        # COLOR: Buscamos el índice actual en tus opciones globales
        color_actual = str(fila_editar["COLOR_EGRESO"])
        idx_color = OPCIONES_COLOR_EGRESO.index(color_actual) if color_actual in OPCIONES_COLOR_EGRESO else 0
        color_egreso_edit = st.selectbox("COLOR (EDICIÓN)", OPCIONES_COLOR_EGRESO, index=idx_color, key="sb_color_edit")
        
        # TIPO DE ESMALTE
        tipo_actual = str(fila_editar["TIPO_ESMALTE_EGRESO"])
        idx_tipo = OPCIONES_TIPO_ESMALTE_EGRESO.index(tipo_actual) if tipo_actual in OPCIONES_TIPO_ESMALTE_EGRESO else 0
        tipo_esmalte_egreso_edit = st.selectbox("TIPO DE ESMALTE (EDICIÓN)", OPCIONES_TIPO_ESMALTE_EGRESO, index=idx_tipo, key="sb_tipo_edit")
        
        # NÚMERO DE CONSUMO
        consumo_edit_text = st.text_input('NÚMERO DE CONSUMO (EDICIÓN)', value=str(fila_editar["NUMERO_CONSUMO_EGRESO"]), key="txt_cons_edit")
        try:
            consumo_egreso_edit_val = float(consumo_edit_text.replace(',', '.')) if consumo_edit_text else 0.000
        except ValueError:
            consumo_egreso_edit_val = 0.000
            st.error("⚠️ Por favor, ingresa un número válido en el NÚMERO DE CONSUMO.")
            
        # CANTIDAD DE CONSUMO
        cantidad_edit_text = st.text_input('CANTIDAD DE CONSUMO (litros) (EDICIÓN)', value=str(fila_editar["CANTIDAD_CONSUMO_EGRESO"]), key="txt_cant_edit")
        try:
            cantidad_egreso_edit_val = float(cantidad_edit_text.replace(',', '.')) if cantidad_edit_text else 0.000
        except ValueError:
            cantidad_egreso_edit_val = 0.000
            st.error("⚠️ Por favor, ingresa un número válido en la CANTIDAD DE CONSUMO.")
            
        # Inicializar estado para confirmación de edición
        if "confirmar_edicion" not in st.session_state:
            st.session_state.confirmar_edicion = False
            
        btn_update = st.button("Actualizar Registro", disabled=st.session_state.confirmar_edicion, key="btn_update_main")
        
        if btn_update:
            if not fecha_egreso_edit or not tipo_esmalte_egreso_edit or consumo_edit_text.strip() == "" or cantidad_edit_text.strip() == "":
                st.error("⚠️ Error: Por favor completa todos los campos para actualizar.")
            else:
                st.session_state.confirmar_edicion = True
                st.rerun()
                
        # --- BLOQUE DE CONFIRMACIÓN DE EDICIÓN ---
        if st.session_state.confirmar_edicion:
            st.warning(f"❓ ¿Está seguro de que desea guardar los cambios en el Registro No. {id_seleccionar_edit}?")
            col_si_ed, col_no_ed = st.columns(2)
            
            with col_si_ed:
                btn_si_ed = st.button("✔️ SÍ, actualizar", use_container_width=True, key="btn_si_ed")
            with col_no_ed:
                btn_no_ed = st.button("❌ NO, cancelar", use_container_width=True, key="btn_no_ed")
                
            if btn_si_ed:
                # Modificamos la fila correspondiente en el DataFrame
                df_editar.loc[df_editar["ID"] == id_seleccionar_edit, "FECHA_EGRESO"] = fecha_egreso_edit.strftime("%d/%m/%Y")
                df_editar.loc[df_editar["ID"] == id_seleccionar_edit, "BATCH_EGRESO"] = batch_egreso_edit_val
                df_editar.loc[df_editar["ID"] == id_seleccionar_edit, "COLOR_EGRESO"] = color_egreso_edit
                df_editar.loc[df_editar["ID"] == id_seleccionar_edit, "TIPO_ESMALTE_EGRESO"] = tipo_esmalte_egreso_edit
                df_editar.loc[df_editar["ID"] == id_seleccionar_edit, "NUMERO_CONSUMO_EGRESO"] = consumo_egreso_edit_val
                df_editar.loc[df_editar["ID"] == id_seleccionar_edit, "CANTIDAD_CONSUMO_EGRESO"] = cantidad_egreso_edit_val
                
                guardar_datos4(df_editar)
                st.success(f"✅ Registro #{id_seleccionar_edit} actualizado exitosamente.")
                st.session_state.confirmar_edicion = False
                time.sleep(2)
                st.rerun()
                
            if btn_no_ed:
                st.session_state.confirmar_edicion = False
                st.rerun()





# ********************************************************************************************************
# ********************************************************************************************************
# --- PESTAÑA: ELIMINAR ---
# ********************************************************************************************************
# ********************************************************************************************************
with tab_del:
    df_eliminar = cargar_datos4()
    st.markdown("<h3 style='text-align: center; color:darkred ;'>🗑️ ELIMINAR REGISTRO</h3>", unsafe_allow_html=True)
    
    if df_eliminar.empty:
        st.warning("📭 No hay registros guardados para eliminar.")
    else:
        
        
        # Selectbox para elegir el ID a borrar
        lista_ids_del = df_eliminar["ID"].tolist()
        id_seleccionar_del = st.selectbox("Seleccione el ID del registro que desea eliminar:", options=lista_ids_del, key="sb_id_eliminar")
        
        # Filtrar y extraer la fila en un DataFrame individual para mostrarlo como tabla previa
        fila_eliminar_df = df_eliminar[df_eliminar["ID"] == id_seleccionar_del]
        
        st.write("📋 **Vista previa del registro seleccionado para eliminar:**")
        # Mostrar el registro en formato tabla limpia
        st.dataframe(
            fila_eliminar_df.style.format(precision=0),
            use_container_width=True,
            hide_index=True
        )
        
        # Inicializar estado para confirmación de eliminación
        if "confirmar_eliminacion" not in st.session_state:
            st.session_state.confirmar_eliminacion = False
            
        btn_delete = st.button("Eliminar Registro", disabled=st.session_state.confirmar_eliminacion, key="btn_del_main")
        
        if btn_delete:
            st.session_state.confirmar_eliminacion = True
            st.rerun()
            
        # --- BLOQUE DE CONFIRMACIÓN DE ELIMINACIÓN ---
        if st.session_state.confirmar_eliminacion:
            st.error(f"🚨 ALERTA: ¿Está completamente seguro de que desea ELIMINAR permanentemente el Registro No. {id_seleccionar_del}?")
            col_si_del, col_no_del = st.columns(2)
            
            with col_si_del:
                btn_si_del = st.button("💥 SÍ, eliminar definitivamente", use_container_width=True, key="btn_si_del")
            with col_no_del:
                btn_no_del = st.button("❌ NO, cancelar", use_container_width=True, key="btn_no_del")
                
            if btn_si_del:
                # Filtrar el DataFrame manteniendo todos los registros menos el ID seleccionado
                df_final_del = df_eliminar[df_eliminar["ID"] != id_seleccionar_del]
                guardar_datos4(df_final_del)
                
                st.success(f"🗑️ Registro #{id_seleccionar_del} eliminado correctamente.")
                st.session_state.confirmar_eliminacion = False
                time.sleep(2)
                st.rerun()
                
            if btn_no_del:
                st.session_state.confirmar_eliminacion = False
                st.rerun()