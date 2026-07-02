import streamlit as st


# Nombre de la pestaña del navegador *************************************************************************************************************************************
# layout="wide" = ocupa todo el espacio de la pantalla
st.set_page_config("ESMALTES", page_icon=":material/home:", layout="wide")


# CREA EL SLIDE BAR *************************************************************************************************************************************
paginas_seleccion = {
    "Contenido: " : [
        #st.Page("a_reg_clientes.py", title="Registro de Clientes", icon="📝"),
        #https://fonts.google.com/icons?selected=Material+Symbols+Outlined:edit_document:FILL@0;wght@400;GRAD@0;opsz@24&icon.size=24&icon.color=%231f1f1f&icon.query=document
        st.Page("b_prep_esm_past.py", title="PREP. ESMALTES PASTELES", icon="📝"),
        st.Page("c_prep_esm_fuer.py", title="PREP. ESMALTES FUERTES", icon="🧮"),
        st.Page("d_verificacion.py", title="VERIFICACIÓN BATCH", icon="📋"),
        st.Page("e_egreso_esm.py", title="EGRESO ESMALTES", icon="📘"),
        st.Page("f_egreso_esm_col.py", title="EGRESOS DE ESMALTES COLOR", icon="📲"),
        #st.Page("f_asistenteUnico.py", title="Asistente IA para un solo Doc", icon="📘"),
        #st.Page("g_asistenteVarios.py", title="Asistente con IA Varios Doc", icon="📚"),
        #st.Page("h_asistenteDatos.py", title="Asistente con IA para Datos", icon="📈"),
        #st.Page("i_analisis.py", title="Analisis de Datos", icon="📊"),
        #st.Page("j_listaAP.py", title="Listado de Alta Presión", icon="✍️"),
        #st.Page("k_llenadasmoldes.py", title="Llenadas de moldes AP", icon="🔢"),
        #st.Page("l_buscador.py", title="Buscador de Datos Sanitarios", icon="🔎"),
        #st.Page("m_datosExcel.py", title="Registro de Datos en Excel", icon="🧮"),
        #st.Page("n_yeseria.py", title="Registro Yeseria", icon="👷‍♂️"),
        #st.Page("o_buscaruc.py", title="Buscar Ruc", icon="🏸")
    ]
}

pg = st.navigation(pages=paginas_seleccion, position="sidebar", expanded=True)
pg.run()
# *************************************************************************************************************************************


