import streamlit as st
import pandas as pd

# Configuración de la página con estética limpia y profesional
st.set_page_config(page_title="Portal de Modulaciones - Taiyo", layout="wide")

st.title("Portal de Modulaciones y Seguimiento de Entregas")
st.markdown("Herramienta automatizada para el cruce de pedidos y monitoreo de estados de entrega.")
st.markdown("---")

# Función robusta para lectura de datos
def cargar_datos(file):
    if file.name.endswith('.xlsx'):
        return pd.read_excel(file)
    else:
        return pd.read_csv(file, sep=None, engine='python')

# Función para aplicar colores según el estado
def color_status(val):
    if pd.isna(val):
        return ''
    val_str = str(val).upper()
    if 'IN_TREATMENT' in val_str:
        return 'background-color: #f8d7da; color: #721c24;' # Rojo
    elif 'NOT_STARTED' in val_str:
        return 'background-color: #e2e3e5; color: #383d41;' # Gris
    elif 'CONCLUDED' in val_str:
        return 'background-color: #d4edda; color: #155724;' # Verde
    return ''

# Secciones de carga de datos
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Base de Clientes y Pedidos")
    file_clientes = st.file_uploader("Cargar archivo de Pedidos (Excel/CSV)", type=["xlsx", "csv"], key="clientes")

with col2:
    st.subheader("2. Base de Despachos y Estados")
    file_entregas = st.file_uploader("Cargar archivo de Despachos (Excel/CSV)", type=["xlsx", "csv"], key="entregas")

if file_clientes and file_entregas:
    try:
        df_clientes = cargar_datos(file_clientes)
        df_entregas = cargar_datos(file_entregas)

        # Identificación automática por posición de columnas
        cols_1 = df_clientes.columns.tolist()
        col_pdv = next((c for c in cols_1 if str(c).upper() == 'PDV'), cols_1[1] if len(cols_1) > 1 else cols_1[0])

        cols_2 = df_entregas.columns.tolist()
        col_cruce_f = next((c for c in cols_2 if 'poc_exter' in str(c).lower()), cols_2[5] if len(cols_2) > 5 else cols_2[0])
        col_datos_d = next((c for c in cols_2 if 'driver' in str(c).lower()), cols_2[3] if len(cols_2) > 3 else cols_2[0])
        col_status_i = next((c for c in cols_2 if 'status' in str(c).lower()), cols_2[8] if len(cols_2) > 8 else cols_2[0])

        st.markdown("### Parámetros de Cruce Confirmados")
        col3, col4, col5, col6 = st.columns(4)
        with col3:
            col_cruce_1 = st.selectbox("Identificador (Base 1):", cols_1, index=cols_1.index(col_pdv))
        with col4:
            col_cruce_2 = st.selectbox("Buscar en (Base 2):", cols_2, index=cols_2.index(col_cruce_f))
        with col5:
            col_mostrar_d = st.selectbox("Extraer conductor de (Base 2):", cols_2, index=cols_2.index(col_datos_d))
        with col6:
            col_mostrar_status = st.selectbox("Extraer status de (Base 2):", cols_2, index=cols_2.index(col_status_i))

        # Estandarización de las llaves de cruce
        df_clientes[col_cruce_1] = df_clientes[col_cruce_1].astype(str).str.strip().str.replace('.0', '', regex=False)
        df_entregas[col_cruce_2] = df_entregas[col_cruce_2].astype(str).str.strip().str.replace('.0', '', regex=False)

        # Reducción de la base de despachos para evitar duplicidad
        df_entregas_subset = df_entregas[[col_cruce_2, col_mostrar_d, col_mostrar_status]].drop_duplicates(subset=[col_cruce_2])

        # Ejecución del Cruce (Left Join)
        df_resultado = pd.merge(
            df_clientes[[col_cruce_1]], 
            df_entregas_subset, 
            left_on=col_cruce_1, 
            right_on=col_cruce_2, 
            how='left'
        )

        # Transformación: Separar texto por guion y convertir a mayúsculas
        if col_mostrar_d in df_resultado.columns:
            df_resultado[col_mostrar_d] = df_resultado[col_mostrar_d].fillna("SIN DATOS")
            split_data = df_resultado[col_mostrar_d].astype(str).str.split('-', expand=True)
            
            if split_data.shape[1] > 1:
                df_resultado['Camion'] = split_data[1].str.strip().str.upper()
            else:
                df_resultado['Camion'] = split_data[0].str.strip().str.upper()
        else:
            df_resultado['Camion'] = "NO ENCONTRADO"

        # Filtrado inicial para consolidar vista
        vista_final = df_resultado[[col_cruce_1, 'Camion', col_mostrar_status]].rename(columns={
            col_cruce_1: 'PDV',
            col_mostrar_status: 'Status'
        })
        
        # Limpieza de nulos en Status para evitar errores en filtros
        vista_final['Status'] = vista_final['Status'].fillna("SIN REGISTRO")

        st.markdown("---")
        st.markdown("### Controles de Búsqueda y Filtrado")
        
        # Implementación de buscador y filtros rápidos
        col_search, col_filter = st.columns([1, 2])
        
        with col_search:
            search_pdv = st.text_input("Buscar por PDV (Código exacto o parcial):", "")
            
        with col_filter:
            status_filter = st.radio(
                "Filtrar por Estado de Entrega:", 
                ["Todos", "CONCLUDED (Verde)", "IN_TREATMENT (Rojo)", "NOT_STARTED (Gris)"], 
                horizontal=True
            )

        # Aplicación lógica de los filtros
        if search_pdv:
            vista_final = vista_final[vista_final['PDV'].astype(str).str.contains(search_pdv, case=False, na=False)]
            
        if status_filter == "CONCLUDED (Verde)":
            vista_final = vista_final[vista_final['Status'].astype(str).str.upper().str.contains('CONCLUDED', na=False)]
        elif status_filter == "IN_TREATMENT (Rojo)":
            vista_final = vista_final[vista_final['Status'].astype(str).str.upper().str.contains('IN_TREATMENT', na=False)]
        elif status_filter == "NOT_STARTED (Gris)":
            vista_final = vista_final[vista_final['Status'].astype(str).str.upper().str.contains('NOT_STARTED', na=False)]

        st.markdown("### Tabla de Modulaciones")
        
        # Aplicar el formato condicional de color a la tabla
        # Se utiliza hasattr para garantizar compatibilidad con distintas versiones de Pandas
        if hasattr(vista_final.style, 'map'):
            styled_df = vista_final.style.map(color_status, subset=['Status'])
        else:
            styled_df = vista_final.style.applymap(color_status, subset=['Status'])
            
        st.dataframe(styled_df, use_container_width=True)

        # Exportación de los datos filtrados
        st.markdown("### Exportar Resultados")
        csv_data = vista_final.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Descargar Reporte Actual",
            data=csv_data,
            file_name="reporte_modulaciones_filtrado.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"Se presentó un error en el procesamiento de las bases de datos: {e}")
