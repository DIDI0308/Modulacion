import streamlit as st
import pandas as pd

# Configuración de la página con estética limpia y profesional
st.set_page_config(page_title="Portal de Modulaciones - Taiyo", layout="wide")

st.title("Portal de Modulaciones y Seguimiento de Entregas")
st.markdown("Herramienta automatizada para el cruce de pedidos, asignación de camiones y monitoreo de estados de entrega.")
st.markdown("---")

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
        # Lectura de los archivos cargados
        df_clientes = pd.read_excel(file_clientes) if file_clientes.name.endswith('xlsx') else pd.read_csv(file_clientes)
        df_entregas = pd.read_excel(file_entregas) if file_entregas.name.endswith('xlsx') else pd.read_csv(file_entregas)

        st.markdown("### Configuración de Parámetros de Cruce")
        col3, col4 = st.columns(2)
        
        with col3:
            cols_clientes = df_clientes.columns.tolist()
            # Intenta preseleccionar 'SPV' si existe
            idx_clientes = cols_clientes.index('SPV') if 'SPV' in cols_clientes else 0
            col_cruce_1 = st.selectbox("Columna clave en Base de Pedidos:", cols_clientes, index=idx_clientes)
            
        with col4:
            cols_entregas = df_entregas.columns.tolist()
            # Intenta detectar automáticamente la columna que contenga 'poc_exter'
            posible_col_2 = next((c for c in cols_entregas if 'poc_exter' in str(c).lower()), cols_entregas[0])
            idx_entregas = cols_entregas.index(posible_col_2) if posible_col_2 in cols_entregas else 0
            col_cruce_2 = st.selectbox("Columna clave en Base de Despachos:", cols_entregas, index=idx_entregas)

        # Estandarización estricta de las columnas de unión para evitar fallas por formato de datos
        df_clientes[col_cruce_1] = df_clientes[col_cruce_1].astype(str).str.strip().str.replace('.0', '', regex=False)
        df_entregas[col_cruce_2] = df_entregas[col_cruce_2].astype(str).str.strip().str.replace('.0', '', regex=False)

        # Ejecución de la unión de datos (Left Join)
        df_resultado = pd.merge(
            df_clientes, 
            df_entregas, 
            left_on=col_cruce_1, 
            right_on=col_cruce_2, 
            how='left'
        )
        
        st.markdown("### Datos Consolidados")
        st.dataframe(df_resultado, use_container_width=True)
        
        # Opciones de descarga en formato Excel y CSV
        st.markdown("### Exportar Resultados")
        
        # Generación de descarga en CSV
        csv_data = df_resultado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Descargar Reporte en CSV",
            data=csv_data,
            file_name="reporte_modulaciones.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"Se presentó un error en el procesamiento de las bases de datos: {e}")
        st.info("Por favor, verifique que las columnas seleccionadas correspondan a los códigos de cliente en ambos archivos.")
