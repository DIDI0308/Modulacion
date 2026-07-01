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

        # 1. Identificación automática de columnas de la Base 1 (Clientes)
        cols_1 = df_clientes.columns.tolist()
        col_pdv = next((c for c in cols_1 if str(c).upper() == 'PDV'), cols_1[1] if len(cols_1) > 1 else cols_1[0])

        # 2. Identificación automática de columnas de la Base 2 (Despachos)
        cols_2 = df_entregas.columns.tolist()
        col_exter = next((c for c in cols_2 if 'poc_exter' in str(c).lower() or 'external' in str(c).lower()), cols_2[6] if len(cols_2) > 6 else cols_2[0])
        col_driver = next((c for c in cols_2 if 'driver' in str(c).lower() or 'chofer' in str(c).lower()), cols_2[3] if len(cols_2) > 3 else cols_2[0])

        st.markdown("### Parámetros de Cruce Confirmados")
        col3, col4, col5 = st.columns(3)
        with col3:
            col_cruce_1 = st.selectbox("Columna Clave (Base 1):", cols_1, index=cols_1.index(col_pdv))
        with col4:
            col_cruce_2 = st.selectbox("Columna Clave (Base 2):", cols_2, index=cols_2.index(col_exter))
        with col5:
            col_driver_sel = st.selectbox("Columna de Conductor (Base 2):", cols_2, index=cols_2.index(col_driver))

        # Estandarización de las llaves de cruce para evitar incompatibilidades de formato
        df_clientes[col_cruce_1] = df_clientes[col_cruce_1].astype(str).str.strip().str.replace('.0', '', regex=False)
        df_entregas[col_cruce_2] = df_entregas[col_cruce_2].astype(str).str.strip().str.replace('.0', '', regex=False)

        # 3. Ejecución del Cruce (Left Join basándose en los PDV de la Base 1)
        df_resultado = pd.merge(
            df_clientes[[col_cruce_1]], 
            df_entregas, 
            left_on=col_cruce_1, 
            right_on=col_cruce_2, 
            how='left'
        )

        # 4. Transformación de Texto: Separar por guión para generar la columna 'camion'
        if col_driver_sel in df_resultado.columns:
            df_resultado[col_driver_sel] = df_resultado[col_driver_sel].fillna("Sin Datos - Sin Camion")
            split_data = df_resultado[col_driver_sel].astype(str).str.split('-', expand=True)
            
            if split_data.shape[1] > 1:
                df_resultado['camion'] = split_data[1].str.strip()
            else:
                df_resultado['camion'] = split_data[0].str.strip()

        st.markdown("### Tabla de Modulaciones Generada")
        
        # Filtrado de columnas asegurando una sola línea limpia de ejecución
        columnas_vista = [col_cruce_1, col_driver_sel, 'camion']
        cols_finales = [c for c in columnas_vista if c in df_resultado.columns]
        
        st.dataframe(df_resultado[cols_finales], use_container_width=True)

        # 5. Exportación de los datos consolidados
        st.markdown("### Exportar Resultados")
        csv_data = df_resultado.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Descargar Reporte en CSV",
            data=csv_data,
            file_name="reporte_modulaciones_final.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"Se presentó un error en el procesamiento de las bases de datos: {e}")
