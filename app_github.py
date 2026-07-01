import streamlit as st
import pandas as pd

# Configuración de la página con estética limpia y profesional
st.set_page_config(page_title="Portal de Modulaciones - Taiyo", layout="wide")

st.title("Portal de Modulaciones y Seguimiento de Entregas")
st.markdown("Herramienta automatizada para el cruce de pedidos y monitoreo de estados de entrega.")
st.markdown("---")

# Función robusta para lectura de datos (Detecta automáticamente el separador)
def cargar_datos(file):
    if file.name.endswith('.xlsx'):
        return pd.read_excel(file)
    else:
        return pd.read_csv(file, sep=None, engine='python')

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
        # Carga de archivos con la función estandarizada
        df_clientes = cargar_datos(file_clientes)
        df_entregas = cargar_datos(file_entregas)

        # 1. Identificación automática por posición de columnas
        cols_1 = df_clientes.columns.tolist()
        col_pdv = next((c for c in cols_1 if str(c).upper() == 'PDV'), cols_1[1] if len(cols_1) > 1 else cols_1[0])

        cols_2 = df_entregas.columns.tolist()
        # Columna clave para cruce (Columna F/G)
        col_cruce_f = next((c for c in cols_2 if 'poc_exter' in str(c).lower()), cols_2[5] if len(cols_2) > 5 else cols_2[0])
        # Columna para extraer Camión (Columna D)
        col_datos_d = next((c for c in cols_2 if 'driver' in str(c).lower()), cols_2[3] if len(cols_2) > 3 else cols_2[0])
        # Columna para extraer Status (Columna I)
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

        # 2. Ejecución del Cruce (Left Join)
        df_resultado = pd.merge(
            df_clientes[[col_cruce_1]], 
            df_entregas_subset, 
            left_on=col_cruce_1, 
            right_on=col_cruce_2, 
            how='left'
        )

        # 3. Transformación: Separar texto por guion y convertir a mayúsculas
        if col_mostrar_d in df_resultado.columns:
            df_resultado[col_mostrar_d] = df_resultado[col_mostrar_d].fillna("SIN DATOS")
            split_data = df_resultado[col_mostrar_d].astype(str).str.split('-', expand=True)
            
            # Si se encuentra el guion, toma la segunda parte, la limpia y la pasa a mayúsculas
            if split_data.shape[1] > 1:
                df_resultado['Camion'] = split_data[1].str.strip().str.upper()
            else:
                df_resultado['Camion'] = split_data[0].str.strip().str.upper()
        else:
            df_resultado['Camion'] = "NO ENCONTRADO"

        st.markdown("### Tabla de Modulaciones Generada")
        
        # Filtrado estricto para mostrar únicamente el PDV, el Camion y el Status
        vista_final = df_resultado[[col_cruce_1, 'Camion', col_mostrar_status]]
        
        # Renombrar columnas para mayor claridad en el reporte
        vista_final = vista_final.rename(columns={
            col_cruce_1: 'PDV',
            col_mostrar_status: 'Status'
        })
        
        st.dataframe(vista_final, use_container_width=True)

        # 4. Exportación de los datos mostrados
        st.markdown("### Exportar Resultados")
        csv_data = vista_final.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Descargar Reporte",
            data=csv_data,
            file_name="reporte_modulaciones_final.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"Se presentó un error en el procesamiento de las bases de datos: {e}")
