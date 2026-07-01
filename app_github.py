import streamlit as st
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="Portal de Modulaciones - Taiyo", layout="wide")

st.title("Portal de Modulaciones y Seguimiento de Entregas")
st.markdown("Herramienta automatizada para el cruce de pedidos y monitoreo de estados de entrega.")
st.markdown("---")

# Función robusta para lectura de datos (Detecta automáticamente el separador)
def cargar_datos(file):
    if file.name.endswith('.xlsx'):
        return pd.read_excel(file)
    else:
        # El uso de sep=None y engine='python' obliga al sistema a identificar si es coma o punto y coma
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
        # Busca 'PDV' o toma la segunda columna (Columna B / Índice 1)
        col_pdv = next((c for c in cols_1 if str(c).upper() == 'PDV'), cols_1[1] if len(cols_1) > 1 else cols_1[0])

        cols_2 = df_entregas.columns.tolist()
        # Selecciona la sexta columna como clave de búsqueda (Columna F / Índice 5)
        col_cruce_f = cols_2[5] if len(cols_2) > 5 else cols_2[0]
        # Selecciona la cuarta columna para extraer datos (Columna D / Índice 3)
        col_datos_d = cols_2[3] if len(cols_2) > 3 else cols_2[0]

        st.markdown("### Parámetros de Cruce Confirmados")
        col3, col4, col5 = st.columns(3)
        with col3:
            col_cruce_1 = st.selectbox("Identificador (Base 1 - Columna B):", cols_1, index=cols_1.index(col_pdv))
        with col4:
            col_cruce_2 = st.selectbox("Buscar en (Base 2 - Columna F):", cols_2, index=cols_2.index(col_cruce_f))
        with col5:
            col_mostrar_d = st.selectbox("Extraer datos de (Base 2 - Columna D):", cols_2, index=cols_2.index(col_datos_d))

        # Estandarización de las llaves de cruce para asegurar la coincidencia exacta de texto
        df_clientes[col_cruce_1] = df_clientes[col_cruce_1].astype(str).str.strip().str.replace('.0', '', regex=False)
        df_entregas[col_cruce_2] = df_entregas[col_cruce_2].astype(str).str.strip().str.replace('.0', '', regex=False)

        # Reducción de la base de despachos para evitar duplicidad de filas innecesarias en la vista
        df_entregas_subset = df_entregas[[col_cruce_2, col_mostrar_d]].drop_duplicates(subset=[col_cruce_2])

        # 2. Ejecución del Cruce (Left Join)
        df_resultado = pd.merge(
            df_clientes[[col_cruce_1]], 
            df_entregas_subset, 
            left_on=col_cruce_1, 
            right_on=col_cruce_2, 
            how='left'
        )

        st.markdown("### Tabla de Modulaciones Generada")
        
        # Filtrado estricto para mostrar únicamente el PDV de origen y la columna D resultante
        vista_final = df_resultado[[col_cruce_1, col_mostrar_d]]
        st.dataframe(vista_final, use_container_width=True)

        # 3. Exportación de los datos mostrados
        st.markdown("### Exportar Resultados")
        csv_data = vista_final.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Descargar Reporte",
            data=csv_data,
            file_name="reporte_modulaciones_simplificado.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"Se presentó un error en el procesamiento de las bases de datos: {e}")
