import streamlit as st
import pandas as pd

# Configuración de la página con estética ejecutiva y minimalista
st.set_page_config(page_title="Portal de Modulaciones", layout="wide")

# Inyección de CSS para estilizar los selectores
st.markdown("""
    <style>
    div.row-widget.stRadio > div {
        flex-direction: row;
        background-color: #f1f3f5;
        padding: 10px 15px;
        border-radius: 6px;
        border: 1px solid #ced4da;
    }
    </style>
""", unsafe_allow_html=True)

st.title("Portal de Modulaciones y Seguimiento de Entregas")
st.markdown("Herramienta automatizada para el cruce de pedidos, tiempos y monitoreo de estados de entrega.")
st.markdown("---")

def cargar_datos(file):
    if file.name.endswith('.xlsx'):
        return pd.read_excel(file)
    else:
        return pd.read_csv(file, sep=None, engine='python')

def color_status(val):
    if pd.isna(val):
        return ''
    val_str = str(val).upper()
    
    if 'IN_TREATMENT' in val_str or 'RESCHEDULED' in val_str:
        return 'background-color: #D32F2F; color: white; font-weight: bold;'
    elif 'NOT_STARTED' in val_str:
        return 'background-color: #757575; color: white; font-weight: bold;'
    elif 'CONCLUDED' in val_str:
        return 'background-color: #2E7D32; color: white; font-weight: bold;'
    return ''

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

        # Identificación automática de columnas clave
        cols_1 = df_clientes.columns.tolist()
        col_pdv = next((c for c in cols_1 if str(c).upper() == 'PDV'), cols_1[1] if len(cols_1) > 1 else cols_1[0])

        cols_2 = df_entregas.columns.tolist()
        col_cruce_f = next((c for c in cols_2 if 'poc_exter' in str(c).lower()), cols_2[5] if len(cols_2) > 5 else cols_2[0])
        col_datos_d = next((c for c in cols_2 if 'driver' in str(c).lower()), cols_2[3] if len(cols_2) > 3 else cols_2[0])
        col_status_i = next((c for c in cols_2 if 'status' in str(c).lower()), cols_2[8] if len(cols_2) > 8 else cols_2[0])
        col_motivo_x = next((c for c in cols_2 if 'reason' in str(c).lower()), cols_2[23] if len(cols_2) > 23 else cols_2[-1])
        
        # Identificación de columnas de tiempos
        col_arrived = next((c for c in cols_2 if 'arrived' in str(c).lower()), cols_2[20] if len(cols_2) > 20 else cols_2[-1])
        col_finished = next((c for c in cols_2 if 'finished' in str(c).lower()), cols_2[21] if len(cols_2) > 21 else cols_2[-1])

        with st.expander("Verificar Parámetros de Cruce"):
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1: col_cruce_1 = st.selectbox("Identificador (Base 1):", cols_1, index=cols_1.index(col_pdv))
            with c2: col_cruce_2 = st.selectbox("Buscar en (Base 2):", cols_2, index=cols_2.index(col_cruce_f))
            with c3: col_mostrar_d = st.selectbox("Conductor:", cols_2, index=cols_2.index(col_datos_d))
            with c4: col_arrived_sel = st.selectbox("Llegada (arrived_at):", cols_2, index=cols_2.index(col_arrived))
            with c5: col_finished_sel = st.selectbox("Salida (finished_at):", cols_2, index=cols_2.index(col_finished))

        # Estandarización
        df_clientes[col_cruce_1] = df_clientes[col_cruce_1].astype(str).str.strip().str.replace('.0', '', regex=False)
        df_entregas[col_cruce_2] = df_entregas[col_cruce_2].astype(str).str.strip().str.replace('.0', '', regex=False)

        # Reducción de la base de despachos para evitar duplicidad
        df_entregas_subset = df_entregas[[col_cruce_2, col_mostrar_d, col_status_i, col_motivo_x, col_arrived_sel, col_finished_sel]].drop_duplicates(subset=[col_cruce_2])

        # Cruce
        df_resultado = pd.merge(
            df_clientes[[col_cruce_1]], 
            df_entregas_subset, 
            left_on=col_cruce_1, 
            right_on=col_cruce_2, 
            how='left'
        )

        # Separar camión
        if col_mostrar_d in df_resultado.columns:
            df_resultado[col_mostrar_d] = df_resultado[col_mostrar_d].fillna("SIN DATOS")
            split_data = df_resultado[col_mostrar_d].astype(str).str.split('-', expand=True)
            if split_data.shape[1] > 1:
                df_resultado['Camion'] = split_data[1].str.strip().str.upper()
            else:
                df_resultado['Camion'] = split_data[0].str.strip().str.upper()
        else:
            df_resultado['Camion'] = "NO ENCONTRADO"

        # Cálculo de Tiempos y Extracción de Hora
        df_resultado[col_arrived_sel] = pd.to_datetime(df_resultado[col_arrived_sel], errors='coerce')
        df_resultado[col_finished_sel] = pd.to_datetime(df_resultado[col_finished_sel], errors='coerce')
        
        # Extraer solo la hora en formato HH:MM:SS
        df_resultado['Hora_Arribo'] = df_resultado[col_arrived_sel].dt.strftime('%H:%M:%S').fillna("-")

        # Diferencia en minutos
        df_resultado['Tiempo_Entrega_Min'] = (df_resultado[col_finished_sel] - df_resultado[col_arrived_sel]).dt.total_seconds() / 60
        df_resultado['Tiempo_Entrega_Min'] = df_resultado['Tiempo_Entrega_Min'].round(2)
        df_resultado['Tiempo_Entrega_Min'] = df_resultado['Tiempo_Entrega_Min'].fillna("-")

        # Configuración de vista final
        vista_final = df_resultado[[col_cruce_1, 'Camion', col_status_i, 'Hora_Arribo', 'Tiempo_Entrega_Min', col_motivo_x]].rename(columns={
            col_cruce_1: 'PDV',
            col_status_i: 'Status',
            col_motivo_x: 'Motivo'
        })
        
        vista_final['Status'] = vista_final['Status'].fillna("SIN REGISTRO")
        
        # Lógica de Motivo
        condicion_motivo = vista_final['Status'].astype(str).str.upper().str.contains('IN_TREATMENT|RESCHEDULED', na=False, regex=True)
        vista_final['Motivo'] = vista_final['Motivo'].where(condicion_motivo, "-")
        vista_final['Motivo'] = vista_final['Motivo'].fillna("SIN DETALLE")

        st.markdown("### Controles de Búsqueda y Filtrado")
        
        col_search, col_filter = st.columns([1, 2])
        with col_search:
            search_pdv = st.text_input("Buscar por Código PDV:")
        with col_filter:
            status_filter = st.radio(
                "Segmentación de Estado:", 
                ["Todos", "CONCLUDED", "IN_TREATMENT / RESCHEDULED", "NOT_STARTED"], 
                horizontal=True
            )

        if search_pdv:
            vista_final = vista_final[vista_final['PDV'].astype(str).str.contains(search_pdv, case=False, na=False)]
            
        if status_filter == "CONCLUDED":
            vista_final = vista_final[vista_final['Status'].astype(str).str.upper().str.contains('CONCLUDED', na=False)]
        elif status_filter == "IN_TREATMENT / RESCHEDULED":
            vista_final = vista_final[vista_final['Status'].astype(str).str.upper().str.contains('IN_TREATMENT|RESCHEDULED', na=False, regex=True)]
        elif status_filter == "NOT_STARTED":
            vista_final = vista_final[vista_final['Status'].astype(str).str.upper().str.contains('NOT_STARTED', na=False)]

        st.markdown("### Consolidado de Modulaciones")
        
        if hasattr(vista_final.style, 'map'):
            styled_df = vista_final.style.map(color_status, subset=['Status'])
        else:
            styled_df = vista_final.style.applymap(color_status, subset=['Status'])
            
        st.dataframe(styled_df, use_container_width=True)

        st.markdown("---")
        
        csv_data = vista_final.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Descargar Reporte Actual",
            data=csv_data,
            file_name="reporte_modulaciones_filtrado.csv",
            mime="text/csv",
            type="primary"
        )
        
    except Exception as e:
        st.error(f"Se presentó un error en el procesamiento de las bases de datos: {e}")
