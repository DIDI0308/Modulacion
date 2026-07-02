import streamlit as st
import pandas as pd
from datetime import datetime
import re

# Configuración de la página con estética ejecutiva
st.set_page_config(page_title="Portal de Modulaciones", layout="wide")

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
st.markdown("Herramienta automatizada para el cruce de pedidos, tiempos, ventanas horarias y estados de entrega.")
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

# Función para estandarizar textos como "ANTES DE 7" a formato hora
def estandarizar_ventana(texto):
    if pd.isna(texto):
        return "SIN ASIGNAR"
    
    t = str(texto).upper().strip()
    
    # Tratamiento de "ANTES DE"
    if "ANTES DE" in t:
        num = ''.join(filter(str.isdigit, t))
        if num:
            return f"00:00 - {num.zfill(2)}:00"
            
    # Tratamiento de "DESPUES DE"
    if "DESPUÉS DE" in t or "DESPUES DE" in t:
        num = ''.join(filter(str.isdigit, t))
        if num:
            return f"{num.zfill(2)}:00 - 23:59"
            
    # Tratamiento AM/PM básico y limpieza de espacios
    t = t.replace(' AM', '').replace(' PM', '').replace('AM', '').replace('PM', '')
    return t

# Inicializar variable en caché para la base de ventanas
if 'df_ventanas' not in st.session_state:
    st.session_state['df_ventanas'] = None

# Creación de Pestañas (Tabs)
tab_principal, tab_drive = st.tabs(["📊 Portal de Cruce", "☁️ Base Ventanas (Google Drive)"])

# ==========================================
# PESTAÑA 2: GOOGLE DRIVE (Extracción Específica)
# ==========================================
with tab_drive:
    st.subheader("Sincronización de Ventanas Horarias")
    st.markdown("Haz clic para extraer PDV y Ventanas de 'VH FIJAS LPZ' (Col B y J) y 'VHs FIJAS EA' (Col A y F).")
    
    if st.button("🔄 Actualizar Base desde Drive", type="primary"):
        with st.spinner("Leyendo documento, esto tomará unos segundos..."):
            try:
                sheet_id = "1OYlT2SVGqxM-C6h27GSrBEcjBOyCixwP"
                url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
                xls = pd.ExcelFile(url)
                
                df_vh_list = []
                
                # Extracción Hoja 1: VH FIJAS LPZ (PDV en B[1], Ventana en J[9])
                if "VH FIJAS LPZ" in xls.sheet_names:
                    df_lpz = pd.read_excel(xls, sheet_name="VH FIJAS LPZ")
                    if len(df_lpz.columns) >= 10:
                        df_temp = df_lpz.iloc[:, [1, 9]].copy()
                        df_temp.columns = ['PDV_Drive', 'Ventana_Tratada']
                        df_vh_list.append(df_temp)
                        
                # Extracción Hoja 2: VHs FIJAS EA (PDV en A[0], Ventana en F[5])
                if "VHs FIJAS EA" in xls.sheet_names:
                    df_ea = pd.read_excel(xls, sheet_name="VHs FIJAS EA")
                    if len(df_ea.columns) >= 6:
                        df_temp = df_ea.iloc[:, [0, 5]].copy()
                        df_temp.columns = ['PDV_Drive', 'Ventana_Tratada']
                        df_vh_list.append(df_temp)
                        
                if df_vh_list:
                    # Unir ambas hojas
                    df_consolidado = pd.concat(df_vh_list, ignore_index=True)
                    df_consolidado.dropna(subset=['PDV_Drive'], inplace=True)
                    
                    # Limpiar PDV de Drive
                    df_consolidado['PDV_Drive'] = df_consolidado['PDV_Drive'].astype(str).str.strip().str.replace('.0', '', regex=False)
                    
                    # Aplicar Tratamiento de Texto (Antes de 7, Después de 16, etc.)
                    df_consolidado['Ventana_Tratada'] = df_consolidado['Ventana_Tratada'].apply(estandarizar_ventana)
                    
                    # Guardar en memoria eliminando duplicados por si acaso
                    df_consolidado = df_consolidado.drop_duplicates(subset=['PDV_Drive'], keep='first')
                    st.session_state['df_ventanas'] = df_consolidado
                    
                    st.success("¡Base actualizada! Se extrajeron y limpiaron las ventanas de LPZ y EA.")
                    st.dataframe(df_consolidado.head(10))
                else:
                    st.error("No se encontraron las hojas solicitadas o su estructura no coincide.")
            except Exception as e:
                st.error(f"Error al conectar con Drive: {e}")

# ==========================================
# PESTAÑA 1: FLUJO PRINCIPAL (Independiente)
# ==========================================
with tab_principal:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. Base Pedidos")
        file_clientes = st.file_uploader("Cargar archivo de Pedidos", type=["xlsx", "csv"], key="clientes")
        tratamiento_clientes = st.radio(
            "Formato de Base Pedidos:", 
            ["Normal", "Requiere Tratamiento"], 
            horizontal=True
        )

    with col2:
        st.subheader("2. Base Despachos")
        file_entregas = st.file_uploader("Cargar archivo de Despachos", type=["xlsx", "csv"], key="entregas")

    if file_clientes and file_entregas:
        try:
            df_clientes = cargar_datos(file_clientes)
            df_entregas = cargar_datos(file_entregas)

            # ---------------------------------------------------------
            # Tratamiento de Base de Clientes (Pedidos)
            # ---------------------------------------------------------
            if tratamiento_clientes == "Requiere Tratamiento":
                if len(df_clientes.columns) == 1:
                    col_name = df_clientes.columns[0]
                    header_parts = col_name.split(',')
                    df_clientes = df_clientes[col_name].astype(str).str.split(',', expand=True)
                    if len(header_parts) == df_clientes.shape[1]:
                        df_clientes.columns = header_parts
                
                col_0 = df_clientes.columns[0]
                split_guion = df_clientes[col_0].astype(str).str.split('-', expand=True)
                
                df_tratado = pd.DataFrame()
                df_tratado['COD'] = split_guion[0].str.strip() if split_guion.shape[1] > 0 else ""
                df_tratado['NOM'] = split_guion[1].str.strip() if split_guion.shape[1] > 1 else ""
                df_tratado['PDV'] = split_guion[2].str.strip() if split_guion.shape[1] > 2 else ""
                
                for c in df_clientes.columns[1:]:
                    df_tratado[c] = df_clientes[c]
                    
                df_clientes = df_tratado
                cols_1 = df_clientes.columns.tolist()
                col_pdv = 'PDV' 
            else:
                cols_1 = df_clientes.columns.tolist()
                col_pdv = next((c for c in cols_1 if str(c).upper() == 'PDV'), cols_1[1] if len(cols_1) > 1 else cols_1[0])

            # ---------------------------------------------------------
            # Identificación de Columnas (Despachos)
            # ---------------------------------------------------------
            cols_2 = df_entregas.columns.tolist()
            col_cruce_f = next((c for c in cols_2 if 'poc_exter' in str(c).lower()), cols_2[5] if len(cols_2) > 5 else cols_2[0])
            col_datos_d = next((c for c in cols_2 if 'driver' in str(c).lower()), cols_2[3] if len(cols_2) > 3 else cols_2[0])
            col_status_i = next((c for c in cols_2 if 'status' in str(c).lower()), cols_2[8] if len(cols_2) > 8 else cols_2[0])
            col_motivo_x = next((c for c in cols_2 if 'reason' in str(c).lower()), cols_2[23] if len(cols_2) > 23 else cols_2[-1])
            col_arrived = next((c for c in cols_2 if 'arrived' in str(c).lower()), cols_2[20] if len(cols_2) > 20 else cols_2[-1])
            col_finished = next((c for c in cols_2 if 'finished' in str(c).lower()), cols_2[21] if len(cols_2) > 21 else cols_2[-1])

            with st.expander("Verificar Parámetros de Cruce"):
                st.markdown("**Bases 1 y 2 (Pedidos y Despachos)**")
                c1, c2, c3, c4, c5 = st.columns(5)
                with c1: col_cruce_1 = st.selectbox("Identificador (Base 1):", cols_1, index=cols_1.index(col_pdv))
                with c2: col_cruce_2 = st.selectbox("Buscar en (Base 2):", cols_2, index=cols_2.index(col_cruce_f))
                with c3: col_mostrar_d = st.selectbox("Conductor:", cols_2, index=cols_2.index(col_datos_d))
                with c4: col_arrived_sel = st.selectbox("Llegada:", cols_2, index=cols_2.index(col_arrived))
                with c5: col_finished_sel = st.selectbox("Salida:", cols_2, index=cols_2.index(col_finished))

            # Estandarización de cruce principal
            df_clientes[col_cruce_1] = df_clientes[col_cruce_1].astype(str).str.strip().str.replace('.0', '', regex=False)
            df_entregas[col_cruce_2] = df_entregas[col_cruce_2].astype(str).str.strip().str.replace('.0', '', regex=False)

            # ---------------------------------------------------------
            # Procesamiento e Inserción de Ventanas (Independiente)
            # ---------------------------------------------------------
            df_entregas_subset = df_entregas[[col_cruce_2, col_mostrar_d, col_status_i, col_motivo_x, col_arrived_sel, col_finished_sel]].drop_duplicates(subset=[col_cruce_2])
            df_resultado = pd.merge(
                df_clientes[[col_cruce_1]], 
                df_entregas_subset, 
                left_on=col_cruce_1, 
                right_on=col_cruce_2, 
                how='left'
            )

            # Lógica Condicional: Si hay ventanas de Drive se cruzan, sino, sigue su camino.
            if st.session_state['df_ventanas'] is not None:
                df_vh = st.session_state['df_ventanas']
                df_resultado = pd.merge(
                    df_resultado,
                    df_vh,
                    left_on=col_cruce_1,
                    right_on='PDV_Drive',
                    how='left'
                )
                col_ventana_view = 'Ventana_Tratada'
                df_resultado[col_ventana_view] = df_resultado[col_ventana_view].fillna("SIN ASIGNAR")
            else:
                st.info("ℹ️ Mostrando tabla sin validación de Ventanas Horarias. (Sincroniza el Drive si las necesitas).")
                df_resultado['Ventana_Tratada'] = "NO CARGADA"
                col_ventana_view = 'Ventana_Tratada'

            # ---------------------------------------------------------
            # Limpieza de Resultados
            # ---------------------------------------------------------
            if col_mostrar_d in df_resultado.columns:
                df_resultado[col_mostrar_d] = df_resultado[col_mostrar_d].fillna("SIN DATOS")
                split_data = df_resultado[col_mostrar_d].astype(str).str.split('-', expand=True)
                if split_data.shape[1] > 1:
                    df_resultado['Camion'] = split_data[1].str.strip().str.upper()
                else:
                    df_resultado['Camion'] = split_data[0].str.strip().str.upper()
            else:
                df_resultado['Camion'] = "NO ENCONTRADO"

            df_resultado[col_arrived_sel] = pd.to_datetime(df_resultado[col_arrived_sel], errors='coerce')
            df_resultado[col_finished_sel] = pd.to_datetime(df_resultado[col_finished_sel], errors='coerce')
            df_resultado['Hora_Arribo'] = df_resultado[col_arrived_sel].dt.strftime('%H:%M:%S').fillna("-")
            df_resultado['Tiempo_Entrega_Min'] = (df_resultado[col_finished_sel] - df_resultado[col_arrived_sel]).dt.total_seconds() / 60
            df_resultado['Tiempo_Entrega_Min'] = df_resultado['Tiempo_Entrega_Min'].round(2).fillna("-")

            # Vista final
            vista_final = df_resultado[[col_cruce_1, 'Camion', col_status_i, col_ventana_view, 'Hora_Arribo', 'Tiempo_Entrega_Min', col_motivo_x]].rename(columns={
                col_cruce_1: 'PDV',
                col_status_i: 'Status',
                col_ventana_view: 'Ventana_Horaria',
                col_motivo_x: 'Motivo'
            })
            
            vista_final['Status'] = vista_final['Status'].fillna("SIN REGISTRO")
            condicion_motivo = vista_final['Status'].astype(str).str.upper().str.contains('IN_TREATMENT|RESCHEDULED', na=False, regex=True)
            vista_final['Motivo'] = vista_final['Motivo'].where(condicion_motivo, "-").fillna("SIN DETALLE")

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
            st.error(f"Se presentó un error en el procesamiento: {e}")
