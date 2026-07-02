import streamlit as st
import pandas as pd
from datetime import datetime

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

# Función de tratamiento de Ventanas Horarias
def procesar_ventanas(df, col_ventana):
    df['Inicio_Ventana'] = pd.to_datetime('23:59:59', format='%H:%M:%S').time()
    df['Fin_Ventana'] = pd.to_datetime('23:59:59', format='%H:%M:%S').time()

    def convertir_a_hora(texto):
        if pd.isna(texto):
            return None, None
        texto = str(texto).upper().strip()
        
        try:
            if '-' in texto:
                partes = texto.split('-')
                inicio_str = partes[0].strip().replace(' AM', '').replace(' PM', '').replace('AM', '').replace('PM', '')
                fin_str = partes[1].strip().replace(' AM', '').replace(' PM', '').replace('AM', '').replace('PM', '')
                
                inicio = pd.to_datetime(inicio_str, format='%H:%M', errors='ignore')
                if isinstance(inicio, str): inicio = pd.to_datetime(inicio_str + ':00', format='%H:%M:%S', errors='coerce')
                
                fin = pd.to_datetime(fin_str, format='%H:%M', errors='ignore')
                if isinstance(fin, str): fin = pd.to_datetime(fin_str + ':00', format='%H:%M:%S', errors='coerce')
                
                return inicio.time() if pd.notna(inicio) else None, fin.time() if pd.notna(fin) else None
                
            elif 'DESPUÉS DE LAS' in texto or 'DESPUES DE LAS' in texto:
                hora = texto.replace('DESPUÉS DE LAS', '').replace('DESPUES DE LAS', '').strip().replace(':00', '').replace(' PM', '').replace(' AM', '')
                inicio = pd.to_datetime(f"{hora}:00:00", format='%H:%M:%S', errors='coerce')
                return inicio.time() if pd.notna(inicio) else None, pd.to_datetime('23:59:59', format='%H:%M:%S').time()
                
            elif 'ANTES DE LAS' in texto:
                hora = texto.replace('ANTES DE LAS', '').strip().replace(':00', '').replace(' PM', '').replace(' AM', '')
                fin = pd.to_datetime(f"{hora}:00:00", format='%H:%M:%S', errors='coerce')
                return pd.to_datetime('00:00:00', format='%H:%M:%S').time(), fin.time() if pd.notna(fin) else None
        except Exception:
            return None, None
            
        return None, None

    for i, fila in df.iterrows():
        inicio, fin = convertir_a_hora(fila[col_ventana])
        if inicio: df.at[i, 'Inicio_Ventana'] = inicio
        if fin: df.at[i, 'Fin_Ventana'] = fin
    
    return df

# Inicializar variable en caché para no perder la descarga al usar filtros
if 'df_ventanas' not in st.session_state:
    st.session_state['df_ventanas'] = None

# Creación de Pestañas (Tabs)
tab_principal, tab_drive = st.tabs(["📊 Portal de Cruce", "☁️ Base Ventanas (Google Drive)"])

# ==========================================
# PESTAÑA 2: GOOGLE DRIVE
# ==========================================
with tab_drive:
    st.subheader("Sincronización de Ventanas Horarias")
    st.markdown("Haz clic en el botón para descargar y consolidar las hojas más recientes directamente desde Google Drive.")
    
    if st.button("🔄 Actualizar Base desde Drive", type="primary"):
        with st.spinner("Leyendo documento, esto tomará unos segundos..."):
            try:
                # El enlace se modifica para forzar una exportación limpia a Excel
                sheet_id = "1OYlT2SVGqxM-C6h27GSrBEcjBOyCixwP"
                url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
                
                xls = pd.ExcelFile(url)
                
                # Orden estricto de hojas solicitado
                hojas_orden = ["VH FIJAS LPZ", "VHs FIJAS EA", "VH FIJA VESPERTINAS LPZ", "BC_LP", "Info", "BC_EA"]
                
                df_list = []
                hojas_leidas = []
                
                for hoja in hojas_orden:
                    if hoja in xls.sheet_names:
                        df_temp = pd.read_excel(xls, sheet_name=hoja)
                        df_list.append(df_temp)
                        hojas_leidas.append(hoja)
                        
                if df_list:
                    # Se consolidan (unen) todas las hojas leídas hacia abajo
                    df_consolidado = pd.concat(df_list, ignore_index=True)
                    st.session_state['df_ventanas'] = df_consolidado
                    st.success(f"¡Base actualizada! Se consolidaron con éxito las siguientes hojas: {', '.join(hojas_leidas)}")
                    st.dataframe(df_consolidado.head(10))
                else:
                    st.error("No se encontró ninguna de las hojas solicitadas en el archivo.")
            except Exception as e:
                st.error(f"Error al conectar con Drive. Revisa que el enlace sea de acceso público. Detalles técnicos: {e}")

# ==========================================
# PESTAÑA 1: FLUJO PRINCIPAL INTACTO
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

    # Verificación: Solo procede si subiste los Excels Y ya actualizaste Drive
    if file_clientes and file_entregas:
        if st.session_state['df_ventanas'] is None:
            st.warning("⚠️ Falta cargar la Base de Ventanas. Ve a la pestaña superior '☁️ Base Ventanas (Google Drive)' y haz clic en Actualizar.")
        else:
            try:
                df_clientes = cargar_datos(file_clientes)
                df_entregas = cargar_datos(file_entregas)
                df_ventanas = st.session_state['df_ventanas'].copy()

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
                # Identificación de Columnas (Despachos y Ventanas)
                # ---------------------------------------------------------
                cols_2 = df_entregas.columns.tolist()
                col_cruce_f = next((c for c in cols_2 if 'poc_exter' in str(c).lower()), cols_2[5] if len(cols_2) > 5 else cols_2[0])
                col_datos_d = next((c for c in cols_2 if 'driver' in str(c).lower()), cols_2[3] if len(cols_2) > 3 else cols_2[0])
                col_status_i = next((c for c in cols_2 if 'status' in str(c).lower()), cols_2[8] if len(cols_2) > 8 else cols_2[0])
                col_motivo_x = next((c for c in cols_2 if 'reason' in str(c).lower()), cols_2[23] if len(cols_2) > 23 else cols_2[-1])
                col_arrived = next((c for c in cols_2 if 'arrived' in str(c).lower()), cols_2[20] if len(cols_2) > 20 else cols_2[-1])
                col_finished = next((c for c in cols_2 if 'finished' in str(c).lower()), cols_2[21] if len(cols_2) > 21 else cols_2[-1])

                cols_3 = df_ventanas.columns.tolist()
                col_cruce_ventanas = next((c for c in cols_3 if 'pdv' in str(c).lower() or 'cod' in str(c).lower()), cols_3[0])
                col_ventana_horaria = next((c for c in cols_3 if 'ventana' in str(c).lower() or 'hora' in str(c).lower()), cols_3[1] if len(cols_3) > 1 else cols_3[0])

                with st.expander("Verificar Parámetros de Cruce"):
                    st.markdown("**Bases 1 y 2 (Pedidos y Despachos)**")
                    c1, c2, c3, c4, c5 = st.columns(5)
                    with c1: col_cruce_1 = st.selectbox("Identificador (Base 1):", cols_1, index=cols_1.index(col_pdv))
                    with c2: col_cruce_2 = st.selectbox("Buscar en (Base 2):", cols_2, index=cols_2.index(col_cruce_f))
                    with c3: col_mostrar_d = st.selectbox("Conductor:", cols_2, index=cols_2.index(col_datos_d))
                    with c4: col_arrived_sel = st.selectbox("Llegada:", cols_2, index=cols_2.index(col_arrived))
                    with c5: col_finished_sel = st.selectbox("Salida:", cols_2, index=cols_2.index(col_finished))
                    
                    st.markdown("**Base 3 (Ventanas Horarias Drive)**")
                    c6, c7 = st.columns(2)
                    with c6: col_cruce_3 = st.selectbox("Identificador (Base 3):", cols_3, index=cols_3.index(col_cruce_ventanas))
                    with c7: col_ventana_sel = st.selectbox("Columna Ventana Horaria:", cols_3, index=cols_3.index(col_ventana_horaria))

                # Estandarización de cruce
                df_clientes[col_cruce_1] = df_clientes[col_cruce_1].astype(str).str.strip().str.replace('.0', '', regex=False)
                df_entregas[col_cruce_2] = df_entregas[col_cruce_2].astype(str).str.strip().str.replace('.0', '', regex=False)
                df_ventanas[col_cruce_3] = df_ventanas[col_cruce_3].astype(str).str.strip().str.replace('.0', '', regex=False)

                # ---------------------------------------------------------
                # Procesamiento y Cruces
                # ---------------------------------------------------------
                # 1. Cruzar Pedidos y Despachos
                df_entregas_subset = df_entregas[[col_cruce_2, col_mostrar_d, col_status_i, col_motivo_x, col_arrived_sel, col_finished_sel]].drop_duplicates(subset=[col_cruce_2])
                df_resultado = pd.merge(
                    df_clientes[[col_cruce_1]], 
                    df_entregas_subset, 
                    left_on=col_cruce_1, 
                    right_on=col_cruce_2, 
                    how='left'
                )

                # 2. Procesar y Cruzar Ventanas Horarias
                df_ventanas = procesar_ventanas(df_ventanas, col_ventana_sel)
                df_ventanas_subset = df_ventanas[[col_cruce_3, col_ventana_sel, 'Inicio_Ventana', 'Fin_Ventana']].drop_duplicates(subset=[col_cruce_3])
                
                df_resultado = pd.merge(
                    df_resultado,
                    df_ventanas_subset,
                    left_on=col_cruce_1,
                    right_on=col_cruce_3,
                    how='left'
                )

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

                # Vista final preliminar
                vista_final = df_resultado[[col_cruce_1, 'Camion', col_status_i, col_ventana_sel, 'Hora_Arribo', 'Tiempo_Entrega_Min', col_motivo_x]].rename(columns={
                    col_cruce_1: 'PDV',
                    col_status_i: 'Status',
                    col_ventana_sel: 'Ventana_Asignada',
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
