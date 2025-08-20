# streamlit_app.py
# Este es el archivo principal de tu aplicación Streamlit.

import gspread
import streamlit as st
import pandas as pd
from io import BytesIO

# ==============================================================================
#                      CONFIGURACIÓN DE GOOGLE SHEETS
# ==============================================================================
# La app lee las credenciales directamente desde st.secrets.
# Asegúrate de tener un archivo 'secrets.toml' en tu repositorio.
# ID de tu hoja de cálculo.
SPREADSHEET_ID = '1b_Ud2KcCKmLW3yp3tjrfWLvywieKwp6LclmetIGtsXA'

# ==============================================================================
#                 CONEXIÓN Y LECTURA DE DATOS (Con caché)
# ==============================================================================
@st.cache_resource
def get_google_sheets_client():
    """Establece la conexión con Google Sheets y devuelve el cliente."""
    try:
        # Usa el método from_service_account_from_dict con st.secrets
        client = gspread.service_account_from_dict(st.secrets["gspread"])
        st.success("Conexión con Google Sheets exitosa.")
        return client
    except Exception as e:
        st.error(f"Error de autenticación: {e}")
        st.warning("Por favor, asegúrate de que el archivo secrets.toml esté correctamente configurado.")
        return None

client = get_google_sheets_client()

# ==============================================================================
#                      INTERFAZ DE USUARIO (UI) DE STREAMLIT
# ==============================================================================

st.title("Carga de Datos de Excel a Google Sheets")
st.markdown("---")

if client is None:
    st.warning("No se pudo conectar a Google Sheets. Por favor, revisa tus credenciales.")
else:
    # 1. Componente para subir múltiples archivos de Excel
    uploaded_files = st.file_uploader("Adjunta tus archivos de Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

    # El botón para subir los datos debe estar fuera del bucle
    if st.button("Subir datos a Google Sheets"):
        if not uploaded_files:
            st.warning("Por favor, adjunta al menos un archivo para subir.")
        else:
            st.info(f"Iniciando la carga de {len(uploaded_files)} archivos...")
            total_rows_uploaded = 0
            
            # Accede a la hoja de cálculo
            sheet = client.open_by_key(SPREADSHEET_ID).sheet1

            # ==============================================================================
            #                    NUEVA LÓGICA: CREAR CABECERAS SI NO EXISTEN
            # ==============================================================================
            required_headers = ['nombre cliente', 'fecha', 'REFERENCIA', 'Referencia', 'cantidad', 'serie']
            try:
                # Lee la primera fila para verificar si ya hay cabeceras
                current_headers = sheet.row_values(1)
                
                # Si la primera fila está vacía o no coincide, se crean las cabeceras
                if not current_headers or current_headers != required_headers:
                    sheet.update('A1:F1', [required_headers])
                    st.info("Se crearon las cabeceras de las columnas en tu hoja de cálculo.")
                else:
                    st.info("Las cabeceras ya existen. Se omitió la creación.")
            except Exception as e:
                st.error(f"Error al verificar/crear las cabeceras de la hoja: {e}")
                st.warning("El proceso de carga no se puede completar.")
                st.stop() # Detiene la ejecución si falla la creación de cabeceras

            for uploaded_file in uploaded_files:
                try:
                    # 2. Leer cada archivo de Excel en un DataFrame de Pandas
                    df = pd.read_excel(uploaded_file)
                    st.write(f"Procesando archivo: {uploaded_file.name}")

                    # ==============================================================================
                    #                  FILTRAR DATOS POR COLUMNAS ESPECÍFICAS
                    # ==============================================================================
                    # Selecciona solo las columnas necesarias para subir.
                    df_all_columns = df[['nombre cliente', 'fecha', 'REFERENCIA', 'Referencia', 'cantidad', 'serie']].copy()

                    # Lógica de filtrado: solo se mantienen las filas donde 'Referencia', 'cantidad' y 'serie' no están vacías
                    df_filtered = df_all_columns[
                        df_all_columns['Referencia'].notna() &
                        df_all_columns['cantidad'].notna() &
                        df_all_columns['serie'].notna()
                    ].copy()

                    st.write("Vista previa de los datos que cumplen con las condiciones y se subirán:")
                    st.dataframe(df_filtered)
                    
                    if df_filtered.empty:
                        st.warning(f"No se encontraron filas válidas para subir en el archivo '{uploaded_file.name}'.")
                        continue # Pasa al siguiente archivo

                    # Convierte el DataFrame filtrado a una lista de listas para gspread
                    records_to_upload = df_filtered.values.tolist()
                    
                    # Sube los datos al final de la hoja de cálculo
                    sheet.append_rows(records_to_upload)
                    total_rows_uploaded += len(df_filtered)
                    st.success(f"Datos de '{uploaded_file.name}' subidos exitosamente.")

                except KeyError as e:
                    # Captura el error si una de las columnas no existe en el archivo Excel
                    st.error(f"Error: La columna {e} no se encontró en el archivo '{uploaded_file.name}'.")
                    st.warning("Por favor, verifica que las columnas 'nombre cliente', 'fecha', 'REFERENCIA', 'Referencia', 'cantidad' y 'serie' existan en tu archivo Excel.")
                except Exception as e:
                    st.error(f"Ocurrió un error al procesar el archivo o subir los datos: {e}")
                    st.warning("Por favor, asegúrate de que el archivo Excel tiene el formato correcto.")
            
            # Mensaje final después de procesar todos los archivos
            st.balloons()
            st.success(f"¡Carga completa! Se añadieron {total_rows_uploaded} filas en total a la hoja de cálculo.")

