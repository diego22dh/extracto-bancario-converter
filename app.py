import streamlit as st
import pdfplumber
import pandas as pd
import re
import base64
from io import BytesIO
import tempfile
import os

# Configuración de la página
st.set_page_config(
    page_title="Conversor de Extracto Bancario",
    page_icon="🏦",
    layout="centered"
)

# Título y descripción de la app
st.title("Conversor de Extracto Bancario Galicia")
st.write("Esta aplicación convierte extractos bancarios del Banco Galicia en formato PDF a archivos Excel estructurados.")

def procesar_descripcion(descripcion_raw):
    """
    Procesa la descripción y la divide en descripción y detalle cuando es posible.
    """
    if ' POR ' in descripcion_raw:
        partes = descripcion_raw.strip().split(' POR ', 1)
        return partes[0].strip(), partes[1].strip()
    elif ' DE ' in descripcion_raw:
        partes = descripcion_raw.strip().split(' DE ', 1)
        return partes[0].strip(), partes[1].strip()
    else:
        return descripcion_raw.strip(), ""

def limpiar_valor_numerico(valor):
    """
    Limpia y convierte valores numéricos del formato argentino al formato numérico de Python.
    Mantiene el signo negativo.
    """
    if not valor:
        return 0.0
    
    # Preservar el signo negativo
    es_negativo = valor.startswith('-')
    valor_sin_signo = valor.replace('-', '')
    
    # Reemplazar separadores
    valor_limpio = valor_sin_signo.replace('.', '').replace(',', '.')
    
    # Convertir a float y aplicar signo si es necesario
    try:
        resultado = float(valor_limpio)
        return -resultado if es_negativo else resultado
    except ValueError:
        st.warning(f"Error al convertir valor: {valor}")
        return 0.0

def extraer_movimientos_del_pdf(pdf_file):
    """
    Extrae los movimientos bancarios de un PDF del Banco Galicia.
    Procesa todas las páginas del PDF.
    """
    movimientos = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            
            # Procesamos el texto de cada página
            lineas = texto.split('\n')
            inicio_movimientos = False
            
            for linea in lineas:
                # Detectamos el inicio de la tabla de movimientos
                if "Fecha" in linea and "Descripción" in linea and ("Crédito" in linea or "Débito" in linea) and "Saldo" in linea:
                    inicio_movimientos = True
                    continue
                
                # Si ya estamos en la sección de movimientos
                if inicio_movimientos:
                    # Intentamos con un patrón más flexible para diversos formatos
                    patron_fecha = r"(\d{2}/\d{2}/\d{2})"
                    
                    # Si la línea comienza con una fecha, es un nuevo movimiento
                    if re.match(patron_fecha, linea):
                        # Intentamos usar diferentes patrones para capturar los formatos
                        # Patrón más completo
                        patron1 = r"(\d{2}/\d{2}/\d{2})\s+(.*?)(?:\s+(\w+))?\s+(?:(\d+[\.,]?\d*\.?\d*,\d+)?\s+)?(?:(-\d+[\.,]?\d*\.?\d*,\d+)?\s+)?(\d+[\.,]?\d*\.?\d*,\d+)$"
                        # Patrón alternativo para casos específicos
                        patron2 = r"(\d{2}/\d{2}/\d{2})\s+(.*?)\s+(\d+[\.,]?\d*\.?\d*,\d+|\-\d+[\.,]?\d*\.?\d*,\d+)\s+(\d+[\.,]?\d*\.?\d*,\d+)$"
                        
                        match = re.search(patron1, linea)
                        
                        if match:
                            fecha, descripcion_raw, origen, credito, debito, saldo = match.groups()
                            
                            # Procesamos la descripción
                            descripcion, detalle = procesar_descripcion(descripcion_raw)
                            
                            # Determinamos tipo de movimiento e importe
                            if credito and credito.strip():
                                importe = limpiar_valor_numerico(credito)
                                tipo_movimiento = "Credito"
                            elif debito and debito.strip():
                                # Mantenemos el signo negativo en el importe para débitos
                                importe = limpiar_valor_numerico(debito)  # Ya mantiene el signo negativo
                                tipo_movimiento = "Debito"
                            else:
                                # Intentar un enfoque alternativo para casos especiales
                                match_alt = re.search(patron2, linea)
                                if match_alt:
                                    fecha, descripcion_raw, importe_str, saldo = match_alt.groups()
                                    descripcion, detalle = procesar_descripcion(descripcion_raw)
                                    
                                    # Procesamos el importe manteniendo el signo
                                    importe = limpiar_valor_numerico(importe_str)
                                    tipo_movimiento = "Debito" if importe < 0 else "Credito"
                                else:
                                    importe = 0.0
                                    tipo_movimiento = "Desconocido"
                            
                            # Procesamos el saldo
                            saldo_valor = limpiar_valor_numerico(saldo) if saldo else 0.0
                            
                            # Agregamos el movimiento a la lista
                            movimientos.append({
                                'fecha': fecha,
                                'descripcion': descripcion,
                                'detalle': detalle,
                                'importe': importe if 'importe' in locals() else 0.0,
                                'saldo': saldo_valor,
                                'tipo_movimiento': tipo_movimiento if 'tipo_movimiento' in locals() else "Desconocido"
                            })
    
    return movimientos

def get_table_download_link(df):
    """Genera un enlace para descargar el DataFrame como un archivo Excel"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    b64 = base64.b64encode(processed_data).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="extracto_bancario.xlsx">Descargar archivo Excel</a>'

# Widget para cargar archivo
uploaded_file = st.file_uploader("Carga tu extracto bancario en PDF", type=['pdf'])

if uploaded_file is not None:
    # Mostrar spinner mientras se procesa
    with st.spinner('Procesando el archivo PDF...'):
        try:
            # Guardar el archivo cargado en un archivo temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
                temp_file.write(uploaded_file.getvalue())
                temp_path = temp_file.name
            
            # Extraer los movimientos del PDF
            movimientos = extraer_movimientos_del_pdf(temp_path)
            
            # Eliminar el archivo temporal
            os.unlink(temp_path)
            
            if movimientos:
                # Crear un DataFrame
                df = pd.DataFrame(movimientos)
                
                # Mostrar éxito
                st.success(f'¡Procesamiento completado! Se encontraron {len(movimientos)} movimientos.')
                
                # Mostrar los datos en una tabla
                st.subheader("Vista previa de los datos extraídos:")
                st.dataframe(df)
                
                # Proporcionar el enlace de descarga
                st.markdown(get_table_download_link(df), unsafe_allow_html=True)
            else:
                st.warning("No se encontraron movimientos en el PDF. Verifica que sea un extracto bancario del Banco Galicia.")
        except Exception as e:
            st.error(f"Error al procesar el archivo: {str(e)}")

# Información adicional
st.markdown("---")
st.markdown("""
### Información
- Esta aplicación está diseñada específicamente para procesar extractos bancarios del Banco Galicia.
- El archivo resultante tendrá las siguientes columnas: fecha, descripción, detalle, importe, saldo y tipo de movimiento.
- Los valores de débito mantienen su signo negativo para facilitar los cálculos.
""")

# Footer
st.markdown("---")
st.markdown("Desarrollado con Streamlit")
