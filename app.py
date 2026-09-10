import streamlit as st
import pdfplumber
import pandas as pd
import re
import base64
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Conversor de Extracto Bancario",
    page_icon="🏦",
    layout="centered"
)

st.title("Conversor de Extracto Bancario Galicia")
st.write("Esta aplicación convierte extractos bancarios del Banco Galicia (formato Office Banking) en PDF a archivos Excel estructurados.")

# --- Patrones ---
PATRON_FECHA = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(.*)$')
PATRON_MONTO = re.compile(r'([+-])\s*\$\s*([\d\.]+,\d{2})')
PATRON_SALDO = re.compile(r'\$\s*([\d\.]+,\d{2})\s*$')

def limpiar_valor_numerico(valor_str):
    """Convierte un número en formato argentino (1.234.567,89) a float."""
    if not valor_str:
        return 0.0
    valor_limpio = valor_str.strip().replace('.', '').replace(',', '.')
    try:
        return float(valor_limpio)
    except ValueError:
        return 0.0

def es_linea_ruido(linea):
    """Detecta líneas que no son parte de un movimiento (encabezados, pie de página, numeración)."""
    if linea in {"Office Banking", "Galicia"}:
        return True
    if linea.startswith("Fecha de descarga"):
        return True
    # timestamp tipo "02/09/26 - 10:11hs"
    if re.match(r'^\d{2}/\d{2}/\d{2}\s*-\s*\d{2}:\d{2}hs$', linea):
        return True
    # numero de página suelto (1 o 2 dígitos solos)
    if re.match(r'^\d{1,2}$', linea):
        return True
    # fila de encabezado de la tabla
    if "Fecha" in linea and "Descripción" in linea and ("Débito" in linea or "Crédito" in linea) and "Saldo" in linea:
        return True
    return False

def extraer_movimientos_del_pdf(pdf_file):
    """
    Extrae los movimientos bancarios de un extracto Galicia - Office Banking.
    Cada movimiento puede ocupar varias líneas: la primera trae
    fecha + concepto + importe + saldo, y las siguientes traen datos
    adicionales (nombre, CUIT, banco, categoría, etc.) que se agrupan en
    la columna 'detalle'. El flag "dentro de la tabla" persiste entre
    páginas, porque el encabezado de columnas solo aparece una vez.
    """
    movimientos = []
    inicio_movimientos = False

    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            lineas = texto.split('\n')

            for linea_raw in lineas:
                linea = linea_raw.strip()
                if not linea:
                    continue

                if not inicio_movimientos:
                    if "Fecha" in linea and "Descripción" in linea and ("Débito" in linea or "Crédito" in linea) and "Saldo" in linea:
                        inicio_movimientos = True
                    continue

                if es_linea_ruido(linea):
                    continue

                match_fecha = PATRON_FECHA.match(linea)

                if match_fecha:
                    # ---- Nueva fila de movimiento ----
                    fecha = match_fecha.group(1)
                    resto = match_fecha.group(2)

                    montos = PATRON_MONTO.findall(resto)
                    saldo_match = PATRON_SALDO.search(resto)
                    saldo_valor = limpiar_valor_numerico(saldo_match.group(1)) if saldo_match else 0.0

                    idx_dollar = resto.find('$')
                    descripcion = resto[:idx_dollar].strip() if idx_dollar != -1 else resto.strip()

                    if montos:
                        signo, valor_str = montos[0]
                        valor = limpiar_valor_numerico(valor_str)
                        importe = valor if signo == '+' else -valor
                        tipo_movimiento = "Credito" if signo == '+' else "Debito"
                    else:
                        importe = 0.0
                        tipo_movimiento = "Desconocido"

                    movimientos.append({
                        'fecha': fecha,
                        'descripcion': descripcion,
                        'detalle_lineas': [],
                        'importe': importe,
                        'saldo': saldo_valor,
                        'tipo_movimiento': tipo_movimiento
                    })
                else:
                    # ---- Línea de detalle del movimiento anterior (incluye continuaciones entre páginas) ----
                    if movimientos:
                        movimientos[-1]['detalle_lineas'].append(linea)

    for mov in movimientos:
        mov['detalle'] = " | ".join(mov.pop('detalle_lineas'))

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
    with st.spinner('Procesando el archivo PDF...'):
        try:
            movimientos = extraer_movimientos_del_pdf(uploaded_file)

            if movimientos:
                df = pd.DataFrame(movimientos, columns=['fecha', 'descripcion', 'detalle', 'importe', 'saldo', 'tipo_movimiento'])

                st.success(f'¡Procesamiento completado! Se encontraron {len(movimientos)} movimientos.')

                st.subheader("Vista previa de los datos extraídos:")
                st.dataframe(df)

                st.markdown(get_table_download_link(df), unsafe_allow_html=True)
            else:
                st.warning("No se encontraron movimientos en el PDF. Verifica que sea un extracto bancario del Banco Galicia en formato Office Banking.")
        except Exception as e:
            st.error(f"Error al procesar el archivo: {str(e)}")

# Información adicional
st.markdown("---")
st.markdown("""
### Información
- Esta aplicación procesa extractos bancarios del Banco Galicia en formato "Office Banking".
- El archivo resultante tendrá las columnas: fecha, descripcion, detalle, importe, saldo y tipo_movimiento.
- Los valores de débito mantienen su signo negativo para facilitar los cálculos.
- Los datos adicionales de cada movimiento (nombre, CUIT, banco, categoría) se agrupan en la columna "detalle", separados por " | ".
""")

st.markdown("---")
st.markdown("Desarrollado con Streamlit")
