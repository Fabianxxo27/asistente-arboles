import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import json
import os

# Configuración de la página
st.set_page_config(
    page_title="🌳 Asistente de Registro - Árboles",
    page_icon="🌳",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS personalizado
st.markdown("""
<style>
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
    }
    .header-style {
        background: linear-gradient(135deg, #2c5f2d 0%, #3d7f3f 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
    }
    .section-header {
        background-color: #2c5f2d;
        color: white;
        padding: 10px;
        border-radius: 5px;
        margin-top: 20px;
        margin-bottom: 10px;
        font-weight: bold;
    }
    div[data-testid="stCheckbox"] {
        display: inline-block;
        margin-right: 10px;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Función para conectar con Google Sheets
@st.cache_resource
def conectar_google_sheets():
    """Conecta a Google Sheets usando credenciales"""
    try:
        # Intenta cargar credenciales desde archivo
        if os.path.exists('credenciales.json'):
            SCOPES = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            creds = Credentials.from_service_account_file('credenciales.json', scopes=SCOPES)
            client = gspread.authorize(creds)
            return client
        # Si no existe archivo, intenta cargar desde secrets de Streamlit
        elif 'gcp_service_account' in st.secrets:
            SCOPES = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=SCOPES
            )
            client = gspread.authorize(creds)
            return client
        else:
            return None
    except Exception as e:
        st.error(f"Error al conectar: {str(e)}")
        return None

def obtener_ultimo_codigo(worksheet):
    """Obtiene el último código de la columna C"""
    try:
        valores = worksheet.col_values(3)  # Columna C
        # Filtrar solo números
        codigos = [int(v) for v in valores[1:] if v and str(v).isdigit()]
        if codigos:
            return max(codigos)
        return 19222  # Valor inicial si no hay códigos
    except:
        return 19222

def agregar_fila_sheets(worksheet, datos):
    """Agrega una fila a Google Sheets"""
    try:
        # Obtener última fila
        valores = worksheet.get_all_values()
        nueva_fila = len(valores) + 1
        
        # Preparar fila con 80 columnas (ajusta según tu Excel)
        fila_datos = [''] * 80
        
        # Datos básicos
        fila_datos[0] = datos.get('entidad', 'OTRO')  # Columna A
        fila_datos[1] = datos.get('nit', '901145808-5')  # Columna B
        fila_datos[2] = datos.get('codigo', '')  # Columna C
        
        # Estado Físico Fuste (columnas 4-19, índices 3-18)
        for col_excel, marcado in datos.get('checks_fuste', {}).items():
            if marcado:
                fila_datos[col_excel - 1] = '1'
        
        # Estados generales
        if datos.get('fuste_general'):
            fila_datos[22] = datos['fuste_general']  # Columna W (23)
        if datos.get('raiz_especifico'):
            fila_datos[23] = datos['raiz_especifico']  # Columna X (24)
        if datos.get('raiz_general'):
            fila_datos[24] = datos['raiz_general']  # Columna Y (25)
        
        # Estado Sanitario Copa (columnas 26-29, 40)
        for col_excel, marcado in datos.get('checks_copa', {}).items():
            if marcado:
                fila_datos[col_excel - 1] = '1'
        
        if datos.get('san_copa_especifico'):
            fila_datos[47] = datos['san_copa_especifico']  # Columna AV (48)
        
        # Estado Sanitario Fuste
        for col_excel, marcado in datos.get('checks_fuste_san', {}).items():
            if marcado:
                fila_datos[col_excel - 1] = '1'
        
        # Estados sanitarios generales
        if datos.get('san_general'):
            fila_datos[48] = datos['san_general']  # Columna AW (49)
        if datos.get('san_copa_general'):
            fila_datos[49] = datos['san_copa_general']  # Columna AX (50)
        if datos.get('san_fuste_general'):
            fila_datos[50] = datos['san_fuste_general']  # Columna AY (51)
        if datos.get('san_raiz_general'):
            fila_datos[51] = datos['san_raiz_general']  # Columna AZ (52)
        
        # Causas de Poda (columnas 61-64)
        for col_excel, marcado in datos.get('checks_poda', {}).items():
            if marcado:
                fila_datos[col_excel - 1] = '1'
        
        # Tipo e intensidad poda
        if datos.get('tipo_poda'):
            fila_datos[65] = datos['tipo_poda']  # Columna BN (66)
        if datos.get('intensidad'):
            fila_datos[66] = datos['intensidad']  # Columna BO (67)
        if datos.get('residuos'):
            fila_datos[76] = datos['residuos']  # Columna BY (77)
        
        # Concepto Técnico
        for col_excel, marcado in datos.get('checks_concepto', {}).items():
            if marcado:
                fila_datos[col_excel - 1] = '1'
        
        # Agregar la fila
        worksheet.append_row(fila_datos)
        return True, nueva_fila
    except Exception as e:
        return False, str(e)

# Header
st.markdown("""
<div class="header-style">
    <h1>🌳 Asistente de Registro de Árboles</h1>
    <p>Registro en tiempo real sincronizado con Google Sheets</p>
</div>
""", unsafe_allow_html=True)

# Intentar conectar
client = conectar_google_sheets()

if client is None:
    st.error("⚠️ **No se pudo conectar a Google Sheets**")
    st.info("""
    **Configuración necesaria:**
    
    1. Crea un proyecto en Google Cloud Platform
    2. Habilita Google Sheets API y Google Drive API
    3. Crea credenciales de cuenta de servicio
    4. Descarga el archivo JSON y guárdalo como `credenciales.json` en esta carpeta
    5. Comparte tu Google Sheet con el email de la cuenta de servicio
    
    📖 Ver instrucciones completas en `README_WEB.md`
    """)
    st.stop()

# Selector de hoja
spreadsheet_id = st.text_input(
    "🔗 ID de Google Sheets",
    placeholder="Pega aquí el ID de tu Google Sheet (desde la URL)",
    help="El ID está en la URL: https://docs.google.com/spreadsheets/d/[ID]/edit"
)

if not spreadsheet_id:
    st.warning("⬆️ Ingresa el ID de tu Google Sheet para continuar")
    st.stop()

try:
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet("BASE DE DATOS")
    
    st.success(f"✅ **Conectado a:** {spreadsheet.title}")
    
    # Obtener último código
    ultimo_codigo = obtener_ultimo_codigo(worksheet)
    siguiente_codigo = ultimo_codigo + 1
    
except Exception as e:
    st.error(f"❌ Error: {str(e)}")
    st.info("Verifica que:\n- El ID sea correcto\n- La hoja se llame 'BASE DE DATOS'\n- Hayas compartido el sheet con la cuenta de servicio")
    st.stop()

# Inicializar session state para el código
if 'codigo_actual' not in st.session_state:
    st.session_state.codigo_actual = siguiente_codigo

# Formulario
with st.form(key="formulario_arbol"):
    
    # Datos básicos
    col1, col2, col3 = st.columns(3)
    with col1:
        entidad = st.text_input("Entidad:", value="OTRO")
    with col2:
        nit = st.text_input("NIT:", value="901145808-5")
    with col3:
        codigo = st.number_input("ID/Código:", value=st.session_state.codigo_actual, step=1)
    
    # Estado Físico Fuste
    st.markdown('<div class="section-header">🌳 Estado Físico Fuste</div>', unsafe_allow_html=True)
    cols_fuste = st.columns(6)
    checks_fuste = {}
    opciones_fuste = [
        ("B", 4), ("Bb", 5), ("BB", 6), ("FR", 7), ("I", 8), ("MI", 9),
        ("To", 10), ("C", 11), ("Rv", 12), ("Ac", 13), ("An", 14), ("Dc", 15),
        ("SB", 16), ("Ag", 17), ("Poe", 18), ("Pe", 19)
    ]
    for idx, (nombre, col) in enumerate(opciones_fuste):
        with cols_fuste[idx % 6]:
            checks_fuste[col] = st.checkbox(nombre, key=f"fuste_{col}")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        fuste_general = st.selectbox("Estado Fuste General:", 
            ["", "Bueno", "Regular", "Malo", "Suprimido"], index=1)
    with col2:
        raiz_especifico = st.selectbox("Estado Raíz Específico:", 
            ["", "No apreciable", "Visible", "Superficial", "Profunda"], index=1)
    with col3:
        raiz_general = st.selectbox("Estado Raíz General:", 
            ["", "Bueno", "Regular", "Malo"], index=1)
    
    # Estado Sanitario Copa
    st.markdown('<div class="section-header">🍃 Estado Sanitario Copa</div>', unsafe_allow_html=True)
    cols_copa = st.columns(5)
    checks_copa = {}
    opciones_copa = [("He", 26), ("An", 27), ("Ag", 28), ("Ne", 29), ("NA", 40)]
    for idx, (nombre, col) in enumerate(opciones_copa):
        with cols_copa[idx]:
            checks_copa[col] = st.checkbox(nombre, key=f"copa_{col}")
    
    san_copa_especifico = st.selectbox("Estado Sanitario Copa Específico:", 
        ["", "Ninguna de las anteriores"], index=1)
    
    # Estado Sanitario Fuste
    st.markdown('<div class="section-header">🏥 Estado Sanitario Fuste</div>', unsafe_allow_html=True)
    checks_fuste_san = {}
    checks_fuste_san[47] = st.checkbox("NA", key="fuste_san_47")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        san_general = st.selectbox("Estado Sanitario General:", 
            ["", "Bueno", "Regular", "Malo"], index=1)
    with col2:
        san_copa_general = st.selectbox("Estado Sanitario Copa General:", 
            ["", "Bueno", "Regular", "Malo"], index=1)
    with col3:
        san_fuste_general = st.selectbox("Estado Sanitario Fuste General:", 
            ["", "Bueno", "Regular", "Malo"], index=1)
    with col4:
        san_raiz_general = st.selectbox("Estado Sanitario Raíz General:", 
            ["", "Bueno", "Regular", "Malo"], index=1)
    
    # Causas de Poda
    st.markdown('<div class="section-header">✂️ Causas de Poda</div>', unsafe_allow_html=True)
    cols_poda = st.columns(3)
    checks_poda = {}
    opciones_poda = [("Ramas rotas", 61), ("Copa asimétrica", 62), ("Ramas pendulares", 64)]
    for idx, (nombre, col) in enumerate(opciones_poda):
        with cols_poda[idx]:
            checks_poda[col] = st.checkbox(nombre, key=f"poda_{col}")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        tipo_poda = st.selectbox("Tipo de Poda:", 
            ["", "De mejoramiento-Estructura", "De mantenimiento", "Especial", "Sanitaria"], 
            index=1)
    with col2:
        intensidad = st.text_input("Intensidad (%):")
    with col3:
        residuos = st.text_input("Residuos (kg):")
    
    # Concepto Técnico
    st.markdown('<div class="section-header">📋 Concepto Técnico</div>', unsafe_allow_html=True)
    checks_concepto = {}
    checks_concepto[69] = st.checkbox("Problemas seguridad", key="concepto_69")
    
    # Botón enviar
    st.markdown("<br>", unsafe_allow_html=True)
    submitted = st.form_submit_button("➕ Agregar Registro a Google Sheets", 
                                      use_container_width=True, type="primary")
    
    if submitted:
        # Preparar datos
        datos = {
            'entidad': entidad,
            'nit': nit,
            'codigo': str(codigo),
            'checks_fuste': {col: True for col, val in checks_fuste.items() if val},
            'fuste_general': fuste_general,
            'raiz_especifico': raiz_especifico,
            'raiz_general': raiz_general,
            'checks_copa': {col: True for col, val in checks_copa.items() if val},
            'san_copa_especifico': san_copa_especifico,
            'checks_fuste_san': {col: True for col, val in checks_fuste_san.items() if val},
            'san_general': san_general,
            'san_copa_general': san_copa_general,
            'san_fuste_general': san_fuste_general,
            'san_raiz_general': san_raiz_general,
            'checks_poda': {col: True for col, val in checks_poda.items() if val},
            'tipo_poda': tipo_poda,
            'intensidad': intensidad,
            'residuos': residuos,
            'checks_concepto': {col: True for col, val in checks_concepto.items() if val}
        }
        
        # Agregar a sheets
        with st.spinner('Guardando en Google Sheets...'):
            exito, resultado = agregar_fila_sheets(worksheet, datos)
        
        if exito:
            st.success(f"✅ **Registro agregado exitosamente en la fila {resultado}**")
            st.balloons()
            
            # Auto-incrementar código
            st.session_state.codigo_actual = codigo + 1
            st.info(f"💡 El siguiente código sugerido es: **{st.session_state.codigo_actual}**")
            st.rerun()
        else:
            st.error(f"❌ Error al guardar: {resultado}")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; font-size: 12px;">
    🌳 Asistente de Registro de Árboles | Datos sincronizados en tiempo real con Google Sheets
</div>
""", unsafe_allow_html=True)
