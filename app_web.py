import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import json
import os
from openpyxl import load_workbook
from io import BytesIO

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

def agregar_fila_excel(worksheet_excel, datos):
    """Agrega o actualiza una fila en Excel según el código/ID"""
    try:
        codigo_a_escribir = str(datos.get('codigo', ''))
        
        # Buscar si el código ya existe
        fila_objetivo = None
        for idx, row in enumerate(worksheet_excel.iter_rows(min_row=2, min_col=3, max_col=3), start=2):
            if str(row[0].value).strip() == codigo_a_escribir.strip():
                fila_objetivo = idx
                break
        
        # Si no existe, agregar al final
        if fila_objetivo is None:
            fila_objetivo = worksheet_excel.max_row + 1
        
        # Escribir datos básicos
        worksheet_excel.cell(fila_objetivo, 1).value = datos.get('entidad', 'OTRO')  # Columna A
        worksheet_excel.cell(fila_objetivo, 2).value = datos.get('nit', '901145808-5')  # Columna B
        worksheet_excel.cell(fila_objetivo, 3).value = codigo_a_escribir  # Columna C
        
        # Estado Físico Fuste
        for col_excel, marcado in datos.get('checks_fuste', {}).items():
            if marcado:
                worksheet_excel.cell(fila_objetivo, col_excel).value = '1'
        
        # Estados generales
        if datos.get('fuste_general'):
            worksheet_excel.cell(fila_objetivo, 23).value = datos['fuste_general']
        if datos.get('raiz_especifico'):
            worksheet_excel.cell(fila_objetivo, 24).value = datos['raiz_especifico']
        if datos.get('raiz_general'):
            worksheet_excel.cell(fila_objetivo, 25).value = datos['raiz_general']
        
        # Estado Sanitario Copa
        for col_excel, marcado in datos.get('checks_copa', {}).items():
            if marcado:
                worksheet_excel.cell(fila_objetivo, col_excel).value = '1'
        
        if datos.get('san_copa_especifico'):
            worksheet_excel.cell(fila_objetivo, 48).value = datos['san_copa_especifico']
        
        # Estado Sanitario Fuste
        for col_excel, marcado in datos.get('checks_fuste_san', {}).items():
            if marcado:
                worksheet_excel.cell(fila_objetivo, col_excel).value = '1'
        
        # Estados sanitarios generales
        if datos.get('san_general'):
            worksheet_excel.cell(fila_objetivo, 49).value = datos['san_general']
        if datos.get('san_copa_general'):
            worksheet_excel.cell(fila_objetivo, 50).value = datos['san_copa_general']
        if datos.get('san_fuste_general'):
            worksheet_excel.cell(fila_objetivo, 51).value = datos['san_fuste_general']
        if datos.get('san_raiz_general'):
            worksheet_excel.cell(fila_objetivo, 52).value = datos['san_raiz_general']
        
        # Causas de Poda
        for col_excel, marcado in datos.get('checks_poda', {}).items():
            if marcado:
                worksheet_excel.cell(fila_objetivo, col_excel).value = '1'
        
        # Tipo e intensidad poda
        if datos.get('tipo_poda'):
            worksheet_excel.cell(fila_objetivo, 66).value = datos['tipo_poda']
        if datos.get('intensidad'):
            worksheet_excel.cell(fila_objetivo, 67).value = datos['intensidad']
        if datos.get('residuos'):
            worksheet_excel.cell(fila_objetivo, 77).value = datos['residuos']
        
        # Concepto Técnico
        for col_excel, marcado in datos.get('checks_concepto', {}).items():
            if marcado:
                worksheet_excel.cell(fila_objetivo, col_excel).value = '1'
        
        return True, fila_objetivo
    except Exception as e:
        return False, str(e)

def agregar_fila_sheets(worksheet, datos):
    """Agrega o actualiza una fila en Google Sheets según el código/ID"""
    try:
        # Obtener todos los valores de la columna C (códigos)
        valores_columna_c = worksheet.col_values(3)  # Columna C
        codigo_a_escribir = str(datos.get('codigo', ''))
        
        # Buscar si el código ya existe
        fila_objetivo = None
        for idx, valor in enumerate(valores_columna_c, start=1):
            if str(valor).strip() == codigo_a_escribir.strip():
                fila_objetivo = idx
                break
        
        # Si no existe, agregar al final
        if fila_objetivo is None:
            valores = worksheet.get_all_values()
            fila_objetivo = len(valores) + 1
        
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
        
        # Escribir en la fila específica (reemplaza si existe, crea si no)
        # Construir el rango A:fila_objetivo hasta la columna que necesites
        rango = f'A{fila_objetivo}:BX{fila_objetivo}'  # BX = columna 76, ajusta si necesitas más
        worksheet.update(rango, [fila_datos[:76]])  # Escribir solo las columnas necesarias
        
        return True, fila_objetivo
    except Exception as e:
        return False, str(e)

# Header
st.markdown("""
<div class="header-style">
    <h1>🌳 Asistente de Registro de Árboles</h1>
    <p>Registro en tiempo real sincronizado con Google Sheets o Excel</p>
</div>
""", unsafe_allow_html=True)

# Selector de modo
st.markdown("### 📂 Selecciona el modo de trabajo:")
modo = st.radio(
    "¿Cómo quieres trabajar?",
    ["🌐 Google Sheets (en la nube)", "📁 Archivo Excel (subir/descargar)"],
    horizontal=True
)

# Variable global para saber qué modo está activo
usa_google_sheets = (modo == "🌐 Google Sheets (en la nube)")

# ============= MODO GOOGLE SHEETS =============
if usa_google_sheets:
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
        
        # Buscar la hoja BASE DE DATOS (con o sin espacios)
        worksheet = None
        for sheet in spreadsheet.worksheets():
            if "BASE DE DATOS" in sheet.title.upper():
                worksheet = sheet
                break
        
        if worksheet is None:
            st.error("❌ No se encontró ninguna hoja con nombre 'BASE DE DATOS'")
            st.info(f"Hojas disponibles: {', '.join([s.title for s in spreadsheet.worksheets()])}")
            st.stop()
        
        st.success(f"✅ **Conectado a:** {spreadsheet.title} (Hoja: {worksheet.title})")
        
        # Obtener último código
        ultimo_codigo = obtener_ultimo_codigo(worksheet)
        siguiente_codigo = ultimo_codigo + 1
        
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.info("Verifica que:\n- El ID sea correcto\n- Hayas compartido el sheet con: asistente@excel-485919.iam.gserviceaccount.com\n- La cuenta de servicio tenga permisos de Editor")
        st.stop()

    # Inicializar session state para el código
    if 'codigo_actual' not in st.session_state:
        st.session_state.codigo_actual = siguiente_codigo

    # Variable para controlar si se debe reiniciar el formulario
    if 'form_key' not in st.session_state:
        st.session_state.form_key = 0
        
# ============= MODO ARCHIVO EXCEL =============
else:
    st.markdown("### 📤 Sube tu archivo Excel")
    
    uploaded_file = st.file_uploader(
        "Selecciona tu archivo Excel (.xlsx o .xlsm)",
        type=['xlsx', 'xlsm'],
        help="Sube el archivo Excel que contiene la hoja 'BASE DE DATOS'"
    )
    
    if uploaded_file is None:
        st.info("⬆️ Sube tu archivo Excel para continuar")
        st.stop()
    
    # Cargar Excel en memoria
    try:
        if 'excel_workbook' not in st.session_state or st.session_state.get('uploaded_filename') != uploaded_file.name:
            # Cargar sin imágenes para evitar errores
            wb = load_workbook(uploaded_file, keep_vba=True, data_only=False, keep_links=False)
            # Eliminar imágenes si existen para evitar problemas al guardar
            for sheet in wb.worksheets:
                if hasattr(sheet, '_images'):
                    sheet._images = []
            st.session_state.excel_workbook = wb
            st.session_state.uploaded_filename = uploaded_file.name
            st.session_state.excel_bytes = uploaded_file.getvalue()
        else:
            wb = st.session_state.excel_workbook
        
        # Buscar hoja BASE DE DATOS
        worksheet_excel = None
        for sheet_name in wb.sheetnames:
            if "BASE DE DATOS" in sheet_name.upper():
                worksheet_excel = wb[sheet_name]
                break
        
        if worksheet_excel is None:
            st.error("❌ No se encontró ninguna hoja con nombre 'BASE DE DATOS'")
            st.info(f"Hojas disponibles: {', '.join(wb.sheetnames)}")
            st.stop()
        
        st.success(f"✅ **Archivo cargado:** {uploaded_file.name} (Hoja: {worksheet_excel.title})")
        
        # Obtener último código del Excel
        ultimo_codigo_excel = 19222
        for row in worksheet_excel.iter_rows(min_row=2, min_col=3, max_col=3):
            valor = row[0].value
            if valor and str(valor).isdigit():
                ultimo_codigo_excel = max(ultimo_codigo_excel, int(valor))
        
        siguiente_codigo = ultimo_codigo_excel + 1
        
        if 'codigo_actual' not in st.session_state:
            st.session_state.codigo_actual = siguiente_codigo
        
        # Variable para controlar si se debe reiniciar el formulario
        if 'form_key' not in st.session_state:
            st.session_state.form_key = 0
        
        worksheet = worksheet_excel  # Para usar en el formulario
        
    except Exception as e:
        st.error(f"formulario_arbol_{st.session_state.form_key}ar Excel: {str(e)}")
        st.stop()

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
    texto_boton = "➕ Agregar Registro a Google Sheets" if usa_google_sheets else "➕ Agregar Registro al Excel"
    submitted = st.form_submit_button(texto_boton, use_container_width=True, type="primary")
    
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
        
        # Guardar según el modo
        if usa_google_sheets:
            with st.spinner('Guardando en Google Sheets...'):
                exito, resultado = agregar_fila_sheets(worksheet, datos)
        else:
            with st.spinner('Guardando en Excel...'):
                exito, resultado = agregar_fila_excel(worksheet, datos)
                if exito:
                    # Guardar el workbook actualizado en session_state
                    st.session_state.excel_workbook = wb
        
        if exito:
            # Auto-incrementar código para el siguiente
            st.session_state.codigo_actual = int(codigo) + 1
            # Cambiar la key del formulario para forzar reset
            st.session_state.form_key += 1
            
            st.success(f"✅ **Registro guardado exitosamente en la fila {resultado}** (ID: {codigo})")
            st.info(f"💡 Siguiente código: **{st.session_state.codigo_actual}**")
            st.balloons()
            st.rerun()
        else:
            st.error(f"❌ Error al guardar: {resultado}")

# Si es modo Excel, mostrar botón de descarga
if not usa_google_sheets and 'excel_workbook' in st.session_state:
    st.markdown("---")
    st.markdown("### 💾 Descargar Excel Actualizado")
    
    # Guardar workbook en bytes
    try:
        output = BytesIO()
        st.session_state.excel_workbook.save(output)
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 Descargar Excel con cambios",
            data=excel_data,
            file_name=f"arboles_actualizado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.info("💡 **Recuerda descargar el archivo** cuando termines de agregar todos los registros.")
    except Exception as e:
        st.warning("""
        ⚠️ **No se pudo generar la descarga automática.**
        
        Esto suele pasar si el Excel tiene imágenes o formatos complejos.
        
        **Solución:** Usa el modo Google Sheets para una experiencia sin problemas.
        """)
        if st.checkbox("Mostrar detalles del error"):
            st.error(f"Error técnico: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; font-size: 12px;">
    🌳 Asistente de Registro de Árboles | Google Sheets o Excel
</div>
""", unsafe_allow_html=True)
