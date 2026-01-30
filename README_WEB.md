# 🌐 Versión Web - Asistente de Registro de Árboles

## 📱 Funciona en Cualquier Dispositivo

✅ Computadora  
✅ Tablet  
✅ Celular  
✅ Sincronización en tiempo real con Google Sheets  

---

## 🚀 Instalación Rápida

### 1. Instalar Dependencias

Ejecuta en PowerShell:

```powershell
pip install -r requirements.txt
```

### 2. Configurar Google Sheets API

#### A. Crear Proyecto en Google Cloud

1. Ve a https://console.cloud.google.com/
2. Crea un nuevo proyecto (ej: "asistente-arboles")
3. Habilita estas APIs:
   - Google Sheets API
   - Google Drive API

#### B. Crear Cuenta de Servicio

1. Ve a "IAM y administración" → "Cuentas de servicio"
2. Click "Crear cuenta de servicio"
3. Nombre: `asistente-sheets`
4. Click "Crear y continuar"
5. Rol: "Editor"
6. Click "Listo"

#### C. Crear Clave JSON

1. Click en la cuenta de servicio recién creada
2. Ve a la pestaña "Claves"
3. Click "Agregar clave" → "Crear nueva clave"
4. Selecciona "JSON"
5. Se descargará un archivo JSON
6. **Renómbralo a `credenciales.json`**
7. **Copia este archivo a la carpeta del proyecto** (donde está `app_web.py`)

#### D. Preparar Google Sheets

1. Abre tu Google Sheet de árboles
2. **Asegúrate de que tenga una hoja llamada "BASE DE DATOS"**
3. Click en "Compartir" (arriba a la derecha)
4. Copia el **email de la cuenta de servicio** (está en el archivo JSON, campo `client_email`)
   - Se ve algo así: `asistente-sheets@proyecto.iam.gserviceaccount.com`
5. Pégalo en "Agregar personas y grupos"
6. Dale permisos de **Editor**
7. Click "Enviar"

#### E. Obtener ID de Google Sheets

1. Abre tu Google Sheet
2. Mira la URL: `https://docs.google.com/spreadsheets/d/[ESTE_ES_EL_ID]/edit`
3. Copia el ID (la parte entre `/d/` y `/edit`)

---

## 🎯 Ejecutar Localmente

### Opción 1: Usando Script

Doble click en:
```
EJECUTAR_WEB.bat
```

### Opción 2: Comando Manual

```powershell
streamlit run app_web.py
```

Se abrirá en tu navegador: `http://localhost:8501`

---

## 📱 Usar desde Celular (Red Local)

1. Ejecuta la app en tu computadora
2. Streamlit mostrará URLs como:
   ```
   Local URL: http://localhost:8501
   Network URL: http://192.168.1.X:8501
   ```
3. Desde tu celular (conectado a la misma WiFi):
   - Abre el navegador
   - Ingresa la **Network URL**
   - ¡Listo! Puedes usarla desde el celular

---

## ☁️ Desplegar en Internet (Gratis)

### Opción 1: Streamlit Cloud (Recomendado)

1. Crea cuenta en https://streamlit.io/cloud
2. Sube tu proyecto a GitHub (sin el archivo `credenciales.json`)
3. En Streamlit Cloud, click "New app"
4. Conecta tu repositorio de GitHub
5. En "Advanced settings" → "Secrets", pega el contenido de `credenciales.json`:

```toml
[gcp_service_account]
type = "service_account"
project_id = "tu-proyecto"
private_key_id = "abc123..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "asistente@proyecto.iam.gserviceaccount.com"
client_id = "123456789"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

6. Click "Deploy"
7. ¡Tu app estará disponible en `https://tu-app.streamlit.app`!

### Opción 2: Render.com (Gratis)

1. Crea cuenta en https://render.com
2. Crea nuevo "Web Service"
3. Conecta tu repositorio de GitHub
4. Configura:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `streamlit run app_web.py --server.port=$PORT`
5. Agrega las credenciales como variables de entorno
6. Deploy

---

## 🔒 Seguridad

⚠️ **IMPORTANTE**: 

- **NUNCA** subas `credenciales.json` a GitHub
- Usa `.gitignore` para excluirlo:
  ```
  credenciales.json
  .streamlit/secrets.toml
  ```
- En producción, usa las variables de entorno/secrets de la plataforma

---

## 🛠️ Estructura del Proyecto

```
pa/
├── app_web.py              # 🆕 Aplicación web con Streamlit
├── asistente_directo.py    # ⚙️ Aplicación de escritorio original
├── requirements.txt        # 📦 Dependencias Python
├── credenciales.json       # 🔑 Credenciales Google (NO subir a GitHub)
├── EJECUTAR_WEB.bat        # 🚀 Script para ejecutar app web
├── EJECUTAR_ASISTENTE.bat  # 🖥️ Script para app de escritorio
├── README.md               # 📖 Documentación original
└── README_WEB.md           # 🌐 Documentación versión web (este archivo)
```

---

## 📊 Flujo de Datos

1. **Usuario** → Llena formulario en navegador (PC/celular)
2. **Streamlit** → Procesa los datos
3. **Google Sheets API** → Escribe en Google Sheets
4. **Google Sheets** → Sincroniza en tiempo real con todos los dispositivos

---

## ❓ Problemas Comunes

### "Error al conectar a Google Sheets"

✅ Verifica que:
- El archivo `credenciales.json` esté en la carpeta correcta
- Hayas compartido el Google Sheet con el email de la cuenta de servicio
- Las APIs estén habilitadas en Google Cloud

### "No se encuentra la hoja BASE DE DATOS"

✅ Asegúrate de que tu Google Sheet tenga una hoja llamada exactamente **"BASE DE DATOS"**

### "Permission denied"

✅ La cuenta de servicio necesita permisos de **Editor** en el Google Sheet

### No se conecta desde celular

✅ Verifica que:
- El celular y la PC estén en la misma red WiFi
- No haya firewall bloqueando el puerto 8501
- Estés usando la Network URL (no la Local URL)

---

## 🎨 Ventajas vs Versión de Escritorio

| Característica | Escritorio (Tkinter) | Web (Streamlit) |
|----------------|----------------------|------------------|
| Instalación | Solo en PC con Python | Cualquier navegador |
| Móvil | ❌ No | ✅ Sí |
| Múltiples usuarios | ❌ Difícil | ✅ Fácil |
| Sincronización | ⚠️ Manual (Excel local) | ✅ Tiempo real (Cloud) |
| Actualizaciones | ⚠️ Manual | ✅ Automáticas |
| Internet requerido | ❌ No | ✅ Sí |
| Velocidad | ⚡ Muy rápida | ⚡ Rápida |

---

## 💡 Próximas Mejoras

- [ ] Autenticación de usuarios
- [ ] Visualización de datos (gráficos)
- [ ] Exportar a Excel
- [ ] Historial de cambios
- [ ] Búsqueda y filtros
- [ ] Modo offline (PWA)

---

## 📞 Soporte

¿Problemas? Revisa:
1. Este archivo (README_WEB.md)
2. Documentación de Google Sheets API: https://developers.google.com/sheets/api
3. Documentación de Streamlit: https://docs.streamlit.io

---

**¡Disfruta tu aplicación web! 🌳📱☁️**
