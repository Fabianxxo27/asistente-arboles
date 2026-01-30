# 🌐 Acceso desde Cualquier Red WiFi/Datos Móviles

## Opción 1: Ngrok (Más Rápida - 5 minutos)

### Paso 1: Instalar ngrok

1. Ve a: https://ngrok.com/download
2. Descarga **ngrok para Windows**
3. Extrae el archivo `ngrok.exe`
4. Mueve `ngrok.exe` a esta carpeta (donde está `app_web.py`)

### Paso 2: Crear cuenta (gratis)

1. Ve a: https://dashboard.ngrok.com/signup
2. Regístrate gratis
3. Copia tu **Authtoken**
4. En PowerShell ejecuta:
   ```powershell
   .\ngrok.exe config add-authtoken TU_TOKEN_AQUI
   ```

### Paso 3: Ejecutar

**Doble click en:** `EJECUTAR_PUBLICO.bat`

Verás algo así:
```
Forwarding   https://abc123.ngrok.io -> http://localhost:8501
```

**¡Esa URL funciona desde cualquier lugar del mundo!** 🌍

### Uso:

1. Copia la URL (ej: `https://abc123.ngrok.io`)
2. Ábrela en tu celular desde datos móviles o cualquier WiFi
3. ¡Ya puedes usarla!

⚠️ **Nota**: La URL cambia cada vez que reinicias. Para URL fija, usa cuenta paga de ngrok o despliega en la nube.

---

## Opción 2: Streamlit Cloud (Permanente - GRATIS)

### Ventajas:
- ✅ URL fija y permanente (no cambia)
- ✅ Funciona 24/7 sin tu computadora encendida
- ✅ Totalmente gratis
- ✅ Actualización automática

### Pasos:

1. **Crear cuenta GitHub** (si no tienes): https://github.com/signup

2. **Subir tu proyecto a GitHub**:
   ```powershell
   git init
   git add app_web.py requirements.txt .gitignore
   git commit -m "Aplicacion web arboles"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/asistente-arboles.git
   git push -u origin main
   ```

3. **Ir a Streamlit Cloud**: https://streamlit.io/cloud

4. **Click "New app"**

5. **Configurar**:
   - Repository: Selecciona tu repo `asistente-arboles`
   - Branch: `main`
   - Main file path: `app_web.py`

6. **Advanced settings** → **Secrets**:
   - Pega el contenido de tu archivo `credenciales.json` en formato TOML:
   
   ```toml
   [gcp_service_account]
   type = "service_account"
   project_id = "tu-proyecto-id"
   private_key_id = "abc123..."
   private_key = "-----BEGIN PRIVATE KEY-----\nTU_CLAVE_AQUI\n-----END PRIVATE KEY-----\n"
   client_email = "tu-email@proyecto.iam.gserviceaccount.com"
   client_id = "123456789"
   auth_uri = "https://accounts.google.com/o/oauth2/auth"
   token_uri = "https://oauth2.googleapis.com/token"
   auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
   client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
   ```

7. **Click "Deploy"**

8. **Listo**: Tu app estará en `https://tu-app.streamlit.app`

---

## Opción 3: Localtunnel (Alternativa a ngrok)

```powershell
npm install -g localtunnel
```

Luego ejecuta:
```powershell
# Terminal 1
.venv\Scripts\python.exe -m streamlit run app_web.py

# Terminal 2 (nueva ventana)
lt --port 8501
```

Te da una URL pública como: `https://funny-tiger-12.loca.lt`

---

## Comparación:

| Método | Configuración | Costo | URL Fija | Sin PC encendida |
|--------|---------------|-------|----------|------------------|
| **Ngrok** | 5 min | Gratis | ❌ | ❌ |
| **Streamlit Cloud** | 15 min | Gratis | ✅ | ✅ |
| **Localtunnel** | 3 min | Gratis | ❌ | ❌ |

**Recomendación**: 
- Para pruebas rápidas → **Ngrok**
- Para producción permanente → **Streamlit Cloud**

---

¿Necesitas ayuda con alguna opción específica?
