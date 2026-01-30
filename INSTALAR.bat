@echo off
chcp 65001 > nul
echo ========================================
echo   🌳 INSTALADOR ASISTENTE EXCEL
echo ========================================
echo.

echo [1/4] Verificando Python...
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python no está instalado
    echo.
    echo Por favor instala Python desde: https://www.python.org/downloads/
    echo IMPORTANTE: Marca "Add Python to PATH" durante la instalación
    echo.
    pause
    exit /b 1
)
echo ✅ Python encontrado

echo.
echo [2/4] Creando entorno virtual...
if exist .venv (
    echo ⚠️  El entorno virtual ya existe, saltando...
) else (
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo ❌ Error al crear entorno virtual
        pause
        exit /b 1
    )
    echo ✅ Entorno virtual creado
)

echo.
echo [3/4] Activando entorno virtual...
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ❌ Error al activar entorno virtual
    pause
    exit /b 1
)
echo ✅ Entorno activado

echo.
echo [4/4] Instalando dependencias...
echo    - openpyxl (para leer/escribir Excel)
echo    - xlwings (para conexión en tiempo real)
pip install --quiet openpyxl xlwings
if %errorlevel% neq 0 (
    echo ❌ Error al instalar dependencias
    pause
    exit /b 1
)
echo ✅ Dependencias instaladas

echo.
echo ========================================
echo   ✅ INSTALACIÓN COMPLETADA
echo ========================================
echo.
echo 📋 Para usar la aplicación:
echo.
echo   Opción 1 (RECOMENDADO):
echo   1. Abre tu Excel
echo   2. Doble clic en "EJECUTAR_ASISTENTE.bat"
echo.
echo   Opción 2:
echo   1. Doble clic en "EJECUTAR_RAPIDO.bat"
echo.
pause
