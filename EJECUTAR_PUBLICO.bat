@echo off
echo.
echo ========================================
echo   Aplicacion Web - Acceso desde Internet
echo ========================================
echo.

REM Verificar si ngrok existe
where ngrok >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] ngrok no esta instalado
    echo.
    echo Descargalo e instalalo desde: https://ngrok.com/download
    echo 1. Descarga ngrok para Windows
    echo 2. Extrae el archivo ngrok.exe
    echo 3. Muevelo a esta carpeta O agregalo al PATH
    echo.
    pause
    exit /b 1
)

echo [1/2] Iniciando Streamlit en segundo plano...
start /B .venv\Scripts\python.exe -m streamlit run app_web.py --server.port=8501 --server.headless=true

echo [2/2] Esperando que Streamlit inicie...
timeout /t 5 /nobreak >nul

echo [3/3] Creando tunel publico con ngrok...
echo.
echo ========================================
echo   COPIA LA URL QUE APARECE ABAJO
echo   Compartela para acceso desde cualquier red
echo ========================================
echo.

ngrok http 8501

pause
