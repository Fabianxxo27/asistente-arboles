@echo off
chcp 65001 > nul
echo ========================================
echo   ⚡ ASISTENTE DIRECTO EXCEL
echo   (Escribe en Excel en tiempo real)
echo ========================================
echo.
echo 📋 ANTES DE CONTINUAR:
echo   1. Abre tu archivo Excel
echo   2. Ve a la hoja "BASE DE DATOS"
echo   3. Luego presiona cualquier tecla aquí
echo.
pause

echo Iniciando aplicación...
echo.

if exist .venv\Scripts\python.exe (
    .venv\Scripts\python.exe asistente_directo.py
) else (
    echo ❌ Entorno virtual no encontrado
    echo Por favor ejecuta primero "INSTALAR.bat"
    echo.
    pause
)
