@echo off
echo ============================================
echo   ACTUALIZANDO PORTAL GHT
echo ============================================
echo.

echo [1/4] Generando datos.json desde los archivos de Excel...
python generar_datos.py
if errorlevel 1 (
    echo.
    echo *** ERROR al generar los datos. Revisa el mensaje de arriba. ***
    pause
    exit /b
)

echo.
echo [2/4] Preparando el cambio para subir...
git add datos.json

echo [3/4] Guardando el cambio...
git commit -m "Actualizar datos del dia"

echo [4/4] Publicando en Vercel...
git push

echo.
echo ============================================
echo   LISTO. En 20-30 segundos el portal
echo   deberia estar actualizado en Vercel.
echo ============================================
pause
