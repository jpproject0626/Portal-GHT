@echo off
if /i "%~1"=="auto" goto auto

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
REM Solo se suben, POR NOMBRE, estos dos archivos. NUNCA se sube nada de salida\,
REM salida_interna\ ni logs\, ni clave_interna.txt ni config.json (estan en .gitignore).
git add datos.json excluidos_avance.json

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
goto :eof

:auto
REM ======================================================================
REM  MODO AUTOMATICO:  actualizar_portal.bat auto
REM  Sin pausas ni preguntas (para el Programador de tareas). Corre
REM  "python generar_datos.py --auto" y SOLO si termina con codigo 0 sube,
REM  por nombre, datos.json y excluidos_avance.json (igual que el modo
REM  normal). Si el codigo es distinto de 0 NO sube nada y este .bat termina
REM  devolviendo ESE MISMO codigo (2 falta archivo, 3 falta clave, 4 archivo
REM  bloqueado, 5 config invalido, 1 error inesperado; ver generar_datos.py).
REM  Si falla git (commit o push), devuelve el codigo de git.
REM  Nunca se sube nada de salida\, salida_interna\, logs\, clave_interna.txt
REM  ni config.json (estan en .gitignore y no se agregan).
REM ======================================================================
setlocal
cd /d "%~dp0"

python generar_datos.py --auto
set "CODIGO=%errorlevel%"
if not "%CODIGO%"=="0" goto auto_error

git add datos.json excluidos_avance.json
git diff --cached --quiet
if not errorlevel 1 goto auto_sin_cambios

git commit -q -m "Actualizar datos del dia"
set "CODIGO=%errorlevel%"
if not "%CODIGO%"=="0" goto auto_error_git

git push
set "CODIGO=%errorlevel%"
if not "%CODIGO%"=="0" goto auto_error_git

echo [auto] Listo: datos subidos.
exit /b 0

:auto_sin_cambios
echo [auto] Sin cambios que subir.
exit /b 0

:auto_error
echo [auto] generar_datos.py termino con codigo %CODIGO%. NO se sube nada.
exit /b %CODIGO%

:auto_error_git
echo [auto] Fallo git (codigo %CODIGO%). Los datos NO quedaron publicados.
exit /b %CODIGO%
