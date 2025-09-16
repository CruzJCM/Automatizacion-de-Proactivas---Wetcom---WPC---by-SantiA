@echo off
setlocal

:: %~dp0 es una variable especial que significa "la carpeta donde está este archivo .bat"
set "BASE_DIR=%~dp0"

echo [PASO 1 de 3] - Ejecutando la recoleccion de datos de vSphere...
echo.

:: Llama a tu script 'run.bat' original y espera a que termine.
call "%BASE_DIR%run.bat"

echo.
echo [PASO 2 de 3] - Recoleccion finalizada. Ejecutando la conversion a Excel...
echo.

:: [CORREGIDO] Añadimos -ExecutionPolicy Bypass para evitar la pregunta de seguridad.
pwsh.exe -ExecutionPolicy Bypass -File "%BASE_DIR%JSONtoExcels.ps1"

echo.
echo [PASO 3 de 3] - Conversion a Excel finalizada. Ejecutando el filtrado y generando el Anexo...
echo.

:: [CORREGIDO] Usamos -ExecutionPolicy Bypass y el nombre de archivo correcto (.ps1).
pwsh.exe -ExecutionPolicy Bypass -File "%BASE_DIR%proactivas-auto2.0.ps1"

echo.
echo Proceso completado.
pause