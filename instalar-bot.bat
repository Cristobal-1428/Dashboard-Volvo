@echo off
rem Prepara el entorno del bot de Academia dentro de esta carpeta (python-portable + Chromium).
rem iniciar-local.bat lo llama solo si falta algo; tambien se puede ejecutar a mano.
setlocal
cd /d "%~dp0"
set "PYP=%~dp0python-portable"
set "PY=%PYP%\python.exe"

echo ================================================================
echo  Preparando el bot de Academia ^(requiere internet, solo esta vez^)
echo ================================================================

rem 1. Habilitar site-packages en el Python portable
findstr /b /c:"import site" "%PYP%\python312._pth" >nul 2>nul
if errorlevel 1 (
    (
        echo python312.zip
        echo .
        echo Lib\site-packages
        echo.
        echo import site
    ) > "%PYP%\python312._pth"
)

rem 2. pip
"%PY%" -m pip --version >nul 2>nul
if errorlevel 1 (
    echo.
    echo [1/3] Instalando pip...
    powershell -NoProfile -Command "[Net.ServicePointManager]::CheckCertificateRevocationList=$false; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest https://bootstrap.pypa.io/get-pip.py -OutFile '%PYP%\get-pip.py'"
    if errorlevel 1 goto error
    "%PY%" "%PYP%\get-pip.py" --no-warn-script-location
    if errorlevel 1 goto error
    del "%PYP%\get-pip.py" >nul 2>nul
)

rem 3. Librerias del bot
echo.
echo [2/3] Instalando librerias del bot...
"%PY%" -m pip install --no-warn-script-location -r "bot\academia\requirements.txt"
if errorlevel 1 goto error

rem 4. Chromium dentro de bot\browsers
echo.
echo [3/3] Descargando Chromium...
"%PY%" "bot\academia\main.py" --install-chromium
if errorlevel 1 goto error

echo.
echo Bot listo.
exit /b 0

:error
echo.
echo ERROR: no se pudo preparar el bot. Revisa la conexion a internet e intenta de nuevo
echo ^(vuelve a abrir iniciar-local.bat^). El portal funciona igual, pero el bot no estara disponible.
exit /b 1
