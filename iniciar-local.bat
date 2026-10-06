@echo off
cd /d "%~dp0"

rem El bot de Academia vive en bot\academia y corre con python-portable.
rem Si falta algo (primera vez en un equipo nuevo), se prepara solo.
if not exist "%~dp0python-portable\Lib\site-packages\pandas" goto instalar_bot
if not exist "%~dp0python-portable\Lib\site-packages\playwright" goto instalar_bot
if not exist "%~dp0bot\browsers" goto instalar_bot
goto servidor
:instalar_bot
call "%~dp0instalar-bot.bat"

:servidor
where python >nul 2>nul
if %errorlevel%==0 (
    echo Python encontrado en el sistema, usando ese.
    set "PYTHON_EXE=python"
) else (
    echo Python no esta instalado en este computador.
    echo Usando la copia portatil incluida en esta carpeta ^(python-portable^), no se instala nada.
    set "PYTHON_EXE=%~dp0python-portable\python.exe"
)

echo.
"%PYTHON_EXE%" server.py
