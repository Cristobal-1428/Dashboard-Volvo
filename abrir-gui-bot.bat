@echo off
rem Abre la interfaz grafica del bot de Academia (sin pasar por el portal).
cd /d "%~dp0bot\academia"
set "PLAYWRIGHT_BROWSERS_PATH=%~dp0bot\browsers"
start "" "%~dp0python-portable\pythonw.exe" gui.py
