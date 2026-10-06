"""Genera un zip liviano del portal para compartir (por ejemplo, subirlo a GitHub).

    python empaquetar.py [ruta_de_salida.zip]

Deja fuera todo lo que se puede regenerar o que no debe compartirse:
  - credenciales (.env) y las correcciones de RUT/correos del bot (*_override.json)
  - las librerías instaladas y Chromium (instalar-bot.bat los descarga solos)
  - el bot compilado antiguo (bot/CertificadosVolvoAuto), el respaldo de Firebase, logs y cachés
Se conserva Tkinter (lo necesita la interfaz del bot y no se reinstala solo).
"""
import os
import sys
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
SALIDA = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ.parent / "portal-volvo.zip"

# Carpetas que no se incluyen (relativas a la raíz, con "/")
EXCLUIR_CARPETAS = {
    "bot/CertificadosVolvoAuto",
    "bot/browsers",
    "_respaldo_firebase",
    "python-portable/Scripts",
    "python-portable/Include",
    "bot/academia/logs",
    "bot/academia/errores",
    "bot/academia/reportes",
    "bot/academia/Excels",
    "cache",
}
EXCLUIR_NOMBRES_CARPETA = {"__pycache__", ".git", ".vscode", ".venv", "venv"}
EXCLUIR_ARCHIVOS = {
    "bot/academia/.env",
    "bot/academia/correos_override.json",
    "bot/academia/razones_override.json",
    "bot/academia/ruts_override.json",
    "bot/academia/forms_file_seleccionado.txt",
    "python-portable/python312._pth.orig",
}
# En python-portable/Lib/site-packages solo se conserva tkinter
SITE_PACKAGES = "python-portable/Lib/site-packages"
CONSERVAR_EN_SITE_PACKAGES = {"tkinter"}
# Ayudante de código que vive dentro de una carpeta excluida
INCLUIR_SIEMPRE = {"bot/academia/Excels/inyectar_registro.py"}


def excluido(rel, es_archivo):
    if rel in INCLUIR_SIEMPRE:
        return False
    if any(rel == c or rel.startswith(c + "/") for c in EXCLUIR_CARPETAS):
        return True
    if es_archivo and (rel in EXCLUIR_ARCHIVOS or rel.endswith((".pyc", ".log"))):
        return True
    if rel.startswith("datos/") and es_archivo and not rel.endswith("LEEME.txt"):
        return True   # copias de Excel de la empresa: no van al repositorio
    if rel.startswith(SITE_PACKAGES + "/"):
        primero = rel[len(SITE_PACKAGES) + 1:].split("/")[0]
        return primero not in CONSERVAR_EN_SITE_PACKAGES
    return False


def vale_entrar(rel_carpeta):
    """Una carpeta excluida igual se recorre si guarda un archivo de INCLUIR_SIEMPRE."""
    if any(i.startswith(rel_carpeta + "/") for i in INCLUIR_SIEMPRE):
        return True
    return not excluido(rel_carpeta, False)


def main():
    if SALIDA.exists():
        SALIDA.unlink()
    total = 0
    with zipfile.ZipFile(SALIDA, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for carpeta, subcarpetas, archivos in os.walk(RAIZ):
            rel_carpeta = Path(carpeta).relative_to(RAIZ).as_posix()
            rel_carpeta = "" if rel_carpeta == "." else rel_carpeta
            subcarpetas[:] = [
                d for d in subcarpetas
                if d not in EXCLUIR_NOMBRES_CARPETA and vale_entrar((rel_carpeta + "/" + d).lstrip("/"))
            ]
            for nombre in archivos:
                rel = (rel_carpeta + "/" + nombre).lstrip("/")
                if excluido(rel, True):
                    continue
                z.write(Path(carpeta) / nombre, "portal-volvo/" + rel)
                total += 1
    print("Listo: %s  (%d archivos, %.1f MB)" % (SALIDA, total, SALIDA.stat().st_size / 1048576))


if __name__ == "__main__":
    main()
