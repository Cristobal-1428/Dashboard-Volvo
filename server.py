import http.server
import json
import os
from pathlib import Path
import socket
import socketserver
import subprocess
import sys
import threading
import time
import urllib.parse
import webbrowser

PORT = int(os.environ.get("VOLVO_PORT", "8000"))
# Solo este PC puede entrar al portal (no otros equipos de la red)
HOST = "127.0.0.1"


def puerto_libre(inicial, intentos=20):
    """Primer puerto desde 'inicial' que ninguna otra aplicación esté usando.

    Windows deja que dos programas escuchen en el mismo puerto si uno lo hace en una
    dirección específica (127.0.0.1) y el otro en todas; en ese caso 'localhost' le llega
    al otro programa y el portal parece no responder. Por eso se prueba con una conexión.
    """
    for puerto in range(inicial, inicial + intentos):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            if s.connect_ex((HOST, puerto)) != 0:
                return puerto
    return inicial
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# El bot corre desde su código fuente (bot/academia) con el Python portable del proyecto:
# lo que se edite ahí aplica en la próxima ejecución, sin recompilar nada.
PYTHON_BOT = os.path.join(BASE_DIR, "python-portable", "python.exe")
BOT_DIR = os.path.join(BASE_DIR, "bot", "academia")
BOT_MAIN = os.path.join(BOT_DIR, "main.py")
PYTHONW_BOT = os.path.join(BASE_DIR, "python-portable", "pythonw.exe")
BOT_GUI = os.path.join(BOT_DIR, "gui.py")
BOT_TKINTER = os.path.join(BASE_DIR, "python-portable", "_tkinter.pyd")
BOT_LIBS = [os.path.join(BASE_DIR, "python-portable", "Lib", "site-packages", n) for n in ("pandas", "playwright", "lxml", "openpyxl")]
BOT_NAVEGADOR = os.path.join(BASE_DIR, "bot", "browsers")

FORMS_FILENAME_DEFAULT = "Resultados_Microsoft6.xlsx"
FORMS_FILENAMES_DISPONIBLES = [
    "Resultados_Microsoft6.xlsx",
    "Resultados_Microsoft Forms.xlsx",
]
FORMS_FILE_CHOICE_PATH = os.path.join(BOT_DIR, "forms_file_seleccionado.txt")

# Base de perfil de usuario
USER_PROFILE = Path(os.environ.get("USERPROFILE", os.path.expanduser("~")))


# Copias de respaldo de los Excel, para PC donde la carpeta de OneDrive no está sincronizada
DATOS_DIR = Path(BASE_DIR) / "datos"
# Última copia leída de cada Excel, por si está abierto en Excel cuando la página lo pide
CACHE_DIR = os.path.join(BASE_DIR, "cache")


_cache_busqueda = {}


def buscar_en_accesos_directos(nombre_archivo, profundidad=3, vigencia=30):
    """Busca un archivo por nombre en cualquier carpeta sincronizada dentro de 'Shortcuts'
    o en la carpeta 'Dashboard VC-AC' del OneDrive personal (ahí se puede dejar una copia,
    a mano o con un flujo de Power Automate).

    Sirve cuando el acceso directo de OneDrive no se llama como se esperaba, o cuando el Excel
    está en otra subcarpeta de una carpeta que ya está sincronizada. El resultado se recuerda
    unos segundos para no recorrer OneDrive en cada consulta.
    """
    ahora = time.time()
    guardado = _cache_busqueda.get(nombre_archivo)
    if guardado and ahora - guardado[0] < vigencia:
        return guardado[1]

    objetivo = nombre_archivo.lower()
    encontrado = None
    base = USER_PROFILE / "OneDrive - Volvo Group"
    pendientes = [(base / "Shortcuts", 0), (base / "Dashboard VC-AC", 1)]
    while pendientes and encontrado is None:
        carpeta, nivel = pendientes.pop(0)
        try:
            with os.scandir(carpeta) as entradas:
                for e in entradas:
                    if e.is_file() and e.name.lower() == objetivo:
                        encontrado = Path(e.path)
                        break
                    if e.is_dir() and nivel < profundidad:
                        pendientes.append((Path(e.path), nivel + 1))
        except OSError:
            continue
    _cache_busqueda[nombre_archivo] = (ahora, encontrado)
    return encontrado


class RutaOnedrive(os.PathLike):
    """Ruta a un Excel de OneDrive, resuelta cada vez que se usa (así no hay que reiniciar el servidor).

    Orden de búsqueda:
      1. OneDrive - Volvo Group/Shortcuts/<partes> (el acceso directo) o, según el PC, OneDrive - Volvo Group/<partes>.
      2. Cualquier carpeta sincronizada dentro de 'Shortcuts', o la carpeta 'Dashboard VC-AC' del OneDrive, con un archivo con ese nombre.
      3. La carpeta 'datos' del proyecto, con el mismo nombre de archivo.
    Si no existe ninguna, se devuelve la de 'Shortcuts' para que el mensaje de error muestre una ruta razonable.
    """

    def __init__(self, *partes):
        self.partes = partes

    def resolver(self):
        base = USER_PROFILE / "OneDrive - Volvo Group"
        directa = base.joinpath(*self.partes)
        acceso_directo = base.joinpath("Shortcuts", *self.partes)
        if acceso_directo.exists():
            return acceso_directo
        if directa.exists():
            return directa
        en_otra_carpeta = buscar_en_accesos_directos(self.partes[-1])
        if en_otra_carpeta:
            return en_otra_carpeta
        local = DATOS_DIR / self.partes[-1]
        return local if local.exists() else acceso_directo

    def __fspath__(self):
        return str(self.resolver())

    __str__ = __fspath__


# 1. Ruta dinámica: Excel de Capacitaciones
REGISTRO_CAPACITACIONES_PATH = RutaOnedrive(
    "Team OCT&P - 6. Registro de capacitaciones",
    "Registro de capacitaciones actualizado.xlsx",
)

# 2. Ruta dinámica: Reporte Maestro de Camiones
REPORTE_MAESTRO_PATH = RutaOnedrive(
    "Team OCT&P - 04.- Conectividad",
    "00.- Negocio",
    "RM_JAVI_VOLVO CONNECT.xlsx",
)

# 3. Ruta dinámica: BBDD Clientes CM
CLIENTES_CM_PATH = RutaOnedrive(
    "Team OCT&P - 04.- Conectividad",
    "Pruebas felipe",
    "BBDD CLIENTES CM.xlsx",
)

# 4. Ruta dinámica: Contactos
CONTACTOS_PATH = RutaOnedrive(
    "Team OCT&P - 04.- Conectividad",
    "Pruebas felipe",
    "contactos.xlsx",
)

def obtener_estado_archivos():
    """Retorna existencia y fecha mtime de las 4 planillas de OneDrive para la campana de notificaciones."""
    archivos = {
        "flota": REPORTE_MAESTRO_PATH,
        "academia": REGISTRO_CAPACITACIONES_PATH,
        "clientes": CLIENTES_CM_PATH,
        "contactos": CONTACTOS_PATH,
    }
    estado_archivos = {}
    for clave, ruta in archivos.items():
        if os.path.isfile(ruta):
            mtime = os.path.getmtime(ruta)
            estado_archivos[clave] = {
                "existe": True,
                "mtime": int(mtime * 1000),  # Milisegundos para JS
                "fecha": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
            }
        else:
            estado_archivos[clave] = {"existe": False}
    return estado_archivos


estado = {
    "ejecutando": False,
    "inicio": None,
    "fin": None,
    "codigo": None,
    "log": ""
}
lock = threading.Lock()

CREATE_NO_WINDOW = 0x08000000


def error_bot():
    """Motivo por el que el bot no puede correr, o None si está listo."""
    if not os.path.isfile(BOT_MAIN):
        return "No se encontró el bot en " + BOT_DIR
    if not os.path.isfile(PYTHON_BOT) or not all(os.path.isdir(d) for d in BOT_LIBS):
        return "El entorno del bot no está instalado. Cierra y vuelve a abrir iniciar-local.bat para instalarlo."
    if not os.path.isdir(BOT_NAVEGADOR):
        return "Falta Chromium para el bot. Cierra y vuelve a abrir iniciar-local.bat para instalarlo."
    return None


gui_proceso = None


def abrir_gui_bot():
    """Abre la interfaz gráfica del bot (gui.py) en una ventana aparte. Devuelve (ok, mensaje)."""
    global gui_proceso
    motivo = error_bot()
    if not motivo and not (os.path.isfile(BOT_GUI) and os.path.isfile(PYTHONW_BOT) and os.path.isfile(BOT_TKINTER)):
        motivo = "Falta la interfaz del bot (gui.py o Tkinter en python-portable)."
    if motivo:
        return False, motivo
    with lock:
        if gui_proceso is not None and gui_proceso.poll() is None:
            return True, "La interfaz del bot ya está abierta."
        gui_proceso = subprocess.Popen([PYTHONW_BOT, BOT_GUI], cwd=BOT_DIR, env=entorno_bot())
    return True, "Interfaz del bot abierta en una ventana nueva."


def entorno_bot():
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    env["PLAYWRIGHT_BROWSERS_PATH"] = BOT_NAVEGADOR
    return env


def ejecutar_bot():
    with lock:
        if estado["ejecutando"]:
            return False
        estado["ejecutando"] = True
        estado["inicio"] = time.strftime("%Y-%m-%d %H:%M:%S")
        estado["fin"] = None
        estado["codigo"] = None
        estado["log"] = ""

    def correr():
        try:
            proceso = subprocess.Popen(
                [PYTHON_BOT, BOT_MAIN],
                cwd=BOT_DIR,
                env=entorno_bot(),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=CREATE_NO_WINDOW
            )
            for linea in proceso.stdout:
                with lock:
                    estado["log"] += linea
            proceso.wait()
            with lock:
                estado["codigo"] = proceso.returncode
        except Exception as e:
            with lock:
                estado["log"] += "\nError al ejecutar el bot: " + str(e)
                estado["codigo"] = -1
        finally:
            with lock:
                estado["ejecutando"] = False
                estado["fin"] = time.strftime("%Y-%m-%d %H:%M:%S")

    threading.Thread(target=correr, daemon=True).start()
    return True


def ejecutar_comando(args):
    try:
        resultado = subprocess.run(
            [PYTHON_BOT, BOT_MAIN] + args,
            cwd=BOT_DIR,
            env=entorno_bot(),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=CREATE_NO_WINDOW,
            timeout=60
        )
    except Exception as e:
        return {"ok": False, "error": "No fue posible ejecutar el bot: " + str(e)}

    lineas = [l for l in resultado.stdout.splitlines() if l.strip()]
    if not lineas:
        return {"ok": False, "error": "El bot no devolvió ninguna respuesta."}
    try:
        return json.loads(lineas[-1])
    except ValueError:
        return {"ok": False, "error": "Respuesta inesperada del bot: " + lineas[-1][:300]}


# Solo estas rutas se sirven como archivos estáticos. El resto de la carpeta (bot con sus
# credenciales .env, server.py, python-portable...) no debe ser descargable.
ESTATICOS_RAIZ = ("favicon.ico",)
ESTATICOS_CARPETAS = ("html", "css", "js", "img")


def ruta_estatica_publica(path):
    partes = [p for p in urllib.parse.unquote(path.split("?", 1)[0].split("#", 1)[0]).split("/") if p]
    if ".." in partes:
        return False
    if not partes:
        return True
    if len(partes) == 1:
        return partes[0] in ESTATICOS_RAIZ
    return partes[0] in ESTATICOS_CARPETAS


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def list_directory(self, path):
        self.send_error(404)

    def do_HEAD(self):
        if not ruta_estatica_publica(self.path):
            self.send_error(404)
            return
        super().do_HEAD()

    def end_headers(self):
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        super().end_headers()

    def _json(self, data, code=200, cabeceras=None):
        body = json.dumps(data).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        for nombre, valor in (cabeceras or {}).items():
            self.send_header(nombre, valor)
        self.end_headers()
        self.wfile.write(body)

    def _cuerpo_json(self):
        """Cuerpo JSON de la petición, o None (ya respondió el error) si es inválido."""
        largo = int(self.headers.get("Content-Length", 0))
        try:
            return json.loads(self.rfile.read(largo) or b"{}")
        except ValueError:
            self._json({"ok": False, "error": "Cuerpo de la petición inválido."}, 400)
            return None

    def _servir_excel(self, ruta_archivo):
        if not os.path.isfile(ruta_archivo):
            self._json({
                "ok": False,
                "error": f"No se encontró el archivo en: {ruta_archivo}"
            }, 404)
            return
        copia = os.path.join(CACHE_DIR, os.path.basename(str(ruta_archivo)))
        desde_copia = False
        try:
            with open(ruta_archivo, "rb") as f:
                contenido = f.read()
            try:   # copia de respaldo para cuando alguien tenga el Excel abierto
                os.makedirs(CACHE_DIR, exist_ok=True)
                with open(copia, "wb") as f:
                    f.write(contenido)
            except OSError:
                pass
        except PermissionError:
            # Excel (u OneDrive) tiene el archivo tomado: se entrega la última copia que se pudo leer
            if not os.path.isfile(copia):
                self._json({"ok": False, "error": "El archivo está abierto en otro programa (probablemente Excel) y aún no hay copia de respaldo. Ciérralo un momento."}, 423)
                return
            with open(copia, "rb") as f:
                contenido = f.read()
            desde_copia = True
        except Exception as e:
            self._json({"ok": False, "error": "Error al leer el archivo Excel: " + str(e)}, 500)
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Length", str(len(contenido)))
        if desde_copia:
            self.send_header("X-Desde-Copia", "1")
        self.end_headers()
        self.wfile.write(contenido)

    def do_GET(self):
        # El portal no tiene inicio de sesión: la raíz lleva directo al inicio
        if self.path.split("?", 1)[0] in ("/", "/index.html"):
            self.send_response(302)
            self.send_header("Location", "/html/home.html")
            self.send_header("Content-Length", "0")
            self.end_headers()
            return

        # 1. Endpoint: Estado de modificación de planillas en OneDrive
        if self.path == "/api/excel-estado":
            self._json(obtener_estado_archivos())
            return

        # 2. Endpoint: Capacitaciones
        if self.path == "/api/academia/excel-capacitaciones":
            self._servir_excel(REGISTRO_CAPACITACIONES_PATH)
            return

        # 3. Endpoint: Reporte Maestro de Camiones
        if self.path == "/api/flota/excel-maestro":
            self._servir_excel(REPORTE_MAESTRO_PATH)
            return

        # 4. Endpoint: BBDD Clientes CM
        if self.path == "/api/clientes/excel-bbdd-cm":
            self._servir_excel(CLIENTES_CM_PATH)
            return

        # 5. Endpoint: Contactos
        if self.path == "/api/contactos/excel-contactos":
            self._servir_excel(CONTACTOS_PATH)
            return

        # Endpoints del Bot
        if self.path == "/api/bot/estado":
            with lock:
                self._json({
                    "ejecutando": estado["ejecutando"],
                    "inicio": estado["inicio"],
                    "fin": estado["fin"],
                    "codigo": estado["codigo"],
                    "log": estado["log"][-8000:]
                })
            return
        if self.path == "/api/bot/config-forms":
            actual = FORMS_FILENAME_DEFAULT
            if os.path.isfile(FORMS_FILE_CHOICE_PATH):
                try:
                    with open(FORMS_FILE_CHOICE_PATH, "r", encoding="utf-8") as f:
                        valor = f.read().strip()
                    if valor:
                        actual = valor
                except OSError:
                    pass
            self._json({
                "ok": True,
                "actual": actual,
                "opciones": FORMS_FILENAMES_DISPONIBLES
            })
            return
        if self.path == "/api/bot/nuevos":
            motivo = error_bot()
            if motivo:
                self._json({"ok": False, "error": motivo}, 500)
                return
            with lock:
                ocupado = estado["ejecutando"]
            if ocupado:
                self._json({"ok": False, "error": "El bot está ejecutándose, espera a que termine."}, 409)
                return
            self._json(ejecutar_comando(["--ver-nuevos"]))
            return
        if not ruta_estatica_publica(self.path):
            self.send_error(404)
            return
        super().do_GET()

    def do_POST(self):
        if self.path == "/api/bot/gui":
            ok, mensaje = abrir_gui_bot()
            self._json({"ok": ok, "mensaje": mensaje, "error": None if ok else mensaje}, 200 if ok else 500)
            return
        if self.path == "/api/bot/ejecutar":
            motivo = error_bot()
            if motivo:
                self._json({"ok": False, "error": motivo}, 500)
                return
            iniciado = ejecutar_bot()
            if not iniciado:
                self._json({"ok": False, "error": "El bot ya está en ejecución."}, 409)
                return
            self._json({"ok": True})
            return
        if self.path == "/api/bot/config-forms":
            with lock:
                ocupado = estado["ejecutando"]
            if ocupado:
                self._json({"ok": False, "error": "El bot está ejecutándose, espera a que termine."}, 409)
                return
            largo = int(self.headers.get("Content-Length", 0))
            try:
                cuerpo = json.loads(self.rfile.read(largo) or b"{}")
            except ValueError:
                self._json({"ok": False, "error": "Cuerpo de la petición inválido."}, 400)
                return
            nombre = str(cuerpo.get("nombre", "")).strip()
            if not nombre or not nombre.lower().endswith(".xlsx"):
                self._json({"ok": False, "error": "El nombre debe terminar en .xlsx"}, 400)
                return
            if not os.path.isdir(BOT_DIR):
                self._json({"ok": False, "error": "No se encontró la carpeta del bot en " + BOT_DIR}, 500)
                return
            try:
                with open(FORMS_FILE_CHOICE_PATH, "w", encoding="utf-8") as f:
                    f.write(nombre)
            except OSError as e:
                self._json({"ok": False, "error": "No fue posible guardar la preferencia: " + str(e)}, 500)
                return
            self._json({"ok": True, "actual": nombre})
            return
        if self.path == "/api/bot/eliminar-registro":
            motivo = error_bot()
            if motivo:
                self._json({"ok": False, "error": motivo}, 500)
                return
            with lock:
                ocupado = estado["ejecutando"]
            if ocupado:
                self._json({"ok": False, "error": "El bot está ejecutándose, espera a que termine."}, 409)
                return
            largo = int(self.headers.get("Content-Length", 0))
            try:
                cuerpo = json.loads(self.rfile.read(largo) or b"{}")
            except ValueError:
                self._json({"ok": False, "error": "Cuerpo de la petición inválido."}, 400)
                return
            rut = str(cuerpo.get("rutPersona", "")).strip()
            fecha = str(cuerpo.get("fechaAtencion", "")).strip()
            if not rut:
                self._json({"ok": False, "error": "Falta el RUT del registro a eliminar."}, 400)
                return
            self._json(ejecutar_comando(["--eliminar-registro", rut, fecha]))
            return
        self.send_response(404)
        self.end_headers()

    def log_message(self, format, *args):
        pass


if __name__ == "__main__":
    puerto_pedido = PORT
    PORT = puerto_libre(PORT)
    if PORT != puerto_pedido:
        print("El puerto %d lo está usando otra aplicación; el portal usará el %d." % (puerto_pedido, PORT))
    with socketserver.ThreadingTCPServer((HOST, PORT), Handler) as httpd:
        print("Servidor local en http://localhost:%d" % PORT)
        motivo = error_bot()
        print("Bot de Academia: " + (motivo or "listo (%s)" % BOT_DIR))
        threading.Timer(1.0, lambda: webbrowser.open("http://localhost:%d/html/home.html" % PORT)).start()
        httpd.serve_forever()