# Portal Volvo Chile (versión local)

Portal web que corre en tu propio PC: flota y dashboard, consulta por VIN, Academia y el bot de certificados.
Todo está dentro de esta carpeta; no hay que instalar Python ni nada más.

## Antes de empezar
- **Windows** con **OneDrive** sincronizado.
- En `OneDrive - Volvo Group\Shortcuts` debe existir el acceso directo **Team OCT&P - 6. Registro de capacitaciones**. De ahí salen el Excel de Forms del bot y el registro de capacitaciones.
- **Internet la primera vez** (unos 700 MB: librerías del bot y Chromium).
- Descomprime en una **ruta corta**, por ejemplo `C:\volvo-portal`. Windows falla con rutas de más de 260 caracteres y algunas librerías traen archivos con nombres largos.

## Primer uso
1. Descomprime el zip.
2. Copia `bot\academia\.env.example` como `bot\academia\.env` y completa `USUARIO_VOLVO` y `CLAVE_VOLVO`. Pídeselas a quien administra el bot: **no van en el repo**.
3. Haz doble clic en `iniciar-local.bat`. La primera vez prepara el bot solo (varios minutos). Después abre el portal directo en el **Inicio** (http://localhost:8000). No hay inicio de sesión.

## Cómo funciona
- `server.py` es el servidor local. Lee los Excel de OneDrive y los entrega a la página; el navegador nunca abre archivos de tu disco.
- **Sin inicio de sesión:** hay un solo perfil con acceso a todo. El servidor solo acepta conexiones de este mismo PC.
- **Datos:** la página revisa cada minuto si los Excel cambiaron y los carga sola. Los de Conectividad (flota, clientes, contactos) salen de `Team OCT&P - 04.- Conectividad`. Si esa carpeta no está sincronizada en tu PC, deja copias en `datos\` (ver `datos\LEEME.txt`); ahí no se actualizan solos.
- **Bot de Academia:** pestaña *Academia → Bot Certificados*. El botón *Ejecutar bot ahora* corre `bot\academia\main.py`; *Abrir interfaz del bot* abre la ventana de escritorio. Lo que edites en `bot\academia` se aplica en la siguiente ejecución.
- **Modo oscuro:** el botón redondo (luna) es arrastrable.

## Archivos importantes
| Archivo | Para qué |
|---|---|
| `iniciar-local.bat` | Arranca el portal (y prepara el bot si falta algo) |
| `instalar-bot.bat` | Instala/repara el entorno del bot |
| `abrir-gui-bot.bat` | Abre la interfaz del bot sin pasar por el portal |
| `empaquetar.py` | Genera el zip liviano para compartir |

## Si algo falla
- **"El entorno del bot no está instalado"**: cierra y abre `iniciar-local.bat` con internet.
- **No aparecen datos**: revisa que exista el acceso directo de OneDrive de arriba o que haya copias en `datos\`.
- **Puerto 8000 ocupado**: ya hay un servidor abierto; ciérralo o usa la variable `VOLVO_PORT`.
