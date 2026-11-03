"""
Punto de entrada principal del Bot de Telegram — CEA Auto Oriente.
Incluye servidor HTTP ligero (Health Check) para despliegue en Render Web Service.
"""

import logging
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters

from config.settings import ASSETS_DIR, PORT, TELEGRAM_BOT_TOKEN
from handlers.admin_handler import cmd_ayuda, cmd_inicializar, cmd_reporte
from handlers.report_handler import procesar_reporte_foto

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


class HealthCheckHandler(BaseHTTPRequestHandler):
    """Responde 200 OK para que Render Web Service mantenga el servicio activo."""

    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write("Bot CEA Auto Oriente activo y funcionando (200 OK)".encode("utf-8"))

    def log_message(self, format, *args):
        # Silenciar logs de cada ping de Render para mantener limpia la consola
        return


def iniciar_servidor_healthcheck() -> None:
    """Inicia el servidor HTTP en segundo plano en el puerto asignado por Render (PORT)."""
    try:
        server = HTTPServer(("0.0.0.0", PORT), HealthCheckHandler)
        hilo = threading.Thread(target=server.serve_forever, daemon=True)
        hilo.start()
        logger.info("🌐 Servidor Health Check escuchando en el puerto %s (Render Web Service)", PORT)
    except Exception as exc:
        logger.warning("No se pudo iniciar el servidor Health Check en puerto %s: %s", PORT, exc)


def main() -> None:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)

    if not TELEGRAM_BOT_TOKEN or TELEGRAM_BOT_TOKEN == "TU_TELEGRAM_BOT_TOKEN":
        logger.error(
            "⚠️ Debes configurar TELEGRAM_BOT_TOKEN en el archivo .env o en las Variables de Entorno de Render."
        )
        sys.exit(1)

    # 1. Iniciar servidor HTTP en segundo plano para Render
    iniciar_servidor_healthcheck()

    # 2. Configurar y arrancar el Bot de Telegram
    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

    # Comandos administrativos
    app.add_handler(CommandHandler(["start", "ayuda", "help"], cmd_ayuda))
    app.add_handler(CommandHandler("inicializar", cmd_inicializar))
    app.add_handler(CommandHandler("reporte", cmd_reporte))

    # Manejador de reportes fotográficos (escucha fotos con caption que mencionen 'clase' o 'cl')
    filtro_reporte = filters.PHOTO & filters.CaptionRegex(r"(?i)\b(clases?|cl\.?)\b")
    app.add_handler(MessageHandler(filtro_reporte, procesar_reporte_foto))

    logger.info("🚗 Bot CEA Auto Oriente iniciado y escuchando mensajes...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
