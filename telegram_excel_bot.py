"""Bot de Telegram que agrega filas a un Excel de OneDrive vía Microsoft Graph.

Flujo:
  1. El usuario manda al bot una línea con valores separados por coma o ';'.
     Ej: "2026-05-10, EURUSD, BUY, 1.0850, 1.0900"
  2. El bot autentica contra Microsoft Graph (device-code la primera vez,
     luego usa cache de tokens) y agrega la fila a la Tabla configurada.
  3. El proceso queda corriendo en tu PC con polling de Telegram.

Comandos:
  /start         - mensaje de bienvenida
  /id            - muestra tu user_id de Telegram
  /add a,b,c     - agrega fila (alternativa a mandar texto suelto)
  /last          - muestra la última fila guardada
  /headers       - muestra los encabezados de la tabla
  /login         - fuerza re-autenticación con Microsoft
"""
from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path
from typing import Any

import msal
import requests
from dotenv import load_dotenv
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# ---------- Configuración ----------
load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
ALLOWED_USER_IDS = {
    int(x) for x in os.getenv("ALLOWED_USER_IDS", "").split(",") if x.strip().isdigit()
}
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "").strip()
MS_TENANT_ID = os.getenv("MS_TENANT_ID", "consumers").strip()
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH", "").strip().lstrip("/")
EXCEL_TABLE_NAME = os.getenv("EXCEL_TABLE_NAME", "Tabla1").strip()
EXCEL_WORKSHEET = os.getenv("EXCEL_WORKSHEET", "Hoja1").strip()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"]
TOKEN_CACHE_FILE = Path(__file__).with_name("ms_token_cache.json")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("telegram-excel-bot")
# Bajar ruido del cliente HTTP de Telegram
logging.getLogger("httpx").setLevel(logging.WARNING)


# ---------- Microsoft Graph ----------
class GraphClient:
    """Cliente mínimo para Graph + autenticación MSAL con device code."""

    def __init__(self, client_id: str, tenant_id: str) -> None:
        if not client_id:
            raise RuntimeError("MS_CLIENT_ID no configurado en .env")
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.cache = msal.SerializableTokenCache()
        if TOKEN_CACHE_FILE.exists():
            self.cache.deserialize(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))
        self.app = msal.PublicClientApplication(
            client_id, authority=self.authority, token_cache=self.cache
        )

    def _save_cache(self) -> None:
        if self.cache.has_state_changed:
            TOKEN_CACHE_FILE.write_text(self.cache.serialize(), encoding="utf-8")

    def acquire_token(self, force_interactive: bool = False) -> str:
        """Devuelve un access_token válido. Usa device-code si hace falta."""
        result: dict[str, Any] | None = None
        accounts = self.app.get_accounts()
        if accounts and not force_interactive:
            result = self.app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])

        if not result:
            flow = self.app.initiate_device_flow(scopes=GRAPH_SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(f"No se pudo iniciar device flow: {flow}")
            print("\n" + "=" * 60)
            print("AUTENTICACIÓN MICROSOFT - HAZLO UNA SOLA VEZ")
            print("=" * 60)
            print(flow["message"])
            print("=" * 60 + "\n", flush=True)
            result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError(
                f"No se obtuvo access_token: {result.get('error_description', result)}"
            )
        self._save_cache()
        return result["access_token"]

    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self.acquire_token()}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    # -- Endpoints Excel --
    def _table_url(self, file_path: str, table: str) -> str:
        return f"{GRAPH_BASE}/me/drive/root:/{file_path}:/workbook/tables/{table}"

    def add_row(self, file_path: str, table: str, values: list[Any]) -> dict[str, Any]:
        url = f"{self._table_url(file_path, table)}/rows/add"
        payload = {"values": [values]}
        r = requests.post(url, headers=self._headers(), json=payload, timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.status_code}: {r.text}")
        return r.json()

    def get_headers(self, file_path: str, table: str) -> list[str]:
        url = f"{self._table_url(file_path, table)}/headerRowRange"
        r = requests.get(url, headers=self._headers(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.status_code}: {r.text}")
        data = r.json()
        return data.get("values", [[]])[0]

    def get_last_row(self, file_path: str, table: str) -> list[Any] | None:
        url = f"{self._table_url(file_path, table)}/rows"
        r = requests.get(url, headers=self._headers(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.status_code}: {r.text}")
        rows = r.json().get("value", [])
        if not rows:
            return None
        return rows[-1].get("values", [[]])[0]


graph: GraphClient  # se inicializa en main()


# ---------- Helpers ----------
def parse_row(text: str) -> list[str]:
    """Parsea 'a, b; c | d' -> ['a','b','c','d']. Acepta , ; y |."""
    raw = text.replace(";", ",").replace("|", ",")
    return [p.strip() for p in raw.split(",") if p.strip() != "" or True][:]


def is_authorized(update: Update) -> bool:
    if not ALLOWED_USER_IDS:
        # Sin lista blanca configurada: rechaza todo por seguridad
        return False
    user = update.effective_user
    return bool(user and user.id in ALLOWED_USER_IDS)


async def deny(update: Update) -> None:
    uid = update.effective_user.id if update.effective_user else "?"
    log.warning("Usuario no autorizado: %s", uid)
    if update.message:
        await update.message.reply_text(
            f"No autorizado. Tu ID es `{uid}`. "
            "Agrégalo a ALLOWED_USER_IDS en el .env y reinicia el bot.",
            parse_mode=ParseMode.MARKDOWN,
        )


# ---------- Handlers de Telegram ----------
async def cmd_start(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update)
        return
    await update.message.reply_text(
        "Hola. Mándame los valores separados por coma y los agrego al Excel.\n\n"
        "Ej: `2026-05-10, EURUSD, BUY, 1.0850, 1.0900`\n\n"
        "Comandos: /headers /last /add /id /login",
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_id(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    uid = update.effective_user.id if update.effective_user else "?"
    await update.message.reply_text(f"Tu user_id es: `{uid}`", parse_mode=ParseMode.MARKDOWN)


async def cmd_headers(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update)
        return
    try:
        headers = graph.get_headers(EXCEL_FILE_PATH, EXCEL_TABLE_NAME)
        await update.message.reply_text("Encabezados: " + " | ".join(map(str, headers)))
    except Exception as e:
        log.exception("headers error")
        await update.message.reply_text(f"Error: {e}")


async def cmd_last(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update)
        return
    try:
        row = graph.get_last_row(EXCEL_FILE_PATH, EXCEL_TABLE_NAME)
        if row is None:
            await update.message.reply_text("La tabla está vacía.")
        else:
            await update.message.reply_text("Última fila: " + " | ".join(map(str, row)))
    except Exception as e:
        log.exception("last error")
        await update.message.reply_text(f"Error: {e}")


async def cmd_login(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update)
        return
    await update.message.reply_text(
        "Mira la consola del PC para completar el login con código de dispositivo."
    )
    try:
        graph.acquire_token(force_interactive=True)
        await update.message.reply_text("Autenticación OK.")
    except Exception as e:
        log.exception("login error")
        await update.message.reply_text(f"Error de login: {e}")


async def cmd_add(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update)
        return
    text = " ".join(context.args) if context.args else ""
    if not text:
        await update.message.reply_text("Uso: /add valor1, valor2, valor3, ...")
        return
    await _append_row(update, text)


async def on_text(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update)
        return
    if not update.message or not update.message.text:
        return
    await _append_row(update, update.message.text)


async def _append_row(update: Update, text: str) -> None:
    values = parse_row(text)
    if not values:
        await update.message.reply_text("No detecté valores. Separa por coma.")
        return
    try:
        graph.add_row(EXCEL_FILE_PATH, EXCEL_TABLE_NAME, values)
        await update.message.reply_text("Fila agregada: " + " | ".join(values))
    except Exception as e:
        log.exception("add_row error")
        await update.message.reply_text(f"Error guardando en Excel: {e}")


async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.exception("Excepción en handler", exc_info=context.error)


# ---------- Main ----------
def _validate_config() -> None:
    missing = []
    if not TELEGRAM_BOT_TOKEN:
        missing.append("TELEGRAM_BOT_TOKEN")
    if not MS_CLIENT_ID:
        missing.append("MS_CLIENT_ID")
    if not EXCEL_FILE_PATH:
        missing.append("EXCEL_FILE_PATH")
    if not ALLOWED_USER_IDS:
        missing.append("ALLOWED_USER_IDS")
    if missing:
        print("Faltan variables en .env: " + ", ".join(missing), file=sys.stderr)
        sys.exit(1)


def main() -> None:
    global graph
    _validate_config()

    log.info("Inicializando cliente Microsoft Graph...")
    graph = GraphClient(MS_CLIENT_ID, MS_TENANT_ID)
    # Forzamos primer login si no hay cache, así el device code aparece al arrancar
    graph.acquire_token()
    log.info("Token Microsoft OK. Archivo objetivo: %s tabla=%s",
             EXCEL_FILE_PATH, EXCEL_TABLE_NAME)

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("id", cmd_id))
    app.add_handler(CommandHandler("headers", cmd_headers))
    app.add_handler(CommandHandler("last", cmd_last))
    app.add_handler(CommandHandler("login", cmd_login))
    app.add_handler(CommandHandler("add", cmd_add))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_error_handler(on_error)

    log.info("Bot corriendo. Ctrl+C para detener.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
