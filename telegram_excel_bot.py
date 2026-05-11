"""Bot de Telegram que agrega filas a tablas de Excel en OneDrive vía Microsoft Graph.

Soporta hasta 4 tablas en el mismo archivo. Cada tabla tiene su comando:
  /t1 valor1, valor2, ...   -> agrega fila a la tabla 1
  /t2 valor1, valor2, ...   -> agrega fila a la tabla 2
  /t3 ...
  /t4 ...

Sin comando: usa la tabla "activa" (cambia con /usar 1..4).

Otros comandos:
  /start         - bienvenida y resumen de tablas
  /id            - tu user_id de Telegram
  /usar N        - fija la tabla activa (1..4)
  /headers N     - encabezados de la tabla N (o de la activa)
  /last N        - última fila de la tabla N (o de la activa)
  /login         - fuerza re-autenticación con Microsoft
"""
from __future__ import annotations

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

# Hasta 4 tablas. Definí en el .env: EXCEL_TABLE_1, EXCEL_TABLE_2, etc.
TABLES: dict[int, str] = {}
for i in range(1, 5):
    name = os.getenv(f"EXCEL_TABLE_{i}", "").strip()
    if name:
        TABLES[i] = name

DEFAULT_TABLE = int(os.getenv("EXCEL_TABLE_DEFAULT", "1") or "1")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"]
TOKEN_CACHE_FILE = Path(__file__).with_name("ms_token_cache.json")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("telegram-excel-bot")
logging.getLogger("httpx").setLevel(logging.WARNING)

# Estado en memoria: tabla activa por usuario
active_table_by_user: dict[int, int] = {}


# ---------- Microsoft Graph ----------
class GraphClient:
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

    def _table_url(self, file_path: str, table: str) -> str:
        return f"{GRAPH_BASE}/me/drive/root:/{file_path}:/workbook/tables/{table}"

    def add_row(self, file_path: str, table: str, values: list[Any]) -> dict[str, Any]:
        url = f"{self._table_url(file_path, table)}/rows/add"
        r = requests.post(url, headers=self._headers(),
                          json={"values": [values]}, timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.status_code}: {r.text}")
        return r.json()

    def get_headers_row(self, file_path: str, table: str) -> list[str]:
        url = f"{self._table_url(file_path, table)}/headerRowRange"
        r = requests.get(url, headers=self._headers(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.status_code}: {r.text}")
        return r.json().get("values", [[]])[0]

    def get_last_row(self, file_path: str, table: str) -> list[Any] | None:
        url = f"{self._table_url(file_path, table)}/rows"
        r = requests.get(url, headers=self._headers(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.status_code}: {r.text}")
        rows = r.json().get("value", [])
        return rows[-1].get("values", [[]])[0] if rows else None


graph: GraphClient


# ---------- Helpers ----------
def parse_row(text: str) -> list[str]:
    raw = text.replace(";", ",").replace("|", ",")
    return [p.strip() for p in raw.split(",") if p.strip() != ""]


def is_authorized(update: Update) -> bool:
    if not ALLOWED_USER_IDS:
        return False
    user = update.effective_user
    return bool(user and user.id in ALLOWED_USER_IDS)


async def deny(update: Update) -> None:
    uid = update.effective_user.id if update.effective_user else "?"
    log.warning("No autorizado: %s", uid)
    if update.message:
        await update.message.reply_text(
            f"No autorizado. Tu ID es `{uid}`. Agrégalo a ALLOWED_USER_IDS en .env y reinicia.",
            parse_mode=ParseMode.MARKDOWN,
        )


def resolve_table_num(update: Update, context: ContextTypes.DEFAULT_TYPE,
                      arg_index: int = 0) -> tuple[int, list[str]]:
    """Devuelve (numero_tabla, args_restantes).

    Si el primer arg es '1'..'4' lo usa como selector. Si no, usa la activa del
    usuario o el DEFAULT_TABLE.
    """
    args = list(context.args or [])
    if args and args[arg_index].isdigit() and 1 <= int(args[arg_index]) <= 4 \
            and int(args[arg_index]) in TABLES:
        n = int(args.pop(arg_index))
        return n, args
    uid = update.effective_user.id if update.effective_user else 0
    n = active_table_by_user.get(uid, DEFAULT_TABLE)
    if n not in TABLES:
        n = next(iter(TABLES))
    return n, args


def tables_summary() -> str:
    if not TABLES:
        return "(ninguna tabla configurada)"
    return "\n".join(f"  /t{i}  →  {name}" for i, name in TABLES.items())


# ---------- Handlers ----------
async def cmd_start(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update); return
    uid = update.effective_user.id
    active = active_table_by_user.get(uid, DEFAULT_TABLE)
    await update.message.reply_text(
        "Hola. Tablas configuradas:\n" + tables_summary() +
        f"\n\nTabla activa: {active} ({TABLES.get(active, '?')})\n\n"
        "Ejemplos:\n"
        "  `/t1 2026-05-10, EURUSD, BUY, 1.0850`\n"
        "  `/usar 2`  (cambia la tabla activa)\n"
        "  `valor1, valor2, valor3`  (usa la tabla activa)\n\n"
        "Otros: /headers /last /id /login",
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_id(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    uid = update.effective_user.id if update.effective_user else "?"
    await update.message.reply_text(f"Tu user_id es: `{uid}`",
                                    parse_mode=ParseMode.MARKDOWN)


async def cmd_usar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update); return
    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text("Uso: /usar 1  (o 2, 3, 4)")
        return
    n = int(context.args[0])
    if n not in TABLES:
        await update.message.reply_text(
            f"No hay tabla {n}. Disponibles:\n" + tables_summary())
        return
    active_table_by_user[update.effective_user.id] = n
    await update.message.reply_text(f"Tabla activa: {n} → {TABLES[n]}")


async def cmd_headers(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update); return
    n, _ = resolve_table_num(update, context)
    try:
        hdrs = graph.get_headers_row(EXCEL_FILE_PATH, TABLES[n])
        await update.message.reply_text(
            f"Encabezados tabla {n} ({TABLES[n]}):\n" + " | ".join(map(str, hdrs)))
    except Exception as e:
        log.exception("headers error")
        await update.message.reply_text(f"Error: {e}")


async def cmd_last(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update); return
    n, _ = resolve_table_num(update, context)
    try:
        row = graph.get_last_row(EXCEL_FILE_PATH, TABLES[n])
        if row is None:
            await update.message.reply_text(f"Tabla {n} ({TABLES[n]}) está vacía.")
        else:
            await update.message.reply_text(
                f"Última fila tabla {n}:\n" + " | ".join(map(str, row)))
    except Exception as e:
        log.exception("last error")
        await update.message.reply_text(f"Error: {e}")


async def cmd_login(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update); return
    await update.message.reply_text("Mira la consola del PC para completar el login.")
    try:
        graph.acquire_token(force_interactive=True)
        await update.message.reply_text("Autenticación OK.")
    except Exception as e:
        log.exception("login error")
        await update.message.reply_text(f"Error de login: {e}")


def make_cmd_tn(n: int):
    async def _handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if not is_authorized(update):
            await deny(update); return
        if n not in TABLES:
            await update.message.reply_text(
                f"Tabla {n} no configurada. Disponibles:\n" + tables_summary())
            return
        text = " ".join(context.args) if context.args else ""
        if not text:
            await update.message.reply_text(f"Uso: /t{n} valor1, valor2, ...")
            return
        await _append_row(update, n, text)
    return _handler


async def on_text(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update):
        await deny(update); return
    if not update.message or not update.message.text:
        return
    uid = update.effective_user.id
    n = active_table_by_user.get(uid, DEFAULT_TABLE)
    if n not in TABLES:
        await update.message.reply_text(
            "No hay tabla activa válida. Usá /usar N o /t1.../t4.")
        return
    await _append_row(update, n, update.message.text)


async def _append_row(update: Update, table_num: int, text: str) -> None:
    values = parse_row(text)
    if not values:
        await update.message.reply_text("No detecté valores. Separá por coma.")
        return
    try:
        graph.add_row(EXCEL_FILE_PATH, TABLES[table_num], values)
        await update.message.reply_text(
            f"OK tabla {table_num} ({TABLES[table_num]}):\n" + " | ".join(values))
    except Exception as e:
        log.exception("add_row error")
        await update.message.reply_text(f"Error guardando en Excel: {e}")


async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.exception("Excepción en handler", exc_info=context.error)


# ---------- Main ----------
def _validate_config() -> None:
    missing = []
    if not TELEGRAM_BOT_TOKEN: missing.append("TELEGRAM_BOT_TOKEN")
    if not MS_CLIENT_ID: missing.append("MS_CLIENT_ID")
    if not EXCEL_FILE_PATH: missing.append("EXCEL_FILE_PATH")
    if not ALLOWED_USER_IDS: missing.append("ALLOWED_USER_IDS")
    if not TABLES: missing.append("EXCEL_TABLE_1 (al menos una tabla)")
    if missing:
        print("Faltan variables en .env: " + ", ".join(missing), file=sys.stderr)
        sys.exit(1)


def main() -> None:
    global graph
    _validate_config()
    log.info("Inicializando Microsoft Graph...")
    graph = GraphClient(MS_CLIENT_ID, MS_TENANT_ID)
    graph.acquire_token()
    log.info("Token OK. Archivo: %s. Tablas: %s", EXCEL_FILE_PATH, TABLES)

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("id", cmd_id))
    app.add_handler(CommandHandler("usar", cmd_usar))
    app.add_handler(CommandHandler("headers", cmd_headers))
    app.add_handler(CommandHandler("last", cmd_last))
    app.add_handler(CommandHandler("login", cmd_login))
    for i in range(1, 5):
        app.add_handler(CommandHandler(f"t{i}", make_cmd_tn(i)))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_error_handler(on_error)

    log.info("Bot corriendo. Ctrl+C para detener.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
