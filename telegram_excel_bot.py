"""Bot de Telegram que agrega filas a tablas de Excel en OneDrive vía Microsoft Graph.

Cada tabla tiene su propio comando, configurado en .env:
  EXCEL_CMD_1 / EXCEL_TABLE_1   ej: paquetes      -> Tabla1
  EXCEL_CMD_2 / EXCEL_TABLE_2   ej: gasolina      -> Tabla13
  EXCEL_CMD_3 / EXCEL_TABLE_3   ej: aceite_filtro -> Tabla14
  EXCEL_CMD_4 / EXCEL_TABLE_4   ej: mantenimiento -> Tabla15

Uso desde Telegram:
  /paquetes valor1, valor2, valor3
  /gasolina 2026-05-10, 100, Shell, 50
  /aceite_filtro ...
  /mantenimiento ...

  /usar paquetes        - fija la tabla activa
  valor1, valor2, ...   - va a la tabla activa
  /headers              - encabezados de la tabla activa
  /headers gasolina     - encabezados de otra tabla
  /last [comando]       - última fila guardada
  /tablas               - lista todas las tablas y comandos
  /id                   - tu user_id de Telegram
  /login                - fuerza re-autenticación con Microsoft
"""
from __future__ import annotations

import logging
import os
import re
import sys
from pathlib import Path
from typing import Any

import msal
import requests
from dotenv import load_dotenv
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import (
    Application, CommandHandler, ContextTypes, MessageHandler, filters,
)

load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
ALLOWED_USER_IDS = {
    int(x) for x in os.getenv("ALLOWED_USER_IDS", "").split(",") if x.strip().isdigit()
}
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "").strip()
MS_TENANT_ID = os.getenv("MS_TENANT_ID", "consumers").strip()
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH", "").strip().lstrip("/")

CMD_RE = re.compile(r"^[a-zA-Z][a-zA-Z0-9_]{0,31}$")

# Mapeo comando -> nombre de Tabla en Excel
COMMAND_TO_TABLE: dict[str, str] = {}
# Mantiene orden de inserción para listar y para DEFAULT por número
COMMAND_ORDER: list[str] = []

for i in range(1, 5):
    cmd = os.getenv(f"EXCEL_CMD_{i}", "").strip().lower()
    table = os.getenv(f"EXCEL_TABLE_{i}", "").strip()
    if cmd and table:
        if not CMD_RE.match(cmd):
            print(f"ERROR: EXCEL_CMD_{i}='{cmd}' inválido. "
                  "Solo letras/dígitos/_ y empezar con letra.", file=sys.stderr)
            sys.exit(1)
        COMMAND_TO_TABLE[cmd] = table
        COMMAND_ORDER.append(cmd)

DEFAULT_TABLE_IDX = int(os.getenv("EXCEL_TABLE_DEFAULT", "1") or "1")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"]
TOKEN_CACHE_FILE = Path(__file__).with_name("ms_token_cache.json")

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
log = logging.getLogger("telegram-excel-bot")
logging.getLogger("httpx").setLevel(logging.WARNING)

# Estado: comando "activo" por usuario
active_cmd_by_user: dict[int, str] = {}


class GraphClient:
    def __init__(self, client_id: str, tenant_id: str) -> None:
        if not client_id:
            raise RuntimeError("MS_CLIENT_ID no configurado")
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.cache = msal.SerializableTokenCache()
        if TOKEN_CACHE_FILE.exists():
            self.cache.deserialize(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))
        self.app = msal.PublicClientApplication(
            client_id, authority=self.authority, token_cache=self.cache)

    def _save_cache(self) -> None:
        if self.cache.has_state_changed:
            TOKEN_CACHE_FILE.write_text(self.cache.serialize(), encoding="utf-8")

    def acquire_token(self, force_interactive: bool = False) -> str:
        result = None
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
            raise RuntimeError(f"Sin access_token: {result.get('error_description', result)}")
        self._save_cache()
        return result["access_token"]

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self.acquire_token()}",
                "Content-Type": "application/json", "Accept": "application/json"}

    def _table_url(self, file_path: str, table: str) -> str:
        return f"{GRAPH_BASE}/me/drive/root:/{file_path}:/workbook/tables/{table}"

    def add_row(self, file_path: str, table: str, values: list[Any]) -> dict[str, Any]:
        r = requests.post(f"{self._table_url(file_path, table)}/rows/add",
                          headers=self._headers(), json={"values": [values]}, timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.status_code}: {r.text}")
        return r.json()

    def get_headers_row(self, file_path: str, table: str) -> list[str]:
        r = requests.get(f"{self._table_url(file_path, table)}/headerRowRange",
                         headers=self._headers(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.status_code}: {r.text}")
        return r.json().get("values", [[]])[0]

    def get_last_row(self, file_path: str, table: str) -> list[Any] | None:
        r = requests.get(f"{self._table_url(file_path, table)}/rows",
                         headers=self._headers(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {r.status_code}: {r.text}")
        rows = r.json().get("value", [])
        return rows[-1].get("values", [[]])[0] if rows else None


graph: GraphClient


def parse_row(text: str) -> list[str]:
    raw = text.replace(";", ",").replace("|", ",")
    return [p.strip() for p in raw.split(",") if p.strip()]


def is_authorized(update: Update) -> bool:
    if not ALLOWED_USER_IDS:
        return False
    return bool(update.effective_user and update.effective_user.id in ALLOWED_USER_IDS)


async def deny(update: Update) -> None:
    uid = update.effective_user.id if update.effective_user else "?"
    if update.message:
        await update.message.reply_text(
            f"No autorizado. Tu ID es `{uid}`. Agregalo a ALLOWED_USER_IDS y reinicia.",
            parse_mode=ParseMode.MARKDOWN)


def default_cmd() -> str | None:
    if not COMMAND_ORDER:
        return None
    idx = max(1, min(DEFAULT_TABLE_IDX, len(COMMAND_ORDER))) - 1
    return COMMAND_ORDER[idx]


def active_cmd_for(user_id: int) -> str | None:
    return active_cmd_by_user.get(user_id) or default_cmd()


def tables_summary() -> str:
    if not COMMAND_ORDER:
        return "(sin tablas)"
    return "\n".join(f"  /{c}  →  {COMMAND_TO_TABLE[c]}" for c in COMMAND_ORDER)


def resolve_target_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> tuple[str | None, list[str]]:
    """Si el primer arg es un comando conocido, lo usa. Si no, devuelve el activo."""
    args = list(context.args or [])
    if args and args[0].lower() in COMMAND_TO_TABLE:
        return args.pop(0).lower(), args
    uid = update.effective_user.id if update.effective_user else 0
    return active_cmd_for(uid), args


# ---------- Handlers ----------
async def cmd_start(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    uid = update.effective_user.id
    active = active_cmd_for(uid) or "?"
    await update.message.reply_text(
        "Hola. Tablas configuradas:\n" + tables_summary() +
        f"\n\nTabla activa: /{active} → {COMMAND_TO_TABLE.get(active, '?')}\n\n"
        "Ejemplos:\n"
        f"  `/{active} 2026-05-10, Cliente, 1500`\n"
        "  `/usar gasolina`  cambia la tabla activa\n"
        "  `valor1, valor2, valor3`  va a la tabla activa\n\n"
        "Más: /tablas /headers /last /id /login",
        parse_mode=ParseMode.MARKDOWN)


async def cmd_id(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    uid = update.effective_user.id if update.effective_user else "?"
    await update.message.reply_text(f"Tu user_id: `{uid}`", parse_mode=ParseMode.MARKDOWN)


async def cmd_tablas(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    await update.message.reply_text("Tablas:\n" + tables_summary())


async def cmd_usar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    if not context.args:
        await update.message.reply_text("Uso: /usar paquetes\n" + tables_summary()); return
    target = context.args[0].lower().lstrip("/")
    if target not in COMMAND_TO_TABLE:
        await update.message.reply_text("No existe.\n" + tables_summary()); return
    active_cmd_by_user[update.effective_user.id] = target
    await update.message.reply_text(f"Tabla activa: /{target} → {COMMAND_TO_TABLE[target]}")


async def cmd_headers(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    target, _rest = resolve_target_cmd(update, context)
    if not target:
        await update.message.reply_text("Sin tabla activa."); return
    try:
        hdrs = graph.get_headers_row(EXCEL_FILE_PATH, COMMAND_TO_TABLE[target])
        await update.message.reply_text(
            f"Encabezados /{target} ({COMMAND_TO_TABLE[target]}):\n" + " | ".join(map(str, hdrs)))
    except Exception as e:
        await update.message.reply_text(f"Error: {e}")


async def cmd_last(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    target, _rest = resolve_target_cmd(update, context)
    if not target:
        await update.message.reply_text("Sin tabla activa."); return
    try:
        row = graph.get_last_row(EXCEL_FILE_PATH, COMMAND_TO_TABLE[target])
        if row is None:
            await update.message.reply_text(f"/{target} está vacía.")
        else:
            await update.message.reply_text(
                f"Última fila /{target}:\n" + " | ".join(map(str, row)))
    except Exception as e:
        await update.message.reply_text(f"Error: {e}")


async def cmd_login(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    await update.message.reply_text("Mira la consola para completar el login.")
    try:
        graph.acquire_token(force_interactive=True)
        await update.message.reply_text("Autenticación OK.")
    except Exception as e:
        await update.message.reply_text(f"Error: {e}")


def make_table_handler(cmd: str):
    async def _handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if not is_authorized(update): await deny(update); return
        text = " ".join(context.args) if context.args else ""
        if not text:
            await update.message.reply_text(f"Uso: /{cmd} valor1, valor2, ..."); return
        await _append_row(update, cmd, text)
    return _handler


async def on_text(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    if not update.message or not update.message.text: return
    target = active_cmd_for(update.effective_user.id)
    if not target:
        await update.message.reply_text(
            "Sin tabla activa. Usá uno de los comandos:\n" + tables_summary())
        return
    await _append_row(update, target, update.message.text)


async def _append_row(update: Update, cmd: str, text: str) -> None:
    values = parse_row(text)
    if not values:
        await update.message.reply_text("No detecté valores. Separá por coma."); return
    table = COMMAND_TO_TABLE[cmd]
    try:
        graph.add_row(EXCEL_FILE_PATH, table, values)
        await update.message.reply_text(
            f"OK /{cmd} ({table}):\n" + " | ".join(values))
    except Exception as e:
        log.exception("add_row")
        await update.message.reply_text(f"Error guardando: {e}")


async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.exception("Excepción", exc_info=context.error)


def _validate_config() -> None:
    missing = []
    if not TELEGRAM_BOT_TOKEN: missing.append("TELEGRAM_BOT_TOKEN")
    if not MS_CLIENT_ID: missing.append("MS_CLIENT_ID")
    if not EXCEL_FILE_PATH: missing.append("EXCEL_FILE_PATH")
    if not ALLOWED_USER_IDS: missing.append("ALLOWED_USER_IDS")
    if not COMMAND_TO_TABLE: missing.append("EXCEL_CMD_1 + EXCEL_TABLE_1 (al menos uno)")
    if missing:
        print("Faltan en .env: " + ", ".join(missing), file=sys.stderr)
        sys.exit(1)


def main() -> None:
    global graph
    _validate_config()
    graph = GraphClient(MS_CLIENT_ID, MS_TENANT_ID)
    graph.acquire_token()
    log.info("Archivo: %s", EXCEL_FILE_PATH)
    log.info("Mapeo comando→tabla: %s", COMMAND_TO_TABLE)

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()

    # Comandos fijos
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("id", cmd_id))
    app.add_handler(CommandHandler("usar", cmd_usar))
    app.add_handler(CommandHandler("tablas", cmd_tablas))
    app.add_handler(CommandHandler("headers", cmd_headers))
    app.add_handler(CommandHandler("last", cmd_last))
    app.add_handler(CommandHandler("login", cmd_login))

    # Un handler por cada comando configurado (paquetes, gasolina, etc.)
    for cmd in COMMAND_ORDER:
        app.add_handler(CommandHandler(cmd, make_table_handler(cmd)))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_error_handler(on_error)

    log.info("Bot corriendo. Ctrl+C para detener.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
