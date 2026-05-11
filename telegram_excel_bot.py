"""Bot de Telegram que escribe en tablas-calendario de Excel en OneDrive.

Cada tabla tiene 364+ columnas (una por día del año) y una sola fila de datos.
El bot localiza la columna del día (por fecha o "hoy") y actualiza esa celda.

Comandos configurados via .env (EXCEL_CMD_N + EXCEL_TABLE_N):
  /paquetes [fecha] valor       suma 'valor' al día (default hoy)
  /gasolina [fecha] valor
  /aceite_filtro [fecha] valor
  /mantenimiento [fecha] valor

Otros:
  /poner comando [fecha] valor  reemplaza (no suma)
  /ver comando [fecha]          muestra el valor actual del día
  /tablas                       lista los comandos configurados
  /id                           tu user_id
  /login                        re-autenticación con Microsoft

Formato:
  valor: 169 ó 98,03 ó 98.03
  fecha: hoy si no la indicás, o '11/05', '11-05', '2026-05-11', '11/05/2026'
"""
from __future__ import annotations

import logging
import os
import re
import sys
from datetime import date, datetime, timedelta
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
COMMAND_TO_TABLE: dict[str, str] = {}
COMMAND_ORDER: list[str] = []
for i in range(1, 5):
    cmd = os.getenv(f"EXCEL_CMD_{i}", "").strip().lower()
    table = os.getenv(f"EXCEL_TABLE_{i}", "").strip()
    if cmd and table:
        if not CMD_RE.match(cmd):
            print(f"ERROR: EXCEL_CMD_{i}='{cmd}' inválido", file=sys.stderr)
            sys.exit(1)
        COMMAND_TO_TABLE[cmd] = table
        COMMAND_ORDER.append(cmd)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"]
TOKEN_CACHE_FILE = Path(__file__).with_name("ms_token_cache.json")
EXCEL_EPOCH = date(1899, 12, 30)
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
                "%d/%m", "%d-%m"]

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
log = logging.getLogger("telegram-excel-bot")
logging.getLogger("httpx").setLevel(logging.WARNING)


# ---------- Microsoft Graph ----------
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
        self._header_cache: dict[str, list[Any]] = {}

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
                raise RuntimeError(f"Device flow falló: {flow}")
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

    def _hdrs(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self.acquire_token()}",
                "Content-Type": "application/json", "Accept": "application/json"}

    def _table_url(self, table: str) -> str:
        return f"{GRAPH_BASE}/me/drive/root:/{EXCEL_FILE_PATH}:/workbook/tables/{table}"

    def get_header_values(self, table: str, refresh: bool = False) -> list[Any]:
        if not refresh and table in self._header_cache:
            return self._header_cache[table]
        r = requests.get(f"{self._table_url(table)}/headerRowRange",
                         headers=self._hdrs(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph headers {r.status_code}: {r.text}")
        values = r.json().get("values", [[]])[0]
        self._header_cache[table] = values
        return values

    def get_column_body(self, table: str, col_index: int) -> dict[str, Any]:
        r = requests.get(
            f"{self._table_url(table)}/columns/itemAt(index={col_index})/dataBodyRange",
            headers=self._hdrs(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph getCol {r.status_code}: {r.text}")
        return r.json()

    def patch_column_body(self, table: str, col_index: int,
                          new_values: list[list[Any]]) -> None:
        r = requests.patch(
            f"{self._table_url(table)}/columns/itemAt(index={col_index})/dataBodyRange",
            headers=self._hdrs(), json={"values": new_values}, timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph patch {r.status_code}: {r.text}")


graph: GraphClient


# ---------- Helpers ----------
def to_excel_serial(d: date) -> int:
    return (d - EXCEL_EPOCH).days


def from_excel_serial(n: float) -> date:
    return EXCEL_EPOCH + timedelta(days=int(n))


def parse_date_token(s: str) -> date | None:
    """Convierte '11/05', '11-05-2026', '2026-05-11', etc. en date. None si no parsea."""
    s = s.strip()
    if not s:
        return None
    for fmt in DATE_FORMATS:
        try:
            d = datetime.strptime(s, fmt).date()
            if "%Y" not in fmt and "%y" not in fmt:
                d = d.replace(year=date.today().year)
            return d
        except ValueError:
            continue
    return None


def parse_number(s: str) -> float | None:
    try:
        return float(s.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return None


def parse_args(args: list[str]) -> tuple[date, float | None]:
    """De los args devuelve (fecha, valor).

    Casos:
      [valor]           -> (hoy, valor)
      [fecha, valor]    -> (fecha, valor)
      [valor, fecha]    -> también acepta orden invertido
    """
    if not args:
        return date.today(), None
    if len(args) == 1:
        v = parse_number(args[0])
        return date.today(), v
    # 2+ args: intentar (fecha, valor)
    d1 = parse_date_token(args[0])
    v2 = parse_number(args[1])
    if d1 is not None and v2 is not None:
        return d1, v2
    # intentar (valor, fecha)
    v1 = parse_number(args[0])
    d2 = parse_date_token(args[1])
    if v1 is not None and d2 is not None:
        return d2, v1
    return date.today(), None


def find_col_index_for_date(headers: list[Any], target: date) -> int:
    """Busca la columna cuyo encabezado coincide con la fecha objetivo."""
    target_serial = to_excel_serial(target)
    for i, h in enumerate(headers):
        if isinstance(h, (int, float)):
            if int(h) == target_serial:
                return i
        elif isinstance(h, str):
            d = parse_date_token(h)
            if d is None:
                continue
            if (d.month, d.day) == (target.month, target.day):
                return i
    return -1


def is_authorized(update: Update) -> bool:
    if not ALLOWED_USER_IDS:
        return False
    return bool(update.effective_user and update.effective_user.id in ALLOWED_USER_IDS)


async def deny(update: Update) -> None:
    uid = update.effective_user.id if update.effective_user else "?"
    if update.message:
        await update.message.reply_text(
            f"No autorizado. Tu ID es `{uid}`.", parse_mode=ParseMode.MARKDOWN)


def tables_summary() -> str:
    if not COMMAND_ORDER:
        return "(sin tablas)"
    return "\n".join(f"  /{c}  →  {COMMAND_TO_TABLE[c]}" for c in COMMAND_ORDER)


def _coerce_current(v: Any) -> float | None:
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        return parse_number(v)
    return None


def write_value(cmd: str, target_date: date, value: float, mode: str
                ) -> tuple[float | None, float]:
    """mode = 'sum' | 'replace'. Devuelve (valor_anterior, valor_nuevo)."""
    table = COMMAND_TO_TABLE[cmd]
    headers = graph.get_header_values(table)
    idx = find_col_index_for_date(headers, target_date)
    if idx < 0:
        # Reintenta con cache fresco
        headers = graph.get_header_values(table, refresh=True)
        idx = find_col_index_for_date(headers, target_date)
    if idx < 0:
        raise RuntimeError(
            f"No encontré la columna del {target_date.isoformat()} en {table}.")

    body = graph.get_column_body(table, idx)
    values = body.get("values") or [[None]]
    row_count = body.get("rowCount", len(values))
    current = values[0][0] if values and values[0] else None
    current_num = _coerce_current(current)

    if mode == "sum":
        new = (current_num or 0.0) + value
    else:
        new = value

    new_values = [row[:] for row in values] if values else [[None]]
    if not new_values:
        new_values = [[None]] * max(row_count, 1)
    new_values[0][0] = new
    graph.patch_column_body(table, idx, new_values)
    return current_num, new


def read_value(cmd: str, target_date: date) -> tuple[int, Any]:
    table = COMMAND_TO_TABLE[cmd]
    headers = graph.get_header_values(table)
    idx = find_col_index_for_date(headers, target_date)
    if idx < 0:
        headers = graph.get_header_values(table, refresh=True)
        idx = find_col_index_for_date(headers, target_date)
    if idx < 0:
        raise RuntimeError(
            f"No encontré la columna del {target_date.isoformat()} en {table}.")
    body = graph.get_column_body(table, idx)
    values = body.get("values") or [[None]]
    return idx, (values[0][0] if values and values[0] else None)


# ---------- Handlers ----------
async def cmd_start(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    await update.message.reply_text(
        "Hola. Tablas configuradas:\n" + tables_summary() +
        "\n\nEjemplos:\n"
        "  `/paquetes 169`           suma 169 al día de hoy\n"
        "  `/gasolina 98,03`         suma 98.03 (acepta coma decimal)\n"
        "  `/paquetes 11/05 50`      suma al 11 de mayo\n"
        "  `/poner paquetes 200`     REEMPLAZA el valor de hoy\n"
        "  `/ver paquetes`           muestra el valor de hoy\n"
        "  `/ver gasolina 11/05`     muestra el valor del 11/05\n\n"
        "Más: /tablas /id /login",
        parse_mode=ParseMode.MARKDOWN)


async def cmd_id(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    uid = update.effective_user.id if update.effective_user else "?"
    await update.message.reply_text(f"Tu user_id: `{uid}`", parse_mode=ParseMode.MARKDOWN)


async def cmd_tablas(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    await update.message.reply_text("Tablas:\n" + tables_summary())


async def cmd_login(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    await update.message.reply_text("Mira la consola del PC.")
    try:
        graph.acquire_token(force_interactive=True)
        await update.message.reply_text("Autenticación OK.")
    except Exception as e:
        await update.message.reply_text(f"Error: {e}")


def make_table_handler(cmd: str):
    async def _h(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if not is_authorized(update): await deny(update); return
        target, value = parse_args(list(context.args or []))
        if value is None:
            await update.message.reply_text(
                f"Uso: /{cmd} [fecha opcional] valor\n"
                f"Ej: /{cmd} 169   o   /{cmd} 11/05 98,03"); return
        try:
            old, new = write_value(cmd, target, value, mode="sum")
        except Exception as e:
            log.exception("write")
            await update.message.reply_text(f"Error: {e}"); return
        old_txt = f"{old}" if old is not None else "(vacío)"
        await update.message.reply_text(
            f"OK /{cmd} {target.strftime('%d/%m/%Y')}\n"
            f"  anterior: {old_txt}\n"
            f"  +{value}\n"
            f"  nuevo: {new}")
    return _h


async def cmd_poner(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    args = list(context.args or [])
    if not args or args[0].lower() not in COMMAND_TO_TABLE:
        await update.message.reply_text(
            "Uso: /poner <comando> [fecha] valor\n" + tables_summary()); return
    cmd = args.pop(0).lower()
    target, value = parse_args(args)
    if value is None:
        await update.message.reply_text(f"Falta valor. Ej: /poner {cmd} 200"); return
    try:
        old, new = write_value(cmd, target, value, mode="replace")
    except Exception as e:
        log.exception("poner")
        await update.message.reply_text(f"Error: {e}"); return
    old_txt = f"{old}" if old is not None else "(vacío)"
    await update.message.reply_text(
        f"OK /{cmd} {target.strftime('%d/%m/%Y')}\n"
        f"  anterior: {old_txt}\n"
        f"  reemplazado por: {new}")


async def cmd_ver(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    args = list(context.args or [])
    if not args or args[0].lower() not in COMMAND_TO_TABLE:
        await update.message.reply_text(
            "Uso: /ver <comando> [fecha]\n" + tables_summary()); return
    cmd = args.pop(0).lower()
    target = date.today()
    if args:
        d = parse_date_token(args[0])
        if d is not None:
            target = d
    try:
        idx, val = read_value(cmd, target)
    except Exception as e:
        await update.message.reply_text(f"Error: {e}"); return
    val_txt = f"{val}" if val not in (None, "") else "(vacío)"
    await update.message.reply_text(
        f"/{cmd} {target.strftime('%d/%m/%Y')} (col {idx}): {val_txt}")


async def on_text(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    if not update.message or not update.message.text: return
    await update.message.reply_text(
        "Usá uno de los comandos:\n" + tables_summary())


async def on_error(_u: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.exception("Excepción", exc_info=context.error)


def _validate_config() -> None:
    missing = []
    if not TELEGRAM_BOT_TOKEN: missing.append("TELEGRAM_BOT_TOKEN")
    if not MS_CLIENT_ID: missing.append("MS_CLIENT_ID")
    if not EXCEL_FILE_PATH: missing.append("EXCEL_FILE_PATH")
    if not ALLOWED_USER_IDS: missing.append("ALLOWED_USER_IDS")
    if not COMMAND_TO_TABLE: missing.append("EXCEL_CMD_1+EXCEL_TABLE_1")
    if missing:
        print("Faltan en .env: " + ", ".join(missing), file=sys.stderr)
        sys.exit(1)


def main() -> None:
    global graph
    _validate_config()
    graph = GraphClient(MS_CLIENT_ID, MS_TENANT_ID)
    graph.acquire_token()
    log.info("Archivo: %s", EXCEL_FILE_PATH)
    log.info("Comandos: %s", COMMAND_TO_TABLE)

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("id", cmd_id))
    app.add_handler(CommandHandler("tablas", cmd_tablas))
    app.add_handler(CommandHandler("login", cmd_login))
    app.add_handler(CommandHandler("poner", cmd_poner))
    app.add_handler(CommandHandler("ver", cmd_ver))
    for cmd in COMMAND_ORDER:
        app.add_handler(CommandHandler(cmd, make_table_handler(cmd)))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_error_handler(on_error)

    log.info("Bot corriendo. Ctrl+C para detener.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
