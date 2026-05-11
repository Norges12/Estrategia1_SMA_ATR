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
from urllib.parse import quote

import msal
import requests
from dotenv import load_dotenv
from telegram import Update
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
        path = quote(EXCEL_FILE_PATH, safe="/")
        # OData function syntax: tables('name') con comillas simples escapadas
        tbl_escaped = table.replace("'", "''")
        tbl_enc = quote(tbl_escaped, safe="'-_,{}.")
        return f"{GRAPH_BASE}/me/drive/root:/{path}:/workbook/tables('{tbl_enc}')"

    def list_children(self, folder: str = "") -> list[dict[str, Any]]:
        if folder:
            url = f"{GRAPH_BASE}/me/drive/root:/{quote(folder, safe='/')}:/children"
        else:
            url = f"{GRAPH_BASE}/me/drive/root/children"
        r = requests.get(url, headers=self._hdrs(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph children {r.status_code}: {r.text}")
        return r.json().get("value", [])

    def list_tables_full(self) -> list[dict[str, Any]]:
        path = quote(EXCEL_FILE_PATH, safe="/")
        url = f"{GRAPH_BASE}/me/drive/root:/{path}:/workbook/tables"
        r = requests.get(url, headers=self._hdrs(), timeout=30)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph listTables {r.status_code}: {r.text}")
        return r.json().get("value", [])

    def list_tables(self) -> list[str]:
        return [t.get("name") for t in self.list_tables_full()]

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
# Mapeo nombre-en-env -> id-real-de-Graph (resuelto al arrancar)
TABLE_NAME_TO_ID: dict[str, str] = {}
# user_id -> comando pendiente (esperando valor en el siguiente mensaje)
pending_cmd_by_user: dict[int, str] = {}


def resolve_table_ref(name_from_env: str) -> str:
    """Devuelve el id real (o el nombre si no se resolvió)."""
    return TABLE_NAME_TO_ID.get(name_from_env, name_from_env)


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
        await update.message.reply_text(f"No autorizado. Tu ID es {uid}.")


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
    table = resolve_table_ref(COMMAND_TO_TABLE[cmd])
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
    table = resolve_table_ref(COMMAND_TO_TABLE[cmd])
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
        "  /paquetes 169            suma 169 al día de hoy\n"
        "  /gasolina 98,03          suma 98.03 (acepta coma decimal)\n"
        "  /paquetes 11/05 50       suma al 11 de mayo\n"
        "  /poner paquetes 200      REEMPLAZA el valor de hoy\n"
        "  /ver paquetes            muestra el valor de hoy\n"
        "  /ver gasolina 11/05      muestra el valor del 11/05\n\n"
        "Más: /tablas /id /login")


async def cmd_id(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    uid = update.effective_user.id if update.effective_user else "?"
    await update.message.reply_text(f"Tu user_id: {uid}")


async def cmd_tablas(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    await update.message.reply_text("Tablas:\n" + tables_summary())


async def cmd_diag(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Lista archivos/carpetas de OneDrive para ubicar el archivo real."""
    if not is_authorized(update): await deny(update); return
    args = list(context.args or [])
    folder = " ".join(args) if args else ""
    try:
        items = graph.list_children(folder)
    except Exception as e:
        await update.message.reply_text(f"Error listando '{folder}': {e}"); return
    if not items:
        await update.message.reply_text(f"'{folder or '(raiz)'}' esta vacio."); return
    lines = [f"Contenido de '{folder or '(raiz)'}':"]
    for it in items[:50]:
        kind = "[DIR]" if "folder" in it else "[xlsx]" if it.get("name", "").endswith(".xlsx") else "     "
        lines.append(f"  {kind} {it.get('name')}")
    lines.append("\nUso: /diag <carpeta>   ej: /diag Documents")
    await update.message.reply_text("\n".join(lines))


async def cmd_tablas_xls(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    """Lista las tablas reales del archivo Excel configurado."""
    if not is_authorized(update): await deny(update); return
    try:
        names = graph.list_tables()
    except Exception as e:
        await update.message.reply_text(f"Error: {e}"); return
    if not names:
        await update.message.reply_text("El archivo no tiene Tablas (Ctrl+T).")
        return
    await update.message.reply_text("Tablas en el Excel:\n" + "\n".join(f"  - {n}" for n in names))


async def cmd_test(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Prueba varios formatos de URL contra Tabla1 para diagnostico."""
    if not is_authorized(update): await deny(update); return
    args = list(context.args or [])
    name = args[0] if args else COMMAND_TO_TABLE.get(COMMAND_ORDER[0], "Tabla1")
    tid = TABLE_NAME_TO_ID.get(name, name)
    path = quote(EXCEL_FILE_PATH, safe="/")
    base = f"{GRAPH_BASE}/me/drive/root:/{path}:/workbook"
    candidates = [
        ("rest-name",   f"{base}/tables/{quote(name, safe='')}/headerRowRange"),
        ("odata-name",  f"{base}/tables('{quote(name, safe=chr(39))}')/headerRowRange"),
        ("rest-id",     f"{base}/tables/{quote(tid, safe='')}/headerRowRange"),
        ("odata-id",    f"{base}/tables('{quote(tid, safe=chr(39))}')/headerRowRange"),
        ("by-index-0",  f"{base}/tables/itemAt(index=0)/headerRowRange"),
    ]
    lines = [f"Probando contra '{name}' (id={tid}):"]
    hdrs = graph._hdrs()
    for label, url in candidates:
        try:
            r = requests.get(url, headers=hdrs, timeout=20)
            ok = r.status_code < 400
            lines.append(f"  [{label}] {'OK' if ok else 'FAIL'} {r.status_code}")
            if not ok:
                snippet = r.text[:120].replace("\n", " ")
                lines.append(f"     {snippet}")
        except Exception as e:
            lines.append(f"  [{label}] EXC {e}")
    await update.message.reply_text("\n".join(lines))


async def cmd_cabeceras(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Muestra los primeros 10 encabezados (raw) de una tabla."""
    if not is_authorized(update): await deny(update); return
    args = list(context.args or [])
    if not args or args[0].lower() not in COMMAND_TO_TABLE:
        await update.message.reply_text(
            "Uso: /cabeceras <comando>\n" + tables_summary()); return
    cmd = args[0].lower()
    table = resolve_table_ref(COMMAND_TO_TABLE[cmd])
    try:
        headers = graph.get_header_values(table, refresh=True)
    except Exception as e:
        await update.message.reply_text(f"Error: {e}"); return
    sample = headers[:10]
    lines = [f"Encabezados de {table} (primeros 10):"]
    for i, h in enumerate(sample):
        lines.append(f"  [{i}] tipo={type(h).__name__}  valor={h!r}")
    lines.append(f"Total columnas: {len(headers)}")
    await update.message.reply_text("\n".join(lines))


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
        uid = update.effective_user.id
        args = list(context.args or [])
        if not args:
            pending_cmd_by_user[uid] = cmd
            table = COMMAND_TO_TABLE[cmd]
            await update.message.reply_text(
                f"OK /{cmd} ({table}). Envia el numero ahora.\n"
                f"Ej:  100   o   98,03   (suma al dia de hoy)\n"
                f"     11/05 50          (suma a otro dia)\n"
                f"Cancela con /cancelar.")
            return
        target, value = parse_args(args)
        if value is None:
            await update.message.reply_text(
                f"No detecte un numero. Probá:\n/{cmd} 100"); return
        await _do_write(update, cmd, target, value)
    return _h


async def _do_write(update: Update, cmd: str, target: date, value: float) -> None:
    try:
        old, new = write_value(cmd, target, value, mode="sum")
    except Exception as e:
        log.exception("write")
        await update.message.reply_text(f"Error: {e}"); return
    old_txt = f"{old}" if old is not None else "(vacio)"
    table = COMMAND_TO_TABLE[cmd]
    await update.message.reply_text(
        f"OK /{cmd} ({table}) {target.strftime('%d/%m/%Y')}\n"
        f"  anterior: {old_txt}\n"
        f"  +{value}\n"
        f"  nuevo: {new}")


async def cmd_cancelar(update: Update, _: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_authorized(update): await deny(update); return
    uid = update.effective_user.id
    if uid in pending_cmd_by_user:
        cmd = pending_cmd_by_user.pop(uid)
        await update.message.reply_text(f"Cancelado /{cmd}.")
    else:
        await update.message.reply_text("Nada pendiente.")


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
    uid = update.effective_user.id
    text = update.message.text.strip()
    cmd = pending_cmd_by_user.get(uid)
    if cmd:
        # Hay un comando esperando valor
        target, value = parse_args(text.split())
        if value is None:
            await update.message.reply_text(
                f"No detecte un numero para /{cmd}. Reenvialo o /cancelar."); return
        pending_cmd_by_user.pop(uid, None)
        await _do_write(update, cmd, target, value)
        return
    await update.message.reply_text(
        "Tocá uno de los comandos y despues envia el numero:\n" + tables_summary())


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

    # Resolver nombres -> ids reales de Graph
    try:
        tables_meta = graph.list_tables_full()
        log.info("Tablas detectadas en el archivo (%d):", len(tables_meta))
        for t in tables_meta:
            name = t.get("name")
            tid = t.get("id")
            log.info("  name=%r  id=%r", name, tid)
        # Match por nombre exacto, luego por strip/case-insensitive
        env_names = set(COMMAND_TO_TABLE.values())
        for t in tables_meta:
            n = t.get("name") or ""
            tid = t.get("id") or ""
            for env_name in env_names:
                if env_name == n:
                    TABLE_NAME_TO_ID[env_name] = tid
                    break
                if env_name.strip().lower() == n.strip().lower():
                    TABLE_NAME_TO_ID[env_name] = tid
                    break
        for env_name in env_names:
            if env_name not in TABLE_NAME_TO_ID:
                log.warning("Tabla '%s' del .env NO esta en el Excel", env_name)
        log.info("Mapeo resuelto: %s", TABLE_NAME_TO_ID)
    except Exception as e:
        log.warning("No pude listar tablas al inicio: %s", e)

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("id", cmd_id))
    app.add_handler(CommandHandler("tablas", cmd_tablas))
    app.add_handler(CommandHandler("diag", cmd_diag))
    app.add_handler(CommandHandler("tablasxls", cmd_tablas_xls))
    app.add_handler(CommandHandler("cabeceras", cmd_cabeceras))
    app.add_handler(CommandHandler("test", cmd_test))
    app.add_handler(CommandHandler("cancelar", cmd_cancelar))
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
