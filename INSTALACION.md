# Bot de Telegram → Excel en OneDrive

Script en Python que recibe mensajes en Telegram y los agrega como filas a un
Excel guardado en tu OneDrive (Microsoft 365 o cuenta personal Outlook/Hotmail)
usando Microsoft Graph API. Pensado para correr 24/7 en tu PC.

## 1. Preparar el Excel en OneDrive

1. Crea/sube un Excel a tu OneDrive, por ejemplo `Documentos/datos.xlsx`.
2. Abre el archivo, escribe los encabezados en la fila 1 (ej: `Fecha`, `Par`,
   `Tipo`, `Entrada`, `Salida`).
3. Selecciona el rango con encabezados y pulsa `Ctrl + T` para convertirlo en
   **Tabla**. Marca "La tabla tiene encabezados".
4. En "Diseño de tabla" anota el **Nombre de la tabla** (ej: `Tabla1`).
5. Anota también el nombre exacto de la **Hoja** (ej: `Hoja1`).

> Sin Tabla con nombre la API no puede insertar filas de forma segura.

## 2. Registrar app en Azure (gratis, una sola vez)

1. Entra a https://portal.azure.com con tu cuenta Microsoft.
2. Ve a **Microsoft Entra ID → App registrations → New registration**.
3. Nombre: `BotExcelTelegram`.
4. Tipos de cuenta admitidos:
   - "Personal Microsoft accounts only" si usas Hotmail/Outlook/Live.
   - "Accounts in any organizational directory and personal" si tienes ambos.
5. **Redirect URI**: tipo *Public client/native* con valor `http://localhost`.
6. Crea la app. Copia el **Application (client) ID** → es tu `MS_CLIENT_ID`.
7. En **Authentication** activa "Allow public client flows" = **Yes** y guarda.
8. En **API permissions** agrega permisos delegados de Microsoft Graph:
   - `Files.ReadWrite`
   - `User.Read`
   Pulsa **Grant admin consent** si te aparece la opción.

## 3. Crear el bot de Telegram

1. En Telegram busca **@BotFather** → `/newbot` → sigue los pasos.
2. Copia el token que te da → es tu `TELEGRAM_BOT_TOKEN`.
3. Habla con **@userinfobot** para conocer tu `user_id` numérico.

## 4. Instalar y configurar en tu PC

Requiere Python 3.10 o superior.

```bash
git clone https://github.com/Norges12/Estrategia1_SMA_ATR.git
cd Estrategia1_SMA_ATR
git checkout claude/telegram-excel-sync-vgOoa

python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux/Mac
source .venv/bin/activate

pip install -r requirements.txt
copy .env.example .env       # en Windows
# cp .env.example .env       # en Linux/Mac
```

Edita `.env` y completa todos los valores.

## 5. Primer arranque

```bash
python telegram_excel_bot.py
```

La primera vez la consola mostrará un código tipo:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code ABCD-1234 to authenticate.
```

Abre la URL, pega el código, inicia sesión con la cuenta que tiene el OneDrive.
El token queda cacheado en `ms_token_cache.json`, así que en arranques
siguientes ya no te lo pedirá.

## 6. Uso desde Telegram

Manda al bot:

```
2026-05-10, EURUSD, BUY, 1.0850, 1.0900
```

El bot agrega esa fila a la tabla. Comandos:

- `/headers` muestra los encabezados de la tabla.
- `/last` muestra la última fila guardada.
- `/add a, b, c` alternativa explícita.
- `/id` te muestra tu user_id (útil si dice "no autorizado").
- `/login` fuerza re-autenticación con Microsoft.

## 7. Que corra siempre en tu PC

### Windows (recomendado: Programador de tareas)

1. Crea `run_bot.bat` en la carpeta del proyecto:
   ```bat
   @echo off
   cd /d "C:\ruta\a\Estrategia1_SMA_ATR"
   call .venv\Scripts\activate
   python telegram_excel_bot.py >> bot.log 2>&1
   ```
2. Abre **Programador de tareas → Crear tarea**.
3. Disparador: "Al iniciar sesión".
4. Acción: ejecutar `run_bot.bat`.
5. En "Configuración" marca "Si la tarea falla, reiniciar cada 1 minuto".

### Linux (systemd)

Crea `/etc/systemd/system/telegram-excel-bot.service`:

```ini
[Unit]
Description=Telegram Excel Bot
After=network.target

[Service]
WorkingDirectory=/home/USUARIO/Estrategia1_SMA_ATR
ExecStart=/home/USUARIO/Estrategia1_SMA_ATR/.venv/bin/python telegram_excel_bot.py
Restart=always
User=USUARIO

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl enable --now telegram-excel-bot
journalctl -u telegram-excel-bot -f
```

## Problemas comunes

- **"No autorizado"** en Telegram → tu user_id no está en `ALLOWED_USER_IDS`.
  Manda `/id` al bot y agrégalo al .env.
- **Graph 404 itemNotFound** → revisa `EXCEL_FILE_PATH` (sin barra inicial,
  con extensión .xlsx) y que el archivo esté en OneDrive personal.
- **Graph 400 unknownTable** → revisa `EXCEL_TABLE_NAME`. Debe ser el nombre
  exacto de la Tabla, no el de la hoja.
- **Token expirado** sin internet → el script reintenta solo cuando vuelve la
  conexión; si falla a fondo manda `/login` desde Telegram.
