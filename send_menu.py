import json
import os
import re
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl
import requests


# === CONFIG ===
MENU_PATH = "menu.xlsx"
DAYSTART_PATH = "day_to_start.txt"
STATE_PATH = Path(".send-state/sent.json")

TELEGRAM_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]
APP_TZ = ZoneInfo(os.getenv("APP_TZ", "Europe/Rome"))

# Opzionali, utili per test manuali/locali senza cambiare codice
TEST_TODAY = os.getenv("TEST_TODAY", "").strip()  # es. "2026-05-04"
DEBUG = os.getenv("DEBUG", "0") == "1"
DRY_RUN = os.getenv("DRY_RUN", "0") == "1"

# Valori impostati dal workflow GitHub Actions
EVENT_NAME = os.getenv("EVENT_NAME", "workflow_dispatch").strip().lower()
REQUEST_SOURCE = os.getenv("REQUEST_SOURCE", "admin").strip().lower()
FORCE_SEND = os.getenv("FORCE_SEND", "0") == "1"
ALREADY_SENT = os.getenv("ALREADY_SENT", "false").strip().lower() == "true"
GITHUB_RUN_ID = os.getenv("GITHUB_RUN_ID", "").strip()

# Obiettivo di pubblicazione in ora italiana.
TARGET_HOUR = 10
TARGET_MINUTE = 0

# I run schedulati sono accettati solo in questa finestra italiana.
# I cron UTC vengono lanciati prima delle 10 per dare tempo a GitHub
# di assegnare il runner; se partono in tempo, lo script attende le 10:00.
SCHEDULE_EARLIEST_MINUTES_BEFORE = 5
SCHEDULE_CUTOFF_HOUR = 12
SCHEDULE_CUTOFF_MINUTE = 45


# === DATE DA ESCLUDERE (MODIFICA QUI) ===
# - Giorni singoli: "YYYY-MM-DD"
# - Range inclusivi: ("YYYY-MM-DD", "YYYY-MM-DD")
# - Range aperti: (None, "YYYY-MM-DD") oppure ("YYYY-MM-DD", None)
EXCLUDED_DATES = [
    # "2025-12-24",
    "2026-05-01",
    "2026-12-07",
    ("2026-06-01", "2026-06-02"),
    ("2026-08-10", "2026-08-21"),
    ("2026-12-24", "2027-01-03"),
    # (None, "2026-01-06"),
    # ("2025-12-27", None),
]


def _parse_iso_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def get_today() -> date:
    if TEST_TODAY:
        return _parse_iso_date(TEST_TODAY)
    return datetime.now(APP_TZ).date()


def is_excluded(d: date) -> bool:
    for item in EXCLUDED_DATES:
        # giorno singolo
        if isinstance(item, str):
            if _parse_iso_date(item) == d:
                return True
        # range
        else:
            start_s, end_s = item
            start = _parse_iso_date(start_s) if start_s else date.min
            end = _parse_iso_date(end_s) if end_s else date.max
            if start <= d <= end:
                return True
    return False


def stop_if_already_sent() -> None:
    if ALREADY_SENT and not FORCE_SEND:
        print("Il menu risulta già inviato oggi. Nessun duplicato.")
        raise SystemExit(0)


def enforce_send_window() -> None:
    """
    Applica la finestra utile a tutti gli invii normali:
    - 09:55-09:59 Europe/Rome: il runner aspetta le 10:00;
    - 10:00-12:45: invia appena il workflow parte;
    - prima delle 09:55 o dalle 12:46: nessun invio.

    L'unica eccezione è un workflow_dispatch amministrativo con force=true,
    utile per emergenze/test intenzionali. Il trigger pubblico non può impostarlo.
    """
    if TEST_TODAY:
        return

    if EVENT_NAME == "workflow_dispatch" and REQUEST_SOURCE == "admin" and FORCE_SEND:
        print("Invio amministrativo forzato: controllo orario bypassato.")
        return

    now = datetime.now(APP_TZ)
    target = now.replace(
        hour=TARGET_HOUR,
        minute=TARGET_MINUTE,
        second=0,
        microsecond=0,
    )
    earliest = target - timedelta(minutes=SCHEDULE_EARLIEST_MINUTES_BEFORE)
    # 12:45 è incluso: il cutoff effettivo è 12:46:00 escluso.
    cutoff_exclusive = now.replace(
        hour=SCHEDULE_CUTOFF_HOUR,
        minute=SCHEDULE_CUTOFF_MINUTE + 1,
        second=0,
        microsecond=0,
    )

    print(
        f"Richiesta ricevuta alle {now.isoformat()} "
        f"(event={EVENT_NAME}, source={REQUEST_SOURCE})."
    )
    print(
        "Finestra valida in Europe/Rome: "
        f"{earliest.isoformat()} -> 12:45:59."
    )

    if now < earliest:
        print(
            "Richiesta arrivata prima della finestra utile (09:55 Europe/Rome). "
            "Nessun invio."
        )
        raise SystemExit(0)

    if now >= cutoff_exclusive:
        print(
            "Richiesta arrivata dopo le 12:45 italiane. "
            "A quest'ora il menu non è più utile: nessun invio."
        )
        raise SystemExit(0)

    if now < target:
        seconds_to_wait = (target - now).total_seconds()
        print(
            f"Runner partito in anticipo: attendo {seconds_to_wait:.0f} secondi "
            "per pubblicare alle 10:00 Europe/Rome."
        )
        time.sleep(seconds_to_wait)

        now_after_wait = datetime.now(APP_TZ)
        print(f"Ora di invio raggiunta: {now_after_wait.isoformat()}.")


def giorni_lavorativi_da_a(data_inizio, data_fine):
    giorni = 0
    giorno = data_inizio
    while giorno < data_fine:
        if giorno.weekday() < 5:  # 0=lun, 4=ven
            giorni += 1
        giorno += timedelta(days=1)
    return giorni


def componi_messaggio_menu(menu_del_giorno, giorno_settimana, data_it):
    msg = (
        f"Buongiorno e buon lavoro.\n\n"
        f"‍ *Menù del giorno* ({giorno_settimana.title()} {data_it})\n\n"
        f"*[Primi]*\n"
        f"{menu_del_giorno[0]}.\n"
        f"{menu_del_giorno[1]}.\n"
        f"{menu_del_giorno[2]}.\n"
        f"*[Pasta o riso in bianco/pomodoro]*\n"
        f"*[Secondi]*\n"
        f"{menu_del_giorno[3]}.\n"
        f"{menu_del_giorno[4]}.\n"
        f"{menu_del_giorno[5]}.\n"
        f"*[Pizza gusti del giorno]*\n"
        f"*[Contorni]*\n"
        f"{menu_del_giorno[6]}.\n"
        f"\nBuon appetito dalla Commissione mensa.\n"
    )
    return msg


def _norm_giorno(s: str) -> str:
    s = str(s).strip().upper()
    return s.replace("À", "A").replace("È", "E").replace("É", "E").replace("Ì", "I").replace("Ò", "O").replace("Ù", "U")


def parse_giorno_settimana(s):
    giorni_sett = ["LUNEDI", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ"]
    s_norm = _norm_giorno(s)
    for g in giorni_sett:
        g_norm = _norm_giorno(g)
        if s_norm.startswith(g_norm):
            n = int(s_norm.replace(g_norm, "").strip())
            return g, n
    raise ValueError(f"Formato giorno_settimana errato: {s}")


def trova_riga_col_settimane(ws):
    """
    Cerca la riga che contiene le intestazioni SETTIMANA 1..4.
    Evita righe tipo titolo con "prima settimana".
    """
    best_row = None
    best_vals = None
    best_score = 0

    for row in ws.iter_rows(min_row=1, max_row=30):
        vals = []
        score = 0
        for cell in row:
            v = cell.value
            s = str(v).strip().upper() if v is not None else ""
            vals.append(s)

            if re.search(r"\bSETTIMANA\s*\d+\b", s):
                score += 1

        if score > best_score:
            best_score = score
            best_row = row
            best_vals = vals

    if best_score >= 2:
        return best_row, best_vals

    raise ValueError(
        "Non trovata riga intestazioni SETTIMANA 1..N (score insufficiente)."
    )


def trova_blocchi_giorni(ws):
    giorni = ["LUNEDI", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ"]
    giorni_norm = [_norm_giorno(g) for g in giorni]
    blocchi = []

    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True)):
        prima_col = str(row[0]) if row[0] else ""
        prima_col_norm = _norm_giorno(prima_col)
        if prima_col_norm in giorni_norm:
            blocchi.append((prima_col_norm, i + 1))

    return blocchi


def _norm(s: str) -> str:
    # normalizza spazi, maiuscole e caratteri strani
    return re.sub(r"\s+", " ", str(s).strip().upper())


def trova_colonna_settimana(intestazioni, settimana_n):
    target = f"SETTIMANA {settimana_n}"
    target_n = _norm(target)

    # 1) match esatto normalizzato
    for idx, v in enumerate(intestazioni):
        if _norm(v) == target_n:
            return idx

    # 2) match "contiene" (es. "SETTIMANA 1 - INVERNO")
    for idx, v in enumerate(intestazioni):
        if target_n in _norm(v):
            return idx

    raise ValueError(
        f"Settimana non trovata: '{target}'. Intestazioni disponibili: "
        + " | ".join([_norm(x) for x in intestazioni if _norm(x)])
    )


def trova_blocco_per_giorno(blocchi, giorno):
    giorno_norm = _norm_giorno(giorno)
    for nome, riga in blocchi:
        if nome == giorno_norm:
            return riga
    raise ValueError(f"Giorno non trovato: {giorno}")


def estrai_menu(ws, riga_giorno, col_settimana):
    menu = []
    num_voci_menu = 7

    for r in range(riga_giorno, riga_giorno + num_voci_menu):
        val = ws.cell(row=r, column=col_settimana + 1).value
        if val:
            menu.append(str(val))

    return menu


def send_telegram_message(token, chat_id, text):
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "Markdown",
    }

    try:
        response = requests.post(url, json=payload, timeout=20)
        response.raise_for_status()
        result = response.json()
    except requests.RequestException as exc:
        raise RuntimeError(f"Errore durante la chiamata a Telegram: {exc}") from exc

    if not result.get("ok"):
        raise RuntimeError(f"Telegram ha rifiutato il messaggio: {result}")

    message_id = result.get("result", {}).get("message_id")
    print(f"Messaggio Telegram inviato correttamente. message_id={message_id}")
    return message_id


def write_sent_marker(d_oggi: date, message_id) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)

    state = {
        "date": d_oggi.isoformat(),
        "sent_at": datetime.now(APP_TZ).isoformat(),
        "event_name": EVENT_NAME,
        "request_source": REQUEST_SOURCE,
        "message_id": message_id,
        "github_run_id": GITHUB_RUN_ID or None,
    }

    STATE_PATH.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Creato marcatore invio: {STATE_PATH}")


# === AVVIO ===
stop_if_already_sent()
enforce_send_window()

# === LEGGI DAY_TO_START ===
with open(DAYSTART_PATH, encoding="utf-8") as f:
    daystart = f.read().strip()

giorno_start, data_start = [x.strip() for x in daystart.split(",")]
d_start = datetime.strptime(data_start, "%Y-%m-%d").date()

d_oggi = get_today()  # OGGI in Europe/Rome

# Salta se la data è esclusa
if is_excluded(d_oggi):
    print(f"Oggi {d_oggi} è escluso (EXCLUDED_DATES). Nessun menu da pubblicare.")
    raise SystemExit(0)

# Salta il weekend (solo pubblicazione giorni lavorativi)
if d_oggi.weekday() >= 5:
    print("Oggi è sabato/domenica, nessun menu da pubblicare.")
    raise SystemExit(0)


giorni_sett = ["LUNEDI", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ"]
g_start, sett_start = parse_giorno_settimana(giorno_start)
idx_giorno_start = giorni_sett.index(g_start)

delta_days = giorni_lavorativi_da_a(d_start, d_oggi)
pos_start = (sett_start - 1) * 5 + idx_giorno_start
pos_oggi = pos_start + delta_days
settimane_totali = 4
settimana_menu = (pos_oggi // 5) % settimane_totali + 1
giorno_menu = giorni_sett[pos_oggi % 5]

if DEBUG:
    print("=== DEBUG MENSA ===")
    print("timezone:", APP_TZ)
    print("now:", datetime.now(APP_TZ).isoformat())
    print("event_name:", EVENT_NAME)
    print("request_source:", REQUEST_SOURCE)
    print("force_send:", FORCE_SEND)
    print("already_sent:", ALREADY_SENT)
    print("day_to_start:", daystart)
    print("d_start:", d_start)
    print("d_oggi:", d_oggi, "weekday:", d_oggi.weekday())
    print("delta_days:", delta_days)
    print("pos_start:", pos_start)
    print("pos_oggi:", pos_oggi)
    print("settimana_menu:", settimana_menu)
    print("giorno_menu:", giorno_menu)
    print("===================")

# --- Estrai menu
wb = openpyxl.load_workbook(MENU_PATH, data_only=True)
ws = wb.worksheets[0]  # Primo foglio del file Excel

_, intestazioni = trova_riga_col_settimane(ws)
blocchi = trova_blocchi_giorni(ws)
riga_giorno = trova_blocco_per_giorno(blocchi, giorno_menu)
col_settimana = trova_colonna_settimana(intestazioni, settimana_menu)
menu = estrai_menu(ws, riga_giorno, col_settimana)

if len(menu) < 7:
    raise RuntimeError(
        f"Menu incompleto: trovate {len(menu)} voci, ne servono almeno 7. "
        "Nessun messaggio Telegram inviato."
    )

# --- Componi messaggio
data_it = d_oggi.strftime("%d/%m/%Y")
msg = componi_messaggio_menu(menu, giorno_menu, data_it)

# --- Manda su Telegram
if DRY_RUN:
    print("DRY_RUN=1: non invio su Telegram.")
    print("\n=== MESSAGGIO ===\n")
    print(msg)
else:
    message_id = send_telegram_message(
        TELEGRAM_TOKEN,
        TELEGRAM_CHAT_ID,
        msg,
    )
    write_sent_marker(d_oggi, message_id)
