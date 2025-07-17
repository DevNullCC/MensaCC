import openpyxl
from datetime import datetime, timedelta
import os
import requests

# === CONFIG ===
MENU_PATH = "menu.xlsx"
DAYSTART_PATH = "day_to_start.txt"
VIRTUAL_TODAY_PATH = "virtual_today.txt"

TELEGRAM_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]

# === FUNZIONI ===

def next_workday(date):
    d = date + timedelta(days=1)
    while d.weekday() >= 5:  # Salta sabato e domenica
        d += timedelta(days=1)
    return d

def giorni_lavorativi_da_a(data_inizio, data_fine):
    giorni = 0
    giorno = data_inizio
    while giorno < data_fine:
        if giorno.weekday() < 5:
            giorni += 1
        giorno += timedelta(days=1)
    return giorni

def componi_messaggio_menu(menu_del_giorno, giorno_settimana, data_it):
    msg = (
        f"Buongiorno e buon lavoro.\n\n"
        f"🧑‍🍳 *Menù del giorno* ({giorno_settimana.title()} {data_it})\n\n"
        f"*[Primi]*\n"
        f"{menu_del_giorno[0]}.\n"
        f"{menu_del_giorno[1]}.\n"
        f"{menu_del_giorno[2]}.\n"
        f"[Pasta o riso in bianco/pomodoro]\n\n"
        f"*[Secondi]*\n"
        f"{menu_del_giorno[3]}.\n"
        f"{menu_del_giorno[4]}.\n"
        f"{menu_del_giorno[5]}.\n\n"
        f"*[Pizza gusti del giorno]*\n"
        f"{menu_del_giorno[6]}.\n\n"
        f"*[Carni o extra]*\n"
        f"{menu_del_giorno[7]}.\n\n"
        f"*[Contorni]*\n"
        f"{menu_del_giorno[8]}.\n\n"
        f"Buon appetito dalla Commissione mensa.\n👋"
    )
    return msg

def parse_giorno_settimana(s):
    giorni_sett = ["LUNEDI", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ"]
    for g in giorni_sett:
        if s.startswith(g):
            n = int(s.replace(g, "").strip())
            return g, n
    raise ValueError("Formato giorno_settimana errato")

def trova_riga_col_settimane(ws):
    for row in ws.iter_rows(min_row=1, max_row=15):
        valori = [str(cell.value).upper() if cell.value else "" for cell in row]
        if any("SETTIMANA" in v for v in valori):
            return row, valori
    raise ValueError("Non trovata riga settimane!")

def trova_blocchi_giorni(ws):
    giorni = ["LUNEDI", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ"]
    blocchi = []
    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True)):
        prima_col = str(row[0]).upper() if row[0] else ""
        if prima_col in giorni:
            blocchi.append( (prima_col, i+1) )
    return blocchi

def trova_colonna_settimana(intestazioni, settimana_n):
    for idx, v in enumerate(intestazioni):
        if v.strip().upper() == f"SETTIMANA {settimana_n}":
            return idx
    raise ValueError("Settimana non trovata")

def trova_blocco_per_giorno(blocchi, giorno):
    for nome, riga in blocchi:
        if nome == giorno:
            return riga
    raise ValueError("Giorno non trovato")

def estrai_menu(ws, riga_giorno, col_settimana):
    menu = []
    for r in range(riga_giorno, riga_giorno + 9):
        val = ws.cell(row=r, column=col_settimana+1).value
        if val: menu.append(str(val))
    return menu

# === 1. LEGGI GIORNO DI PARTENZA E DATA VIRTUALE ===

with open(DAYSTART_PATH, encoding="utf-8") as f:
    daystart = f.read().strip()
giorno_start, data_start = [x.strip() for x in daystart.split(",")]
d_start = datetime.strptime(data_start, "%Y-%m-%d").date()

if os.path.exists(VIRTUAL_TODAY_PATH):
    with open(VIRTUAL_TODAY_PATH) as f:
        data_oggi = f.read().strip()
else:
    data_oggi = data_start

d_oggi = datetime.strptime(data_oggi, "%Y-%m-%d").date()

# === 2. CALCOLA MENU DEL GIORNO VIRTUALE ===

giorni_sett = ["LUNEDI", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ"]
g_start, sett_start = parse_giorno_settimana(giorno_start)
idx_giorno_start = giorni_sett.index(g_start)
delta_days = giorni_lavorativi_da_a(d_start, d_oggi)
pos_start = (sett_start - 1) * 5 + idx_giorno_start
pos_oggi = pos_start + delta_days
settimane_totali = 4
settimana_menu = (pos_oggi // 5) % settimane_totali + 1
giorno_menu = giorni_sett[pos_oggi % 5]

wb = openpyxl.load_workbook(MENU_PATH, data_only=True)
ws = wb.active
row_sett, intestazioni = trova_riga_col_settimane(ws)
blocchi = trova_blocchi_giorni(ws)
riga_giorno = trova_blocco_per_giorno(blocchi, giorno_menu)
col_settimana = trova_colonna_settimana(intestazioni, settimana_menu)
menu = estrai_menu(ws, riga_giorno, col_settimana)

# === 3. COMPONI MESSAGGIO E INVIA ===
data_it = d_oggi.strftime("%d/%m/%Y")
msg = componi_messaggio_menu(menu, giorno_menu, data_it)

def send_telegram_message(token, chat_id, text):
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "Markdown"
    }
    r = requests.post(url, json=payload)
    print(r.text)

send_telegram_message(TELEGRAM_TOKEN, TELEGRAM_CHAT_ID, msg)

# === 4. AVANZA DATA VIRTUALE ===
next_day = next_workday(d_oggi)
with open(VIRTUAL_TODAY_PATH, "w") as f:
    f.write(next_day.strftime("%Y-%m-%d"))
