import openpyxl
from datetime import datetime, timedelta
import os
import requests

# === CONFIGURAZIONE ===
excel_path = "menu.xlsx"
giorno_start, data_start = open("day_to_start.txt").read().strip().split(",")
virtual_today_path = "virtual_today.txt"

# 1. Gestione della data virtuale
def next_workday(date):
    d = date + timedelta(days=1)
    while d.weekday() >= 5:  # 5 = sabato, 6 = domenica
        d += timedelta(days=1)
    return d

if os.path.exists(virtual_today_path):
    with open(virtual_today_path) as f:
        data_oggi = f.read().strip()
else:
    data_oggi = data_start

d_oggi = datetime.strptime(data_oggi, "%Y-%m-%d").date()

# (qui tutta la tua logica di calcolo come prima)
# ... COPIA QUI il tuo codice per calcolare il menu ...
# alla fine: ottieni la variabile menu (lista di 9 voci), giorno_menu, settimana_menu

# Esempio di funzione componi_messaggio_menu, adattala al tuo codice
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
        f"{menu_del_giorno[5]}.\n"
        f"[Pizza gusti del giorno]\n"
        f"{menu_del_giorno[6]}.\n"
        f"{menu_del_giorno[7]}.\n\n"
        f"*[Contorni]*\n"
        f"{menu_del_giorno[8]}.\n\n"
        f"Buon appetito dalla Commissione mensa.\n👋"
    )
    return msg

# Invia messaggio Telegram
def send_telegram_message(token, chat_id, text):
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "Markdown"
    }
    r = requests.post(url, json=payload)
    print(r.text)

# --- Chiamata invio ---
# ... trova giorno_settimana, settimana_menu, menu ...
# (come nei tuoi script precedenti)
data_it = d_oggi.strftime("%d/%m/%Y")
msg = componi_messaggio_menu(menu, giorno_menu, data_it)
send_telegram_message(
    os.environ["TELEGRAM_BOT_TOKEN"],
    os.environ["TELEGRAM_CHAT_ID"],
    msg
)

# 3. Avanza la data virtuale
next_day = next_workday(d_oggi)
with open(virtual_today_path, "w") as f:
    f.write(next_day.strftime("%Y-%m-%d"))
