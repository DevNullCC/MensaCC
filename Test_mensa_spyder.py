import openpyxl
from datetime import datetime, timedelta

# === CONFIGURAZIONE UTENTE ===
excel_path = r"C:\Users\48077\Desktop\Appoggio\MENU' ESTATE 2025 CASCINA COSTA.xlsx"      # <--- Modifica con il tuo percorso!
day_to_start = "LUNEDI 1,2025-06-23"        # <--- Copia da day_to_start.txt oppure scrivila qui
data_oggi = "2025-07-23"                    # <--- Data che vuoi simulare (YYYY-MM-DD)

# --- FUNZIONI ---
def giorni_lavorativi_da_a(data_inizio, data_fine):
    giorni = 0
    giorno = data_inizio
    while giorno < data_fine:
        if giorno.weekday() < 5:
            giorni += 1
        giorno += timedelta(days=1)
    return giorni

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
    # Modifica qui: metti il numero giusto di righe del nuovo menu (ad esempio 6)
    NUM_VOCI_MENU = 7
    menu = []
    for r in range(riga_giorno, riga_giorno + NUM_VOCI_MENU):
        val = ws.cell(row=r, column=col_settimana+1).value
        if val: menu.append(str(val))
    return menu

def componi_messaggio_menu(menu_del_giorno, giorno_settimana, data_it):
    # Adatta qui se vuoi cambiare il layout o le etichette!
    msg = (
        f"Buongiorno e buon lavoro.\n\n"
        f"🧑‍🍳 *Menù del giorno* ({giorno_settimana.title()} {data_it})\n\n"
        f"*[Primi]*\n"
        f"{menu_del_giorno[0]}.\n"
        f"{menu_del_giorno[1]}.\n"
        f"{menu_del_giorno[2]}.\n"
        f"[Pasta o riso in bianco/pomodoro]\n"
        f"\n"
        f"*[Secondi]*\n"
        f"{menu_del_giorno[3]}.\n"
        f"{menu_del_giorno[4]}.\n"
        f"{menu_del_giorno[5]}.\n"
        f"[Pizza gusti del giorno]\n"
        f"\n"
        f"*[Contorni]*\n"
        f"{menu_del_giorno[6]}.\n"       
        f"\nBuon appetito dalla Commissione mensa.\n👋"
    )
    return msg

# --- INPUT DAY_TO_START ---
giorno_start, data_start = [x.strip() for x in day_to_start.split(",")]
d_start = datetime.strptime(data_start, "%Y-%m-%d").date()
d_oggi = datetime.strptime(data_oggi, "%Y-%m-%d").date()

# --- CALCOLA GIORNO E SETTIMANA
giorni_sett = ["LUNEDI", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ"]
g_start, sett_start = parse_giorno_settimana(giorno_start)
idx_giorno_start = giorni_sett.index(g_start)
delta_days = giorni_lavorativi_da_a(d_start, d_oggi)
pos_start = (sett_start - 1) * 5 + idx_giorno_start
pos_oggi = pos_start + delta_days
settimane_totali = 4
settimana_menu = (pos_oggi // 5) % settimane_totali + 1
giorno_menu = giorni_sett[pos_oggi % 5]

print(f"Data simulata: {data_oggi}")
print(f"Corrisponde a: {giorno_menu} settimana {settimana_menu}")

# --- Estrai menu dal file Excel ---
wb = openpyxl.load_workbook(excel_path, data_only=True)
#ws = wb.active
ws = wb.worksheets[0]  # Primo foglio del file Excel
row_sett, intestazioni = trova_riga_col_settimane(ws)
blocchi = trova_blocchi_giorni(ws)
riga_giorno = trova_blocco_per_giorno(blocchi, giorno_menu)
col_settimana = trova_colonna_settimana(intestazioni, settimana_menu)
menu = estrai_menu(ws, riga_giorno, col_settimana)

# --- Componi messaggio e stampa ---
data_it = d_oggi.strftime("%d/%m/%Y")
msg = componi_messaggio_menu(menu, giorno_menu, data_it)
print("\nMessaggio che verrebbe inviato:")
print(msg)
