import pytest
import os
import openpyxl

os.environ["TELEGRAM_BOT_TOKEN"] = "fake"
os.environ["TELEGRAM_CHAT_ID"] = "fake"
os.environ["DRY_RUN"] = "1"
os.environ["TEST_TODAY"] = "2026-10-07"

import send_menu

def test_norm_giorno():
    assert send_menu._norm_giorno("LUNEDI") == "LUNEDI"
    assert send_menu._norm_giorno("Lunedì") == "LUNEDI"
    assert send_menu._norm_giorno("MARTEDÌ") == "MARTEDI"
    assert send_menu._norm_giorno("   MERCOLEDÌ  ") == "MERCOLEDI"

def test_parse_giorno_settimana():
    assert send_menu.parse_giorno_settimana("LUNEDI 1") == ("LUNEDI", 1)
    assert send_menu.parse_giorno_settimana("LUNEDÌ 1") == ("LUNEDI", 1)
    assert send_menu.parse_giorno_settimana("MARTEDÌ 2") == ("MARTEDÌ", 2)
    assert send_menu.parse_giorno_settimana("MERCOLEDI 3") == ("MERCOLEDÌ", 3)

def test_trova_blocchi_giorni_and_trova_blocco_per_giorno():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["MENU AUTUNNO"])
    ws.append(["LUNEDI"])
    ws.append(["Pasta"])
    ws.append(["MARTEDI"])
    ws.append(["Riso"])
    ws.append(["MERCOLEDI"])
    ws.append(["Zuppa"])

    blocchi = send_menu.trova_blocchi_giorni(ws)
    assert blocchi == [("LUNEDI", 2), ("MARTEDI", 4), ("MERCOLEDI", 6)]

    assert send_menu.trova_blocco_per_giorno(blocchi, "MERCOLEDÌ") == 6
    assert send_menu.trova_blocco_per_giorno(blocchi, "MERCOLEDI") == 6

    with pytest.raises(ValueError):
        send_menu.trova_blocco_per_giorno(blocchi, "SABATO")

def test_estrai_menu_invalid_length():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["MERCOLEDI", "SETTIMANA 1"])
    ws.append(["", "Pasta"])
    ws.append(["", "Riso"])

    menu = send_menu.estrai_menu(ws, 2, 1)
    assert menu == ["Pasta", "Riso"]

def test_giorno_senza_menu():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["MERCOLEDI", "SETTIMANA 1"])
    menu = send_menu.estrai_menu(ws, 2, 1)
    assert menu == []

def test_data_non_valida():
    with pytest.raises(ValueError):
        send_menu.parse_giorno_settimana("INVALIDO 1")
