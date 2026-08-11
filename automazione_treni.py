#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automazione Ricerca Treni - RFI QSCS
Legge numeri treni e date da Excel e li ricerca automaticamente nel browser
"""

import pyautogui
import openpyxl
import time
from datetime import datetime
import os
import sys

# ==============================================
# CONFIGURAZIONE - COORDINATE CALIBRATE
# ==============================================
COORD_NUMERO = (143, 1)       # Campo Numero (dove c'è 61655)
COORD_DATA = (274, 1)         # Campo Data (dove c'è 03/08/2026)
COORD_CERCA = (1861, 1)       # Pulsante Cerca (blu in alto a destra)

# Nome del file Excel
EXCEL_FILE = 'COP 269_Elenco cause_Tabella treni - 2026-08-11(CARGO dal 01.07 al 10.08.26).xlsx'

# Riga di inizio dati (la riga 25 contiene i dati)
START_ROW = 25

# Tempo di attesa tra operazioni (in secondi)
DELAY_NORMAL = 0.3
DELAY_RICERCA = 2

# ==============================================
# FUNZIONI
# ==============================================

def trova_coordinate():
    """
    Utility per trovare le coordinate esatte dei campi
    Stampa le coordinate del mouse in tempo reale
    """
    print("=" * 60)
    print("UTILITY RICERCA COORDINATE")
    print("=" * 60)
    print("\n📍 Muovi il mouse sui campi della pagina")
    print("🎯 Leggi le coordinate e annotale")
    print("⏹️  Premi CTRL+C per fermare\n")
    
    try:
        while True:
            x, y = pyautogui.position()
            print(f"Coordinate: X={x:4d}, Y={y:4d}", end="\r")
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\n\n✅ Utility terminata!")
        sys.exit(0)


def valida_file_excel():
    """Verifica che il file Excel esista"""
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ ERRORE: File non trovato: {EXCEL_FILE}")
        print(f"📁 Cartella corrente: {os.getcwd()}")
        print(f"📂 File disponibili: {os.listdir('.')}")
        sys.exit(1)
    print(f"✅ File Excel trovato: {EXCEL_FILE}")


def carica_dati_excel():
    """Carica i dati dal file Excel"""
    print(f"📖 Caricamento dati da Excel (riga {START_ROW} in poi)...")
    
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        
        dati = []
        for row in ws.iter_rows(min_row=START_ROW, values_only=True):
            numero_treno = row[0]
            data_treno = row[1]
            
            if numero_treno is None:
                break
            
            # Formatta la data
            if isinstance(data_treno, datetime):
                data_formattata = data_treno.strftime("%d/%m/%Y")
            else:
                data_formattata = str(data_treno)
            
            dati.append({
                'numero': int(numero_treno),
                'data': data_formattata
            })
        
        print(f"✅ Caricati {len(dati)} treni")
        return dati
    
    except Exception as e:
        print(f"❌ ERRORE nel caricamento Excel: {e}")
        sys.exit(1)


def prepara_browser():
    """Prepara il browser per l'automazione"""
    print("\n" + "=" * 60)
    print("PREPARAZIONE")
    print("=" * 60)
    print("\n⚠️  IMPORTANTE!")
    print("1️⃣  Assicurati che il browser sia in primo piano")
    print("2️⃣  La pagina deve essere visibile")
    print("3️⃣  L'automazione inizierà tra 5 secondi...\n")
    
    for i in range(5, 0, -1):
        print(f"⏳ Avvio tra {i} secondi...", end="\r")
        time.sleep(1)
    
    print("\n🚀 Automazione avviata!\n")


def pulisci_campo(coordinate):
    """
    Clicca su un campo e lo pulisce
    """
    pyautogui.click(coordinate)
    time.sleep(DELAY_NORMAL)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.1)
    pyautogui.press('backspace')
    time.sleep(0.1)


def inserisci_valore(coordinate, valore, is_numero=False):
    """
    Clicca su un campo, lo pulisce e inserisce il valore
    """
    pulisci_campo(coordinate)
    
    if is_numero:
        valore = str(int(valore))
    else:
        valore = str(valore)
    
    pyautogui.typewrite(valore, interval=0.05)
    time.sleep(DELAY_NORMAL)


def esegui_ricerca(numero_treno, data_treno):
    """
    Esegue una singola ricerca
    1. Inserisce il numero treno
    2. Inserisce la data
    3. Clicca Cerca
    """
    try:
        # Inserisci numero treno
        inserisci_valore(COORD_NUMERO, numero_treno, is_numero=True)
        
        # Inserisci data
        inserisci_valore(COORD_DATA, data_treno, is_numero=False)
        
        # Clicca Cerca
        pyautogui.click(COORD_CERCA)
        time.sleep(DELAY_RICERCA)
        
        return True
    
    except Exception as e:
        print(f"❌ ERRORE durante la ricerca: {e}")
        return False


def automazione_principale():
    """
    Ciclo principale di automazione
    """
    print("=" * 60)
    print("AUTOMAZIONE RICERCA TRENI")
    print("=" * 60)
    
    # Valida file
    valida_file_excel()
    
    # Carica dati
    dati = carica_dati_excel()
    
    if not dati:
        print("❌ ERRORE: Nessun dato trovato nel file Excel")
        sys.exit(1)
    
    # Prepara browser
    prepara_browser()
    
    # Esegui ricerche
    print(f"🔍 Inizio ricerche ({len(dati)} treni)...\n")
    print("-" * 60)
    
    successi = 0
    errori = 0
    
    for idx, treno in enumerate(dati, 1):
        numero = treno['numero']
        data = treno['data']
        
        print(f"[{idx}/{len(dati)}] 🚂 Treno: {numero:6d} | 📅 Data: {data}")
        
        if esegui_ricerca(numero, data):
            print(f"      ✅ Completato\n")
            successi += 1
        else:
            print(f"      ❌ Errore\n")
            errori += 1
    
    # Risultati finali
    print("-" * 60)
    print("\n" + "=" * 60)
    print("RISULTATI FINALI")
    print("=" * 60)
    print(f"✅ Successi: {successi}")
    print(f"❌ Errori:   {errori}")
    print(f"📊 Totale:   {len(dati)}")
    print("=" * 60 + "\n")
    
    if errori == 0:
        print("🎉 Automazione completata con successo!\n")
    else:
        print(f"⚠️  Automazione completata con {errori} errori\n")


def main():
    """
    Punto di ingresso del programma
    """
    if len(sys.argv) > 1:
        if sys.argv[1] == '--coordinate':
            trova_coordinate()
        elif sys.argv[1] == '--help':
            mostra_aiuto()
        else:
            mostra_aiuto()
    else:
        automazione_principale()


def mostra_aiuto():
    """Mostra il messaggio di aiuto"""
    print("""
╔════════════════════════════════════════════════════════════════╗
║         AUTOMAZIONE RICERCA TRENI - RFI QSCS                   ║
╚════════════════════════════════════════════════════════════════╝

UTILIZZO:
  python automazione_treni.py           # Esegui automazione
  python automazione_treni.py --coordinate  # Trova coordinate
  python automazione_treni.py --help    # Mostra questo messaggio

CONFIGURAZIONE ATTUALE:
  - COORD_NUMERO = (143, 1)   ✓ Calibrate
  - COORD_DATA = (274, 1)     ✓ Calibrate
  - COORD_CERCA = (1861, 1)   ✓ Calibrate
  - EXCEL_FILE = 'COP 269_Elenco cause_Tabella treni...'
  - START_ROW = 25

COME RICALIBRARE LE COORDINATE (se necessario):
  1. Esegui: python automazione_treni.py --coordinate
  2. Muovi il mouse sui campi della pagina
  3. Leggi le coordinate X, Y
  4. Modifica COORD_NUMERO, COORD_DATA, COORD_CERCA
  5. Salva il file e riavvia

REQUISITI:
  pip install pyautogui openpyxl

FILE EXCEL:
  - La colonna A deve contenere i numeri dei treni
  - La colonna B deve contenere le date (formato DD/MM/YYYY)
  - I dati iniziano dalla riga 25
""")


if __name__ == "__main__":
    main()
