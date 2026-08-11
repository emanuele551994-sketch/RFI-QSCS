#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automazione Ricerca Treni - RFI QSCS (VERSIONE SELENIUM)
Legge numeri treni e date da Excel e li ricerca automaticamente nel browser
Funziona sia in ambiente desktop che in Codespaces
"""

import sys
import os
import subprocess
import time
from datetime import datetime

# ==============================================
# INSTALLAZIONE AUTOMATICA DIPENDENZE (PRIMO)
# ==============================================

def installa_dipendenze():
    """Installa automaticamente le dipendenze necessarie"""
    dipendenze = ['selenium', 'openpyxl', 'webdriver-manager']
    
    print("=" * 60)
    print("VERIFICA DIPENDENZE")
    print("=" * 60)
    print()
    
    da_installare = []
    
    for dipendenza in dipendenze:
        try:
            __import__(dipendenza)
            print(f"✅ {dipendenza:20s} - Installato")
        except ImportError:
            print(f"❌ {dipendenza:20s} - NON installato")
            da_installare.append(dipendenza)
    
    print()
    
    if da_installare:
        print(f"⚠️  Installazione richiesta per: {', '.join(da_installare)}")
        print()
        print("Installazione in corso...\n")
        
        for dipendenza in da_installare:
            print(f"📦 Installazione di {dipendenza}...")
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', dipendenza, '-q'])
                print(f"✅ {dipendenza} installato con successo\n")
            except subprocess.CalledProcessError:
                print(f"❌ Errore nell'installazione di {dipendenza}")
                print("Prova a installare manualmente:")
                print(f"   pip install {dipendenza}\n")
                sys.exit(1)
        
        print("=" * 60)
        print("✅ Tutte le dipendenze sono state installate!")
        print("=" * 60 + "\n")
    else:
        print("=" * 60)
        print("✅ Tutte le dipendenze sono già installate!")
        print("=" * 60 + "\n")


# Installa le dipendenze PRIMA di importarle
installa_dipendenze()

# ORA possiamo importare i moduli
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl

# ==============================================
# CONFIGURAZIONE
# ==============================================

# URL della pagina di ricerca (MODIFICA SE NECESSARIO)
URL_RICERCA = "https://piccwebrifiol/wol/EMCAM/ValidareTreni/?CercaTreno=true&CodGiurisAI=0"

# Nome del file Excel
EXCEL_FILE = 'COP 269_Elenco cause_Tabella treni - 2026-08-11(CARGO dal 01.07 al 10.08.26).xlsx'

# Riga di inizio dati (la riga 25 contiene i dati)
START_ROW = 25

# Tempo di attesa tra operazioni (in secondi)
DELAY_NORMALE = 1
DELAY_RICERCA = 2

# ==============================================
# FUNZIONI
# ==============================================

def valida_file_excel():
    """Verifica che il file Excel esista"""
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ ERRORE: File non trovato: {EXCEL_FILE}")
        print(f"📁 Cartella corrente: {os.getcwd()}")
        print(f"📂 File disponibili:")
        for f in os.listdir('.'):
            if f.endswith('.xlsx'):
                print(f"   - {f}")
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


def inizializza_browser():
    """Inizializza il browser Chrome/Chromium con Selenium"""
    print("\n" + "=" * 60)
    print("INIZIALIZZAZIONE BROWSER")
    print("=" * 60 + "\n")
    
    try:
        # Opzioni Chrome
        chrome_options = Options()
        
        # Rileva se siamo in Codespaces
        is_codespaces = os.environ.get('GITHUB_WORKSPACE') is not None
        
        if is_codespaces:
            print("🌐 Ambiente Codespaces rilevato\n")
            print("📦 Verifica di Chromium...")
            
            # Verifica se Chromium è installato
            result = subprocess.run(['which', 'chromium-browser'], capture_output=True)
            
            if result.returncode != 0:
                print("❌ Chromium non trovato!")
                print("\n📥 Installa Chromium con:")
                print("   apt-get update && apt-get install -y chromium-browser")
                print("\nPoi riavvia il programma.\n")
                sys.exit(1)
            
            print("✅ Chromium trovato\n")
            
            # Configura per Chromium in Codespaces
            chrome_options.binary_location = '/usr/bin/chromium-browser'
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
        else:
            print("💻 Ambiente Desktop rilevato\n")
        
        # Opzioni comuni
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Inizializza il driver
        print("📦 Caricamento WebDriver Chrome...")
        
        if is_codespaces:
            # In Codespaces usa Chromium direttamente
            driver = webdriver.Chrome(options=chrome_options)
        else:
            # Su desktop scarica ChromeDriver
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
        
        print("✅ Browser inizializzato\n")
        return driver
    
    except Exception as e:
        print(f"❌ ERRORE nell'inizializzazione del browser: {e}")
        print("\n💡 Suggerimenti:")
        print("   - Su Codespaces: installa Chromium con:")
        print("     apt-get update && apt-get install -y chromium-browser")
        print("   - Su Desktop: scarica Chrome da https://www.google.com/chrome/")
        sys.exit(1)


def apri_pagina(driver):
    """Apre la pagina di ricerca"""
    print(f"🌐 Apertura pagina: {URL_RICERCA}\n")
    
    try:
        driver.get(URL_RICERCA)
        time.sleep(3)  # Aspetta il caricamento
        print("✅ Pagina caricata\n")
    except Exception as e:
        print(f"❌ ERRORE nell'apertura della pagina: {e}")
        driver.quit()
        sys.exit(1)


def esegui_ricerca_selenium(driver, numero_treno, data_treno):
    """
    Esegue una singola ricerca usando Selenium
    1. Trova il campo numero treno
    2. Inserisce il numero
    3. Trova il campo data
    4. Inserisce la data
    5. Clicca il pulsante Cerca
    """
    try:
        wait = WebDriverWait(driver, 10)
        
        # Trova il campo numero (cerca per attributo value che contiene 61655)
        print(f"      🔍 Ricerca campi...")
        
        # Prova a trovare il campo input con name="Numero"
        try:
            campo_numero = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
            )
        except:
            # Se non trova, usa il primo input disponibile
            campo_numero = driver.find_element(By.CSS_SELECTOR, "input[type='text']")
        
        # Pulisci e inserisci il numero
        campo_numero.clear()
        campo_numero.send_keys(str(int(numero_treno)))
        time.sleep(DELAY_NORMALE)
        
        print(f"      📝 Numero inserito: {numero_treno}")
        
        # Trova tutti gli input
        inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
        
        # Il secondo input dovrebbe essere la data
        if len(inputs) >= 2:
            campo_data = inputs[1]
            campo_data.clear()
            campo_data.send_keys(data_treno)
            time.sleep(DELAY_NORMALE)
            print(f"      📅 Data inserita: {data_treno}")
        else:
            print(f"      ⚠️  Campo data non trovato")
        
        # Trova il pulsante Cerca (cerca per testo o classe)
        try:
            pulsante_cerca = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Cerca')]"))
            )
        except:
            try:
                pulsante_cerca = driver.find_element(By.CSS_SELECTOR, "button")
            except:
                print(f"      ⚠️  Pulsante Cerca non trovato")
                return False
        
        # Clicca il pulsante
        driver.execute_script("arguments[0].click();", pulsante_cerca)
        print(f"      🔘 Pulsante Cerca cliccato")
        
        time.sleep(DELAY_RICERCA)
        print(f"      ✅ Ricerca completata\n")
        
        return True
    
    except Exception as e:
        print(f"      ❌ ERRORE: {e}\n")
        return False


def automazione_principale():
    """
    Ciclo principale di automazione
    """
    print("=" * 60)
    print("AUTOMAZIONE RICERCA TRENI (SELENIUM)")
    print("=" * 60)
    print()
    
    # Valida file
    valida_file_excel()
    
    # Carica dati
    dati = carica_dati_excel()
    
    if not dati:
        print("❌ ERRORE: Nessun dato trovato nel file Excel")
        sys.exit(1)
    
    # Inizializza browser
    driver = inizializza_browser()
    
    try:
        # Apri pagina
        apri_pagina(driver)
        
        # Esegui ricerche
        print(f"🔍 Inizio ricerche ({len(dati)} treni)...\n")
        print("-" * 60)
        
        successi = 0
        errori = 0
        
        for idx, treno in enumerate(dati, 1):
            numero = treno['numero']
            data = treno['data']
            
            print(f"[{idx}/{len(dati)}] 🚂 Treno: {numero:6d} | 📅 Data: {data}")
            
            if esegui_ricerca_selenium(driver, numero, data):
                successi += 1
            else:
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
        
        # Tieni il browser aperto per 10 secondi per vedere i risultati
        is_codespaces = os.environ.get('GITHUB_WORKSPACE') is not None
        if not is_codespaces:
            print("⏳ Browser resterà aperto per 10 secondi...")
            time.sleep(10)
    
    finally:
        driver.quit()
        print("🌐 Browser chiuso")


def main():
    """
    Punto di ingresso del programma
    """
    if len(sys.argv) > 1:
        if sys.argv[1] == '--help':
            mostra_aiuto()
        else:
            mostra_aiuto()
    else:
        automazione_principale()


def mostra_aiuto():
    """Mostra il messaggio di aiuto"""
    print("""
╔════════════════════════════════════════════════════════════════╗
║    AUTOMAZIONE RICERCA TRENI - RFI QSCS (VERSIONE SELENIUM)    ║
╚════════════════════════════════════════════════════════════════╝

UTILIZZO:
  python automazione_treni_selenium.py  # Esegui automazione
  python automazione_treni_selenium.py --help  # Mostra questo messaggio

CONFIGURAZIONE:
  - URL_RICERCA: URL della pagina di ricerca RFI
  - EXCEL_FILE: Nome del file Excel con i dati
  - START_ROW: Riga di inizio dei dati (default: 25)

REQUISITI (installati automaticamente):
  - selenium
  - openpyxl
  - webdriver-manager

INSTALLAZIONE SU CODESPACES:
  1. Apri il terminale
  2. Esegui: apt-get update && apt-get install -y chromium-browser
  3. Esegui: python automazione_treni_selenium.py

INSTALLAZIONE SU DESKTOP:
  1. Scarica Chrome da https://www.google.com/chrome/
  2. Metti il file nella cartella con l'Excel
  3. Esegui: python automazione_treni_selenium.py

FILE EXCEL:
  - La colonna A deve contenere i numeri dei treni
  - La colonna B deve contenere le date (formato DD/MM/YYYY)
  - I dati iniziano dalla riga 25

FUNZIONA SU:
  ✓ Windows / Mac / Linux (con Chrome)
  ✓ Codespaces (con Chromium)
  ✓ Desktop con browser
  ✓ Ambienti cloud

NOTE:
  - Non necessita PyAutoGUI
  - Non ha problemi con display X11
  - Rileva automaticamente l'ambiente
""")


if __name__ == "__main__":
    main()
