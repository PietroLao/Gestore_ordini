import contextlib
import sys
import os
import re
import time
import shutil
from datetime import datetime
from PyPDF2 import PdfReader
import pywhatkit as kit
import pyautogui
import pandas as pd

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QPushButton, QListWidget, QMessageBox, QInputDialog, QCheckBox, QLineEdit)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QKeySequence

# Forza il desktop come cartella di lavoro
os.chdir(os.path.expanduser("~/Desktop"))

# --- CONFIGURAZIONE PERCORSI ---
CARTELLA_ARCHIVIO = r"/Users/pietro/Desktop/ArchivioOrdini"
FILE_EXCEL = r"/Users/pietro/Desktop/ArchivioOrdini/Resoconto_Periodo.xlsx"
FILE_CONFIG_ID = os.path.join(CARTELLA_ARCHIVIO, "config_gruppo.txt")

# --- REGEX ---
REGEX_CODICE = r"Num\.ns\.rif #:\s*(\d+)"
REGEX_DATA = r"Data:\s*(\d{2}/\d{2}/\d{4})"
REGEX_TOTALE = r"Totale:\s*([\d,]+)"
REGEX_PAGAMENTO = r"(?i)(carta|contanti|pos|paypal)"

class ListaDropPDF(QListWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setStyleSheet("background-color: #f0f0f0; border: 2px dashed #aaa; border-radius: 5px; padding: 5px;")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.acceptProposedAction()
        else: event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls(): event.acceptProposedAction()
        else: event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            percorso = url.toLocalFile()
            if percorso.lower().endswith('.pdf'):
                items = [self.item(i).text() for i in range(self.count())]
                if percorso not in items: self.addItem(percorso)
        event.acceptProposedAction()

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Paste):
            appunti = QApplication.clipboard()
            if appunti.mimeData().hasUrls():
                for url in appunti.mimeData().urls():
                    percorso = url.toLocalFile()
                    if percorso.lower().endswith('.pdf'):
                        items = [self.item(i).text() for i in range(self.count())]
                        if percorso not in items: self.addItem(percorso)
        else:
            super().keyPressEvent(event)

class BotApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestore Ordini WhatsApp V2.4")
        self.setGeometry(100, 100, 500, 650)
        
        os.makedirs(CARTELLA_ARCHIVIO, exist_ok=True)
        
        layout = QVBoxLayout()
        
        # --- SEZIONE ID GRUPPO ---
        layout_id = QHBoxLayout()
        layout_id.addWidget(QLabel("ID Gruppo WA:"))
        self.input_id = QLineEdit()
        self.input_id.setPlaceholderText("Incolla qui l'ID del gruppo")
        self.input_id.setText(self.carica_id())
        layout_id.addWidget(self.input_id)
        
        btn_salva_id = QPushButton("Salva ID")
        btn_salva_id.clicked.connect(self.salva_id)
        layout_id.addWidget(btn_salva_id)
        layout.addLayout(layout_id)
        
        titolo = QLabel("Trascina qui i PDF degli ordini")
        titolo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(titolo)
        
        self.lista_file = ListaDropPDF()
        layout.addWidget(self.lista_file)
        
        # --- OPZIONI DI INVIO/SALVATAGGIO ---
        self.check_whatsapp = QCheckBox("Invia codici sul gruppo WhatsApp")
        self.check_whatsapp.setChecked(True)
        self.check_whatsapp.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.check_whatsapp)

        self.check_excel = QCheckBox("Aggiorna file Excel e archivia PDF")
        self.check_excel.setChecked(True)
        self.check_excel.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.check_excel)
        
        self.btn_avvia = QPushButton("Avvia Elaborazione")
        self.btn_avvia.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px; font-weight: bold; border-radius: 5px;")
        self.btn_avvia.clicked.connect(self.avvia_processo)
        layout.addWidget(self.btn_avvia)
        
        self.btn_chiudi_periodo = QPushButton("Chiudi Periodo (Archivia Excel)")
        self.btn_chiudi_periodo.setStyleSheet("background-color: #E67E22; color: white; padding: 8px; font-weight: bold; border-radius: 5px; margin-top: 10px;")
        self.btn_chiudi_periodo.clicked.connect(self.chiudi_periodo)
        layout.addWidget(self.btn_chiudi_periodo)
        
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    def carica_id(self):
        if os.path.exists(FILE_CONFIG_ID):
            with open(FILE_CONFIG_ID, "r") as f:
                return f.read().strip()
        return "KX4yP8FCKQa8PswmRK7TiN"

    def salva_id(self):
        if nuovo_id := self.input_id.text().strip():
            with open(FILE_CONFIG_ID, "w") as f:
                f.write(nuovo_id)
            QMessageBox.information(self, "Salvato", "ID Gruppo aggiornato correttamente!")

    def _verifica_excel_aperto(self):
        """Restituisce True se l'Excel è aperto (bloccato dal Mac o con file fantasma)."""
        if not os.path.exists(FILE_EXCEL):
            return False
            
        cartella = os.path.dirname(FILE_EXCEL)
        nome_file = os.path.basename(FILE_EXCEL)
        file_fantasma = os.path.join(cartella, f"~${nome_file}")
        
        file_bloccato_da_os = False
        try:
            os.rename(FILE_EXCEL, FILE_EXCEL)
        except OSError:
            file_bloccato_da_os = True

        return os.path.exists(file_fantasma) or file_bloccato_da_os

    def chiudi_periodo(self):
        if not os.path.exists(FILE_EXCEL):
            QMessageBox.information(self, "Info", "Nessun file Excel trovato.")
            return
        
        # SCUDO ANTI-EXCEL APERTO (Versione Mac-Friendly)
        if self._verifica_excel_aperto():
            QMessageBox.critical(self, "Attenzione!", "⚠️ Il file Excel risulta APERTO.\n\nChiudi completamente la finestra di Excel per archiviarlo.")
            return
            
        risposta = QMessageBox.question(self, "Conferma", "Vuoi archiviare il periodo attuale?",
                                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if risposta == QMessageBox.StandardButton.Yes:
            try:
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                os.rename(FILE_EXCEL, FILE_EXCEL.replace(".xlsx", f"_{timestamp}.xlsx"))
                QMessageBox.information(self, "Successo", "Periodo chiuso!")
            except Exception as e:
                QMessageBox.critical(self, "Errore", f"Errore: {e}")

    def avvia_processo(self):
        # SCUDO ANTI-EXCEL APERTO (Versione Mac-Friendly)
        if self.check_excel.isChecked() and self._verifica_excel_aperto():
            QMessageBox.critical(self, "Attenzione!", "⚠️ Il file Excel risulta APERTO.\n\nChiudi completamente la finestra di Excel per permettere il salvataggio dei nuovi dati, poi riprova.")
            return

        percorsi_pdf = [self.lista_file.item(i).text() for i in range(self.lista_file.count())]
        if not percorsi_pdf:
            QMessageBox.warning(self, "Attenzione", "Inserisci dei PDF!")
            return

        ordini, saltati = self.estrai_dati_ordini(percorsi_pdf)
        if saltati and self.check_excel.isChecked():
            QMessageBox.information(self, "Duplicati", f"Ordini già presenti in Excel saltati:\n{', '.join(saltati)}")
        
        if not ordini:
            self.lista_file.clear()
            return

        if self.check_excel.isChecked():
            ordini = self.fai_domande_interattive(ordini)
        
        self.btn_avvia.setEnabled(False)
        self.btn_chiudi_periodo.setEnabled(False)
        self.repaint()

        try:
            if self.check_whatsapp.isChecked():
                self.invia_messaggi(ordini)
            
            if self.check_excel.isChecked():
                self.aggiorna_excel(ordini)
                self.archivia_ordini(ordini)
            
            QMessageBox.information(self, "Completato", "✅ Operazione terminata con successo!")
            self.lista_file.clear()
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Errore durante l'elaborazione:\n{e}")
        finally:
            self.btn_avvia.setEnabled(True)
            self.btn_chiudi_periodo.setEnabled(True)

    def invia_messaggi(self, ordini):
        id_attuale = self.input_id.text().strip()
        kit.sendwhatmsg_to_group_instantly(id_attuale, ordini[0]["Codice_Ordine"], wait_time=25, tab_close=False)
        if len(ordini) > 1:
            time.sleep(3) 
            for i in range(1, len(ordini)):
                pyautogui.typewrite(ordini[i]["Codice_Ordine"])
                time.sleep(1)
                pyautogui.press('enter')    
                time.sleep(2)               
        time.sleep(4) 
        pyautogui.hotkey('command', 'w')

    def _carica_codici_esistenti(self):
        codici = set()
        if os.path.exists(FILE_EXCEL):
            with contextlib.suppress(Exception):
                df = pd.read_excel(FILE_EXCEL, sheet_name="Dati_Ordini", dtype={"Codice_Ordine": str})
                if 'Codice_Ordine' in df.columns:
                    codici = set(df["Codice_Ordine"].str.split('.').str[0].str.strip())
        return codici

    def estrai_dati_ordini(self, percorsi):
        esistenti = self._carica_codici_esistenti() if self.check_excel.isChecked() else set()
        estratti, saltati = [], []
        for p in percorsi:
            try:
                testo = "".join(pag.extract_text() for pag in PdfReader(p).pages)
                if match_c := re.search(REGEX_CODICE, testo):
                    cod = match_c[1].strip()
                    if cod in esistenti: 
                        saltati.append(cod)
                    else:
                        match_d = re.search(REGEX_DATA, testo)
                        data = datetime.strptime(match_d[1], "%d/%m/%Y").strftime("%Y-%m-%d") if match_d else datetime.now().strftime("%Y-%m-%d")
                        
                        match_t = re.search(REGEX_TOTALE, testo)
                        tot = float(match_t[1].replace(",", ".")) if match_t else 0.0
                        
                        match_p = re.search(REGEX_PAGAMENTO, testo)
                        if match_p:
                            parola = match_p.group(1).lower()
                            if "carta" in parola: met = "Carta"
                            elif "contanti" in parola: met = "Contanti"
                            elif "pos" in parola: met = "POS"
                            elif "paypal" in parola: met = "PayPal"
                            else: met = "Sconosciuto"
                        else:
                            met = "Sconosciuto"
                            
                        estratti.append({
                            "Data": data, "Codice_Ordine": cod, "Metodo_PDF": met, "Totale_PDF_€": tot,
                            "Metodo_Finale": met, "Totale_Finale_€": tot, "Uso_Scooter": "Da definire",
                            "nome_file": os.path.basename(p), "percorso_originale": p
                        })
            except Exception as e: print(f"Errore file {os.path.basename(p)}: {e}")
        return estratti, saltati

    def fai_domande_interattive(self, ordini):
        date_presenti = {o["Data"] for o in ordini}
        for d in date_presenti:
            risp = QMessageBox.question(self, "Scooter", f"🛵 Hai usato lo SCOOTER per le consegne del {d}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            uso = "Si" if risp == QMessageBox.StandardButton.Yes else "No"
            for o in ordini:
                if o["Data"] == d: o["Uso_Scooter"] = uso

        lista_c = [f"{o['Codice_Ordine']} ({o['Totale_Finale_€']}€ - {o['Metodo_Finale']})" for o in ordini]

        while True:
            risp = QMessageBox.question(self, "Pagamenti", "💳 Qualcuno ha pagato con un METODO DIFFERENTE?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if risp == QMessageBox.StandardButton.Yes:
                scelta, ok = QInputDialog.getItem(self, "Seleziona Ordine", "Quale ordine ha cambiato metodo?", lista_c, 0, False)
                if ok and scelta:
                    codice_puro = scelta.split(" ")[0]
                    nuovo_metodo, ok2 = QInputDialog.getItem(self, "Nuovo Metodo", "Scegli il nuovo metodo:", ["Carta", "Contanti", "POS", "PayPal"], 0, False)
                    if ok2 and nuovo_metodo:
                        for o in ordini:
                            if o["Codice_Ordine"] == codice_puro:
                                o["Metodo_Finale"] = nuovo_metodo
                        lista_c = [f"{o['Codice_Ordine']} ({o['Totale_Finale_€']}€ - {o['Metodo_Finale']})" for o in ordini]
            else:
                break

        while True:
            risp = QMessageBox.question(self, "Totali", "💰 Ci sono state MODIFICHE AL TOTALE di qualche ordine?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if risp == QMessageBox.StandardButton.Yes:
                scelta, ok = QInputDialog.getItem(self, "Seleziona Ordine", "Quale ordine ha cambiato totale?", lista_c, 0, False)
                if ok and scelta:
                    codice_puro = scelta.split(" ")[0]
                    nuovo_totale, ok2 = QInputDialog.getDouble(self, "Nuovo Totale", "Inserisci il nuovo totale in Euro:", 0.0, 0, 1000, 2)
                    if ok2:
                        for o in ordini:
                            if o["Codice_Ordine"] == codice_puro:
                                o["Totale_Finale_€"] = nuovo_totale
                        lista_c = [f"{o['Codice_Ordine']} ({o['Totale_Finale_€']}€ - {o['Metodo_Finale']})" for o in ordini]
            else:
                break
                
        return ordini

    def archivia_ordini(self, ordini):
        for o in ordini:
            dest = os.path.join(CARTELLA_ARCHIVIO, f"Ordini_{o['Data']}")
            os.makedirs(dest, exist_ok=True)
            shutil.copy(o["percorso_originale"], os.path.join(dest, o["nome_file"]))
            os.remove(o["percorso_originale"])

    def aggiorna_excel(self, ordini):
        dati_puliti = [{k: v for k, v in o.items() if k not in ["nome_file", "percorso_originale"]} for o in ordini]
        df_nuovi = pd.DataFrame(dati_puliti)
        
        if os.path.exists(FILE_EXCEL):
            try: 
                df_old = pd.read_excel(FILE_EXCEL, sheet_name='Dati_Ordini', dtype={'Codice_Ordine': str})
            except ValueError: 
                df_old = pd.read_excel(FILE_EXCEL)
            df_finale = pd.concat([df_old, df_nuovi], ignore_index=True)
        else: 
            df_finale = df_nuovi
            
        df_finale.sort_values(by='Data', inplace=True)
        
        riepilogo = df_finale.pivot_table(index='Data', columns='Metodo_Finale', values='Totale_Finale_€', aggfunc='sum', fill_value=0).reset_index()
        for col in ['Carta', 'Contanti', 'POS', 'PayPal']:
            if col not in riepilogo.columns: riepilogo[col] = 0.0

        colonne_pagamento = [c for c in riepilogo.columns if c != 'Data']
        riepilogo['Incasso_Lordo_€'] = riepilogo[colonne_pagamento].sum(axis=1)
        
        conteggio = df_finale.groupby('Data').size()
        uso_scooter = df_finale.groupby('Data')['Uso_Scooter'].first()
        riepilogo['La_Mia_Paga_€'] = [25.0 + float(conteggio.get(d, 0)) if str(uso_scooter.get(d, 'Si')).lower() in ['no', 'n'] else 25.0 for d in riepilogo['Data']]
        
        col_somma = colonne_pagamento + ['Incasso_Lordo_€', 'La_Mia_Paga_€']
        riga_totale = pd.DataFrame([['TOTALE PERIODO'] + riepilogo[col_somma].sum().tolist()], columns=['Data'] + col_somma)
        riepilogo = pd.concat([riepilogo, riga_totale], ignore_index=True)

        with pd.ExcelWriter(FILE_EXCEL, engine='openpyxl') as writer:
            df_finale.to_excel(writer, sheet_name='Dati_Ordini', index=False)
            riepilogo.to_excel(writer, sheet_name='Riepilogo_Totali', index=False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    finestra = BotApp()
    finestra.show()
    sys.exit(app.exec())