# Documentazione Tecnica & Architettura (GEMINI)

Questo file definisce il contesto tecnico dell'applicativo **Gestore Ordini WhatsApp V2**, utile per qualsiasi futuro intervento o refactoring sull'applicazione Desktop.

## Contesto Generale
L'applicazione (`gestore_v2.py`) è un eseguibile sviluppato in **Python 3** utilizzando **PyQt6** per la GUI.  
Nasce come mezzo di automazione per fattorini (rider aziendali) per chiudere in blocco le procedure di:
1. Segnalazione della conclusione del trasporto al ristorante/centrale tramite un unico gruppo WhatsApp.
2. Contabilità dei pagamenti incassati, delle mance gestite sul momento (es. variazioni del totale dei PDF) rispetto alle indicazioni d'origine, salvando tutto in file Excel calcolati dinamicamente.

## Come È Strutturata L'Applicazione?

L'intera logica è attualmente compattata in un unico file monolitico (`gestore_v2.py`) ed orientato agli oggetti:

### 1. Interfaccia e Moduli Generali (`BotApp` & `ListaDropPDF`)
- **PyQt6 (GUI Framework)**: La finestra (QMainWindow) dispone di un widget personalizzato `ListaDropPDF` che permette il Drag&Drop dei file `.pdf` da elaborare. È predisposto pure il Paste (`Cmd` + `V`) diretto dagli appunti.
- L'interfaccia chiede di specificare (e salvare via file `config_gruppo.txt`) l'ID del gruppo WhatsApp dove inviare l'esito.
- Altri controlli in UI permettono di separare l'azione ("Manda su WhatsApp", "Genera/Modifica Excel", o entrambe) avvalendosi di processi sequenziali attivati tramite pulsante "Avvia Elaborazione".

### 2. Regex e Parsing (`estrai_dati_ordini`)
I parser agiscono iterando sui percorsi dei PDF raccolti:
- **PyPDF2**: Utilizzato in sola lettura, unisce tutte le pagine in un Testo unico.
- Vengono usate le Espressioni Regolari specificatamente studiate per il modello della bolla in uso:
  - `REGEX_CODICE`: Cerca numeri subito in seguito alla formattazione fissa es: *"Num.ns.rif #: "*
  - `REGEX_DATA`: Cerca date per instradarvi l'accodamento in Excel ed identificare in quale sub-folder destinarlo.
  - `REGEX_TOTALE` e `REGEX_PAGAMENTO`: Calcolano il float di partenza (sostituendo le `,` con i paramentri decimali standard per Python) e rintracciano se l'originale segnava Pos, Carta, PayPal, Contanti, ecc. Se la regex di un pagamento fallisce o non coincide il fallback assegnato è "Sconosciuto".

### 3. Dinamismi & Gestione del Workflow utente (`fai_domande_interattive`)
Questa funzione incrocia le date estratte nell'iterazione per interrogare il fattorino su variabili contingenti tramite pop-up (`QMessageBox` & `QInputDialog`):
- "Hai usato lo SCOOTER per le consegne del *DATA*?" (modifica la paga aziendale del fattorino).
- Richiede ricorsivamente fino allo sbocco l'intervento manuale su ordini a scelta per aggiustare il *Totale* finale battuto al POS/incassato o confermare l'eventuale variazione di metodo (es. "Doveva essere in contanti ma ha pagato al POS!"). Esempi operativi reali.

### 4. Excel & Storage (`aggiorna_excel`, `archivia_ordini`)
- Usa **pandas** per l'I/O documentale Excel e logica da DataFrame.
- Gestisce l'aggiunta su fogli vecchi leggendo l'Excel esistente, concatenando in modalità `ignore_index=True`, o generandolo ex-novo se assente.
- Ordina proceduralmente la Pivot per `Data` accertandosi di sommare i totali in euro divisi per colonne dei metodi di pagamento.
- **La Mia Paga €**: Incorpora una formula fissa di calcolo (`25.0` euro base + quantitativo del numero di consegne giornaliero se l'utente HA dichiarato di usare il proprio mezzo e NON lo scooter aziendale).
- I file Processati `.pdf` subiscono `shutil.copy` per spostarli su `~/Desktop/ArchivioOrdini/Ordini_{DATA}` e quindi il pre-esistente viene eliminato con `os.remove`. *Rispetta questa semantica qualora si aggiorni su OS differenti (poiché ora pre-formatta in percorsi Mac/Unix).*

### 5. Broadcast WhatsApp (`invia_messaggi`)
- **pywhatkit**: Effettua chiamate al browser per generare link URL che atterrano sul web, portando con se l'ID Gruppo e il Codice Primo.
- **pyautogui**: Risulta essenziale che simuli, dopo l'apertura del tab, di digitare a raffica gli altri codici degli ordini (i successivi dal secondo all'ultimo estratti), schiacciando per ogni order il tasto "Invio", chiudendo alla fine con `Command+w` per non far accatastare Tab aperti di Chrome. Essenziale mantenere i delay stabiliti `time.sleep` vista la reazione asincrona del Browser all'hardware.

## Note Specifiche dell'Ambiente
L'applicativo gestisce anche un modulo "Chiudi Periodo" che, onde evitare conflitti tra OS ed una sessione aperta dell'User su Excel, previene i blocchi OS tramite la funzione di Scudo locale _verifica_excel_aperto(). Implementa un trucco basato sul check dell'esistenza del file fantasma `~$[Nome].xlsx` ed l'eventuale throw da parte di `os.rename`.

1. Costruito e impacchettato presumibilmente con **PyInstaller** usando `gestore_v2.spec`.
2. Contiene `os.chdir(os.path.expanduser("~/Desktop"))` all'avvio: il programma vincola e hard-codifica l'attività ad operare rispetto la Home e la scrivania del Mac. Nessuno degli applicativi funzionerebbe o verrebbe salvato se eseguito su Windows senza le opportune conversioni dei path di sistema.