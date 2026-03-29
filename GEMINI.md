# Documentazione Tecnica & Architettura (GEMINI)

Questo file definisce il contesto tecnico dell'applicativo **Gestore Ordini WhatsApp V2.6**, utile per qualsiasi futuro intervento o refactoring sull'applicazione Desktop.

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
  - Modalità a **Regex Intelligenti con "Ancora"**:
    - `REGEX_DATA_CONSEGNA`: Sfrutta il flag `(?is)` agganciandosi al prefisso *"data consegna"*, andando poi a cacciare la prima data utile all'interno di un raggio di 150 caratteri (non-greedy `.{0,150}?`). Include un robusto **Fallback**: qualora fallisca, l'app pesca brutalmente tutte le stringhe di formato data (\`\\d{2}/\\d{2}/\\d{4}\`) che passano indenni un controllo `datetime.strptime` all'interno di un blocco `try-except ValueError`, prelevando sempre la cronologicamente più recente (`[-1]`). Questo impedisce crash nel caso il PDF presenti stringhe false-positive di date non valide.
    - `REGEX_PAGAMENTO`: Stessa architettura, utilizza come ancora i termini *"pagamento"* o *"modalità"*, scandagliando i successivi 150 caratteri. Limitando il raggio visivo della ricerca a partire dalle keyword originali della tabella del PDF, l'app evita di leggere le fastidiose variazioni digitate per sbaglio a mano libere dal locale nei campi "Note: ".
  - `REGEX_TOTALE`: Estrae il calcolo float di partenza recuperando unicamente l'importo totale. Se falliscono, i valori assumono i default "Sconosciuto" e `0.0`.

### 3. Dinamismi & Gestione del Workflow utente (`fai_domande_interattive`)
Questa funzione incrocia le date estratte nell'iterazione per interrogare il fattorino su variabili contingenti tramite pop-up (`QMessageBox` & `QInputDialog`):
- "Hai usato lo SCOOTER per le consegne del *DATA*?" (modifica la paga aziendale del fattorino).
- Richiede ricorsivamente fino allo sbocco l'intervento manuale su ordini a scelta per aggiustare il *Totale* finale battuto al POS/incassato o confermare l'eventuale variazione di metodo (es. "Doveva essere in contanti ma ha pagato al POS!"). Tra le opzioni manuali per il metodo di pagamento è stato aggiunto "**Non pagato**" (per gestire le eccezioni o ordini offerti dal capo non rilevabili dai PDF).
- **Tracking Modifiche**: Ogni volta che l'utente sovrascrive manualmente un totale o un metodo di pagamento via GUI, viene impostato un flag `Modificato: True` per quel record.

### 4. Excel & Storage (`aggiorna_excel`, `archivia_ordini`)
- Usa **pandas** per l'I/O documentale Excel e logica da DataFrame, coadiuvato da **openpyxl** per l'applicazione di stili nativi.
- Gestisce l'aggiunta su fogli vecchi leggendo l'Excel esistente, concatenando in modalità `ignore_index=True`, o generandolo ex-novo se assente.
- Ordina proceduralmente la Pivot per `Data` accertandosi di sommare i totali in euro divisi per colonne dei metodi di pagamento (allocando opportunamente anche la colonna per gli ordini "**Non pagato**").
- **Highlighting Visivo**: Al momento del salvataggio nel foglio excel 'Dati_Ordini', un loop con `openpyxl` intercetta le righe contrassegnate dal flag `Modificato` e ne colora lo sfondo interamente di **Giallo** (`FFFF00`), rendendole nettamente distinguibili a colpo d'occhio nel report contabile periodico.
- **La Mia Paga €**: Incorpora una formula fissa di calcolo (`25.0` euro base + quantitativo del numero di consegne giornaliero se l'utente HA dichiarato di usare il proprio mezzo e NON lo scooter aziendale).
- I file Processati `.pdf` subiscono `shutil.copy` per spostarli su `~/Desktop/ArchivioOrdini/Ordini_{DATA}` e quindi il pre-esistente viene eliminato con `os.remove`. *Rispetta questa semantica qualora si aggiorni su OS differenti (poiché ora pre-formatta in percorsi Mac/Unix).*

### 5. Broadcast WhatsApp (`invia_messaggi`)
- **pywhatkit**: Effettua chiamate al browser per generare link URL che atterrano sul web, portando con se l'ID Gruppo e il Codice Primo.
- **Variazioni Pagamento**: Se l'utente in fase di revisione ("fai_domande_interattive") modifica il metodo di incasso, il nuovo metodo viene accodato in stringa al messaggio inviato sul gruppo WhatsApp (es. `26001072 POS`).
- **pyautogui**: Risulta essenziale che simuli, dopo l'apertura del tab, di digitare a raffica gli altri codici degli ordini (i successivi dal secondo all'ultimo estratti), schiacciando per ogni order il tasto "Invio", chiudendo alla fine con `Command+w` per non far accatastare Tab aperti di Chrome. Essenziale mantenere i delay stabiliti `time.sleep` vista la reazione asincrona del Browser all'hardware.

## Note Specifiche dell'Ambiente
L'applicativo gestisce anche un modulo "Chiudi Periodo" che, onde evitare conflitti tra OS ed una sessione aperta dell'User su Excel, previene i blocchi OS tramite la funzione di Scudo locale _verifica_excel_aperto(). Implementa un trucco basato sul check dell'esistenza del file fantasma `~$[Nome].xlsx` ed l'eventuale throw da parte di `os.rename`.

1. Costruito e impacchettato presumibilmente con **PyInstaller** usando `gestore_v2.spec`.
2. Contiene `os.chdir(os.path.expanduser("~/Desktop"))` all'avvio: il programma vincola e hard-codifica l'attività ad operare rispetto la Home e la scrivania del Mac. Nessuno degli applicativi funzionerebbe o verrebbe salvato se eseguito su Windows senza le opportune conversioni dei path di sistema.
3. **App Ambiente Virtuale:** È richiesto l'utilizzo di `Conda` per gestire le dipendenze. Prima di qualsiasi esecuzione dell'IDE per lo sviluppo ed i test si raccomanda l'attivazione preventiva dell'environment designato tramite il comando: `conda activate bot_ordini`.

## Struttura e Template dei PDF (`esempi_pdf`)
All'interno della cartella `esempi_pdf` sono forniti modelli di riferimento per testare i diversi output e le eccezioni prodotte dalla formattazione dei documenti:
- **spiga (1).pdf**: Mostra un ordine saldato con **POS** , e la presenza di testo nel campo *Note*.
- **Pala  (11).pdf**: Modello per i pagamenti finalizzati con **PayPal**.
- **moretto (1).pdf**: Modello per i pagamenti finalizzati con **Carta**.
- **schievinin.pdf**: Modello per i pagamenti finalizzati in **Contanti**.

⚠️ **CRITICITÀ SULLE DATE & PAGAMENTI: RISOLTO (V2.6)**
I documenti di riferimento contenevano la criticità legata alla presenza di "due date diverse" (ricezione ordine in alta impaginazione e consegna in bassa impaginazione) nonché distorsioni da parole chiave "Contanti" o "POS" digitate casualmente nelle note testuali. 
L'emergenza è stata arginata attraverso le nuove **Regex di Ancoraggio** (`(?is)data\s+consegna.{0,150}?` e `(?:pagamento|modalit[aà]).{0,150}?`) capaci di agganciare esclusivamente le sezioni desiderate in modo intelligente ed estrarne i valori su base controllata (nei 150 char successivi).
Un sistema a *fallback* avanzato cicla su tutte le date trovate verificandole con `try-except`, estraendo la data cronologicamente più recente e scongiurando errori o crash fatali sul `datetime.strptime`.
Inoltre, per migliorare la comunicazione tra l'app e il gruppo WhatsApp, ogni override manuale da parte del rider in merito a mezzi di pagamento (es. "POS" al posto di "Contanti") viene adesso accodato dinamicamente nella stringa prodotta da `pyautogui`.