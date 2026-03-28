# Gestore Ordini WhatsApp

Applicazione desktop per automatizzare la lettura dei PDF degli ordini, l'aggiornamento della contabilità in Excel e l'invio dei codici completati su un gruppo WhatsApp in modo istantaneo. Progettato specificamente per chi gestisce le consegne.

## Funzionalità Principali

1. **Elaborazione PDF Drag&Drop:** Trascina i file PDF degli ordini nell'interfaccia. L'app estrarrà in automatico tramite Regex (codice ordine, data, importo totale e metodo di pagamento).
2. **Controllo Metodi di Pagamento:** L'app chiederà se si sono verificati cambiamenti dell'ultimo minuto rispetto a quanto riportato nel PDF (es: un ordine con pagamento alla consegna tramutato in POS o viceversa) oppure variazioni del totale.
3. **Calcolo Paga Giornaliera:** Integrata una domanda inerente al mezzo utilizzato (se lo scooter aziendale proprio o meno) per determinare automaticamente una quota extra sulla busta o rimborso, con calcoli gestiti in Background.
4. **Automazione WhatsApp:** Tramite l'ID del gruppo salvato nell'app, il tool apre automaticamente WhatsApp Web sul browser e digita i codici di tutti gli ordini appena processati.
5. **Esportazione e Riepilogo Excel:** Crea (o aggiorna, se esistente) un file Excel nominato `Resoconto_Periodo.xlsx` contenente i fogli di dettaglio di ogni consegna e la Pivot di resoconto giornaliera con l'incasso e le quote del fattorino per ogni singola giornata.
6. **Archiviazione:** I PDF vengono spostati automaticamente in una cartella di archivio e categorizzati per data (`ArchivioOrdini/Ordini_YYYY-MM-DD`).
7. **Chiusura del Periodo:** Permette di congelare il file Excel al termine del periodo/settimana, rinominandolo con il timestamp (ed iniziando il file successivo vuoto ai futuri utilizzi).

## Requisiti di Sistema
- **macOS** obbligatorio causa impostazioni di percorso e le automazioni tramite PyAutoGUI (`command` + `w` per chiudere il tab di Chrome).
- Permessi di Sicurezza ed Accessibilità (su Mac) attivi per permettere a PyAutoGUI di replicare le battiture dei codici nella chat in automatico.

## Uso Rapido
Una volta avviato l'eseguibile:
1. Inserisci e salva il codice/ID del Gruppo WhatsApp (in alto).
2. Trascina all'interno del riquadro tratteggiato tutti i PDF che intendi smazzare.
3. Seleziona l'azione desiderata (Inserimento in WA e/o Archivio in Excel).
4. Clicca su "Avvia Elaborazione" e rispondi alle domande per confermare i totali. Lascia che il browser faccia il resto senza toccare nulla nel mentre.