# 🖋️ Spaziopoesia: Gmail AI Autoresponder

<p align="center">
  <a href="README.md"><b>🇮🇹 Italiano</b></a> | 
  <a href="README.en.md"><b>🇬🇧 English</b></a>
</p>

---

**Spaziopoesia** è un sistema avanzato di autoresponder per Gmail basato su **Google Apps Script** e l'intelligenza artificiale **Gemini**. È progettato per gestire automaticamente le comunicazioni relative a concorsi di poesia, richieste di informazioni e invio di componimenti, garantendo risposte rapide e contestualizzate.

## 🚀 Come Funziona

Il sistema opera come una pipeline automatizzata che segue questi passaggi:

1.  **Trigger Temporale**: Ogni 5 minuti (configurabile), il sistema si attiva automaticamente.
2.  **Caricamento Risorse**: Legge le istruzioni (Knowledge Base) e le configurazioni da un Google Sheet dedicato. Utilizza un sistema di cache sofisticato per ottimizzare le performance.
3.  **Analisi Email**: Identifica le email non lette, ignorando spam e contatti in blacklist.
4.  **Recupero Memoria**: Grazie al `MemoryService`, il bot "ricorda" le interazioni precedenti nello stesso thread per mantenere una conversazione coerente.
5.  **Prompt Engineering**: Il `PromptEngine` costruisce una richiesta dettagliata per Gemini, includendo il contesto del thread, la Knowledge Base e gli eventuali allegati (OCR incluso).
6.  **Generazione & Validazione**: Gemini genera la risposta, che viene poi validata dal `ResponseValidator` per assicurarne la qualità e la coerenza.
7.  **Invio & Etichettatura**: La risposta viene inviata e l'email viene archiviata con un'etichetta (es. `IA`, `Verifica` o `Errore`).

## ✨ Caratteristiche Principali

-   🤖 **AI Multimodale**: Supporto per l'analisi di allegati (immagini, PDF, documenti Office) tramite Gemini Vision.
-   📚 **Knowledge Base Dinamica**: Istruzioni facilmente aggiornabili tramite un foglio Google ("Istruzioni").
-   🧠 **Memoria Contestuale**: Gestione dei thread per evitare risposte ripetitive e mantenere il filo del discorso.
-   🛡️ **Validazione Semantica**: Sistema di scoring per filtrare risposte potenzialmente errate o fuori contesto.
-   🚦 **Rate Limiting**: Gestione intelligente delle quote API per evitare blocchi.
-   📊 **Metriche**: Esportazione giornaliera delle statistiche di utilizzo su Google Sheets.

## 🛠️ Struttura del Progetto

Il codice è suddiviso in moduli specializzati:

-   `gas_main.js`: Orchestratore principale e gestione dei trigger.
-   `gas_email_processor.js`: Logica di elaborazione del batch di email.
-   `gas_gemini_service.js`: Interfaccia con le API di Gemini.
-   `gas_prompt_engine.js`: Costruzione dei prompt per l'IA.
-   `gas_memory_service.js`: Gestione dello storico conversazioni su Sheet.
-   `gas_config.js`: Configurazione centralizzata (API key, modelli, limiti).
-   `gas_setup_ui.js`: Interfaccia per la configurazione iniziale dei trigger.

## ⚙️ Configurazione Rapida

1.  **Script Properties**: Imposta `GEMINI_API_KEY`, `SPREADSHEET_ID` e `METRICS_SHEET_ID`.
2.  **Google Sheets**: Crea un foglio con i tab `Istruzioni`, `Sostituzioni`, `Controllo`, `ConversationMemory` e `DailyMetrics`.
3.  **Trigger**: Esegui la funzione `setupAllTriggers()` dal file `gas_main.js` per attivare il sistema.

---

<p align="center">
  Creato con ❤️ per la gestione automatizzata della poesia.
</p>