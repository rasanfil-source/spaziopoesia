/**
 * Config.js - Configurazione centralizzata del sistema
 * Tutti i parametri configurabili sono definiti qui
 */

const CONFIG = {
  // === API ===
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  MODEL_NAME: 'gemini-2.5-flash',

  // === Generazione ===
  TEMPERATURE: 0.5,
  MAX_OUTPUT_TOKENS: 6000,

  // === Validazione ===
  VALIDATION_ENABLED: true,
  VALIDATION_MIN_SCORE: 0.6,
  VALIDATION_WARNING_THRESHOLD: 0.9,  // Soglia warning sotto cui aggiungere etichetta Verifica
  SEMANTIC_VALIDATION: {
    enabled: true,
    activationThreshold: 0.9,
    cacheEnabled: true,
    cacheTTL: 300,
    taskType: 'semantic',
    maxRetries: 1,
    fallbackOnError: true
  },

  // === Gmail ===
  LABEL_NAME: 'IA',                    // Label per email processate
  ERROR_LABEL_NAME: 'Errore',          // Label per errori
  VALIDATION_ERROR_LABEL: 'Verifica',  // Label per risposte da rivedere
  // Ridotto a 3 per supportare strategia "Cross-Key Quality First"
  // Fino a 4 chiamate API per email → batch ridotto per prevenire timeout GAS (6 min)
  MAX_EMAILS_PER_RUN: 3,
  MAX_CONSECUTIVE_EXTERNAL: 5,        // Soglia per rilevamento email loop
  EMPTY_INBOX_WARNING_THRESHOLD: 5,   // Soglia per warning inbox vuota
  SUSPENSION_STALE_UNREAD_HOURS: 12,    // Paracadute: processa unread vecchie anche in fascia sospesa
  MIN_REMAINING_TIME_MS: 90000,      // Stop preventivo se resta meno di 90 secondi
  EXECUTION_LOCK_WAIT_MS: 5000,      // Timeout acquisizione lock esecuzione (ms)
  SEARCH_PAGE_SIZE: 50,              // Numero massimo thread restituiti da GmailApp.search
  MAX_EXECUTION_TIME_MS: 280000,    // Budget massimo per run (default GAS trigger ~6 minuti)
  GMAIL_LABEL_CACHE_TTL: 3600000,      // 1 ora in millisecondi
  MAX_HISTORY_MESSAGES: 10,            // Massimo messaggi in cronologia thread
  ATTACHMENT_CONTEXT: {
    enabled: true,                   // Includi allegati multimodali (Vision) nel prompt
    maxFiles: 4,                     // Numero massimo di allegati da processare
    maxBytesPerFile: 5 * 1024 * 1024,// 5 MB per file
    ocrTriggerKeywords: [            // Attiva Vision solo se il contenuto è rilevante
      'poesia', 'componimento', 'testo', 'immagine', 'allegato', 'foto', 'foto poesia'
    ]
  },
  // Stima token per allegati multimodali (Gemini Vision)
  ATTACHMENT_TOKEN_ESTIMATE: {
    image: 258,          // Token stimati per immagine inviata come inlineData
    pdf: 1032,           // Token stimati per PDF inviato come inlineData
    defaultDoc: 1032     // Token stimati per documento generico (es. Office convertito in PDF)
  },

  // === Cache e Lock ===
  CACHE_LOCK_TTL: 240,                 // Secondi (copre OCR + AI + validazione semantica)
  CACHE_RACE_SLEEP_MS: 200,             // Attesa anti-race condition

  // === Alias noti (anti-loop: il bot riconosce sé stesso anche quando invia da alias) ===
  KNOWN_ALIASES: ['info@parrocchiasanteugenio.it'],

  // === Knowledge Base ===
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),
  KB_SHEET_NAME: 'Istruzioni',
  REPLACEMENTS_SHEET_NAME: 'Sostituzioni',

  MEMORY_SHEET_NAME: 'ConversationMemory',
  MAX_PROVIDED_TOPICS: 50,             // Limite massimo topic in memoria
  MEMORY_LOCK_TTL: 30,                 // Lock TTL in secondi per MemoryService (>= timeout lock Sheet)
  SHEET_WRITE_LOCK_TIMEOUT_MS: 10000,  // Timeout attesa ScriptLock prima di scrivere su Sheet

  // === Retry API Sheets ===
  SHEETS_RETRY_MAX: 3,                 // Tentativi massimi
  SHEETS_RETRY_BACKOFF_MS: 1000,       // Backoff iniziale (raddoppia ad ogni retry)

  // === Modalità ===
  DRY_RUN: false,                      // True per test senza invio email
  FORCE_RELOAD: false,                 // Forza ricaricamento cache KB
  USE_RATE_LIMITER: true,              // Rate limiter intelligente abilitato

  // === Limiti Token (Prompt Engine) ===
  MAX_SAFE_TOKENS: 50000,              // Limite massimo token per prompt (più sicuro per timeout GAS)
  KB_TOKEN_BUDGET_RATIO: 0.5,          // Percentuale budget KB rispetto a max token

  // === Limiti Thread ===
  MAX_THREAD_LENGTH: 10,               // Messaggi massimi per thread prima di anti-loop

  // === Logging ===
  LOGGING: {
    LEVEL: 'INFO',                     // DEBUG, INFO, WARN, ERROR
    STRUCTURED: true,                  // Log in formato JSON
    SEND_ERROR_NOTIFICATIONS: true,    // Invia email per errori critici
    ADMIN_EMAIL: 'rasanfil@gmail.com'  // Email admin per notifiche
  },

  // === Metriche Giornaliere ===
  // Configurare METRICS_SHEET_ID in Script Properties per abilitare export
  METRICS_SHEET_ID: PropertiesService.getScriptProperties().getProperty('METRICS_SHEET_ID'),
  METRICS_SHEET_NAME: 'DailyMetrics',

  // === Modelli Gemini (configurazione centralizzata) ===
  // Aggiornato: Gennaio 2026
  GEMINI_MODELS: {
    // Modello premium: qualità massima per generazione risposte
    'flash-2.5': {
      name: 'gemini-2.5-flash',
      rpm: 15,
      tpm: 1000000,
      rpd: 1500,
      useCases: ['generation']
    },
    // Modello veloce: quick check, classificazione, validazione semantica e fallback
    'flash-lite': {
      name: 'gemini-2.5-flash-lite',
      rpm: 10,
      tpm: 1000000,
      rpd: 1500,
      useCases: ['quick_check', 'classification', 'semantic', 'fallback']
    }
  },

  MODEL_STRATEGY: {
    'quick_check': ['flash-lite', 'flash-2.5'],
    'generation': ['flash-2.5', 'flash-lite'],
    'semantic': ['flash-lite', 'flash-2.5'],
    'fallback': ['flash-lite', 'flash-2.5']
  },

  // === Liste di esclusione ===
  // Nota: lista intentionally mista (domini + email complete).
  // Il matcher supporta sia exact match (email) sia suffisso dominio in _shouldIgnoreEmail.
  IGNORE_DOMAINS: [
    'noreply', 'no-reply', 'newsletter', 'marketing',
    'promo', 'ads', 'notifications',
    'amazon.com', 'eventbrite.com', 'paypal.com', 'ebay.com',
    'subito.it', 'mailchimp.com', 'mailup.com',
    'unclickperlascuolaelosport.it', 'sendinblue.com',
    'miqueldg63@gmail.com', 'rego.juan@gmail.com'
  ],
  IGNORE_KEYWORDS: [
    'unsubscribe', 'opt-out', 'newsletter',
    'disiscriviti', 'disiscrizione', 'annulla iscrizione',
    'annulla l\'iscrizione', 'gestisci la tua iscrizione',
    'gestisci le tue preferenze', 'aggiorna le tue preferenze',
    'cancella iscrizione', 'mailing list', 'inviato con mailup',
    'messaggio inviato con', 'non rispondere a questo messaggio',
    'avviso di sicurezza'
  ]
};

// ====================================================================
// MARCATORI LINGUA (costante condivisa tra moduli)
// ====================================================================

const LANGUAGE_MARKERS = {
  'it': ['grazie', 'cordiali', 'saluti', 'gentile', 'poesia', 'concorso', 'testo', 'versi', 'buongiorno', 'buonasera'],
  'en': ['thank', 'regards', 'dear', 'poetry', 'contest', 'verses', 'poem', 'could', 'please', 'sincerely'],
  'es': ['gracias', 'saludos', 'estimado', 'poesía', 'concurso', 'versos', 'poema', 'buenos', 'días'],
  'pt': ['obrigado', 'obrigada', 'atenciosamente', 'prezado', 'poesia', 'concurso', 'versos', 'bom', 'dia', 'tarde'],
  'fr': ['merci', 'cordialement', 'cher', 'poésie', 'concours', 'vers', 'poème', 'bonjour', 'bonsoir'],
  'de': ['danke', 'grüße', 'liebe', 'poesie', 'wettbewerb', 'verse', 'gedicht', 'guten', 'tag']
};

// ====================================================================
// CACHE GLOBALE
// ====================================================================
// NOTA: GLOBAL_CACHE è dichiarata con init difensiva in gas_main.js.
// NON ridichiarare qui per evitare conflitti di ordine esecuzione file GAS.

// ====================================================================
// VALIDAZIONE CONFIGURAZIONE
// ====================================================================

/**
 * Valida la configurazione all'avvio con schema rigoroso
 * Previene silent failures da typo o tipi errati
 * @returns {Object} Risultato validazione {valid: boolean, errors: string[]}
 */
function validateConfig() {
  const errors = [];

  // Helper per validazione tipo
  const checkType = (path, value, expectedType) => {
    if (typeof value !== expectedType) {
      errors.push(`Errore Config: '${path}' deve essere di tipo ${expectedType}, trovato ${typeof value}`);
    }
  };

  // Helper per validazione range
  const checkRange = (path, value, min, max) => {
    if (typeof value === 'number' && (value < min || value > max)) {
      errors.push(`Errore Config: '${path}' (${value}) fuori range [${min}, ${max}]`);
    }
  };

  // 1. Validazione Campi Critici (fail-fast)
  if (!CONFIG.GEMINI_API_KEY) errors.push('CRITICO: GEMINI_API_KEY mancante');
  if (!CONFIG.SPREADSHEET_ID) errors.push('CRITICO: SPREADSHEET_ID mancante');

  // 2. Validazione Tipi e Valori Logici
  // API & Gen
  checkType('MODEL_NAME', CONFIG.MODEL_NAME, 'string');
  checkType('TEMPERATURE', CONFIG.TEMPERATURE, 'number');
  checkRange('TEMPERATURE', CONFIG.TEMPERATURE, 0.0, 1.0);
  checkType('MAX_OUTPUT_TOKENS', CONFIG.MAX_OUTPUT_TOKENS, 'number');

  // Gmail & Process
  checkType('MAX_EMAILS_PER_RUN', CONFIG.MAX_EMAILS_PER_RUN, 'number');
  checkRange('MAX_EMAILS_PER_RUN', CONFIG.MAX_EMAILS_PER_RUN, 1, 50); // Sanity check
  checkType('LABEL_NAME', CONFIG.LABEL_NAME, 'string');

  // Cache & Lock
  checkType('CACHE_LOCK_TTL', CONFIG.CACHE_LOCK_TTL, 'number');

  // Validation Logic
  checkType('VALIDATION_ENABLED', CONFIG.VALIDATION_ENABLED, 'boolean');
  checkType('VALIDATION_MIN_SCORE', CONFIG.VALIDATION_MIN_SCORE, 'number');
  checkRange('VALIDATION_MIN_SCORE', CONFIG.VALIDATION_MIN_SCORE, 0.0, 1.0);

  // Arrays
  if (!Array.isArray(CONFIG.IGNORE_DOMAINS)) errors.push("Errore Config: 'IGNORE_DOMAINS' deve essere un array");
  if (!Array.isArray(CONFIG.IGNORE_KEYWORDS)) errors.push("Errore Config: 'IGNORE_KEYWORDS' deve essere un array");

  // Omitted advanced parsing errors as this is simplified.
  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * Ottiene la configurazione
 * @returns {Object} Oggetto CONFIG
 */
function getConfig() {
  return CONFIG;
}

/**
 * Healthcheck del sistema
 * @returns {Object} Stato dei componenti
 */
function healthCheck() {
  const health = {
    timestamp: new Date().toISOString(),
    status: 'OK',
    components: {}
  };

  try {
    // Controllo configurazione
    const configValidation = validateConfig();
    health.components.config = {
      status: configValidation.valid ? 'OK' : 'ERROR',
      errors: configValidation.errors
    };

    // Controllo Gmail
    try {
      GmailApp.getInboxThreads(0, 1);
      health.components.gmail = { status: 'OK' };
    } catch (e) {
      health.components.gmail = { status: 'ERROR', error: e.message };
    }

    // Controllo Knowledge Base
    try {
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      health.components.knowledgeBase = { status: 'OK' };
    } catch (e) {
      health.components.knowledgeBase = { status: 'ERROR', error: e.message };
    }

    // Controllo Properties Service
    try {
      PropertiesService.getScriptProperties().getProperty('test');
      health.components.properties = { status: 'OK' };
    } catch (e) {
      health.components.properties = { status: 'ERROR', error: e.message };
    }

    // Determina stato complessivo
    const hasErrors = Object.values(health.components).some(c => c.status === 'ERROR');
    health.status = hasErrors ? 'DEGRADED' : 'OK';

  } catch (e) {
    health.status = 'ERROR';
    health.error = e.message;
  }

  return health;
}