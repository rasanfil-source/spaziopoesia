/**
 * Main.js - Entry point del sistema autoresponder
 * Gestisce trigger, inizializzazione e orchestrazione principale
 * Include logica sospensione oraria e festività italiane
 */

// Inizializzazione difensiva cache globale condivisa tra moduli
var GLOBAL_CACHE = (typeof GLOBAL_CACHE !== 'undefined' && GLOBAL_CACHE) ? GLOBAL_CACHE : {
  loaded: false,
  lastLoadedAt: 0,
  knowledgeBase: '',
  systemEnabled: true,
  vacationPeriods: [],
  ignoreDomains: [],
  ignoreKeywords: [],
  replacements: {}
};

// TTL in-RAM valido nella singola esecuzione: cross-esecuzione la cache reale è CacheService.
// ⚠️ Scelta blindata: questo TTL è allineato ai 21600s della ScriptCache.
// Cambiare solo insieme ai comandi manuali di riallineamento (primeCache/clearKnowledgeCache)
// per evitare disallineamenti "RAM fresca / ScriptCache scaduta" e ricariche imprevedibili.
const RESOURCE_CACHE_TTL_MS = 6 * 60 * 60 * 1000; // 6 ore
const RESOURCE_CACHE_TTL_SECONDS = 21600; // 6 ore
const RESOURCE_CACHE_KEY_V2 = 'SPA_KNOWLEDGE_BASE_V2';
const RESOURCE_CACHE_KEY_V1 = 'SPA_KNOWLEDGE_BASE_V1';
const RESOURCE_CACHE_PARTS_KEY = `${RESOURCE_CACHE_KEY_V2}:parts`;
const RESOURCE_CACHE_PART_PREFIX = `${RESOURCE_CACHE_KEY_V2}:part:`;
const RESOURCE_CACHE_MAX_PART_SIZE = 95000;

// Sospensione e festività rimosse.

// ====================================================================
// CARICAMENTO RISORSE
// ====================================================================

function withSheetsRetry(fn, context = 'Operazione Sheets') {
  const maxRetries = 3;
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return fn();
    } catch (error) {
      if (attempt < maxRetries - 1) {
        console.warn(`⚠️ [${context}] Tentativo ${attempt + 1}/${maxRetries} fallito: ${error.message}. Retry in ${1000 * Math.pow(2, attempt)}ms...`);
        Utilities.sleep(1000 * Math.pow(2, attempt));
        continue;
      }
      console.error(`❌ [${context}] Tutti i ${maxRetries} tentativi esauriti. Ultimo errore: ${error.message}`);
      throw error;
    }
  }
}

function loadResources(acquireLock = true, hasExternalLock = false) {
  // ⚠️ Invariante blindante: niente reload senza lock.
  // Questo evita race condition in cui due trigger sovrascrivono la cache a metà serializzazione.
  if (!acquireLock && !hasExternalLock) {
    throw new Error('loadResources richiede un lock preventivo.');
  }

  if (typeof GLOBAL_CACHE === 'undefined' || !GLOBAL_CACHE) {
    throw new Error('GLOBAL_CACHE non inizializzata: impossibile caricare risorse in sicurezza.');
  }
  // Nota: da qui in avanti GLOBAL_CACHE è garantita; evitiamo fallback silenziosi
  // perché maschererebbero regressioni d'inizializzazione del runtime.

  const now = Date.now();
  const cacheIsFresh = GLOBAL_CACHE.loaded && GLOBAL_CACHE.lastLoadedAt && ((now - GLOBAL_CACHE.lastLoadedAt) < RESOURCE_CACHE_TTL_MS);
  if (cacheIsFresh) return;

  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    if (acquireLock) {
      lockAcquired = lock.tryLock(10000);
      if (!lockAcquired) {
        if (!GLOBAL_CACHE.loaded) {
          throw new Error('Impossibile acquisire lock per caricamento risorse.');
        }
        console.warn('⚠️ Lock non acquisito ma cache già presente: evito reload concorrente non protetto.');
        return;
      }
    }

    _loadResourcesInternal();
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function _loadResourcesInternal() {
  const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;

  // ⚠️ Scelta blindata: la cache persiste SEMPRE il payload completo delle risorse.
  // Eventuali riduzioni/riassunti vanno fatte solo a runtime nel PromptEngine,
  // mai qui, altrimenti si degrada sistematicamente il caso normale.

  // 1. Prova a leggere dalla vera Cache di Apps Script
  if (cache) {
    const cachedData = _readResourceCachePayload(cache);
    if (cachedData) {
      try {
        const parsedData = _deserializeResourceCache(cachedData);
        Object.assign(GLOBAL_CACHE, parsedData);
        GLOBAL_CACHE.loaded = true;
        console.log('✓ Risorse caricate dalla Cache veloce.');
        return;
      } catch (e) {
        console.warn('⚠️ Cache corrotta o obsoleta, ricaricamento dai fogli...');
      }
    }
  }

  const spreadsheetId = (typeof CONFIG !== 'undefined' && CONFIG.SPREADSHEET_ID) ? CONFIG.SPREADSHEET_ID : null;
  if (!spreadsheetId) {
    throw new Error('Impossibile aprire il foglio: CONFIG.SPREADSHEET_ID non configurato.');
  }

  let ss;
  try {
    ss = withSheetsRetry(
      () => SpreadsheetApp.openById(spreadsheetId),
      'Apertura Spreadsheet da CONFIG.SPREADSHEET_ID'
    );
  } catch (e) {
    throw new Error('Impossibile aprire il foglio. Verifica CONFIG.SPREADSHEET_ID. Dettaglio: ' + e.message);
  }

  // Hardening: evita crash se CONFIG non è ancora inizializzato (ordine file GAS)
  // Nota manutenzione: questo caricamento risorse deve avere priorità su processUnreadEmails
  // per assicurare che KB e Dottrina siano disponibili.
  const cfg = (typeof CONFIG !== 'undefined') ? CONFIG : {};
  const newCacheData = {};

  const replacementsSheetName = cfg.REPLACEMENTS_SHEET_NAME || 'Sostituzioni';
  const replacementsSheet = withSheetsRetry(() => ss.getSheetByName(replacementsSheetName), 'Recupero foglio Sostituzioni');
  newCacheData.replacements = {};
  if (replacementsSheet) {
    withSheetsRetry(() => {
      const replacementRows = replacementsSheet.getDataRange().getValues();
      replacementRows.forEach(row => {
        const from = String((row && row[0]) || '').trim();
        const to = String((row && row[1]) || '').trim();
        if (from) {
          newCacheData.replacements[from] = to;
        }
      });
    }, 'Lettura Sostituzioni');
  }

  // 1.5 Caricamento Knowledge Base
  const kbSheetName = cfg.KB_SHEET_NAME || 'Istruzioni';
  const kbSheet = withSheetsRetry(() => ss.getSheetByName(kbSheetName), 'Recupero foglio Knowledge Base');

  newCacheData.knowledgeBase = '';
  if (kbSheet) {
    withSheetsRetry(() => {
      const kbRows = kbSheet.getDataRange().getValues();
      newCacheData.knowledgeBase = _sheetRowsToText(kbRows);
    }, 'Lettura Knowledge Base');
  }

  // Config Avanzata
  const adv = withSheetsRetry(() => _loadAdvancedConfig(ss), 'Lettura Configurazione Avanzata');
  newCacheData.systemEnabled = adv.systemEnabled;
  newCacheData.ignoreDomains = adv.ignoreDomains;
  newCacheData.ignoreKeywords = adv.ignoreKeywords;
  newCacheData.loaded = true;
  newCacheData.lastLoadedAt = Date.now();

  // 3. Salva nella RAM dell'esecuzione corrente
  Object.assign(GLOBAL_CACHE, newCacheData);

  // 4. Salva nel CacheService (6 ore)
  if (cache) {
    try {
      const serialized = _serializeResourceCache(newCacheData, false);
      _writeResourceCachePayload(cache, serialized);
      console.log('✓ Risorse caricate da Fogli e salvate in Cache.');
    } catch (e) {
      console.warn('⚠️ Salvataggio cache standard fallito: ' + e.message);
      try {
        const compressedPayload = _serializeResourceCache(newCacheData, true);
        _writeResourceCachePayload(cache, compressedPayload);
        console.warn('⚠️ Cache risorse salvata in formato compresso (payload vicino limite 100KB).');
      } catch (compressionError) {
        console.warn('⚠️ Impossibile salvare in cache anche in formato compresso: ' + compressionError.message);
      }
    }
  }
}

/**
 * Serializza payload risorse per CacheService.
 * Usa JSON diretto; opzionalmente comprime con gzip+base64 quando disponibile.
 */
function _serializeResourceCache(data, forceCompression) {
  const json = JSON.stringify(data);
  if (!forceCompression) {
    return json;
  }

  if (typeof Utilities === 'undefined' || !Utilities || typeof Utilities.newBlob !== 'function' || typeof Utilities.gzip !== 'function' || typeof Utilities.base64Encode !== 'function') {
    throw new Error('Utilities gzip/base64 non disponibili');
  }

  const gzipped = Utilities.gzip(Utilities.newBlob(json, 'application/json'));
  const base64 = Utilities.base64Encode(gzipped.getBytes());
  return JSON.stringify({
    encoding: 'gzip_base64_json_v1',
    payload: base64
  });
}

/**
 * Deserializza payload risorse da CacheService (plain JSON o gzip+base64).
 */
function _deserializeResourceCache(serializedPayload) {
  const parsed = JSON.parse(serializedPayload);
  if (!parsed || typeof parsed !== 'object' || parsed.encoding !== 'gzip_base64_json_v1') {
    return parsed;
  }

  if (typeof Utilities === 'undefined' || !Utilities || typeof Utilities.ungzip !== 'function' || typeof Utilities.base64Decode !== 'function' || typeof Utilities.newBlob !== 'function') {
    throw new Error('Utilities ungzip/base64 non disponibili per cache compressa');
  }

  const bytes = Utilities.base64Decode(parsed.payload || '');
  const uncompressedBlob = Utilities.ungzip(Utilities.newBlob(bytes));
  const json = uncompressedBlob.getDataAsString('UTF-8');
  return JSON.parse(json);
}

function _splitCachePayload(payload, maxSize) {
  const parts = [];
  for (let i = 0; i < payload.length; i += maxSize) {
    parts.push(payload.substring(i, i + maxSize));
  }
  return parts;
}

function _readResourceCachePayload(cache) {
  if (!cache) return null;

  // Lettura cache con supporto chiave precedente per continuità operativa.
  // Manteniamo questo ramo per non perdere warm-cache durante deploy graduali.
  const v1Payload = cache.get(RESOURCE_CACHE_KEY_V1);
  if (v1Payload) {
    return v1Payload;
  }

  const v2Inline = cache.get(RESOURCE_CACHE_KEY_V2);
  if (v2Inline) {
    return v2Inline;
  }

  const partsCountRaw = cache.get(RESOURCE_CACHE_PARTS_KEY);
  const partsCount = parseInt(partsCountRaw || '0', 10);
  if (!Number.isFinite(partsCount) || partsCount <= 0) {
    return null;
  }

  const keys = [];
  for (let i = 0; i < partsCount; i++) {
    keys.push(`${RESOURCE_CACHE_PART_PREFIX}${i}`);
  }

  const chunks = cache.getAll(keys);
  const missing = keys.find(k => !chunks[k]);
  if (missing) {
    console.warn(`⚠️ Cache multipart incompleta (${missing}), invalido payload e forzo reload.`);
    _invalidateResourceCacheStorage(cache);
    return null;
  }

  return keys.map(k => chunks[k]).join('');
}

function _writeResourceCachePayload(cache, payload) {
  if (!cache) return;

  // ⚠️ Prima invalidiamo sempre: evita mix V2-inline + multipart stale dopo downgrade/upgrade.
  _invalidateResourceCacheStorage(cache);

  if (payload.length <= RESOURCE_CACHE_MAX_PART_SIZE) {
    cache.put(RESOURCE_CACHE_KEY_V2, payload, RESOURCE_CACHE_TTL_SECONDS);
    return;
  }

  // ⚠️ Multipart è una protezione tecnica contro il limite CacheService (~100KB/entry),
  // non una riduzione funzionale della KB.
  const parts = _splitCachePayload(payload, RESOURCE_CACHE_MAX_PART_SIZE);
  if (!parts.length) {
    throw new Error('Payload cache vuoto: impossibile salvare');
  }

  const values = {};
  parts.forEach((part, idx) => {
    values[`${RESOURCE_CACHE_PART_PREFIX}${idx}`] = part;
  });
  values[RESOURCE_CACHE_PARTS_KEY] = String(parts.length);

  cache.putAll(values, RESOURCE_CACHE_TTL_SECONDS);
  console.warn(`⚠️ Cache risorse salvata in modalità multipart (${parts.length} chunk).`);
}

function _invalidateResourceCacheStorage(cache) {
  if (!cache) return;

  const toRemove = [RESOURCE_CACHE_KEY_V1, RESOURCE_CACHE_KEY_V2, RESOURCE_CACHE_PARTS_KEY];
  const partsCountRaw = cache.get(RESOURCE_CACHE_PARTS_KEY);
  const partsCount = parseInt(partsCountRaw || '0', 10);
  if (Number.isFinite(partsCount) && partsCount > 0) {
    for (let i = 0; i < partsCount; i++) {
      toRemove.push(`${RESOURCE_CACHE_PART_PREFIX}${i}`);
    }
  }

  cache.removeAll(toRemove);
}

/**
 * Svuota manualmente la cache globale (knowledge/config) per forzare reload.
 * Utile come comando da eseguire a mano dall'editor Apps Script.
 */
function clearKnowledgeCache() {
  // ⚠️ Comando operativo ufficiale: resetta RAM + ScriptCache in modo coerente.
  // Evitare reset "parziali" altrove: causano stati fantasma e reload intermittenti.
  GLOBAL_CACHE.loaded = false;
  GLOBAL_CACHE.lastLoadedAt = 0;
  GLOBAL_CACHE.knowledgeBase = '';
  GLOBAL_CACHE.systemEnabled = true;
  GLOBAL_CACHE.ignoreDomains = [];
  GLOBAL_CACHE.ignoreKeywords = [];
  GLOBAL_CACHE.replacements = {};

  // Invalida anche la cache di sistema (CacheService)
  try {
    const cache = CacheService.getScriptCache();
    _invalidateResourceCacheStorage(cache);
  } catch (e) {
    // best effort
  }

  console.log('🗑️ Cache conoscenza/config svuotata manualmente (RAM + ScriptCache)');
}

// Alias per invocazione manuale o da trigger precedenti.
function clearCache() {
  clearKnowledgeCache();
}

/**
 * Forza invalidazione + ricarica immediata usando le funzioni operative già esistenti
 * (`clearKnowledgeCache`/`clearCache` + `loadResources`).
 * Da usare quando cambia il contenuto dei fogli e non si vuole attendere il TTL.
 */
function primeCache() {
  // ⚠️ Orchestrazione voluta: 1) invalidate totale, 2) reload immediato.
  // Non invertire l'ordine (reload->clear) o si ottiene una cache svuotata subito dopo il warm-up.
  clearKnowledgeCache();
  loadResources(true, false);
  const kbSize = (GLOBAL_CACHE.knowledgeBase || '').length;
  console.log(`🔄 Cache primed manualmente (KB=${kbSize} chars).`);
  return {
    loaded: GLOBAL_CACHE.loaded,
    lastLoadedAt: GLOBAL_CACHE.lastLoadedAt,
    knowledgeBaseChars: kbSize
  };
}

function _parseSheetToStructured(data) {
  if (!data || data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  const firstEmptyHeaderIndex = headers.findIndex(h => !h || h === 'null' || h === 'undefined');
  const usedHeaders = (firstEmptyHeaderIndex === -1)
    ? headers
    : headers.slice(0, firstEmptyHeaderIndex);

  return data.slice(1).map(row => {
    const obj = {};
    usedHeaders.forEach((h, i) => {
      if (h) obj[h] = row[i];
    });
    return obj;
  });
}

function _loadAdvancedConfig(ss) {
  const config = { systemEnabled: true, vacationPeriods: [], suspensionRules: {}, ignoreDomains: [], ignoreKeywords: [] };
  const sheet = ss.getSheetByName('Controllo');
  if (!sheet) return config;

  withSheetsRetry(() => {
    // Interruttore
    const status = sheet.getRange('B2').getValue();
    if (String(status).toUpperCase().includes('SPENTO')) config.systemEnabled = false;

    // Filtri anti-spam (layout single-sheet: E11:F)
    const lastDataRow = sheet.getLastRow();
    const filterStartRow = 11;
    const filterRows = lastDataRow >= filterStartRow ? (lastDataRow - filterStartRow + 1) : 0;
    if (filterRows > 0) {
      const filters = sheet.getRange(filterStartRow, 5, filterRows, 2).getValues();
      filters.forEach(row => {
        const domain = String(row[0] || '').trim().toLowerCase();
        const keyword = String(row[1] || '').trim().toLowerCase();
        if (domain) config.ignoreDomains.push(domain);
        if (keyword) config.ignoreKeywords.push(keyword);
      });
    }
  }, 'Lettura configurazione avanzata');

  // Dedup + fallback su config statica
  const staticDomains = (typeof CONFIG !== 'undefined' && Array.isArray(CONFIG.IGNORE_DOMAINS)) ? CONFIG.IGNORE_DOMAINS : [];
  const staticKeywords = (typeof CONFIG !== 'undefined' && Array.isArray(CONFIG.IGNORE_KEYWORDS)) ? CONFIG.IGNORE_KEYWORDS : [];
  config.ignoreDomains = Array.from(new Set([...staticDomains, ...config.ignoreDomains].map(v => String(v).trim().toLowerCase()).filter(Boolean)));
  config.ignoreKeywords = Array.from(new Set([...staticKeywords, ...config.ignoreKeywords].map(v => String(v).trim().toLowerCase()).filter(Boolean)));

  return config;
}

// ====================================================================
// ENTRY POINT PRINCIPALE (TRIGGER)
// ====================================================================



/**
 * Configura tutti i trigger necessari al sistema.
 * Eseguire manualmente una volta per ripristinare i trigger principali.
 */
function setupAllTriggers() {
  // 1. Trigger Principale (Autoresponder)
  setupMainTrigger(5);

  // 2. Trigger Pulizia Memoria (Settimanale)
  setupWeeklyCleanupTrigger();

  // 3. Trigger Metriche/Statistiche (Giornaliero)
  setupMetricsTrigger();

  console.log('✅ Tutti i trigger sono stati riattivati correttamente.');
}

/**
 * Configura il trigger di elaborazione email.
 */
function setupMainTrigger(minutes) {
  const intervalMinutes = parseInt(minutes, 10) || 5;
  deleteTriggersByHandler_('main');
  deleteTriggersByHandler_('processEmailsMain');

  ScriptApp.newTrigger('main')
    .timeBased()
    .everyMinutes(intervalMinutes)
    .create();
}

/**
 * Configura il trigger per l'export delle metriche (ore 23:00).
 */
function setupMetricsTrigger() {
  deleteTriggersByHandler_('exportMetricsToSheet');

  ScriptApp.newTrigger('exportMetricsToSheet')
    .timeBased()
    .atHour(23)
    .everyDays(1)
    .create();
}

/**
 * Elimina trigger esistenti per uno specifico handler, evitando duplicati.
 */
function deleteTriggersByHandler_(handlerName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Configura il trigger per la pulizia settimanale della memoria.
 * La funzione associata (cleanupOldMemory) deve esistere nel progetto.
 */
function setupWeeklyCleanupTrigger() {
  deleteTriggersByHandler_('cleanupOldMemory');
  ScriptApp.newTrigger('cleanupOldMemory')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create();
}

/**
 * Entrypoint alternativo per trigger e script preesistenti.
 * Delega direttamente a main().
 */
function processEmailsMain() {
  return main();
}

/**
 * Verifica che il Gmail Advanced Service sia disponibile e autorizzato.
 * Usa una chiamata leggera (Users.getProfile) per conferma realistica.
 * @returns {{ok: boolean, reason?: string}}
 */
function _checkGmailAdvancedService_() {
  try {
    if (typeof Gmail === 'undefined' || !Gmail.Users || typeof Gmail.Users.getProfile !== 'function') {
      return { ok: false, reason: 'Gmail Advanced Service non disponibile (abilitare in Apps Script).' };
    }

    const profile = Gmail.Users.getProfile('me');
    if (!profile || !profile.emailAddress) {
      return { ok: false, reason: 'Profilo Gmail non accessibile (permessi/scopes mancanti?).' };
    }

    return { ok: true };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    return { ok: false, reason: msg };
  }
}

/**
 * Funzione principale invocata dal trigger temporale (es. ogni 5 min)
 */
function main() {
  console.log('🚀 Avvio pipeline principale');

  // 0. Controllo Preventivo API Avanzate
  const advancedCheck = _checkGmailAdvancedService_();
  if (!advancedCheck.ok) {
    console.error(`💥 CRITICO: Gmail Advanced Service non abilitato o non autorizzato. ${advancedCheck.reason || ''}`);
    return;
  }

  const executionLock = LockService.getScriptLock();
  let hasExecutionLock = false;

  try {
    // 1. Sincronizzazione Esecuzione (Prevenzione concurrency)
    hasExecutionLock = executionLock.tryLock(10000);
    if (!hasExecutionLock) {
      // Nota progettuale: evitiamo di accodare trigger aggiuntivi qui per non creare
      // una tempesta di trigger concorrenti; il trigger periodico successivo riproverà.
      console.warn('⚠️ Esecuzione già in corso o lock bloccato. Salto turno.');
      return;
    }

    // 2. Caricamento Risorse (Config, KB, Blacklist)
    withSheetsRetry(() => loadResources(false, true), 'loadResources(main)');

    if (!GLOBAL_CACHE.loaded) {
      console.error('💥 Risorse non caricate correttamente (GLOBAL_CACHE.loaded=false). Interruzione preventiva.');
      return;
    }

    // 3. Controllo Stato Sistema
    if (!GLOBAL_CACHE.systemEnabled) {
      console.log('🛑 Sistema disattivato da foglio Controllo.');
      return;
    }

    // 4. Orchestrazione Pipeline (Delegato alle classi di servizio)
    const processor = new EmailProcessor();
    const knowledgeBase = GLOBAL_CACHE.knowledgeBase || '';

    const results = processor.processUnreadEmails(knowledgeBase, true); // true = skip double lock

    if (results) {
      console.log(`📊 Batch completato: ${results.total || 0} analizzati, ${results.replied || 0} risposte, ${results.errors || 0} errori.`);
    }

  } catch (error) {
    console.error(`💥 Errore fatale in main: ${error.message}`);
    if (typeof createLogger === 'function') {
      try {
        const logger = createLogger('Main');
        logger.error(`Errore fatale in main: ${error.message}`, {
          stack: error && error.stack ? error.stack : null
        });
      } catch (logError) {
        // Fallback silente
      }
    }
  } finally {
    if (typeof hasExecutionLock !== 'undefined' && hasExecutionLock && executionLock) {
      executionLock.releaseLock();
    }
  }
}



/**
 * Serializza righe foglio in testo robusto per prompt/validator.
 * - converte in stringa e trimma ogni cella
 * - rimuove celle vuote
 * - rimuove righe completamente vuote
 */
function _sheetRowsToText(rows) {
  if (!Array.isArray(rows)) return '';

  return rows
    .map(row => {
      const safeRow = Array.isArray(row) ? row : [row];
      return safeRow
        .map(cell => _formatCellForKnowledgeText(cell))
        .filter(Boolean)
        .join(' | ');
    })
    .filter(Boolean)
    .join('\n');
}

/**
 * Normalizza la serializzazione celle per evitare output locale-dipendente.
 * In particolare, le Date di Google Sheets vengono convertite in formato stabile
 * (YYYY-MM-DD oppure YYYY-MM-DD HH:mm) invece di "Tue May 12 2026 ...".
 */
function _formatCellForKnowledgeText(cell) {
  if (cell == null) return '';

  if (typeof cell === 'string' && cell.trim().startsWith('#')) return ''; // Salta errori formula tipo #REF! o #N/A

  if (cell instanceof Date && !isNaN(cell.getTime())) {
    return _formatDateForKnowledgeText(cell);
  }

  // Evita che ritorni a capo dentro una singola cella spezzino la struttura
  // del testo KB (una riga Sheet deve restare una riga logica nel prompt).
  return String(cell)
    .replace(/\r\n?|\n/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function _formatDateForKnowledgeText(date) {
  if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
    const tz = (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function')
      ? Session.getScriptTimeZone()
      : 'UTC';

    // Rileviamo se ha una parte oraria in base al fuso orario target
    const hasTime = (tz === 'UTC')
      ? (date.getUTCHours() !== 0 || date.getUTCMinutes() !== 0 || date.getUTCSeconds() !== 0)
      : (date.getHours() !== 0 || date.getMinutes() !== 0 || date.getSeconds() !== 0);

    const pattern = hasTime ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd';

    return Utilities.formatDate(date, tz, pattern);
  }

  // Fallback Node/tests: serializzazione stabile UTC
  const hasUtcTime = date.getUTCHours() !== 0 || date.getUTCMinutes() !== 0 || date.getUTCSeconds() !== 0;
  const yyyy = date.getUTCFullYear();
  const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(date.getUTCDate()).padStart(2, '0');

  if (!hasUtcTime) {
    return `${yyyy}-${mm}-${dd}`;
  }

  const hh = String(date.getUTCHours()).padStart(2, '0');
  const min = String(date.getUTCMinutes()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
}
