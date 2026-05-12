/**
 * GeminiRateLimiter.gs - Gestione Quote API Gemini
 * 
 * SINCRONO - Compatibile con Google Apps Script
 * Configurazione modelli centralizzata in gas_config.js
 * Reset quota: ore 9:00 italiane (mezzanotte Pacific)
 * Cache ottimizzata per ridurre letture PropertiesService
 * 
 * FUNZIONALITÀ:
 * - Traccia utilizzo RPM (richieste/minuto), TPM (token/minuto), RPD (richieste/giorno)
 * - Seleziona automaticamente il modello disponibile
 * - Applica throttling quando ci si avvicina ai limiti
 * - Passa al modello di riserva se il principale è esaurito
 */
class GeminiRateLimiter {
  constructor() {
    console.log('\uD83D\uDEA6 Inizializzazione GeminiRateLimiter...');

    // ================================================================
    // CONFIGURAZIONE MODELLI (Legge da CONFIG)
    // ================================================================

    // Legge modelli da CONFIG.GEMINI_MODELS (centralizzato)
    if (typeof CONFIG !== 'undefined' && CONFIG.GEMINI_MODELS) {
      this.models = this._normalizeDeprecatedModelNames(CONFIG.GEMINI_MODELS);
      console.log('   \u2713 Modelli caricati da CONFIG.GEMINI_MODELS');
    } else {
      // Fallback se CONFIG non disponibile
      console.warn('   \u26A0\uFE0F CONFIG.GEMINI_MODELS non trovato, uso default');
      this.models = {
        'flash-3.1-lite': {
          name: 'gemini-3.1-flash-lite',
          rpm: 2000, tpm: 2000000, rpd: 3500,
          useCases: ['generation', 'all']
        },
        'flash-lite': {
          name: 'gemini-3.1-flash-lite',
          rpm: 2000, tpm: 2000000, rpd: 3500,
          useCases: ['fallback', 'classification', 'quick_check']
        }
      };
    }

    // Legge strategia da CONFIG.MODEL_STRATEGY (centralizzato)
    if (typeof CONFIG !== 'undefined' && CONFIG.MODEL_STRATEGY) {
      this.strategies = CONFIG.MODEL_STRATEGY;
      console.log('   \u2713 Strategia caricata da CONFIG.MODEL_STRATEGY');
    } else {
      // Fallback default
      this.strategies = {
        'quick_check': ['flash-lite', 'flash-2.5'],
        // Per generation evitiamo loop sul solo "lite" e usiamo fallback premium-safe.
        'generation': ['flash-2.5', 'flash-lite', 'flash-2.5-lite-backup'],
        'fallback': ['flash-lite', 'flash-2.5-lite-backup', 'flash-2.5']
      };
    }

    // Modello di default (primo nella lista generation)
    this.defaultModel = Object.keys(this.models)[0] || 'flash-2.5';

    // ================================================================
    // CACHE IN-MEMORY (Session-based)
    // ================================================================

    this.cache = {
      rpmWindow: [],
      tpmWindow: [],
      lastCacheUpdate: 0
    };

    // ================================================================
    // PERSISTENZA (PropertiesService Optimized)
    // ================================================================

    this.props = (typeof PropertiesService !== 'undefined' && PropertiesService && typeof PropertiesService.getScriptProperties === 'function')
      ? PropertiesService.getScriptProperties()
      : {
        getProperty: () => null,
        setProperty: () => { },
        getProperties: () => ({}),
        setProperties: () => { },
        deleteProperty: () => { }
      };

    // ⚠️ OTTIMIZZAZIONE: Caricamento batch iniziale per evitare PropertiesService.getProperty() ripetuti.
    this.allProps = this.props.getProperties() || {};
    this.dirtyProps = {};

    // Sincronizza cache con stato persistito iniziale
    this.cache.rpmWindow = this._safeParseJson(this.allProps['rpm_window'], []);
    this.cache.tpmWindow = this._safeParseJson(this.allProps['tpm_window'], []);
    this.cache.lastCacheUpdate = Date.now();

    // Recupera da WAL se presente (crash recovery)
    this._recoverFromWAL();

    // Inizializza contatori se non esistono
    this._initializeCounters();

    // ================================================================
    // CONFIGURAZIONE THROTTLING
    // ================================================================

    this.safetyMargin = {
      rpm: 0.8,
      tpm: 0.8,
      rpd: 0.9
    };

    this.throttleDelays = {
      rpm: 5000,   // 5 secondi
      tpm: 3000,
      rpd: 10000
    };

    // Exponential backoff
    this.backoffBase = 2000;
    this.backoffMultiplier = 2;
    this.maxBackoff = 60000;

    console.log('✓ GeminiRateLimiter inizializzato');
    console.log(`   Modelli: ${Object.keys(this.models).join(', ')}`);
    console.log(`   Default: ${this.defaultModel}`);
  }

  /**
   * Sostituisce nomi modello deprecati/ritirati con equivalenti supportati.
   * Mantiene la stessa struttura dell'oggetto modelli.
   *
   * @param {Object<string, {name: string}>} models
   * @returns {Object<string, {name: string}>}
   */
  _normalizeDeprecatedModelNames(models) {
    const deprecatedMap = {
      // Mappatura modelli ritirati verso lo standard 3.1 Lite.
      // Il modello 2.5 Flash standard viene mantenuto come da richiesta utente.
      'gemini-2.5-flash-lite': 'gemini-3.1-flash-lite',
      'gemini-2.5-flash-exp': 'gemini-3.1-flash-lite',
      'gemini-2.0-flash-exp': 'gemini-3.1-flash-lite',
      'gemini-2.0-flash': 'gemini-3.1-flash-lite'
    };

    const normalized = {};
    Object.keys(models || {}).forEach(modelKey => {
      const modelConfig = models[modelKey] || {};
      const currentName = modelConfig.name;
      const replacement = deprecatedMap[currentName];

      if (replacement) {
        console.warn(`⚠️ Modello deprecato rilevato per '${modelKey}': ${currentName} → ${replacement}`);
      }

      normalized[modelKey] = Object.assign({}, modelConfig, {
        name: replacement || currentName
      });
    });

    return normalized;
  }

  // ================================================================
  // INIZIALIZZAZIONE
  // ================================================================

  _initializeCounters() {
    // Usa ScriptLock per sincronizzare il reset tra esecuzioni parallele
    const lock = LockService.getScriptLock();
    let lockAcquired = false;
    try {
      // Tenta di acquisire il lock per 5 secondi
      if (lock.tryLock(5000)) {
        lockAcquired = true;
        // Usa data Pacific per allinearsi al reset reale delle quote Google
        // Il reset Google avviene a mezzanotte Pacific = 9:00 AM italiana
        const pacificDate = this._getPacificDate();
        const storedDate = this.props.getProperty('rate_limit_date');

        // Reset quando cambia la data Pacific (non italiana!)
        if (storedDate !== pacificDate) {
          console.log(`📅 Giorno Pacific cambiato (${pacificDate}), reset contatori giornalieri`);
          console.log(`   (Ora italiana: ${Utilities.formatDate(new Date(), 'Europe/Rome', 'HH:mm')})`);
          this._resetDailyCounters();
          this.props.setProperty('rate_limit_date', pacificDate);
        }
      } else {
        console.warn('⚠️ Impossibile acquisire lock per reset quota, salto controllo');
      }
    } catch (e) {
      console.error(`❌ Errore durante lock inizializzazione quota: ${e.message}`);
    } finally {
      if (lockAcquired) {
        try {
          lock.releaseLock();
        } catch (e) {
          console.warn(`⚠️ Errore rilascio lock (QuotaReset): ${e.message}`);
        }
      }
    }
  }

  _resetDailyCounters() {
    const todayPacific = this._getPacificDate();
    for (const modelKey in this.models) {
      this.props.setProperty(`rpd_${modelKey}`, '0');
      this.props.setProperty(`rpd_date_${modelKey}`, todayPacific);
      this.props.setProperty(`tokens_${modelKey}`, '0');
    }
    // Reset anche cache
    this.props.setProperty('rpm_window', JSON.stringify([]));
    this.props.setProperty('tpm_window', JSON.stringify([]));
    console.log('✓ Contatori giornalieri resettati');
  }

  /**
   * Ottieni data in formato italiano (per logging user-friendly)
   */
  _getItalianDate() {
    const now = new Date();
    const italianDate = Utilities.formatDate(now, 'Europe/Rome', 'yyyy-MM-dd');
    return italianDate;
  }

  /**
   * Ottieni data Pacific (per reset quote Google)
   * Il reset delle quote Google avviene a mezzanotte Pacific Time.
   * Mezzanotte Pacific = 9:00 AM italiana (in inverno, 8:00 in estate con DST)
   */
  _getPacificDate() {
    const now = new Date();
    try {
      // America/Los_Angeles gestisce automaticamente DST (PST/PDT)
      const pacificDate = Utilities.formatDate(now, 'America/Los_Angeles', 'yyyy-MM-dd');

      const month = now.getMonth();
      if (month === 2 || month === 10) {
        const hour = parseInt(Utilities.formatDate(now, 'America/Los_Angeles', 'HH'), 10);
        if (hour >= 0 && hour <= 3) {
          console.warn(`⚠️ Possibile transizione DST in corso, ora Pacific: ${hour}`);
        }
      }

      return pacificDate;
    } catch (error) {
      console.error(`❌ Errore getPacificDate: ${error.message}`);
      return Utilities.formatDate(now, 'UTC', 'yyyy-MM-dd');
    }
  }

  // ================================================================
  // SELEZIONE MODELLO
  // ================================================================

  selectModel(taskType, options) {
    options = options || {};
    const preferQuality = options.preferQuality || false;
    const forceModel = options.forceModel || null;
    const estimatedTokens = options.estimatedTokens || 1000;

    // Override manuale
    if (forceModel && this.models[forceModel]) {
      return this._validateModelAvailability(forceModel, estimatedTokens);
    }

    // Usa strategia da CONFIG (o fallback)
    // Aggiunge 'classification' come alias di 'quick_check' se non definito
    const taskStrategies = Object.assign({}, this.strategies);
    if (!taskStrategies['classification']) {
      taskStrategies['classification'] = taskStrategies['quick_check'] || ['flash-lite', 'flash-2.5'];
    }

    const candidates = taskStrategies[taskType] || taskStrategies['fallback'] || ['flash-lite', 'flash-2.5'];

    // Trova primo modello disponibile (Punto 11: Protezione con lock per atomicità check+use)
    const lock = LockService.getScriptLock();
    // Retry breve con backoff: evita falsi "non disponibile" sotto burst concorrenti.
    let lockAcquired = false;
    const maxLockAttempts = 3;
    for (let attempt = 0; attempt < maxLockAttempts; attempt++) {
      lockAcquired = lock.tryLock(5000);
      if (lockAcquired) break;
      if (attempt < maxLockAttempts - 1) {
        const backoffMs = 150 * Math.pow(2, attempt) + Math.floor(Math.random() * 100);
        Utilities.sleep(backoffMs);
      }
    }

    if (!lockAcquired) {
      console.warn('⚠️ Lock non acquisito per selezione modello dopo retry, fallback conservativo');
      return {
        available: false,
        modelKey: null,
        reason: 'lock_timeout',
        nextResetTime: this._getNextResetTime()
      };
    }

    let selectedResult = null;
    try {
      // Rinfresca i contatori per avere dati aggiornati sotto lock
      this._refreshCache();

      for (var i = 0; i < candidates.length; i++) {
        const modelKey = candidates[i];
        const result = this._validateModelAvailability(modelKey, estimatedTokens);
        if (result.available) {
          console.log(`✓ Selezionato: ${modelKey} per ${taskType}`);
          selectedResult = result;
          break;
        }
      }

    } finally {
      if (lockAcquired) lock.releaseLock();
    }

    return selectedResult || {
      available: false,
      modelKey: null,
      reason: 'all_quotas_exhausted',
      nextResetTime: this._getNextResetTime()
    };
  }

  _validateModelAvailability(modelKey, estimatedTokens) {
    const model = this.models[modelKey];
    if (!model) {
      return { available: false, reason: 'model_not_found' };
    }

    // Identifica il modello fisico condiviso per aggregare i consumi (es. flash-lite e flash-3.1-lite)
    const physicalModelName = model.name;
    let rpdUsed = 0;
    let rpmUsed = 0;
    let tpmUsed = 0;

    for (const key of Object.keys(this.models)) {
      if (this.models[key].name === physicalModelName) {
        rpdUsed += parseInt(this.props.getProperty(`rpd_${key}`) || '0', 10) || 0;
        rpmUsed += this._getRequestsInWindow('rpm', key);
        tpmUsed += this._getTokensInWindow('tpm', key);
      }
    }

    // Controllo RPD
    const rpdLeft = model.rpd - rpdUsed;

    if (rpdLeft <= 0) {
      return {
        available: false,
        modelKey: modelKey,
        reason: 'rpd_exhausted',
        quotaLeft: { rpd: 0 }
      };
    }

    // Controllo RPM (ultimo minuto)
    const rpmLeft = model.rpm - rpmUsed;

    if (rpmLeft <= 0) {
      return {
        available: false,
        modelKey: modelKey,
        reason: 'rpm_exhausted',
        retryAfter: 60
      };
    }

    // Controllo TPM (ultimo minuto)
    const tpmLeft = model.tpm - tpmUsed;

    if (tpmLeft < estimatedTokens) {
      return {
        available: false,
        modelKey: modelKey,
        reason: 'tpm_insufficient',
        quotaLeft: { tpm: tpmLeft },
        retryAfter: 60
      };
    }

    // Modello disponibile
    return {
      available: true,
      modelKey: modelKey,
      model: model,
      quotaLeft: {
        rpd: rpdLeft,
        rpm: rpmLeft,
        tpm: tpmLeft
      },
      shouldThrottle: this._shouldThrottle(modelKey, rpdUsed, rpmUsed, tpmUsed)
    };
  }

  _shouldThrottle(modelKey, rpdUsed, rpmUsed, tpmUsed) {
    const model = this.models[modelKey];

    const rpdRatio = rpdUsed / model.rpd;
    const rpmRatio = rpmUsed / model.rpm;
    const tpmRatio = tpmUsed / model.tpm;

    if (rpdRatio >= this.safetyMargin.rpd) {
      return { needed: true, reason: 'rpd', delay: this.throttleDelays.rpd };
    }
    if (rpmRatio >= this.safetyMargin.rpm) {
      return { needed: true, reason: 'rpm', delay: this.throttleDelays.rpm };
    }
    if (tpmRatio >= this.safetyMargin.tpm) {
      return { needed: true, reason: 'tpm', delay: this.throttleDelays.tpm };
    }

    return { needed: false };
  }

  // ================================================================
  // ESECUZIONE RICHIESTA (SINCRONO)
  // ================================================================

  /**
   * Esegue richiesta con rate limiting
   * VERSIONE SINCRONA per Google Apps Script
   * 
   * @param {string} taskType - Tipo task: 'quick_check', 'generation', etc.
   * @param {Function} requestFn - Funzione che riceve modelName ed esegue la richiesta
   * @param {Object} options - {estimatedTokens, maxRetries, preferQuality}
   * @returns {Object} {success, result, modelUsed, quotaUsed}
   */
  executeRequest(taskType, requestFn, options) {
    options = options || {};

    // ═══════════════════════════════════════════════════════════════════
    // BYPASS PER CHIAVE DI RISERVA
    // Se stiamo usando una chiave esterna, NON dobbiamo tracciare i consumi
    // sul Rate Limiter locale per non inquinare le statistiche della chiave principale
    // ═══════════════════════════════════════════════════════════════════
    if (options.skipRateLimit) {
      console.warn('\u23E9 RateLimiter BYPASSED (Chiave di Riserva in uso)');
      try {
        const startTime = Date.now();
        // Esecuzione diretta senza controlli quota
        const result = requestFn(options.modelNameOverride || 'gemini-2.5-flash');
        const duration = Date.now() - startTime;

        return {
          success: true,
          result: result,
          modelUsed: options.modelNameOverride || 'backup-model',
          quotaUsed: { rpd: 0, rpm: 0 }, // Statistiche fittizie per non sporcare contatori
          duration: duration
        };
      } catch (e) {
        // Se fallisce, rilancia l'errore per gestione esterna
        throw e;
      }
    }
    // ═══════════════════════════════════════════════════════════════════

    const estimatedTokens = options.estimatedTokens || 1000;
    const maxRetries = options.maxRetries || 3;
    const preferQuality = options.preferQuality || false;

    // 1. Selezione modello standard
    const selection = this.selectModel(taskType, { preferQuality: preferQuality });

    if (!selection.available) {
      console.error(`\u274C Nessun modello disponibile: ${selection.reason}`);
      throw new Error('QUOTA_EXHAUSTED: ' + selection.reason);
    }

    const modelKey = selection.modelKey;
    const model = selection.model;
    const shouldThrottle = selection.shouldThrottle;

    // 2. Throttling
    if (shouldThrottle && shouldThrottle.needed) {
      console.warn(`\u23F8\uFE0F Throttling (${shouldThrottle.reason}): ${shouldThrottle.delay}ms`);
      Utilities.sleep(shouldThrottle.delay);
    }

    // 3. Esecuzione con retry (sincrono)
    var lastError = null;
    for (var attempt = 0; attempt < maxRetries; attempt++) {
      try {
        const startTime = Date.now();

        console.log(`🚀 Tentativo richiesta ${attempt + 1}/${maxRetries}`);
        console.log(`   Modello: ${model.name}, Task: ${taskType}`);

        // CHIAMATA SINCRONA (no await)
        const result = requestFn(model.name);

        const duration = Date.now() - startTime;

        // Traccia richiesta riuscita
        this._trackRequest(modelKey, estimatedTokens, duration);

        console.log(`✓ Successo (${duration}ms)`);

        return {
          success: true,
          result: result,
          modelUsed: model.name,
          modelKey: modelKey,
          duration: duration,
          quotaUsed: {
            rpd: parseInt(this._getProp(`rpd_${modelKey}`) || '0', 10) || 0,
            rpm: this._getRequestsInWindow('rpm', modelKey)
          }
        };

      } catch (error) {
        lastError = error;
        const errorMsg = error.message || '';

        // Controllo se errore 429
        if (errorMsg.indexOf('429') !== -1 ||
          errorMsg.indexOf('rate limit') !== -1 ||
          errorMsg.indexOf('quota') !== -1) {

          console.warn(`⚠️  Limite quota (429) al tentativo ${attempt + 1}`);

          if (attempt < maxRetries - 1) {
            const backoffDelay = Math.min(
              this.backoffBase * Math.pow(this.backoffMultiplier, attempt),
              this.maxBackoff
            );
            console.log(`   Attesa ${backoffDelay}ms...`);
            Utilities.sleep(backoffDelay);
            continue;
          }
        } else {
          // Errore non ritentabile
          throw error;
        }
      }
    }

    // Tutti i tentativi falliti
    console.error(`❌ Tutti i ${maxRetries} tentativi falliti`);
    throw lastError || new Error('Richiesta fallita dopo tutti i tentativi');
  }

  /**
   * Traccia una richiesta ausiliaria (es. creazione cache)
   * Non influisce sui token ma consuma RPD
   */
  trackAuxiliaryRequest(modelKey, tokensUsed, reason) {
    console.log(`📊 Tracciamento richiesta ausiliaria (${reason || 'aux'}): ${modelKey}`);
    this._trackRequest(modelKey, tokensUsed || 0, 0);
  }

  /**
   * Traccia una richiesta con Grounding (Google Search)
   * Consuma quota specifica per Grounding
   */
  trackGroundingRequest(modelKey, tokensUsed) {
    console.log(`🌐 Tracciamento Grounding: ${modelKey}`);
    // Incrementa contatore grounding specifico
    this._incrementGroundingCounter_(modelKey);
    // E poi traccia come richiesta normale
    this._trackRequest(modelKey, tokensUsed, 0);
  }

  _incrementGroundingCounter_(modelKey) {
    // ⚠️ OTTIMIZZAZIONE: Incremento in memoria (deferred persistence)
    const key = 'grounding_count_' + this._getItalianDate();
    const current = parseInt(this._getProp(key) || '0', 10);
    this._setProp(key, String(current + 1));
  }

  // ================================================================
  // TRACKING (Ottimizzato con Session Cache)
  // ================================================================

  _trackRequest(modelKey, tokensUsed, duration) {
    const now = Date.now();
    const nonce = `${Math.floor(Math.random() * 1000000)}`;

    // 1. Contatori RPD/Tokens (In memoria)
    const counters = this._incrementCountersAtomic(modelKey, tokensUsed);

    // 2. Aggiornamento finestre RPM/TPM in memoria
    this._updateWindow('rpm', {
      timestamp: now,
      nonce: nonce,
      modelKey: modelKey
    });

    this._updateWindow('tpm', {
      timestamp: now,
      nonce: nonce,
      modelKey: modelKey,
      tokens: tokensUsed
    });

    // ⚠️ OTTIMIZZAZIONE: Rimosso _persistCache() ad ogni richiesta.
    // La persistenza è ora differita al termine della pipeline (metodo flush()).
    console.log(`📊 Tracciato: ${modelKey} | RPD: ${counters.rpd}`);
  }

  /**
   * Incrementa contatori RPD e Token in modo atomico.
   * In questa versione ottimizzata, lavora primariamente in memoria
   * poiché l'esecuzione è protetta da un lock globale nel main().
   */
  _incrementCountersAtomic(modelKey, tokensUsed) {
    const rpdKey = 'rpd_' + modelKey;
    const rpdDateKey = 'rpd_date_' + modelKey;
    const tokensKey = 'tokens_' + modelKey;
    const todayPacific = this._getPacificDate();

    const lastRpdDate = this._getProp(rpdDateKey) || '';
    let currentRpd = parseInt(this._getProp(rpdKey) || '0', 10) || 0;
    let currentTokens = parseInt(this._getProp(tokensKey) || '0', 10) || 0;

    if (lastRpdDate !== todayPacific) {
      currentRpd = 0;
      currentTokens = 0;
    }

    const nextRpd = currentRpd + 1;
    const nextTokens = currentTokens + (tokensUsed || 0);

    this._setProp(rpdKey, String(nextRpd));
    this._setProp(tokensKey, String(nextTokens));
    this._setProp(rpdDateKey, todayPacific);

    return { rpd: nextRpd, tokens: nextTokens };
  }

  _updateWindow(windowType, entry) {
    const now = Date.now();

    // Aggiungi a cache in memoria
    const cacheKey = windowType + 'Window';
    this.cache[cacheKey].push(entry);

    // Pulisci vecchie entry (>60 secondi)
    this.cache[cacheKey] = this.cache[cacheKey].filter(function (e) {
      return now - e.timestamp < 60000;
    });

    // Limita dimensioni array (max 100 per finestra)
    if (this.cache[cacheKey].length > 100) {
      this.cache[cacheKey] = this.cache[cacheKey].slice(-100);
    }
  }

  /**
   * Salva tutte le modifiche pendenti su PropertiesService in un unico colpo.
   * Deve essere chiamato alla fine della pipeline (main finally block).
   */
  flush() {
    const keys = Object.keys(this.dirtyProps);
    const hasCacheChanges = true; // Assumiamo sempre RPM/TPM aggiornati in una run

    if (keys.length === 0 && !hasCacheChanges) return;

    const lock = LockService.getScriptLock();
    if (lock.tryLock(10000)) {
      try {
        console.log(`💾 Persistenza differita Rate Limiter: salvataggio proprietà...`);

        // 1. Sincronizza finestre RPM/TPM
        this.dirtyProps['rpm_window'] = JSON.stringify(this.cache.rpmWindow);
        this.dirtyProps['tpm_window'] = JSON.stringify(this.cache.tpmWindow);

        // 2. Checkpoint WAL (per sicurezza batch)
        const wal = {
          timestamp: Date.now(),
          rpm: this.cache.rpmWindow.slice(),
          tpm: this.cache.tpmWindow.slice()
        };
        this.dirtyProps['rate_limit_wal'] = JSON.stringify(wal);

        // 3. Scrittura batch atomica
        this.props.setProperties(this.dirtyProps);

        // 4. Rimuovi WAL
        this.props.deleteProperty('rate_limit_wal');

        // Reset stato dirty
        this.dirtyProps = {};
        console.log('✅ Persistenza completata.');
      } catch (e) {
        console.error(`❌ Errore durante flush Rate Limiter: ${e.message}`);
      } finally {
        lock.releaseLock();
      }
    } else {
      console.warn('⚠️ Impossibile acquisire lock per flush Rate Limiter. Modifiche perse.');
    }
  }

  // ================================================================
  // HELPERS PROPRIETÀ (Memory First)
  // ================================================================

  _getProp(key) {
    return this.dirtyProps[key] !== undefined ? this.dirtyProps[key] : this.allProps[key];
  }

  _setProp(key, val) {
    this.dirtyProps[key] = val;
  }

  _safeParseJson(str, fallback) {
    try {
      return str ? JSON.parse(str) : fallback;
    } catch (e) {
      return fallback;
    }
  }

  /**
   * Recupera dati da WAL dopo un crash
   * Chiamato nel constructor prima di inizializzare i contatori
   */
  _recoverFromWAL() {
    // Punto 1: Aggiunto lock per garantire atomicità durante il recovery
    const lock = LockService.getScriptLock();
    const lockAcquired = lock.tryLock(5000);
    if (!lockAcquired) {
      console.warn('⚠️ Recovery WAL saltato: impossibile acquisire lock entro 5s');
      return;
    }

    try {
      const walData = this.props.getProperty('rate_limit_wal');
      if (!walData) return;

      console.warn('⚠️ WAL trovato - recupero dati dopo crash...');
      const wal = JSON.parse(walData);
      if (!wal || typeof wal !== 'object' || !wal.timestamp) {
        console.error('❌ WAL corrotto, ignoro');
        this.props.deleteProperty('rate_limit_wal');
        return;
      }
      if (!Array.isArray(wal.rpm) || !Array.isArray(wal.tpm)) {
        console.error('❌ WAL con array invalidi');
        this.props.deleteProperty('rate_limit_wal');
        return;
      }

      // Verifica che il WAL non sia troppo vecchio (> 5 minuti)
      const age = Date.now() - wal.timestamp;
      if (age > 300000) {
        console.warn('   WAL troppo vecchio, ignorato');
        this.props.deleteProperty('rate_limit_wal');
        return;
      }

      // Leggi dati attuali
      const currentRpm = JSON.parse(this.props.getProperty('rpm_window') || '[]');
      const currentTpm = JSON.parse(this.props.getProperty('tpm_window') || '[]');

      // Merge WAL con dati esistenti (evita duplicati per timestamp)
      const mergedRpm = this._mergeWindowData(currentRpm, wal.rpm || []);
      const mergedTpm = this._mergeWindowData(currentTpm, wal.tpm || []);

      // Salva dati recuperati
      this.props.setProperty('rpm_window', JSON.stringify(mergedRpm));
      this.props.setProperty('tpm_window', JSON.stringify(mergedTpm));

      // Rimuovi WAL dopo recovery
      this.props.deleteProperty('rate_limit_wal');

      // Aggiorna cache in-memory
      this.cache.rpmWindow = mergedRpm;
      this.cache.tpmWindow = mergedTpm;
      this.cache.lastCacheUpdate = Date.now();

      console.log('✓ Dati recuperati da WAL con successo e cache aggiornata');

    } catch (error) {
      console.error(`❌ Errore recovery WAL: ${error.message}`);
      // Rimuovi WAL corrotto
      try { this.props.deleteProperty('rate_limit_wal'); } catch (e) { }
    } finally {
      if (lockAcquired) lock.releaseLock();
    }
  }

  /**
   * Merge dati finestra evitando duplicati
   */
  _mergeWindowData(existing, walData) {
    const toKey = (entry) => {
      const ts = entry && typeof entry.timestamp !== 'undefined' ? entry.timestamp : 'na';
      const nonce = entry && typeof entry.nonce !== 'undefined' ? entry.nonce : 'na';
      const model = entry && entry.modelKey ? entry.modelKey : 'na';
      const tokens = entry && typeof entry.tokens !== 'undefined' ? entry.tokens : 0;
      return `${ts}|${nonce}|${model}|${tokens}`;
    };

    const existingEntries = new Set(existing.map(toKey));
    const merged = JSON.parse(JSON.stringify(existing));

    for (const entry of walData) {
      const entryKey = toKey(entry);
      if (!existingEntries.has(entryKey)) {
        merged.push(Object.assign({}, entry));
        existingEntries.add(entryKey);
      }
    }

    // Ordina per timestamp e limita
    return merged
      .sort((a, b) => a.timestamp - b.timestamp)
      .slice(-100);
  }

  _getRequestsInWindow(type, modelKey) {
    const window = (type === 'rpm') ? this.cache.rpmWindow : this.cache.tpmWindow;
    const now = Date.now();
    const oneMinuteAgo = now - 60000;

    const filtered = window.filter(entry =>
      entry.timestamp > oneMinuteAgo && entry.modelKey === modelKey
    );

    if (type === 'rpm') {
      return filtered.length;
    } else {
      return filtered.reduce((sum, entry) => sum + (entry.tokens || 0), 0);
    }
  }

  _getTokensInWindow(type, modelKey) {
    // Riusa la logica unificata di _getRequestsInWindow per coerenza
    return this._getRequestsInWindow(type, modelKey);
  }

  // ================================================================
  // STATISTICHE
  // ================================================================

  getUsageStats() {
    const stats = {
      date: this._getItalianDate(),
      italianTime: Utilities.formatDate(new Date(), 'Europe/Rome', 'HH:mm'),
      pacificTime: Utilities.formatDate(new Date(), 'America/Los_Angeles', 'HH:mm') + ' (PST/PDT)',
      nextReset: this._getNextResetTime(),
      nextResetPacific: '00:00 Pacific Time', // Reset Google è sempre mezzanotte Pacific
      models: {}
    };

    for (var modelKey in this.models) {
      const model = this.models[modelKey];
      const rpdUsed = parseInt(this._getProp('rpd_' + modelKey) || '0', 10);
      const tokensUsed = parseInt(this._getProp('tokens_' + modelKey) || '0', 10);
      const rpmUsed = this._getRequestsInWindow('rpm', modelKey);
      const tpmUsed = this._getTokensInWindow('tpm', modelKey);

      stats.models[modelKey] = {
        name: model.name,
        rpd: {
          used: rpdUsed,
          limit: model.rpd,
          percent: (rpdUsed / model.rpd * 100).toFixed(1)
        },
        rpm: {
          used: rpmUsed,
          limit: model.rpm,
          percent: (rpmUsed / model.rpm * 100).toFixed(1)
        },
        tpm: {
          used: tpmUsed,
          limit: model.tpm,
          percent: (tpmUsed / model.tpm * 100).toFixed(1)
        },
        tokensToday: tokensUsed
      };
    }

    return stats;
  }

  logUsageStats() {
    const stats = this.getUsageStats();

    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('📊 UTILIZZO QUOTA GEMINI - ' + stats.date + ' ' + stats.italianTime);
    console.log('⏰ Prossimo reset: ' + stats.nextReset + ' (9:00 AM italiana)');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    for (var modelKey in stats.models) {
      const model = stats.models[modelKey];
      console.log('\n' + modelKey.toUpperCase() + ' (' + model.name + '):');
      console.log('  RPD: ' + model.rpd.used + '/' + model.rpd.limit + ' (' + model.rpd.percent + '%)');
      console.log('  RPM: ' + model.rpm.used + '/' + model.rpm.limit + ' (' + model.rpm.percent + '%)');
      console.log('  Token oggi: ' + model.tokensToday.toLocaleString());
    }

    console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  }

  _getNextResetTime() {
    const nowMs = Date.now();
    const currentPacificDate = this._getPacificDate();

    // Trova il primo istante in cui cambia la data Pacific (mezzanotte locale Pacific)
    let low = nowMs;
    let high = nowMs + (36 * 60 * 60 * 1000); // margine abbondante per DST

    while (high - low > 1000) {
      const mid = Math.floor((low + high) / 2);
      const midPacificDate = Utilities.formatDate(new Date(mid), 'America/Los_Angeles', 'yyyy-MM-dd');

      if (midPacificDate === currentPacificDate) {
        low = mid;
      } else {
        high = mid;
      }
    }

    return new Date(high).toISOString();
  }

  /**
   * Stima il numero di token per un testo
   * Formula: parole * 1.25 + overhead 10%
   */
  _estimateTokens(text) {
    if (!text) return 0;

    const wordCount = text.trim().split(/\s+/).filter(Boolean).length;
    const baseTokens = Math.ceil(wordCount * 1.25);
    const overhead = Math.ceil(baseTokens * 0.1);
    const charEstimate = Math.ceil(text.length / 3.5);

    return Math.max(baseTokens + overhead, charEstimate, 1);
  }
}

// ================================================================
// FUNZIONI UTILITÀ (per dashboard e manutenzione)
// ================================================================

/**
 * Dashboard quota (esegui manualmente da editor script)
 */
function showQuotaDashboard() {
  const limiter = new GeminiRateLimiter();
  limiter.logUsageStats();

  // Avviso se >80%
  const stats = limiter.getUsageStats();
  for (var modelKey in stats.models) {
    const model = stats.models[modelKey];
    if (parseFloat(model.rpd.percent) > 80) {
      console.warn('⚠️  ATTENZIONE: ' + modelKey + ' RPD > 80% (' + model.rpd.percent + '%)');
    }
  }
}

/**
 * Reset manuale contatori (usare con cautela!)
 */
function resetQuotaCounters() {
  const limiter = new GeminiRateLimiter();
  limiter._resetDailyCounters();
  limiter.props.setProperty('rate_limit_date', limiter._getPacificDate());
  console.log('✓ Contatori quota resettati manualmente (usando data Pacific)');
}