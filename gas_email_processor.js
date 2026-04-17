/**
 * EmailProcessor.js - Orchestratore Pipeline Email
 * 
 * PIPELINE ELABORAZIONE:
 * 1. FILTRA: Dobbiamo processare questa email?
 * 2. CLASSIFICA: Che tipo di richiesta è?
 * 3. GENERA: Crea risposta AI
 * 4. VALIDA: Controlla qualità risposta
 * 5. INVIA: Rispondi all'email
 * 
 * FUNZIONALITÀ AVANZATE:
 * - Lock a livello thread (anti race condition)
 * - Anti-loop detection
 * - Salutation mode (full/soft/none_or_continuity/session)
 * - KB enrichment condizionale
 * - Memory tracking
 */

class EmailProcessor {
  constructor(options = {}) {
    // Logger strutturato
    this.logger = (typeof createLogger === 'function')
      ? createLogger('EmailProcessor')
      : {
        info: (...args) => console.log(...args),
        warn: (...args) => console.warn(...args),
        error: (...args) => console.error(...args),
        debug: (...args) => console.log(...args),
      };
    this.logger.info('Inizializzazione EmailProcessor');

    // Inietta dipendenze o crea default
    this.geminiService = options.geminiService || new GeminiService();
    this.classifier = options.classifier || new Classifier();
    this.requestClassifier = options.requestClassifier ||
      (typeof RequestTypeClassifier !== 'undefined'
        ? new RequestTypeClassifier()
        : {
          classify: () => ({ type: 'technical' }),
          getRequestTypeHint: () => ''
        });
    this.validator = options.validator || new ResponseValidator();
    this.gmailService = options.gmailService || new GmailService();
    this.promptEngine = options.promptEngine ||
      (typeof PromptEngine !== 'undefined'
        ? new PromptEngine()
        : { buildPrompt: (opts) => (opts && opts.emailContent) ? opts.emailContent : '' });
    this.memoryService = options.memoryService ||
      (typeof MemoryService !== 'undefined'
        ? new MemoryService()
        : {
          getMemory: () => ({}),
          updateMemoryAtomic: () => { },
          updateReaction: () => { }
        });
    // Integrazione TerritoryValidator rimossa
    // Configurazione
    this.config = {
      validationEnabled: typeof CONFIG !== 'undefined' ? CONFIG.VALIDATION_ENABLED : true,
      dryRun: typeof CONFIG !== 'undefined' ? CONFIG.DRY_RUN : false,
      maxEmailsPerRun: typeof CONFIG !== 'undefined' ? CONFIG.MAX_EMAILS_PER_RUN : 10,
      maxExecutionTimeMs: typeof CONFIG !== 'undefined' && CONFIG.MAX_EXECUTION_TIME_MS
        ? CONFIG.MAX_EXECUTION_TIME_MS
        : 280 * 1000,
      minRemainingTimeMs: typeof CONFIG !== 'undefined' && typeof CONFIG.MIN_REMAINING_TIME_MS === 'number'
        ? CONFIG.MIN_REMAINING_TIME_MS
        : 90 * 1000,
      labelName: typeof CONFIG !== 'undefined' ? CONFIG.LABEL_NAME : 'IA',
      errorLabelName: typeof CONFIG !== 'undefined' ? CONFIG.ERROR_LABEL_NAME : 'Errore',
      validationErrorLabel: typeof CONFIG !== 'undefined' ? CONFIG.VALIDATION_ERROR_LABEL : 'Verifica',
      validationWarningThreshold: typeof CONFIG !== 'undefined' && typeof CONFIG.VALIDATION_WARNING_THRESHOLD === 'number'
        ? CONFIG.VALIDATION_WARNING_THRESHOLD
        : 0.9,
      maxConsecutiveExternal: typeof CONFIG !== 'undefined' && typeof CONFIG.MAX_CONSECUTIVE_EXTERNAL === 'number'
        ? CONFIG.MAX_CONSECUTIVE_EXTERNAL
        : 5,
      emptyInboxWarningThreshold: typeof CONFIG !== 'undefined' && typeof CONFIG.EMPTY_INBOX_WARNING_THRESHOLD === 'number'
        ? CONFIG.EMPTY_INBOX_WARNING_THRESHOLD
        : 5,
      searchPageSize: typeof CONFIG !== 'undefined' && typeof CONFIG.SEARCH_PAGE_SIZE === 'number'
        ? CONFIG.SEARCH_PAGE_SIZE
        : 50
    };

    this.logger.info('EmailProcessor inizializzato', {
      validazione: this.config.validationEnabled,
      dryRun: this.config.dryRun
    });

    // Timestamp run corrente (usato da _isNearDeadline/_getRemainingTimeMs).
    // Viene poi resettato all'avvio di ogni batch in processUnreadEmails.
    this._startTime = Date.now();
  }

  /**
   * Classifica l'errore per determinare la strategia di retry.
   * Gli errori non retryable (chiave invalida, risposta malformata) sono fatali.
   * @param {Error} error
   * @returns {{ type: string, retryable: boolean, message: string }}
   */
  _classifyError(error) {
    const result = classifyError(error);
    if (!result.retryable) {
      result.type = 'FATAL';
    }
    return result;
  }

  /**
   * Elabora il singolo thread (analisi, categorizzazione, generazione risposta, invio)
   * @param {GmailThread} thread 
   * @param {string} knowledgeBase - KB testo semplice
   * @param {Set} labeledMessageIds - ID messaggi già etichettati (opzionale)
   * @param {boolean} skipLock - Se true, salta acquisizione lock
   */
  processThread(thread, knowledgeBase, labeledMessageIds = new Set(), skipLock = false) {
    const threadId = thread.getId();
    const startTime = Date.now();

    // ====================================================================================================
    // ACQUISIZIONE LOCK (LIVELLO-THREAD) - Previene condizioni di conflitto
    // ====================================================================================================

    // Stato lock per rilascio in finally
    let lockAcquired = false;
    const scriptCache = (typeof CacheService !== 'undefined' && CacheService && typeof CacheService.getScriptCache === 'function')
      ? CacheService.getScriptCache()
      : null;
    const threadLockKey = `thread_lock_${threadId}`;
    let lockValue = null;

    if (skipLock) {
      console.log(`🔒 Lock saltato per thread ${threadId} (chiamante ha già lock)`);
    } else if (!scriptCache || typeof LockService === 'undefined' || !LockService || typeof LockService.getScriptLock !== 'function') {
      console.warn(`⚠️ Lock service/cache non disponibili per thread ${threadId}: procedo senza lock`);
    } else {
      const configuredTtl = (typeof CONFIG !== 'undefined' && Number(CONFIG.CACHE_LOCK_TTL))
        ? Number(CONFIG.CACHE_LOCK_TTL)
        : 240; // Fallback allineato al default config per coprire OCR + validazione semantica

      const ttlSeconds = configuredTtl;
      const lockTtlMs = ttlSeconds * 1000;

      // Aggiunge entropia per evitare collisioni se due thread identici partono nello stesso Ms 
      const entropy = Math.random().toString(36).substring(2, 8);
      lockValue = `${Date.now()}_${entropy}`;

      const scriptLock = LockService.getScriptLock();
      try {
        if (!scriptLock.tryLock(5000)) {
          console.warn(`🔒 Impossibile acquisire lock globale per thread ${threadId}, salto`);
          return { status: 'skipped', reason: 'global_lock_unavailable' };
        }

        const existingLock = scriptCache.get(threadLockKey);
        if (existingLock) {
          const existingTimestamp = Number(String(existingLock).split('_')[0]);
          const isStale = !isNaN(existingTimestamp) && (Date.now() - existingTimestamp) > lockTtlMs;

          if (isStale) {
            console.warn(`🔓 Lock stale rilevato per thread ${threadId}, pulizia`);
            scriptCache.remove(threadLockKey);
          } else {
            console.warn(`🔒 Thread ${threadId} lockato da altro processo, salto`);
            return { status: 'skipped', reason: 'thread_locked' };
          }
        }

        scriptCache.put(threadLockKey, lockValue, ttlSeconds);
        const confirmValue = scriptCache.get(threadLockKey);
        if (confirmValue !== lockValue) {
          console.warn(`🔒 Collisione lock cache su thread ${threadId}, salto`);
          return { status: 'skipped', reason: 'thread_lock_collision' };
        }

        lockAcquired = true;
        console.log(`🔒 Lock acquisito per thread ${threadId}`);
      } catch (e) {
        console.warn(`⚠️ Errore acquisizione lock thread: ${e.message}`);
        return { status: 'error', error: 'Lock acquisition failed' };
      } finally {
        if (scriptLock && typeof scriptLock.releaseLock === 'function') {
          try {
            scriptLock.releaseLock();
          } catch (_) { }
        }
      }
    }

    const result = {
      status: 'unknown',
      validationFailed: false,
      dryRun: false,
      error: null
    };

    // Snapshot robusto del classificatore errori — restituisce sempre { type, retryable, message }
    // per coerenza con il contratto di gas_error_types.classifyError
    const classifyError = (typeof this._classifyError === 'function')
      ? this._classifyError.bind(this)
      // Manteniamo fallback locale per non dipendere dal globale in runtime modulari.
      : function fallbackClassifyError(err) { return { type: 'UNKNOWN', retryable: false, message: String(err) }; };

    let candidate = null;
    try {
      // Raccogli informazioni su thread e messaggi
      // Ottieni ultimo messaggio NON LETTO nel thread
      const messages = thread.getMessages();
      const unreadMessages = messages.filter(m => m.isUnread());

      // Recupero indirizzo email corrente con fallback robusto
      let myEmail = '';
      try {
        const effectiveUser = Session.getEffectiveUser();
        myEmail = effectiveUser ? effectiveUser.getEmail() : '';

        if (!myEmail) {
          const activeUser = Session.getActiveUser();
          myEmail = activeUser ? activeUser.getEmail() : '';
        }
      } catch (e) {
        console.warn(`⚠️ Impossibile recuperare email utente: ${e.message}`);
      }

      if (!myEmail) {
        const adminEmailProperty = (typeof PropertiesService !== 'undefined' && PropertiesService && typeof PropertiesService.getScriptProperties === 'function')
          ? PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')
          : '';
        const adminEmailConfig = (typeof CONFIG !== 'undefined' && CONFIG.LOGGING && CONFIG.LOGGING.ADMIN_EMAIL)
          ? CONFIG.LOGGING.ADMIN_EMAIL
          : '';
        const adminEmail = adminEmailProperty || adminEmailConfig || '';
        const botEmailProperty = (typeof PropertiesService !== 'undefined' && PropertiesService && typeof PropertiesService.getScriptProperties === 'function')
          ? PropertiesService.getScriptProperties().getProperty('BOT_EMAIL')
          : '';
        const botEmailConfig = (typeof CONFIG !== 'undefined' && CONFIG.BOT_EMAIL) ? CONFIG.BOT_EMAIL : '';

        myEmail = adminEmail || botEmailProperty || botEmailConfig || '';

        if (myEmail) {
          console.warn(`⚠️ Session email non disponibile: uso fallback anti-loop (${myEmail})`);
        } else {
          console.warn('⚠️ Session email non disponibile e nessun fallback configurato (ScriptProperties.ADMIN_EMAIL/CONFIG.LOGGING.ADMIN_EMAIL/BOT_EMAIL/CONFIG.BOT_EMAIL)');
        }
      }

      // ====================================================================================================
      // FILTRO A LIVELLO MESSAGGIO
      // ====================================================================================================
      const effectiveLabeledIds = (labeledMessageIds && labeledMessageIds.size > 0)
        ? labeledMessageIds
        : this.gmailService.getMessageIdsWithLabel(this.config.labelName);

      const unlabeledUnread = unreadMessages.filter(message => {
        return !effectiveLabeledIds.has(message.getId());
      });

      // Build set of our own addresses (primary + aliases) per filtro early-stage
      const ownAddresses = new Set();
      if (myEmail) ownAddresses.add(myEmail.toLowerCase());
      const knownAliasesArray = (typeof CONFIG !== 'undefined' && Array.isArray(CONFIG.KNOWN_ALIASES))
        ? CONFIG.KNOWN_ALIASES : [];
      knownAliasesArray.forEach(alias => {
        if (alias) ownAddresses.add(String(alias).toLowerCase());
      });

      const externalUnread = unlabeledUnread.filter(message => {
        // B07: Usa getFrom() leggero anziché la costosa extractMessageDetails()
        const rawFrom = (message.getFrom() || '');
        const senderEmail = this.gmailService._extractEmailAddress(rawFrom);

        // Se non riusciamo ad estrarre l'email, consideriamo il mittente come esterno per sicurezza
        if (!senderEmail) return true;
        return !ownAddresses.has(senderEmail.toLowerCase());
      });

      // GUARDRAIL (critico): se un messaggio è già stato etichettati IA, non deve
      // rientrare nel ciclo di risposta automatica anche se il thread è ancora aperto.
      // Questo evita doppie risposte su stesso messaggio.
      // Se non ci sono messaggi non letti non ancora etichettati → skip
      if (unlabeledUnread.length === 0) {
        console.log('   ⊖ Thread già elaborato (nessun nuovo messaggio non letto)');
        result.status = 'skipped';
        result.reason = 'already_labeled_no_new_unread';
        return result;
      }

      // GUARDRAIL (critico): rispondiamo solo a mittenti esterni.
      // I messaggi interni (noi/alias) vengono esclusi per evitare loop e risposte non dovute.
      // Se non ci sono messaggi da esterni → skip
      if (externalUnread.length === 0) {
        console.log('   ⊖ Saltato: nessun nuovo messaggio esterno non letto');
        unlabeledUnread.forEach(message => this._markMessageAsProcessed(message, labeledMessageIds));
        result.status = 'skipped';
        result.reason = 'no_external_unread';
        return result;
      }

      // Seleziona ultimo messaggio non letto non etichettato da esterni
      candidate = externalUnread[externalUnread.length - 1];
      const messageDetails = this.gmailService.extractMessageDetails(candidate);

      console.log(`\n📧 Elaborazione: ${(messageDetails.subject || '').substring(0, 50)}...`);
      console.log(`   Da: ${messageDetails.senderEmail} (${messageDetails.senderName})`);


      if (messageDetails.isNewsletter) {
        console.log('   ⊖ Saltato: rilevata newsletter (List-Unsubscribe/Precedence)');
        this._markMessageAsProcessed(candidate, labeledMessageIds);
        result.status = 'filtered';
        result.reason = 'newsletter_header';
        return result;
      }

      // ====================================================================================================
      // STEP 0: CONTROLLO ULTIMO MITTENTE (Anti-Loop & Ownership)
      // GUARDRAIL (critico): se l'ultimo intervento nel thread è nostro,
      // il bot NON deve inviare un'altra risposta. Si attende il prossimo messaggio utente.
      // ====================================================================================================
      const lastMessage = messages[messages.length - 1];
      const lastSenderRaw = lastMessage.getFrom() || '';
      const lastSenderEmail = (this.gmailService && typeof this.gmailService._extractEmailAddress === 'function')
        ? this.gmailService._extractEmailAddress(lastSenderRaw).toLowerCase()
        : '';
      const knownAliases = (typeof CONFIG !== 'undefined' && Array.isArray(CONFIG.KNOWN_ALIASES)) ? CONFIG.KNOWN_ALIASES : [];
      const normalizedMyEmail = myEmail ? myEmail.toLowerCase() : '';
      const lastSpeakerIsUs = Boolean(lastSenderEmail) && Boolean(normalizedMyEmail) && (
        lastSenderEmail === normalizedMyEmail ||
        knownAliases.some(alias => lastSenderEmail === alias.toLowerCase())
      );

      if (lastSpeakerIsUs) {
        console.log('   ⊖ Saltato: l\'ultimo messaggio del thread è già nostro (bot o segreteria)');
        // Non marchiamo nulla, semplicemente ci fermiamo finché l'utente non risponde
        result.status = 'skipped';
        result.reason = 'last_speaker_is_me';
        return result;
      }

      // ====================================================================================================
      // STEP 0.1: ANTI-AUTO-RISPOSTA (Safe Sender Check)
      // ====================================================================================================
      const safeSenderEmail = (messageDetails.senderEmail || '').toLowerCase();

      // Verifica mittente: confronta solo con email proprie e alias configurati.
      // Il confronto con i destinatari (To/CC) è stato rimosso perché genera
      // falsi positivi su mailing list e thread con CC multipli.
      const isMe = Boolean(safeSenderEmail) && (
        (Boolean(normalizedMyEmail) && safeSenderEmail === normalizedMyEmail) ||
        knownAliases.some(alias => safeSenderEmail === alias.toLowerCase())
      );

      if (isMe) {
        console.log('   ⊖ Saltato: messaggio auto-inviato (o da alias conosciuto)');
        this._markMessageAsProcessed(candidate, labeledMessageIds);
        result.status = 'skipped';
        result.reason = 'self_sent';
        return result;
      }

      // ====================================================================================================
      // STEP 0.2: AUTO-REPLY / OUT-OF-OFFICE DETECTION
      // ====================================================================================================
      const headers = messageDetails.headers || {};
      // Lookup case-insensitive: i server SMTP possono restituire header in casing arbitrario
      const getHeader = (name) => {
        const lower = name.toLowerCase();
        for (const key in headers) {
          if (key.toLowerCase() === lower) return headers[key];
        }
        return '';
      };
      const autoSubmitted = getHeader('auto-submitted');
      const precedence = getHeader('precedence');
      const xAutoReply = getHeader('x-autoreply');
      const xAutoResponseSuppress = getHeader('x-auto-response-suppress');

      if (
        /auto-replied|auto-generated/i.test(autoSubmitted) ||
        /bulk|auto_reply/i.test(precedence) ||
        /auto-reply|autoreply/i.test(xAutoReply) ||
        /oof|all|dr|rn|nri|auto/i.test(xAutoResponseSuppress)
      ) {
        console.log('   ⊖ Saltato: risposta automatica (header SMTP)');
        this._markMessageAsProcessed(candidate, labeledMessageIds);
        result.status = 'filtered';
        result.reason = 'out_of_office';
        return result;
      }

      const outOfOfficePatterns = [
        /\b(out of office|away from office|fuori ufficio|assente)\b/i,
        /\b(automatic reply|risposta automatica)\b/i,
        /\breturn(ing)? on\b/i,
        /\bdi ritorno (il|dal)\b/i,
        /\b(thank you for your message|mailbox monitored periodically|messaggio ricevuto)\b/i
      ];

      const oooSubject = messageDetails.subject || '';
      const oooBody = messageDetails.body || '';
      if (outOfOfficePatterns.some(p => p.test(`${oooSubject} ${oooBody}`))) {
        console.log('   ⊖ Saltato: risposta automatica out-of-office (testo)');
        this._markMessageAsProcessed(candidate, labeledMessageIds);
        result.status = 'filtered';
        result.reason = 'out_of_office';
        return result;
      }

      const candidateIndex = messages.findIndex(msg => msg.getId() === candidate.getId());
      if (candidateIndex > 0 && messages[candidateIndex - 1]) {
        const previousMessage = messages[candidateIndex - 1];
        const previousSender = (previousMessage.getFrom() || '').toLowerCase();
        const candidateDate = messageDetails.date ? messageDetails.date.getTime() : null;
        const previousDate = previousMessage.getDate() ? previousMessage.getDate().getTime() : null;
        const arrivedSoonAfterUs = candidateDate && previousDate
          ? Math.abs(candidateDate - previousDate) <= 10 * 60 * 1000
          : false;
        const previousIsUs = myEmail ? previousSender.includes(myEmail.toLowerCase()) : false;
        const candidateBody = messageDetails.body || '';
        const candidateWords = candidateBody.trim().split(/\s+/).filter(Boolean);
        const isShortClosureReply = candidateWords.length > 0 && candidateWords.length < 10 &&
          /\b(grazie|ok|perfetto)\b/i.test(candidateBody);

        if (previousIsUs && arrivedSoonAfterUs && isShortClosureReply) {
          console.log('   ⊖ Saltato: risposta breve di chiusura (grazie/ok/perfetto)');
          this._markMessageAsProcessed(candidate, labeledMessageIds);
          result.status = 'filtered';
          result.reason = 'short_closure_reply';
          return result;
        }
      }

      // ====================================================================================================
      // STEP 0.5: ANTI-LOOP (rilevamento intelligente)
      // ====================================================================================================
      const MAX_THREAD_LENGTH = (typeof CONFIG !== 'undefined' && CONFIG.MAX_THREAD_LENGTH) ? CONFIG.MAX_THREAD_LENGTH : 10;
      const MAX_CONSECUTIVE_EXTERNAL = this.config.maxConsecutiveExternal;

      if (messages.length > MAX_THREAD_LENGTH) {
        if (!myEmail) {
          console.warn('   ⚠️ Email utente non disponibile: skip controllo anti-loop basato su mittente');
        }
        let consecutiveExternal = 0;
        const knownAliases = (typeof CONFIG !== 'undefined' && Array.isArray(CONFIG.KNOWN_ALIASES))
          ? CONFIG.KNOWN_ALIASES.map(alias => String(alias || '').toLowerCase())
          : [];
        const normalizedMyEmail = myEmail ? myEmail.toLowerCase() : '';

        for (let i = messages.length - 1; i >= 0; i--) {
          const msgFrom = messages[i].getFrom().toLowerCase();
          const isUs = Boolean(normalizedMyEmail) && (
            msgFrom.includes(normalizedMyEmail) ||
            knownAliases.some(alias => alias && msgFrom.includes(alias))
          );

          if (!isUs) {
            consecutiveExternal++;
          } else {
            break;
          }
        }

        if (normalizedMyEmail && consecutiveExternal >= MAX_CONSECUTIVE_EXTERNAL) {
          console.log(`   ⊖ Saltato: probabile loop email (${consecutiveExternal} esterni consecutivi)`);
          this._markMessageAsProcessed(candidate, labeledMessageIds);
          result.status = 'filtered';
          result.reason = 'email_loop_detected';
          return result;
        }

        console.warn(`   ⚠️ Thread lungo (${messages.length} messaggi) ma non loop - elaboro`);
      }

      // ====================================================================================================
      // STEP 0.8: ANTI-MITTENTE-NOREPLY
      // ====================================================================================================
      const senderInfo = `${messageDetails.senderEmail} ${messageDetails.senderName}`.toLowerCase();
      const autoPattern = /no-reply|do-not-reply|noreply|daemon|postmaster|bounce|mailer/i;
      if (autoPattern.test(senderInfo)) {
        console.log('   ⊖ Saltato: mittente rilevato come casella automatica o no-reply');
        this._markMessageAsProcessed(candidate, labeledMessageIds);
        result.status = 'filtered';
        result.reason = 'no_reply_sender';
        return result;
      }

      // ====================================================================================================
      // STEP 1: FILTRO - Domini/parole chiave ignorati
      // ====================================================================================================
      if (this._shouldIgnoreEmail(messageDetails)) {
        console.log('   ⊖ Filtrato: domain/keyword ignore');
        this._markMessageAsProcessed(candidate, labeledMessageIds);
        result.status = 'filtered';
        return result;
      }

      // ====================================================================================================
      // STEP 2: CLASSIFICAZIONE - Filtro ack/greeting ultra-semplice
      // ====================================================================================================
      const MAX_SUBJECT_LENGTH = 1000;
      const safeSubject = (messageDetails.subject || '').substring(0, MAX_SUBJECT_LENGTH);
      const safeBody = (messageDetails.body || '');
      const safeSubjectLower = safeSubject.toLowerCase();

      const classification = this.classifier.classifyEmail(
        safeSubject,
        safeBody,
        safeSubjectLower.startsWith('re:')
      );

      if (!classification.shouldReply) {
        // NUOVO: Se ci sono allegati, ignoriamo il filtro "no reply" del classificatore locale
        // per permettere a Gemini Vision di analizzare il contenuto degli allegati.
        if (this._shouldTryOcr(safeBody, safeSubject, candidate)) {
          console.log(`   📎 Classificatore locale: shouldReply=false ma rilevati allegati, proseguo per analisi Vision.`);
        } else {
          console.log(`   ⊖ Filtrato dal classifier: ${classification.reason}`);
          this._markMessageAsProcessed(candidate, labeledMessageIds);
          result.status = 'filtered';
          return result;
        }
      }

      // ====================================================================================================
      // STEP 3: CONTROLLO RAPIDO - Gemini decide se serve risposta
      // ====================================================================================================
      const quickCheck = this.geminiService.shouldRespondToEmail(
        messageDetails.body,
        messageDetails.subject
      );

      if (!quickCheck.shouldRespond) {
        // NUOVO: Se ci sono allegati, ignoriamo il "no respond" di Gemini Quick Check
        // perché Gemini non ha ancora visto il contenuto visivo degli allegati.
        if (this._shouldTryOcr(messageDetails.body, messageDetails.subject, candidate)) {
          console.log(`   📎 Gemini quick check negativo, ma rilevati allegati: procedo con analisi Vision`);
        } else {
          console.log(`   ⊖ Gemini quick check: nessuna risposta necessaria (${quickCheck.reason})`);
          this._markMessageAsProcessed(candidate, labeledMessageIds);
          result.status = 'filtered';
          return result;
        }
      }

      const languageDetection = this.geminiService.detectEmailLanguage(
        messageDetails.body,
        messageDetails.subject
      ) || {};
      const detectedLanguage = (quickCheck.language || languageDetection.lang || 'it').toLowerCase();
      console.log(`   🌐 Lingua: ${detectedLanguage.toUpperCase()}`);

      // ====================================================================================================
      // STEP 4: CLASSIFICAZIONE TIPO RICHIESTA (Multi-dimensionale)
      // ====================================================================================================
      const requestType = this.requestClassifier.classify(
        messageDetails.subject,
        messageDetails.body,
        quickCheck.classification
      );

      // Estrai dati dalla nuova struttura classificazione
      const categoryMatch = (messageDetails.body || '').match(/"category"\s*:\s*"(TECHNICAL|PASTORAL|POETRY|FORMAL|MIXED)"/i);
      const isPastoral = requestType.dimensions ? (requestType.dimensions.pastoral > 0.6) : (requestType.type === 'pastoral'); // Compatibilità

      // ====================================================================================================
      // STEP 5: KB ENRICHMENT CONDIZIONALE
      // ====================================================================================================
      const knowledgeSections = [];

      // KB Principale (Istruzioni dal foglio)
      knowledgeSections.push('--- ISTRUZIONI OPERATIVE ---\\n' + knowledgeBase);

      const enrichedKnowledgeBase = knowledgeSections.filter(Boolean).join('\n\n');

      // ====================================================================================================
      // STEP 6: STORICO CONVERSAZIONE
      // ====================================================================================================
      let conversationHistory = '';
      if (messages.length > 1) {
        const candidateId = candidate.getId();
        const historyMessages = messages.filter(m => m.getId() !== candidateId);

        if (historyMessages.length > 0) {
          conversationHistory = this.gmailService.buildConversationHistory(
            historyMessages,
            10,
            myEmail
          );
        }
      }

      // ====================================================================================================
      // STEP 6.5: CONTESTO MEMORIA
      // ====================================================================================================
      const memoryContext = this.memoryService.getMemory(threadId) || {};

      if (Object.keys(memoryContext).length > 0) {
        console.log(`   🧠 Memoria trovata: lang=${memoryContext.language}, topics=${(memoryContext.providedInfo || []).length}`);
      }

      // ====================================================================================================
      // STEP 6.6: CALCOLO DINAMICO SALUTO E RITARDO
      // ====================================================================================================
      const memoryMessageCount = Number.isFinite(Number(memoryContext.messageCount))
        ? Number(memoryContext.messageCount)
        : 0;

      const salutationMode = computeSalutationMode({
        isReply: safeSubjectLower.startsWith('re:'),
        messageCount: memoryMessageCount,
        memoryExists: Object.keys(memoryContext).length > 0,
        lastUpdated: memoryContext.lastUpdated || null,
        now: new Date()
      });
      console.log(`   📊 Modalità saluto: ${salutationMode}`);

      const responseDelay = computeResponseDelay({
        messageDate: messageDetails.date,
        now: new Date()
      });
      if (responseDelay.shouldApologize) {
        console.log(`   🕐 Ritardo risposta: ${responseDelay.days} giorni`);
      }

      // ====================================================================================================
      // STEP 7: COSTRUISCI PROMPT
      // ====================================================================================================
      let { greeting, closing } = this.geminiService.getAdaptiveGreeting(
        messageDetails.senderName,
        detectedLanguage
      );

      // Override strutturale: nessun saluto in conversazioni attive
      if (salutationMode === 'none_or_continuity' || salutationMode === 'session') {
        greeting = '';
      } else if (salutationMode === 'soft') {
        greeting = '[Inizia con breve frase di riaggancio cordiale]';
      }

      // ====================================================================================================
      // VERIFICA TERRITORIO RIMOSSA
      // ====================================================================================================


      // ====================================================================================================
      // STEP 7.2: PROMPT CONTEXT (profilo e concern dinamici)
      // ====================================================================================================
      let promptProfile = 'standard';
      let activeConcerns = [];
      if (typeof createPromptContext === 'function') {
        const promptContext = createPromptContext({
          email: {
            subject: safeSubject,
            body: messageDetails.body,
            isReply: safeSubjectLower.startsWith('re:'),
            detectedLanguage: detectedLanguage
          },
          classification: {
            category: classification.category,
            subIntents: classification.subIntents || {},
            confidence: classification.confidence || 0.8
          },
          requestType: requestType,
          memory: {
            exists: Object.keys(memoryContext).length > 0,
            providedInfoCount: (memoryContext.providedInfo || []).length,
            lastUpdated: memoryContext.lastUpdated || null
          },
          conversation: { messageCount: memoryMessageCount },
          knowledgeBase: { length: enrichedKnowledgeBase.length, containsDates: /\d{4}/.test(enrichedKnowledgeBase) },
          temporal: {
            mentionsDates: this._detectTemporalMentions(messageDetails.body, detectedLanguage) || /\b\d{1,2}\/\d{1,2}\b/.test(messageDetails.body),
            mentionsTimes: /\d{1,2}[:.]\d{2}/.test(messageDetails.body)
          },
          salutationMode: salutationMode
        });
        promptProfile = promptContext.profile;
        activeConcerns = promptContext.concerns;
        console.log(`   🧠 PromptContext: profilo=${promptProfile}`);
      }

      const categoryHintSource = classification.category || requestType.type;

      // ====================================================================================================
      // STEP 7.1: ESTRAZIONE ALLEGATI MULTIMODALI (Gemini Vision)
      // ====================================================================================================
      let attachmentBlobs = [];
      let textFromAttachments = '';
      let attachmentSkipped = [];

      if (typeof CONFIG !== 'undefined' && CONFIG.ATTACHMENT_CONTEXT && CONFIG.ATTACHMENT_CONTEXT.enabled) {
        if (this._isNearDeadline(this.config.maxExecutionTimeMs)) {
          attachmentSkipped.push({ reason: 'near_deadline' });
          console.warn('   ⏳ Elaborazione allegati saltata: tempo residuo insufficiente.');
        } else {
          let hasAttachments = false;
          try {
            const attachments = candidate.getAttachments({ includeInlineImages: true, includeAttachments: true }) || [];
            hasAttachments = attachments.length > 0;
          } catch (e) {
            console.warn(`⚠️ Impossibile leggere allegati per pre-check: ${e.message}`);
          }

          if (!hasAttachments) {
            attachmentSkipped.push({ reason: 'no_attachments' });
            console.log('   📎 Elaborazione allegati saltata: nessun allegato nel messaggio candidato');
          } else if (this._shouldTryOcr(messageDetails.body, messageDetails.subject, candidate)) {
            console.log('   📎 Elaborazione allegati multimodale (Vision)...');
            const attachResult = this.gmailService.getProcessableAttachments(candidate);
            attachmentBlobs = attachResult.blobs || [];
            textFromAttachments = attachResult.textContext || '';
            attachmentSkipped = attachResult.skipped || [];
          } else {
            attachmentSkipped.push({ reason: 'precheck_no_ocr' });
            console.log('   📎 Allegati saltati: pre-check negativo');
          }
        }
      }
      if (attachmentBlobs.length > 0) {
        console.log(`   📎 ${attachmentBlobs.length} allegati visivi pronti per Gemini Vision`);
      }
      if (attachmentSkipped.length > 0) {
        const skippedNames = attachmentSkipped.map(item => item.name || item.reason).join(', ');
        console.log(`   📎 Allegati saltati: ${attachmentSkipped.length} (${skippedNames})`);
      }

      const promptOptions = {
        emailContent: messageDetails.body,
        emailSubject: messageDetails.subject,
        knowledgeBase: enrichedKnowledgeBase,
        senderName: messageDetails.senderName,
        senderEmail: messageDetails.senderEmail,
        conversationHistory: conversationHistory,
        category: requestType.type === 'parish_info_request' ? 'parish_info_request' : (classification.category || requestType.type),
        topic: quickCheck.classification ? quickCheck.classification.topic : '',
        detectedLanguage: detectedLanguage,
        currentSeason: this._getCurrentSeason(),
        currentDate: new Date().toISOString().split('T')[0],
        salutation: greeting,
        closing: closing,
        subIntents: classification.subIntents || {},
        memoryContext: memoryContext,
        salutationMode: salutationMode,
        responseDelay: responseDelay,
        promptProfile: promptProfile,
        activeConcerns: activeConcerns,

        requestType: requestType, // Aggiunto per recupero selettivo Dottrina
        attachmentsContext: textFromAttachments
      };

      const prompt = this.promptEngine.buildPrompt(promptOptions);

      // Aggiungi hint tipo richiesta (nuovo metodo blended)
      const typeHint = this.requestClassifier.getRequestTypeHint(requestType);
      const fullPrompt = typeHint + '\n\n' + prompt;

      // ====================================================================================================
      // STEP 8: GENERA RISPOSTA (STRATEGIA "CROSS-KEY QUALITY FIRST")
      // ====================================================================================================
      // NOTA ARCHITETTURALE:
      // Questa fase può richiedere più tempo del normale (fino a 4 tentativi API).
      // SCELTA DELIBERATA: Privilegiamo la qualità della risposta (Modello Flash 2.5)
      // rispetto alla velocità. 
      // 1. Proviamo Flash 2.5 sulla chiave primaria.
      // 2. Se fallisce, proviamo Flash 2.5 sulla chiave di RISERVA.
      // 3. Solo se entrambe falliscono, degradiamo al modello Lite (più economico).
      // Questo "costo" in termini di tempo è gestito riducendo MAX_EMAILS_PER_RUN a 3.
      // ====================================================================================================

      let response = null;
      let generationError = null;
      let initialError = null;
      let strategyUsed = null;

      if (this._isNearDeadline(this.config.maxExecutionTimeMs)) {
        console.warn('⏳ Tempo residuo insufficiente prima della generazione AI: rimando il thread al prossimo turno.');
        result.status = 'skipped';
        result.reason = 'near_deadline_before_generation';
        return result;
      }

      // Punto 12: Utilizzo del metodo di classe centralizzato per la classificazione degli errori


      // Definizione strategie di generazione (Punti di robustezza cross-key)
      const geminiModels = (typeof CONFIG !== 'undefined' && CONFIG.GEMINI_MODELS) ? CONFIG.GEMINI_MODELS : {};
      const flashModel = (geminiModels['flash-2.5'] && geminiModels['flash-2.5'].name) ? geminiModels['flash-2.5'].name : 'gemini-2.5-flash';
      const liteModel = (geminiModels['flash-lite'] && geminiModels['flash-lite'].name) ? geminiModels['flash-lite'].name : 'gemini-2.5-flash-lite';

      const attemptStrategy = [
        { name: 'Primary-Flash2.5', key: this.geminiService.primaryKey, model: flashModel, skipRateLimit: false },
        { name: 'Backup-Flash2.5', key: this.geminiService.backupKey, model: flashModel, skipRateLimit: true },
        { name: 'Fallback-Lite', key: this.geminiService.backupKey || this.geminiService.primaryKey, model: liteModel, skipRateLimit: true }
      ];

      // Esecuzione Loop Strategico
      for (const plan of attemptStrategy) {
        // Salta se manca la chiave (es. backupKey non configurata)
        if (!plan.key) continue;

        try {
          console.log(`🔄 Tentativo Generazione: ${plan.name}...`);

          response = this.geminiService.generateResponse(fullPrompt, {
            apiKey: plan.key,
            modelName: plan.model,
            skipRateLimit: plan.skipRateLimit,
            attachments: attachmentBlobs
          });

          // GeminiService può restituire stringa semplice oppure
          // oggetto strutturato { success, text, modelUsed }.
          if (response && typeof response === 'object') {
            response = response.text;
          }

          if (response) {
            strategyUsed = plan.name;
            console.log(`✅ Generazione riuscita con strategia: ${plan.name}`);
            break; // Successo! Esci dal loop
          }

        } catch (err) {
          generationError = err; // Salva l'ultimo errore
          if (!initialError) initialError = err; // Salva il primo errore per il reporting
          const errorClass = this._classifyError(err);
          console.warn(`⚠️ Strategia '${plan.name}' fallita: ${err.message} [${errorClass.type}]`);

          if (errorClass.type === 'FATAL') {
            console.error('🛑 Errore fatale rilevato, interrompo strategia.');
            break;
          }

          if (errorClass.type === 'NETWORK') {
            console.warn('🌐 Errore di rete/timeout, continuo con prossima strategia.');
            continue;
          }
          // QUOTA_EXCEEDED, INVALID_RESPONSE, UNKNOWN: continua
        }
      }

      // Verifiche finali post-loop
      if (!response) {
        const errorToReport = initialError || generationError;
        const errorClass = errorToReport ? this._classifyError(errorToReport) : { type: 'UNKNOWN', retryable: false, message: 'Generation strategies exhausted' };
        console.error('🛑 TUTTE le strategie di generazione sono fallite.');
        this._addErrorLabel(thread);
        this._markMessageAsProcessed(candidate, labeledMessageIds);
        result.status = 'error';
        result.error = errorToReport ? errorToReport.message : 'Generation strategies exhausted';
        if (initialError && generationError && initialError !== generationError) {
          result.error += ` (Ultimo fallback: ${generationError.message})`;
        }
        result.errorClass = errorClass.type;
        return result;
      }

      // Controlla marcatore NO_REPLY
      if (typeof response !== 'string') {
        console.error(`🛑 Risposta non valida da Gemini: tipo ricevuto '${typeof response}'`);
        this._addErrorLabel(thread);
        this._markMessageAsProcessed(candidate, labeledMessageIds);
        result.status = 'error';
        result.error = 'Invalid response type from GeminiService';
        result.errorClass = 'DATA';
        return result;
      }

      if (response.trim() === 'NO_REPLY') {
        console.log('   ⊖ AI ha restituito NO_REPLY');
        this._markMessageAsProcessed(candidate, labeledMessageIds);
        result.status = 'filtered';
        return result;
      }

      // Nota: con il sistema multimodale Vision, Gemini legge direttamente gli allegati.
      // Non serve più la nota OCR bassa confidenza.

      // ====================================================================================================
      // STEP 9: VALIDA RISPOSTA
      // ====================================================================================================
      if (this.config.validationEnabled) {
        const validation = this.validator.validateResponse(
          response,
          detectedLanguage,
          enrichedKnowledgeBase,
          messageDetails.body,
          messageDetails.subject,
          salutationMode
        );

        if (!validation.isValid) {
          console.warn(`   🛑 Validazione FALLITA (punteggio: ${validation.score.toFixed(2)})`);

          // Gestione errore validazione critica
          if (validation.details && validation.details.exposedReasoning && validation.details.exposedReasoning.score === 0.0) {
            console.warn("⚠️ Risposta bloccata per Thinking Leak. Invio a etichetta 'Verifica'.");
            result.reason = 'thinking_leak';
          }

          this._addValidationErrorLabel(thread);
          this._markMessageAsProcessed(candidate, labeledMessageIds);
          result.status = 'validation_failed';
          result.validationFailed = true;
          // Garantisce che ogni rifiuto abbia un motivo tracciabile
          if (!result.reason) {
            result.reason = 'validation_score_below_threshold';
          }
          return result;
        }

        // Se ci sono WARNING e il punteggio è sotto la soglia di sicurezza, aggiungi etichetta "verifica"
        // Ignoriamo i warning per punteggi alti (es. >= 0.90) assumendo siano nits minori (es. firma)
        const warningThreshold = this.config.validationWarningThreshold || 0.90;

        if (validation.warnings && validation.warnings.length > 0 && validation.score < warningThreshold) {
          console.log(
            `   ⚠️ Validazione: Punteggio ${validation.score.toFixed(2)} < ${warningThreshold} con warning - Aggiungo etichetta '${this.config.validationErrorLabel}'`
          );
          this.gmailService.addLabelToMessage(candidate.getId(), this.config.validationErrorLabel);
        } else if (validation.warnings && validation.warnings.length > 0) {
          console.log(`   ℹ️ Validazione: Punteggio alto (${validation.score.toFixed(2)}). Warning ignorati: ${validation.warnings.join(', ')}`);
        }

        if (validation.fixedResponse) {
          console.log('   🩹 Usa risposta corretta automaticamente (Self-Healing)');
          response = validation.fixedResponse;
        }

        console.log(`   ✓ Validazione PASSATA (punteggio: ${validation.score.toFixed(2)})`);
      }

      // ====================================================================================================
      // STEP 10: INVIA RISPOSTA
      // ====================================================================================================
      if (this.config.dryRun) {
        console.log('   🔴 DRY RUN - Risposta non inviata');
        console.log(`   📄 Invierebbe: ${response.substring(0, 100)}...`);
        result.dryRun = true;
        // In DRY_RUN non aggiorniamo memoria né label per non avere effetti permanenti
        result.status = 'replied';
        result.durationMs = Date.now() - startTime;
        this.logger.info(`Thread processato in ${result.durationMs}ms`, { threadId: threadId, duration: result.durationMs });
        return result;
      }

      try {
        this.gmailService.sendHtmlReply(candidate, response, messageDetails);
      } catch (e) {
        const errorMessage = e && e.message ? e.message : String(e);
        console.error(`   🛑 Errore invio Gmail: ${errorMessage}`);
        this._addErrorLabel(thread);
        if (candidate) {
          this._markMessageAsProcessed(candidate, labeledMessageIds);
        }
        result.status = 'error';
        result.error = `gmail_send_failed: ${errorMessage}`;
        return result;
      }

      // ====================================================================================================
      // STEP 11: AGGIORNA MEMORIA (solo se non DRY_RUN)
      // ====================================================================================================
      const providedTopics = this._detectProvidedTopics(response);

      // Strutturazione Oggetti Topic
      const topicsWithObjects = providedTopics.map(topic => ({
        topic: topic,
        userReaction: 'unknown',
        context: null,
        timestamp: new Date().toISOString()
      }));

      const memorySummary = this._buildMemorySummary({
        existingSummary: memoryContext.memorySummary || '',
        responseText: response,
        providedTopics: providedTopics
      });

      // Inferisci reazione utente su topic precedenti (se presenti)
      if (memoryContext.providedInfo && memoryContext.providedInfo.length > 0) {
        this._inferUserReaction(messageDetails.body, memoryContext.providedInfo, threadId);
      }

      const memoryUpdate = {
        language: detectedLanguage,
        category: classification.category || requestType.type
      };

      if (memorySummary) {
        memoryUpdate.memorySummary = memorySummary;
      }

      this.memoryService.updateMemoryAtomic(threadId, memoryUpdate, topicsWithObjects.length > 0 ? topicsWithObjects : null);

      if (candidate) {
        this._markMessageAsProcessed(candidate, labeledMessageIds);
      }
      result.status = 'replied';
      result.durationMs = Date.now() - startTime;
      this.logger.info(`Thread processato in ${result.durationMs}ms`, { threadId: threadId, duration: result.durationMs });
      return result;

    } catch (error) {
      console.error(`   🛑 Errore elaborazione thread: ${error.message}`);
      this._addErrorLabel(thread);
      if (candidate) {
        this._markMessageAsProcessed(candidate, labeledMessageIds);
      }
      result.status = 'error';
      result.error = error.message;
      return result;

    } finally {
      // Nota manutenzione: questo finally copre TUTTI i return anticipati nel try,
      // quindi non duplicare scriptCache.remove(...) in ogni ramo di skip.
      // Rilascia lock (solo se acquisito)
      if (lockAcquired && scriptCache && threadLockKey) {
        try {
          const currentLockValue = scriptCache.get(threadLockKey);
          // Rilasciamo sia quando il lock è nostro sia quando è già scaduto.
          if (!currentLockValue || currentLockValue === lockValue) {
            scriptCache.remove(threadLockKey);
            console.log(`🔓 Lock rilasciato per thread ${threadId}`);
          } else {
            console.warn(`⚠️ Rilascio lock saltato per thread ${threadId} (lock scaduto o di altro processo)`);
          }
        } catch (e) {
          // Non forziamo remove() se get() fallisce: potremmo cancellare il lock
          // legittimo di un altro worker appena subentrato dopo la scadenza TTL.
          console.warn('⚠️ Errore rilascio lock:', e.message);
        }
      }
    }
  }

  /**
   * Processa tutte le email non lette
   * @param {string} knowledgeBase
   * @param {boolean} skipExecutionLock - Evita il double-locking se chiamato da main()
   */
  processUnreadEmails(knowledgeBase, skipExecutionLock = true) {
    console.log('\n' + '='.repeat(70));
    console.log('📬 Inizio elaborazione email...');
    console.log('='.repeat(70));

    const executionLock = LockService.getScriptLock();
    let lockAcquiredHere = false;

    // Nota: non rimuoviamo questo blocco. Serve quando processUnreadEmails è invocata
    // fuori dall'entry point principale (job manuali/test), evitando corse concorrenti.
    // In main il parametro skipExecutionLock=true impedisce il double-lock/release prematuro.
    if (!skipExecutionLock) {
      const lockWaitMs = (typeof CONFIG !== 'undefined' && CONFIG.EXECUTION_LOCK_WAIT_MS)
        ? CONFIG.EXECUTION_LOCK_WAIT_MS : 5000;

      if (!executionLock.tryLock(lockWaitMs)) {
        console.warn('⚠️ Un\'altra esecuzione è già attiva: salto questo turno per evitare doppie risposte.');
        return { total: 0, replied: 0, filtered: 0, errors: 0, skipped: 1, reason: 'execution_locked' };
      }
      lockAcquiredHere = true;
    }

    try {
      if (knowledgeBase === null || typeof knowledgeBase === 'undefined') {
        this.logger.error('Knowledge base non disponibile: interrompo batch per evitare risposte senza contesto.');
        return { total: 0, replied: 0, filtered: 0, errors: 1, skipped: 0, reason: 'knowledge_base_missing' };
      }

      if (this.config.dryRun) {
        console.warn('🔴 MODALITÀ DRY_RUN ATTIVA - Email NON inviate!');
      }

      // Cerca thread non letti nella inbox
      // Utilizziamo un buffer di ricerca più ampio per gestire thread saltati (es. loop interni)
      // Rimuoviamo il filtro etichetta per permettere la gestione dei follow-up in thread già elaborati
      const processedLabelQuery = this._formatLabelQueryValue(this.config.labelName);
      const errorLabelQuery = this._formatLabelQueryValue(this.config.errorLabelName);
      const validationLabelQuery = this._formatLabelQueryValue(this.config.validationErrorLabel);

      // La query esclude i thread già etichettati (il prefisso "-" è obbligatorio)
      const searchQuery = `is:unread -label:${processedLabelQuery} -label:${errorLabelQuery} -label:${validationLabelQuery} in:inbox`;
      const searchLimit = (this.config.searchPageSize || 50);

      const threads = GmailApp.search(
        searchQuery,
        0,
        searchLimit
      );

      if (threads.length === 0) {
        const emptyStreak = this._trackEmptyInboxStreak(true);
        console.log(`Nessuna email da elaborare (query: ${searchQuery}).`);

        if (emptyStreak >= this.config.emptyInboxWarningThreshold) {
          console.warn(`⚠️ Inbox vuota da ${emptyStreak} esecuzioni consecutive. Verificare filtri Gmail/trigger in ingresso.`);
        }

        return { total: 0, replied: 0, filtered: 0, errors: 0, emptyStreak: emptyStreak };
      }

      this._trackEmptyInboxStreak(false);
      console.log(`📬 Trovate ${threads.length} email da elaborare (query: ${searchQuery})`);

      // Carica etichette una sola volta
      let labeledMessageIds;
      try {
        labeledMessageIds = this.gmailService.getMessageIdsWithLabel(this.config.labelName, true, { throwOnError: true });
      } catch (e) {
        console.warn(`⚠️ Impossibile caricare le etichette IA: ${e.message}. Per evitare doppie risposte, salto il batch.`);
        return {
          total: 0,
          replied: 0,
          filtered: 0,
          validationFailed: 0,
          errors: 1,
          dryRun: 0,
          skipped: 1,
          skipped_locked: 0,
          skipped_processed: 0,
          skipped_internal: 0,
          skipped_loop: 0,
          reason: 'label_lookup_failed'
        };
      }
      console.log(`📦 Trovati in cache ${labeledMessageIds.size} messaggi già elaborati`);

      // Statistiche
      const stats = {
        total: 0,
        replied: 0,
        filtered: 0,
        validationFailed: 0,
        errors: 0,
        dryRun: 0,
        skipped: 0,
        skipped_locked: 0,
        skipped_processed: 0,
        skipped_internal: 0,
        skipped_loop: 0
      };

      // Processa ogni thread fino a raggiungere il limite di elaborazione
      this._startTime = Date.now();
      const MAX_EXECUTION_TIME = this.config.maxExecutionTimeMs;
      let processedCount = 0; // Contatore thread effettivamente elaborati
      const maxLimit = parseInt(this.config.maxEmailsPerRun, 10);
      const safeLimit = Number.isNaN(maxLimit) ? 10 : maxLimit;

      for (let index = 0; index < threads.length; index++) {
        // Stop se abbiamo raggiunto il target di elaborazione effettiva
        if (processedCount >= safeLimit) {
          console.log(`🛑 Raggiunti ${safeLimit} thread elaborati. Stop.`);
          break;
        }

        const thread = threads[index];

        // Controllo tempo residuo (rigoroso): evita avvio di un nuovo thread senza buffer minimo
        const remainingTimeMs = this._getRemainingTimeMs(MAX_EXECUTION_TIME);
        if (remainingTimeMs < this.config.minRemainingTimeMs || this._isNearDeadline(MAX_EXECUTION_TIME)) {
          console.warn(`⏳ Tempo insufficiente per un nuovo thread (${Math.round(remainingTimeMs / 1000)}s restanti). Stop preventivo.`);
          break;
        }

        // Fast-skip: evita pipeline completa/lock quando i non letti sono già etichettati.
        // Riduce overhead su thread marcati manualmente come "non letto" ma già gestiti.
        if (!this._hasUnreadMessagesToProcess(thread, labeledMessageIds)) {
          console.log(`\n--- Thread ${index + 1}/${threads.length} ---`);
          console.log('   ⊖ Fast-skip: thread con soli non letti già etichettati IA');
          stats.total++;
          stats.skipped++;
          stats.skipped_processed++;
          continue;
        }

        console.log(`\n--- Thread ${index + 1}/${threads.length} ---`);

        const result = this.processThread(thread, knowledgeBase, labeledMessageIds);
        stats.total++;

        // Incrementa contatore solo se c'è stata un'azione significativa o decisione esplicita dell'AI
        const isEffectiveWork = (
          result.status === 'replied' ||
          result.status === 'error' ||
          result.status === 'validation_failed' ||
          result.status === 'filtered'
        );

        if (isEffectiveWork) {
          processedCount++;
        }

        if (result.validationFailed) {
          stats.validationFailed++;
        } else if (result.status === 'replied') {
          stats.replied++;
          if (result.dryRun) stats.dryRun++;
        } else if (result.status === 'skipped') {
          stats.skipped++;
          if (result.reason === 'thread_locked' || result.reason === 'thread_locked_race') stats.skipped_locked++;
          if (result.reason === 'already_labeled_no_new_unread') stats.skipped_processed++;
          if (result.reason === 'no_external_unread' || result.reason === 'self_sent') stats.skipped_internal++;
          if (result.reason === 'email_loop_detected') stats.skipped_loop++;
        } else if (result.status === 'filtered') {
          stats.filtered++;
          if (result.reason === 'email_loop_detected') stats.skipped_loop++;
        } else if (result.status === 'error') {
          stats.errors++;
        }
      }

      // Stampa riepilogo
      console.log('\n' + '='.repeat(70));
      console.log('📊 RIEPILOGO ELABORAZIONE');
      console.log('='.repeat(70));
      console.log(`   Totale analizzate (buffer): ${stats.total}`);
      console.log(`   ✓ Risposte inviate: ${stats.replied}`);
      if (stats.dryRun > 0) console.warn(`   🔴 DRY RUN: ${stats.dryRun}`);

      if (stats.skipped > 0) {
        console.log(`   ⊖ Saltate (Totale): ${stats.skipped}`);
      }

      console.log(`   ⊖ Filtrate (AI/Regole): ${stats.filtered}`);
      if (stats.validationFailed > 0) console.warn(`   🛑 Validazione fallita: ${stats.validationFailed}`);
      if (stats.errors > 0) console.error(`   🛑 Errori: ${stats.errors}`);
      stats.processed = processedCount;
      console.log('='.repeat(70));

      return stats;

    } finally {
      if (lockAcquiredHere) {
        try {
          executionLock.releaseLock();
        } catch (e) {
          console.warn(`⚠️ Errore rilascio execution lock: ${e.message}`);
        }
      }
    }
  }

  // ====================================================================================================
  // RILEVAMENTO TEMPORALE (Date/Orari)
  // ====================================================================================================

  /**
   * Verifica se l'email deve essere ignorata (blacklist, auto-reply, notifiche)
   * Usa le liste UNIFICATE (Codice + Foglio) presenti in GLOBAL_CACHE
   */
  _shouldIgnoreEmail(messageDetails) {
    const email = (messageDetails.senderEmail || '').toLowerCase();
    const subject = (messageDetails.subject || '').toLowerCase();
    const body = (messageDetails.body || '').toLowerCase();

    // 1. Controllo Blacklist Domini/Email
    // NOTA: GLOBAL_CACHE.ignoreDomains include già CONFIG.IGNORE_DOMAINS (merge in _loadAdvancedConfig)
    const ignoreDomains = (typeof GLOBAL_CACHE !== 'undefined' && Array.isArray(GLOBAL_CACHE.ignoreDomains))
      ? GLOBAL_CACHE.ignoreDomains.map(d => String(d).toLowerCase())
      : ((typeof CONFIG !== 'undefined' && CONFIG.IGNORE_DOMAINS) ? CONFIG.IGNORE_DOMAINS.map(d => String(d).toLowerCase()) : []);

    if (ignoreDomains.some(domain => {
      const isExactMatch = email === domain;
      const isDomainMatch = email.endsWith(domain.startsWith('@') ? domain : '@' + domain);
      return isExactMatch || isDomainMatch;
    })) {
      console.log(`🚫 Ignorato: mittente in blacklist (${email})`);
      return true;
    }

    // 2. Controllo Keyword Oggetto/Corpo
    // NOTA: GLOBAL_CACHE.ignoreKeywords include già CONFIG.IGNORE_KEYWORDS (merge in _loadAdvancedConfig)
    const ignoreKeywords = (typeof GLOBAL_CACHE !== 'undefined' && Array.isArray(GLOBAL_CACHE.ignoreKeywords))
      ? GLOBAL_CACHE.ignoreKeywords.map(k => String(k).toLowerCase())
      : ((typeof CONFIG !== 'undefined' && CONFIG.IGNORE_KEYWORDS) ? CONFIG.IGNORE_KEYWORDS.map(k => String(k).toLowerCase()) : []);

    if (ignoreKeywords.some(keyword => subject.includes(keyword) || body.includes(keyword))) {
      console.log(`🚫 Ignorato: oggetto o corpo contiene keyword vietata`);
      return true;
    }

    // 3. Controllo Auto-reply e Notifiche (Standard)
    // NOTA: no-reply/noreply sono anche controllati in STEP 0.8 (defense-in-depth)
    if (
      email.includes('no-reply') ||
      email.includes('noreply') ||
      email.includes('mailer-daemon') ||
      email.includes('postmaster') ||
      email.includes('notification') ||
      email.includes('alert') ||
      subject.includes('delivery status notification') ||
      subject.includes('automatic reply') ||
      subject.includes('fuori sede') ||
      subject.includes('out of office') ||
      body.includes('this is an automatically generated message') ||
      body.includes('do not reply to this email')
    ) {
      console.log('🚫 Ignorato: auto-reply o notifica di sistema');
      return true;
    }

    return false;
  }

  _shouldTryOcr(body, subject, message = null) {
    if (!message) return false;
    try {
      const attachments = message.getAttachments({ includeInlineImages: true, includeAttachments: true }) || [];
      if (attachments.length > 0) {
        console.log('   📎 OCR attivo: allegati trovati per potenziale analisi poetica');
        return true;
      }
    } catch (e) {
      console.warn(`⚠️ Impossibile verificare allegati per OCR: ${e.message}`);
    }
    return false;
  }

  _getCurrentSeason() {
    const month = new Date().getMonth() + 1;
    return (month >= 6 && month <= 9) ? 'estivo' : 'invernale';
  }


  _markMessageAsProcessed(message, labeledMessageIds = null) {
    // SCELTA OPERATIVA INTENZIONALE:
    // - etichetta IA a livello *messaggio* (non thread), usando Gmail Advanced Service;
    // - NON marcare come letto qui.
    // Motivo: il segretario deve vedere a colpo d'occhio i non letti, anche se già gestiti da IA.
    // Cambiare questo comportamento altera la triage operativa.
    const messageId = message.getId();
    this.gmailService.addLabelToMessage(messageId, this.config.labelName);
    if (labeledMessageIds && typeof labeledMessageIds.add === 'function') {
      labeledMessageIds.add(messageId);
    }
  }

  // Calcola se il tempo residuo è sufficiente per elaborare un nuovo thread
  _isNearDeadline(maxExecutionTimeMs) {
    const budgetMs = Number(maxExecutionTimeMs) || 330000;
    const minRemainingMs = (typeof this.config.minRemainingTimeMs === 'number')
      ? this.config.minRemainingTimeMs
      : 90000; // 90 secondi margine sicurezza

    if (!this._startTime) return false;

    const elapsed = Date.now() - this._startTime;
    return elapsed > Math.max(0, budgetMs - minRemainingMs);
  }

  _getRemainingTimeMs(maxExecutionTimeMs) {
    const budgetMs = Number(maxExecutionTimeMs) || 330000;
    const start = Number(this._startTime) || Date.now();
    const elapsed = Date.now() - start;
    return Math.max(0, budgetMs - elapsed);
  }

  _formatLabelQueryValue(labelName) {
    const raw = String(labelName || '').trim();
    if (!raw) return '""';
    // Scelta intenzionale: in query Gmail il quoting è la strategia più stabile.
    // Evitiamo escaping aggressivo di caratteri non-quote perché può alterare il matching label:nome.
    return `"${raw.replace(/"/g, '\\"')}"`;
  }

  _hasUnreadMessagesToProcess(thread, labeledMessageIds) {
    try {
      const messages = thread.getMessages() || [];
      const unreadMessages = messages.filter(m => m.isUnread());

      // Nessun non letto: non c'è lavoro da fare.
      if (unreadMessages.length === 0) {
        return false;
      }

      const effectiveLabeledIds = (labeledMessageIds && labeledMessageIds.size > 0)
        ? labeledMessageIds
        : this.gmailService.getMessageIdsWithLabel(this.config.labelName);

      return unreadMessages.some(message => !effectiveLabeledIds.has(message.getId()));
    } catch (e) {
      // Fallback sicuro: in caso di errore non bloccare il thread, lasciamo decidere a processThread.
      this.logger.warn(`⚠️ Fast-skip check fallito: ${e.message}`);
      return true;
    }
  }

  _addErrorLabel(thread) {
    this.gmailService.addLabelToThread(thread, this.config.errorLabelName);
  }

  _addValidationErrorLabel(thread) {
    this.gmailService.addLabelToThread(thread, this.config.validationErrorLabel);
  }

  // Costruisce un sommario incrementale delle risposte inviate al thread
  _buildMemorySummary({ existingSummary, responseText, providedTopics }) {
    const maxChars = 2000;
    const maxBullets = 5;

    let summaryLines = [];
    if (existingSummary && typeof existingSummary === 'string') {
      summaryLines = existingSummary
        .split('\n')
        .map(line => line.trim())
        .filter(Boolean);
    }

    if (!responseText) {
      return summaryLines.slice(-maxBullets).join('\n') || null;
    }

    const plainText = responseText
      .replace(/<[^>]*>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    const sentenceMatches = plainText.match(/[^.!?]+[.!?]+|[^.!?]+$/g) || [];
    const ignorePatterns = [
      /^ciao\b/i,
      /^buongiorno/i,
      /^buonasera/i,
      /^gentile/i,
      /^salve\b/i,
      /^grazie\b/i,
      /^cordiali saluti/i,
      /^saluti\b/i
    ];

    const candidateSentences = sentenceMatches
      .map(sentence => sentence.trim())
      .filter(sentence => sentence.length > 20)
      .filter(sentence => !ignorePatterns.some(pattern => pattern.test(sentence)));

    let summarySentence = candidateSentences.slice(0, 2).join(' ');

    if (!summarySentence && providedTopics && providedTopics.length > 0) {
      summarySentence = `Risposta con informazioni su: ${providedTopics.join(', ')}.`;
    }
    if (!summarySentence) {
      summarySentence = plainText.slice(0, 200);
    }

    // Aggiungiamo data contestuale per non collassare topic se esauditi in momenti diversi
    const newBullet = summarySentence ? `• [${new Date().toISOString().split('T')[0]}] ${summarySentence}` : '';

    // Controlliamo l'overlap ma con tolleranza (data aiuta l'unicità)
    if (newBullet && !summaryLines.some(line => line.toLowerCase() === newBullet.toLowerCase())) {
      summaryLines.push(newBullet);
    }

    const trimmedLines = summaryLines.slice(-maxBullets);
    let summary = trimmedLines.join('\n').trim();

    if (summary.length > maxChars) {
      const truncated = summary.slice(0, maxChars);
      const lastBreak = truncated.lastIndexOf('\n');
      const lastSpace = truncated.lastIndexOf(' ');
      const cutIndex = lastBreak > 0 ? lastBreak : (lastSpace > 0 ? lastSpace : maxChars);
      summary = truncated.slice(0, cutIndex).trim() + '...';
    }

    return summary || null;
  }

  /**
   * Rileva topic forniti nella risposta (per anti-ripetizione memoria)
   */
  _detectProvidedTopics(response) {
    if (!response || typeof response !== 'string') return [];
    const topics = [];
    // I pattern usano già il flag /i, nessuna normalizzazione necessaria

    const patterns = {
      'orari_messe': /messe?\b.*\d{1,2}[:.]\d{2}|orari\w*\s+messe/i,
      'contatti': /telefono|email|@|segreteria/i,
      'battesimo_info': /battesimo.*documento|documento.*battesimo/i,
      'comunione_info': /comunione.*catechismo|catechismo.*comunione/i,
      'cresima_info': /cresima.*percorso|percorso.*cresima/i,
      'matrimonio_info': /matrimonio.*corso|corso.*matrimonio/i,
      'territorio': /rientra|non rientra|parrocchia.*competenza/i,
      'indirizzo': /(?:via|viale|corso|piazza|largo|circonvallazione)\s+(?:[a-zA-ZàèéìòùÀÈÉÌÒÙ']+\s+){0,10}[a-zA-ZàèéìòùÀÈÉÌÒÙ']+\s*,?\s*\d+/i
    };

    for (const [topic, pattern] of Object.entries(patterns)) {
      if (pattern.test(response)) {
        topics.push(topic);
      }
    }

    return topics;
  }
  /**
   * Inferisce la reazione dell'utente rispetto ai topic forniti in precedenza
   */
  _inferUserReaction(userBody, previousTopics, threadId) {
    if (!previousTopics || previousTopics.length === 0) return;
    if (!userBody || typeof userBody !== 'string') return;

    const bodyLower = userBody.toLowerCase();

    // Pattern semplici di reazione
    const patterns = {
      questioned: [
        'non ho capito', 'non capisco', 'mi scusi non ho capito', 'non mi è chiaro',
        'non è chiaro', 'può chiarire', 'potrebbe chiarire', 'potrebbe spiegare',
        'cosa significa', 'dubbio', 'confuso', 'mi aiuta a capire',
        'i did not understand', 'i don\'t understand', 'not clear',
        'could you clarify', 'could you please clarify', 'could you explain',
        'no entiendo', 'no entendí', 'no me queda claro', 'podría aclarar',
        'podría explicar', 'podría ayudarme a entender'
      ],
      acknowledged: [
        'ho capito', 'tutto chiaro', 'grazie per la spiegazione', 'ok grazie',
        'perfetto', 'chiarissimo', 'ricevuto', 'la ringrazio', 'grazie',
        'gentilissimi', 'va benissimo', 'compreso',
        'thank you', 'thanks', 'understood', 'all clear', 'received',
        'gracias', 'entendido', 'entendida', 'recibido', 'recibida', 'perfecto', 'clarísimo'
      ],
      needs_expansion: [
        'potrebbe aggiungere', 'potrebbe fornire maggiori dettagli', 'maggiori dettagli',
        'più dettagli', 'approfondire', 'potrebbe spiegare meglio', 'potrebbe ampliare',
        'sarebbe possibile avere più informazioni', 'servirebbero più informazioni',
        'potrebbe indicare i passaggi',
        'could you provide more details', 'more details', 'could you elaborate',
        'would it be possible to have more information', 'could you outline the steps',
        'podría ampliar', 'más detalles', 'podría proporcionar más información',
        'sería posible tener más información', 'podría indicar los pasos'
      ]
    };

    const matchedQuestioned = patterns.questioned.find(p => bodyLower.includes(p));
    const matchedAcknowledged = patterns.acknowledged.find(p => bodyLower.includes(p));
    const matchedExpansion = patterns.needs_expansion.find(p => bodyLower.includes(p));

    let inferredReaction = null;
    if (matchedQuestioned) {
      inferredReaction = { type: 'questioned', match: matchedQuestioned };
    } else if (matchedExpansion) {
      inferredReaction = { type: 'needs_expansion', match: matchedExpansion };
    } else if (matchedAcknowledged) {
      inferredReaction = { type: 'acknowledged', match: matchedAcknowledged };
    }

    if (!inferredReaction) return;

    // 1. Trova TUTTI i topic menzionati esplicitamente
    const normalizedTopics = previousTopics
      .map(info => (typeof info === 'object' ? info.topic : info))
      .map(topic => this._normalizeTopicKey(topic));

    const mentionedTopics = normalizedTopics.filter(topic => {
      if (!topic) return false;
      const topicWithSpaces = topic.replace(/_/g, ' ');
      return bodyLower.includes(topic) || bodyLower.includes(topicWithSpaces);
    });

    let targetTopics = [];

    if (mentionedTopics.length > 0) {
      // Se l'utente cita esplicitamente dei topic, applica a tutti quelli trovati
      targetTopics = mentionedTopics;
    } else {
      // Fallback: applica all'ultimo topic discusso
      targetTopics = [normalizedTopics[normalizedTopics.length - 1]].filter(Boolean);
    }

    if (targetTopics.length === 0) return;

    const context = {
      source: 'user_reply',
      matchedPhrase: inferredReaction.match,
      excerpt: userBody.substring(0, 160)
    };

    targetTopics.forEach(topic => {
      console.log(`🧠 Inferred Reaction: ${inferredReaction.type.toUpperCase()} su topic '${topic}'`);
      this.memoryService.updateReaction(threadId, topic, inferredReaction.type, context);
    });
  }

  /**
   * Normalizza una chiave topic (lowercase, trim, spaces to underscore)
   */
  _normalizeTopicKey(topic) {
    if (topic == null) return '';
    return String(topic)
      .toLowerCase()
      .trim()
      .replace(/[_\-]+/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/ /g, '_');
  }
  /**
   * Punto 12: Classificazione centralizzata degli errori API
   * Determina se un errore è fatale, legato alla quota o alla rete.
   */
  _classifyError(error) {
    const mkResult = (type, retryable, message) => ({ type, retryable, message });

    if (!error) {
      console.warn('⚠️ _classifyError chiamato con errore nullo');
      return mkResult('UNKNOWN', false, '');
    }

    // Delega al classificatore centralizzato se disponibile
    if (typeof classifyError === 'function' && typeof ErrorTypes !== 'undefined') {
      const normalized = classifyError(error);
      switch (normalized.type) {
        case ErrorTypes.QUOTA_EXCEEDED:
          return mkResult('QUOTA', true, normalized.message);
        case ErrorTypes.TIMEOUT:
        case ErrorTypes.NETWORK:
          return mkResult('NETWORK', true, normalized.message);
        case ErrorTypes.INVALID_API_KEY:
        case ErrorTypes.INVALID_RESPONSE:
          return mkResult('FATAL', false, normalized.message);
        default:
          return mkResult('UNKNOWN', false, normalized.message);
      }
    }

    // Classificazione locale (standalone) — stesso contratto { type, retryable, message }
    const msg = String(error.message || error.toString() || '');
    if (!msg) {
      console.warn('⚠️ Errore senza messaggio utile:', error);
      return mkResult('UNKNOWN', false, '');
    }

    const RETRYABLE_ERRORS = ['429', 'rate limit', 'quota', 'RESOURCE_EXHAUSTED'];
    const FATAL_ERRORS = ['INVALID_ARGUMENT', 'PERMISSION_DENIED', 'UNAUTHENTICATED'];

    for (const fatal of FATAL_ERRORS) {
      if (msg.includes(fatal)) return mkResult('FATAL', false, msg);
    }

    for (const retryable of RETRYABLE_ERRORS) {
      if (msg.includes(retryable)) return mkResult('QUOTA', true, msg);
    }

    if (msg.includes('timeout') || msg.includes('ECONNRESET') || msg.includes('503') || msg.includes('500') || msg.includes('502') || msg.includes('504')) {
      return mkResult('NETWORK', true, msg);
    }

    return mkResult('UNKNOWN', false, msg);
  }

  /**
   * Traccia il contatore di inbox vuote consecutive (per avvisi diagnostici)
   */
  _trackEmptyInboxStreak(isEmpty) {
    const cache = CacheService.getScriptCache();
    const key = 'empty_inbox_streak';
    let streak = parseInt(cache.get(key) || '0');

    if (isEmpty) {
      streak++;
      cache.put(key, streak.toString(), 21600); // 6 ore
    } else {
      streak = 0;
      cache.remove(key);
    }
    return streak;
  }

  _detectTemporalMentions(text, language) {
    // Punto 15: Protezione contro input nulli o non validi
    if (!text || typeof text !== 'string') return false;
    const monthPatterns = {
      'it': /\b(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)\b/i,
      'en': /\b(january|february|march|april|may|june|july|august|september|october|november|december)\b/i,
      'es': /\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\b/i,
      'pt': /\b(janeiro|fevereiro|mar\u00E7o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\b/i
    };

    // Fallback su italiano se lingua non supportata
    const pattern = monthPatterns[language] || monthPatterns['it'];
    return pattern.test(text);
  }
}

// ====================================================================================================
// CALCOLATORE MODALITÀ SALUTO
// ====================================================================================================

/**
 * Calcola modalità saluto basata su segnali strutturali
 * @param {Object} params - Parametri di input
 * @returns {'full'|'soft'|'none_or_continuity'|'session'}
 */ // Fix: added 'session' to contract
function computeSalutationMode({ isReply = false, messageCount = 0, memoryExists = false, lastUpdated = null, now = new Date() } = {}) {
  const SESSION_WINDOW_MINUTES = 15;
  // 0️⃣ Nuovo contatto (non reply): privilegia sempre un saluto completo.
  // Anche in presenza di memoria pregressa, un nuovo thread/messaggio iniziale
  // deve evitare modalità "none_or_continuity".
  if (!isReply) {
    return 'full';
  }

  // 1️⃣ Primo messaggio assoluto
  if (!memoryExists && messageCount <= 1) {
    return 'full';
  }

  // 2️⃣ Conversazione attiva (qui isReply è necessariamente true)
  if (!lastUpdated) {
    return 'none_or_continuity';
  }

  const parsedLastUpdated = new Date(lastUpdated);
  if (isNaN(parsedLastUpdated.getTime())) {
    return 'none_or_continuity';
  }

  const timeSinceLastMs = now.getTime() - parsedLastUpdated.getTime();
  const minutesSinceLast = timeSinceLastMs / (1000 * 60);
  const hoursSinceLast = timeSinceLastMs / (1000 * 60 * 60);

  if (isNaN(hoursSinceLast) || hoursSinceLast < 0) {
    console.warn('⚠️ Timestamp futuro o invalido');
    return 'full';
  }

  // Sessione conversazionale ravvicinata (entro 15 minuti)
  if (minutesSinceLast <= SESSION_WINDOW_MINUTES) {
    return 'session';
  }

  // Follow-up ravvicinato (entro 48h)
  if (hoursSinceLast <= 48) {
    return 'none_or_continuity';
  }

  // Conversazione ripresa dopo pausa (48h - 4 giorni)
  if (hoursSinceLast <= 96) {
    return 'soft';
  }

  // Troppo tempo passato (> 4 giorni) → nuovo contatto
  return 'full';
}

// Compatibilità: rende la funzione disponibile anche in runtime che usano moduli/isolamento
if (typeof globalThis !== 'undefined' && typeof globalThis.computeSalutationMode !== 'function') {
  globalThis.computeSalutationMode = computeSalutationMode;
}

// Funzione factory
function createEmailProcessor(options) {
  return new EmailProcessor(options);
}

/**
 * Calcola ritardo risposta rispetto alla data del messaggio
 * @param {Object} params
 * @returns {{shouldApologize: boolean, hours: number, days: number}}
 */
function computeResponseDelay({ messageDate, now = new Date(), thresholdHours = 72 }) {
  if (!messageDate) {
    return { shouldApologize: false, hours: 0, days: 0 };
  }

  const parsedMessageDate = new Date(messageDate);
  const diffMs = now.getTime() - parsedMessageDate.getTime();

  if (isNaN(diffMs) || diffMs < 0) {
    return { shouldApologize: false, hours: 0, days: 0 };
  }

  const hours = diffMs / (1000 * 60 * 60);
  const days = Math.floor(hours / 24);

  return {
    shouldApologize: hours >= thresholdHours,
    hours: Math.round(hours),
    days: days
  };
}

// ====================================================================================================
// ENTRY POINT PRINCIPALE
// ====================================================================================================

/**
 * Alias dell'entry point principale processEmailsMain() (gas_main.js).
 * Mantenuta per compatibilità con trigger preesistenti.
 */
function processUnreadEmailsMain() {
  if (typeof processEmailsMain === 'function') {
    processEmailsMain();
  } else {
    console.error('🛑 processEmailsMain non trovata — impossibile delegare.');
  }
}
