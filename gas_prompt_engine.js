/**
 * PromptEngine.gs - Generazione prompt modulare
 * 19 classi template per composizione prompt
 * Supporta filtro dinamico basato su profilo
 * 
 * Include:
 * - Recupero selettivo Dottrina
 * - Checklist contestuale
 * - Ottimizzazione struttura prompt
 */

class PromptEngine {
  constructor() {
    // Logger strutturato
    this.logger = createLogger('PromptEngine');
    this.logger.info('Inizializzazione PromptEngine con recupero selettivo');

    // Configurazione filtering template per profilo
    this.LITE_SKIP_TEMPLATES = [
      'ExamplesTemplate',
      'FormattingGuidelinesTemplate',
      'HumanToneGuidelinesTemplate',
      'SpecialCasesTemplate'
    ];

    this.STANDARD_SKIP_TEMPLATES = [
      'ExamplesTemplate'
    ];

    this.logger.info('PromptEngine inizializzato', { templates: 20 });
  }

  /**
   * Stima token (approx 4 char/token per l'italiano/inglese)
   */
  estimateTokens(text) {
    if (!text || typeof text !== 'string') return 0;
    return Math.ceil(text.length / 4);
  }

  /**
   * Normalizza valori eterogenei in stringa sicura per il prompt.
   * Evita output "[object Object]" quando una risorsa viene passata in forma non-stringa.
   */
  _normalizePromptTextInput(value, fallback = '') {
    if (value == null) return fallback;
    if (typeof value === 'string') return value;

    try {
      const serialized = JSON.stringify(value);
      return typeof serialized === 'string' ? serialized : String(value);
    } catch (e) {
      return String(value);
    }
  }

  /**
   * Determina se un template deve essere incluso in base a profilo e concern
   */
  _shouldIncludeTemplate(templateName, promptProfile, activeConcerns = {}) {
    if (promptProfile === 'heavy') {
      return true; // Profilo heavy include tutto
    }

    if (promptProfile === 'lite') {
      if (this.LITE_SKIP_TEMPLATES.includes(templateName)) {
        return false;
      }
    }

    if (promptProfile === 'standard') {
      if (this.STANDARD_SKIP_TEMPLATES.includes(templateName)) {
        // Salta esempi a meno che formatting_risk non sia attivo
        if (!activeConcerns.formatting_risk) {
          return false;
        }
      }
    }

    return true;
  }

  /**
  * Costruisce il prompt completo dal contesto
  * Supporta filtro dinamico template basato su profilo
  * 
  * ORDINE SEZIONI:
  * 1. Setup critico (Ruolo, Lingua, KB, Territorio) - Priorità alta
  * 2. Contesto (Memoria, Cronologia, Email)
  * 3. Linee guida (Formattazione, Tono, Esempi)
  * 4. Rinforzo finale (Errori critici, Checklist)
  */
  buildPrompt(options) {
    const {
      emailContent,
      emailSubject,
      knowledgeBase,
      senderName = 'Utente',
      senderEmail = '',
      conversationHistory = '',
      category = null,
      topic = '',
      detectedLanguage = 'it',
      currentSeason = 'invernale',
      currentDate = Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd'),
      salutation = 'Buongiorno.',
      closing = 'Cordiali saluti,',
      subIntents = {},
      memoryContext = {},
      promptProfile = 'heavy',
      activeConcerns = {},
      salutationMode = 'full',
      responseDelay = null,
      attachmentsContext = ''
    } = options;

    let sections = [];
    let skippedCount = 0;

    // ══════════════════════════════════════════════════════
    // PRE-STIMA E BUDGETING TOKEN (Protezione Miglioramento #6 - Memory Growth)
    // ══════════════════════════════════════════════════════
    const MAX_SAFE_TOKENS = typeof CONFIG !== 'undefined' && CONFIG.MAX_SAFE_TOKENS
      ? CONFIG.MAX_SAFE_TOKENS : 50000;

    const OVERHEAD_TOKENS = (typeof CONFIG !== 'undefined' && CONFIG.PROMPT_ENGINE && CONFIG.PROMPT_ENGINE.OVERHEAD_TOKENS)
      ? CONFIG.PROMPT_ENGINE.OVERHEAD_TOKENS : 15000; // Riserva per istruzioni e sistema
    const KB_BUDGET_RATIO = (typeof CONFIG !== 'undefined' && typeof CONFIG.KB_TOKEN_BUDGET_RATIO === 'number')
      ? CONFIG.KB_TOKEN_BUDGET_RATIO
      : 0.5; // La KB può occupare max il 50% dello spazio rimanente
    const availableForKB = Math.max(0, (MAX_SAFE_TOKENS - OVERHEAD_TOKENS) * KB_BUDGET_RATIO);
    // Stima conservativa 1 token ≈ 4 caratteri: volutamente prudente per evitare overflow
    // con input multilingua/rumorosi. Ridurre il fattore aumenterebbe il rischio di prompt troppo lunghi.
    const kbCharsLimit = Math.round(availableForKB * 4);

    let workingKnowledgeBase = this._normalizePromptTextInput(knowledgeBase, '');
    let kbWasTruncated = false;

    // Troncamento proattivo della KB PRIMA di assemblare il prompt
    // ⚠️ Scelta blindata: questo è l'UNICO punto dove la KB può essere ridotta.
    // La cache risorse deve restare completa; qui applichiamo solo una riduzione runtime
    // per rispettare il budget token quando il contesto del singolo messaggio è eccezionalmente grande.
    if (workingKnowledgeBase && workingKnowledgeBase.length > kbCharsLimit) {
      console.warn(`⚠️ KB eccede il budget (${workingKnowledgeBase.length} chars), tronco a ${kbCharsLimit}`);
      // _truncateKbSemantically è implementato in questa classe: preserva paragrafi completi
      // invece di fare uno slice cieco che può spezzare contesto e istruzioni operative.
      workingKnowledgeBase = this._truncateKbSemantically(workingKnowledgeBase, kbCharsLimit);
      kbWasTruncated = true;
    }

    let workingAttachmentsContext = this._normalizePromptTextInput(attachmentsContext, '');
    if (kbWasTruncated && workingAttachmentsContext) {
      const attachmentSettings = (typeof CONFIG !== 'undefined' && CONFIG.ATTACHMENT_CONTEXT)
        ? CONFIG.ATTACHMENT_CONTEXT
        : {};
      const attachmentLimit = attachmentSettings.maxCharsWhenKbTruncated || 2000;
      if (workingAttachmentsContext.length > attachmentLimit) {
        console.warn(`⚠️ KB troncata: riduco allegati da ${workingAttachmentsContext.length} a ${attachmentLimit} chars`);
        workingAttachmentsContext = workingAttachmentsContext.slice(0, Math.max(0, attachmentLimit - 1)).trim() + '…';
      }
    }

    let usedTokens = 0;

    /**
     * Helper per aggiungere sezioni tracciando il budget token
     */
    const addSection = (section, label, options = {}) => {
      if (!section) return;
      const sectionTokens = this.estimateTokens(section);

      // Se superiamo il budget, saltiamo a meno che non sia forzato (es. istruzioni critiche)
      if (!options.force && usedTokens + sectionTokens > MAX_SAFE_TOKENS) {
        console.warn(`⚠️ Budget esaurito, sezione saltata: ${label}`);
        skippedCount++;
        return;
      }

      // Protezione memoria: Limita numero massimo sezioni
      if (sections.length >= 30) {
        console.warn(`⚠️ Limite sezioni raggiunto (30), salto sezione non critica: ${label}`);
        skippedCount++;
        return;
      }

      sections.push(section);
      usedTokens += sectionTokens;
    };

    /**
     * Helper per aggiungere template condizionali
     */
    const addTemplate = (templateName, content, label) => {
      if (this._shouldIncludeTemplate(templateName, promptProfile, activeConcerns)) {
        addSection(content, label || templateName);
      } else {
        skippedCount++;
      }
    };

    // ══════════════════════════════════════════════════════
    // BLOCCO 1: SETUP CRITICO (Priorità Massima)
    // ══════════════════════════════════════════════════════

    // 1. RUOLO SISTEMA
    addSection(this._renderSystemRole(), 'SystemRole', { force: true });

    // 2. ISTRUZIONI LINGUA
    addSection(this._renderLanguageInstruction(detectedLanguage), 'LanguageInstruction', { force: true });

    // 3. KNOWLEDGE BASE (Già troncata se necessario)
    addSection(this._renderKnowledgeBase(workingKnowledgeBase), 'KnowledgeBase');

    // 4. VERIFICA TERRITORIO RIMOSSA
    // ══════════════════════════════════════════════════════
    // BLOCCO 2: CONTESTO E CONTINUITÀ
    // ══════════════════════════════════════════════════════

    // 5. CONTESTO MEMORIA
    addSection(this._renderMemoryContext(memoryContext), 'MemoryContext');

    // 6. CONTINUITÀ CONVERSAZIONALE
    addSection(this._renderConversationContinuity(salutationMode), 'ConversationContinuity');

    // 7. SCUSE PER RITARDO
    addSection(this._renderResponseDelay(responseDelay, detectedLanguage), 'ResponseDelay');

    // 8. FOCUS UMANO (Condizionale)
    const shouldAddContinuityFocus =
      (memoryContext && Object.keys(memoryContext).length > 0) ||
      (salutationMode && salutationMode !== 'full') ||
      activeConcerns.emotional_sensitivity ||
      activeConcerns.repetition_risk;
    if (shouldAddContinuityFocus) {
      addSection(this._renderContinuityHumanFocus(), 'ContinuityHumanFocus');
    }

    // 9. CONTESTO STAGIONALE E TEMPORALE
    addSection(this._renderSeasonalContext(currentSeason), 'SeasonalContext');
    addSection(this._renderTemporalAwareness(currentDate, detectedLanguage), 'TemporalAwareness');

    // 10. SUGGERIMENTO CATEGORIA
    addSection(this._renderCategoryHint(category), 'CategoryHint');

    // ══════════════════════════════════════════════════════
    // BLOCCO 2b: ISTRUZIONE ANALISI POETICA
    // ══════════════════════════════════════════════════════
    const isPoetry = category === 'poem_submission' ||
      (category && category.toString().toUpperCase() === 'POETRY') ||
      workingAttachmentsContext;

    if (isPoetry) {
      addSection(this._renderPoetryAnalysisInstruction(), 'PoetryAnalysisInstruction');
    }

    // 12. CRONOLOGIA E CONTENUTO EMAIL
    if (conversationHistory) {
      addSection(this._renderConversationHistory(conversationHistory), 'ConversationHistory');
    }
    addSection(this._renderEmailContent(emailContent, emailSubject, senderName, senderEmail, detectedLanguage), 'EmailContent');
    addSection(this._renderAttachmentContext(workingAttachmentsContext), 'AttachmentsContext');

    // ══════════════════════════════════════════════════════
    // BLOCCO 3: LINEE GUIDA E TEMPLATE
    // ══════════════════════════════════════════════════════

    // 13. REGOLE NO REPLY
    addSection(this._renderNoReplyRules(), 'NoReplyRules');

    // 14. LINEE GUIDA (Filtrabili per profilo)
    addTemplate('FormattingGuidelinesTemplate', this._renderFormattingGuidelines(), 'FormattingGuidelines');

    // 15. STRUTTURA RISPOSTA
    addSection(this._renderResponseStructure(category, subIntents), 'ResponseStructure');

    // 16. TEMPLATE SPECIALI
    const normalizedTopic = (topic || '').toLowerCase();

    addTemplate('HumanToneGuidelinesTemplate', this._renderHumanToneGuidelines(), 'HumanToneGuidelines');
    addTemplate('ExamplesTemplate', this._renderExamples(category), 'Examples');

    // 17. REGOLE FINALI
    addSection(this._renderResponseGuidelines(detectedLanguage, currentSeason, salutation, closing, category, emailContent, workingAttachmentsContext, memoryContext), 'ResponseGuidelines');

    if (category !== 'formal') {
      addTemplate('SpecialCasesTemplate', this._renderSpecialCases(), 'SpecialCases');
    }

    // ══════════════════════════════════════════════════════
    // BLOCCO 4: RINFORZO FINALE
    // ══════════════════════════════════════════════════════

    addSection(this._renderCriticalErrorsReminder(), 'CriticalErrorsReminder');
    // Nota: il parametro è volutamente territoryContext (senza refusi) perché viene passato dal chiamante con lo stesso nome.
    addSection(this._renderContextualChecklist(detectedLanguage, salutationMode), 'ContextualChecklist');

    addSection('**Genera la risposta completa seguendo le linee guida sopra:**', 'FinalInstruction', { force: true });

    // Componi prompt finale tramite concatenazione efficiente
    const prompt = sections.join('\n\n');
    const finalTokens = this.estimateTokens(prompt);

    console.log(`📝 Prompt generato: ${prompt.length} caratteri (~${finalTokens} token) | Profilo: ${promptProfile} | Saltati: ${skippedCount}`);

    return prompt;
  }

  // ========================================================================
  // TEMPLATE 1: ERRORI CRITICI REMINDER (VERSIONE CONDENSATA)
  // ========================================================================
  // Una sola volta nel prompt

  _renderCriticalErrorsReminder() {
    return `
🚨 REMINDER ERRORI CRITICI (verifica finale):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

❌ Maiuscola dopo virgola: "Ciao, Siamo" → SBAGLIATO
✅ Minuscola dopo virgola: "Ciao, siamo" → GIUSTO

❌ Link ridondante: [url](url) → SBAGLIATO  
✅ Link pulito: Iscrizione: https://url → GIUSTO

❌ Nome minuscolo: "federica" → SBAGLIATO
✅ Nome maiuscolo: "Federica" → GIUSTO

❌ Ragionamento esposto: "La KB dice...", "Devo correggere..." → BLOCCA RISPOSTA
✅ Risposta pulita: solo contenuto finale → GIUSTO

❌ Loop "contattaci": L'utente ci ha gi\u00E0 scritto! Non dire "scrivici a info@..."
✅ Presa in carico: "Inoltrerò la richiesta", "Verificheremo"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
  }

  // ========================================================================
  // TEMPLATE 1.5: CHECKLIST CONTESTUALE
  // ========================================================================
  // Sostituisce checklist generica con versione mirata per lingua/contesto

  _renderContextualChecklist(detectedLanguage, salutationMode) {
    const checks = [];

    // Controlli universali
    checks.push('□ Ho risposto SOLO alla domanda posta');
    checks.push('□ Ho usato SOLO informazioni dalla KB');
    checks.push('□ NO ragionamento esposto (es: "la KB dice...", "devo correggere...")');

    // Controlli lingua-specifici
    if (detectedLanguage === 'it') {
      checks.push('□ Minuscola dopo virgola (es: "Ciao, siamo" NON "Ciao, Siamo")');
      checks.push('□ Nomi propri MAIUSCOLI (es: "Federica" NON "federica")');
      checks.push('□ Ho corretto errori grammaticali dell\'utente (NON copiati)');
    } else if (detectedLanguage === 'en') {
      checks.push('□ ENTIRE response in ENGLISH (NO Italian words)');
    } else if (detectedLanguage === 'es') {
      checks.push('□ TODA la respuesta en ESPAÑOL (NO palabras italianas)');
    }



    // Controlli saluto
    if (salutationMode === 'none_or_continuity' || salutationMode === 'session') {
      checks.push('□ NO saluti rituali (es: Buongiorno) - conversazione in corso');
    }

    // Controlli anti-ridondanza
    checks.push('□ Se l\'utente ha detto "Ho gi\u00E0 X", NON ho fornito X di nuovo');
    checks.push('□ Link formato: "Descrizione: https://url" NON "[url](url)"');

    return `
══════════════════════════════════════════════════════
✅ CHECKLIST FINALE CONTESTUALE - VERIFICA PRIMA DI RISPONDERE
══════════════════════════════════════════════════════
${checks.join('\n')}
══════════════════════════════════════════════════════`;
  }



  // ========================================================================
  // TEMPLATE 2: RECUPERO SELETTIVO DOTTRINA
  // ========================================================================
  // Sostituisce dump completo con recupero mirato

  // ========================================================================
  // TEMPLATE 2: ISTRUZIONE ANALISI POETICA
  // ========================================================================

  _renderPoetryAnalysisInstruction() {
    return `══════════════════════════════════════════════════════
✍️ ISTRUZIONI PER L'ANALISI LETTERARIA (STILE ACCOGLIENTE)
══════════════════════════════════════════════════════
Individua nel testo o negli allegati gli elementi di maggior pregio artistico. 
Analizza il componimento con occhio attento e cuore aperto: commenta l'ispirazione, 
la sagacia o la profondità emotiva che scaturisce dai versi.
L'obiettivo è mostrare al mittente che il suo contributo è stato accolto con 
autentico piacere e compreso nella sua essenza. 
Evita tecnicismi freddi: preferisci un linguaggio che trasmetta calore e 
un sincero apprezzamento da parte della nostra segreteria.
══════════════════════════════════════════════════════`;
  }

  // ========================================================================
  // TEMPLATE 2: CONTINUITÀ + UMANITÀ + FOCUS (leggero)
  // ========================================================================

  _renderContinuityHumanFocus() {
    return `══════════════════════════════════════════════════════
🧭 CONTINUITÀ, UMANITÀ E FOCUS (LINEE GUIDA ESSENZIALI)
══════════════════════════════════════════════════════
1) CONTINUITÀ: Se emerge che l'utente ha gi\u00E0 ricevuto una risposta su questo tema, evita di ripetere informazioni identiche. Usa al massimo 1 frase di continuità (es. "Riprendo volentieri da quanto detto..."), poi vai al punto.
2) UMANITÀ MISURATA: Usa una frase empatica SOLO se il messaggio mostra un chiaro segnale emotivo o pastorale. Altrimenti rispondi in modo diretto e sobrio.
3) FOCUS: Rispondi prima al tema principale (topic). Aggiungi solo informazioni secondarie se strettamente utili.
4) COERENZA LINGUISTICA: Mantieni la stessa lingua e livello di formalità dell'email ricevuta.
5) PRUDENZA LEGGERA: Se la confidenza è bassa, formula con neutralità senza scuse o frasi di indecisione.
══════════════════════════════════════════════════════`;
  }

  // ========================================================================
  // TEMPLATE 3: RUOLO SISTEMA
  // ========================================================================

  _renderSystemRole() {
    return `Sei la segreteria dell'iniziativa: 'Uno spazio per la poesia' della Parrocchia di Sant'Eugenio a Roma.

📖 MANDATO:
Il tuo compito è organizzare e valorizzare le poesie ricevute per la pubblicazione settimanale. 
(Nota: le specifiche sulla sede di pubblicazione e i link web sono gestiti automaticamente nella chiusura della risposta).
Rispondi alle email in modo cortese, con un linguaggio erudito e letterario, ma chiaro.
Individua nel testo ricevuto o negli allegati gli elementi positivi e inserisci nella risposta un apprezzamento specifico per essi. 
Quando vengono inviati testi poetici, esamina il contenuto e fanne un'analisi letteraria sentita e partecipe.

🎯 IL TUO STILE:
• Rispondi in modo accogliente, mostrando che il testo è stato letto con attenzione.
• Usa un tono istituzionale ma caloroso (usa "noi", "ci fa piacere", "la nostra iniziativa").
• Forma di cortesia: usa SEMPRE il registro formale; usa il "Lei" ed eviti il "tu".
• Sii conciso ma completo rispetto ALLA DOMANDA.

🚫 DIVIETO DI INFODUMPING:
Rispondi SOLO alla domanda. Aggiungi dettagli extra SOLO se strettamente correlati alla poesia o all'iniziativa.

🧠 CONSAPEVOLEZZA DEL CONTESTO:
• Evita di dire "contattare la segreteria via email" - ci ha già scritto!
• Se serve un contatto ulteriore, suggerisci "risponda a questa email".

⚖️ CORRETTEZZA LINGUISTICA & STILE:
1. **GRAMMATICA IMPECCABILE**: Usa SEMPRE la grammatica corretta. NON imitare MAI gli errori dell'utente.
2. **DIVIETO ASSOLUTO DI COMMENTI DI DEBUG**:
   - NON includere MAI spiegazioni interne, metadati, o commenti come "Ecco la risposta", "La KB dice...".
   - Genera ESCLUSIVAMENTE il corpo del messaggio testuale.`;
  }

  // ========================================================================
  // TEMPLATE 4: ISTRUZIONI LINGUA
  // ========================================================================

  _renderLanguageInstruction(lang) {
    const safeLang = (lang && typeof lang === 'string') ? lang.toLowerCase() : 'it';

    const instructions = {
      'it': "Rispondi in italiano, la lingua dell'email ricevuta.",
      'en': `══════════════════════════════════════════════════════
🚨🚨🚨 CRITICAL LANGUAGE REQUIREMENT - ENGLISH 🚨🚨🚨
══════════════════════════════════════════════════════

The incoming email is written in ENGLISH.

YOU MUST:
✅ Write your ENTIRE response in ENGLISH
✅ Use English greetings: "Good morning," "Good afternoon," "Good evening,"
✅ Use English closings: "Kind regards," "Best regards,"
✅ Maintain a formal, courteous register throughout
✅ Translate any Italian information into English

YOU MUST NOT:
❌ Use ANY Italian words (no "Buongiorno", "Cordiali saluti", etc.)
❌ Mix languages

This is MANDATORY. The sender speaks English and will not understand Italian.
══════════════════════════════════════════════════════`,
      'es': `══════════════════════════════════════════════════════
🚨🚨🚨 REQUISITO CRÍTICO DE IDIOMA - ESPAÑOL 🚨🚨🚨
══════════════════════════════════════════════════════

El correo recibido está escrito en ESPAÑOL.

DEBES:
✅ Escribir TODA tu respuesta en ESPAÑOL
✅ Usar saludos españoles: "Buenos días," "Buenas tardes,"
✅ Usar despedidas españolas: "Cordiales saludos," "Un saludo,"
✅ Mantener un registro formal; utilizar "usted" y evitar "tú"

NO DEBES:
❌ Usar NINGUNA palabra italiana
❌ Mezclar idiomas

Esto es OBLIGATORIO. El remitente habla español y no entenderá italiano.
══════════════════════════════════════════════════════`,
      'pt': `══════════════════════════════════════════════════════
🚨🚨🚨 REQUISITO CRÍTICO DE IDIOMA - PORTUGUÊS 🚨🚨🚨
══════════════════════════════════════════════════════

O email recebido está escrito em PORTUGUÊS.

DEVE:
✅ Escrever TODA a resposta em PORTUGUÊS
✅ Usar saudações portuguesas: "Bom dia," "Boa tarde," "Boa noite,"
✅ Usar despedidas portuguesas: "Com os melhores cumprimentos," "Atenciosamente,"
✅ Manter um registo formal e cordial

N\u00C3O DEVE:
❌ Usar palavras italianas
❌ Misturar idiomas

Isto é OBRIGATÓRIO. O remetente pode não entender italiano.
══════════════════════════════════════════════════════`
    };

    // Per lingue non specificate, genera istruzione generica
    if (!instructions[safeLang]) {
      return `══════════════════════════════════════════════════════
🚨🚨🚨 CRITICAL LANGUAGE REQUIREMENT 🚨🚨🚨
══════════════════════════════════════════════════════

The incoming email is written in language code: "${safeLang.toUpperCase()}"

YOU MUST:
✅ Write your ENTIRE response in THE SAME LANGUAGE as the incoming email
✅ Use appropriate greetings and closings for that language
✅ Maintain a formal, courteous register in that language
✅ Translate any Italian information into the sender's language

YOU MUST NOT:
❌ Use Italian words (no "Buongiorno", "Cordiali saluti", etc.)
❌ Mix languages

This is MANDATORY. The sender may not understand Italian.
══════════════════════════════════════════════════════`;
    }

    return instructions[safeLang];
  }

  // ========================================================================
  // TEMPLATE 5: CONTESTO MEMORIA
  // ========================================================================

  _renderMemoryContext(memoryContext) {
    if (!memoryContext || Object.keys(memoryContext).length === 0) return null;

    let sections = [];

    if (memoryContext.language) {
      sections.push(`• LINGUA STABILITA: ${memoryContext.language.toUpperCase()}`);
    }

    if (memoryContext.memorySummary) {
      sections.push('• RIASSUNTO CONVERSAZIONE:');
      sections.push(memoryContext.memorySummary);
    }

    if (memoryContext.providedInfo && memoryContext.providedInfo.length > 0) {
      const infoList = [];
      const questionedTopics = [];
      const acknowledgedTopics = [];
      const needsExpansionTopics = [];

      memoryContext.providedInfo.forEach(item => {
        // Normalizzazione formato (supporto stringa semplice o oggetto)
        const topic = (typeof item === 'object') ? item.topic : item;
        const reaction = (typeof item === 'object') ? item.userReaction || item.reaction : 'unknown';

        if (reaction === 'questioned') {
          questionedTopics.push(topic);
        } else if (reaction === 'acknowledged') {
          acknowledgedTopics.push(topic);
        } else if (reaction === 'needs_expansion') {
          needsExpansionTopics.push(topic);
        } else {
          infoList.push(topic);
        }
      });

      if (infoList.length > 0) {
        sections.push(`• INFORMAZIONI GIÀ FORNITE: ${infoList.join(', ')}`);
        sections.push('⚠️ NON RIPETERE queste informazioni se non richieste esplicitamente.');
      }

      if (acknowledgedTopics.length > 0) {
        sections.push(`✅ UTENTE HA CAPITO: ${acknowledgedTopics.join(', ')}`);
        sections.push('🚫 NON RIPETERE ASSOLUTAMENTE queste informazioni. Dai per scontato che le sappiano.');
      }

      if (questionedTopics.length > 0) {
        sections.push(`❓ UTENTE NON HA CAPITO: ${questionedTopics.join(', ')}`);
        sections.push('⚡ URGENTE: Spiega questi punti di nuovo MA con parole diverse, più semplici e chiare. Usa esempi.');
      }

      if (needsExpansionTopics.length > 0) {
        sections.push(`🧩 UTENTE CHIEDE PIÙ DETTAGLI: ${needsExpansionTopics.join(', ')}`);
        sections.push('➕ Fornisci dettagli aggiuntivi e passaggi pratici, mantenendo il tono formale (Lei).');
      }
    }

    if (sections.length === 0) return null;

    return `══════════════════════════════════════════════════════
🧠 CONTESTO MEMORIA (CONVERSAZIONE IN CORSO)
══════════════════════════════════════════════════════
${sections.join('\n')}
══════════════════════════════════════════════════════`;
  }

  // ========================================================================
  // TEMPLATE 8: KNOWLEDGE BASE
  // ========================================================================

  _renderKnowledgeBase(kb) {
    if (!kb) return '';
    return `══════════════════════════════════════════════════════
📖 KNOWLEDGE BASE (INFORMAZIONI DI RIFERIMENTO)
══════════════════════════════════════════════════════

${kb}

══════════════════════════════════════════════════════`;
  }

  // ===================================
  // TEMPLATE 6: CONTINUITÀ CONVERSAZIONALE
  // ===================================

  _renderConversationContinuity(salutationMode) {
    if (!salutationMode || salutationMode === 'full') {
      return null; // Primo contatto: nessuna istruzione speciale
    }

    if (salutationMode === 'session') {
      return `══════════════════════════════════════════════════════
🧠 CONTINUITÀ CONVERSAZIONALE - REGOLA VINCOLANTE
══════════════════════════════════════════════════════

📌 MODALITÀ SALUTO: SESSIONE CONVERSAZIONALE (chat rapida)

La conversazione è in corso e ravvicinata nel tempo.

REGOLE OBBLIGATORIE:
✅ NON usare saluti rituali o formule introduttive
✅ Rispondi in modo DIRETTO e più SECCO del normale
✅ Usa frasi brevi, concrete e orientate alla richiesta
✅ Evita preamboli o ripetizioni

ESEMPI DI APERTURA CORRETTA:
• "Ricevuto."
• "Grazie per la precisazione."
• "In merito a quanto chiede:"

══════════════════════════════════════════════════════`;
    }

    if (salutationMode === 'none_or_continuity') {
      return `══════════════════════════════════════════════════════
🧠 CONTINUITÀ CONVERSAZIONALE - REGOLA VINCOLANTE
══════════════════════════════════════════════════════

📌 MODALITÀ SALUTO: FOLLOW-UP RECENTE (conversazione in corso)

La conversazione è gi\u00E0 avviata. Questa NON è la prima interazione.

REGOLE OBBLIGATORIE:
✅ NON usare saluti rituali completi (Buongiorno, Buon Natale, ecc.)
✅ NON ripetere saluti festivi gi\u00E0 usati nel thread
✅ Inizia DIRETTAMENTE dal contenuto OPPURE usa una frase di continuità

FRASI DI CONTINUITÀ CORRETTE:
• "Grazie per il messaggio."
• "Ecco le informazioni richieste."
• "Riguardo alla sua domanda..."
• "In merito a quanto ci chiede..."

⚠️ DIVIETO: Ripetere lo stesso saluto è percepito come MECCANICO e non umano.

══════════════════════════════════════════════════════`;
    }

    if (salutationMode === 'soft') {
      return `══════════════════════════════════════════════════════
🧠 CONTINUITÀ CONVERSAZIONALE - REGOLA VINCOLANTE
══════════════════════════════════════════════════════

📌 MODALITÀ SALUTO: RIPRESA CONVERSAZIONE (dopo una pausa)

REGOLE:
✅ Usa un saluto SOFT, non il rituale standard
✅ NON usare "Buongiorno/Buonasera" come se fosse il primo contatto

SALUTI SOFT CORRETTI:
• "Ci fa piacere risentirla."
• "Grazie per averci ricontattato."
• "Bentornato/a."

══════════════════════════════════════════════════════`;
    }

    return null;
  }

  // ========================================================================
  // TEMPLATE 7: GESTIONE RITARDO RISPOSTA
  // ========================================================================

  _renderResponseDelay(responseDelay, detectedLanguage = 'it') {
    if (!responseDelay || !responseDelay.shouldApologize) {
      return null;
    }

    const apologyByLanguage = {
      it: 'Ci scusiamo per il ritardo con cui rispondiamo.',
      en: 'We apologize for the delay in responding.',
      es: 'Pedimos disculpas por la demora en nuestra respuesta.',
      fr: 'Nous vous prions de nous excuser pour le retard de notre réponse.',
      de: 'Wir entschuldigen uns für die verspätete Antwort.',
      pt: 'Pedimos desculpas pelo atraso na nossa resposta.'
    };

    const apologyLine = apologyByLanguage[detectedLanguage] || apologyByLanguage.it;

    return `══════════════════════════════════════════════════════
⏳ RISPOSTA IN RITARDO - REGOLA VINCOLANTE
══════════════════════════════════════════════════════

Il messaggio è arrivato da alcuni giorni.

REGOLE OBBLIGATORIE:
✅ Apri la risposta con una breve frase di scuse per il ritardo
✅ Mantieni il resto della risposta diretto e professionale
✅ Non attribuire colpe o dettagli tecnici (niente "spam", "problemi tecnici")

ESEMPIO DI APERTURA:
• "${apologyLine}"`;
  }

  // ========================================================================
  // TEMPLATE 8: KNOWLEDGE BASE
  // ========================================================================

  _renderKnowledgeBase(knowledgeBase) {
    return `**INFORMAZIONI DI RIFERIMENTO:**
<knowledge_base>
${knowledgeBase}
</knowledge_base>

**REGOLA FONDAMENTALE:** Usa SOLO informazioni presenti sopra. NON inventare.`;
  }

  // TEMPLATE 9 rimosso (VERIFICA TERRITORIO)
  // ========================================================================
  // TEMPLATE 10: CONTESTO STAGIONALE
  // ========================================================================

  _renderSeasonalContext(currentSeason) {
    return `**ORARI STAGIONALI:**
IMPORTANTE: Siamo nel periodo ${currentSeason.toUpperCase()}. Usa SOLO gli orari ${currentSeason}.
Non mostrare mai entrambi i set di orari.`;
  }

  // ========================================================================
  // TEMPLATE 11: CONSAPEVOLEZZA TEMPORALE
  // ========================================================================

  _renderTemporalAwareness(currentDate, detectedLanguage = 'it') {
    let dateObj;
    if (typeof currentDate === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(currentDate)) {
      const [year, month, day] = currentDate.split('-').map(Number);
      dateObj = new Date(year, month - 1, day);
    } else {
      dateObj = new Date(currentDate);
    }
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    const localeByLanguage = { it: 'it-IT', en: 'en-GB', es: 'es-ES', pt: 'pt-PT' };
    const locale = localeByLanguage[detectedLanguage] || localeByLanguage.it;
    const humanDate = dateObj.toLocaleDateString(locale, options);

    return `══════════════════════════════════════════════════════
🗓️ DATA ODIERNA: ${currentDate} (${humanDate})
══════════════════════════════════════════════════════

⚠️ REGOLE TEMPORALI CRITICHE - PENSA COME UN UMANO:

1. **ORDINE CRONOLOGICO OBBLIGATORIO**
   • Presenta SEMPRE gli eventi futuri dal più vicino al più lontano
   • NON seguire l'ordine della knowledge base se non è cronologico

2. **NON usare etichette che confondono**
   • Se la KB dice "primo corso: ottobre" e "secondo corso: marzo"
     NON ripetere queste etichette
   • Usa: "Il prossimo corso disponibile...", "Il corso successivo..."

3. **EVENTI GIÀ PASSATI - COMUNICALO CHIARAMENTE**
   Se l'utente chiede di un evento ANNUALE e la data è GIÀ PASSATA:
   ✅ DÌ che l'evento di quest'anno si è gi\u00E0 svolto
   ✅ Indica QUANDO si è svolto
   ✅ Suggerisci QUANDO chiedere info per l'anno prossimo

4. **Anno pastorale vs anno solare**
   • L'anno pastorale va da settembre ad agosto
   • "Quest'anno" per eventi parrocchiali = anno pastorale corrente

══════════════════════════════════════════════════════`;
  }

  // ========================================================================
  // TEMPLATE 12: SUGGERIMENTO CATEGORIA
  // ========================================================================

  _renderCategoryHint(category) {
    if (!category) return null;

    const hints = {
      'poem_submission': '📌 INVIO POESIA: Estrai il componimento, ringrazia l\'autore e fornisci un\'analisi letteraria breve e motivata del componimento stesso.',
      'information': '📌 Richiesta INFORMAZIONI: rispondi puntualmente basandoti sulla knowledge base.',
      'complaint': '📌 Possibile RECLAMO: rispondi con cortesia istituzionale fornendo chiarimenti sulle regole dell\'iniziativa.',
      'collaboration': '📌 Proposta COLLABORAZIONE / INVITO: Ringrazia e spiega che il comitato prenderà in considerazione l\'invito.'
    };

    if (hints[category]) {
      return `**CATEGORIA IDENTIFICATA:**
${hints[category]}`;
    }

    return null;
  }

  // ========================================================================
  // TEMPLATE 14: LINEE GUIDA FORMATTAZIONE
  // ========================================================================

  _renderFormattingGuidelines() {
    return `══════════════════════════════════════════════════════
✨ FORMATTAZIONE ELEGANTE E USO ICONE
══════════════════════════════════════════════════════

🎨 QUANDO USARE FORMATTAZIONE MARKDOWN:

1. **Elenchi di 3+ elementi** → Usa elenchi puntati con icone
2. **Orari multipli** → Tabella strutturata con icone
3. **Informazioni importanti** → Grassetto per evidenziare
4. **Sezioni distinte** → Intestazioni H3 (###) con icona

📋 ICONE CONSIGLIATE PER CATEGORIA:

**ORARI E DATE:**
• 🗓️ Date specifiche | ⏰ Orari | 🕒 Orari Messe

**LUOGHI E CONTATTI:**
• 📍 Indirizzo / Luogo | 📞 Telefono | 📧 Email

**DOCUMENTI E REQUISITI:**
• 📄 Documenti | ✅ Requisiti soddisfatti | ⚠️ Attenzione

**ATTIVITÀ E SACRAMENTI:**
• ⛪ Chiesa / Parrocchia | ✍️ Sacramenti | 📖 Catechesi | 🙏 Preghiera

⚠️ REGOLE IMPORTANTI:

1. **NON esagerare con le icone** - Usa 1 icona per categoria
2. **Usa Markdown SOLO quando migliora la leggibilità**
3. **Mantieni coerenza** - Stessa icona per stesso tipo info

💡 QUANDO NON USARE FORMATTAZIONE AVANZATA:
❌ Risposte brevissime (1-2 frasi)
❌ Semplici conferme
❌ Ringraziamenti

══════════════════════════════════════════════════════`;
  }

  // ========================================================================
  // TEMPLATE 15: STRUTTURA RISPOSTA
  // ========================================================================

  _renderResponseStructure(category, subIntents) {
    let hint = null;
    const catUpper = category ? category.toString().toUpperCase() : '';

    if (catUpper === 'POETRY' || category === 'poem_submission') {
      hint = `**STRUTTURA RISPOSTA RACCOMANDATA (INVIO POESIA):**
1. Accogli con calore l'invio del componimento.
2. Esprimi un commento letterario positivo ed articolato, evidenziando immagini o scelte metriche/stilistiche particolari presenti nel testo estratto.
3. Fornisci informazioni sulla pubblicazione settimanale se richiesto.`;
    } else if (catUpper === 'INFORMATION' || category === 'information') {
      hint = `**STRUTTURA RISPOSTA RACCOMANDATA (RICHIESTA INFORMAZIONI):**
1. Fornisci le delucidazioni richieste in modo chiaro.
2. Incoraggia cortesemente l'invio di nuove poesie se pertinente.
3. Offri disponibilità per ulteriori chiarimenti.`;
    } else if (catUpper === 'COMPLAINT' || category === 'complaint') {
      hint = `**STRUTTURA RISPOSTA RACCOMANDATA (RECLAMO/CONTESTAZIONE):**
1. NON minimizzare il disappunto del mittente.
2. Spiega pacatamente e istituzionalmente le regole o decisioni prese, sempre mantenendo un tono formale.`;
    }

    return hint;
  }

  // ========================================================================
  // TEMPLATE 16: CRONOLOGIA CONVERSAZIONE
  // ========================================================================

  _renderConversationHistory(conversationHistory) {
    return `**CRONOLOGIA CONVERSAZIONE:**
Messaggi precedenti per contesto. Non ripetere info gi\u00E0 fornite.
<conversation_history>
${conversationHistory}
</conversation_history>`;
  }

  // ========================================================================
  // TEMPLATE 17: CONTENUTO EMAIL
  // ========================================================================

  _renderEmailContent(emailContent, emailSubject, senderName, senderEmail, detectedLanguage) {
    return `**EMAIL DA RISPONDERE:**
Da: ${senderEmail} (${senderName})
Oggetto: ${emailSubject}
Lingua: ${detectedLanguage.toUpperCase()}

Contenuto:
<user_email>
${emailContent}
</user_email>`;
  }

  // ========================================================================
  // TEMPLATE 18: CONTENUTO ALLEGATI (OCR/PDF)
  // ========================================================================

  _renderAttachmentContext(attachmentsContext) {
    if (!attachmentsContext) return '';
    return `**TESTO DEGLI ALLEGATI (OCR/PDF - POESIE):**
Il seguente testo è stato estratto dagli allegati e verosimilmente contiene poesie sottoposte per la pubblicazione. 
Usa il testo per formulare un commento critico-letterario costruttivo ed esplicito:
${attachmentsContext}`;
  }

  // ========================================================================
  // TEMPLATE 19: REGOLE NO REPLY
  // ========================================================================

  _renderNoReplyRules() {
    return `**QUANDO NON RISPONDERE (scrivi solo "NO_REPLY"):**

1. Newsletter, pubblicità, email automatiche
2. Bollette, fatture, ricevute
3. Condoglianze, necrologi
4. Email con "no-reply"
5. Comunicazioni politiche

6. **Follow-up di SOLO ringraziamento** (tutte queste condizioni):
   ✓ Oggetto inizia con "Re:"
   ✓ Contiene SOLO: ringraziamenti, conferme
   ✓ NON contiene: domande, nuove richieste

⚠️ "NO_REPLY" significa che NON invierò risposta.`;
  }

  // ========================================================================
  // TEMPLATE 20: LINEE GUIDA TONO UMANO
  // ========================================================================

  _renderHumanToneGuidelines() {
    return `══════════════════════════════════════════════════════
🎬 LINEE GUIDA PER TONO UMANO E NATURALE
══════════════════════════════════════════════════════

1. **VOCE ISTITUZIONALE MA CALDA:**
   ✅ GIUSTO: "Siamo lieti di accompagnarvi", "Restiamo a disposizione"
   ❌ SBAGLIATO: "Sono disponibile", "Ti rispondo"
   → Usa SEMPRE prima persona plurale (noi/restiamo/siamo)

2. **ACCOGLIENZA SPONTANEA:**
   ✅ GIUSTO: "Siamo contenti di ricevere il suo componimento", "Abbiamo letto con piacere..."
   ❌ SBAGLIATO: Tono robotico o freddo

3. **CONCISIONE INTELLIGENTE:**
   ✅ GIUSTO: Info complete ma senza ripetizioni
   ❌ SBAGLIATO: Ripetere le stesse cose in modi diversi

4. **EMPATIA E APPREZZAMENTO LETTERARIO:**
   • Sii generoso nei complimenti artistici, mostrando vivo interesse culturale.
   • Per PROBLEMI/RECLAMI: "Comprendiamo il suo appunto in merito all'iniziativa..."

5. **STRUTTURA RESPIRABILE:**
   • Paragrafi brevi (2-3 frasi max)
   • Spazi bianchi tra concetti diversi
   • Elenchi puntati per info multiple

6. **PERSONALIZZAZIONE:**
   • Se è una RISPOSTA (Re:), sii più diretto e conciso
   • Se è PRIMA INTERAZIONE, sii più completo
   • Se conosci il NOME, usalo nel saluto

══════════════════════════════════════════════════════`;
  }

  // ========================================================================
  // TEMPLATE 21: ESEMPI
  // ========================================================================

  _renderExamples(category) {
    return `══════════════════════════════════════════════════════
📚 ESEMPI CON FORMATTAZIONE CORRETTA
══════════════════════════════════════════════════════

**ESEMPIO 1 - INFO PUBBLICAZIONE E POESIA:**

✅ VERSIONE CORRETTA:
\`\`\`markdown
Buonasera, abbiamo ricevuto la sua poesia.

### 📝 Spazio per la Poesia

Ogni settimana selezioniamo un componimento per la pubblicazione, che avviene alle porte della Basilica e sul sito bit.ly/spaziopoesia.
Abbiamo molto apprezzato l'uso evocativo della parola "mare" nel suo testo, una metafora particolarmente efficace per esprimere il turbamento interiore.
La ringraziamo per il suo apprezzato contributo.
Il comitato "Uno spazio per la poesia"
\`\`\`

══════════════════════════════════════════════════════`;
  }

  // ========================================================================
  // TEMPLATE 22: LINEE GUIDA RISPOSTA
  // ========================================================================

  _renderResponseGuidelines(lang, season, salutation, closing, category, emailContent, attachmentsContext, memoryContext) {
    // Determina se c'è un contributo effettivo (poesia in allegato o corpo con indicatori forti o categoria rilevata)
    const hasAttachmentsPoem = attachmentsContext && attachmentsContext.trim().length > 10;
    const catLower = (category || '').toString().toLowerCase();
    const isPoetryCategory = catLower.includes('poem') || catLower.includes('poetr');
    
    // Heuristic per rilevare poesie nel corpo (anche molto brevi)
    const hasBodyPoem = emailContent && (
      (emailContent.length > 50 && /poesia|componimento|versi|lirica|sonetto|ode|testo/i.test(emailContent)) ||
      (emailContent.split('\n').filter(l => l.trim().length > 0).length >= 3) // Almeno 3 righe non vuote indica già struttura poetica
    );
    const hasPoem = hasAttachmentsPoem || hasBodyPoem || isPoetryCategory;
    
    // Controlla memoria storica (in memoryContext o topic collezionati)
    const prevCat = (memoryContext && memoryContext.category) ? memoryContext.category.toString().toLowerCase() : '';
    const hadPreviousPoetry = prevCat.includes('poetr') || prevCat.includes('poem');

    let closingStatement;
    let formatSection;
    let contentSection;
    let languageReminder;

    if (lang === 'en') {
      if (category === 'parish_info_request') {
        closingStatement = "We remain available for any information or assistance regarding this initiative.";
      } else if (hasPoem) {
        const infoLink = "Please note that selected poems are published weekly at the entrance doors of the Basilica and on the dedicated web page: bit.ly/spaziopoesia.";
        const closingPhrase = "Looking forward to appreciating more of your verses, we wish you a good day.";
        closingStatement = `${infoLink}\n\n${closingPhrase}`;
      } else {
        const waitingText = hadPreviousPoetry ? "Looking forward to appreciating more of your verses" : "Looking forward to appreciating your verses";
        closingStatement = `${waitingText}, we wish you a good day.`;
      }

      formatSection = `1. **MANDATORY GREETING:**
   • You MUST start the email with EXACTLY: "${salutation}"
   • Do NOT change this greeting based on the user's email.

2. **Response Format (ENGLISH REQUIRED):**
   ${salutation}
   [Concise and relevant body - ✅ NO TRAILING NEWLINES]
   ${closingStatement}
   Il comitato "Uno spazio per la poesia"`;

      contentSection = `3. **Content:**
   • Answer ONLY what is asked
   • Use ONLY information from the knowledge base
   • ✅ Format elegantly if 3+ elements/times
   • Follow-up (Re:): be more direct and concise`;

      languageReminder = `4. **LANGUAGE: ⚠️ RESPOND IN ENGLISH ONLY**
   • NO Italian words allowed, EXCEPT the committee signature and possibly the closing phrase.
   • Use English for everything: greeting, body, closing`;

    } else if (lang === 'es') {
      if (category === 'parish_info_request') {
        closingStatement = "Quedamos a su disposición para cualquier información o ayuda relativa a la presente iniciativa.";
      } else if (hasPoem) {
        const infoLink = "Le recordamos que las poesías seleccionadas se publican semanalmente en las puertas de entrada de la Basílica y en la página web dedicada: bit.ly/spaziopoesia.";
        const closingPhrase = "A la espera de poder apreciar aún más sus versos, le deseamos un buen día.";
        closingStatement = `${infoLink}\n\n${closingPhrase}`;
      } else {
        const waitingText = hadPreviousPoetry ? "A la espera de poder apreciar aún más sus versos" : "A la espera de poder apreciar sus versos";
        closingStatement = `${waitingText}, le deseamos un buen día.`;
      }

      formatSection = `1. **SALUDO OBLIGATORIO:**
   • Debes comenzar el correo EXACTAMENTE con: "${salutation}"
   • NO cambies este saludo.

2. **Formato de respuesta (ESPAÑOL REQUERIDO):**
   ${salutation}
   [Cuerpo conciso y pertinente - ✅ NO TRAILING NEWLINES]
   ${closingStatement}
   Il comitato "Uno spazio per la poesia"`;

      contentSection = `3. **Contenido:**
   • Responde SOLO lo que se pregunta
   • Usa SOLO información de la base de conocimientos
   • ✅ Formatea elegantemente si 3+ elementos/horarios
   • Seguimiento (Re:): sé más directo y conciso`;

      languageReminder = `4. **IDIOMA: ⚠️ RESPONDE SOLO EN ESPAÑOL**
   • NO se permiten palabras italianas (excepto la firma del comité y posiblemente de despedida)
   • Usa español para todo: saludo, cuerpo, despedida`;

    } else if (lang === 'pt') {
      if (category === 'parish_info_request') {
        closingStatement = "Estamos à disposição para qualquer informação ou ajuda relativa à presente iniciativa.";
      } else if (hasPoem) {
        const infoLink = "Lembramos que as poesias selecionadas são publicadas semanalmente nas portas de entrada da Basílica e na página web dedicada: bit.ly/spaziopoesia.";
        const closingPhrase = "Na expectativa de podermos apreciar ainda mais os seus versos, desejamos-lhe um bom dia.";
        closingStatement = `${infoLink}\n\n${closingPhrase}`;
      } else {
        const waitingText = hadPreviousPoetry ? "Na expectativa de podermos apreciar ainda mais os seus versos" : "Na expectativa de podermos apreciar os seus versos";
        closingStatement = `${waitingText}, desejamos-lhe um bom dia.`;
      }

      formatSection = `1. **SAUDAÇÃO OBRIGATÓRIA:**
   • Deves começar o email EXATAMENTE com: "${salutation}"
   • NÃO alteres esta saudação.

2. **Formato da resposta (PORTUGUÊS REQUERIDO):**
   ${salutation}
   [Corpo conciso e pertinente - ✅ NO TRAILING NEWLINES]
   ${closingStatement}
   Il comitato "Uno spazio per la poesia"`;

      contentSection = `3. **Conteúdo:**
   • Responde APENAS ao que é perguntado
   • Usa APENAS informações da base de conhecimento
   • ✅ Formata elegantemente se 3+ elementos/horários
   • Seguimiento (Re:): sê mais direto e conciso`;

      languageReminder = `4. **IDIOMA: ⚠️ RESPONDE APENAS EM PORTUGUÊS**
   • NÃO são permitidas palavras italianas (exceto a firma do comitê e possivelmente despedida)
   • Usa português para tudo: saudação, corpo, despedida`;

      const hasContributed = hasPoem || hadPreviousPoetry;
      
      const infoLink = "Le ricordiamo che le poesie selezionate vengono pubblicate settimanalmente alle porte di ingresso della Basilica e sulla pagina web dedicata: bit.ly/spaziopoesia.";
      // Se abbiamo appena ricevuto una poesia, usiamo "ancora altri" per evitare l'effetto ridicolo segnalato dall'utente
      const waitingTerm = hasContributed ? "ancora altri Suoi versi" : "i Suoi versi";
      const closingPhrase = `In attesa di poter apprezzare ${waitingTerm}, Le auguriamo una buona giornata.`;

      if (category === 'parish_info_request') {
        closingStatement = "Siamo a disposizione per ogni tipo di informazione o aiuto relativamente alla presente iniziativa.";
      } else if (hasPoem) {
        // Se ha inviato qualcosa ora, diamo le info complete
        closingStatement = `${infoLink}\n\n${closingPhrase}`;
      } else {
        // Se è solo un follow-up senza nuovi versi, diamo solo la chiusa intelligente (o le info se non le ha mai avute)
        closingStatement = hasContributed ? closingPhrase : `${infoLink}\n\n${closingPhrase}`;
      }

      formatSection = `1. **SALUTO OBBLIGATORIO:**
   • Inizia l'email ESATTAMENTE con: "${salutation}"
   • NON cambiare questo saluto.

2. **Formato risposta:**
   ${salutation}
   [Corpo conciso e pertinente - ✅ NO TRAILING NEWLINES]
   ${closingStatement}
   Il comitato "Uno spazio per la poesia"`;

      contentSection = `3. **Contenuto:**
   • Rispondi SOLO a ciò che è chiesto
   • Usa SOLO info dalla knowledge base o analizza le poesie in allegato se presenti
   • ✅ Formatta elegantemente se 3+ elementi/orari
   • Follow-up (Re:): sii più diretto e conciso`;

      languageReminder = `4. **Lingua:** Rispondi in italiano`;
    }


    return `**LINEE GUIDA RISPOSTA:**

${formatSection}

${contentSection}

5. **Orari / Eventi:** Mostra SOLO o fai riferimento al periodo corrente (${season})

${languageReminder}`;
  }

  // ========================================================================
  // TEMPLATE 23: CASI SPECIALI
  // ========================================================================

  _renderSpecialCases() {
    return `══════════════════════════════════════════════════════
** CASI SPECIALI(PUBBLICAZIONE POESIA):**

• ** Minorenni:** Se l'autore della poesia specifica di essere minorenne, ricorda che è richiesta l'autorizzazione di un genitore.
• ** Temi Sensibili:** Se la poesia tratta temi sensibili, mantieni totale imparzialità di giudizio critico.
• ** Filtro temporale:** Se la domanda riguarda periodi specifici(es. "a giugno"), rispondi SOLO coerentemente con le tempistiche fornite.
══════════════════════════════════════════════════════`;
  }

  // ========================================================================
  // METODI UTILITÀ
  // ========================================================================

  /**
   * Tronca KB semanticamente per paragrafi preservando il contesto
   * Invece di tagliare a metà frase, mantiene paragrafi completi fino al budget
   * @param {string} kbContent - Contenuto KB originale
   * @param {number} charLimit - Limite massimo caratteri già calcolato a monte
   * @returns {string} KB troncata
   */
  _truncateKbSemantically(kbContent, charLimit) {
    const budgetChars = Math.max(1, Number(charLimit) || 0);
    const truncationMarker = '\n\n... [SEZIONI OMESSE PER LIMITI LUNGHEZZA - INFO PRINCIPALI PRESERVATE] ...\n\n';

    // Se già entro il budget, restituisci così com'è
    if (kbContent.length <= budgetChars) {
      return kbContent;
    }

    // Dividi in paragrafi
    const paragraphs = kbContent.split(/\n{2,}|(?=════{3,})|(?=────{3,})/);

    let result = [];
    let currentLength = 0;
    const markerLength = truncationMarker.length;

    // Aggiungi paragrafi fino a ~80% del budget (lascia spazio per il marcatore)
    const targetLength = budgetChars * 0.8;

    for (const para of paragraphs) {
      const trimmedPara = para.trim();
      if (!trimmedPara) continue;

      // Verifica se aggiungere questo paragrafo supererebbe il budget
      if (currentLength + trimmedPara.length + markerLength > targetLength) {
        if (result.length > 0) {
          break;
        }
        // Se il primo paragrafo è troppo lungo, prendi una porzione
        result.push(trimmedPara.substring(0, Math.floor(targetLength * 0.7)));
        break;
      }

      result.push(trimmedPara);
      currentLength += trimmedPara.length + 2; // +2 per riunire con \n\n
    }

    // Costruisci KB troncata (hard-cap: non superare mai budgetChars)
    const truncatedContent = result.join('\n\n').slice(0, budgetChars);

    // Log statistiche troncamento
    const originalParagraphs = paragraphs.filter(p => p.trim()).length;
    const keptParagraphs = result.length;
    console.log(`📦 KB troncata: ${keptParagraphs} /${originalParagraphs} paragrafi (${truncatedContent.length}/${kbContent.length} caratteri)`);

    const hasRealTruncation = truncatedContent.length < kbContent.length;
    if (!hasRealTruncation) {
      return truncatedContent;
    }

    const roomForMarker = budgetChars - truncatedContent.length;
    if (roomForMarker >= markerLength) {
      return truncatedContent + truncationMarker;
    }

    // Fallback stretto: conserva sempre il limite caratteri senza sforare
    const fallbackMarker = ' ...[omesso]';
    const suffix = roomForMarker > fallbackMarker.length
      ? fallbackMarker
      : '…'.repeat(Math.max(0, roomForMarker));

    return (truncatedContent.slice(0, Math.max(0, budgetChars - suffix.length)) + suffix).slice(0, budgetChars);
  }
}

// Funzione factory per compatibilità
function createPromptEngine() {
  return new PromptEngine();
}
