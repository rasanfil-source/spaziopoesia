/**
 * RequestTypeClassifier.gs - Classificazione Poesia / Informazioni
 * 
 * TIPI RICHIESTA:
 * - POEM_SUBMISSION: invio di una poesia o testo per pubblicazione settimanale
 * - INFORMATION_REQUEST: richiesta di info su scadenze, premi, regole
 */
class RequestTypeClassifier {
  constructor() {
    console.log('📊 Inizializzazione RequestTypeClassifier...');

    this.POEM_INDICATORS = [
      { pattern: /\bpoesia\b/i, weight: 3 },
      { pattern: /\bcomponimento\b/i, weight: 3 },
      { pattern: /\bversi\b/i, weight: 2 },
      { pattern: /\bpartecipo\b/i, weight: 2 },
      { pattern: /\ballego\b/i, weight: 2 },
      { pattern: /\binvio\b/i, weight: 2 },
      { pattern: /\biniziativa\b/i, weight: 2 },
      { pattern: /\btesto\b/i, weight: 1 }
    ];

    this.INFO_INDICATORS = [
      { pattern: /\bcome (?:si )?fa\b/i, weight: 2 },
      { pattern: /\bquando\b/i, weight: 2 },
      { pattern: /\bregolamento\b/i, weight: 3 },
      { pattern: /\binformazioni\b/i, weight: 2 },
      { pattern: /\bdove\b/i, weight: 1 },
      { pattern: /\bposso\b/i, weight: 1 }
    ];

    this.PARISH_INDICATORS = [
      { pattern: /\bmess[ae]\b/i, weight: 3 },
      { pattern: /\borari(?:o)?\b/i, weight: 2 },
      { pattern: /\bpellegrinaggi(?:o)?\b/i, weight: 3 },
      { pattern: /\bincontr[io]\b/i, weight: 2 },
      { pattern: /\bcatechesi\b/i, weight: 3 },
      { pattern: /\bparrocchi[ae]\b/i, weight: 2 },
      { pattern: /\battivit[aà]\b/i, weight: 2 },
      { pattern: /\bcammino di santiago\b/i, weight: 4 },
      { pattern: /\bscout\b/i, weight: 2 },
      { pattern: /\bchies[ae]\b/i, weight: 2 },
      { pattern: /\bterra santa\b/i, weight: 1 },
      { pattern: /\bliturgia\b/i, weight: 1 },
      { pattern: /\bconfession[ei]\b/i, weight: 1 },
      { pattern: /\bcresim[ae]\b/i, weight: 1 },
      { pattern: /\bprim[ae]\s+comunion[ei]\b/i, weight: 1 },
      { pattern: /\bprematrimonial[ei]\b/i, weight: 1 }
    ];

    console.log('✓ RequestTypeClassifier inizializzato');
  }

  /**
   * Classifica la richiesta email
   * Restituisce dimensioni continue, complessità e tono suggerito.
   */
  classify(subject, body, externalHint = null) {
    // Smart Truncation (primi 1500 + ultimi 1500 caratteri)
    const MAX_ANALYSIS_LENGTH = 3000;
    const sanitizedText = this._sanitizeText(subject, body);
    const text = sanitizedText.length > MAX_ANALYSIS_LENGTH
      ? (
        sanitizedText.substring(0, 1500) +
        ' ... ' +
        sanitizedText.substring(sanitizedText.length - 1500)
      ).toLowerCase()
      : sanitizedText.toLowerCase();

    // 1. Calcola punteggi grezzi
    const poemResult = this._calculateScore(text, this.POEM_INDICATORS);
    const infoResult = this._calculateScore(text, this.INFO_INDICATORS);
    const parishResult = this._calculateScore(text, this.PARISH_INDICATORS);

    // 2. Normalizzazione Punteggi (0.0 - 1.0)
    // Soglia saturazione arbitraria: 4 match = 1.0
    const SATURATION_POINT = 4;
    let dimensions = {
      poem_submission: Math.min(poemResult.score / SATURATION_POINT, 1.0),
      information_request: Math.min(infoResult.score / SATURATION_POINT, 1.0),
      parish_info_request: Math.min(parishResult.score / SATURATION_POINT, 1.0)
    };

    // 3. Determinazione Tipo Primario
    let requestType = 'poem_submission';

    if (dimensions.parish_info_request >= 0.6 && dimensions.parish_info_request > dimensions.poem_submission && dimensions.parish_info_request >= dimensions.information_request) {
      requestType = 'parish_info_request';
    } else if (dimensions.information_request >= 0.6 && dimensions.information_request > dimensions.poem_submission) {
      requestType = 'information_request';
    } else if (dimensions.poem_submission >= 0.6) {
      requestType = 'poem_submission';
    } else {
      // Caso dubbio, default a invio poesia
      requestType = 'poem_submission';
    }

    // 4. Calcolo Metriche Derivate

    // Complessità: Somma delle dimensioni attive (> 0.2)
    const activeDims = Object.values(dimensions).filter(v => v > 0.2).length;
    let complexity = 'Low';
    if (activeDims >= 2 || Math.max(...Object.values(dimensions)) > 0.8) complexity = 'Medium';

    // Tono Suggerito
    let suggestedTone = 'Incoraggiante e Apprezzativo';
    if (requestType === 'information_request') suggestedTone = 'Professionale e Chiaro';
    if (requestType === 'parish_info_request') suggestedTone = 'Cortese e Indirizzante';

    const result = {
      type: requestType, // Categoria classica
      source: 'regex',
      dimensions: dimensions, // Nuova metrica
      complexity: complexity,
      suggestedTone: suggestedTone,

      poemScore: dimensions.poem_submission, // Normalizzati
      infoScore: dimensions.information_request,
      parishScore: dimensions.parish_info_request,

      detectedIndicators: [
        ...poemResult.matched,
        ...infoResult.matched,
        ...parishResult.matched
      ]
    };

    console.log(`   📊 Classificazione: ${requestType.toUpperCase()} (Tono: ${suggestedTone})`);
    return result;
  }

  /**
   * Interfaccia semplificata (soggetto + corpo) con ordine parametri invertito.
   */
  classifyRequest(body, subject, externalHint = null) {
    return this.classify(subject, body, externalHint);
  }

  /**
   * Calcola punteggio ponderato per set di indicatori
   */
  _calculateScore(text, indicators) {
    let total = 0;
    const matched = [];
    let matchCount = 0;

    for (const indicator of indicators) {
      // Forza flag globale per conteggio corretto di tutte le occorrenze
      const flags = indicator.pattern.flags.includes('g')
        ? indicator.pattern.flags
        : indicator.pattern.flags + 'g';

      const pattern = new RegExp(indicator.pattern.source, flags);

      const matches = text.match(pattern);
      if (matches) {
        // match() con global flag ritorna array di stringhe matchate, non gruppi
        total += indicator.weight * matches.length;
        matchCount += matches.length;
        matched.push(indicator.pattern.source);
      }
    }

    return { score: total, matched: matched, matchCount: matchCount };
  }

  /**
   * Sanitizza il testo evitando falsi positivi da quote e firme
   */
  _sanitizeText(subject, body) {
    let text = `${subject || ''}\n${body || ''}`;

    let iterations = 0;
    while (/<blockquote/i.test(text) && iterations < 10) {
      const previousText = text;
      text = text.replace(/<blockquote[^>]*>[\s\S]*?<\/blockquote>/gi, '');
      if (text === previousText) {
        break;
      }
      iterations++;
    }
    if (iterations >= 10) {
      text = text.replace(/<blockquote[^>]*>[\s\S]*$/gi, '');
    }
    text = text.replace(/<div\s+class=["']gmail_quote["'][^>]*>[\s\S]*?<\/div>/gi, '');
    text = text.replace(/<div\s+id=["']?divRplyFwdMsg["']?[^>]*>[\s\S]*?$/gi, '');

    const lines = text.split('\n');
    const cleaned = [];
    let inQuotedSection = false;
    let inSignature = false;

    for (const line of lines) {
      const stripped = line.trim();

      if (stripped === '') {
        cleaned.push('');
        continue;
      }

      if (/^--\s*$/.test(stripped) || /^__+$/.test(stripped) || /^inviato da/i.test(stripped)) {
        inSignature = true;
      }

      if (inSignature) {
        continue;
      }

      if (
        /^On .* wrote:.*$/i.test(stripped) ||
        /^Il giorno .* ha scritto:.*$/i.test(stripped) ||
        /^Il .* alle .* ha scritto:.*$/i.test(stripped) ||
        /^-{3,}.*Messaggio originale.*$/i.test(stripped) ||
        /^-{3,}.*Original Message.*$/i.test(stripped)
      ) {
        inQuotedSection = true;
      }

      if (inQuotedSection) {
        continue;
      }

      if (/^>/.test(stripped)) {
        continue;
      }

      cleaned.push(stripped);
    }

    return cleaned.join(' ').replace(/\s+/g, ' ').trim();
  }

  /**
   * Ottiene suggerimento tipo richiesta per iniezione nel prompt
   * Supporta sia stringa pura che oggetto classificazione completo
   */
  getRequestTypeHint(classificationOrType) {
    const requestType = typeof classificationOrType === 'string'
      ? classificationOrType
      : (classificationOrType && classificationOrType.type ? classificationOrType.type : 'poem_submission');

    if (requestType === 'information_request') {
      return `
🎯 TIPO RICHIESTA RILEVATO: RICHIESTA INFORMAZIONI
────────────────────────────────────────────────────────
Linee guida per la risposta:
- Rispondi in modo CHIARO e BREVE alle domande
- Fornisci informazioni sul regolamento e modalità dalla Knowledge Base
- Tono professionale ma cordiale
────────────────────────────────────────────────────────`;
    } else if (requestType === 'parish_info_request') {
      return `
🎯 TIPO RICHIESTA RILEVATO: RICHIESTA INFORMAZIONI PARROCCHIALI (NON PERTINENTE)
────────────────────────────────────────────────────────
Linee guida per la risposta:
- Spiega cortesemente che questa casella di posta è dedicata esclusivamente all'iniziativa "Uno spazio per la poesia".
- Per richieste relative a orari delle Sante Messe, pellegrinaggi, incontri, catechesi e altre attività parrocchiali, ti chiediamo di invitare CHIARAMENTE l'utente a scrivere all'indirizzo email: info@parrocchiasanteugenio.it
- Tono: formale, cortese e indirizzante.
────────────────────────────────────────────────────────`;
    } else {
      return `
🎯 TIPO RICHIESTA RILEVATO: INVIO POESIA
────────────────────────────────────────────────────────
Linee guida per la risposta:
- Accogli caldamente il contributo
- Ringrazia per la partecipazione
- Leggi attentamente il componimento (o testo OCR) ed esprimi un breve e sincero apprezzamento letterario pertinente ai versi
- Fornisci (se utile) promemoria amichevole sulla pubblicazione settimanale
────────────────────────────────────────────────────────`;
    }
  }
}

// Funzione factory
function createRequestTypeClassifier() {
  return new RequestTypeClassifier();
}
