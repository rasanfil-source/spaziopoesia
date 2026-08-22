/**
 * GmailService.gs - Gestione operazioni Gmail
 * 
 * FUNZIONALITÀ:
 * - Label cache per performance
 * - Supporto header Reply-To per form web
 * - Costruttore cronologia conversazione
 * - Rimozione citazioni/firme
 * - Threading corretto (In-Reply-To, References)
 * - Markdown to HTML
 */

class GmailService {
    constructor() {
        console.log('📧 Inizializzazione GmailService...');

        // Cache etichette: in-memory (stessa esecuzione) + CacheService (cross-esecuzione)
        this._labelCache = new Map();
        this._cacheTTL = (typeof CONFIG !== 'undefined' && CONFIG.GMAIL_LABEL_CACHE_TTL) ? CONFIG.GMAIL_LABEL_CACHE_TTL : 3600000;
        this._cacheTtlSeconds = Math.max(60, Math.floor(this._cacheTTL / 1000));
        this._scriptCache = CacheService.getScriptCache();

        console.log('✓ GmailService inizializzato con cache etichette (TTL 1h)');
    }

    // ========================================================================
    // GESTIONE ETICHETTE (con cache)
    // ========================================================================

    /**
     * Ottiene o crea un'etichetta Gmail con caching
     */
    getOrCreateLabel(labelName) {
        const cacheKey = `gmail_label_exists:${labelName}`;
        const cachedEntry = this._labelCache.get(labelName);
        const now = Date.now();
        if (cachedEntry && (now - cachedEntry.ts) < this._cacheTTL) {
            console.log(`📦 Label '${labelName}' trovata in cache`);
            return cachedEntry.label;
        } else if (cachedEntry) {
            this._labelCache.delete(labelName);
        }

        const cachedExists = this._scriptCache.get(cacheKey);
        if (cachedExists) {
            const label = GmailApp.getUserLabelByName(labelName);
            if (label) {
                this._labelCache.set(labelName, { label: label, ts: now });
                console.log(`📦 Label '${labelName}' trovata in cache persistente`);
                return label;
            }
            this._scriptCache.remove(cacheKey);
        }

        const labels = GmailApp.getUserLabels();
        for (const label of labels) {
            if (label.getName() === labelName) {
                this._labelCache.set(labelName, { label: label, ts: now });
                this._scriptCache.put(cacheKey, '1', this._cacheTtlSeconds);
                console.log(`✓ Label '${labelName}' trovata`);
                return label;
            }
        }

        const newLabel = GmailApp.createLabel(labelName);
        this._labelCache.set(labelName, { label: newLabel, ts: now });
        this._scriptCache.put(cacheKey, '1', this._cacheTtlSeconds);
        console.log(`✓ Creata nuova label: ${labelName}`);
        return newLabel;
    }

    clearLabelCache() {
        this._labelCache.clear();
        console.log('🗑️ Cache label svuotata');
    }

    _clearPersistentLabelCache(labelName) {
        if (!labelName) return;
        this._scriptCache.remove(`gmail_label_exists:${labelName}`);
    }

    addLabelToThread(thread, labelName) {
        try {
            const label = this.getOrCreateLabel(labelName);
            thread.addLabel(label);
            console.log(`✓ Aggiunta label '${labelName}' al thread`);
        } catch (e) {
            console.warn(`⚠️ addLabelToThread fallito per '${labelName}': ${e.message}`);
            if (this._isLabelNotFoundError(e)) {
                this._clearPersistentLabelCache(labelName);
                this.clearLabelCache();
                const label = this.getOrCreateLabel(labelName);
                thread.addLabel(label);
                console.log(`✓ Aggiunta label '${labelName}' al thread (retry dopo cache reset)`);
                return;
            }

            // Non nascondere errori non correlati alla cache etichette (permessi, quota, thread invalido...)
            throw e;
        }
    }

    /**
     * Aggiunge etichetta a un messaggio specifico (Gmail API avanzata)
     */
    addLabelToMessage(messageId, labelName) {
        try {
            const label = this.getOrCreateLabel(labelName);
            const labelId = label.getId();
            Gmail.Users.Messages.modify({
                addLabelIds: [labelId],
                removeLabelIds: []
            }, 'me', messageId);
            console.log(`✓ Aggiunta label '${labelName}' al messaggio ${messageId}`);
        } catch (e) {
            console.warn(`⚠️ addLabelToMessage fallito per messaggio ${messageId}: ${e.message}`);
            if (this._isLabelNotFoundError(e)) {
                this._clearPersistentLabelCache(labelName);
                this.clearLabelCache();
                try {
                    const label = this.getOrCreateLabel(labelName);
                    const labelId = label.getId();
                    Gmail.Users.Messages.modify({
                        addLabelIds: [labelId],
                        removeLabelIds: []
                    }, 'me', messageId);
                    console.log(`✓ Aggiunta label '${labelName}' al messaggio ${messageId} (retry dopo cache reset)`);
                } catch (retryError) {
                    console.warn(`⚠️ Retry addLabelToMessage fallito per messaggio ${messageId}: ${retryError.message}`);
                    throw retryError;
                }
                return;
            }
            throw e;
        }
    }

    _isLabelNotFoundError(error) {
        const message = (error && error.message) ? error.message.toLowerCase() : '';
        return (message.includes('label') && message.includes('not found')) ||
            message.includes('etichetta non trovata') ||
            message.includes('invalid label') ||
            message.includes('404');
    }

    /**
     * Ottiene gli ID di tutti i messaggi con una specifica etichetta
     */
    getMessageIdsWithLabel(labelName, onlyInbox = true, options = {}) {
        try {
            const label = this.getOrCreateLabel(labelName);
            const labelId = label ? label.getId() : null;

            const messageIds = new Set();
            let pageToken;

            const safeWindowDays = parseInt(options.windowDays, 10);
            const useWindowDays = Number.isFinite(safeWindowDays) && safeWindowDays > 0
                ? safeWindowDays
                : ((typeof CONFIG !== 'undefined' && CONFIG.GMAIL_LABEL_LOOKBACK_DAYS) || 0);
            const maxPages = this._safePositiveInt(
                options.maxPages,
                ((typeof CONFIG !== 'undefined' && CONFIG.GMAIL_LIST_MAX_PAGES) || 20),
                1
            );
            const maxMessages = this._safePositiveInt(
                options.maxMessages,
                ((typeof CONFIG !== 'undefined' && CONFIG.GMAIL_LIST_MAX_MESSAGES) || 5000),
                1
            );
            const pageSize = this._safePositiveInt(options.pageSize, 500, 50, 500);

            // Query composita: inbox opzionale + finestra temporale opzionale
            const queryParts = [];
            if (onlyInbox) queryParts.push('in:inbox');
            if (useWindowDays > 0) queryParts.push(`after:${this._getNDaysAgo(useWindowDays)}`);
            const query = queryParts.join(' ').trim();
            let pageCount = 0;

            do {
                if (pageCount >= maxPages || messageIds.size >= maxMessages) {
                    console.warn(`⚠️ Interruzione list label '${labelName}': limite raggiunto (pages=${pageCount}/${maxPages}, messages=${messageIds.size}/${maxMessages})`);
                    break;
                }

                const response = Gmail.Users.Messages.list('me', {
                    labelIds: [labelId],
                    q: query,
                    maxResults: pageSize,
                    pageToken: pageToken
                });
                pageCount++;

                if (response.messages) {
                    for (const m of response.messages) {
                        messageIds.add(m.id);
                        if (messageIds.size >= maxMessages) {
                            break;
                        }
                    }
                }

                pageToken = response.nextPageToken;
            } while (pageToken);

            console.log(`📦 Trovati ${messageIds.size} messaggi con label '${labelName}' (inbox: ${onlyInbox}, windowDays: ${useWindowDays || 'all'}, pages: ${pageCount})`);
            return messageIds;
        } catch (e) {
            console.warn(`⚠️ Impossibile ottenere messaggi con label ${labelName}: ${e.message}`);
            if (options && options.throwOnError) {
                throw e;
            }
            return new Set();
        }
    }

    _safePositiveInt(value, fallback, min, max = null) {
        const parsed = parseInt(value, 10);
        const fallbackParsed = parseInt(fallback, 10);
        let safe = Number.isFinite(parsed) ? parsed : (Number.isFinite(fallbackParsed) ? fallbackParsed : min);

        safe = Math.max(min, safe);
        if (max !== null) {
            safe = Math.min(max, safe);
        }

        return safe;
    }

    _getNDaysAgo(n) {
        const days = Math.max(0, parseInt(n, 10) || 0);
        const d = new Date();
        d.setDate(d.getDate() - days);
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    }

    // ========================================================================
    // ESTRAZIONE MESSAGGI (con supporto Reply-To)
    // ========================================================================

    /**
     * Estrae dettagli messaggio con supporto Reply-To e threading
     */
    extractMessageDetails(message) {
        const subject = message.getSubject();
        const sender = message.getFrom();
        const date = message.getDate();
        const body = message.getPlainBody() || this._htmlToPlainText(message.getBody());
        const messageId = message.getId();

        // Estrai RFC 2822 Message-ID e header utili per filtraggio
        let rfc2822MessageId = null;
        let existingReferences = null;
        let isNewsletter = false;
        const headers = {};
        try {
            const rawMessage = Gmail.Users.Messages.get('me', messageId, {
                format: 'metadata',
                metadataHeaders: [
                    'Message-ID',
                    'References',
                    'Auto-Submitted',
                    'Precedence',
                    'X-Autoreply',
                    'X-Auto-Response-Suppress',
                    'Reply-To',
                    'List-Unsubscribe'
                ]
            });
            if (rawMessage && rawMessage.payload && rawMessage.payload.headers) {
                for (const header of rawMessage.payload.headers) {
                    if (header && header.name) {
                        headers[header.name.toLowerCase()] = header.value || '';
                    }
                    if (header.name === 'Message-ID' || header.name === 'Message-Id') {
                        rfc2822MessageId = header.value;
                    }
                    if (header.name === 'References') {
                        existingReferences = header.value;
                    }
                }
            }

            // Calcolo flag newsletter basato su header raccolti
            if (
                headers['list-unsubscribe'] ||
                /bulk|list/i.test(headers['precedence'] || '') ||
                /auto-generated|auto-replied/i.test(headers['auto-submitted'] || '')
            ) {
                isNewsletter = true;
            }
        } catch (e) {
            console.warn(`⚠️ Impossibile estrarre RFC 2822 Message-ID: ${e.message}`);
        }

        const replyTo = message.getReplyTo();

        let effectiveSender;
        let hasReplyTo = false;

        if (replyTo && replyTo.includes('@') && replyTo !== sender) {
            effectiveSender = replyTo;
            hasReplyTo = true;
            console.log(`   📧 Uso Reply-To: ${replyTo} (From originale: ${sender})`);
        } else {
            effectiveSender = sender;
        }

        const senderName = this._extractSenderName(effectiveSender);
        const senderEmail = this._extractEmailAddress(effectiveSender);

        let recipientEmail = null;
        try {
            recipientEmail = message.getTo();
        } catch (e) {
            const effectiveUser = Session.getEffectiveUser();
            recipientEmail = effectiveUser ? effectiveUser.getEmail() : '';
            if (!recipientEmail) {
                // Nota: Session.getActiveUser() in questo contesto GAS potrebbe restituire stringa vuota 
                // se non ci sono permessi specifici o se è un trigger.
                const activeUser = Session.getActiveUser();
                recipientEmail = activeUser ? activeUser.getEmail() : '';
            }
        }

        let recipientCc = '';
        try {
            recipientCc = message.getCc() || '';
        } catch (e) {
            recipientCc = '';
        }

        return {
            id: messageId,
            subject: subject,
            sender: effectiveSender,
            senderName: senderName,
            senderEmail: senderEmail,
            date: date,
            body: body,
            originalFrom: sender,
            hasReplyTo: hasReplyTo,
            rfc2822MessageId: rfc2822MessageId,
            existingReferences: existingReferences,
            recipientEmail: recipientEmail,
            recipientCc: recipientCc,
            headers: headers,
            isNewsletter: isNewsletter
        };
    }

    // ========================================================================
    // ALLEGATI: GESTIONE MULTIMODALE (Gemini Vision)
    // ========================================================================

    /**
     * Estrae gli allegati processabili in modalità multimodale.
     * - TXT/CSV: estratti come testo di contesto
     * - PDF/Immagini: passati come Blob per inlineData Gemini
     * - DOC/DOCX/PPT/PPTX: convertiti al volo in PDF
     * @param {GmailMessage} message
     * @param {object} options
     * @returns {{textContext: string, blobs: Array<Blob>, skipped: Array}}
     */
    getProcessableAttachments(message, options = {}) {
        const defaults = (typeof CONFIG !== 'undefined' && CONFIG.ATTACHMENT_CONTEXT)
            ? CONFIG.ATTACHMENT_CONTEXT
            : {};
        const settings = Object.assign({
            maxFiles: 4,
            maxBytesPerFile: 5 * 1024 * 1024
        }, defaults, options);

        const result = {
            textContext: '',
            blobs: [],
            skipped: []
        };

        let attachments = [];
        try {
            attachments = message.getAttachments({ includeInlineImages: true, includeAttachments: true }) || [];
        } catch (e) {
            console.warn(`⚠️ Impossibile leggere allegati: ${e.message}`);
            result.skipped.push({ reason: 'read_error', error: e.message });
            return result;
        }

        if (attachments.length === 0) {
            return result;
        }
        console.log(`   📎 Allegati trovati: ${attachments.length}`);

        let filesConsidered = 0;
        for (const attachment of attachments) {
            const name = attachment.getName ? (attachment.getName() || 'allegato') : 'allegato';

            if (this.isLikelySignatureImage(attachment)) {
                result.skipped.push({ name: name, reason: 'likely_signature_image' });
                console.log(`   🖋️ Immagine firma ignorata: ${name}`);
                continue;
            }

            if (filesConsidered >= settings.maxFiles) {
                result.skipped.push({ name: name, reason: 'max_files' });
                continue;
            }
            filesConsidered++;

            const size = attachment.getSize ? attachment.getSize() : 0;
            if (size > settings.maxBytesPerFile) {
                result.skipped.push({ name: name, reason: 'too_large', size: size });
                continue;
            }

            const mimeType = (attachment.getContentType() || '').toLowerCase();

            // TXT/CSV → estrazione testuale diretta
            if (mimeType.includes('text/plain') || mimeType.includes('text/csv')) {
                try {
                    const text = attachment.getDataAsString();
                    result.textContext += `\n\n--- Contenuto file: ${name} ---\n${text.substring(0, 5000)}`;
                    console.log(`   📄 Testo estratto da: ${name} (${text.length} caratteri)`);
                } catch (e) {
                    result.skipped.push({ name: name, reason: 'text_extract_error', error: e.message });
                }
                continue;
            }

            // PDF/Immagini → Blob diretto per Gemini Vision inlineData
            if (mimeType.startsWith('image/') || mimeType === 'application/pdf') {
                result.blobs.push(attachment.copyBlob());
                console.log(`   🖼️ Blob multimodale preparato: ${name} (${mimeType}, ${Math.round(size / 1024)}KB)`);
                continue;
            }

            // Office (Word/PowerPoint) → conversione al volo in PDF
            const isWord = mimeType.includes('msword') || mimeType.includes('wordprocessingml');
            const isPowerPoint = mimeType.includes('mspowerpoint') || mimeType.includes('presentationml');

            if (isWord || isPowerPoint) {
                try {
                    console.log(`   🔄 Conversione al volo in PDF per: ${name}`);
                    const convertedPdf = this._convertOfficeToPdf(attachment);
                    if (convertedPdf) {
                        convertedPdf.setName(`${name}.pdf`);
                        result.blobs.push(convertedPdf);
                        console.log(`   ✓ Convertito e pronto: ${name}.pdf`);
                    } else {
                        result.skipped.push({ name: name, reason: 'conversion_failed' });
                    }
                } catch (e) {
                    console.warn(`   ⚠️ Errore conversione per ${name}: ${e.message}`);
                    result.skipped.push({ name: name, reason: 'conversion_error', error: e.message });
                }
                continue;
            }

            result.skipped.push({ name: name, reason: 'unsupported_type', mimeType: mimeType });
        }

        if (result.blobs.length > 0) {
            console.log(`   📎 Pronti ${result.blobs.length} allegati visivi per Gemini Vision`);
        }

        return result;
    }

    /**
     * Euristica conservativa per distinguere loghi/firme da immagini significative.
     * Un'immagine viene esclusa solo con segnali forti; in caso di dubbio passa a Vision.
     */
    isLikelySignatureImage(attachment) {
        if (!attachment) return false;

        let mimeType = '';
        let name = '';
        let size = 0;
        try {
            mimeType = String(attachment.getContentType ? attachment.getContentType() : '').toLowerCase();
            name = String(attachment.getName ? attachment.getName() : '').toLowerCase();
            size = Number(attachment.getSize ? attachment.getSize() : 0) || 0;
        } catch (e) {
            return false;
        }

        if (!mimeType.startsWith('image/')) return false;

        const dimensions = this._getImageDimensions_(attachment, mimeType);
        const width = dimensions ? dimensions.width : 0;
        const height = dimensions ? dimensions.height : 0;
        const maxSide = Math.max(width, height);
        const minSide = Math.min(width, height);
        const aspectRatio = minSide > 0 ? maxSide / minSide : 0;

        // Tracking pixel o spacer quasi certamente non è contenuto inviato dall'utente.
        if (width > 0 && height > 0 && width <= 16 && height <= 16) return true;

        const strongNamePattern = /(^|[_.\-\s])(firma|signature|sig|logo|logotipo|footer|spacer|tracking|pixel|social|facebook|instagram|linkedin|twitter|youtube|tiktok|icon|icona)([_.\-\s]|$)/i;
        const hasSignatureName = strongNamePattern.test(name);

        // Nome esplicito + file contenuto: firma/logo con elevata probabilità.
        if (hasSignatureName && size <= 300 * 1024 && (maxSide === 0 || maxSide <= 1200)) return true;

        if (width > 0 && height > 0) {
            const bannerLike = width <= 900 && height <= 250 && aspectRatio >= 3 && size <= 200 * 1024;
            if (bannerLike) return true;
        }

        return false;
    }

    /** Legge le dimensioni senza librerie esterne per PNG, GIF e JPEG. */
    _getImageDimensions_(attachment, mimeType) {
        let bytes;
        try {
            bytes = attachment.getBytes ? attachment.getBytes() : null;
        } catch (e) {
            return null;
        }
        if (!bytes || bytes.length < 10) return null;

        const u8 = index => {
            const value = bytes[index];
            return value < 0 ? value + 256 : value;
        };
        const u16be = index => (u8(index) << 8) | u8(index + 1);
        const u32be = index => ((u8(index) * 0x1000000) + (u8(index + 1) << 16) + (u8(index + 2) << 8) + u8(index + 3));

        if (mimeType === 'image/png' && bytes.length >= 24 && u8(0) === 137 && u8(1) === 80) {
            return { width: u32be(16), height: u32be(20) };
        }
        if (mimeType === 'image/gif' && bytes.length >= 10) {
            return {
                width: u8(6) | (u8(7) << 8),
                height: u8(8) | (u8(9) << 8)
            };
        }
        if ((mimeType === 'image/jpeg' || mimeType === 'image/jpg') && u8(0) === 0xFF && u8(1) === 0xD8) {
            let offset = 2;
            while (offset + 8 < bytes.length) {
                if (u8(offset) !== 0xFF) { offset++; continue; }
                const marker = u8(offset + 1);
                if ([0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF].includes(marker)) {
                    return { height: u16be(offset + 5), width: u16be(offset + 7) };
                }
                if (marker === 0xD8 || marker === 0xD9) { offset += 2; continue; }
                const segmentLength = u16be(offset + 2);
                if (segmentLength < 2) break;
                offset += 2 + segmentLength;
            }
        }
        return null;
    }

    /**
     * Converte un file Office in PDF usando Drive Advanced Service.
     * Crea un file temporaneo, lo esporta in PDF e lo cancella sempre.
     * @param {Blob} attachmentBlob
     * @returns {Blob}
     */
    _convertOfficeToPdf(attachmentBlob) {
        if (typeof Drive === 'undefined' || !Drive.Files) {
            throw new Error('Drive Advanced Service non abilitato. Attivare il servizio Drive nel progetto Apps Script.');
        }

        let fileId = null;
        try {
            const blob = attachmentBlob.copyBlob();
            const fileName = attachmentBlob.getName ? attachmentBlob.getName() : 'allegato';

            if (typeof Drive.Files.insert === 'function') {
                // Drive API v2
                const resource = {
                    title: `TEMP_CONV_${fileName}`,
                    mimeType: blob.getContentType()
                };
                const file = Drive.Files.insert(resource, blob, { convert: true });
                fileId = file && file.id ? file.id : null;
            } else if (typeof Drive.Files.create === 'function') {
                // Drive API v3
                const sourceMime = (blob.getContentType ? blob.getContentType() : '').toLowerCase();
                const isPowerPoint = sourceMime.includes('mspowerpoint') || sourceMime.includes('presentationml');
                const targetMime = isPowerPoint
                    ? 'application/vnd.google-apps.presentation'
                    : 'application/vnd.google-apps.document';
                const resource = {
                    name: `TEMP_CONV_${fileName}`,
                    mimeType: targetMime
                };
                const file = Drive.Files.create(resource, blob, {
                    fields: 'id,mimeType'
                });
                
                if (file && file.mimeType && file.mimeType !== targetMime) {
                    throw new Error(`Conversione Office non applicata (mimeType=${file.mimeType})`);
                }
                fileId = file && file.id ? file.id : null;
            } else {
                throw new Error('Drive.Files non espone metodi compatibili (insert/create)');
            }

            if (!fileId) {
                throw new Error('Conversione fallita: file temporaneo senza id.');
            }

            const pdfBlob = DriveApp.getFileById(fileId).getAs('application/pdf');
            return pdfBlob;
        } finally {
            if (fileId) {
                try {
                    if (typeof Drive.Files.remove === 'function') {
                        Drive.Files.remove(fileId);
                    } else if (typeof Drive.Files.delete === 'function') {
                        Drive.Files.delete(fileId);
                    }
                } catch (e) {
                    console.warn(`⚠️ Errore cancellazione file temporaneo ${fileId}: ${e.message}`);
                }
            }
        }
    }

    _extractSenderName(fromField) {
        if (!fromField || typeof fromField !== 'string') {
            return 'Utente';
        }

        const match = fromField.match(/^"?(.+?)"?\s*</);
        let name = null;

        if (match) {
            name = match[1].trim();
        } else {
            const email = this._extractEmailAddress(fromField);
            if (email) {
                name = email.split('@')[0];
            }
        }

        if (name) {
            return this._capitalizeName(name);
        }

        return 'Utente';
    }

    _capitalizeName(name) {
        if (!name) return name;

        return name
            .split(/[\s-]+/)
            .map(word => {
                if (word.length === 0) return word;
                return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
            })
            .join(' ');
    }

    _extractEmailAddress(fromField) {
        if (typeof fromField !== 'string') return '';

        const angleMatch = fromField.match(/<([^>]+@[^>]+)>/);
        if (angleMatch) {
            return angleMatch[1];
        }

        // Evita regex RFC5322 troppo complesse (rischio backtracking su input malevoli).
        // Header From di Gmail sono già sanificati: pattern snello e lineare è sufficiente.
        const safeFromField = fromField.length > 512 ? fromField.substring(0, 512) : fromField;
        const emailMatch = safeFromField.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
        if (emailMatch) {
            return emailMatch[0];
        }

        return '';
    }

    _htmlToPlainText(html) {
        if (!html) return '';

        let text = html;
        // Preserva separatori strutturali per evitare blocchi di testo illeggibili.
        // È intenzionale: evitare il "muro di testo" migliora la qualità del parsing Gemini.
        text = text.replace(/<br\s*\/?\s*>/gi, '\n');
        text = text.replace(/<\/p\s*>/gi, '\n\n');
        text = text.replace(/<\/div\s*>/gi, '\n');

        text = text.replace(/<[^>]+>/g, ' ');
        text = text.replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&#39;/g, "'");
        // Riduce spazi/tabs senza distruggere i newline significativi
        text = text
            .replace(/\r\n?/g, '\n')
            .replace(/[ \t]{2,}/g, ' ')
            .replace(/\n{3,}/g, '\n\n')
            .trim();

        return text;
    }

    // ========================================================================
    // CRONOLOGIA CONVERSAZIONE
    // ========================================================================

    /**
     * Costruisce cronologia conversazione da messaggi thread
     */
    buildConversationHistory(messages, maxMessages = 10, ourEmail = '') {
        if (!ourEmail) {
            ourEmail = Session.getEffectiveUser().getEmail();
        }

        if (messages.length > maxMessages) {
            console.warn(`⚠️ Thread con ${messages.length} messaggi, limitato a ultimi ${maxMessages}`);
            messages = messages.slice(-maxMessages);
        }

        const history = [];

        for (const msg of messages) {
            const details = this.extractMessageDetails(msg);
            const isOurs = ourEmail && details.senderEmail.toLowerCase() === ourEmail.toLowerCase();

            const prefix = isOurs ? 'Segreteria' : `Utente (${details.senderName})`;

            let body = details.body;
            if (body.length > 2000) {
                body = body.substring(0, 2000) + '\n[... messaggio troncato ...]';
            }

            history.push(`${prefix}: ${body}\n---`);
        }

        return history.join('\n');
    }

    // ========================================================================
    // RIMOZIONE CITAZIONI/FIRME
    // ========================================================================

    extractMainReply(content) {
        const markers = [
            /^>/m,
            /^On .* wrote:/m,
            /^Il giorno .* ha scritto:/m,
            /^-{3,}.*Original Message/m
        ];

        let result = content;

        for (const marker of markers) {
            const match = result.search(marker);
            if (match !== -1) {
                result = result.substring(0, match);
                break;
            }
        }

        const sigMarkers = [
            /^cordiali saluti\b/im,
            /^distinti saluti\b/im,
            /^in fede\b/im,
            /best regards/i,
            /sincerely/i,
            /sent from my iphone/i,
            /inviato da/i
        ];

        const signatureSearchStart = Math.floor(result.length * 0.7);
        const signatureTail = result.substring(signatureSearchStart);

        for (const marker of sigMarkers) {
            const match = signatureTail.search(marker);
            if (match === -1) continue;

            const absoluteMatch = signatureSearchStart + match;
            const prefix = result.substring(0, absoluteMatch);

            // Tronca solo se la firma è su una nuova sezione (dopo riga vuota)
            if (/\n\s*\n\s*$/.test(prefix) || absoluteMatch === 0) {
                result = result.substring(0, absoluteMatch);
                break;
            }
        }

        return result.trim();
    }

    // ========================================================================
    // INVIO RISPOSTA
    // ========================================================================

    sendReply(thread, replyText, messageDetails) {
        const gmailThread = typeof thread === 'string' ?
            GmailApp.getThreadById(thread) : thread;

        const messages = gmailThread.getMessages();
        const lastMsg = messages[messages.length - 1];
        lastMsg.reply(replyText);

        console.log(`✓ Risposta inviata a ${messageDetails.senderEmail}`);

        if (messageDetails.hasReplyTo) {
            console.log("   📧 Risposta inviata all'indirizzo Reply-To");
        }

        return true;
    }


    /**
     * Invia risposta come HTML con threading corretto
     */
    sendHtmlReply(resource, responseText, messageDetails) {
        const sanitizedText = this._sanitizeHeaders(responseText);

        let finalResponse = sanitizedText;

        // ⚠️ Ordine intenzionale: fixPunctuation PRIMA, applyReplacements DOPO.
        // Le sostituzioni dal foglio Sostituzioni hanno l'ultima parola,
        // così non vengono annullate da fixPunctuation (es. "Roma" → "roma").
        finalResponse = this.fixPunctuation(finalResponse, messageDetails.senderName);
        finalResponse = this.ensureGreetingLineBreak(finalResponse);

        if (typeof GLOBAL_CACHE !== 'undefined' && GLOBAL_CACHE.replacements) {
            const replacementCount = Object.keys(GLOBAL_CACHE.replacements).length;
            if (replacementCount > 0) {
                finalResponse = this.applyReplacements(finalResponse, GLOBAL_CACHE.replacements);
                console.log(`   ✓ Applicate ${replacementCount} regole sostituzione`);
            }
        }

        const htmlBody = (typeof markdownToHtml === 'function')
            ? markdownToHtml(finalResponse)
            : finalResponse
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/\n/g, '<br>');
        const plainText = this._htmlToPlainText(htmlBody);

        const hasThreadingInfo = messageDetails.rfc2822MessageId;

        if (hasThreadingInfo) {
            try {
                let threadId = null;
                if (typeof resource === 'string') {
                    threadId = resource;
                } else if (resource && typeof resource.getId === 'function') {
                    if (typeof resource.getThread === 'function') {
                        threadId = resource.getThread().getId();
                    } else {
                        threadId = resource.getId();
                    }
                }

                let replySubject = messageDetails.subject;
                if (!replySubject.toLowerCase().startsWith('re:')) {
                    replySubject = 'Re: ' + replySubject;
                }

                let referencesHeader = messageDetails.rfc2822MessageId;
                if (messageDetails.existingReferences) {
                    referencesHeader = messageDetails.existingReferences + ' ' + messageDetails.rfc2822MessageId;
                }

                // From stabile: usa sempre l'account attivo (evita errori "non autorizzato")
                const stableFrom = Session.getEffectiveUser().getEmail();

                // Reply-To: usa alias solo se presente in To/Cc del messaggio originale
                let replyToEmail = null;
                const recipientHeaders = `${messageDetails.recipientEmail || ''},${messageDetails.recipientCc || ''}`;
                const emailRegex = /\b[A-Za-z0-9][A-Za-z0-9._%+-]{0,63}@(?!-)(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}\b/gi;
                const recipientAddresses = (recipientHeaders.match(emailRegex) || [])
                    .map(addr => addr.replace(/[\r\n]+/g, '').trim().toLowerCase());
                const knownAliases = (typeof CONFIG !== 'undefined' && Array.isArray(CONFIG.KNOWN_ALIASES))
                    ? CONFIG.KNOWN_ALIASES.map(alias => (alias || '').toLowerCase())
                    : [];

                const matchedAlias = recipientAddresses.find(addr => knownAliases.includes(addr));
                if (matchedAlias && matchedAlias !== stableFrom.toLowerCase()) {
                    replyToEmail = matchedAlias;
                }

                const boundary = 'boundary_' + Date.now() + '_' + Math.random().toString(36).substring(2, 15);
                const rawHeaders = [
                    'MIME-Version: 1.0',
                    `From: ${stableFrom}`,
                    `To: ${messageDetails.senderEmail}`,
                    `Subject: =?UTF-8?B?${Utilities.base64Encode(replySubject, Utilities.Charset.UTF_8)}?=`,
                    `In-Reply-To: ${messageDetails.rfc2822MessageId}`,
                    `References: ${referencesHeader}`,
                    `Content-Type: multipart/alternative; boundary="${boundary}"`
                ];

                if (replyToEmail) {
                    rawHeaders.splice(2, 0, `Reply-To: ${replyToEmail}`);
                }

                // Manteniamo `rawHeaders` come array e lo espandiamo nel payload MIME.
                const rawMessage = [
                    rawHeaders.join('\r\n'),
                    '',
                    `--${boundary}`,
                    'Content-Type: text/plain; charset=UTF-8',
                    'Content-Transfer-Encoding: base64',
                    '',
                    Utilities.base64Encode(plainText, Utilities.Charset.UTF_8),
                    '',
                    `--${boundary}`,
                    'Content-Type: text/html; charset=UTF-8',
                    'Content-Transfer-Encoding: base64',
                    '',
                    Utilities.base64Encode(htmlBody, Utilities.Charset.UTF_8),
                    '',
                    `--${boundary}--`,
                    ''
                ].join('\r\n');

                // Gmail API RAW richiede base64url RFC4648 senza padding finale '='.
                const encodedMessage = Utilities.base64EncodeWebSafe(rawMessage).replace(/=+$/, '');

                Gmail.Users.Messages.send({
                    raw: encodedMessage,
                    threadId: threadId
                }, 'me');

                console.log(`✓ Risposta HTML inviata via Gmail API a ${messageDetails.senderEmail}`);
                console.log(`   📧 Threading headers: In-Reply-To=${messageDetails.rfc2822MessageId.substring(0, 30)}...`);
                return;

            } catch (apiError) {
                console.warn(`⚠️ Gmail API fallita, ripiego su GmailApp: ${apiError.message}`);
            }
        }

        // Alternativa: metodo tradizionale
        // Nel fallback nativo prediligiamo il cast esplicito a GmailMessage (se disponibile)
        // affinché la libreria interna mantenga al meglio il riferimento al messaggio specifico
        const isMessage = resource && typeof resource.reply === 'function' && typeof resource.getThread === 'function';
        let mailEntity = null;

        if (isMessage) {
            mailEntity = resource;
        } else if (typeof resource === 'string') {
            const threadEntity = GmailApp.getThreadById(resource);
            const threadMessages = threadEntity ? threadEntity.getMessages() : [];
            mailEntity = threadMessages.length > 0 ? threadMessages[threadMessages.length - 1] : threadEntity;
        } else {
            mailEntity = resource;
        }

        if (!mailEntity || typeof mailEntity.reply !== 'function') {
            throw new Error('Entità Gmail non valida per reply() nel fallback HTML');
        }

        try {
            mailEntity.reply('', { htmlBody: htmlBody });
            console.log(`✓ Risposta HTML inviata a ${messageDetails.senderEmail} (metodo alternativo nativo)`);
        } catch (error) {
            console.error(`❌ Risposta fallita: ${error.message}`);
            try {
                mailEntity.reply(plainText || this._stripHtmlTags(finalResponse));
                console.log(`✓ Risposta plain text inviata a ${messageDetails.senderEmail} (alternativa)`);
            } catch (fallbackError) {
                console.error(`❌ CRITICO: Invio risposta alternativo fallito: ${fallbackError.message}`);
                const errorLabel = (typeof CONFIG !== 'undefined' && CONFIG.ERROR_LABEL_NAME) ? CONFIG.ERROR_LABEL_NAME : 'Errore';
                if (mailEntity) {
                    const targetThread = (typeof mailEntity.getThread === 'function')
                        ? mailEntity.getThread()
                        : mailEntity;

                    if (targetThread && typeof targetThread.getMessages === 'function') {
                        this.addLabelToThread(targetThread, errorLabel);
                    }
                }
                throw new Error(`Invio Gmail fallito sia in HTML sia in testo semplice: ${fallbackError.message}`);
            }
        }
    }

    _stripHtmlTags(text) {
        if (!text) return '';
        return text
            .replace(/<[^>]+>/g, ' ')
            .replace(/\*\*(.+?)\*\*/g, '$1')
            .replace(/\*(.+?)\*/g, '$1')
            .replace(/#{1,4}\s+/g, '')
            // Mantieni link leggibile: [Testo](URL) -> Testo (URL)
            .replace(/\[(.+?)\]\((.+?)\)/g, '$1 ($2)');
    }

    // ========================================================================
    // SAFEGUARD DI FORMATTAZIONE
    // ========================================================================

    /**
     * Corregge errori comuni di punteggiatura
     * Gestisce eccezioni per nomi doppi (es. "Maria Isabella")
     */
    fixPunctuation(text, senderName = '') {
        if (!text) return text;

        // Intenzionale: array locale ricreato a ogni chiamata, quindi la mutazione
        // serve solo ad ampliare le eccezioni per il messaggio corrente.
        // Include nomi propri (città, luoghi) che non devono MAI essere minuscolizzati.
        const exceptions = [
            'Don', 'Padre', 'Suor', 'Monsignor', 'Papa', 'Signore', 'Signora',
            'Roma', 'Milano', 'Napoli', 'Torino', 'Firenze', 'Genova', 'Bologna',
            'Venezia', 'Palermo', 'Catania', 'Bari', 'Verona', 'Padova',
            'Italia', 'Europa', 'Vaticano',
            'Dio', 'Gesù', 'Maria', 'Cristo', 'Chiesa', 'Eugenio'
        ];

        if (senderName) {
            const nameParts = senderName.split(/\s+/);
            for (const part of nameParts) {
                if (part && !exceptions.includes(part)) {
                    exceptions.push(part);
                }
            }
        }

        return text.replace(/,\s+([A-ZÀÈÉÌÒÙ])([a-zàèéìòù]*)/g, (match, firstLetter, rest, offset) => {
            const word = firstLetter + rest;

            if (exceptions.includes(word)) {
                return match;
            }

            const afterMatch = text.substring(offset + match.length);

            // Eccezione per virgola/punto successivo
            if (afterMatch.match(/^\s*[,.]/)) {
                return match;
            }

            // Eccezione per congiunzione "e" seguita da nome (es. "Maria e Giovanni,")
            if (afterMatch.match(/^\s+e\s+[A-ZÀÈÉÌÒÙ][a-zàèéìòù]*\s*[,.]/)) {
                return match;
            }

            // Euristica nomi doppi: se la parola è seguita da un'altra parola maiuscola,
            // probabilmente sono nomi propri (es. "Maria Isabella", "Gian Luca", "Carlo Alberto")
            if (afterMatch.match(/^\s+[A-ZÀÈÉÌÒÙ][a-zàèéìòù]+/)) {
                return match;
            }

            return `, ${firstLetter.toLowerCase()}${rest}`;
        });
    }

    ensureGreetingLineBreak(text) {
        if (!text) return text;

        const lines = text.split('\n');
        if (lines.length > 1) {
            const firstLine = lines[0].trim();
            if (/^(Buongiorno|Buonasera|Salve|Gentile|Egregio|Ciao)/i.test(firstLine)) {
                if (lines[1].trim() !== '') {
                    lines.splice(1, 0, '');
                    return lines.join('\n');
                }
            }
        }
        return text;
    }

    /**
     * Applica sostituzioni testo dal foglio Sostituzioni
     */
    applyReplacements(text, replacements) {
        if (!text || !replacements) return text;

        let result = text;
        let count = 0;

        for (const [bad, good] of Object.entries(replacements)) {
            const regex = new RegExp(bad.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
            const before = result;
            result = result.replace(regex, good);

            if (result !== before) {
                count++;
            }
        }

        if (count > 0) {
            console.log(`✓ Applicate ${count} sostituzioni`);
        }

        return result;
    }

    _sanitizeHeaders(text) {
        if (!text) return '';
        return text
            .replace(/(^|\n)(To|Cc|Bcc|From|Subject|Reply-To):/gi, '$1[$2]:')
            .replace(/\r\n|\r/g, '\n');
    }

    // ========================================================================
    // VERIFICA STATO
    // ========================================================================

    testConnection() {
        const results = {
            connectionOk: false,
            canListMessages: false,
            canCreateLabels: false,
            errors: []
        };

        try {
            const threads = GmailApp.search('is:unread', 0, 1);
            results.connectionOk = true;
            results.canListMessages = true;

            try {
                const testLabel = this.getOrCreateLabel('_TEST_LABEL_');
                results.canCreateLabels = true;

                try {
                    testLabel.deleteLabel();
                } catch (e) { }
            } catch (e) {
                results.errors.push(`Impossibile creare label: ${e.message}`);
            }

        } catch (e) {
            results.errors.push(`Errore connessione: ${e.message}`);
        }

        results.isHealthy = results.connectionOk && results.canListMessages;
        return results;
    }

    _detectDocumentType(fileName, text) {
        const source = `${fileName || ''}\n${text || ''}`.toLowerCase();
        const docPatterns = [
            { type: 'Modulo iscrizione cresima', patterns: ['cresima', 'confermazione'], minMatches: 1 },
            { type: 'Modulo iscrizione prima comunione/catechesi', patterns: ['prima comunione', 'catechesi', 'catechismo'], minMatches: 1 },
            { type: 'Modulo corso prematrimoniale', patterns: ['prematrimonial', 'fidanzati', 'matrimonio'], minMatches: 1 },
            { type: 'Certificato di battesimo', patterns: ['certificato', 'battesimo', 'battezz'], minMatches: 2 },
            { type: 'Certificato di cresima', patterns: ['certificato', 'cresima', 'confermazion'], minMatches: 2 },
            { type: 'Documento identità/passaporto', patterns: ["carta d'identit", "documento di identit", 'passaporto'], minMatches: 1 },
            { type: 'Tessera sanitaria/codice fiscale', patterns: ['tessera sanitaria', 'codice fiscale'], minMatches: 1 }
        ];

        for (const rule of docPatterns) {
            const matches = rule.patterns.reduce((acc, pattern) => acc + (source.includes(pattern) ? 1 : 0), 0);
            if (matches >= (rule.minMatches || rule.patterns.length)) {
                return rule.type;
            }
        }

        if (source.includes('certificato')) return 'Certificato (non specificato)';
        if (source.includes('modulo') || source.includes('iscrizione')) return 'Modulo parrocchiale';
        return 'Documento generico';
    }

    _extractDocumentFields(text, shouldMask = true) {
        const value = `${text || ''}`;
        if (!value) return [];

        const extract = [];
        const patterns = [
            { label: 'Nome e cognome', regex: /(?:nome\s*(?:e\s*cognome)?|cognome\s*e\s*nome)\s*[:\-]\s*([^\n;]{3,80})/i },
            { label: 'Data di nascita', regex: /(?:data\s*di\s*nascita|nato\/a\s*il)\s*[:\-]?\s*(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})/i },
            { label: 'Luogo di nascita', regex: /(?:luogo\s*di\s*nascita|nato\/a\s*a)\s*[:\-]\s*([^\n,;]{2,80})/i },
            { label: 'Codice fiscale', regex: /\b([A-Z]{6}[0-9]{2}[A-Z][0-9]{2}[A-Z][0-9]{3}[A-Z])\b/i },
            { label: 'Documento', regex: /(?:numero\s*(?:documento|doc\.)|n\.\s*documento)\s*[:\-]?\s*([A-Z0-9\-]{5,20})/i },
            { label: 'Contatto email', regex: /\b([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})\b/i },
            { label: 'Telefono', regex: /(?:tel(?:efono)?|cell(?:ulare)?)\s*[:\-]?\s*(\+?[0-9\s]{7,16})/i }
        ];

        for (const p of patterns) {
            const m = value.match(p.regex);
            if (!m || !m[1]) continue;
            const normalized = m[1].trim();
            extract.push(`${p.label}: ${shouldMask ? this._maskSensitiveValue(normalized) : normalized}`);
        }

        return extract.slice(0, 8);
    }

    _maskSensitiveValue(raw) {
        const value = `${raw || ''}`.trim();
        if (!value) return '';
        if (value.length <= 4) return '****';
        const visiblePrefix = value.slice(0, 2);
        const visibleSuffix = value.slice(-2);
        return `${visiblePrefix}${'*'.repeat(Math.max(4, value.length - 4))}${visibleSuffix}`;
    }
}

// Funzione factory
function createGmailService() {
    return new GmailService();
}

// ====================================================================
// MARKDOWN → HTML
// ====================================================================

/**
 * Sanitizzazione URL robusta con whitelist di protocolli
 */
function sanitizeUrl(url) {
    if (!url || typeof url !== 'string') return null;

    let decoded = url
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>');

    try {
        decoded = decodeURIComponent(decoded);
    } catch (e) {
        console.warn('⚠️ URL decode fallito, uso raw');
    }

    decoded = decoded.replace(/[\u0000-\u001F\u007F-\u009F]/g, '');
    const normalized = decoded.toLowerCase().trim();

    const FORBIDDEN_PROTOCOLS = /^\s*(javascript|vbscript|data|file):/i;
    const ALLOWED_PROTOCOLS = /^\s*(https?|mailto):/i;

    if (FORBIDDEN_PROTOCOLS.test(normalized)) {
        console.warn(`🛑 Bloccato protocollo pericoloso: ${decoded}`);
        return null;
    }

    if (!ALLOWED_PROTOCOLS.test(normalized)) {
        console.warn(`🛑 Bloccato protocollo non whitelisted: ${decoded}`);
        return null;
    }

    // SSRF: blocco IP interni, IPv6 loopback/link-local, IP decimali
    const INTERNAL_IP_PATTERN = /^(https?:\/\/)?(localhost|127\.|10\.|172\.(1[6-9]|2[0-9]|3[01])\.|192\.168\.|169\.254\.)/i;
    const DECIMAL_IP = /^https?:\/\/\d{8,10}(\/|$)/i;
    const USERINFO_BYPASS = /^https?:\/\/[^@]+@/i;

    // Blocca rappresentazioni numeriche alternative localhost (hex/octal/miste)
    // es: 0x7f000001, 0177.0.0.1, 0x7f.0.0.1
    const ALT_LOCALHOST_NUMERIC = /^https?:\/\/(?:0x[0-9a-f]+|0[0-7]+|\d+)(?::\d+)?(?:\/|$)/i;

    if (INTERNAL_IP_PATTERN.test(normalized) ||
        DECIMAL_IP.test(normalized) ||
        ALT_LOCALHOST_NUMERIC.test(normalized) ||
        USERINFO_BYPASS.test(normalized)) {
        console.warn(`🛑 Bloccato tentativo SSRF: ${decoded}`);
        return null;
    }

    // Validazione hostname post-parse per bloccare dotted-quad in notazione esadecimale/ottale
    // es: http://0x7f.0x0.0x0.0x1/
    try {
        const parseHostFromUrl = (value) => {
            const match = value.match(/^https?:\/\/(\[[^\]]+\]|[^\/?#:]+)/i);
            return match ? match[1] : '';
        };

        const host = String(parseHostFromUrl(decoded) || '').toLowerCase();
        const hostNoBrackets = host.replace(/^\[|\]$/g, '');
        const normalizedHost = hostNoBrackets.replace(/\.+$/, '');

        if (normalizedHost === 'localhost') {
            console.warn(`🛑 Bloccato tentativo SSRF localhost canonico: ${decoded}`);
            return null;
        }

        const isBlockedIpv6Host = (ipv6Host) => {
            if (!ipv6Host || !ipv6Host.includes(':')) return false;

            const normalizedIpv6 = ipv6Host.toLowerCase();
            // Block loopback and unspecified
            if (normalizedIpv6 === '::' || normalizedIpv6 === '::1') return true;
            // Block link-local
            if (normalizedIpv6.startsWith('fe80:')) return true;
            // Block unique-local (ULA)
            if (normalizedIpv6.startsWith('fc') || normalizedIpv6.startsWith('fd')) return true;

            return false;
        };

        if (isBlockedIpv6Host(hostNoBrackets)) {
            console.warn(`🛑 Bloccato tentativo SSRF IPv6 locale: ${decoded}`);
            return null;
        }
        const parts = normalizedHost.split('.');

        if (parts.length === 4) {
            const parsedOctets = parts.map(part => {
                if (/^0x[0-9a-f]+$/i.test(part)) return parseInt(part, 16);
                if (/^0[0-7]+$/.test(part)) return parseInt(part, 8);
                if (/^\d+$/.test(part)) return parseInt(part, 10);
                return NaN;
            });

            const isNumericHost = parsedOctets.every(v => Number.isInteger(v) && v >= 0 && v <= 255);
            const firstOctet = parsedOctets[0];
            const secondOctet = parsedOctets[1];

            const isLoopback = firstOctet === 127;
            const isPrivate10 = firstOctet === 10;
            const isPrivate172 = firstOctet === 172 && secondOctet >= 16 && secondOctet <= 31;
            const isPrivate192 = firstOctet === 192 && secondOctet === 168;
            const isLinkLocal = firstOctet === 169 && secondOctet === 254;
            const isZeroNet = firstOctet === 0;
            const isCgnat = firstOctet === 100 && secondOctet >= 64 && secondOctet <= 127;
            const isBenchmarkNet = firstOctet === 198 && (secondOctet === 18 || secondOctet === 19);

            if (isNumericHost && (isLoopback || isPrivate10 || isPrivate172 || isPrivate192 || isLinkLocal || isZeroNet || isCgnat || isBenchmarkNet)) {
                console.warn(`🛑 Bloccato tentativo SSRF hostname numerico: ${decoded}`);
                return null;
            }
        }

        // Blocca IPv4-mapped IPv6 verso reti locali/loopback
        // es: http://[::ffff:127.0.0.1]/, http://[::ffff:7f00:1]/
        const mappedPatterns = [
            /^::ffff:(.+)$/i,
            /^0:0:0:0:0:ffff:(.+)$/i,
            /^0000:0000:0000:0000:0000:ffff:(.+)$/i
        ];

        const mappedMatch = mappedPatterns
            .map(pattern => hostNoBrackets.match(pattern))
            .find(match => match && match[1]);

        if (mappedMatch && mappedMatch[1]) {
            const mapped = mappedMatch[1].replace(/^\[|\]$/g, '');
            let mappedOctets = null;

            if (/^\d+\.\d+\.\d+\.\d+$/.test(mapped)) {
                mappedOctets = mapped.split('.').map(v => parseInt(v, 10));
            } else if (/^[0-9a-f]{1,4}:[0-9a-f]{1,4}$/i.test(mapped)) {
                const [highHex, lowHex] = mapped.split(':');
                const high = parseInt(highHex, 16);
                const low = parseInt(lowHex, 16);
                mappedOctets = [
                    (high >> 8) & 0xff,
                    high & 0xff,
                    (low >> 8) & 0xff,
                    low & 0xff
                ];
            }

            if (mappedOctets && mappedOctets.length === 4 && mappedOctets.every(v => Number.isInteger(v) && v >= 0 && v <= 255)) {
                const firstOctet = mappedOctets[0];
                const secondOctet = mappedOctets[1];
                const isLoopback = firstOctet === 127;
                const isPrivate10 = firstOctet === 10;
                const isPrivate172 = firstOctet === 172 && secondOctet >= 16 && secondOctet <= 31;
                const isPrivate192 = firstOctet === 192 && secondOctet === 168;
                const isLinkLocal = firstOctet === 169 && secondOctet === 254;
                const isZeroNet = firstOctet === 0;
                const isCgnat = firstOctet === 100 && secondOctet >= 64 && secondOctet <= 127;
                const isBenchmarkNet = firstOctet === 198 && (secondOctet === 18 || secondOctet === 19);

                if (isLoopback || isPrivate10 || isPrivate172 || isPrivate192 || isLinkLocal || isZeroNet || isCgnat || isBenchmarkNet) {
                    console.warn(`🛑 Bloccato tentativo SSRF IPv4-mapped IPv6: ${decoded}`);
                    return null;
                }
            }
        }
    } catch (e) {
        console.warn(`⚠️ URL parse fallito, blocco prudenziale: ${decoded}`);
        return null;
    }

    return decoded
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/**
 * Escape HTML per prevenire XSS.
 * Applicato PRIMA delle trasformazioni markdown.
 */
function escapeHtml(text) {
    const value = (text === null || typeof text === 'undefined') ? '' : String(text);
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/**
 * Converte Markdown in HTML sicuro.
 * Strategia: escape-first, poi trasformazioni markdown.
 */
function markdownToHtml(text) {
    if (text === null || typeof text === 'undefined') return '';
    const inputText = (typeof text === 'string') ? text : String(text);
    const normalizedInputText = inputText.replace(/\r\n?/g, '\n');

    const replaceMarkdownLinks = (input, replacer) => {
        let result = '';
        let cursor = 0;

        while (cursor < input.length) {
            const openBracket = input.indexOf('[', cursor);
            if (openBracket === -1) {
                result += input.slice(cursor);
                break;
            }

            result += input.slice(cursor, openBracket);

            const closeBracket = input.indexOf(']', openBracket + 1);
            if (closeBracket === -1 || input[closeBracket + 1] !== '(') {
                result += input.slice(openBracket, closeBracket === -1 ? input.length : closeBracket + 1);
                cursor = closeBracket === -1 ? input.length : closeBracket + 1;
                continue;
            }

            const linkText = input.slice(openBracket + 1, closeBracket);
            let i = closeBracket + 2;
            let depth = 0;
            let foundClosingParen = false;

            while (i < input.length) {
                const ch = input[i];
                if (ch === '(') {
                    depth++;
                } else if (ch === ')') {
                    if (depth === 0) {
                        foundClosingParen = true;
                        break;
                    }
                    depth--;
                }
                i++;
            }

            if (!foundClosingParen) {
                result += input.slice(openBracket);
                break;
            }

            const url = input.slice(closeBracket + 2, i);
            result += replacer(linkText, url);
            cursor = i + 1;
        }

        return result;
    };

    // 1. Proteggi code blocks (prima dell'escape globale)
    const codeBlocks = [];
    let html = normalizedInputText.replace(/```[\s\S]*?```/g, (match) => {
        const sanitized = escapeHtml(match.replace(/```/g, '').trim());
        const token = `@@CODEBLOCK_PLACEHOLDER_${codeBlocks.length}_${Utilities.getUuid()}@@`;
        codeBlocks.push({ token: token, value: sanitized });
        return token;
    });

    // 2. Proteggi link markdown (prima dell'escape globale)
    const links = [];
    html = replaceMarkdownLinks(html, (linkText, url) => {
        const sanitizedUrl = sanitizeUrl(url);
        const escapedText = escapeHtml(linkText);
        const token = `@@LINK_PLACEHOLDER_${links.length}_${Utilities.getUuid()}@@`;
        if (sanitizedUrl) {
            const hrefSafe = sanitizedUrl.replace(/&(?!amp;|lt;|gt;|quot;|#39;)/g, '&amp;');
            links.push({ token: token, value: `<a href="${hrefSafe}" style="color:#660000; font-weight:bold;">${escapedText}</a>` });
        } else {
            console.warn(`⚠️ URL bloccato per sicurezza: ${url}`);
            links.push({ token: token, value: escapedText });
        }
        return token;
    });

    // 3. Escape globale (tutto il testo rimanente diventa sicuro)
    html = escapeHtml(html);

    // 4. Trasformazioni markdown su testo già escaped
    // Manteniamo una proporzione fissa tra corpo testo e titoli.
    const baseBodyFontPx = 20;
    const headingPx = {
        h4: Math.round(baseBodyFontPx * 1.00),
        h3: Math.round(baseBodyFontPx * 1.15),
        h2: Math.round(baseBodyFontPx * 1.30),
        h1: Math.round(baseBodyFontPx * 1.50)
    };

    // Headers
    html = html.replace(/^####\s+(.+)$/gm, `<p style="font-size:${headingPx.h4}px;font-weight:bold;margin:8px 0 4px;">$1</p>`);
    html = html.replace(/^###\s+(.+)$/gm, `<p style="font-size:${headingPx.h3}px;font-weight:bold;margin:10px 0 4px;">$1</p>`);
    html = html.replace(/^##\s+(.+)$/gm, `<p style="font-size:${headingPx.h2}px;font-weight:bold;margin:12px 0 4px;">$1</p>`);
    html = html.replace(/^#\s+(.+)$/gm, `<p style="font-size:${headingPx.h1}px;font-weight:bold;margin:14px 0 6px;">$1</p>`);

    // Bold / Italic (asterischi già escaped come testo, usiamo la versione escaped)
    // Nota: gli asterischi NON vengono escaped da escapeHtml(), quindi funzionano normalmente
    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/(?<!\*)\*(?!\*)(.+?)\*(?!\*)/g, '<em>$1</em>');

    // 5. Liste markdown (bullet e numerate) -> <ul>/ <ol> + <li>
    // Liste puntate (- item  oppure  * item all'inizio riga)
    // Raggruppa righe consecutive con lo stesso prefisso in un unico <ul>
    html = html.replace(/((?:^[ \t]*[-*][ \t]+.+(?:\n|$))+)/gm, (block) => {
        const items = block
            .split('\n')
            .filter(l => l.trim())
            .map(l => `<li>${l.replace(/^[ \t]*[-*][ \t]+/, '')}</li>`)
            .join('');
        return `<ul style="margin:6px 0;padding-left:20px;">${items}</ul>`;
    });

    // Liste numerate (1. item)
    html = html.replace(/((?:^[ \t]*\d+\.[ \t]+.+(?:\n|$))+)/gm, (block) => {
        const items = block
            .split('\n')
            .filter(l => l.trim())
            .map(l => `<li>${l.replace(/^[ \t]*\d+\.[ \t]+/, '')}</li>`)
            .join('');
        return `<ol style="margin:6px 0;padding-left:20px;">${items}</ol>`;
    });

    html = html.trim();

    // 7. Ripristina link e code blocks
    links.forEach((entry) => {
        html = html.split(entry.token).join(entry.value);
    });

    codeBlocks.forEach((entry) => {
        html = html.split(entry.token).join(
            `<pre style="background:#f4f4f4;padding:10px;border-radius:4px;font-family:monospace;">${entry.value}</pre>`
        );
    });

    // 8. Emoji to HTML entities
    html = Array.from(html).map(char => {
        const codePoint = char.codePointAt(0);
        if (codePoint > 0xFFFF) {
            return '&#' + codePoint + ';';
        }
        return char;
    }).join('');

    // 9. Costruzione paragrafi evitando nesting invalido di <p>
    // Inserisce separatori intorno ai blocchi per evitare casi tipo:
    // "Intro\n<ul>...</ul>" -> <p>Intro<br><ul>...</ul></p> (HTML invalido)
    html = html.replace(/(<\/?(?:ul|ol|pre|p|div|h[1-6])\b[^>]*>)/gi, '\n$1\n');

    const isBlockHtml = (fragment) => /^<(p|ul|ol|pre|div|h[1-6])\b/i.test(fragment.trim());
    const cleanedHtml = html
        .split(/\n\n+/)
        .map(fragment => fragment.trim())
        .filter(fragment => fragment.length > 0)
        .map(fragment => {
            const withLineBreaks = isBlockHtml(fragment)
                ? fragment
                : fragment.replace(/\n/g, '<br>');
            if (!withLineBreaks || withLineBreaks === '<br>') return withLineBreaks;
            return isBlockHtml(withLineBreaks) ? withLineBreaks : `<p>${withLineBreaks}</p>`;
        })
        .join('');

    const startsWithBlock = /^\s*<(p|ul|ol|pre|h[1-6])/i.test(cleanedHtml);
    const bodyContent = startsWithBlock ? cleanedHtml : `<p>${cleanedHtml}</p>`;

    // Manteniamo il corpo risposta a 20px come da stile originale.
    return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family: Arial, Helvetica, sans-serif; font-size: ${baseBodyFontPx}px; color: #660000; line-height: 1.6;">
  <div style="font-family: Arial, Helvetica, sans-serif; font-size: ${baseBodyFontPx}px; color: #660000; line-height: 1.6;">
    ${bodyContent}
  </div>
</body>
</html>`;
}
