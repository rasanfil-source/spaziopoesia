/**
 * PromptContext.gs - Calcola profilo e concern runtime
 * Determina dinamicamente quali regole del prompt sono attive
 */

class PromptContext {
    constructor(input) {
        // Validazione input
        if (!input || typeof input !== 'object') {
            console.error(`🚨 PromptContext: input non valido`);
            input = {};
        }

        // Sanitizza lastUpdated non valido.
        // Nota: il doppio controllo su input.memory è intenzionale per evitare TypeError
        // quando il chiamante non passa il blocco memoria.
        if (input.memory && input.memory.lastUpdated) {
            if (isNaN(new Date(input.memory.lastUpdated).getTime())) {
                console.warn(`⚠️ PromptContext: lastUpdated non valido, reset a null`);
                input.memory.lastUpdated = null;
            }
        }

        this.input = this._normalizeInput(input);
        this.concerns = this._computeConcerns();
        this.profile = this._computeProfile();
        this.meta = this._buildMeta();
    }

    _normalizeInput(input) {
        const normalizedInput = Object.assign({}, input);

        // Mantiene knowledgeBase originale inalterata e crea metadati separati.
        // NOTA: i metadati restano dentro input (knowledgeBaseMeta) per compatibilita con i test
        // e con il codice che legge this.input in debug/telemetria: non usiamo un campo esterno.
        if (normalizedInput.knowledgeBase) {
            const isString = typeof normalizedInput.knowledgeBase === 'string';
            let knowledgeBaseRaw = '';

            if (isString) {
                knowledgeBaseRaw = normalizedInput.knowledgeBase;
            } else {
                knowledgeBaseRaw = this._safeStringify(normalizedInput.knowledgeBase);
            }

            normalizedInput.knowledgeBaseRaw = knowledgeBaseRaw;
            normalizedInput.knowledgeBaseMeta = {
                length: knowledgeBaseRaw.length,
                containsDates: (isString && /\d{4}/.test(knowledgeBaseRaw)) || (typeof normalizedInput.knowledgeBase === 'object' && normalizedInput.knowledgeBase !== null)
            };
        }

        return normalizedInput;
    }

    /**
     * Serializzazione robusta della KB:
     * - gestisce riferimenti circolari senza lanciare eccezioni;
     * - evita fallback generici tipo "[object Object]" che perderebbero informazione.
     */
    _safeStringify(value) {
        const seen = new WeakSet();

        try {
            return JSON.stringify(value, (key, current) => {
                if (typeof current === 'object' && current !== null) {
                    if (seen.has(current)) {
                        return '[Circular]';
                    }
                    seen.add(current);
                }
                return current;
            });
        } catch (e) {
            console.warn('⚠️ PromptContext: knowledgeBase non serializzabile, uso fallback controllato');
            if (value === null || typeof value === 'undefined') {
                return '';
            }
            return typeof value === 'string' ? value : Object.prototype.toString.call(value);
        }
    }

    _computeConcerns() {
        const i = this.input;

        return {
            language_safety:
                i.email?.detectedLanguage !== 'it' ||
                (i.classification?.confidence ?? 1) < 0.8,

            hallucination_risk:
                (i.knowledgeBaseMeta?.length ?? i.knowledgeBase?.length ?? 0) > 800 ||
                i.temporal?.mentionsDates ||
                i.temporal?.mentionsTimes,

            formatting_risk:
                i.temporal?.mentionsTimes ||
                ['information', 'sacrament'].includes(i.classification?.category),

            temporal_risk:
                i.temporal?.mentionsDates ||
                i.knowledgeBaseMeta?.containsDates,

            discernment_risk:
                i.requestType?.needsDiscernment ||
                i.territory?.addressFound,

            emotional_sensitivity:
                i.requestType?.type === 'pastoral' ||
                i.classification?.subIntents?.emotional_distress ||
                i.classification?.subIntents?.bereavement,

            repetition_risk:
                i.memory?.exists ||
                (i.conversation?.messageCount ?? 0) > 1,

            identity_consistency:
                !i.email?.isReply ||
                i.requestType?.type !== 'technical',

            response_scope_control:
                i.email?.isReply ||
                (i.classification?.confidence ?? 1) < 0.7,

            salutation_control:
                i.salutationMode && i.salutationMode !== 'full'
        };
    }

    _computeProfile() {
        const c = this.concerns;

        if (c.discernment_risk || c.emotional_sensitivity) {
            return 'heavy';
        }

        if (c.hallucination_risk || c.formatting_risk || c.temporal_risk) {
            return 'standard';
        }

        return 'lite';
    }

    _buildMeta() {
        const active = Object.entries(this.concerns)
            .filter(([_, v]) => v)
            .map(([k]) => k);

        return {
            profile: this.profile,
            activeConcerns: active
        };
    }
}

// Factory function
function createPromptContext(input) {
    return new PromptContext(input);
}
