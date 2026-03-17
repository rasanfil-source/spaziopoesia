/**
 * gas_error_types.js - Classificazione centralizzata errori API
 *
 * Fornisce ErrorTypes e classifyError() per categorizzare
 * gli errori in modo uniforme in tutto il sistema.
 * Usato da GeminiService._withRetry e dai test.
 */

const ErrorTypes = {
    QUOTA_EXCEEDED: 'QUOTA_EXCEEDED',
    INVALID_API_KEY: 'INVALID_API_KEY',
    TIMEOUT: 'TIMEOUT',
    INVALID_RESPONSE: 'INVALID_RESPONSE',
    NETWORK: 'NETWORK',
    UNKNOWN: 'UNKNOWN'
};

/**
 * Classifica un errore in una categoria standard.
 * @param {Error|string} error - Errore da classificare
 * @returns {{ type: string, retryable: boolean, message: string }}
 */
function classifyError(error) {
    const message = String((error != null && error.message != null ? error.message : error) || '').toLowerCase();

    // I messaggi 5xx possono contenere la parola "quota" (es. "Errore rete/server o quota (503)").
    // Manteniamo priorità alla classificazione NETWORK per evitare falsi positivi QUOTA_EXCEEDED.
    if (message.includes('rete/server') || message.includes('network') ||
        message.includes('500') || message.includes('502') ||
        message.includes('503') || message.includes('504') ||
        message.includes('service unavailable')) {
        return { type: ErrorTypes.NETWORK, retryable: true, message: message };
    }

    if (message.includes('quota') || message.includes('rate limit') ||
        message.includes('429') || message.includes('resource_exhausted')) {
        return { type: ErrorTypes.QUOTA_EXCEEDED, retryable: true, message: message };
    }

    if (message.includes('api key') || message.includes('unauthorized') ||
        message.includes('unauthenticated') || message.includes('permission_denied') ||
        message.includes('403')) {
        return { type: ErrorTypes.INVALID_API_KEY, retryable: false, message: message };
    }

    if (message.includes('timeout') || message.includes('deadline exceeded') ||
        message.includes('econnreset') || message.includes('econnaborted') ||
        message.includes('408') || message.includes('request timed out')) {
        return { type: ErrorTypes.TIMEOUT, retryable: true, message: message };
    }

    if (message.includes('invalid') || message.includes('malformed') ||
        message.includes('invalid_argument') || message.includes('non json valida')) {
        return { type: ErrorTypes.INVALID_RESPONSE, retryable: false, message: message };
    }

    return { type: ErrorTypes.UNKNOWN, retryable: false, message: message };
}
