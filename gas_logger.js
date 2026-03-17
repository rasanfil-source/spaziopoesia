/**
 * Logger.gs - Sistema di logging strutturato e centralizzato
 */

const LogLevel = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3
};

// Nota: non usare il nome `Logger` per evitare shadowing del built-in GAS `Logger.log()`.
class AppLogger {
  constructor(context = 'System') {
    this.context = context;
    this.config = typeof getConfig === 'function' ? getConfig() : (typeof CONFIG !== 'undefined' ? CONFIG : {});
    const levelStr = (this.config.LOGGING && this.config.LOGGING.LEVEL) || 'INFO';
    this.minLevel = Object.prototype.hasOwnProperty.call(LogLevel, levelStr)
      ? LogLevel[levelStr]
      : LogLevel.INFO;
  }

  /**
   * Log generico
   */
  _log(level, message, data = {}) {
    if (LogLevel[level] < this.minLevel) return;

    const logEntry = {
      timestamp: new Date().toISOString(),
      level: level,
      context: this.context,
      message: message,
      ...data
    };

    const loggingConfig = (this.config && this.config.LOGGING) ? this.config.LOGGING : {};

    if (loggingConfig.STRUCTURED) {
      console.log(JSON.stringify(logEntry));
    } else {
      console.log(`[${logEntry.timestamp}] [${level}] [${this.context}] ${message}`);
      if (Object.keys(data).length > 0) {
        console.log(JSON.stringify(data, null, 2));
      }
    }

    // Invia notifica per errori critici
    if (level === 'ERROR' && loggingConfig.SEND_ERROR_NOTIFICATIONS) {
      this._sendErrorNotification(logEntry);
    }
  }

  debug(message, data) {
    this._log('DEBUG', message, data);
  }

  info(message, data) {
    this._log('INFO', message, data);
  }

  warn(message, data) {
    this._log('WARN', message, data);
  }

  error(message, data) {
    this._log('ERROR', message, data);
  }

  /**
   * Log specifico per thread email
   */
  logThread(threadId, action, status, details = {}) {
    this.info(`Thread ${action}`, {
      threadId: threadId,
      action: action,
      status: status,
      ...details
    });
  }

  /**
   * Log metriche di esecuzione
   */
  logMetrics(metrics) {
    this.info('Metriche di esecuzione', metrics);
  }

  /**
   * Invia notifica email per errori critici
   */
  _sendErrorNotification(logEntry) {
    try {
      const loggingConfig = (this.config && this.config.LOGGING) ? this.config.LOGGING : {};
      const adminEmailProperty = (typeof PropertiesService !== 'undefined' && PropertiesService && typeof PropertiesService.getScriptProperties === 'function')
        ? PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')
        : '';
      const adminEmail = adminEmailProperty || loggingConfig.ADMIN_EMAIL || '';
      if (!adminEmail || adminEmail.includes('[') || adminEmail.includes('YOUR_')) return;

      const subject = `[${this.config.PROJECT_NAME || 'GAS_BOT'}] Avviso Errore: ${logEntry.message}`;
      const body = `
Errore nel sistema autoresponder:

Timestamp: ${logEntry.timestamp}
Context: ${logEntry.context}
Message: ${logEntry.message}

Dettagli:
${JSON.stringify(logEntry, null, 2)}

---
Sistema: ${this.config.PROJECT_NAME || 'GAS_BOT'}
Script ID: ${this.config.SCRIPT_ID || 'Unknown'}
      `.trim();

      // Rate limit: max 1 email ogni 5 minuti
      const lastNotification = PropertiesService.getScriptProperties()
        .getProperty('last_error_notification');
      const now = Date.now();

      if (!lastNotification || (now - parseInt(lastNotification)) > 300000) {
        GmailApp.sendEmail(adminEmail, subject, body);
        PropertiesService.getScriptProperties()
          .setProperty('last_error_notification', now.toString());
      }
    } catch (e) {
      console.error('Invio notifica errore fallito:', e.message);
    }
  }

  /**
   * Crea logger con contesto specifico
   */
  withContext(newContext) {
    return new AppLogger(`${this.context}:${newContext}`);
  }
}

/**
 * Factory function per creare logger
 */
function createLogger(context) {
  return new AppLogger(context);
}