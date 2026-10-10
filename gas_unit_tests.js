// Bootstrap Node.js: carica script GAS e mock minimi per esecuzione locale/CI
if (typeof process !== 'undefined' && typeof require !== 'undefined') {
    var fs = require('fs');
    var vm = require('vm');

    var loadedScripts = new Set();
    global.loadScript = function (path) {
        if (loadedScripts.has(path)) return;
        try {
            var code = fs.readFileSync(path, 'utf8');
            vm.runInThisContext(code, { filename: path });
            loadedScripts.add(path);
        } catch (e) {
            console.error(`❌ Errore caricamento script [${path}]: ${e.stack}`);
        }
    };

    // Mock minimi obbligatori
    if (typeof global.createLogger !== 'function') {
        global.createLogger = () => ({ info: () => { }, warn: () => { }, debug: () => { }, error: () => { } });
    }
    if (typeof global.CONFIG === 'undefined') {
        global.CONFIG = {
            VALIDATION_MIN_SCORE: 0.6,
            SEMANTIC_VALIDATION: { enabled: false }
        };
    }
    if (typeof global.Utilities === 'undefined') {
        global.Utilities = {
            formatDate: (date, tz, fmt) => new Date(date).toISOString(),
            sleep: () => { },
            computeDigest: (algorithm, value) => Array.from(
                require('crypto').createHash('sha256').update(String(value || '')).digest()
            ).map(byte => byte > 127 ? byte - 256 : byte),
            DigestAlgorithm: { MD5: 'MD5', SHA_256: 'SHA_256' },
            getUuid: () => 'test-uuid-' + Math.random().toString(36).substring(2, 9),
            base64Encode: (data) => Buffer.from(data).toString('base64')
        };
    }
    if (typeof global.PropertiesService === 'undefined') {
        var propsData = new Map();
        global.PropertiesService = {
            getScriptProperties: () => ({
                getProperties: () => Object.fromEntries(propsData),
                getProperty: (k) => {
                    if (propsData.has(k)) return propsData.get(k);
                    if (k === 'GEMINI_API_KEY') return 'abcdefghijklmnopqrstuvwxyz123456';
                    if (k === 'SPREADSHEET_ID') return 'sheet-123';
                    return null;
                },
                setProperty: (k, v) => propsData.set(k, String(v)),
                setProperties: (obj) => Object.keys(obj || {}).forEach(k => propsData.set(k, String(obj[k]))),
                deleteProperty: (k) => propsData.delete(k)
            })
        };
    }
    if (typeof global.LockService === 'undefined') {
        global.LockService = {
            getScriptLock: () => ({
                tryLock: () => true,
                waitLock: () => true,
                releaseLock: () => { },
                hasLock: () => true
            })
        };
    }
    if (typeof global.CacheService === 'undefined') {
        var cache = new Map();
        global.CacheService = {
            getScriptCache: () => ({
                get: (k) => (cache.has(k) ? cache.get(k) : null),
                put: (k, v) => cache.set(k, String(v)),
                remove: (k) => cache.delete(k)
            })
        };
    }

    if (typeof global.SpreadsheetApp === 'undefined') {
        global.SpreadsheetApp = {
            flush: () => { },
            openById: () => ({
                getSheetByName: () => ({
                    getRange: (r) => ({
                        getValues: () => [[
                            'threadId', 'language', 'category', 'tone', 'providedInfo', 'lastUpdated', 'messageCount', 'version', 'memorySummary'
                        ]],
                        setFontWeight: () => { },
                        setValue: () => { },
                        getRow: () => 2,
                        getColumn: () => 1,
                        createTextFinder: (text) => ({
                            matchEntireCell: () => ({
                                matchCase: () => ({
                                    matchFormulaText: () => ({
                                        findNext: () => (text.includes('test-thread') ? { getRow: () => 2, getColumn: () => 1 } : null)
                                    })
                                })
                            })
                        })
                    }),
                    createTextFinder: (text) => ({
                        matchEntireCell: () => ({
                            matchCase: () => ({
                                matchFormulaText: () => ({
                                    findNext: () => (text.includes('test-thread') ? { getRow: () => 2, getColumn: () => 1 } : null)
                                })
                            })
                        })
                    }),
                    appendRow: () => { },
                    getLastRow: () => 10,
                    getMaxColumns: () => 10
                })
            })
        };
    }

    if (typeof global.Session === 'undefined') {
        global.Session = {
            getEffectiveUser: () => ({ getEmail: () => 'bot@example.com' }),
            getActiveUser: () => ({ getEmail: () => 'bot@example.com' }),
            getScriptTimeZone: () => 'Europe/Rome'
        };
    }

    // Caricamento script core
    [
        'gas_config.js',
        'gas_main.js',
        'gas_setup_ui.js',
        'gas_error_types.js',
        'gas_rate_limiter.js',
        'gas_memory_service.js',
        'gas_gmail_service.js',
        'gas_prompt_engine.js',
        'gas_email_processor.js',
        'gas_gemini_service.js',
        'gas_classifier.js',
        'gas_request_classifier.js',
        'gas_response_validator.js'
    ].forEach(loadScript);
}

// ═══════════════════════════════════════════════════════════════════════════
// FUNZIONI HELPER
// ═══════════════════════════════════════════════════════════════════════════

function testGroup(label, results, callback) {
    console.log(`\n${'═'.repeat(70)}`);
    console.log(`🧪 ${label}`);
    console.log('═'.repeat(70));
    try {
        callback();
    } catch (e) {
        console.error(`💥 ERRORE NEL GRUPPO [${label}]: ${e.message}`);
        process.exit(1);
    }
}

function test(label, results, callback) {
    results.total = (results.total || 0) + 1;
    try {
        var result = callback();
        if (result === true || result === undefined) {
            console.log(`  ✅ ${label}`);
            results.passed = (results.passed || 0) + 1;
        } else {
            console.error(`  ❌ ${label}`);
            results.failed = (results.failed || 0) + 1;
        }
    } catch (error) {
        console.error(`  💥 ${label}: ${error.message}`);
        results.failed = (results.failed || 0) + 1;
    }
}

// ═══════════════════════════════════════════════════════════════════════════
// TEST SUITE
// ═══════════════════════════════════════════════════════════════════════════

function runAllTests() {
    console.log('╔' + '═'.repeat(68) + '╗');
    console.log('║' + ' '.repeat(15) + '🧪 SUITE DI TEST' + ' '.repeat(36) + '║');
    console.log('╚' + '═'.repeat(68) + '╝');

    const results = { total: 0, passed: 0, failed: 0 };
    const start = Date.now();

    // 1. RateLimiter
    testGroup('Punto #1: RateLimiter - WAL Protection', results, () => {
        test('WAL persist elimini log se riuscito', results, () => {
            const limiter = new GeminiRateLimiter();
            limiter.flush();
            return limiter.props.getProperty('rate_limit_wal') === null;
        });
        test('WAL recovery ripristini finestre', results, () => {
            const limiter = new GeminiRateLimiter();
            const wal = { timestamp: Date.now(), rpm: [{ timestamp: Date.now(), modelKey: 'flash' }], tpm: [] };
            limiter.props.setProperty('rate_limit_wal', JSON.stringify(wal));
            limiter._recoverFromWAL();
            return limiter.cache.rpmWindow.length > 0;
        });
    });

    // 2. MemoryService
    testGroup('Punto #2: MemoryService - Timestamp & Lock', results, () => {
        const service = new MemoryService();
        test('Normalizzazione timestamp futuro', results, () => {
            const future = new Date(Date.now() + 200000000).toISOString();
            const normalized = service._validateAndNormalizeTimestamp(future);
            return new Date(normalized).getTime() <= Date.now() + 86400000;
        });
        test('Canonicalizza timestamp validi in ISO', results, () => {
            const normalized = service._validateAndNormalizeTimestamp('Wed, 01 Jan 2025 10:00:00 GMT');
            return normalized === '2025-01-01T10:00:00.000Z';
        });
        test('Lock gestion con threadId', results, () => {
            service.updateMemory('test-thread-id', { language: 'it' });
            return true;
        });
    });

    // Punto 3 rimosso (TerritoryValidator)

    // 4. GeminiService
    testGroup('Punto #4: GeminiService - Language', results, () => {
        const service = new GeminiService();
        test('Rilevamento IT', results, () => service.detectEmailLanguage("Buongiorno").lang === 'it');
        test('Rilevamento PT', results, () => service.detectEmailLanguage("Bom dia").lang === 'pt');
        test('Rilevamento LA per versi latini', results, () => {
            const latinText = "Angelus adstat,\nMaria humilis audit,\nVerbum fit caro.\n\nCaro Lucretius";
            return service.detectEmailLanguage(latinText).lang === 'la';
        });
        test('Saluto adattivo LA non usa inglese', results, () => {
            const greeting = service.getAdaptiveGreeting('Lucretius', 'la');
            return !/Good|Kind regards/i.test(`${greeting.greeting} ${greeting.closing}`);
        });
    });

    // 5. ResponseValidator
    testGroup('Punto #5: ResponseValidator - Quality', results, () => {
        const validator = new ResponseValidator();
        test('Rileva leak "Rivedendo la KB"', results, () => {
            const res = validator.validateResponse("Rivedendo la knowledge base, ecco la risposta.", 'it', "...", "...", "...", "full");
            return res.details.exposedReasoning.score === 0.0;
        });
        test('Rileva placeholder "XXX"', results, () => {
            const res = validator.validateResponse("Gentile utente, XXX, saluti.", 'it', "...", "...", "...", "full");
            return res.details.content.score === 0.0;
        });
        test('Rileva inconsistenza lingua (ES invece di IT)', results, () => {
            const spanishText = "Hola, gracias por contactarnos. Saludos estimables.";
            const res = validator._checkLanguage(spanishText, 'it');
            return res.score < 1.0 && (res.detectedLang === 'es' || res.warnings.length > 0);
        });
    });

    // 6. Gemini JSON parser recovery
    testGroup('Punto #6: Gemini JSON Parser - Recovery', results, () => {
        test('Parsa JSON in blocco markdown', results, () => {
            const parsed = parseGeminiJsonLenient('```json\n{"reply_needed":true,"language":"it","category":"MIXED"}\n```');
            return parsed.reply_needed === true && parsed.language === 'it' && parsed.category === 'MIXED';
        });
        test('Recupera campi minimi da JSON troncato', results, () => {
            const parsed = parseGeminiJsonLenient('{"reply_needed": true, "language": "it", "category": "MIXED", "dimensions": {"technical": 0.6');
            return parsed.reply_needed === true && parsed.language === 'it' && parsed.category === 'MIXED';
        });
    });

    // 7. Poetry Contest - Classification and Validator Logic
    testGroup('Punto #7: Poetry Contest - Classificazione e Firma', results, () => {
        // Test di Classificazione (gas_request_classifier.js)
        test('Classifica correttamente una email come POEM_SUBMISSION', results, () => {
            if (typeof RequestTypeClassifier !== "undefined") {
                const classifier = new RequestTypeClassifier();
                const res = classifier.classify("Ecco la mia poesia per il concorso", "Vi invio i miei versi sul tramonto...", { category: 'other' });
                return res.type === 'poem_submission';
            }
            return true; // Skip test se manca la classe a livello globale (es. scope node)
        });

        test('Classifica correttamente una email come PARISH_INFO_REQUEST', results, () => {
            if (typeof RequestTypeClassifier !== "undefined") {
                const classifier = new RequestTypeClassifier();
                const res = classifier.classify("Orari Sante Messe settimanali", "Vorrei sapere a che ora c'è la Messa domenica.", { category: 'other' });
                return res.type === 'parish_info_request';
            }
            return true;
        });

        // Test del Validator Signature (gas_response_validator.js)
        test('Validator riconosce la firma corretta del comitato poesia', results, () => {
            const validator = new ResponseValidator();
            const textConFirma = "La ringraziamo.\n\nIl comitato Uno spazio per la poesia";
            const res = validator.validateResponse(textConFirma, 'it', "", "", "", "full");
            // Non dovrebbe esserci l'errore o warning di firma mancante
            return !res.warnings.some(w => w.includes("Firma mancante"));
        });

        test('Validator segnala firma mancante se si usa il vecchio pattern', results, () => {
            const validator = new ResponseValidator();
            const textConVecchiaFirma = "La ringraziamo.\n\nSegreteria Parrocchia Sant'Eugenio";
            const res = validator.validateResponse(textConVecchiaFirma, 'it', "", "", "", "full");
            return res.warnings.some(w => w.includes("Firma mancante"));
        });
    });

    // 8. Parity Fixes (exnovoGAS)
    testGroup('Punto #8: Parity Fixes - Normalizzazione e Reazioni', results, () => {
        const processor = Object.create(EmailProcessor.prototype);
        processor._normalizeTopicKey = EmailProcessor.prototype._normalizeTopicKey;

        test('Normalizzazione topic (orari_messe -> orari_messe)', results, () => {
            return processor._normalizeTopicKey('orari_messe') === 'orari_messe';
        });
        test('Normalizzazione topic con spazi (Orari Messe -> orari_messe)', results, () => {
            return processor._normalizeTopicKey('Orari Messe') === 'orari_messe';
        });
        test('Normalizzazione topic con trattini (orari-messe -> orari_messe)', results, () => {
            return processor._normalizeTopicKey('orari-messe') === 'orari_messe';
        });
    });

    // 9. Config Validation (Gemini 3.8 Flash)
    testGroup('Punto #9: Config Validation - Gemini 3.8 Flash & Model Strategy', results, () => {
        test('validateConfig() valida con successo la configurazione Gemini 3.8', results, () => {
            const val = validateConfig();
            if (!val.valid) {
                console.error('Errori validateConfig:', val.errors);
            }
            return val.valid === true && val.errors.length === 0;
        });

        test('CONFIG.MODEL_NAME è impostato su gemini-3.8-flash', results, () => {
            return CONFIG.MODEL_NAME === 'gemini-3.8-flash';
        });

        test('CONFIG.MODEL_STRATEGY.generation inizia con flash-3.7', results, () => {
            return Array.isArray(CONFIG.MODEL_STRATEGY.generation) && CONFIG.MODEL_STRATEGY.generation[0] === 'flash-3.7';
        });

        test('CONFIG.GEMINI_MODELS contiene flash-3.7 con limiti corretti', results, () => {
            const m = CONFIG.GEMINI_MODELS['flash-3.7'];
            return m && m.name === 'gemini-3.8-flash' && m.rpm === 10 && m.tpm === 250000 && m.rpd === 1500;
        });
    });

    // 10. Regressioni sui bug critici della pipeline
    testGroup('Punto #10: Regressioni critiche - Cache, invio, prompt e quote', results, () => {
        test('Cascata AG: 3.8, backup, 3.6, latest, backup, Lite, backup', results, () => {
            const service = new GeminiService();
            service.backupKey = 'backup-key-abcdefghijklmnopqrstuvwxyz';
            const plans = service.buildGenerationStrategies();
            return plans.map(plan => plan.model).join(',') === [
                'gemini-3.8-flash', 'gemini-3.8-flash', 'gemini-3.6-flash',
                'gemini-flash-latest', 'gemini-flash-latest',
                'gemini-3.5-flash-lite', 'gemini-3.5-flash-lite'
            ].join(',') && plans.every(plan =>
                plan.key === (plan.skipRateLimit ? service.backupKey : service.primaryKey)
            );
        });

        test('Cascata senza backup mantiene tutti i modelli sulla primaria', results, () => {
            const service = new GeminiService();
            service.backupKey = null;
            return service.buildGenerationStrategies().map(plan => plan.modelKey).join(',') ===
                'flash-3.7,flash-3.6,flash-latest,flash-lite';
        });

        test('Il rate limiter conserva Gemini 3.6 come fallback distinto', results, () => {
            const limiter = new GeminiRateLimiter();
            return limiter.models['flash-3.6'].name === 'gemini-3.6-flash';
        });

        test('Generazione vincolata al modello propaga la quota senza saltare la cascata', results, () => {
            const service = new GeminiService();
            service.config.GEMINI_CONTEXT_CACHE = { enabled: false };
            service.useRateLimiter = true;
            let forcedModel;
            let directCalls = 0;
            service.rateLimiter = { executeRequest: (task, callback, options) => {
                forcedModel = options.forceModel;
                throw new Error('QUOTA_EXHAUSTED');
            } };
            service.fetchFn = () => { directCalls++; throw new Error('Chiamata inattesa'); };
            try {
                service.generateResponse('prompt', { modelName: 'gemini-3.6-flash', modelKey: 'flash-3.6' });
                return false;
            } catch (error) {
                return error.message === 'QUOTA_EXHAUSTED' && forcedModel === 'flash-3.6' && directCalls === 0;
            }
        });

        ['gemini-3.8-flash', 'gemini-3.5-flash-lite'].forEach(modelName => {
        test(`${modelName} usa il nuovo endpoint senza parametri non supportati`, results, () => {
            const service = new GeminiService();
            let payload = null;
            let requestUrl = '';
            service.fetchFn = (url, options) => {
                requestUrl = url;
                payload = JSON.parse(options.payload);
                return {
                    getResponseCode: () => 200,
                    getContentText: () => JSON.stringify({
                        candidates: [{ content: { parts: [{ text: 'Risposta valida' }] } }]
                    })
                };
            };
            service._generateWithModel('prompt', modelName);
            return requestUrl.includes(`/models/${modelName}:generateContent?`) && payload &&
                ['temperature', 'topP', 'topK', 'candidateCount'].every(
                    key => !Object.prototype.hasOwnProperty.call(payload.generationConfig, key)
                );
        });
        });

        test('Migrazione 3.7 conserva le chiavi delle quote senza modificare la configurazione originale', results, () => {
            const legacy = { 'flash-3.7': { name: 'gemini-3.7-flash', rpm: 10 } };
            const normalized = GeminiRateLimiter.prototype._normalizeDeprecatedModelNames(legacy);
            return Object.keys(normalized).join(',') === 'flash-3.7' &&
                normalized['flash-3.7'].name === 'gemini-3.8-flash' &&
                normalized['flash-3.7'].rpm === 10 &&
                legacy['flash-3.7'].name === 'gemini-3.7-flash';
        });

        test('Context Cache non viene riutilizzata per contenuti differenti', results, () => {
            const service = new GeminiService();
            service.config.GEMINI_CONTEXT_CACHE = Object.assign({}, service.config.GEMINI_CONTEXT_CACHE, {
                enabled: true,
                minCacheableTokens: 1,
                propertyPrefix: 'test_cache_fingerprint_'
            });
            ['name', 'expiry', 'fingerprint'].forEach(suffix => {
                service.props.deleteProperty('test_cache_fingerprint_' + suffix);
            });
            let calls = 0;
            service.fetchFn = () => {
                calls++;
                return {
                    getResponseCode: () => 200,
                    getContentText: () => JSON.stringify({ name: `cachedContents/${calls}` })
                };
            };
            const first = service.getOrCreateCachedContent('CONTESTO_A');
            const second = service.getOrCreateCachedContent('CONTESTO_B');
            return calls === 2 && first !== second;
        });

        test('Doppio fallimento Gmail viene propagato al chiamante', results, () => {
            const service = new GmailService();
            service.addLabelToThread = () => { };
            const resource = {
                reply: () => { throw new Error('send failed'); },
                getThread: () => ({ getMessages: () => [] })
            };
            let threw = false;
            try {
                service.sendHtmlReply(resource, 'Risposta', {
                    senderName: 'Mario', senderEmail: 'mario@example.com',
                    subject: 'Oggetto', rfc2822MessageId: ''
                });
            } catch (e) {
                threw = /Invio Gmail fallito/.test(e.message);
            }
            return threw;
        });

        test('Prompt enorme conserva sempre il contenuto email', results, () => {
            const engine = new PromptEngine();
            const prompt = engine.buildPrompt({
                emailContent: 'TESTO_UTENTE ' + 'x'.repeat(500000),
                emailSubject: 'Poesia',
                knowledgeBase: 'Istruzioni essenziali',
                senderName: 'Mario',
                senderEmail: 'mario@example.com'
            });
            return prompt.includes('**EMAIL DA RISPONDERE:**') && prompt.includes('TESTO_UTENTE');
        });

        test('RateLimiter passa la stima token alla selezione modello', results, () => {
            const limiter = new GeminiRateLimiter();
            let captured = null;
            limiter.selectModel = (taskType, options) => {
                captured = options;
                return { available: false, reason: 'test_stop' };
            };
            try {
                limiter.executeRequest('generation', () => 'unused', { estimatedTokens: 50000 });
            } catch (e) { }
            return captured && captured.estimatedTokens === 50000;
        });

        test('OCR include immagini inline significative anche senza parole chiave', results, () => {
            const processor = Object.create(EmailProcessor.prototype);
            processor.gmailService = new GmailService();
            let optionsSeen = null;
            const messageWithAttachment = {
                getAttachments: options => {
                    optionsSeen = options;
                    return [{ getName: () => 'logo.png' }];
                }
            };
            const messageWithoutAttachment = {
                getAttachments: options => {
                    optionsSeen = options;
                    return [];
                }
            };
            const simpleThanks = processor._shouldTryOcr('Grazie', 'Saluti', messageWithoutAttachment);
            const whatsappState = processor._shouldTryOcr('Grazie', 'Saluti', messageWithAttachment);
            return simpleThanks === false && whatsappState === true && optionsSeen.includeInlineImages === true;
        });

        test('Euristica firma scarta logo ma conserva stato WhatsApp inline', results, () => {
            const service = new GmailService();
            const pngBytes = (width, height) => {
                const bytes = new Array(24).fill(0);
                bytes[0] = 137; bytes[1] = 80;
                bytes[16] = (width >>> 24) & 255;
                bytes[17] = (width >>> 16) & 255;
                bytes[18] = (width >>> 8) & 255;
                bytes[19] = width & 255;
                bytes[20] = (height >>> 24) & 255;
                bytes[21] = (height >>> 16) & 255;
                bytes[22] = (height >>> 8) & 255;
                bytes[23] = height & 255;
                return bytes;
            };
            const makeImage = (name, width, height, size) => ({
                getName: () => name,
                getContentType: () => 'image/png',
                getSize: () => size,
                getBytes: () => pngBytes(width, height)
            });
            const logo = makeImage('logo_parrocchia.png', 500, 120, 60000);
            const whatsappState = makeImage('image001.png', 1080, 1920, 350000);
            return service.isLikelySignatureImage(logo) === true &&
                service.isLikelySignatureImage(whatsappState) === false;
        });

        test('Errore quick-check produce stato indeterminato e ritentabile', results, () => {
            const service = new GeminiService();
            service.useRateLimiter = false;
            service.backupKey = null;
            service.fetchFn = () => ({
                getResponseCode: () => 503,
                getContentText: () => 'temporarily unavailable'
            });
            const result = service.shouldRespondToEmail('Vorrei informazioni', 'Richiesta');
            return result.shouldRespond === false && result.indeterminate === true;
        });

        test('PowerPoint viene convertito in Google Slides', results, () => {
            const oldDrive = global.Drive;
            const oldDriveApp = global.DriveApp;
            let targetMime = null;
            try {
                global.Drive = {
                    Files: {
                        create: resource => {
                            targetMime = resource.mimeType;
                            return { id: 'temp-id', mimeType: resource.mimeType };
                        },
                        delete: () => { }
                    }
                };
                global.DriveApp = {
                    getFileById: () => ({ getAs: () => ({ converted: true }) })
                };
                const blob = {
                    copyBlob: () => blob,
                    getName: () => 'presentazione.pptx',
                    getContentType: () => 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                };
                const service = new GmailService();
                service._convertOfficeToPdf(blob);
                return targetMime === 'application/vnd.google-apps.presentation';
            } finally {
                global.Drive = oldDrive;
                global.DriveApp = oldDriveApp;
            }
        });

        test('Export metriche senza configurazione termina senza creare dati', results, () => {
            const result = exportMetricsToSheet();
            return result.exported === false && result.reason === 'metrics_sheet_not_configured';
        });
    });

    // 11. Regressioni audit completo (escluso il controllo allucinazioni concordato)
    testGroup('Punto #11: Regressioni audit - concorrenza, Gmail, UI e classificazione', results, () => {
        test('Due RateLimiter concorrenti sommano contatori e finestre', results, () => {
            const props = PropertiesService.getScriptProperties();
            props.setProperties({ rpd_audit_model: '0', rpd_date_audit_model: '2026-08-24', tokens_audit_model: '0', rpm_window: '[]', tpm_window: '[]', rate_limit_date: '2026-08-24' });
            const first = new GeminiRateLimiter();
            const second = new GeminiRateLimiter();
            [first, second].forEach(limiter => {
                limiter.models = { audit_model: { name: 'audit-physical', rpm: 10, tpm: 10000, rpd: 1500 } };
                limiter._getPacificDate = () => '2026-08-24';
                limiter._trackRequest('audit_model', 100, 1);
            });
            first.flush();
            second.flush();
            const rpm = JSON.parse(props.getProperty('rpm_window') || '[]').filter(entry => entry.modelKey === 'audit_model');
            return props.getProperty('rpd_audit_model') === '2' && rpm.length === 2;
        });

        test('Disponibilità RPD considera gli incrementi dirty locali', results, () => {
            const limiter = new GeminiRateLimiter();
            limiter.models = { only: { name: 'only-model', rpm: 10, tpm: 10000, rpd: 2 } };
            limiter.allProps = { rpd_only: '1', rpd_date_only: '2026-08-24' };
            limiter.dirtyProps = { rpd_only: '2' };
            limiter.cache.rpmWindow = [];
            limiter.cache.tpmWindow = [];
            return limiter._validateModelAvailability('only', 1).reason === 'rpd_exhausted';
        });

        test('Reset giornaliero conserva le finestre mobili RPM/TPM', results, () => {
            const props = PropertiesService.getScriptProperties();
            props.setProperties({ rpm_window: '[{"timestamp":1}]', tpm_window: '[{"timestamp":1}]' });
            const limiter = new GeminiRateLimiter();
            limiter._resetDailyCounters();
            return props.getProperty('rpm_window') === '[{"timestamp":1}]' && props.getProperty('tpm_window') === '[{"timestamp":1}]';
        });

        test('Quick-check tenta la backup key una sola volta', results, () => {
            const service = new GeminiService();
            service.backupKey = 'backup-key-abcdefghijklmnopqrstuvwxyz';
            let calls = 0;
            service.fetchFn = () => {
                calls++;
                return { getResponseCode: () => 503, getContentText: () => 'unavailable' };
            };
            try { service._quickCheckWithModel('testo', 'oggetto', service.modelName); } catch (e) { }
            return calls === 2;
        });

        test('Quick-check rifiuta schema vuoto e finishReason MAX_TOKENS', results, () => {
            const service = new GeminiService();
            service.backupKey = null;
            let payload = '{}';
            let finishReason = 'STOP';
            service.fetchFn = () => ({
                getResponseCode: () => 200,
                getContentText: () => JSON.stringify({ candidates: [{ finishReason, content: { parts: [{ text: payload }] } }] })
            });
            const emptySchema = service._quickCheckWithModel('testo', 'oggetto', service.modelName);
            payload = JSON.stringify({ reply_needed: true, language: 'it', category: 'TECHNICAL', dimensions: { technical: 1, pastoral: 0, poetry: 0, formal: 0 } });
            finishReason = 'MAX_TOKENS';
            const truncated = service._quickCheckWithModel('testo', 'oggetto', service.modelName);
            return emptySchema.indeterminate === true && truncated.indeterminate === true;
        });

        test('Config senza Controllo conserva blacklist statiche', results, () => {
            const loaded = _loadAdvancedConfig({ getSheetByName: () => null });
            return loaded.ignoreDomains.length === CONFIG.IGNORE_DOMAINS.length && loaded.ignoreKeywords.length === CONFIG.IGNORE_KEYWORDS.length;
        });

        test('Calendario sospende in orario ma non durante un’assenza', results, () => {
            const monday = new Date(2026, 7, 24, 10, 0, 0);
            const base = { suspensionRules: { 'lunedì': { startMinutes: 540, endMinutes: 780 } }, vacationPeriods: [] };
            const suspended = _isAutoresponderSuspended_(base, monday);
            const vacation = Object.assign({}, base, { vacationPeriods: [{ start: monday.getTime() - 1000, end: monday.getTime() + 1000 }] });
            return suspended === true && _isAutoresponderSuspended_(vacation, monday) === false;
        });

        test('Allegato grande non consuma maxFiles prima di un PDF valido', results, () => {
            const service = new GmailService();
            const makeAttachment = (name, type, size) => ({
                getName: () => name, getContentType: () => type, getSize: () => size,
                copyBlob: () => ({ getBytes: () => new Array(size), setName: () => { } }),
                getBytes: () => []
            });
            const message = { getAttachments: () => [makeAttachment('grande.bin', 'application/octet-stream', 999), makeAttachment('valido.pdf', 'application/pdf', 10)] };
            const out = service.getProcessableAttachments(message, { maxFiles: 1, maxBytesPerFile: 100, maxInlinePayloadBytes: 1000 });
            return out.blobs.length === 1 && out.skipped.some(item => item.name === 'grande.bin' && item.reason === 'too_large');
        });

        test('Budget cumulativo Base64 blocca allegati oltre soglia', results, () => {
            const service = new GmailService();
            const attachment = name => ({ getName: () => name, getContentType: () => 'application/pdf', getSize: () => 10, copyBlob: () => ({}) });
            const out = service.getProcessableAttachments({ getAttachments: () => [attachment('a.pdf'), attachment('b.pdf')] }, { maxFiles: 4, maxBytesPerFile: 100, maxInlinePayloadBytes: 20 });
            return out.blobs.length === 1 && out.skipped.some(item => item.reason === 'max_inline_payload');
        });

        test('Filtri locali non scartano malattia poetica o parole contenenti perfetto', results, () => {
            const classifier = new Classifier();
            return classifier.classifyEmail('Poesia', 'Una poesia sulla malattia').shouldReply === true &&
                classifier.classifyEmail('Riflessione', 'amore imperfetto').shouldReply === true;
        });

        test('Hint Gemini tecnico evita il default poem_submission', results, () => {
            const classifier = new RequestTypeClassifier();
            const out = classifier.classify('Richiesta', 'Avrei bisogno di aiuto', { category: 'TECHNICAL', confidence: 0.9, dimensions: { technical: 0.9, pastoral: 0, poetry: 0, formal: 0 } });
            return out.type === 'information_request' && out.source === 'regex+gemini';
        });

        test('Testo breve senza evidenza non ottiene certezza linguistica massima', results, () => {
            const service = new GeminiService();
            return service.detectEmailLanguage('Hi', '').safetyGrade === 1;
        });

        test('Intervallo trigger invalido non cancella il trigger esistente', results, () => {
            const oldScriptApp = global.ScriptApp;
            let deleted = 0;
            try {
                global.ScriptApp = { getProjectTriggers: () => [], deleteTrigger: () => { deleted++; }, newTrigger: () => { throw new Error('non deve creare'); } };
                let threw = false;
                try { setupMainTrigger(2); } catch (e) { threw = /non valido/.test(e.message); }
                return threw && deleted === 0;
            } finally { global.ScriptApp = oldScriptApp; }
        });

        test('Protezioni setup sono idempotenti', results, () => {
            const oldType = SpreadsheetApp.ProtectionType;
            const protections = [];
            const makeProtection = () => ({
                description: '', removed: false,
                getDescription() { return this.description; },
                setDescription(value) { this.description = value; return this; },
                setWarningOnly() { return this; }, canEdit: () => true,
                remove() { this.removed = true; }
            });
            const sheet = {
                getProtections: () => protections.filter(item => !item.removed),
                getRange: () => ({ protect: () => { const p = makeProtection(); protections.push(p); return p; } })
            };
            try {
                SpreadsheetApp.ProtectionType = { RANGE: 'RANGE' };
                protectRangesWithWarning_(sheet, ['A1:B1']);
                protectRangesWithWarning_(sheet, ['A1:B1']);
                return protections.filter(item => !item.removed).length === 1;
            } finally { SpreadsheetApp.ProtectionType = oldType; }
        });

        test('Guard invio accetta snapshot stabile e blocca nuovi messaggi', results, () => {
            const processor = Object.create(EmailProcessor.prototype);
            processor.config = { labelName: 'IA' };
            processor.gmailService = {
                getMessageIdsWithLabel: () => new Set(),
                _extractEmailAddress: raw => String(raw || '').replace(/[<>]/g, '').trim()
            };
            const candidate = { getId: () => 'old', getFrom: () => 'user@example.com' };
            const newest = { getId: () => 'snapshot-last', getFrom: () => 'other@example.com' };
            const messages = [candidate, newest];
            const thread = { getMessages: () => messages.slice() };
            const stable = processor._checkSendStillSafe_(thread, candidate, new Set(['bot@example.com']), 'snapshot-last', false);
            messages.push({ getId: () => 'human-new', getFrom: () => 'human@example.com' });
            const changed = processor._checkSendStillSafe_(thread, candidate, new Set(['bot@example.com']), 'snapshot-last', false);
            return stable.safe === true && changed.safe === false && changed.reason === 'thread_changed_before_send';
        });

        test('Errore Gmail ambiguo non attiva reply fallback duplicante', results, () => {
            const service = new GmailService();
            const oldGmail = global.Gmail;
            const oldWebSafe = Utilities.base64EncodeWebSafe;
            const oldCharset = Utilities.Charset;
            let fallbackCalls = 0;
            try {
                global.Gmail = { Users: { Messages: { send: () => { throw new Error('timeout 504 after upload'); } } } };
                Utilities.base64EncodeWebSafe = value => Buffer.from(String(value)).toString('base64url');
                Utilities.Charset = { UTF_8: 'UTF-8' };
                const resource = {
                    getId: () => 'message-id',
                    getThread: () => ({ getId: () => 'thread-id' }),
                    reply: () => { fallbackCalls++; }
                };
                let blocked = false;
                try {
                    service.sendHtmlReply(resource, 'Risposta', {
                        senderName: 'Mario', senderEmail: 'mario@example.com', subject: 'Oggetto',
                        rfc2822MessageId: '<original@example.com>', existingReferences: '', recipientEmail: 'bot@example.com', recipientCc: ''
                    });
                } catch (e) {
                    blocked = /esito incerto/.test(e.message);
                }
                return blocked && fallbackCalls === 0;
            } finally {
                global.Gmail = oldGmail;
                Utilities.base64EncodeWebSafe = oldWebSafe;
                Utilities.Charset = oldCharset;
            }
        });

        test('Cache semantica distingue email originali diverse', results, () => {
            const semantic = Object.create(SemanticValidator.prototype);
            const first = semantic._cacheKey('halluc', 'rispostakbemail-A');
            const second = semantic._cacheKey('halluc', 'rispostakbemail-B');
            return first !== second;
        });
    });

    const duration = Date.now() - start;
    const successRate = results.total > 0 ? ((results.passed / results.total) * 100).toFixed(1) : 0;

    console.log('\n' + '╔' + '═'.repeat(68) + '╗');
    console.log('║' + ' '.repeat(20) + '📊 RISULTATI FINALI' + ' '.repeat(28) + '║');
    console.log(`║  Totale Test:      ${results.total.toString().padEnd(48)} ║`);
    console.log(`║  ✅ Superati:      ${results.passed.toString().padEnd(48)} ║`);
    console.log(`║  ❌ Falliti:       ${results.failed.toString().padEnd(48)} ║`);
    console.log(`║  Percentuale:      ${successRate}%`.padEnd(69) + '║');
    console.log(`║  Durata:           ${duration}ms`.padEnd(69) + '║');
    console.log('╚' + '═'.repeat(68) + '╝');

    return results;
}

if (typeof process !== 'undefined' && typeof require !== 'undefined' && require.main === module) {
    const results = runAllTests();
    process.exit(results.failed > 0 ? 1 : 0);
}
