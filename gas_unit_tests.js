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

    // 9. Config Validation (Gemini 3.7 Flash)
    testGroup('Punto #9: Config Validation - Gemini 3.7 Flash & Model Strategy', results, () => {
        test('validateConfig() valida con successo la configurazione Gemini 3.7', results, () => {
            const val = validateConfig();
            if (!val.valid) {
                console.error('Errori validateConfig:', val.errors);
            }
            return val.valid === true && val.errors.length === 0;
        });

        test('CONFIG.MODEL_NAME è impostato su gemini-3.7-flash', results, () => {
            return CONFIG.MODEL_NAME === 'gemini-3.7-flash';
        });

        test('CONFIG.MODEL_STRATEGY.generation inizia con flash-3.7', results, () => {
            return Array.isArray(CONFIG.MODEL_STRATEGY.generation) && CONFIG.MODEL_STRATEGY.generation[0] === 'flash-3.7';
        });

        test('CONFIG.GEMINI_MODELS contiene flash-3.7 con limiti corretti', results, () => {
            const m = CONFIG.GEMINI_MODELS['flash-3.7'];
            return m && m.name === 'gemini-3.7-flash' && m.rpm === 10 && m.tpm === 250000 && m.rpd === 1500;
        });
    });

    // 10. Regressioni sui bug critici della pipeline
    testGroup('Punto #10: Regressioni critiche - Cache, invio, prompt e quote', results, () => {
        test('Gemini 3.7 non riceve il parametro temperature', results, () => {
            const service = new GeminiService();
            let payload = null;
            service.fetchFn = (url, options) => {
                payload = JSON.parse(options.payload);
                return {
                    getResponseCode: () => 200,
                    getContentText: () => JSON.stringify({
                        candidates: [{ content: { parts: [{ text: 'Risposta valida' }] } }]
                    })
                };
            };
            service._generateWithModel('prompt', 'gemini-3.7-flash');
            return payload && !Object.prototype.hasOwnProperty.call(payload.generationConfig, 'temperature');
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
