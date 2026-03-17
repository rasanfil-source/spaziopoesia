/**
 * @file gas_setup_ui.js
 * @description Setup UI "single-sheet first" per Controllo, con validazioni e protezioni robuste.
 */

const UI_CONFIG = {
  CONTROLLO_SHEET: 'Controllo',
  SHEET_RESET_RANGE: 'A1:Z300',
  DAYS: ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica']
};

function logMessage_(message) {
  if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
    Logger.log(message);
    return;
  }
  if (typeof console !== 'undefined' && console && typeof console.log === 'function') {
    console.log(message);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Poesia IA')
    .addItem('Configura foglio Controllo (Reset Layout)', 'setupConfigurationSheets')
    .addItem('🔧 Applica regole validazione e basta', 'applyValidationOnly')
    .addItem('🗑️ Svuota cache conoscenza', 'clearKnowledgeCache')
    .addItem('Test configurazione', 'testConfiguration')
    .addToUi();
}

/**
 * Applica SOLO le regole di validazione dati al foglio Controllo.
 * Non modifica layout, colori o testi. Utile per correggere i formati.
 * ORA RICREA ANCHE I NAMED RANGES MANCANTI.
 */
function applyValidationOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(UI_CONFIG.CONTROLLO_SHEET);
  const setupWarnings = [];

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Errore: Foglio "Controllo" non trovato.');
    return;
  }

  // PULIZIA Totale delle validazioni per evitare conflitti
  const rangesToClean = ['B2', 'B5:B7', 'D5:D7', 'B10:B16', 'D10:D16', 'E11:E120', 'F11:F120'];
  rangesToClean.forEach(a1 => {
    try {
      sheet.getRange(a1).clearDataValidations();
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      const warning = `clearDataValidations ${a1}: ${msg}`;
      setupWarnings.push(warning);
      console.warn('⚠️ ' + warning);
    }
  });

  SpreadsheetApp.flush();

  // Applica constraints usando la funzione helper dell'utente
  applyControlloInputConstraints_(sheet);

  // CRUCIALE: Ricrea i Named Ranges se mancano!
  createNamedRanges(ss, setupWarnings);

  SpreadsheetApp.flush();

  if (setupWarnings.length > 0) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Regole applicate con avvisi:\n\n' +
      setupWarnings.map(w => '• ' + w).join('\n')
    );
    return;
  }

  SpreadsheetApp.getUi().alert('✅ Regole di validazione applicate e Named Ranges ripristinati!');
}

/**
 * Setup principale: mantiene il layout nel foglio Controllo (stile base), senza imporre multi-foglio.
 */
function setupConfigurationSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];

  runSetupStep_('Controllo', function () {
    setupControlloSheet(ss);
  }, errors);

  runSetupStep_('Named Ranges', function () {
    createNamedRanges(ss);
  }, errors);

  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert('Configurazione completata con avvisi\n\n' + errors.join('\n'));
    return;
  }

  SpreadsheetApp.getUi().alert('Configurazione completata con successo sul foglio Controllo.');
}

function runSetupStep_(stepName, fn, errors) {
  try {
    fn();
  } catch (e) {
    const message = e && e.message ? e.message : String(e);
    const formatted = '• ' + stepName + ': ' + message;
    errors.push(formatted);
    logMessage_('[setupConfigurationSheets] ' + formatted);
  }
}

function setupControlloSheet(ss) {
  const sheet = getOrCreateSheet(ss, UI_CONFIG.CONTROLLO_SHEET, '#4285F4');
  resetSheetLayout(sheet);

  // HERO
  safeMerge(sheet.getRange('A1:B1'));
  sheet.getRange('A1:B1')
    .setValue('STATO DEL SISTEMA')
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('left');

  // Nota: B2 viene gestito da applyControlloInputConstraints_

  safeMerge(sheet.getRange('E2:F2'));
  applyFormulaWithLocaleFallback_(
    sheet.getRange('E2:F2'),
    '=IF($B$2="Spento";"🔴 Spento";IF(SUMPRODUCT((TODAY()>=$B$5:$B$7)*(TODAY()<=$D$5:$D$7)*($B$5:$B$7<>""))>0;"🟢 Attiva (Ferie/H24)";IFERROR(IF(AND(HOUR(NOW())>=INDEX($B$10:$B$16;MATCH(LOWER(TEXT(TODAY();"dddd"));LOWER($A$10:$A$16);0));HOUR(NOW())<INDEX($D$10:$D$16;MATCH(LOWER(TEXT(TODAY();"dddd"));LOWER($A$10:$A$16);0)));"🟡 Sospesa (orari)";"🟢 Attiva");"🟢 Attiva")))'
  );
  sheet.getRange('E2:F2')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#E6F4EA');

  safeMerge(sheet.getRange('A3:F3'));
  sheet.getRange('A3:F3').setValue('La risposta automatica è attiva fuori dalla presenza segreteria e quando non ci sono assenze attive.');

  // ASSENZE compatte nel Controllo
  sheet.getRange('A4').setValue('🟢 Assenze segretario').setFontWeight('bold');
  sheet.getRange('C4').setValue('dal').setFontWeight('bold');
  sheet.getRange('D4').setValue('al').setFontWeight('bold');

  sheet.getRange('A5:A7').setValues([['Ferie'], ['Permesso'], ['Malattia']]);

  // Le regole date sono applicate dopo

  // RIASSUNTO
  sheet.getRange('E4').setValue('RIASSUNTO').setFontWeight('bold');
  sheet.getRange('E5').setValue('Risposta automatica:').setFontWeight('bold');
  sheet.getRange('E6').setValue('Segretario:').setFontWeight('bold');
  sheet.getRange('E7').setValue('Oggi:').setFontWeight('bold');
  sheet.getRange('E8').setValue('Motivo:').setFontWeight('bold');
  // Sospensione oraria rimossa

  // Filtri nella stessa pagina
  safeMerge(sheet.getRange('E9:F9'));
  sheet.getRange('E9:F9').setValue('email cui non rispondere').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#F4CCCC');
  sheet.getRange('E10').setValue('domini').setFontWeight('bold').setBackground('#F97316').setHorizontalAlignment('center');
  sheet.getRange('F10').setValue('parole').setFontWeight('bold').setBackground('#F97316').setHorizontalAlignment('center');

  // Layout generale
  sheet.setColumnWidth(1, 330);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 290);
  sheet.setColumnWidth(6, 290);
  sheet.setFrozenRows(1);

  // Applica le constraints GUIDATE dall'utente
  applyControlloInputConstraints_(sheet);
}

function applyFormulaWithLocaleFallback_(range, formula) {
  if (!range) return;
  const safeFormula = typeof formula === 'string' ? formula : '';
  const hasLocalSeparators = safeFormula.includes(';');

  if (hasLocalSeparators && typeof range.setFormulaLocal === 'function') {
    try {
      range.setFormulaLocal(safeFormula);
      return;
    } catch (err) {
      logMessage_('[applyFormulaWithLocaleFallback_] setFormulaLocal fallita, provo fallback US: ' + (err && err.message ? err.message : String(err)));
    }
  }

  if (typeof range.setFormula !== 'function') return;

  try {
    range.setFormula(safeFormula);
  } catch (err) {
    if (!hasLocalSeparators) throw err;
    const fallbackFormula = safeFormula.replace(/;/g, ',');
    range.setFormula(fallbackFormula);
  }
}

function createNamedRanges(ss, warningsCollector) {
  const ranges = [
    { name: 'cfg_system_master', range: "'Controllo'!B2" },
    { name: 'cfg_timezone', range: "'Controllo'!B4" },
    // Legacy: nome ambiguo mantenuto per retrocompatibilità (punta alla prima data ferie).
    { name: 'cfg_holidays_mode', range: "'Controllo'!B5" },
    // Nome esplicito consigliato per nuove integrazioni.
    { name: 'cfg_vacation_start_date', range: "'Controllo'!B5" },
    { name: 'sum_auto_status', range: "'Controllo'!F5" },
    { name: 'sum_secretary_status', range: "'Controllo'!F6" },
    { name: 'sum_today_date', range: "'Controllo'!F7" },
    { name: 'sum_today_reason', range: "'Controllo'!F8" },
    { name: 'tbl_absences', range: "'Controllo'!A5:D7" }, // Assenze da A5 a D7
    { name: 'lst_ignore_domains', range: "'Controllo'!E11:E" },
    { name: 'lst_ignore_keywords', range: "'Controllo'!F11:F" }
  ];

  const documentLock = LockService.getDocumentLock();
  let lockAcquired = false;

  try {
    // Evita race condition in esecuzioni concorrenti (es. trigger multipli setup).
    // Timeout breve: se il lock non arriva, preferiamo warning esplicito a stato intermedio silenzioso.
    lockAcquired = documentLock.tryLock(5000);
    if (!lockAcquired) {
      const warning = 'Lock documento non acquisito: creazione named ranges saltata per evitare stato incoerente.';
      console.warn(warning);
      if (Array.isArray(warningsCollector)) warningsCollector.push(warning);
      return;
    }

    ranges.forEach(def => {
      removeNamedRangeIfExists(ss, def.name);
      try { ss.setNamedRange(def.name, ss.getRange(def.range)); } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        const warning = 'Errore creazione named range ' + def.name + ': ' + msg;
        console.warn(warning);
        if (Array.isArray(warningsCollector)) {
          warningsCollector.push(warning);
        }
      }
    });
  } finally {
    if (lockAcquired) {
      documentLock.releaseLock();
    }
  }
}

function testConfiguration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checks = [];

  checks.push({ name: 'Foglio Controllo', ok: !!ss.getSheetByName('Controllo') });

  const nrList = ['cfg_system_master', 'sum_auto_status', 'tbl_absences', 'lst_ignore_domains'];
  nrList.forEach(n => {
    const nr = ss.getRangeByName(n);
    checks.push({ name: 'Named range ' + n, ok: !!nr });
  });

  const lines = checks.map(c => (c.ok ? '✅ ' : '❌ ') + c.name);
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

function getOrCreateSheet(ss, name, tabColor) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.setTabColor(tabColor);
  return sheet;
}

function resetSheetLayout(sheet) {
  const range = sheet.getRange(UI_CONFIG.SHEET_RESET_RANGE);
  range.breakApart();
  range.clear();
}

function safeMerge(range) {
  range.breakApart();
  range.merge();
}

function applyFormulaValidationWithFallback_(range, formulas, helpText) {
  let lastError = null;
  for (let i = 0; i < formulas.length; i++) {
    try {
      const rule = SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied(formulas[i])
        .setAllowInvalid(false)
        .setHelpText(helpText)
        .build();
      range.setDataValidation(rule);
      return;
    } catch (e) {
      lastError = e;
    }
  }
  console.warn('Formula validation failed: ' + (lastError ? lastError.message : 'Unknown error'));
}

function removeNamedRangeIfExists(ss, name) {
  const existing = ss.getNamedRanges().find(nr => nr.getName() === name);
  if (existing) ss.removeNamedRange(name);
}

/**
 * Applica validazioni e protezioni input sul foglio Controllo
 * (Logica suggerita dall'utente)
 */
function applyControlloInputConstraints_(sheet) {
  // B2: interruttore
  const switchRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Acceso', 'Spento'], true)
    .setAllowInvalid(false)
    .setHelpText('Inserisci solo "Acceso" o "Spento".')
    .build();
  sheet.getRange('B2').setDataValidation(switchRule).setHorizontalAlignment('center').setFontWeight('bold');

  if (!sheet.getRange('B2').getDisplayValue()) {
    sheet.getRange('B2').setValue('Acceso');
  }

  // Date ferie/malattia: B5,D5 - B7,D7
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Inserisci una data valida (formato italiano: gg/mm/aaaa).')
    .build();

  ['B5', 'D5', 'B6', 'D6', 'B7', 'D7'].forEach(a1 => {
    sheet.getRange(a1).setDataValidation(dateRule).setNumberFormat('dd/MM/yyyy');
  });

  // Coerenza data inizio <= data fine per ogni riga 5..7
  applyFormulaValidationWithFallback_(
    sheet.getRange('D5:D7'),
    ['=OR($B5="";$D5="";$B5<=$D5)', '=OR($B5="",$D5="",$B5<=$D5)'],
    'La data fine deve essere uguale o successiva alla data inizio.'
  );

  // Orari sospensione rimossi

  // Filtri domini/keyword direttamente su Controllo
  const domainRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=OR(E11="";REGEXMATCH(E11;"^(?:[a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,}$"))')
    .setAllowInvalid(false)
    .setHelpText('Inserisci solo dominio/email parziale (es. amazon.com).')
    .build();
  sheet.getRange('E11:E120').setDataValidation(domainRule);

  const keywordRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=OR(F11="";LEN(TRIM(F11))>0)')
    .setAllowInvalid(false)
    .setHelpText('Inserisci una parola/frase da escludere.')
    .build();
  sheet.getRange('F11:F120').setDataValidation(keywordRule);

  // Protezioni Warning Only sulle etichette
  protectRangesWithWarning_(sheet, ['A1:F1', 'A3:A8', 'C4:D4', 'C5:C7', 'E4:F10']);
}

/**
 * Imposta protezioni in modalità warning-only sui range specificati.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} rangesA1
 */
function protectRangesWithWarning_(sheet, rangesA1) {
  rangesA1.forEach(a1 => {
    const range = sheet.getRange(a1);
    const protection = range.protect();
    protection.setWarningOnly(true);
    protection.setDescription('Area protetta setup UI: ' + a1);
  });
}
