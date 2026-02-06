/**
 * Google Apps Script Backend für Arbeitszeit Tracker
 * Version 2.0 - Verbessertes Datenmodell
 */

// Name des Tabellenblattes
const SHEET_NAME = 'Zeiterfassung';
const THEMES_SHEET = 'Themen'; // Für Projekt/Thema-Vorschläge
const MAIN_HEADERS = [
  'Datum',
  'Modus',
  'Pendel-Art',
  'Losfahrt Zuhause',
  'Ankunft Arbeit',
  'Losfahrt Arbeit',
  'Ankunft Zuhause',
  'Thema/Projekt',
  'Beschreibung',
  'Arbeitszeit',
  'Pendelzeit'
];
const LEGACY_HEADERS = [
  'Datum',
  'Modus',
  'Losfahrt Zuhause',
  'Ankunft Arbeit',
  'Losfahrt Arbeit',
  'Ankunft Zuhause',
  'Arbeitszeit',
  'Pendelzeit'
];
const HEADER_KEYS = {
  date: 'Datum',
  mode: 'Modus',
  pendelArt: 'Pendel-Art',
  homeDeparture: 'Losfahrt Zuhause',
  workArrival: 'Ankunft Arbeit',
  workDeparture: 'Losfahrt Arbeit',
  homeArrival: 'Ankunft Zuhause',
  theme: 'Thema/Projekt',
  description: 'Beschreibung',
  workTime: 'Arbeitszeit',
  commuteTime: 'Pendelzeit'
};

/**
 * HTTP GET Handler - Verarbeitet alle Anfragen
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const debug = String(e.parameter.debug || '').toLowerCase();
    const wantDebug = debug === '1' || debug === 'true';
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    let themesSheet = spreadsheet.getSheetByName(THEMES_SHEET);
    
    // Erstelle Sheets falls nicht vorhanden
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      initializeSheet(sheet);
    }
    ensureSheetSchema(sheet);
    
    if (!themesSheet) {
      themesSheet = spreadsheet.insertSheet(THEMES_SHEET);
      initializeThemesSheet(themesSheet);
    }
    
    if (action === 'save') {
      // Speichern der Zeitdaten
      const data = {
        date: e.parameter.date,
        workMode: e.parameter.workMode,
        pendelArt: e.parameter.pendelArt || '',
        timeType: e.parameter.timeType, // home-departure, work-arrival, etc.
        timeValue: e.parameter.timeValue,
        theme: e.parameter.theme || '',
        description: e.parameter.description || ''
      };
      
      const result = saveTimeData(sheet, themesSheet, data);
      
      const response = { 
        success: true, 
        message: 'Daten gespeichert',
        data: result
      };
      
      if (wantDebug) {
        response.debug = buildDebugInfo(spreadsheet, sheet);
        response.request = {
          action: action,
          date: e.parameter.date || '',
          workMode: e.parameter.workMode || '',
          pendelArt: e.parameter.pendelArt || '',
          timeType: e.parameter.timeType || '',
          timeValue: e.parameter.timeValue || '',
          theme: e.parameter.theme || '',
          description: e.parameter.description || ''
        };
      }
      
      return createJsonResponse(response);
      
    } else if (action === 'getToday') {
      const date = e.parameter.date;
      const data = getTodayData(sheet, date);
      
      return createJsonResponse({ 
        success: true, 
        data: data 
      });
      
    } else if (action === 'getAllData') {
      const allData = getAllData(sheet);
      
      return createJsonResponse({ 
        success: true, 
        data: allData 
      });
      
    } else if (action === 'getThemes') {
      const themes = getThemes(themesSheet);
      
      return createJsonResponse({ 
        success: true, 
        data: themes 
      });
      
    } else if (action === 'test') {
      return createJsonResponse({ 
        success: true, 
        message: 'Apps Script funktioniert!',
        sheetName: SHEET_NAME,
        spreadsheetId: spreadsheet.getId()
      });
    } else if (action === 'saveMeta') {
      const data = {
        date: e.parameter.date,
        workMode: e.parameter.workMode,
        pendelArt: e.parameter.pendelArt || '',
        theme: e.parameter.theme || '',
        description: e.parameter.description || ''
      };
      
      const result = saveTimeData(sheet, themesSheet, data);
      
      const response = {
        success: true,
        message: 'Meta gespeichert',
        data: result
      };
      
      if (wantDebug) {
        response.debug = buildDebugInfo(spreadsheet, sheet);
        response.request = {
          action: action,
          date: e.parameter.date || '',
          workMode: e.parameter.workMode || '',
          pendelArt: e.parameter.pendelArt || '',
          theme: e.parameter.theme || '',
          description: e.parameter.description || ''
        };
      }
      
      return createJsonResponse(response);
    } else if (action === 'debug') {
      return createJsonResponse({
        success: true,
        debug: buildDebugInfo(spreadsheet, sheet)
      });
    }
    
    return createJsonResponse({ 
      success: false, 
      error: 'Unbekannte Aktion' 
    });
    
  } catch (error) {
    return createJsonResponse({ 
      success: false, 
      error: error.toString() 
    });
  }
}

/**
 * Erstellt JSON Response
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Normalisiert Datum auf dd.MM.yyyy für konsistente Vergleiche
 */
function normalizeDate(value) {
  if (!value) return '';
  
  // Date-Objekt
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  }
  
  const str = String(value).trim();
  if (!str) return '';
  
  // dd.mm.yyyy (ggf. ohne führende Nullen)
  const parts = str.split('.');
  if (parts.length === 3) {
    const day = parts[0].padStart(2, '0');
    const month = parts[1].padStart(2, '0');
    const year = parts[2].length === 2 ? `20${parts[2]}` : parts[2];
    return `${day}.${month}.${year}`;
  }
  
  // Fallback: versuche zu parsen
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  }
  
  return str;
}

/**
 * Erstellt Mapping von Header-Name zu Spaltenindex (0-basiert)
 */
function getHeaderMap(sheet) {
  const columnCount = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  const headerRow = sheet.getRange(1, 1, 1, columnCount).getValues()[0];
  const map = {};
  
  headerRow.forEach((value, idx) => {
    const key = normalizeHeader(value);
    if (key) map[key] = idx;
  });
  
  return map;
}

/**
 * Normalisiert Headertexte (Whitespace, Zeilenumbrüche, Slash-Abstände)
 */
function normalizeHeader(value) {
  const str = String(value || '').trim();
  if (!str) return '';
  
  // Mehrfach-Whitespace/Zeilenumbrüche -> ein Leerzeichen
  let out = str.replace(/\u00a0/g, ' ').replace(/[\u200B-\u200D\uFEFF]/g, '');
  out = out.replace(/\s+/g, ' ');
  // "Thema / Projekt" -> "Thema/Projekt"
  out = out.replace(/\s*\/\s*/g, '/');
  
  return out.toLowerCase();
}

/**
 * Liefert Spaltenindex über mehrere Kandidaten
 */
function findHeaderIndex(map, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const key = normalizeHeader(candidates[i]);
    if (key in map) return map[key];
  }
  return undefined;
}

/**
 * Liefert Standardindex basierend auf MAIN_HEADERS
 */
function defaultIndex(headerName) {
  const idx = MAIN_HEADERS.indexOf(headerName);
  return idx >= 0 ? idx : undefined;
}

/**
 * Stellt sicher, dass die Zeile die richtige Länge hat
 */
function ensureRowLength(row, length) {
  if (row.length < length) {
    return row.concat(new Array(length - row.length).fill(''));
  }
  if (row.length > length) {
    return row.slice(0, length);
  }
  return row;
}

/**
 * Stellt sicher, dass das Sheet die aktuelle Spaltenstruktur hat
 */
function ensureSheetSchema(sheet) {
  const colCount = Math.max(1, sheet.getLastColumn(), MAIN_HEADERS.length);
  const headerRow = sheet.getRange(1, 1, 1, colCount).getDisplayValues()[0];
  const normalized = headerRow.map(v => normalizeHeader(v));
  
  const matches = (current, target) => target.every((v, i) => (current[i] || '') === v);
  
  // Aktuelles Schema vorhanden
  if (matches(normalized, MAIN_HEADERS.map(h => normalizeHeader(h)))) return;
  
  // Altes Schema migrieren
  if (matches(normalized, LEGACY_HEADERS.map(h => normalizeHeader(h)))) {
    // Pendel-Art (Spalte C) einfügen
    sheet.insertColumns(3, 1);
    // Thema/Projekt + Beschreibung vor Arbeitszeit einfügen
    sheet.insertColumns(8, 2);
    initializeSheet(sheet);
    return;
  }
  
  // Leeres/ungewöhnliches Sheet: Header setzen
  const hasAnyHeader = normalized.some(v => v !== '');
  if (!hasAnyHeader) {
    initializeSheet(sheet);
  }
}

/**
 * Initialisiert das Hauptsheet mit Headern
 */
function initializeSheet(sheet) {
  sheet.getRange(1, 1, 1, MAIN_HEADERS.length).setValues([MAIN_HEADERS]);
  
  // Formatierung
  sheet.getRange(1, 1, 1, MAIN_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#667eea')
    .setFontColor('#ffffff');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, MAIN_HEADERS.length);
  
  // Spaltenbreiten optimieren
  sheet.setColumnWidth(1, 100); // Datum
  sheet.setColumnWidth(2, 120); // Modus
  sheet.setColumnWidth(3, 120); // Pendel-Art
  sheet.setColumnWidth(8, 150); // Thema/Projekt
  sheet.setColumnWidth(9, 200); // Beschreibung
}

/**
 * Initialisiert das Themen-Sheet
 */
function initializeThemesSheet(sheet) {
  const headers = ['Thema/Projekt', 'Häufigkeit', 'Zuletzt verwendet'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#667eea')
    .setFontColor('#ffffff');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Speichert Zeiterfassungsdaten
 */
function saveTimeData(sheet, themesSheet, data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  
  try {
  const date = normalizeDate(data.date);
  const allData = sheet.getDataRange().getValues();
  const headerMap = getHeaderMap(sheet);
  const columnCount = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  
  const idxDate = findHeaderIndex(headerMap, [HEADER_KEYS.date]);
  const idxMode = findHeaderIndex(headerMap, [HEADER_KEYS.mode]);
  const idxPendelArt = findHeaderIndex(headerMap, [HEADER_KEYS.pendelArt]);
  const idxHomeDeparture = findHeaderIndex(headerMap, [HEADER_KEYS.homeDeparture]);
  const idxWorkArrival = findHeaderIndex(headerMap, [HEADER_KEYS.workArrival, 'Ankunft\nArbeit']);
  const idxWorkDeparture = findHeaderIndex(headerMap, [HEADER_KEYS.workDeparture]);
  const idxHomeArrival = findHeaderIndex(headerMap, [HEADER_KEYS.homeArrival, 'Ankunft\nZuhause']);
  const idxTheme = findHeaderIndex(headerMap, [HEADER_KEYS.theme, 'Thema / Projekt']);
  const idxDescription = findHeaderIndex(headerMap, [HEADER_KEYS.description]);
  const idxWorkTime = findHeaderIndex(headerMap, [HEADER_KEYS.workTime]);
  const idxCommuteTime = findHeaderIndex(headerMap, [HEADER_KEYS.commuteTime]);
  
  const colDate = idxDate !== undefined ? idxDate : defaultIndex(HEADER_KEYS.date);
  const colMode = idxMode !== undefined ? idxMode : defaultIndex(HEADER_KEYS.mode);
  const colPendelArt = idxPendelArt !== undefined ? idxPendelArt : defaultIndex(HEADER_KEYS.pendelArt);
  const colHomeDeparture = idxHomeDeparture !== undefined ? idxHomeDeparture : defaultIndex(HEADER_KEYS.homeDeparture);
  const colWorkArrival = idxWorkArrival !== undefined ? idxWorkArrival : defaultIndex(HEADER_KEYS.workArrival);
  const colWorkDeparture = idxWorkDeparture !== undefined ? idxWorkDeparture : defaultIndex(HEADER_KEYS.workDeparture);
  const colHomeArrival = idxHomeArrival !== undefined ? idxHomeArrival : defaultIndex(HEADER_KEYS.homeArrival);
  const colTheme = idxTheme !== undefined ? idxTheme : defaultIndex(HEADER_KEYS.theme);
  const colDescription = idxDescription !== undefined ? idxDescription : defaultIndex(HEADER_KEYS.description);
  const colWorkTime = idxWorkTime !== undefined ? idxWorkTime : defaultIndex(HEADER_KEYS.workTime);
  const colCommuteTime = idxCommuteTime !== undefined ? idxCommuteTime : defaultIndex(HEADER_KEYS.commuteTime);
  
  // Finde existierende Zeile für heute
  let rowIndex = -1;
  let existingRow = null;
  
  for (let i = 1; i < allData.length; i++) {
    const rowDate = normalizeDate(colDate !== undefined ? allData[i][colDate] : allData[i][0]);
    if (rowDate && rowDate === date) {
      rowIndex = i + 1; // +1 wegen 1-basiertem Index
      existingRow = allData[i];
      break;
    }
  }
  
  if (rowIndex === -1) {
    // Neue Zeile - initialisiere alle Felder
    const newRow = new Array(columnCount).fill('');
    if (colDate !== undefined) newRow[colDate] = date;
    if (colMode !== undefined) newRow[colMode] = data.workMode || '';
    if (colPendelArt !== undefined) newRow[colPendelArt] = data.pendelArt || '';
    sheet.appendRow(newRow);
    rowIndex = sheet.getLastRow();
  } else {
    // Aktualisiere Modus falls angegeben
    if (data.workMode && colMode !== undefined) {
      sheet.getRange(rowIndex, colMode + 1).setValue(data.workMode);
    }
    if (data.pendelArt && colPendelArt !== undefined) {
      sheet.getRange(rowIndex, colPendelArt + 1).setValue(data.pendelArt);
    }
  }
  
  // Für Urlaub/Krank/Abwesend: Nur Modus setzen, keine Zeiten
  if (data.workMode === 'Urlaub' || data.workMode === 'Krank' || data.workMode === 'Abwesend') {
    if (colHomeDeparture !== undefined) sheet.getRange(rowIndex, colHomeDeparture + 1).setValue('');
    if (colWorkArrival !== undefined) sheet.getRange(rowIndex, colWorkArrival + 1).setValue('');
    if (colWorkDeparture !== undefined) sheet.getRange(rowIndex, colWorkDeparture + 1).setValue('');
    if (colHomeArrival !== undefined) sheet.getRange(rowIndex, colHomeArrival + 1).setValue('');
    if (colWorkTime !== undefined) sheet.getRange(rowIndex, colWorkTime + 1).setValue('');
    if (colCommuteTime !== undefined) sheet.getRange(rowIndex, colCommuteTime + 1).setValue('');
  } else if (data.timeType && data.timeValue) {
    // Setze die spezifische Zeit
    const timeMapping = {
      'home-departure': colHomeDeparture,
      'work-arrival': colWorkArrival,
      'work-departure': colWorkDeparture,
      'home-arrival': colHomeArrival
    };
    
    if (timeMapping[data.timeType] !== undefined) {
      sheet.getRange(rowIndex, timeMapping[data.timeType] + 1).setValue(data.timeValue);
    }
    
    // Berechne Zeiten nach dem Setzen
    const arrival = colWorkArrival !== undefined ? sheet.getRange(rowIndex, colWorkArrival + 1).getValue() : '';
    const departure = colWorkDeparture !== undefined ? sheet.getRange(rowIndex, colWorkDeparture + 1).getValue() : '';
    if (colWorkTime !== undefined) {
      sheet.getRange(rowIndex, colWorkTime + 1).setValue(calculateWorkTime(arrival, departure));
    }
    
    // Pendelzeit nur bei Pendeln-Modus berechnen
    if (data.workMode === 'Pendeln' && colCommuteTime !== undefined) {
      const homeDeparture = colHomeDeparture !== undefined ? sheet.getRange(rowIndex, colHomeDeparture + 1).getValue() : '';
      const workArrival = colWorkArrival !== undefined ? sheet.getRange(rowIndex, colWorkArrival + 1).getValue() : '';
      const workDeparture = colWorkDeparture !== undefined ? sheet.getRange(rowIndex, colWorkDeparture + 1).getValue() : '';
      const homeArrival = colHomeArrival !== undefined ? sheet.getRange(rowIndex, colHomeArrival + 1).getValue() : '';
      sheet.getRange(rowIndex, colCommuteTime + 1)
        .setValue(calculateCommuteTime(homeDeparture, workArrival, workDeparture, homeArrival));
    } else {
      if (colCommuteTime !== undefined) sheet.getRange(rowIndex, colCommuteTime + 1).setValue('');
    }
  }
  
  // Setze Thema/Beschreibung
  if (data.theme !== undefined && data.theme !== null && String(data.theme).trim() !== '') {
    if (colTheme !== undefined) sheet.getRange(rowIndex, colTheme + 1).setValue(data.theme);
    // Speichere Thema für Vorschläge (nur wenn nicht leer)
    if (data.theme.trim() !== '') {
      saveTheme(themesSheet, data.theme);
    }
  }
  if (data.description !== undefined && data.description !== null && String(data.description).trim() !== '') {
    if (colDescription !== undefined) sheet.getRange(rowIndex, colDescription + 1).setValue(data.description);
  }
  
  // Formatierung
  if (rowIndex > 1) {
    const range = sheet.getRange(rowIndex, 1, 1, columnCount);
    
    // Spezielle Formatierung für Urlaub/Krank/Abwesend
    if (data.workMode === 'Urlaub') {
      range.setBackground('#fff3e0');
    } else if (data.workMode === 'Krank') {
      range.setBackground('#ffebee');
    } else if (data.workMode === 'Abwesend') {
      range.setBackground('#f3e5f5');
    } else if (rowIndex % 2 === 0) {
      // Zebra-Streifen für normale Einträge
      range.setBackground('#f8f9fa');
    }
  }
  
  const updatedRow = sheet.getRange(rowIndex, 1, 1, columnCount).getValues()[0];
  
  return {
    row: rowIndex,
    data: updatedRow,
    existed: existingRow !== null
  };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Speichert ein Thema für spätere Vorschläge
 */
function saveTheme(themesSheet, theme) {
  if (!theme || String(theme).trim() === '') return;
  
  const normalizedDisplay = normalizeThemeDisplay(theme);
  if (!normalizedDisplay) return;
  
  const allThemes = themesSheet.getDataRange().getValues();
  let found = false;
  const normalizedKey = normalizeThemeKey(normalizedDisplay);
  
  // Suche existierendes Thema
  for (let i = 1; i < allThemes.length; i++) {
    const existing = normalizeThemeDisplay(allThemes[i][0]);
    if (existing && normalizeThemeKey(existing) === normalizedKey) {
      // Erhöhe Häufigkeit
      const count = (allThemes[i][1] || 0) + 1;
      const now = new Date().toLocaleDateString('de-DE');
      themesSheet.getRange(i + 1, 1, 1, 3).setValues([[normalizedDisplay, count, now]]);
      found = true;
      break;
    }
  }
  
  // Neues Thema hinzufügen
  if (!found) {
    const now = new Date().toLocaleDateString('de-DE');
    themesSheet.appendRow([normalizedDisplay, 1, now]);
  }
}

/**
 * Lädt alle gespeicherten Themen
 */
function getThemes(themesSheet) {
  const allData = themesSheet.getDataRange().getValues();
  const themeMap = {};
  
  for (let i = 1; i < allData.length; i++) {
    if (!allData[i][0]) continue;
    
    const name = normalizeThemeDisplay(allData[i][0]);
    if (!name) continue;
    
    const key = normalizeThemeKey(name);
    const count = Number(allData[i][1] || 0) || 0;
    const lastUsed = allData[i][2] || '';
    
    if (!themeMap[key]) {
      themeMap[key] = {
        name: name,
        count: count,
        lastUsed: lastUsed
      };
    } else {
      themeMap[key].count += count;
      
      // Nimm das zuletzt verwendete Datum
      const existingDate = parseGermanDate(themeMap[key].lastUsed);
      const incomingDate = parseGermanDate(lastUsed);
      if (incomingDate && (!existingDate || incomingDate > existingDate)) {
        themeMap[key].lastUsed = lastUsed;
      }
    }
  }
  
  // Sortiere nach Häufigkeit
  const themes = Object.values(themeMap);
  themes.sort((a, b) => b.count - a.count);
  
  return themes;
}

/**
 * Normalisiert Theme für Anzeige/Speicherung
 */
function normalizeThemeDisplay(value) {
  const str = String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return str;
}

/**
 * Normalisiert Theme für Vergleich (case-insensitive)
 */
function normalizeThemeKey(value) {
  return normalizeThemeDisplay(value).toLowerCase();
}

/**
 * Parst deutsches Datum (dd.MM.yyyy) zu Date
 */
function parseGermanDate(dateStr) {
  if (!dateStr) return null;
  const parts = String(dateStr).split('.');
  if (parts.length !== 3) return null;
  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const year = parseInt(parts[2], 10);
  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
  return new Date(year, month, day);
}

/**
 * Lädt Daten für einen bestimmten Tag
 */
function getTodayData(sheet, date) {
  const targetDate = normalizeDate(date);
  const allData = sheet.getDataRange().getValues();
  const headerMap = getHeaderMap(sheet);
  
  const idxDate = findHeaderIndex(headerMap, [HEADER_KEYS.date]);
  const idxMode = findHeaderIndex(headerMap, [HEADER_KEYS.mode]);
  const idxPendelArt = findHeaderIndex(headerMap, [HEADER_KEYS.pendelArt]);
  const idxHomeDeparture = findHeaderIndex(headerMap, [HEADER_KEYS.homeDeparture]);
  const idxWorkArrival = findHeaderIndex(headerMap, [HEADER_KEYS.workArrival, 'Ankunft\nArbeit']);
  const idxWorkDeparture = findHeaderIndex(headerMap, [HEADER_KEYS.workDeparture]);
  const idxHomeArrival = findHeaderIndex(headerMap, [HEADER_KEYS.homeArrival, 'Ankunft\nZuhause']);
  const idxTheme = findHeaderIndex(headerMap, [HEADER_KEYS.theme, 'Thema / Projekt']);
  const idxDescription = findHeaderIndex(headerMap, [HEADER_KEYS.description]);
  
  const colDate = idxDate !== undefined ? idxDate : defaultIndex(HEADER_KEYS.date);
  const colMode = idxMode !== undefined ? idxMode : defaultIndex(HEADER_KEYS.mode);
  const colPendelArt = idxPendelArt !== undefined ? idxPendelArt : defaultIndex(HEADER_KEYS.pendelArt);
  const colHomeDeparture = idxHomeDeparture !== undefined ? idxHomeDeparture : defaultIndex(HEADER_KEYS.homeDeparture);
  const colWorkArrival = idxWorkArrival !== undefined ? idxWorkArrival : defaultIndex(HEADER_KEYS.workArrival);
  const colWorkDeparture = idxWorkDeparture !== undefined ? idxWorkDeparture : defaultIndex(HEADER_KEYS.workDeparture);
  const colHomeArrival = idxHomeArrival !== undefined ? idxHomeArrival : defaultIndex(HEADER_KEYS.homeArrival);
  const colTheme = idxTheme !== undefined ? idxTheme : defaultIndex(HEADER_KEYS.theme);
  const colDescription = idxDescription !== undefined ? idxDescription : defaultIndex(HEADER_KEYS.description);
  
  for (let i = 1; i < allData.length; i++) {
    const rowDate = normalizeDate(colDate !== undefined ? allData[i][colDate] : allData[i][0]);
    if (rowDate && rowDate === targetDate) {
      return {
        date: rowDate,
        workMode: colMode !== undefined ? (allData[i][colMode] || 'Pendeln') : 'Pendeln',
        pendelArt: colPendelArt !== undefined ? (allData[i][colPendelArt] || '') : '',
        'home-departure': colHomeDeparture !== undefined ? (allData[i][colHomeDeparture] || '') : '',
        'work-arrival': colWorkArrival !== undefined ? (allData[i][colWorkArrival] || '') : '',
        'work-departure': colWorkDeparture !== undefined ? (allData[i][colWorkDeparture] || '') : '',
        'home-arrival': colHomeArrival !== undefined ? (allData[i][colHomeArrival] || '') : '',
        theme: colTheme !== undefined ? (allData[i][colTheme] || '') : '',
        description: colDescription !== undefined ? (allData[i][colDescription] || '') : ''
      };
    }
  }
  
  return null;
}

/**
 * Lädt alle Daten
 */
function getAllData(sheet) {
  const allData = sheet.getDataRange().getValues();
  const headerMap = getHeaderMap(sheet);
  const idxDate = findHeaderIndex(headerMap, [HEADER_KEYS.date]);
  const colDate = idxDate !== undefined ? idxDate : defaultIndex(HEADER_KEYS.date);
  const data = [];
  
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i].slice();
    if (colDate !== undefined) {
      row[colDate] = normalizeDate(row[colDate]);
    } else {
      row[0] = normalizeDate(row[0]);
    }
    data.push(row);
  }
  
  return data;
}

/**
 * Liefert Debug-Infos zum Sheet/Mapping
 */
function buildDebugInfo(spreadsheet, sheet) {
  const colCount = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  const headerRow = sheet.getRange(1, 1, 1, colCount).getDisplayValues()[0];
  const normalized = headerRow.map(v => normalizeHeader(v));
  const headerMap = getHeaderMap(sheet);
  
  const indexFor = (label) => {
    const idx = findHeaderIndex(headerMap, [label, `${label.replace(' ', '\n')}`]);
    return idx !== undefined ? idx : defaultIndex(label);
  };
  
  return {
    spreadsheetId: spreadsheet.getId(),
    sheetName: sheet.getName(),
    lastRow: sheet.getLastRow(),
    lastColumn: sheet.getLastColumn(),
    headerRow: headerRow,
    headerRowNormalized: normalized,
    indices: {
      date: indexFor(HEADER_KEYS.date),
      mode: indexFor(HEADER_KEYS.mode),
      pendelArt: indexFor(HEADER_KEYS.pendelArt),
      homeDeparture: indexFor(HEADER_KEYS.homeDeparture),
      workArrival: indexFor(HEADER_KEYS.workArrival),
      workDeparture: indexFor(HEADER_KEYS.workDeparture),
      homeArrival: indexFor(HEADER_KEYS.homeArrival),
      theme: indexFor(HEADER_KEYS.theme),
      description: indexFor(HEADER_KEYS.description),
      workTime: indexFor(HEADER_KEYS.workTime),
      commuteTime: indexFor(HEADER_KEYS.commuteTime)
    }
  };
}

/**
 * Berechnet Arbeitszeit
 */
function calculateWorkTime(arrival, departure) {
  if (!arrival || !departure) return '';
  
  const arrivalTime = parseTime(arrival);
  const departureTime = parseTime(departure);
  
  if (!arrivalTime || !departureTime) return '';
  
  let diff = departureTime - arrivalTime;
  if (diff < 0) diff += 24 * 60;
  
  const hours = Math.floor(diff / 60);
  const minutes = diff % 60;
  
  return `${hours}h ${minutes}m`;
}

/**
 * Berechnet Pendelzeit
 */
function calculateCommuteTime(homeDeparture, workArrival, workDeparture, homeArrival) {
  let total = 0;
  
  if (homeDeparture && workArrival) {
    const start = parseTime(homeDeparture);
    const end = parseTime(workArrival);
    if (start && end) {
      let diff = end - start;
      if (diff < 0) diff += 24 * 60;
      total += diff;
    }
  }
  
  if (workDeparture && homeArrival) {
    const start = parseTime(workDeparture);
    const end = parseTime(homeArrival);
    if (start && end) {
      let diff = end - start;
      if (diff < 0) diff += 24 * 60;
      total += diff;
    }
  }
  
  if (total === 0) return '';
  
  const hours = Math.floor(total / 60);
  const minutes = total % 60;
  
  return `${hours}h ${minutes}m`;
}

/**
 * Parst Zeitstring zu Minuten seit Mitternacht
 */
function parseTime(timeStr) {
  if (!timeStr) return null;
  
  const parts = timeStr.split(':');
  if (parts.length !== 2) return null;
  
  const hours = parseInt(parts[0], 10);
  const minutes = parseInt(parts[1], 10);
  
  if (isNaN(hours) || isNaN(minutes)) return null;
  
  return hours * 60 + minutes;
}
