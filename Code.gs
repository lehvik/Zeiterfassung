/**
 * Google Apps Script Backend für Arbeitszeit Tracker
 * Version 2.0 - Verbessertes Datenmodell
 */

// Name des Tabellenblattes
const SHEET_NAME = 'Zeiterfassung';
const THEMES_SHEET = 'Themen'; // Für Projekt/Thema-Vorschläge
const STATS_SHEET = 'Statistik';
const MAIN_HEADERS = [
  'KW',
  'Wochentag',
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
  'Pendelzeit',
  'Status'
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
const HEADERS_V2 = [
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
const HEADER_KEYS = {
  week: 'KW',
  weekday: 'Wochentag',
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
  commuteTime: 'Pendelzeit',
  status: 'Status'
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
    let statsSheet = spreadsheet.getSheetByName(STATS_SHEET);
    
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
    
    if (!statsSheet) {
      statsSheet = spreadsheet.insertSheet(STATS_SHEET);
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
      description: e.parameter.description || '',
      status: e.parameter.status || ''
    };
      
      const result = saveTimeData(sheet, themesSheet, data);
      updateStatistics(sheet, statsSheet);
      
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
        description: e.parameter.description || '',
        status: e.parameter.status || ''
      };
      
      const result = saveTimeData(sheet, themesSheet, data);
      updateStatistics(sheet, statsSheet);
      
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
 * Parst Datum zu Date-Objekt
 */
function parseDateValue(value) {
  if (!value) return null;
  
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }
  
  const str = String(value).trim();
  if (!str) return null;
  
  const parts = str.split('.');
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const year = parseInt(parts[2], 10);
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
      return new Date(year, month, day);
    }
  }
  
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  
  return null;
}

/**
 * ISO Kalenderwoche (1-53)
 */
function getIsoWeekNumber(dateObj) {
  const d = new Date(dateObj.getTime());
  d.setHours(0, 0, 0, 0);
  // Donnerstag der aktuellen Woche
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
  const week1 = new Date(d.getFullYear(), 0, 4);
  return 1 + Math.round(
    ((d - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7
  );
}

/**
 * Wochentag kurz (Mo, Di, Mi...)
 */
function getWeekdayShort(dateObj) {
  const days = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
  return days[dateObj.getDay()];
}

/**
 * Erstellt Mapping von Header-Name zu Spaltenindex (0-basiert)
 */
function getHeaderMap(sheet) {
  const columnCount = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  const headerRow = sheet.getRange(1, 1, 1, columnCount).getDisplayValues()[0];
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
  
  // V2 Schema (11 Spalten) -> KW/Wochentag vorne, Status hinten
  if (matches(normalized, HEADERS_V2.map(h => normalizeHeader(h)))) {
    sheet.insertColumns(1, 2);
    sheet.insertColumnAfter(sheet.getLastColumn());
    initializeSheet(sheet);
    fillDerivedColumns(sheet);
    return;
  }
  
  // Altes Schema migrieren
  if (matches(normalized, LEGACY_HEADERS.map(h => normalizeHeader(h)))) {
    // Pendel-Art (Spalte C) einfügen
    sheet.insertColumns(3, 1);
    // Thema/Projekt + Beschreibung vor Arbeitszeit einfügen
    sheet.insertColumns(8, 2);
    // KW/Wochentag vorne
    sheet.insertColumns(1, 2);
    // Status hinten
    sheet.insertColumnAfter(sheet.getLastColumn());
    initializeSheet(sheet);
    fillDerivedColumns(sheet);
    return;
  }
  
  // Leeres/ungewöhnliches Sheet: Header setzen
  const hasAnyHeader = normalized.some(v => v !== '');
  if (!hasAnyHeader) {
    initializeSheet(sheet);
    return;
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
  sheet.setColumnWidth(1, 70);  // KW
  sheet.setColumnWidth(2, 90);  // Wochentag
  sheet.setColumnWidth(3, 110); // Datum
  sheet.setColumnWidth(4, 120); // Modus
  sheet.setColumnWidth(5, 120); // Pendel-Art
  sheet.setColumnWidth(10, 150); // Thema/Projekt
  sheet.setColumnWidth(11, 200); // Beschreibung
  sheet.setColumnWidth(14, 100); // Status
  
  // KW Format "KW 06"
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('"KW "00');
}

/**
 * Füllt KW/Wochentag für bestehende Daten und setzt Wochen-Trennlinien
 */
function fillDerivedColumns(sheet) {
  const headerMap = getHeaderMap(sheet);
  const colWeek = findHeaderIndex(headerMap, [HEADER_KEYS.week, 'Kalenderwoche']);
  const colWeekday = findHeaderIndex(headerMap, [HEADER_KEYS.weekday]);
  const colDate = findHeaderIndex(headerMap, [HEADER_KEYS.date]);
  const columnCount = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  
  if (colDate === undefined) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const dateValues = sheet.getRange(2, colDate + 1, lastRow - 1, 1).getValues();
  const weekValues = [];
  const weekdayValues = [];
  
  dateValues.forEach((row, idx) => {
    const dateObj = parseDateValue(row[0]);
    if (dateObj) {
      weekValues.push([getIsoWeekNumber(dateObj)]);
      weekdayValues.push([getWeekdayShort(dateObj)]);
    } else {
      weekValues.push(['']);
      weekdayValues.push(['']);
    }
  });
  
  if (colWeek !== undefined) {
    sheet.getRange(2, colWeek + 1, weekValues.length, 1).setValues(weekValues);
  }
  if (colWeekday !== undefined) {
    sheet.getRange(2, colWeekday + 1, weekdayValues.length, 1).setValues(weekdayValues);
  }
  
  // Trennlinien pro Woche
  for (let r = 2; r <= lastRow; r++) {
    if (colWeek !== undefined) {
      applyWeekSeparator(sheet, r, colWeek, columnCount);
    }
  }
}

/**
 * Setzt Trennlinie vor neuer Kalenderwoche
 */
function applyWeekSeparator(sheet, rowIndex, colWeek, columnCount) {
  if (rowIndex <= 2) return;
  
  const current = sheet.getRange(rowIndex, colWeek + 1).getValue();
  const previous = sheet.getRange(rowIndex - 1, colWeek + 1).getValue();
  
  if (!current || !previous) return;
  
  if (current !== previous) {
    sheet.getRange(rowIndex, 1, 1, columnCount).setBorder(true, null, null, null, null, null);
  }
}

/**
 * Statistik-Sheet aktualisieren
 */
function updateStatistics(mainSheet, statsSheet) {
  if (!mainSheet || !statsSheet) return;
  
  const headerMap = getHeaderMap(mainSheet);
  const colDate = findHeaderIndex(headerMap, [HEADER_KEYS.date]);
  const colWorkTime = findHeaderIndex(headerMap, [HEADER_KEYS.workTime]);
  const colCommuteTime = findHeaderIndex(headerMap, [HEADER_KEYS.commuteTime]);
  const colWorkArrival = findHeaderIndex(headerMap, [HEADER_KEYS.workArrival, 'Ankunft\nArbeit']);
  const colWorkDeparture = findHeaderIndex(headerMap, [HEADER_KEYS.workDeparture]);
  const colTheme = findHeaderIndex(headerMap, [HEADER_KEYS.theme, 'Thema / Projekt']);
  const colStatus = findHeaderIndex(headerMap, [HEADER_KEYS.status]);
  
  if (colDate === undefined) return;
  
  const data = mainSheet.getDataRange().getValues();
  if (data.length < 2) {
    statsSheet.clearContents();
    return;
  }
  
  const weekly = {};
  const monthly = {};
  const yearly = {};
  const projects = {};
  const quarterProjects = {};
  const monthlyProjects = {};
  const modeStats = {};
  const weekdayStats = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateObj = parseDateValue(row[colDate]);
    if (!dateObj) continue;
    
    const status = colStatus !== undefined ? String(row[colStatus] || '').trim() : '';
    
    const workMinutes = computeWorkMinutes(row, colWorkTime, colWorkArrival, colWorkDeparture);
    const commuteMinutes = computeDurationMinutes(row[colCommuteTime]);
    const startMinutes = parseTime(colWorkArrival !== undefined ? row[colWorkArrival] : '');
    const endMinutes = parseTime(colWorkDeparture !== undefined ? row[colWorkDeparture] : '');
    
    const week = getIsoWeekNumber(dateObj);
    const weekYear = getIsoWeekYear(dateObj);
    const weekKey = `${weekYear}-W${String(week).padStart(2, '0')}`;
    const monthKey = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}`;
    const yearKey = String(dateObj.getFullYear());
    const quarterKey = `${dateObj.getFullYear()}-Q${getQuarter(dateObj)}`;
    const weekdayKey = getWeekdayShort(dateObj);
    
    const w = weekly[weekKey] || createStatsBucket(weekYear, week);
    const m = monthly[monthKey] || createStatsBucket(dateObj.getFullYear(), dateObj.getMonth() + 1);
    const y = yearly[yearKey] || createStatsBucket(dateObj.getFullYear(), dateObj.getFullYear());
    
    if (status === 'Fehlt') {
      w.missingDays++;
      m.missingDays++;
      y.missingDays++;
    } else {
      if (workMinutes > 0) {
        w.workMinutes += workMinutes;
        m.workMinutes += workMinutes;
        y.workMinutes += workMinutes;
        w.workDays++;
        m.workDays++;
        y.workDays++;
        
        const rawTheme = colTheme !== undefined ? row[colTheme] : '';
        const themeName = normalizeThemeDisplay(rawTheme) || 'Ohne Projekt';
        const themeKey = normalizeThemeKey(themeName);
        if (!projects[themeKey]) {
          projects[themeKey] = { name: themeName, minutes: 0 };
        }
        projects[themeKey].minutes += workMinutes;
        
        if (!quarterProjects[quarterKey]) {
          quarterProjects[quarterKey] = {};
        }
        quarterProjects[quarterKey][themeKey] = (quarterProjects[quarterKey][themeKey] || 0) + workMinutes;
        
        if (!monthlyProjects[monthKey]) {
          monthlyProjects[monthKey] = {};
        }
        monthlyProjects[monthKey][themeKey] = (monthlyProjects[monthKey][themeKey] || 0) + workMinutes;
        
        const mode = String(row[colMode] || '').trim();
        if (mode) {
          modeStats[mode] = (modeStats[mode] || 0) + workMinutes;
        }
        
        if (!weekdayStats[weekdayKey]) {
          weekdayStats[weekdayKey] = { minutes: 0, days: 0 };
        }
        weekdayStats[weekdayKey].minutes += workMinutes;
        weekdayStats[weekdayKey].days += 1;
      }
      
      if (commuteMinutes > 0) {
        w.commuteMinutes += commuteMinutes;
        m.commuteMinutes += commuteMinutes;
        y.commuteMinutes += commuteMinutes;
      }
      
      if (startMinutes !== null) {
        w.startMinutes += startMinutes;
        w.startCount++;
        m.startMinutes += startMinutes;
        m.startCount++;
        y.startMinutes += startMinutes;
        y.startCount++;
      }
      
      if (endMinutes !== null) {
        w.endMinutes += endMinutes;
        w.endCount++;
        m.endMinutes += endMinutes;
        m.endCount++;
        y.endMinutes += endMinutes;
        y.endCount++;
      }
      
      if (status === 'Schätzung') {
        w.estimateDays++;
        m.estimateDays++;
        y.estimateDays++;
      }
    }
    
    weekly[weekKey] = w;
    monthly[monthKey] = m;
    yearly[yearKey] = y;
  }
  
  writeStatisticsSheet(
    statsSheet,
    weekly,
    monthly,
    yearly,
    projects,
    quarterProjects,
    monthlyProjects,
    modeStats,
    weekdayStats
  );
}

function createStatsBucket(year, period) {
  return {
    year: year,
    period: period,
    workMinutes: 0,
    commuteMinutes: 0,
    startMinutes: 0,
    endMinutes: 0,
    startCount: 0,
    endCount: 0,
    workDays: 0,
    estimateDays: 0,
    missingDays: 0
  };
}

function computeWorkMinutes(row, colWorkTime, colWorkArrival, colWorkDeparture) {
  let minutes = computeDurationMinutes(colWorkTime !== undefined ? row[colWorkTime] : '');
  if (minutes > 0) return minutes;
  
  const start = parseTime(colWorkArrival !== undefined ? row[colWorkArrival] : '');
  const end = parseTime(colWorkDeparture !== undefined ? row[colWorkDeparture] : '');
  if (start === null || end === null) return 0;
  
  let diff = end - start;
  if (diff < 0) diff += 24 * 60;
  return diff;
}

function computeDurationMinutes(value) {
  if (!value) return 0;
  
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value.getHours() * 60 + value.getMinutes();
  }
  
  const str = String(value).trim();
  if (!str) return 0;
  
  const hourMatch = str.match(/(\d+)\s*h/i);
  const minuteMatch = str.match(/(\d+)\s*m/i);
  if (hourMatch || minuteMatch) {
    const h = hourMatch ? parseInt(hourMatch[1], 10) : 0;
    const m = minuteMatch ? parseInt(minuteMatch[1], 10) : 0;
    return h * 60 + m;
  }
  
  if (str.includes(':')) {
    const parsed = parseTime(str);
    return parsed !== null ? parsed : 0;
  }
  
  const num = Number(str);
  return isNaN(num) ? 0 : Math.round(num * 60);
}

function writeStatisticsSheet(statsSheet, weekly, monthly, yearly, projects, quarterProjects, monthlyProjects, modeStats, weekdayStats) {
  statsSheet.clearContents();
  statsSheet.clearCharts();
  statsSheet.setFrozenRows(1);
  
  const monthlyOverviewRows = Object.values(monthly)
    .sort((a, b) => a.year === b.year ? a.period - b.period : a.year - b.year)
    .map(bucket => [
      `${bucket.year}-${String(bucket.period).padStart(2, '0')}`,
      minutesToHours(bucket.workMinutes),
      minutesToHours(bucket.commuteMinutes)
    ]);

  const weeklyTitle = [['Wochenstatistik']];
  const weeklyHeader = [[
    'Jahr', 'KW', 'Arbeitsstunden', 'Ø/Tag', 'Pendelstunden', 'Ø Pendel/Tag',
    'Ø Start', 'Ø Ende', 'Arbeitstage', 'Schätzung', 'Fehlt'
  ]];
  
  const weeklyRows = Object.values(weekly)
    .sort((a, b) => a.year === b.year ? b.period - a.period : b.year - a.year)
    .map(bucket => [
      bucket.year,
      bucket.period,
      minutesToHours(bucket.workMinutes),
      bucket.workDays ? minutesToHours(bucket.workMinutes / bucket.workDays) : '',
      minutesToHours(bucket.commuteMinutes),
      bucket.workDays ? minutesToHours(bucket.commuteMinutes / bucket.workDays) : '',
      bucket.startCount ? minutesToHours(bucket.startMinutes / bucket.startCount) : '',
      bucket.endCount ? minutesToHours(bucket.endMinutes / bucket.endCount) : '',
      bucket.workDays,
      bucket.estimateDays,
      bucket.missingDays
    ]);
  
  const monthlyTitle = [['Monatsstatistik']];
  const monthlyHeader = [[
    'Monat', 'Arbeitsstunden', 'Ø/Tag', 'Pendelstunden', 'Ø Pendel/Tag',
    'Ø Start', 'Ø Ende', 'Arbeitstage', 'Schätzung', 'Fehlt'
  ]];
  
  const monthlyRows = Object.values(monthly)
    .sort((a, b) => a.year === b.year ? b.period - a.period : b.year - a.year)
    .map(bucket => [
      `${bucket.year}-${String(bucket.period).padStart(2, '0')}`,
      minutesToHours(bucket.workMinutes),
      bucket.workDays ? minutesToHours(bucket.workMinutes / bucket.workDays) : '',
      minutesToHours(bucket.commuteMinutes),
      bucket.workDays ? minutesToHours(bucket.commuteMinutes / bucket.workDays) : '',
      bucket.startCount ? minutesToHours(bucket.startMinutes / bucket.startCount) : '',
      bucket.endCount ? minutesToHours(bucket.endMinutes / bucket.endCount) : '',
      bucket.workDays,
      bucket.estimateDays,
      bucket.missingDays
    ]);
  
  const yearlyTitle = [['Jahresstatistik']];
  const yearlyHeader = [[
    'Jahr', 'Arbeitsstunden', 'Ø/Tag', 'Pendelstunden', 'Ø Pendel/Tag',
    'Ø Start', 'Ø Ende', 'Arbeitstage', 'Schätzung', 'Fehlt'
  ]];
  
  const yearlyRows = Object.values(yearly)
    .sort((a, b) => b.year - a.year)
    .map(bucket => [
      bucket.year,
      minutesToHours(bucket.workMinutes),
      bucket.workDays ? minutesToHours(bucket.workMinutes / bucket.workDays) : '',
      minutesToHours(bucket.commuteMinutes),
      bucket.workDays ? minutesToHours(bucket.commuteMinutes / bucket.workDays) : '',
      bucket.startCount ? minutesToHours(bucket.startMinutes / bucket.startCount) : '',
      bucket.endCount ? minutesToHours(bucket.endMinutes / bucket.endCount) : '',
      bucket.workDays,
      bucket.estimateDays,
      bucket.missingDays
    ]);
  
  const projectRows = buildProjectRows(projects, 8);
  const quarterTable = buildQuarterTable(projects, quarterProjects, 5);
  const projectTrend = buildProjectTrendTable(projects, monthlyProjects, 3);
  const modeRows = buildModeRows(modeStats);
  const weekdayRows = buildWeekdayRows(weekdayStats);

  let row = 1;
  const overviewTitle = [['Überblick (Monat)']];
  const overviewHeader = [['Monat', 'Arbeitsstunden', 'Pendelstunden']];
  statsSheet.getRange(row, 1, 1, overviewTitle[0].length).setValues(overviewTitle);
  statsSheet.getRange(row, 1, 1, overviewTitle[0].length).setFontWeight('bold');
  row += 1;
  statsSheet.getRange(row, 1, 1, overviewHeader[0].length).setValues(overviewHeader);
  statsSheet.getRange(row, 1, 1, overviewHeader[0].length).setFontWeight('bold');
  row += 1;
  const overviewStart = row;
  if (monthlyOverviewRows.length > 0) {
    statsSheet.getRange(row, 1, monthlyOverviewRows.length, overviewHeader[0].length).setValues(monthlyOverviewRows);
    row += monthlyOverviewRows.length;
  }
  const overviewEnd = row - 1;
  
  // Chart: Monatsüberblick
  if (overviewEnd >= overviewStart) {
    const range = statsSheet.getRange(overviewStart - 1, 1, overviewEnd - overviewStart + 2, 3);
    const chart = statsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(range)
      .setOption('title', 'Arbeits- und Pendelstunden (Monat)')
      .setOption('legend', { position: 'bottom' })
      .setOption('curveType', 'function')
      .setOption('height', 240)
      .setPosition(2, 6, 0, 0)
      .build();
    statsSheet.insertChart(chart);
  }
  
  row += 2;
  statsSheet.getRange(row, 1, 1, weeklyTitle[0].length).setValues(weeklyTitle);
  statsSheet.getRange(row, 1, 1, weeklyTitle[0].length).setFontWeight('bold');
  row += 1;
  statsSheet.getRange(row, 1, 1, weeklyHeader[0].length).setValues(weeklyHeader);
  statsSheet.getRange(row, 1, 1, weeklyHeader[0].length).setFontWeight('bold');
  row += 1;
  const weeklyDataStart = row;
  if (weeklyRows.length > 0) {
    statsSheet.getRange(row, 1, weeklyRows.length, weeklyHeader[0].length).setValues(weeklyRows);
    row += weeklyRows.length;
  }
  const weeklyDataEnd = row - 1;
  
  row += 2;
  statsSheet.getRange(row, 1, 1, monthlyTitle[0].length).setValues(monthlyTitle);
  statsSheet.getRange(row, 1, 1, monthlyTitle[0].length).setFontWeight('bold');
  row += 1;
  statsSheet.getRange(row, 1, 1, monthlyHeader[0].length).setValues(monthlyHeader);
  statsSheet.getRange(row, 1, 1, monthlyHeader[0].length).setFontWeight('bold');
  row += 1;
  const monthlyDataStart = row;
  if (monthlyRows.length > 0) {
    statsSheet.getRange(row, 1, monthlyRows.length, monthlyHeader[0].length).setValues(monthlyRows);
  }
  const monthlyDataEnd = monthlyDataStart + monthlyRows.length - 1;
  
  row = monthlyDataEnd + 3;
  statsSheet.getRange(row, 1, 1, yearlyTitle[0].length).setValues(yearlyTitle);
  statsSheet.getRange(row, 1, 1, yearlyTitle[0].length).setFontWeight('bold');
  row += 1;
  statsSheet.getRange(row, 1, 1, yearlyHeader[0].length).setValues(yearlyHeader);
  statsSheet.getRange(row, 1, 1, yearlyHeader[0].length).setFontWeight('bold');
  row += 1;
  const yearlyDataStart = row;
  if (yearlyRows.length > 0) {
    statsSheet.getRange(row, 1, yearlyRows.length, yearlyHeader[0].length).setValues(yearlyRows);
  }
  const yearlyDataEnd = yearlyDataStart + yearlyRows.length - 1;
  
  row = yearlyDataEnd + 3;
  statsSheet.getRange(row, 1, 1, 1).setValues([['Projektverteilung']]);
  statsSheet.getRange(row, 1, 1, 1).setFontWeight('bold');
  row += 1;
  statsSheet.getRange(row, 1, 1, 2).setValues([['Projekt', 'Arbeitsstunden']]);
  statsSheet.getRange(row, 1, 1, 2).setFontWeight('bold');
  row += 1;
  const projectStart = row;
  if (projectRows.length > 0) {
    statsSheet.getRange(row, 1, projectRows.length, 2).setValues(projectRows);
    row += projectRows.length;
  }
  const projectEnd = row - 1;
  
  if (projectEnd >= projectStart) {
    const range = statsSheet.getRange(projectStart - 1, 1, projectEnd - projectStart + 2, 2);
    const chart = statsSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(range)
      .setOption('title', 'Projektverteilung')
      .setOption('pieHole', 0.35)
      .setOption('legend', { position: 'right' })
      .setPosition(projectStart - 1, 6, 0, 0)
      .build();
    statsSheet.insertChart(chart);
  }
  
  row += 2;
  statsSheet.getRange(row, 1, 1, 1).setValues([['Projekt nach Quartal']]);
  statsSheet.getRange(row, 1, 1, 1).setFontWeight('bold');
  row += 1;
  if (quarterTable.header.length > 0) {
    statsSheet.getRange(row, 1, 1, quarterTable.header.length).setValues([quarterTable.header]);
    statsSheet.getRange(row, 1, 1, quarterTable.header.length).setFontWeight('bold');
    row += 1;
  }
  const quarterStart = row;
  if (quarterTable.rows.length > 0) {
    statsSheet.getRange(row, 1, quarterTable.rows.length, quarterTable.header.length).setValues(quarterTable.rows);
    row += quarterTable.rows.length;
  }
  const quarterEnd = row - 1;
  
  if (quarterEnd >= quarterStart && quarterTable.header.length > 1) {
    const range = statsSheet.getRange(quarterStart - 1, 1, quarterEnd - quarterStart + 2, quarterTable.header.length);
    const chart = statsSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(range)
      .setOption('title', 'Arbeitsstunden je Projekt (Quartal)')
      .setOption('isStacked', true)
      .setOption('legend', { position: 'right' })
      .setPosition(quarterStart - 1, 6, 0, 0)
      .build();
    statsSheet.insertChart(chart);
  }
  
  row = quarterEnd + 3;
  statsSheet.getRange(row, 1, 1, 1).setValues([['Projekttrend (Monat)']]);
  statsSheet.getRange(row, 1, 1, 1).setFontWeight('bold');
  row += 1;
  if (projectTrend.header.length > 0) {
    statsSheet.getRange(row, 1, 1, projectTrend.header.length).setValues([projectTrend.header]);
    statsSheet.getRange(row, 1, 1, projectTrend.header.length).setFontWeight('bold');
    row += 1;
  }
  const trendStart = row;
  if (projectTrend.rows.length > 0) {
    statsSheet.getRange(row, 1, projectTrend.rows.length, projectTrend.header.length).setValues(projectTrend.rows);
    row += projectTrend.rows.length;
  }
  const trendEnd = row - 1;
  
  if (trendEnd >= trendStart && projectTrend.header.length > 1) {
    const range = statsSheet.getRange(trendStart - 1, 1, trendEnd - trendStart + 2, projectTrend.header.length);
    const chart = statsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(range)
      .setOption('title', 'Arbeitsstunden pro Projekt (Monat)')
      .setOption('legend', { position: 'bottom' })
      .setOption('curveType', 'function')
      .setOption('height', 240)
      .setPosition(trendStart - 1, 8, 0, 0)
      .build();
    statsSheet.insertChart(chart);
  }
  
  row = trendEnd + 3;
  statsSheet.getRange(row, 1, 1, 1).setValues([['Modus-Verteilung']]);
  statsSheet.getRange(row, 1, 1, 1).setFontWeight('bold');
  row += 1;
  statsSheet.getRange(row, 1, 1, 2).setValues([['Modus', 'Arbeitsstunden']]);
  statsSheet.getRange(row, 1, 1, 2).setFontWeight('bold');
  row += 1;
  const modeStart = row;
  if (modeRows.length > 0) {
    statsSheet.getRange(row, 1, modeRows.length, 2).setValues(modeRows);
    row += modeRows.length;
  }
  const modeEnd = row - 1;
  
  if (modeEnd >= modeStart) {
    const range = statsSheet.getRange(modeStart - 1, 1, modeEnd - modeStart + 2, 2);
    const chart = statsSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(range)
      .setOption('title', 'Arbeitsstunden nach Modus')
      .setOption('legend', { position: 'right' })
      .setOption('pieHole', 0.3)
      .setPosition(modeStart - 1, 4, 0, 0)
      .build();
    statsSheet.insertChart(chart);
  }
  
  row = modeEnd + 3;
  statsSheet.getRange(row, 1, 1, 1).setValues([['Wochentag-Muster']]);
  statsSheet.getRange(row, 1, 1, 1).setFontWeight('bold');
  row += 1;
  statsSheet.getRange(row, 1, 1, 2).setValues([['Wochentag', 'Ø Arbeitsstunden']]);
  statsSheet.getRange(row, 1, 1, 2).setFontWeight('bold');
  row += 1;
  const weekdayStart = row;
  if (weekdayRows.length > 0) {
    statsSheet.getRange(row, 1, weekdayRows.length, 2).setValues(weekdayRows);
    row += weekdayRows.length;
  }
  const weekdayEnd = row - 1;
  
  if (weekdayEnd >= weekdayStart) {
    const range = statsSheet.getRange(weekdayStart - 1, 1, weekdayEnd - weekdayStart + 2, 2);
    const chart = statsSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(range)
      .setOption('title', 'Ø Arbeitsstunden nach Wochentag')
      .setOption('legend', { position: 'none' })
      .setOption('height', 240)
      .setPosition(weekdayStart - 1, 4, 0, 0)
      .build();
    statsSheet.insertChart(chart);
  }
  
  // Formatierungen
  if (weeklyDataEnd >= weeklyDataStart) {
    statsSheet.getRange(weeklyDataStart, 2, weeklyDataEnd - weeklyDataStart + 1, 1).setNumberFormat('00');
    statsSheet.getRange(weeklyDataStart, 3, weeklyDataEnd - weeklyDataStart + 1, 6).setNumberFormat('0.0');
  }
  if (monthlyDataEnd >= monthlyDataStart) {
    statsSheet.getRange(monthlyDataStart, 2, monthlyDataEnd - monthlyDataStart + 1, 6).setNumberFormat('0.0');
  }
  if (yearlyDataEnd >= yearlyDataStart) {
    statsSheet.getRange(yearlyDataStart, 2, yearlyDataEnd - yearlyDataStart + 1, 6).setNumberFormat('0.0');
  }
  if (overviewEnd >= overviewStart) {
    statsSheet.getRange(overviewStart, 2, overviewEnd - overviewStart + 1, 2).setNumberFormat('0.0');
  }
  if (projectEnd >= projectStart) {
    statsSheet.getRange(projectStart, 2, projectEnd - projectStart + 1, 1).setNumberFormat('0.0');
  }
  if (quarterEnd >= quarterStart && quarterTable.header.length > 1) {
    statsSheet.getRange(quarterStart, 2, quarterEnd - quarterStart + 1, quarterTable.header.length - 1).setNumberFormat('0.0');
  }
  if (trendEnd >= trendStart && projectTrend.header.length > 1) {
    statsSheet.getRange(trendStart, 2, trendEnd - trendStart + 1, projectTrend.header.length - 1).setNumberFormat('0.0');
  }
  if (modeEnd >= modeStart) {
    statsSheet.getRange(modeStart, 2, modeEnd - modeStart + 1, 1).setNumberFormat('0.0');
  }
  if (weekdayEnd >= weekdayStart) {
    statsSheet.getRange(weekdayStart, 2, weekdayEnd - weekdayStart + 1, 1).setNumberFormat('0.0');
  }
  statsSheet.autoResizeColumns(1, 11);
}

function minutesToHours(valueMinutes) {
  if (valueMinutes === '' || valueMinutes === null || valueMinutes === undefined) return '';
  if (typeof valueMinutes !== 'number') return valueMinutes;
  return Math.round((valueMinutes / 60) * 10) / 10;
}

function getIsoWeekYear(dateObj) {
  const d = new Date(dateObj.getTime());
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
  return d.getFullYear();
}

function getQuarter(dateObj) {
  return Math.floor(dateObj.getMonth() / 3) + 1;
}

function buildProjectRows(projects, topN) {
  const list = Object.values(projects)
    .map(p => ({ name: p.name, minutes: p.minutes }))
    .sort((a, b) => b.minutes - a.minutes);
  
  if (list.length === 0) return [];
  
  const rows = [];
  const main = list.slice(0, topN);
  let otherMinutes = 0;
  
  main.forEach(item => {
    rows.push([item.name, minutesToHours(item.minutes)]);
  });
  
  list.slice(topN).forEach(item => {
    otherMinutes += item.minutes;
  });
  
  if (otherMinutes > 0) {
    rows.push(['Sonstige', minutesToHours(otherMinutes)]);
  }
  
  return rows;
}

function buildQuarterTable(projects, quarterProjects, topN) {
  const list = Object.entries(projects)
    .map(([key, value]) => ({ key: key, name: value.name, minutes: value.minutes }))
    .sort((a, b) => b.minutes - a.minutes);
  
  const top = list.slice(0, topN);
  const topKeys = top.map(t => t.key);
  const header = ['Quartal', ...top.map(t => t.name), 'Sonstige'];
  
  const quarters = Object.keys(quarterProjects).sort();
  const rows = quarters.map(q => {
    const projectMap = quarterProjects[q] || {};
    let sumTop = 0;
    const values = topKeys.map(k => {
      const minutes = projectMap[k] || 0;
      sumTop += minutes;
      return minutesToHours(minutes);
    });
    const totalMinutes = Object.values(projectMap).reduce((a, b) => a + b, 0);
    const otherMinutes = Math.max(0, totalMinutes - sumTop);
    return [q, ...values, minutesToHours(otherMinutes)];
  });
  
  return { header, rows };
}

function buildProjectTrendTable(projects, monthlyProjects, topN) {
  const list = Object.entries(projects)
    .map(([key, value]) => ({ key: key, name: value.name, minutes: value.minutes }))
    .sort((a, b) => b.minutes - a.minutes);
  
  const top = list.slice(0, topN);
  const topKeys = top.map(t => t.key);
  const header = ['Monat', ...top.map(t => t.name), 'Sonstige'];
  
  const months = Object.keys(monthlyProjects).sort();
  const rows = months.map(month => {
    const projectMap = monthlyProjects[month] || {};
    let sumTop = 0;
    const values = topKeys.map(k => {
      const minutes = projectMap[k] || 0;
      sumTop += minutes;
      return minutesToHours(minutes);
    });
    const totalMinutes = Object.values(projectMap).reduce((a, b) => a + b, 0);
    const otherMinutes = Math.max(0, totalMinutes - sumTop);
    return [month, ...values, minutesToHours(otherMinutes)];
  });
  
  return { header, rows };
}

function buildModeRows(modeStats) {
  const entries = Object.entries(modeStats || {})
    .map(([mode, minutes]) => ({ mode, minutes }))
    .sort((a, b) => b.minutes - a.minutes);
  
  return entries.map(item => [item.mode, minutesToHours(item.minutes)]);
}

function buildWeekdayRows(weekdayStats) {
  const order = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
  const rows = [];
  
  order.forEach(day => {
    const data = weekdayStats[day];
    if (!data || !data.days) return;
    rows.push([day, minutesToHours(data.minutes / data.days)]);
  });
  
  return rows;
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
  const dateObj = parseDateValue(data.date);
  const allData = sheet.getDataRange().getValues();
  const headerMap = getHeaderMap(sheet);
  const columnCount = Math.max(sheet.getLastColumn(), MAIN_HEADERS.length);
  
  const idxWeek = findHeaderIndex(headerMap, [HEADER_KEYS.week, 'Kalenderwoche']);
  const idxWeekday = findHeaderIndex(headerMap, [HEADER_KEYS.weekday]);
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
  const idxStatus = findHeaderIndex(headerMap, [HEADER_KEYS.status]);
  
  const colWeek = idxWeek !== undefined ? idxWeek : defaultIndex(HEADER_KEYS.week);
  const colWeekday = idxWeekday !== undefined ? idxWeekday : defaultIndex(HEADER_KEYS.weekday);
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
  const colStatus = idxStatus !== undefined ? idxStatus : defaultIndex(HEADER_KEYS.status);
  
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

  if (colDate !== undefined && date) {
    sheet.getRange(rowIndex, colDate + 1).setValue(date);
  }
  if (dateObj) {
    if (colWeek !== undefined) sheet.getRange(rowIndex, colWeek + 1).setValue(getIsoWeekNumber(dateObj));
    if (colWeekday !== undefined) sheet.getRange(rowIndex, colWeekday + 1).setValue(getWeekdayShort(dateObj));
  }
  
  // Für Urlaub/Krank/Abwesend/Feiertag: Nur Modus setzen, keine Zeiten
  if (data.workMode === 'Urlaub' || data.workMode === 'Krank' || data.workMode === 'Abwesend' || data.workMode === 'Feiertag') {
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
  
  if (data.status !== undefined && data.status !== null && String(data.status).trim() !== '') {
    if (colStatus !== undefined) sheet.getRange(rowIndex, colStatus + 1).setValue(data.status);
  }

  // Arbeits- und Pendelzeit immer aus aktuellen Zellen berechnen
  if (colWorkTime !== undefined || colCommuteTime !== undefined) {
    const arrival = colWorkArrival !== undefined ? sheet.getRange(rowIndex, colWorkArrival + 1).getValue() : '';
    const departure = colWorkDeparture !== undefined ? sheet.getRange(rowIndex, colWorkDeparture + 1).getValue() : '';
    if (colWorkTime !== undefined) {
      sheet.getRange(rowIndex, colWorkTime + 1).setValue(calculateWorkTime(arrival, departure));
    }
    
    const effectiveMode = data.workMode || (colMode !== undefined ? sheet.getRange(rowIndex, colMode + 1).getValue() : '');
    if (colCommuteTime !== undefined) {
      if (String(effectiveMode).trim() === 'Pendeln') {
        const homeDeparture = colHomeDeparture !== undefined ? sheet.getRange(rowIndex, colHomeDeparture + 1).getValue() : '';
        const workArrival = colWorkArrival !== undefined ? sheet.getRange(rowIndex, colWorkArrival + 1).getValue() : '';
        const workDeparture = colWorkDeparture !== undefined ? sheet.getRange(rowIndex, colWorkDeparture + 1).getValue() : '';
        const homeArrival = colHomeArrival !== undefined ? sheet.getRange(rowIndex, colHomeArrival + 1).getValue() : '';
        sheet.getRange(rowIndex, colCommuteTime + 1)
          .setValue(calculateCommuteTime(homeDeparture, workArrival, workDeparture, homeArrival));
      } else {
        sheet.getRange(rowIndex, colCommuteTime + 1).setValue('');
      }
    }
  }
  
  // Formatierung
  if (rowIndex > 1) {
    const range = sheet.getRange(rowIndex, 1, 1, columnCount);
    
    // Spezielle Formatierung für Urlaub/Krank/Abwesend
    if (data.status === 'Fehlt') {
      range.setBackground('#ffebee'); // rot (Fehlt)
    } else if (data.status === 'Schätzung') {
      range.setBackground('#ffe0e0'); // hellrot (Schätzung)
    } else if (data.workMode === 'Urlaub') {
      range.setBackground('#fff3e0');
    } else if (data.workMode === 'Krank') {
      range.setBackground('#e3f2fd'); // gedämpftes Blau
    } else if (data.workMode === 'Feiertag') {
      range.setBackground('#fff9c4');
    } else if (data.workMode === 'Abwesend') {
      range.setBackground('#f3e5f5');
    } else if (rowIndex % 2 === 0) {
      // Zebra-Streifen für normale Einträge
      range.setBackground('#f8f9fa');
    }
  }
  
  const updatedRow = sheet.getRange(rowIndex, 1, 1, columnCount).getValues()[0];
  if (colWeek !== undefined) {
    applyWeekSeparator(sheet, rowIndex, colWeek, columnCount);
  }
  
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
        'home-departure': colHomeDeparture !== undefined ? normalizeTime(allData[i][colHomeDeparture]) : '',
        'work-arrival': colWorkArrival !== undefined ? normalizeTime(allData[i][colWorkArrival]) : '',
        'work-departure': colWorkDeparture !== undefined ? normalizeTime(allData[i][colWorkDeparture]) : '',
        'home-arrival': colHomeArrival !== undefined ? normalizeTime(allData[i][colHomeArrival]) : '',
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
      week: indexFor(HEADER_KEYS.week),
      weekday: indexFor(HEADER_KEYS.weekday),
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
      commuteTime: indexFor(HEADER_KEYS.commuteTime),
      status: indexFor(HEADER_KEYS.status)
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
  
  // Date-Objekt aus Google Sheets (Zeit)
  if (Object.prototype.toString.call(timeStr) === '[object Date]' && !isNaN(timeStr)) {
    return timeStr.getHours() * 60 + timeStr.getMinutes();
  }
  
  const str = String(timeStr).trim();
  if (!str) return null;
  
  // ISO-String fallback
  if (str.includes('T')) {
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      return parsed.getHours() * 60 + parsed.getMinutes();
    }
  }
  
  const parts = str.split(':');
  if (parts.length !== 2) return null;
  
  const hours = parseInt(parts[0], 10);
  const minutes = parseInt(parts[1], 10);
  
  if (isNaN(hours) || isNaN(minutes)) return null;
  
  return hours * 60 + minutes;
}

/**
 * Normalisiert Zeitwert zu "HH:mm"
 */
function normalizeTime(value) {
  if (!value) return '';
  
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  
  const str = String(value).trim();
  if (!str) return '';
  
  if (/^\d{1,2}:\d{2}$/.test(str)) return str;
  
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'HH:mm');
  }
  
  return str;
}
