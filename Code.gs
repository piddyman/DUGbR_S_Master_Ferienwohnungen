/**
 * MAIN-FUNCTION --> Reagiert auf Daten, die vom Formular gesendet worden sind.
 * =============
  * @param {Objekt-EreignisDaten} _e - EreignisDaten.
 */
function processFormSubmission(_e) {

  const FCT_NAME = "processFormSubmission";
  var logMsg = "";
  var logMsgPrinted = "";
  

  // START
  // =====
  // 1.CHECK: Sicherheitsprüfung am Anfang.
  // ========
  if (!_e || !_e.range) {
    // Logging
    logMsg = "ERROR - '" + FCT_NAME + "': Die Funktion wurde manuell gestartet (ohne Ereignis-Daten).";
    SpreadsheetApp.getUi().alert(logMsg);
    console.error(logMsg);
    //
    return; // Beendet die Funktion hier sicher.
  }  

  // 2. GET_SCRIPT_LOCK: Sperre sofort zu erhalten (0ms Wartezeit)!
  // ===================
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    logMsg = "WARNUNG - '" + FCT_NAME + "': Bitte warten - Eine andere Operation wird gerade noch ausgeführt.";
    console.warn(logMsg);
    SpreadsheetApp.getActive().toast(logMsg, FCT_NAME, 3);    
    SpreadsheetApp.getUi().alert(logMsg);
    return;
  }

  // 3. Erzeuge Variablen für diese Funktion.
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 4. MAIN-MAIN-MAIN
  // =================
  try {
    // Logging
    logMsg = "[START]";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // 1. Entnehme die Daten aus dem Ereignis-Objekt.
    var arrFormularValues = _e.values;
    // 2. Füge die vom Formular gesendeten und neu empfangenen Werte in die MASTER-Tabelle ein.
    newFormularValues(arrFormularValues);

    // 99. SENDE eine E-Mail.
    sendEmailNewEntryAndReportAsHtml(arrFormularValues);
    //sendEmailNewEntryAndReport(arrFormularValues);

  } catch (e) {
    // Logging
    logMsg = "Folgender Fehler ist aufgetreten: '" + e.toString() + "'.";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_ERR, FCT_NAME);

    // BenutzerInteraktion.
    SpreadsheetApp.getUi().alert(logMsg);

  } finally {
    // Logging
    logMsg = "[E_N_D]";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // RELEASE_SCRIPT_LOCK: Gibt die Sperre wieder FREI.
    // ====================    
    SpreadsheetApp.flush();
    if (lock && lock.hasLock()) {
      lock.releaseLock();
    }
  }

} // END-Funktion: 'function processFormSubmission(_e)'.

/**
 * Trage die empfangenen Werte aus dem Formular in die Master-Tabelle ein.
 * 
 * @param {array} _values - Werte aus dem Formular.
 */
function newFormularValues(_values) {

  const FCT_NAME = "newFormularValues";
  var logMsg = "";
  var logMsgPrinted = "";

  // HINWEIS: Kein eigener Script-Lock notwendig.
  // Der Lock wird bereits vom Aufrufer (processFormSubmission) gehalten.
  // Ein erneutes tryLock(0) würde sofort fehlschlagen, da LockService nicht reentrant ist.

  // START: Erstelle eine Sammel-PDF-Abrechnung.
  // =====
  try {
    // Logging
    logMsg = "[START]";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // 1. Erzeuge Konstanten für diese Funktion.
    const ssActive = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ssActive.getSheetByName(CONFIG.SN_Master);
    var range = ssActive.getRangeByName(CONFIG.TBL_Master);

    // 2. CHECK: Wurde Bereich in Sheet gefunden?
    // =========
    if (!range) {
        // CASE: NEIN, Bereich in Sheet wurde NICHT gefunden.
        // =====
        // Log-and-throw.
        logMsg = "Bereich '" + CONFIG.TBL_Master + "' in Sheet '" + CONFIG.SN_Master + "' wurde NICHT gefunden!";
        logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_ERR, FCT_NAME);
        //
        throw new Error(logMsg);
    }

    // 3. Bestimmung der Zielzeile (erste leere Zeile direkt unter dem Bereich).
    var row = range.getLastRow() + 1;
    var colStart = range.getColumn();

    // 4. Werte aus dem Formular-Array zuordnen.
    //    Hinweis: Werte werden direkt in die berechneten Spalten geschrieben.
    //    ======== 
    // 4.1 Spalte A (Zeitstempel).
    sheet.getRange(row, colStart).setValue(_values[0]);
    //
    // Logging
    logMsg = "Spalte A (Zeitstempel) = '" + _values[0] + "'.";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // 4.2 Spalte B (Objekt).
    sheet.getRange(row, colStart + 1).setValue(_values[2]);
    //
    // Logging
    logMsg = "Spalte B (Objekt) - _values[2] = '" + _values[2] + "'.";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // 4.3 Spalte F (Stunde) - _values[5] -> col + 5.
    sheet.getRange(row, colStart + 5).setValue(_values[5]);
    //
    // Logging
    logMsg = "Spalte F (Stunde) - _values[5] -> col + 5 = '" + _values[5] + "'.";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // 4.4 Spalte G (TätigkeitsID): Extrahiere nur die Zahl vor dem "-" aus _values[6] -> col + 6.
    //     Bsp.: Aus " 1 - Kontrollgang" wird "1".
    //     =====
    var doingID = "";
    if (_values[6] && _values[6].indexOf("-") !== -1) {
      doingID = _values[6].split("-")[0].trim();
    } else {
      // Fallback falls kein "-" vorhanden.
      doingID = _values[6];
    }
    sheet.getRange(row, colStart + 6).setValue(doingID);
    //
    // Logging
    logMsg = "Spalte G (TätigkeitsID) - _values[6] -> col + 6 = '" + doingID + "'.";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // 4.5 Spalte H (Dauer) - _values[7] -> col + 7.
    sheet.getRange(row, colStart + 7).setValue(_values[7]);
    //
    // Logging
    logMsg = "Spalte H (Dauer) - _values[7] -> col + 7 = '" + _values[7] + "'.";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // 4.6 Spalte I (Kommentar, Hinweis) - _values[8] -> col + 8.
    sheet.getRange(row, colStart + 8).setValue(_values[8]);
    //
    // Logging
    logMsg = "Spalte I (Kommentar, Hinweis) - _values[8] -> col + 8 = '" + _values[8] + "'.";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // 4.7 Spalte J (Arbeiter) - _values[1] -> col + 9.
    sheet.getRange(row, colStart + 9).setValue(_values[1]);
    //
    // Logging
    logMsg = "Spalte J (Arbeiter) - _values[1] -> col + 9 = '" + _values[1] + "'.";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    // 4.8 Zusätzliche Informationen berechnen (Datumswerte).
    // ======================================================
    var timeZone = Session.getScriptTimeZone(); 
    var format = "dd.MM.yyyy HH:mm:ss";
    var dateString = _values[0];

    // 4.8.1 TRY: Bestimme aus einem JavaScript-DatumsObjekt: Jahr, Monat, Tag.
    // ==========
    try {
      // 1. Macht aus dem Text "17.01.2026..." ein echtes JavaScript-Datumsobjekt.
      var dateTimeObject = Utilities.parseDate(dateString, timeZone, format);
      
      // 2. CHECK: Ist das DateTime-Objekt gültig?
      // =========
      if (dateTimeObject && !isNaN(dateTimeObject.getTime())) {
        // CASE: DateTime-Objekt ist gültig.
        // =====        
        // Spalte C: Jahr.
        var timeStampYear = dateTimeObject.getFullYear();
        sheet.getRange(row, colStart + 2).setValue(timeStampYear);

        // Spalte D: Monat.
        var timeStampMonth;
        if (_values[3] == CONFIG.MONTH_CALC_STRING_CURRENT) {
          timeStampMonth = dateTimeObject.getMonth() + 1;
        } else {
          timeStampMonth = dateTimeObject.getMonth();
        }
        sheet.getRange(row, colStart + 3).setValue(timeStampMonth);

        // Spalte E: Tag - Wert errechnet anhand es aktuellen Datums.
        // 1. CHECK: Wurde "HEUTE" übergeben und soll somit automatisch der aktuelle Tag ermittelt werden oder NICHT?
        // =========
        if (_values[4] == CONFIG.DAY_CALC_STRING_AUTO) {
          // CASE: JA, es wurde HEUTE übergeben, d.h. es soll automatisch der aktuelle Tag ermittelt werden.
          // ===== --> Setze Flag entsprechend.
          CONFIG.USE_DAY_GIVEN_BY = CONFIG.USE_DAY_GIVEN_BY_CALC;

        } else {
          // CASE: NEIN, es wurde HEUTE übergeben, d.h. es soll der Wert, der vom Formular übergeben worden ist verwendet werden.
          // ===== --> Setze Flag entsprechend.
          CONFIG.USE_DAY_GIVEN_BY = CONFIG.USE_DAY_GIVEN_BY_FORUMLAR;

        }
        // 2. Setze Wert für Spalte E.
        var timeStampDay = "";
        if (CONFIG.USE_DAY_GIVEN_BY == CONFIG.USE_DAY_GIVEN_BY_CALC) {
          // CASE: A) - Errechne den Tag.
          // =====
          timeStampDay = dateTimeObject.getDate();
        } else if (CONFIG.USE_DAY_GIVEN_BY == CONFIG.USE_DAY_GIVEN_BY_FORUMLAR) {
          // CASE: B) - Entnehme den Tag aus der Formulars-Spalte E: Tag - _values[2] -> col + 4.
          // =====
          timeStampDay = _values[4];
        } else {

          // Log-and-throw.
          logMsg = "CONFIG.USE_DAY_GIVEN_BY -> Unbekannter CASE: '" + CONFIG.USE_DAY_GIVEN_BY + "'!";
          logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_ERR, FCT_NAME);
          //
          throw new Error(logMsg);          
        }
        sheet.getRange(row, colStart + 4).setValue(timeStampDay);

        // Logging
        logMsg = "Spalte C (Jahr) - '" + timeStampYear + "', Spalte D (Monat) - '" + timeStampMonth + "', Spalte E (Tag) - '" + timeStampDay + "'.";
        logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);
      } else {
        // CASE: DateTime-Objekt ist NICHT gültig.
        // =====

        // Log-and-throw.
        logMsg = "DateTimeObject ist kein gültiges Datum!";
        logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_ERR, FCT_NAME);
        //
        throw new Error(logMsg);
      }

    } catch (e) {
      // Logging
      logMsg = "Fehler bei Utilities.parseDate. Folgender Fehler ist aufgetreten: '" + e.toString() + "'.";
      logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_ERR, FCT_NAME);

      // BenutzerInteraktion.
      SpreadsheetApp.getUi().alert(logMsg);
    }

    // 5. Benannten Bereich "Tbl_Master" um die neue Zeile erweitern.
    var updatedRange = sheet.getRange(
      range.getRow(), 
      range.getColumn(), 
      range.getNumRows() + 1, 
      range.getLastColumn()
    );
    ssActive.setNamedRange(CONFIG.TBL_Master, updatedRange);

  } catch (e) {
    // Log-and-throw.
    logMsg = "Folgender Fehler ist aufgetreten: '" + e.toString() + "'.";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_ERR, FCT_NAME);

    // BenutzerInteraktion.
    SpreadsheetApp.getUi().alert(logMsg);
    //
    throw new Error(logMsg);

  } finally {
    // Logging
    logMsg = "[E_N_D]";
    logMsgPrinted = logLOGGING(logMsg, CONFIG.LOG_LEVEL_DBG, FCT_NAME);

    SpreadsheetApp.flush();
  }

} // END-Funktion: 'function newFormularValues(_values)'.

/**
 * Sendet eine E-Mail mit dem neuen Eintrag und einem CSV-Bericht (Plaintext).
 *
 * @param {array} _values - Werte aus dem Formular.
 */
function sendEmailNewEntryAndReport(_values) {

  // Erzwingt das Schreiben aller ausstehenden Änderungen.
  SpreadsheetApp.flush();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetReportFormularResponses = ss.getSheetByName(CONFIG.SN_FormularResponses);
  var sheetReportMaster = ss.getSheetByName(CONFIG.SN_Master);
  
  
   // Umwandlung mit Semikolon als Trennzeichen.
  var valuesAsSemiColonString = _values.join(";");
  
  // 1. Header und letzte Zeile aus dem Sheet "Formularantworten 1" holen.
  // HINWEIS: Zeitstempel;Welches Objekt?;Tag des Monats?;Stunde der Uhrzeit?;Dauer in Minuten?;Beschreibung der sonstigen Tätigkeiten?;Tätigkeit?;Welcher Arbeiter?;Welcher Monat?
  // =======            0;              1;              2;                  3;                4;                                      5;         6;                7;            8
  // 1.1 Liest alle Daten in ein 2D-Array.
  var dataFormularResponses = sheetReportFormularResponses.getDataRange().getValues();
  // 1.2 Erste Zeile (Index 0 im Array).
  var valueFormularHeader = dataFormularResponses[0];             
  var valueFormularHeaderCsv = valueFormularHeader.join(";");
  // 1.3 Letzte Zeile (letzter Index im Array).
  var valueFormularRowLast = dataFormularResponses[dataFormularResponses.length - 1];
  var valueFormularRowLastCsv = valueFormularRowLast.join(";");

  // 2. Header und letzte Zeile aus dem Sheet "Master" holen.
  // HINWEIS: ZeitStempel;ObjektID;Jahr [2026..ABCD];Monat [1..12];Tag [1..31];Stunde [0..23];TätigkeitsID;Dauer [min, 0..60];Kommentar;Arbeiter
  // ========
  // 2.1 Liest alle Daten in ein 2D-Array.
  var dataMaster = sheetReportMaster.getDataRange().getValues();
  // 2.2 Erste Zeile (Index 0 im Array).
  var valueMasterHeader = dataMaster[0];
  var valueMasterHeaderCsv = valueMasterHeader.join(";");
  // 2.3 Letzte Zeile (letzter Index im Array).
  var valueMasterRowLast = dataMaster[dataMaster.length - 1];
  var valueMasterRowLastCsv = valueMasterRowLast.join(";");

  // 3. Erzeuge: E-MAIL.  
  // 3.1 Setze den E-Mail-Empfänger.
  var recipient = "info@solation.de";
  // 3.2 Setze den E-Mail-Betreff.
  var subject = "Neuer Eintrag für Ferienwohnung von : '" + _values[7] + "'.";   
  // 3.3 Erzeuge den E-Mail-Body.
  var body = "Ein neuer Eintrag für Ferienwohnung wurde am " + _values[0] + " erstellt.\n\n" +
             "DETAILS\n" +
             "FormularResponses -----------------------------------\n" +
             "valueFormularHeaderCsv: " + valueFormularHeaderCsv + "\n" +
             "valueFormularRowLastCsv: " + valueFormularRowLastCsv + "\n" +
             "-----------------------------------\n\n" + 
             "Master----------- -----------------------------------\n" +
             "valueMasterHeaderCsv: " + valueMasterHeaderCsv + "\n" +
             "valueMasterRowLastCsv: " + valueMasterRowLastCsv + "\n" +
             "-----------------------------------\n\n" + 

             "Link zur Tabelle: " + ss.getUrl();

  // 99. SENDE: E-MAIL!!!
  // ====================
  GmailApp.sendEmail(recipient, subject, body);
}

/**
 * Sendet eine HTML-formatierte E-Mail mit dem neuen Eintrag und tabellarischem Bericht.
 *
 * @param {array} _values - Werte aus dem Formular.
 */
function sendEmailNewEntryAndReportAsHtml(_values) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetReportFormularResponses = ss.getSheetByName(CONFIG.SN_FormularResponses);
  var sheetReportMaster = ss.getSheetByName(CONFIG.SN_Master);
  
  
   // Umwandlung mit Semikolon als Trennzeichen.
  var valuesAsSemiColonString = _values.join(";");
  
  // 1. Header und letzte Zeile aus dem Sheet "Formularantworten 1" holen.
  // HINWEIS: Zeitstempel;Welches Objekt?;Tag des Monats?;Stunde der Uhrzeit?;Dauer in Minuten?;Beschreibung der sonstigen Tätigkeiten?;Tätigkeit?;Welcher Arbeiter?;Welcher Monat?
  // =======            0;              1;              2;                  3;                4;                                      5;         6;                7;            8
  // 1.1 Liest alle Daten in ein 2D-Array.
  var dataFormularResponses = sheetReportFormularResponses.getDataRange().getValues();
  // 1.2 Erste Zeile (Index 0 im Array).
  var valueFormularHeader = dataFormularResponses[0];             
  var valueFormularHeaderCsv = valueFormularHeader.join(";");
  // 1.3 Letzte Zeile (letzter Index im Array).
  var valueFormularRowLast = dataFormularResponses[dataFormularResponses.length - 1];
  var valueFormularRowLastCsv = valueFormularRowLast.join(";");

  // 2. Header und letzte Zeile aus dem Sheet "Master" holen.
  // HINWEIS: ZeitStempel;ObjektID;Jahr [2026..ABCD];Monat [1..12];Tag [1..31];Stunde [0..23];TätigkeitsID;Dauer [min, 0..60];Kommentar;Arbeiter
  // ========
  // 2.1 Liest alle Daten in ein 2D-Array.
  var dataMaster = sheetReportMaster.getDataRange().getValues();
  // 2.2 Erste Zeile (Index 0 im Array).
  var valueMasterHeader = dataMaster[0];
  var valueMasterHeaderCsv = valueMasterHeader.join(";");
  // 2.3 Letzte Zeile (letzter Index im Array).
  var valueMasterRowLast = dataMaster[dataMaster.length - 1];
  var valueMasterRowLastCsv = valueMasterRowLast.join(";");

  // 3. Erzeuge: HTML E-MAIL.
  // 3.1 Setze den E-Mail-Empfänger.
  var recipient = "info@solation.de";
  // 3.2 Setze den E-Mail-Betreff.
  var subject = "Neuer Eintrag für Ferienwohnung von : '" + _values[7] + "'.";   
  // 3.3 Erzeuge den HTML-Body mit CSS-Styling.
  //     Erzeuge die Tabellen direkt aus den Arrays.
  var tableFormular = createVerticalHtmlTable(valueFormularHeader, valueFormularRowLast);
  var tableMaster = createVerticalHtmlTable(valueMasterHeader, valueMasterRowLast);
  // 3.3.1 Der HTML-Body.
  var htmlBody = 
    '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #eee; padding: 20px;">' +
      '<h2 style="color: #0056b3;">Ferienwohnung Protokoll</h2>' +
      '<p>Neuer Eintrag von: <strong>' + _values[7] + '</strong></p>' +
      
      '<h3 style="color: #666; border-bottom: 1px solid #ccc;">Details: Formularantwort</h3>' +
      tableFormular +
      
      '<h3 style="color: #666; border-bottom: 1px solid #ccc; margin-top: 25px;">Details: Master-Datenbank</h3>' +
      tableMaster +
      
      '<div style="margin-top: 30px;">' +
        '<a href="' + ss.getUrl() + '" style="background: #0056b3; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px;">Tabelle öffnen</a>' +
      '</div>' +
    '</div>';

  // Senden
  GmailApp.sendEmail(recipient, subject, "Bitte HTML-Client nutzen", { htmlBody: htmlBody });

}

/**
 * Erzeugt eine vertikale HTML-Tabelle aus einem Header- und einem Werte-Array.
 *
 * @param {array} headerArray - Array mit den Spaltenüberschriften.
 * @param {array} valueArray  - Array mit den zugehörigen Werten.
 *
 * @return {string} HTML-String der erzeugten Tabelle.
 */
function createVerticalHtmlTable(headerArray, valueArray) {
  var html = '<table style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 13px;">';
  var tz = Session.getScriptTimeZone();

  for (var i = 0; i < headerArray.length; i++) {
    // ÜBERPRÜFUNG DER LÄNGE & EXISTENZ
    // Wir prüfen: Gibt es das Element im valueArray überhaupt?
    var val = (valueArray && i < valueArray.length) ? valueArray[i] : "-";
    
    // DATUM FORMATIEREN
    if (val instanceof Date) {
      val = Utilities.formatDate(val, tz, "dd.MM.yyyy HH:mm");
    }

    var bgColor = (i % 2 === 0) ? '#ffffff' : '#f9f9f9';
    html += '<tr style="background-color: ' + bgColor + ';">';
    html += '<td style="border: 1px solid #ddd; padding: 8px; font-weight: bold; width: 35%; color: #555;">' + headerArray[i] + '</td>';
    html += '<td style="border: 1px solid #ddd; padding: 8px; color: #000;">' + (val !== "" ? val : "-") + '</td>';
    html += '</tr>';
  }
  
  html += '</table>';
  return html;
}

/**
 * Zentrale Logging-Funktion zur Vereinfachung des Codes.
 * 
 * @param {string} _logMsg - Die LogMessage.
 * @param {string} _logLevel - Der gewünschte LogLevel der LogMessage.
 * @param {string} _fctName - Der Funktionsname.
 * @param {string} _logWay - Die gewünschte Art des Loggings.
 * 
 * @return {string} Die zusammengesetzte Log-Nachricht.
 */
function logLOGGING(_logMsg, _logLevel, _fctName, _logWay = CONFIG.LOG_WAY) {

  // 1. Präfix bestimmen.
  let prefix = "INFO";
  if (_logLevel === CONFIG.LOG_LEVEL_DBG) prefix = "DEBUG";
  if (_logLevel === CONFIG.LOG_LEVEL_WRN) prefix = "WARN";
  if (_logLevel === CONFIG.LOG_LEVEL_ERR) prefix = "ERROR";

  // 2. Komplette Message erzeugen.
  const fullMsg = `${prefix} - '${_fctName}': ${_logMsg}`;

  // 3. Wenn Level == OFF oder zu niedrig, hier abbrechen.
  if ( (CONFIG.LOG_LEVEL === CONFIG.LOG_LEVEL_OFF) || (_logLevel < CONFIG.LOG_LEVEL) ) {
    return fullMsg;
  } else {
    // 1. Log ggf. in Console.
    const isConsole = (_logWay === CONFIG.LOG_WAY_CONSOLE ||_logWay === CONFIG.LOG_WAY_CONSOLE_and_TOAST);
    if (isConsole) {
      if (_logLevel === CONFIG.LOG_LEVEL_ERR) console.error(fullMsg);
      else if (_logLevel === CONFIG.LOG_LEVEL_WRN) console.warn(fullMsg);
      else console.info(fullMsg);
    }
    
    // 2. Log ggf. als Toast.
    const isToast = (_logWay === CONFIG.LOG_WAY_TOAST || _logWay === CONFIG.LOG_WAY_CONSOLE_and_TOAST);
    if (isToast) {
      const duration = (_logLevel === CONFIG.LOG_LEVEL_ERR) ? 6 : 3;
      SpreadsheetApp.getActive().toast(fullMsg, _fctName, duration);
    }

    return fullMsg;
  }
}