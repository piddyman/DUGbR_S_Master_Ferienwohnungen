const CONFIG = {
  // LOG_LEVEL
  LOG_LEVEL_OFF: 0, // OFF
  LOG_LEVEL_DBG: 1, // DEBUG
  LOG_LEVEL_INF: 2, // INFO
  LOG_LEVEL_WRN: 3, // WARNING
  LOG_LEVEL_ERR: 4, // ERROR
  LOG_LEVEL: 1,
  // LOG_WAY
  LOG_WAY_CONSOLE: 0,           // CONSOLE
  LOG_WAY_TOAST: 1,             // TOAST
  LOG_WAY_CONSOLE_and_TOAST: 2, // CONSOLE und TOAST
  LOG_WAY: 2,

  // VERWENDE als Wochentag den Wert, der
  // A) den errechneten Wert vom aktuellen Datum.
  USE_DAY_GIVEN_BY_CALC: 0,
  // B) übergeben worden ist (d.h. vom Formular stammt)
  USE_DAY_GIVEN_BY_FORUMLAR: 1,  
  // --
  USE_DAY_GIVEN_BY: 1,

  // TAG für USE_DAY_GIVEN_BY_CALC: Wenn der Wert GLEICH dem String "HEUTE" ist, dann automatisch auf "USE_DAY_GIVEN_BY_CALC: 0" setzen.
  DAY_CALC_STRING_AUTO: "HEUTE",

  // TAG für die Monatsberechnung.
  MONTH_CALC_STRING_CURRENT: "AKTUELLER",
  MONTH_CALC_STRING_LAST: "LETZTER",

  // ##__SPREADSHEET__##: "Master_FerienwohnungDienst" ------------------------------------ [START]
  // Sheet: "Master"
  SN_Master: "Master",
  SN_FormularResponses: "Formularantworten 3",
  // Tabelle: "Tbl_Master"
  TBL_Master: "Tbl_Master"
};  