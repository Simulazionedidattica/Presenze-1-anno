const SPREADSHEET_ID = "1CDxFGSFvKIa-yrZahvhYbkxxZg21V1wC5H1DHCa2S90";
const SHEET_GID = 1375747828;
const NOTE_PREFIX = "PRESENZA_QR_ID:";
const TZ = "Europe/Rome";

/**
 * Endpoint per GitHub Pages.
 * Azioni:
 * - ping
 * - append
 * - listToday
 * - delete
 *
 * Colonne foglio:
 * A = Informazioni cronologiche
 * B = DATA
 * C = ORA
 * D = COGNOME E NOME
 * E = NOTE
 */
function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const callback = String(p.callback || "").replace(/[^\w.$]/g, "");

  try {
    const action = String(p.action || "ping").toLowerCase();
    const sh = getSheet_();

    if (action === "ping") {
      return output_({
        ok: true,
        message: "Collegamento attivo",
        sheet: sh.getName(),
        today: today_()
      }, callback);
    }

    if (action === "listtoday") {
      return output_({
        ok: true,
        date: today_(),
        records: listToday_(sh)
      }, callback);
    }

    if (action === "delete") {
      const row = Number(p.row || 0);

      if (!row || row < 2) {
        throw new Error("Riga non valida per la cancellazione.");
      }

      sh.deleteRow(row);
      SpreadsheetApp.flush();

      return output_({
        ok: true,
        deleted: true,
        records: listToday_(sh)
      }, callback);
    }

    if (action !== "append") {
      throw new Error("Azione non riconosciuta.");
    }

    const nominativo = String(p.cognome_nome || p.nome_completo || "").trim().toUpperCase();
    if (!nominativo) {
      throw new Error("Cognome e nome mancanti.");
    }

    const turno = String(p.turno || "").trim();
    if (!turno) {
      throw new Error("Turno mancante.");
    }

    const data = today_();
    const ora = Utilities.formatDate(new Date(), TZ, "HH:mm:ss");
    const note = "Turno " + turno;
    const id = String(p.id_registrazione || Utilities.getUuid()).trim();

    const duplicateRow = findDuplicateToday_(sh, data, turno, nominativo);
    if (duplicateRow) {
      return output_({
        ok: true,
        duplicate: true,
        row: duplicateRow,
        nome: nominativo,
        records: listToday_(sh)
      }, callback);
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    let row;
    try {
      row = Math.max(sh.getLastRow() + 1, 2);

      sh.getRange(row, 1, 1, 5).setValues([[
        new Date(),
        data,
        ora,
        nominativo,
        note
      ]]);

      sh.getRange(row, 1)
        .setNumberFormat("dd/MM/yyyy HH:mm:ss")
        .setNote(NOTE_PREFIX + id);

      sh.getRange(row, 4).setFontWeight("bold");
      SpreadsheetApp.flush();

    } finally {
      lock.releaseLock();
    }

    return output_({
      ok: true,
      duplicate: false,
      row: row,
      nome: nominativo,
      records: listToday_(sh)
    }, callback);

  } catch (error) {
    return output_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    }, callback);
  }
}

function doPost(e) {
  return doGet(e);
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetById(SHEET_GID);

  if (!sh) {
    throw new Error("Scheda non trovata. Controllare il gid del foglio.");
  }

  return sh;
}

function today_() {
  return Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy");
}

function listToday_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
  const notes = sh.getRange(2, 1, lastRow - 1, 1).getNotes();
  const today = today_();
  const out = [];

  for (let i = 0; i < values.length; i++) {
    const row = i + 2;
    const data = String(values[i][1] || "").trim();

    if (data !== today) continue;

    const ora = String(values[i][2] || "").trim();
    const nominativo = String(values[i][3] || "").trim().toUpperCase();
    const note = String(values[i][4] || "").trim();
    const turno = extractTurno_(note);
    const idNote = String(notes[i][0] || "").replace(NOTE_PREFIX, "").trim();

    out.push({
      row: row,
      id: idNote || String(row),
      data: data,
      ora: ora,
      cognome_nome: nominativo,
      note: note,
      turno: turno
    });
  }

  out.sort(function(a, b) {
    return String(b.ora).localeCompare(String(a.ora));
  });

  return out;
}

function findDuplicateToday_(sh, data, turno, nominativo) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const values = sh.getRange(2, 1, lastRow - 1, 5).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const rowData = String(values[i][1] || "").trim();
    const rowName = String(values[i][3] || "").trim().toUpperCase();
    const rowTurno = extractTurno_(String(values[i][4] || "").trim());

    if (rowData === data && rowTurno === turno && rowName === nominativo) {
      return i + 2;
    }
  }

  return null;
}

function extractTurno_(note) {
  const text = String(note || "");
  if (text.indexOf("13:45") !== -1) return "13:45";
  if (text.indexOf("07:00") !== -1) return "07:00";
  return "";
}

function output_(obj, callback) {
  const json = JSON.stringify(obj);

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Test manuale da eseguire una sola volta dall'editor Apps Script
 * per autorizzare la scrittura sul foglio.
 */
function testAutorizzazioneScrittura() {
  const sh = getSheet_();
  const row = Math.max(sh.getLastRow() + 1, 2);

  sh.getRange(row, 1, 1, 5).setValues([[
    new Date(),
    today_(),
    Utilities.formatDate(new Date(), TZ, "HH:mm:ss"),
    "TEST AUTORIZZAZIONE",
    "Turno 07:00"
  ]]);

  sh.getRange(row, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange(row, 4).setFontWeight("bold");
  SpreadsheetApp.flush();

  Logger.log("Test riuscito: riga scritta nel foglio alla riga " + row);
}
