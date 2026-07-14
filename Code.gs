const SPREADSHEET_ID = "1CDxFGSFvKIa-yrZahvhYbkxxZg21V1wC5H1DHCa2S90";
const SHEET_GID = 1375747828;
const NOTE_PREFIX = "PRESENZA_QR_ID:";

/**
 * Endpoint per GitHub Pages.
 * Riceve i dati via GET/JSONP e scrive nel foglio Google.
 *
 * Colonne:
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
        sheet: sh.getName()
      }, callback);
    }

    if (action === "delete") {
      const id = String(p.id_registrazione || "").trim();
      if (!id) throw new Error("Identificativo mancante.");

      const row = findRowById_(sh, id);

      if (row) {
        sh.deleteRow(row);
        SpreadsheetApp.flush();
      }

      return output_({
        ok: true,
        deleted: Boolean(row),
        id_registrazione: id
      }, callback);
    }

    if (action !== "append") {
      throw new Error("Azione non riconosciuta.");
    }

    const id = String(p.id_registrazione || "").trim();
    if (!id) throw new Error("Identificativo registrazione mancante.");

    const existingRow = findRowById_(sh, id);
    if (existingRow) {
      return output_({
        ok: true,
        duplicate: true,
        row: existingRow,
        id_registrazione: id
      }, callback);
    }

    const nominativo = String(p.cognome_nome || p.nome_completo || "").trim().toUpperCase();
    if (!nominativo) throw new Error("Cognome e nome mancanti.");

    const timestamp = p.timestamp_iso ? new Date(p.timestamp_iso) : new Date();
    const data = String(p.data || "").trim();
    const ora = String(p.ora || "").trim();
    const turno = String(p.turno || "").trim();
    const note = String(p.note || ("Turno " + turno)).trim();

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    let row;
    try {
      row = Math.max(sh.getLastRow() + 1, 2);

      sh.getRange(row, 1, 1, 5).setValues([[
        timestamp,
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

      const checkName = String(sh.getRange(row, 4).getDisplayValue()).trim();
      const checkNote = String(sh.getRange(row, 1).getNote()).trim();

      if (checkName !== nominativo || checkNote !== NOTE_PREFIX + id) {
        throw new Error("Scrittura non confermata dal foglio.");
      }

    } finally {
      lock.releaseLock();
    }

    return output_({
      ok: true,
      duplicate: false,
      row: row,
      id_registrazione: id,
      nome: nominativo
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
    throw new Error("Scheda non trovata. Controllare il gid.");
  }

  return sh;
}

function findRowById_(sh, id) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const notes = sh.getRange(2, 1, lastRow - 1, 1).getNotes();
  const target = NOTE_PREFIX + id;

  for (let i = 0; i < notes.length; i++) {
    if (String(notes[i][0] || "").trim() === target) {
      return i + 2;
    }
  }

  return null;
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
