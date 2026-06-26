const SPREADSHEET_ID = "1CDxFGSFvKIa-yrZahvhYbkxxZg21V1wC5H1DHCa2S90";
const SHEET_GID = 1375747828;
const NOTE_PREFIX = "PRESENZA_QR_ID:";

/*
STRUTTURA DEL FOGLIO
A = Informazioni cronologiche
B = DATA
C = ORA
D = COGNOME E NOME (in stampatello)
E = NOTE

Non viene utilizzata alcuna sesta colonna.
L'ID tecnico viene salvato in modo invisibile come nota della cella A.
*/

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || "");

  try {
    const action = String(p.action || "ping").toLowerCase();
    const sh = getSheet_();

    if (action === "ping") {
      return output_({
        ok: true,
        message: "Collegamento al foglio attivo"
      }, callback);
    }

    if (action === "delete") {
      const id = String(p.id_registrazione || "").trim();
      if (!id) throw new Error("Identificativo mancante per la cancellazione");

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
      throw new Error("Operazione non riconosciuta");
    }

    const id = String(p.id_registrazione || "").trim();
    if (!id) throw new Error("Identificativo della registrazione mancante");

    const existingRow = findRowById_(sh, id);
    if (existingRow) {
      return output_({
        ok: true,
        duplicate: true,
        row: existingRow,
        id_registrazione: id
      }, callback);
    }

    const cognome = String(p.cognome || "").trim();
    const nome = String(p.nome || "").trim();
    const nominativo = [cognome, nome]
      .filter(Boolean)
      .join(" ")
      .toUpperCase();

    if (!nominativo) {
      throw new Error("Cognome e nome dello studente mancanti");
    }

    const timestamp = p.timestamp_iso
      ? new Date(p.timestamp_iso)
      : new Date();

    const data = String(p.data || "").trim();
    const ora = String(p.ora || "").trim();
    const turno = String(p.turno || "").trim();
    const note = String(p.note || ("Turno " + turno)).trim();

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    let row;
    try {
      row = Math.max(sh.getLastRow() + 1, 2);

      // Scrive ESCLUSIVAMENTE nelle colonne A-E.
      sh.getRange(row, 1, 1, 5).setValues([[
        timestamp,   // A - Informazioni cronologiche
        data,        // B - DATA
        ora,         // C - ORA
        nominativo,  // D - COGNOME E NOME
        note         // E - NOTE
      ]]);

      sh.getRange(row, 1)
        .setNumberFormat("dd/MM/yyyy HH:mm:ss")
        .setNote(NOTE_PREFIX + id);

      sh.getRange(row, 4).setFontWeight("bold");
      SpreadsheetApp.flush();

      // Conferma reale della scrittura.
      const savedName = String(sh.getRange(row, 4).getDisplayValue()).trim();
      const savedNote = String(sh.getRange(row, 1).getNote()).trim();

      if (savedName !== nominativo || savedNote !== NOTE_PREFIX + id) {
        throw new Error("Il foglio Google non ha confermato la registrazione");
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

  } catch (err) {
    return output_({
      ok: false,
      error: err && err.message ? err.message : String(err)
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
    throw new Error("La scheda con gid 1375747828 non è stata trovata");
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
    const safeCallback = callback.replace(/[^\w.$]/g, "");

    return ContentService
      .createTextOutput(safeCallback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
