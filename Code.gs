const SPREADSHEET_ID = "1CDxFGSFvKIa-yrZahvhYbkxxZg21V1wC5H1DHCa2S90";
const SHEET_GID = 1375747828;
const NOTE_PREFIX = "PRESENZA_QR_ID:";

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Presenze QR · Infermieristica 1° anno")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no");
}

function salvaPresenza(record) {
  try {
    if (!record || !record.id_registrazione) {
      throw new Error("Dati della presenza mancanti");
    }

    const sh = getSheet_();
    const id = String(record.id_registrazione).trim();

    const existingRow = findRowById_(sh, id);
    if (existingRow) {
      return {
        ok: true,
        duplicate: true,
        row: existingRow,
        id_registrazione: id
      };
    }

    const cognome = String(record.cognome || "").trim();
    const nome = String(record.nome || "").trim();
    const nominativo = [cognome, nome]
      .filter(Boolean)
      .join(" ")
      .toUpperCase();

    if (!nominativo) {
      throw new Error("Cognome e nome dello studente mancanti");
    }

    const timestamp = record.timestamp_iso
      ? new Date(record.timestamp_iso)
      : new Date();

    const data = String(record.data || "").trim();
    const ora = String(record.ora || "").trim();
    const turno = String(record.turno || "").trim();
    const note = String(record.note || ("Turno " + turno)).trim();

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

      const savedName = String(
        sh.getRange(row, 4).getDisplayValue()
      ).trim();

      const savedId = String(
        sh.getRange(row, 1).getNote()
      ).trim();

      if (savedName !== nominativo || savedId !== NOTE_PREFIX + id) {
        throw new Error("Il foglio non ha confermato la registrazione");
      }

    } finally {
      lock.releaseLock();
    }

    return {
      ok: true,
      duplicate: false,
      row: row,
      id_registrazione: id,
      nome: nominativo
    };

  } catch (error) {
    return {
      ok: false,
      error: error && error.message
        ? error.message
        : String(error)
    };
  }
}

function cancellaPresenza(payload) {
  try {
    const id = String(
      payload && payload.id_registrazione
        ? payload.id_registrazione
        : ""
    ).trim();

    if (!id) {
      throw new Error("Identificativo della presenza mancante");
    }

    const sh = getSheet_();
    const row = findRowById_(sh, id);

    if (row) {
      sh.deleteRow(row);
      SpreadsheetApp.flush();
    }

    return {
      ok: true,
      deleted: Boolean(row),
      id_registrazione: id
    };

  } catch (error) {
    return {
      ok: false,
      error: error && error.message
        ? error.message
        : String(error)
    };
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetById(SHEET_GID);

  if (!sh) {
    throw new Error(
      "La scheda del foglio Google non è stata trovata"
    );
  }

  return sh;
}

function findRowById_(sh, id) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const notes = sh
    .getRange(2, 1, lastRow - 1, 1)
    .getNotes();

  const target = NOTE_PREFIX + id;

  for (let i = 0; i < notes.length; i++) {
    if (String(notes[i][0] || "").trim() === target) {
      return i + 2;
    }
  }

  return null;
}
