const SPREADSHEET_ID = "1CDxFGSFvKIa-yrZahvhYbkxxZg21V1wC5H1DHCa2S90";
const SHEET_GID = 1375747828;

/*
A = Informazioni cronologiche
B = DATA
C = ORA
D = COGNOME E NOME (in stampatello)
E = NOTE
F = ID tecnico nascosto
*/

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || "");

  try {
    const action = String(p.action || "ping").toLowerCase();
    const sh = getSheet_();

    if (action === "ping") {
      return output_({ok:true, message:"Collegamento attivo"}, callback);
    }

    if (action === "delete") {
      const id = String(p.id_registrazione || "").trim();
      if (!id) throw new Error("ID mancante per la cancellazione");

      const row = findRowById_(sh, id);
      if (row) sh.deleteRow(row);
      SpreadsheetApp.flush();

      return output_({
        ok:true,
        deleted:Boolean(row),
        id_registrazione:id
      }, callback);
    }

    if (action !== "append") {
      throw new Error("Azione non riconosciuta");
    }

    const id = String(p.id_registrazione || "").trim();
    if (!id) throw new Error("ID registrazione mancante");

    const existing = findRowById_(sh, id);
    if (existing) {
      return output_({
        ok:true,
        duplicate:true,
        row:existing,
        id_registrazione:id
      }, callback);
    }

    const cognome = String(p.cognome || "").trim();
    const nome = String(p.nome || "").trim();
    const nominativo = [cognome, nome].filter(Boolean).join(" ").toUpperCase();
    if (!nominativo) throw new Error("Cognome e nome mancanti");

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

      sh.getRange(row, 1, 1, 6).setValues([[
        timestamp,
        data,
        ora,
        nominativo,
        note,
        id
      ]]);

      sh.getRange(row, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
      sh.getRange(row, 4).setFontWeight("bold");
      sh.hideColumns(6);
      SpreadsheetApp.flush();

      const confirmed = String(sh.getRange(row, 6).getValue()).trim();
      if (confirmed !== id) {
        throw new Error("Il foglio non ha confermato la registrazione");
      }
    } finally {
      lock.releaseLock();
    }

    return output_({
      ok:true,
      duplicate:false,
      row:row,
      id_registrazione:id,
      nome:nominativo
    }, callback);

  } catch (err) {
    return output_({
      ok:false,
      error:err && err.message ? err.message : String(err)
    }, callback);
  }
}

function doPost(e) {
  return doGet(e);
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetById(SHEET_GID);
  if (!sh) throw new Error("Scheda Google non trovata");
  return sh;
}

function findRowById_(sh, id) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const values = sh.getRange(2, 6, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === id) return i + 2;
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
