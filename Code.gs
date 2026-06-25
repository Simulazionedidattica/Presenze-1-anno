const SPREADSHEET_ID = "1CDxFGSFvKIa-yrZahvhYbkxxZg21V1wC5H1DHCa2S90";
const SHEET_GID = 1375747828;
const META_KEY = "ID_REGISTRAZIONE";

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetById(SHEET_GID);
    if (!sh) throw new Error("Scheda Google non trovata: gid=" + SHEET_GID);

    const p = (e && e.parameter) || {};
    const action = String(p.action || "append").toLowerCase();
    const id = String(p.id_registrazione || "").trim();

    if (action === "delete") {
      if (!id) throw new Error("ID registrazione mancante per la cancellazione");
      const row = findRowByMetadata(sh, id);
      if (row) sh.deleteRow(row);
      return output({ok:true, deleted:Boolean(row), id_registrazione:id}, p.callback);
    }

    if (!id) throw new Error("ID registrazione mancante");

    const existingRow = findRowByMetadata(sh, id);
    if (existingRow) {
      return output({ok:true, duplicate:true, id_registrazione:id}, p.callback);
    }

    const cognome = String(p.cognome || "").trim();
    const nome = String(p.nome || "").trim();
    const nominativo = [cognome, nome].filter(Boolean).join(" ").toUpperCase();

    if (!nominativo) throw new Error("Cognome e nome mancanti");

    const timestamp = p.timestamp_iso ? new Date(p.timestamp_iso) : new Date();
    const data = String(p.data || "").trim();
    const ora = String(p.ora || "").trim();
    const note = String(p.note || ("Turno " + (p.turno || ""))).trim();

    const nextRow = Math.max(sh.getLastRow() + 1, 2);

    // A: Informazioni cronologiche
    // B: DATA
    // C: ORA
    // D: COGNOME E NOME (in stampatello)
    // E: NOTE
    sh.getRange(nextRow, 1, 1, 5).setValues([[
      timestamp,
      data,
      ora,
      nominativo,
      note
    ]]);

    sh.getRange(nextRow, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sh.getRange(nextRow, 4).setFontWeight("bold");
    sh.getRange(nextRow, 1, 1, 5).addDeveloperMetadata(META_KEY, id);

    return output({
      ok:true,
      duplicate:false,
      id_registrazione:id,
      row:nextRow,
      nome:nominativo
    }, p.callback);

  } catch (err) {
    const callback = e && e.parameter ? e.parameter.callback : "";
    return output({ok:false, error:String(err)}, callback);
  } finally {
    lock.releaseLock();
  }
}

function findRowByMetadata(sh, id) {
  const results = sh.createDeveloperMetadataFinder()
    .withKey(META_KEY)
    .withValue(id)
    .find();

  if (!results || !results.length) return null;

  const location = results[0].getLocation();
  const range = location.getRange();
  return range ? range.getRow() : null;
}

function output(obj, callback) {
  const json = JSON.stringify(obj);

  if (callback) {
    const safeCallback = String(callback).replace(/[^\w.$]/g, "");
    return ContentService
      .createTextOutput(safeCallback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
