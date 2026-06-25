const SPREADSHEET_ID = "1CDxFGSFvKIa-yrZahvhYbkxxZg21V1wC5H1DHCa2S90";
const SHEET_GID = 1375747828;

function doGet() {
  return jsonOutput({ok:true, servizio:"Presenze QR 1° anno", sheetGid:SHEET_GID});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetById(SHEET_GID);
    if (!sh) throw new Error("Scheda Google non trovata: gid=" + SHEET_GID);

    ensureHeader(sh);
    const p = e.parameter || {};
    const action = String(p.action || "append").toLowerCase();
    const id = String(p.id_registrazione || "").trim();

    if (action === "delete") {
      if (!id) throw new Error("ID registrazione mancante per la cancellazione");
      const row = findRowById(sh, id);
      if (row) sh.deleteRow(row);
      return jsonOutput({ok:true, deleted:Boolean(row), id_registrazione:id});
    }

    if (id) {
      const existingRow = findRowById(sh, id);
      if (existingRow) return jsonOutput({ok:true, duplicate:true, id_registrazione:id});
    }

    sh.appendRow([
      new Date(), p.data || "", p.ora || "", p.turno || "",
      p.nome || "", p.cognome || "", p.matricola || "",
      p.codice_qr || "", p.corso || "Infermieristica 1° anno",
      p.dispositivo || "", id
    ]);
    return jsonOutput({ok:true, duplicate:false, id_registrazione:id});
  } catch(err) {
    return jsonOutput({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

function ensureHeader(sh) {
  const headers = ["Timestamp ricezione","Data","Ora scansione","Turno","Nome","Cognome","Matricola","Codice QR","Corso","Dispositivo","ID registrazione"];
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#153E5C").setFontColor("#ffffff");
    sh.autoResizeColumns(1,headers.length);
  }
}

function findRowById(sh, id) {
  if (sh.getLastRow() < 2) return null;
  const found = sh.getRange(2,11,sh.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();
  return found ? found.getRow() : null;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
