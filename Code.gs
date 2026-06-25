const SPREADSHEET_ID = "1CDxFGSFvKIa-yrZahvhYbkxxZg21V1wC5H1DHCa2S90";
const SHEET_NAME = "Presenze_1_Anno";

function doGet() {
  return jsonOutput({ok:true, servizio:"Presenze QR 1° anno"});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(["Timestamp ricezione","Data","Ora scansione","Turno","Nome","Cognome","Matricola","Codice QR","Corso","Dispositivo","ID registrazione"]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,11).setFontWeight("bold").setBackground("#153E5C").setFontColor("#ffffff");
      sh.autoResizeColumns(1,11);
    }
    const p=e.parameter||{};
    const id=String(p.id_registrazione||"").trim();
    if(id && sh.getLastRow()>1){
      const found=sh.getRange(2,11,sh.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();
      if(found) return jsonOutput({ok:true,duplicate:true,id_registrazione:id});
    }
    sh.appendRow([new Date(),p.data||"",p.ora||"",p.turno||"",p.nome||"",p.cognome||"",p.matricola||"",p.codice_qr||"",p.corso||"Infermieristica 1° anno",p.dispositivo||"",id]);
    return jsonOutput({ok:true,duplicate:false,id_registrazione:id});
  } catch(err) { return jsonOutput({ok:false,error:String(err)}); } finally { lock.releaseLock(); }
}
function jsonOutput(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
