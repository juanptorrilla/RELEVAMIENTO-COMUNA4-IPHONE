// Relevamiento Comuna 4 — backend para Google Sheets y Google Drive
const TOKEN = 'C4-ARBOLADO-2026-PLANILLA-4';
const SPREADSHEET_ID = '1l_58opct0coaz2suIsez7UDtB5Wa1RiF';
const SHEET_GID = 1068617970;
function doPost(e) {
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (p.token !== TOKEN) return json_({ok:false,error:'Token de acceso inválido.'});
    if (p.action === 'savePdf') return json_(savePdf_(p));
    if (p.action === 'saveTree') return json_(saveTree_(p.record || {}));
    return json_({ok:false,error:'Acción no reconocida.'});
  } catch (err) { return json_({ok:false,error:String(err && err.message || err)}); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function savePdf_(p) {
  if (!p.data || !p.fileName) throw new Error('Faltan datos del PDF o nombre del archivo.');
  const bytes = Utilities.base64Decode(p.data);
  const blob = Utilities.newBlob(bytes, p.mimeType || 'application/pdf', p.fileName);
  const file = DriveApp.createFile(blob);
  return {ok:true,message:'PDF guardado en Google Drive.',fileId:file.getId(),fileName:file.getName()};
}
function saveTree_(r) {
  const aviso = String(r.aviso || '').trim();
  const submissionId = String(r.submissionId || '').trim();
  if (!aviso) throw new Error('El número de reclamo (AVISO) está vacío. No se modificó la planilla.');
  if (!submissionId) throw new Error('Falta el identificador del envío.');
  const props = PropertiesService.getScriptProperties();
  const sentKey = 'sent_' + Utilities.base64EncodeWebSafe(submissionId).slice(0,180);
  if (props.getProperty(sentKey)) return {ok:true,message:'Este envío ya había sido registrado; no se duplicó.'};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheets().find(s => s.getSheetId() === SHEET_GID);
  if (!sh) throw new Error('No se encontró la hoja configurada. No se modificó la planilla.');
  const last = sh.getLastRow();
  if (last < 2) throw new Error('La planilla no contiene números de reclamo.');
  // Búsqueda EXCLUSIVA por número de reclamo en columna A. No usar calle ni altura.
  const avisos = sh.getRange(2,1,last-1,1).getDisplayValues().flat().map(v => String(v).trim());
  const rows = [];
  avisos.forEach((v,i) => { if (v === aviso) rows.push(i+2); });
  if (!rows.length) throw new Error('No se encontró el reclamo ' + aviso + ' en la columna A. No se escribió en ninguna fila.');
  const values = [[r.fecha || '',r.intervencion || '',r.especie || '',r.dap || '',r.altura || '',r.m2 || '',r.codigoBaldosa || '',r.riesgo || '',r.observacion || '',r.inspector || '']];
  // Primera fila del mismo reclamo sin datos I:R; si ya hay árbol registrado, crear fila hermana.
  let target = rows.find(row => sh.getRange(row,9,1,10).getDisplayValues()[0].every(v => String(v).trim()===''));
  if (!target) {
    target = sh.getLastRow()+1;
    const source = rows[0];
    sh.getRange(source,1,1,24).copyTo(sh.getRange(target,1,1,24), {contentsOnly:false});
    // Se conservan A:H y S:X de la fila de reclamo; no se alteran las fórmulas/controles.
  }
  sh.getRange(target,9,1,10).setValues(values);
  SpreadsheetApp.flush();
  props.setProperty(sentKey, new Date().toISOString());
  return {ok:true,message:'Datos del árbol asociados al reclamo ' + aviso + ' (fila ' + target + ').',row:target};
}
