// Relevamiento Comuna 4 — backend V7 ampliado
// Mantiene la escritura de datos del árbol en I:R y limita la actualización de altura a la columna C.
const TOKEN = 'C4-ARBOLADO-2026-PLANILLA-4';
const SPREADSHEET_ID = '1l_58opct0coaz2suIsez7UDtB5Wa1RiF';
const SHEET_GID = 1068617970;
const EXTRACCIONES_SPREADSHEET_ID = '1qXcMluBqKpDHCEdtOF73pQA1oYPTR9q3hyvo7usxndM';
const EXTRACCIONES_SHEET_GID = 0;
const BASE_GENERAL_ID = '1DxbjbjFwOvhhBYI6i7_CjlrepAOxU7cSvr2BCMn_s_w';
const BASE_GENERAL_SHEET = 'BASE_RECLAMOS';

function doPost(e) {
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (p.token !== TOKEN) return json_({ok:false,error:'Token de acceso inválido.'});
    if (p.action === 'savePdf') return json_(savePdf_(p));
    if (p.action === 'listActas') return json_(listActas_());
    if (p.action === 'getActa') return json_(getActa_(p));
    if (p.action === 'saveTree') return json_(saveTree_(p.record || {}));
    if (p.action === 'denyClaim') return json_(denyClaim_(p));
    return json_({ok:false,error:'Acción no reconocida.'});
  } catch (err) { return json_({ok:false,error:String(err && err.message || err)}); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function savePdf_(p) {
  if (!p.data || !p.fileName) throw new Error('Faltan datos del PDF o nombre del archivo.');
  const bytes = Utilities.base64Decode(p.data);
  const blob = Utilities.newBlob(bytes, p.mimeType || 'application/pdf', p.fileName);
  const pdf = DriveApp.createFile(blob);
  const savedAt = new Date().toISOString();
  const sidecar = {app:'Relevamiento Comuna 4',pdfFileId:pdf.getId(),fileName:pdf.getName(),savedAt:savedAt,snapshot:p.snapshot || null};
  const record = DriveApp.createFile(p.fileName + '.acta.json', JSON.stringify(sidecar), MimeType.PLAIN_TEXT);
  return {ok:true,message:'PDF y datos de recuperación guardados en Google Drive.',fileId:pdf.getId(),recordId:record.getId(),fileName:pdf.getName()};
}
function listActas_() {
  const files = DriveApp.getFiles();
  const items = [];
  while (files.hasNext() && items.length < 100) {
    const f = files.next();
    if (!f.getName().endsWith('.acta.json') || f.isTrashed()) continue;
    try {
      const data = JSON.parse(f.getBlob().getDataAsString());
      if (data && data.app === 'Relevamiento Comuna 4' && data.snapshot) {
        items.push({fileId:f.getId(),fileName:data.fileName || f.getName().replace(/\.acta\.json$/,''),savedAt:data.savedAt || ''});
      }
    } catch (_) {}
  }
  items.sort((a,b)=>String(b.savedAt).localeCompare(String(a.savedAt)));
  return {ok:true,items:items};
}
function getActa_(p) {
  if (!p.fileId) throw new Error('Falta identificar el relevamiento de Drive.');
  const file = DriveApp.getFileById(p.fileId);
  const data = JSON.parse(file.getBlob().getDataAsString());
  if (!data || data.app !== 'Relevamiento Comuna 4' || !data.snapshot) throw new Error('El archivo no contiene datos recuperables de la aplicación.');
  return {ok:true,snapshot:data.snapshot,fileName:data.fileName || ''};
}

function saveTree_(r) {
  const aviso = String(r.aviso || '').trim();
  const submissionId = String(r.submissionId || '').trim();
  const isExtraction = /extracci[oó]n/i.test(String(r.intervencion || ''));

  if (!submissionId) throw new Error('Falta el identificador del envío.');
  if (!aviso) {
    if (!isExtraction) throw new Error('Para enviar un relevamiento sin número de reclamo, la intervención debe ser una extracción.');
    // Sin AVISO no se consulta ni se modifica RECLAMOS SUACI SLA FINAL.
    // La extracción se registra exclusivamente en EXTRACCIONES SLA.
    return saveExtraction_(r, '', submissionId);
  }

  const props = PropertiesService.getScriptProperties();
  const sentKey = 'sent_' + Utilities.base64EncodeWebSafe(submissionId).slice(0,180);
  const treeAlreadySent = !!props.getProperty(sentKey);
  let target = null;

  // El envío principal conserva su comportamiento actual: búsqueda exclusiva por AVISO
  // y escritura de los datos del árbol en la fila correspondiente.
  if (!treeAlreadySent) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheets().find(s => s.getSheetId() === SHEET_GID);
    if (!sh) throw new Error('No se encontró la hoja configurada. No se modificó la planilla.');
    const last = sh.getLastRow();
    if (last < 2) throw new Error('La planilla no contiene números de reclamo.');
    const avisos = sh.getRange(2,1,last-1,1).getDisplayValues().flat().map(v => String(v).trim());
    const rows = [];
    avisos.forEach((v,i) => { if (v === aviso) rows.push(i+2); });
    if (!rows.length) throw new Error('No se encontró el reclamo ' + aviso + ' en la columna A. No se escribió en ninguna fila.');

    const values = [[r.fecha || '',r.intervencion || '',r.especie || '',r.dap || '',r.altura || '',r.m2 || '',r.codigoBaldosa || '',r.riesgo || '',r.observacion || '',r.inspector || '']];
    target = rows.find(row => sh.getRange(row,9,1,10).getDisplayValues()[0].every(v => String(v).trim()===''));
    if (!target) {
      target = sh.getLastRow()+1;
      const source = rows[0];
      sh.getRange(source,1,1,24).copyTo(sh.getRange(target,1,1,24), {contentsOnly:false});
    }

    // Mantiene la actualización existente de la chapa catastral en la columna C.
    const numeroRaw = r.numero != null ? r.numero : r['N°'];
    const numero = numeroRaw == null ? '' : String(numeroRaw);
    if (numero.trim() !== '') {
      const cell = sh.getRange(target,3);
      const current = String(cell.getDisplayValue() == null ? '' : cell.getDisplayValue());
      if (current !== numero) {
        cell.setNumberFormat('@');
        cell.setValue(numero);
      }
    }

    sh.getRange(target,9,1,10).setValues(values);
    SpreadsheetApp.flush();
    props.setProperty(sentKey, new Date().toISOString());
  }

  // Si es extracción, también se registra en EXTRACCIONES SLA.
  // Un error en la segunda planilla no duplica el registro principal al reintentar.
  let extractionResult = null;
  if (isExtraction) {
    try {
      extractionResult = saveExtraction_(r, aviso, submissionId);
    } catch (extractionError) {
      return {
        ok:true,
        extractionOk:false,
        extractionError:String(extractionError && extractionError.message || extractionError),
        message:'Los datos quedaron registrados en RECLAMOS SUACI SLA FINAL, pero no se pudo completar el envío adicional a EXTRACCIONES SLA. Se puede reintentar sin duplicar el registro principal.',
        row:target
      };
    }
  }

  const message = treeAlreadySent
    ? 'Este árbol ya estaba registrado en RECLAMOS SUACI SLA FINAL; no se duplicó.'
    : 'Datos asociados al reclamo ' + aviso + (target ? ' (fila ' + target + ').' : '.');
  return {
    ok:true,
    extractionOk:extractionResult ? extractionResult.ok : null,
    extractionMessage:extractionResult ? extractionResult.message : '',
    message:message + (extractionResult ? ' ' + extractionResult.message : '')
  };
}

function saveExtraction_(r, aviso, submissionId) {
  const intervencion = String(r.intervencion || '').trim();
  if (!/extracci[oó]n/i.test(intervencion)) return null;

  const props = PropertiesService.getScriptProperties();
  const extractionKey = 'sent_extraction_' + Utilities.base64EncodeWebSafe(submissionId).slice(0,180);
  if (props.getProperty(extractionKey)) {
    return {ok:true,message:'Esta acta ya estaba registrada en EXTRACCIONES SLA; no se duplicó.'};
  }

  const numeroRaw = r.numero != null ? r.numero : r['N°'];
  const altura = numeroRaw == null ? '' : String(numeroRaw).trim();
  if (!altura) throw new Error('El campo N° de la aplicación está vacío; no se agregó la fila a EXTRACCIONES SLA.');

  // Con AVISO, la calle se obtiene de la columna B del reclamo encontrado.
  // Sin AVISO, se toma la calle ingresada en el formulario y no se consulta RECLAMOS SUACI.
  const calle = aviso
    ? getClaimStreetForExtraction_(aviso)
    : String(r.calle || '').trim();
  if (!calle) throw new Error(aviso
    ? 'No se encontró la CALLE del reclamo ' + aviso + ' en la columna B de RECLAMOS SUACI SLA.'
    : 'Completá la CALLE para registrar la extracción sin número de reclamo.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (props.getProperty(extractionKey)) {
      return {ok:true,message:'Esta acta ya estaba registrada en EXTRACCIONES SLA; no se duplicó.'};
    }
    const ss = SpreadsheetApp.openById(EXTRACCIONES_SPREADSHEET_ID);
    const sh = ss.getSheets().find(s => s.getSheetId() === EXTRACCIONES_SHEET_GID);
    if (!sh) throw new Error('No se encontró la pestaña de EXTRACCIONES SLA con gid 0.');
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,5).setValues([['N° RECLAMO','DIRECCION','ALTURA','INTERVENCION','FECHA DE RELEVAMIENTO']]);
    }
    sh.appendRow([aviso,calle,altura,intervencion,r.fecha || '']);
    SpreadsheetApp.flush();
    props.setProperty(extractionKey, new Date().toISOString());
    return {ok:true,message:aviso
      ? 'Acta agregada también a EXTRACCIONES SLA.'
      : 'Extracción registrada únicamente en EXTRACCIONES SLA, sin número de reclamo.'};
  } finally {
    lock.releaseLock();
  }
}

function getClaimStreetForExtraction_(aviso) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheets().find(s => s.getSheetId() === SHEET_GID);
  if (!sh) throw new Error('No se encontró la hoja de RECLAMOS SUACI SLA FINAL.');
  const last = sh.getLastRow();
  if (last < 2) return '';
  const rows = sh.getRange(2,1,last-1,2).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(aviso).trim()) {
      return String(rows[i][1] || '').trim();
    }
  }
  return '';
}

function denyClaim_(p) {
  const aviso = String(p.aviso || '').trim();
  const reason = String(p.reason || '').trim();
  const extra = String(p.extra || '').trim().slice(0,200);
  if (!aviso || !reason) throw new Error('Falta el número de reclamo o el motivo de denegación.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(BASE_GENERAL_ID);
    const sh = ss.getSheetByName(BASE_GENERAL_SHEET);
    if (!sh) throw new Error('No se encontró la hoja BASE_RECLAMOS de la base general.');
    const last = sh.getLastRow();
    const avisos = sh.getRange(1,1,last,1).getDisplayValues().flat().map(v=>String(v).trim());
    const index = avisos.findIndex(v=>v===aviso);
    if (index < 0) throw new Error('No se encontró el reclamo ' + aviso + ' en la BASE GENERAL.');
    const row = index + 1;
    const email = String(sh.getRange(row,7).getDisplayValue() || '').trim();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('El reclamo no tiene una dirección de correo válida en la columna G.');
    const prior = String(sh.getRange(row,13).getDisplayValue() || '').toUpperCase();
    if (prior.indexOf('DENEGACIÓN ENVIADA') >= 0 || prior.indexOf('DENEGACION ENVIADA') >= 0) return {ok:true,message:'La notificación de denegación ya figura enviada. No se volvió a enviar.'};
    const subject = 'Respuesta a su reclamo de arbolado — ' + aviso;
    const body = 'Estimado/a vecino/a:\n\nEn relación con el reclamo ' + aviso + ', le informamos que no corresponde realizar la intervención solicitada por el siguiente motivo:\n\n' + reason + (extra ? '\n\n' + extra : '') + '\n\nSaludos cordiales,\nÁrea de Arbolado\nComuna 4';
    MailApp.sendEmail({to:email,subject:subject,body:body,name:'Arbolado Comuna 4'});
    sh.getRange(row,11).setValue('DENEGADO');
    sh.getRange(row,13).setValue('DENEGACIÓN ENVIADA');
    sh.getRange(row,14).setValue(new Date());
    SpreadsheetApp.flush();
    return {ok:true,message:'Correo de denegación enviado y notificación registrada en la BASE GENERAL.'};
  } finally { lock.releaseLock(); }
}
