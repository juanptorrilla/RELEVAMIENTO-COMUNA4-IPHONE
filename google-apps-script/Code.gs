// --------------------------------------------------
// RELEVAMIENTO COMUNA 4 — BACKEND V14
// Mantiene el registro original y permite extracciones de oficio.
// --------------------------------------------------
// Mantiene el registro en RECLAMOS SUACI SLA y agrega extracciones a EXTRACCIONES SLA.
// Busca por AVISO en columna A.
// Guarda el N° exacto de la aplicación en columna C.
// Guarda los datos del árbol en columnas I:R.
// DENEGAR deriva el envío al servicio institucional.

const TOKEN = 'C4-ARBOLADO-2026-PLANILLA-4';

const SPREADSHEET_ID = '1l_58opct0coaz2suIsez7UDtB5Wa1RiF';
const BASE_GENERAL_ID = '1DxbjbjFwOvHHBYI6i7_CjlrepAOxU7cSvr2BCMn_s_w';
const BASE_GENERAL_SHEET = 'BASE_RECLAMOS';
const RECLAMOS_SHEET = 'Reclamos';

// Servicio institucional de notificaciones.
const NOTIFICACIONES_WEBAPP_URL =
  'https://script.google.com/macros/s/AKfycbyxZ6IcdMNaffNH3fXutTi3W6Dm9dopgqOSVMty1NI5AJHrhfjrbLT5Pw6Rrkw5ASxT/exec';

// Planilla adicional: EXTRACCIONES SLA (pestaña gid=0).
const EXTRACCIONES_SPREADSHEET_ID = '1qXcMluBqKpDHCEdtOF73pQA1oYPTR9q3hyvo7usxndM';
const EXTRACCIONES_SHEET_ID = 0;

// --------------------------------------------------
// RECEPCIÓN DE SOLICITUDES
// --------------------------------------------------

function doPost(e) {
  try {
    const p = JSON.parse(
      (e && e.postData && e.postData.contents) || '{}'
    );

    if (p.token !== TOKEN) {
      return json_({ok:false,error:'Token de acceso inválido.'});
    }

    if (p.action === 'savePdf') return json_(savePdf_(p));
    if (p.action === 'listActas') return json_(listActas_());
    if (p.action === 'getActa') return json_(getActa_(p));

    if (p.action === 'saveExtractionDirect') {
      return json_(saveExtractionDirect_(p.record || {}));
    }

    if (p.action === 'saveTree') {
      const record = p.record || {};
      const aviso = String(record.aviso || '').trim();
      const intervencion = String(record.intervencion || '').trim();

      // Las extracciones de oficio se registran solo en EXTRACCIONES SLA.
      if (/extracci[oó]n/i.test(intervencion) && !aviso) {
        return json_(saveExtractionDirect_(record));
      }

      // Los demás casos conservan el flujo original de V13.
      return json_(saveTree_(record));
    }

    if (p.action === 'denyClaim') return json_(denyClaim_(p));

    return json_({ok:false,error:'Acción no reconocida.'});

  } catch (err) {
    return json_({
      ok:false,
      error:String(err && err.message || err)
    });
  }
}

function json_(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------------------------------------------
// GUARDAR PDF
// --------------------------------------------------

function savePdf_(p) {
  if (!p.data || !p.fileName) {
    throw new Error('Faltan datos del PDF o nombre del archivo.');
  }

  const bytes = Utilities.base64Decode(p.data);

  const blob = Utilities.newBlob(
    bytes,
    p.mimeType || 'application/pdf',
    p.fileName
  );

  const pdf = DriveApp.createFile(blob);
  const savedAt = new Date().toISOString();

  const sidecar = {
    app:'Relevamiento Comuna 4',
    pdfFileId:pdf.getId(),
    fileName:pdf.getName(),
    savedAt:savedAt,
    snapshot:p.snapshot || null
  };

  const record = DriveApp.createFile(
    p.fileName + '.acta.json',
    JSON.stringify(sidecar),
    MimeType.PLAIN_TEXT
  );

  return {
    ok:true,
    message:'PDF y datos de recuperación guardados en Google Drive.',
    fileId:pdf.getId(),
    recordId:record.getId(),
    fileName:pdf.getName()
  };
}

// --------------------------------------------------
// LISTAR ACTAS RECUPERABLES
// --------------------------------------------------

function listActas_() {
  const files = DriveApp.getFiles();
  const items = [];

  while (files.hasNext() && items.length < 100) {
    const f = files.next();

    if (!f.getName().endsWith('.acta.json') || f.isTrashed()) {
      continue;
    }

    try {
      const data = JSON.parse(f.getBlob().getDataAsString());

      if (
        data &&
        data.app === 'Relevamiento Comuna 4' &&
        data.snapshot
      ) {
        items.push({
          fileId:f.getId(),
          fileName:data.fileName ||
            f.getName().replace(/\.acta\.json$/,''),
          savedAt:data.savedAt || ''
        });
      }

    } catch (_) {}
  }

  items.sort((a,b) =>
    String(b.savedAt).localeCompare(String(a.savedAt))
  );

  return {ok:true,items:items};
}

// --------------------------------------------------
// RECUPERAR ACTA
// --------------------------------------------------

function getActa_(p) {
  if (!p.fileId) {
    throw new Error('Falta identificar el relevamiento de Drive.');
  }

  const file = DriveApp.getFileById(p.fileId);
  const data = JSON.parse(file.getBlob().getDataAsString());

  if (
    !data ||
    data.app !== 'Relevamiento Comuna 4' ||
    !data.snapshot
  ) {
    throw new Error(
      'El archivo no contiene datos recuperables de la aplicación.'
    );
  }

  return {
    ok:true,
    snapshot:data.snapshot,
    fileName:data.fileName || ''
  };
}

// --------------------------------------------------
// RECUPERAR CALLE DEL RECLAMO (SOLO LECTURA)
// --------------------------------------------------

function getClaimStreetForExtraction_(aviso) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(RECLAMOS_SHEET);

  if (!sh) {
    throw new Error('No se encontró la pestaña Reclamos para recuperar la calle.');
  }

  const last = sh.getLastRow();

  if (last < 2) {
    throw new Error('La pestaña Reclamos no contiene registros.');
  }

  const normalizar = function(v) {
    return String(v == null ? '' : v).trim().toUpperCase();
  };

  const buscado = normalizar(aviso);
  const datos = sh.getRange(2, 1, last - 1, 2).getDisplayValues();

  for (let i = 0; i < datos.length; i++) {
    if (normalizar(datos[i][0]) === buscado) {
      const calle = String(datos[i][1] || '').trim();

      if (calle) return calle;
    }
  }

  throw new Error('No se encontró la CALLE del AVISO ' + aviso + '.');
}

// --------------------------------------------------
// GUARDAR ÁRBOL EN RECLAMOS
// --------------------------------------------------

function saveTree_(r) {

  // AVISO identifica el reclamo. numero es la altura/chapa de la dirección.
  const aviso = String(r.aviso || '').trim();

  // La CALLE se consulta en RECLAMOS SUACI SLA FINAL solo para lectura.
  // ALTURA se toma del campo N° enviado por la aplicación.

  const numeroRaw = r.numero != null ? r.numero : r['N°'];
  const numero = String(numeroRaw == null ? '' : numeroRaw).trim();

  const intervencion = String(r.intervencion || '').trim();
  const fecha = r.fecha || '';
  const submissionId = String(r.submissionId || '').trim();

  if (!aviso) throw new Error('Falta el AVISO del reclamo.');
  if (!numero) throw new Error('Falta el N° de dirección ingresado en la aplicación.');
  if (!submissionId) throw new Error('Falta el identificador del envío.');

  const esExtraccion = /extracci[oó]n/i.test(intervencion);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();
    const encodedId = Utilities.base64EncodeWebSafe(submissionId).slice(0,180);
    const treeKey = 'sent_tree_' + encodedId;
    const extractionKey = 'sent_extraction_' + encodedId;
    const sentKey = 'sent_' + encodedId;

    if (props.getProperty(sentKey)) {
      return {ok:true, message:'Este envío ya había sido registrado; no se duplicó.'};
    }

    let target = null;
    let addedAdditionalTree = false;

    // Primera etapa: mantener el guardado original en RECLAMOS SUACI SLA.
    // La clave propia permite reintentar la segunda planilla sin duplicar este árbol.
    if (!props.getProperty(treeKey)) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sh = ss.getSheetByName(RECLAMOS_SHEET);

      if (!sh) {
        throw new Error('No se encontró la pestaña Reclamos en RECLAMOS SUACI SLA FINAL.');
      }

      const last = sh.getLastRow();
      if (last < 2) throw new Error('No se encontraron registros.');

      const normalizar = function(v) {
        return String(v == null ? '' : v).trim().toUpperCase();
      };

      const avisoBuscado = normalizar(aviso);
      const datos = sh.getRange(2,1,last-1,1).getDisplayValues();
      const rows = [];

      datos.forEach(function(fila,i) {
        if (normalizar(fila[0]) === avisoBuscado) rows.push(i + 2);
      });

      if (rows.length === 0) {
        throw new Error('No se encontró el AVISO ' + aviso +
          ' en la columna A. No se modificó la planilla.');
      }

      const values = [[
        fecha,
        intervencion,
        r.especie || '',
        r.dap || '',
        r.altura || '',
        r.m2 || '',
        r.codigoBaldosa || '',
        r.riesgo || '',
        r.observacion || '',
        r.inspector || ''
      ]];

      // Regla de identificación y corrección:
      // 1) Si ya existe AVISO + chapa, actualizar esa misma fila.
      // 2) Si la chapa cambió y todavía hay una fila original de ese AVISO
      //    sin datos de intervención (I:R vacías), corregir esa fila original.
      // 3) Si todas las filas del AVISO ya tienen intervención, el árbol es
      //    adicional: insertar una nueva fila y conservar las anteriores.
      const norm = function(v) {
        return String(v == null ? '' : v).trim().replace(/\s+/g, ' ').toUpperCase();
      };
      const numeroBuscado = norm(numero);

      target = rows.find(function(row) {
        return norm(sh.getRange(row,3).getDisplayValue()) === numeroBuscado;
      });

      if (!target) {
        const filaOriginalSinIntervencion = rows.find(function(row) {
          const datosArbol = sh.getRange(row,9,1,10).getDisplayValues()[0];
          return datosArbol.every(function(valor) {
            return String(valor == null ? '' : valor).trim() === '';
          });
        });

        if (filaOriginalSinIntervencion) {
          // Primera acta: reemplaza la chapa errónea en la fila del reclamo.
          target = filaOriginalSinIntervencion;
        } else {
          // Árbol adicional: nueva fila, sin borrar ni vaciar los anteriores.
          const lastClaimRow = rows[rows.length - 1];
          sh.insertRowAfter(lastClaimRow);
          target = lastClaimRow + 1;
          sh.getRange(lastClaimRow,1,1,24).copyTo(
            sh.getRange(target,1,1,24), {contentsOnly:false}
          );
          addedAdditionalTree = true;
        }
      }

      sh.getRange(target,3).setNumberFormat('@');
      sh.getRange(target,3).setValue(numero);
      sh.getRange(target,9,1,10).setValues(values);

      SpreadsheetApp.flush();

      // Registrar inmediatamente la primera escritura para que un reintento
      // por fallo en EXTRACCIONES SLA no agregue otra fila del mismo árbol.
      props.setProperty(treeKey, new Date().toISOString());
    }

    // Segunda etapa: una fila por acta/árbol de extracción.
    if (esExtraccion && !props.getProperty(extractionKey)) {
      const extSS = SpreadsheetApp.openById(EXTRACCIONES_SPREADSHEET_ID);
      const extSh = extSS.getSheetById(EXTRACCIONES_SHEET_ID);

      if (!extSh) {
        throw new Error('No se encontró la pestaña gid=0 en EXTRACCIONES SLA.');
      }

      const headers = [
        'N° RECLAMO',
        'CALLE',
        'ALTURA',
        'INTERVENCIÓN',
        'FECHA DE RELEVAMIENTO'
      ];

      if (extSh.getLastRow() === 0) {
        extSh.getRange(1,1,1,headers.length).setValues([headers]);
      }

      // CALLE se consulta en la columna B de RECLAMOS, solo para lectura.
      // ALTURA se toma del campo N° enviado por la aplicación.
      const calle = getClaimStreetForExtraction_(aviso);

      // En EXTRACCIONES SLA, AVISO + chapa catastral identifican el árbol.
      // Si ambos coinciden, actualizar esa fila; si cambia la chapa, agregar
      // una nueva fila sin borrar los registros anteriores.
      const extractionRow = upsertExtractionRow_(extSh, aviso, calle, numero, intervencion, fecha);
      SpreadsheetApp.flush();

      props.setProperty(extractionKey, new Date().toISOString());
    }

    props.setProperty(sentKey, new Date().toISOString());

    return {
      ok:true,
      message: esExtraccion
        ? 'Árbol guardado en RECLAMOS SUACI SLA y extracción registrada en EXTRACCIONES SLA.'
        : 'Datos del árbol guardados correctamente para el AVISO ' + aviso + '.',
      row:target,
      addedAdditionalTree:addedAdditionalTree,
      extractionRegistered:esExtraccion
    };

  } finally {
    lock.releaseLock();
  }
}


// --------------------------------------------------
// INSERTAR O ACTUALIZAR EN EXTRACCIONES SLA
// Con AVISO: AVISO + chapa. Sin AVISO: CALLE + chapa.
// --------------------------------------------------

function upsertExtractionRow_(sh, aviso, calle, numero, intervencion, fecha) {
  const norm = function(v) {
    return String(v == null ? '' : v).trim().replace(/\s+/g, ' ').toUpperCase();
  };

  const avisoKey = norm(aviso);
  const calleKey = norm(calle);
  const chapaKey = norm(numero);
  const last = sh.getLastRow();
  let targetRow = null;

  if (last >= 2) {
    const existing = sh.getRange(2,1,last - 1,5).getDisplayValues();

    // Primero: misma combinación AVISO + chapa.
    for (let i = 0; i < existing.length; i++) {
      const rowAviso = norm(existing[i][0]);
      const rowCalle = norm(existing[i][1]);
      const rowChapa = norm(existing[i][2]);

      const sameTree = avisoKey
        ? rowAviso === avisoKey && rowChapa === chapaKey
        : !rowAviso && rowCalle === calleKey && rowChapa === chapaKey;

      if (sameTree) {
        targetRow = i + 2;
        break;
      }
    }

    // Si el AVISO existe pero la chapa fue corregida, reutilizar la fila
    // original que todavía no tiene intervención cargada (columna D vacía).
    if (!targetRow && avisoKey) {
      for (let i = 0; i < existing.length; i++) {
        if (
          norm(existing[i][0]) === avisoKey &&
          String(existing[i][3] || '').trim() === ''
        ) {
          targetRow = i + 2;
          break;
        }
      }
    }
  }

  // Si ya hay intervención en las filas del AVISO, es otro árbol: agregar fila.
  if (!targetRow) targetRow = Math.max(last + 1, 2);

  sh.getRange(targetRow,1,1,5).setValues([[
    aviso || '',
    calle,
    numero,
    intervencion,
    fecha
  ]]);

  return targetRow;
}

// --------------------------------------------------
// GUARDAR EXTRACCIÓN DE OFICIO (SIN AVISO)
// Registra únicamente en EXTRACCIONES SLA.
// No consulta ni modifica RECLAMOS SUACI SLA FINAL.
// --------------------------------------------------

function saveExtractionDirect_(r) {
  const numeroRaw = r.numero != null ? r.numero : r['N°'];
  const numero = String(numeroRaw == null ? '' : numeroRaw).trim();
  const intervencion = String(r.intervencion || '').trim();
  const fecha = r.fecha || '';
  const submissionId = String(r.submissionId || '').trim();

  // Busca la calle en los campos de dirección enviados por la aplicación.
  const calle = String(
    r.calle || r.direccion || r.street || r.ubicacion ||
    r.address || r.nombreCalle || ''
  ).trim();

  if (!/extracci[oó]n/i.test(intervencion)) {
    throw new Error('El registro sin AVISO debe corresponder a una extracción.');
  }

  if (!calle) {
    throw new Error('Falta la CALLE de la extracción de oficio en los datos enviados por la aplicación.');
  }

  if (!numero) {
    throw new Error('Falta el N° de dirección ingresado en la aplicación.');
  }

  if (!submissionId) {
    throw new Error('Falta el identificador del envío.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();
    const encodedId = Utilities.base64EncodeWebSafe(submissionId).slice(0,180);
    const extractionKey = 'sent_extraction_' + encodedId;

    if (props.getProperty(extractionKey)) {
      return {
        ok:true,
        message:'Esta extracción de oficio ya había sido registrada; no se duplicó.',
        extractionRegistered:true,
        alreadySent:true
      };
    }

    const extSS = SpreadsheetApp.openById(EXTRACCIONES_SPREADSHEET_ID);
    const extSh = extSS.getSheetById(EXTRACCIONES_SHEET_ID);

    if (!extSh) {
      throw new Error('No se encontró la pestaña gid=0 en EXTRACCIONES SLA.');
    }

    const headers = [
      'N° RECLAMO',
      'CALLE',
      'ALTURA',
      'INTERVENCIÓN',
      'FECHA DE RELEVAMIENTO'
    ];

    if (extSh.getLastRow() === 0) {
      extSh.getRange(1,1,1,headers.length).setValues([headers]);
    }

    // Sin AVISO, CALLE + chapa catastral identifican la extracción.
    // Misma calle y chapa: actualizar la fila; chapa distinta: agregar otra.
    const extractionRow = upsertExtractionRow_(extSh, '', calle, numero, intervencion, fecha);

    SpreadsheetApp.flush();

    props.setProperty(extractionKey, new Date().toISOString());

    return {
      ok:true,
      message:'Extracción de oficio registrada correctamente en EXTRACCIONES SLA. No se modificó RECLAMOS SUACI SLA FINAL.',
      row:extractionRow,
      extractionRegistered:true,
      claimSheetModified:false
    };

  } finally {
    lock.releaseLock();
  }
}

// --------------------------------------------------
// DENEGACIÓN DE RECLAMOS
// ENVÍO DELEGADO AL APPS SCRIPT INSTITUCIONAL
// RECIBE Y ENVÍA EL TEXTO EDITADO POR EL INSPECTOR
// --------------------------------------------------

function denyClaim_(p) {

  const aviso = String(p.aviso || '').trim();
  const reason = String(p.reason || '').trim();
  const extra = String(p.extra || '').trim().slice(0, 200);

  // Se conserva el texto completo tal como fue editado en la vista previa.
  // No aplicar trim() al contenido que se enviará al vecino.
  const emailBody = String(p.emailBody || '');

  if (!aviso || !reason) {
    throw new Error(
      'Falta el número de reclamo o el motivo de denegación.'
    );
  }

  if (!emailBody.trim()) {
    throw new Error(
      'Falta el texto revisado del correo de denegación.'
    );
  }

  const payload = {
    token: TOKEN,
    action: 'denyClaim',
    aviso: aviso,
    reason: reason,
    extra: extra,
    emailBody: emailBody
  };

  const response = UrlFetchApp.fetch(
    NOTIFICACIONES_WEBAPP_URL,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  const responseText = response.getContentText();

  let result;

  try {
    result = JSON.parse(responseText);
  } catch (err) {
    throw new Error(
      'El servicio institucional respondió con un formato inesperado.'
    );
  }

  if (
    response.getResponseCode() < 200 ||
    response.getResponseCode() >= 300
  ) {
    throw new Error(
      result.error ||
      'No se pudo contactar al servicio institucional.'
    );
  }

  if (!result.ok) {
    throw new Error(
      result.error ||
      result.message ||
      'No se pudo enviar la notificación institucional.'
    );
  }

  return {
    ok:true,
    message:result.message ||
      'Solicitud procesada por el servicio institucional.',
    reclamo:aviso,
    institutional:true,
    alreadySent:result.alreadySent || false
  };
}

function autorizarConexionNotificaciones() {
  const respuesta = UrlFetchApp.fetch(
    NOTIFICACIONES_WEBAPP_URL,
    {
      method: 'get',
      muteHttpExceptions: true
    }
  );

  Logger.log(
    'Respuesta del servicio: ' + respuesta.getResponseCode()
  );
}