/*******************************************************
 * BABIA · Control de Embarques — API v2 (Apps Script)
 * Hoja: Babia - Control de Embarques
 *
 * INSTALACIÓN:
 * 1. Abre el Sheet → Extensiones → Apps Script
 * 2. Borra todo y pega este código
 * 3. Ejecuta la función setup() una vez (botón ▶) y acepta permisos
 * 4. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copia la URL /exec y pégala en el index.html (const API)
 *******************************************************/

const SHEET_ID = '1hN4oKXWWrv9qFjSmGf2MAzY7WsKJ2tpXr2i9I4PmwIw';
const TAB = 'Sheet1';

function sheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
}

// Corre esto UNA VEZ a mano: nombra la columna de incidencias y crea la columna KEY
function setup() {
  const sh = sheet_();
  ensure_(sh);
  backfillKeys_();
}

function colMap_(sh) {
  const lastCol = sh.getLastColumn();
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toUpperCase());
  const map = {};
  hdr.forEach((h, i) => { if (h && map[h] === undefined) map[h] = i + 1; });
  return map;
}

function ensure_(sh) {
  let map = colMap_(sh);
  // La columna P de tu hoja tiene incidencias pero sin encabezado — la nombramos
  if (!map['INCIDENCIAS']) {
    const p16 = String(sh.getRange(1, 16).getValue()).trim();
    if (!p16) sh.getRange(1, 16).setValue('INCIDENCIAS');
    else sh.getRange(1, sh.getLastColumn() + 1).setValue('INCIDENCIAS');
    map = colMap_(sh);
  }
  // Columnas nuevas para tarifas y KPIs de desempeño
  ['CITACARGA', 'TEMPLLEGADA', 'TARIFA', 'POD', 'FACTURA', 'DOCS', 'FECHAFACTURA', 'FECHAPAGOPROG', 'FECHAPAGO'].forEach(function (h) {
    if (!map[h]) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      map = colMap_(sh);
    }
  });
  // Llave interna única por registro (los PO se repiten, esto evita editar la fila equivocada)
  if (!map['KEY']) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue('KEY');
    map = colMap_(sh);
  }
  return map;
}

function backfillKeys_() {
  const sh = sheet_();
  const map = ensure_(sh);
  const last = sh.getLastRow();
  if (last < 2) return;
  const n = last - 1;
  const ids = sh.getRange(2, map['ID'], n, 1).getValues();
  const clientes = map['CLIENTE'] ? sh.getRange(2, map['CLIENTE'], n, 1).getValues() : ids;
  const keys = sh.getRange(2, map['KEY'], n, 1).getValues();
  let changed = false;
  for (let i = 0; i < n; i++) {
    const hasData = String(ids[i][0]).trim() !== '' || String(clientes[i][0]).trim() !== '';
    if (hasData && !String(keys[i][0]).trim()) {
      keys[i][0] = Utilities.getUuid();
      changed = true;
    }
  }
  if (changed) sh.getRange(2, map['KEY'], n, 1).setValues(keys);
}

const OUT_FIELDS = {
  ID: 'id', CLIENTE: 'cliente', PRODUCTO: 'producto', TIPO: 'tipo',
  ORIGEN: 'origen', DESTINO: 'destino', TRANSPORTISTA: 'transportista',
  TEMP: 'temp', PESO: 'peso', CAJAS: 'cajas',
  DIACARGA: 'diaCarga', DIADESCARGA: 'diaDescarga', FECHAENTREGA: 'fechaEntrega',
  DIASTR: 'diasTr', ESTATUS: 'estatus', INCIDENCIAS: 'incidencias', ACCIONES: 'acciones',
  CITACARGA: 'citaCarga', TEMPLLEGADA: 'tempLlegada', TARIFA: 'tarifa',
  POD: 'pod', FACTURA: 'factura', DOCS: 'docs',
  FECHAFACTURA: 'fechaFactura', FECHAPAGOPROG: 'fechaPagoProg', FECHAPAGO: 'fechaPago',
  KEY: 'key'
};

function doGet() {
  const sh = sheet_();
  backfillKeys_();
  const map = ensure_(sh);
  const last = sh.getLastRow();
  const rows = [];
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getDisplayValues();
    vals.forEach(v => {
      const get = n => map[n] ? String(v[map[n] - 1]).trim() : '';
      if (!get('ID') && !get('CLIENTE')) return;
      const o = {};
      Object.keys(OUT_FIELDS).forEach(col => { o[OUT_FIELDS[col]] = get(col); });
      rows.push(o);
    });
  }
  return json_({ ok: true, rows: rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const sh = sheet_();
    const map = ensure_(sh);

    if (b.action === 'add') {
      const key = Utilities.getUuid();
      const row = sh.getLastRow() + 1;
      writeRow_(sh, map, row, b);
      sh.getRange(row, map['KEY']).setValue(key);
      return json_({ ok: true, key: key });
    }

    const row = findRow_(sh, map, b.key);
    if (!row) return json_({ ok: false, error: 'No se encontró el registro (key inválida)' });

    if (b.action === 'update') {
      writeRow_(sh, map, row, b);
      return json_({ ok: true, key: b.key });
    }
    if (b.action === 'delete') {
      sh.deleteRow(row);
      return json_({ ok: true });
    }
    if (b.action === 'upload') {
      var id = String(sh.getRange(row, map['ID']).getDisplayValue());
      var folder = folderFor_(id, b.key);
      var blob = Utilities.newBlob(Utilities.base64Decode(b.data), b.mimeType || 'application/octet-stream', b.filename || 'documento');
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var entry = (b.filename || 'documento') + '::' + file.getUrl();
      var col = b.docType === 'POD' ? 'POD' : (b.docType === 'FACTURA' ? 'FACTURA' : 'DOCS');
      if (col === 'DOCS') {
        var cur = String(sh.getRange(row, map['DOCS']).getValue()).trim();
        sh.getRange(row, map['DOCS']).setValue(cur ? cur + ' | ' + entry : entry);
      } else {
        sh.getRange(row, map[col]).setValue(entry);
      }
      return json_({ ok: true, url: file.getUrl(), entry: entry });
    }
    if (b.action === 'carta') {
      if (!b.to) return json_({ ok: false, error: 'Falta el correo destinatario' });
      var g = function (n) { return map[n] ? String(sh.getRange(row, map[n]).getDisplayValue()).trim() : ''; };
      var opts = { htmlBody: cartaHtml_(g, b.extra || ''), name: 'BABIA Logística' };
      if (b.cc) opts.cc = b.cc;
      GmailApp.sendEmail(b.to, 'Carta de Instrucciones · Embarque ' + g('ID') + ' · BABIA', 'Carta de Instrucciones — embarque ' + g('ID'), opts);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'Acción desconocida: ' + b.action });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function writeRow_(sh, map, row, b) {
  const pairs = {
    ID: b.id, CLIENTE: b.cliente, PRODUCTO: b.producto, TIPO: b.tipo,
    ORIGEN: b.origen, DESTINO: b.destino, TRANSPORTISTA: b.transportista,
    TEMP: b.temp, PESO: b.peso, CAJAS: b.cajas,
    DIACARGA: b.diaCarga, DIADESCARGA: b.diaDescarga, FECHAENTREGA: b.fechaEntrega,
    DIASTR: b.diasTr, ESTATUS: b.estatus, INCIDENCIAS: b.incidencias, ACCIONES: b.acciones,
    CITACARGA: b.citaCarga, TEMPLLEGADA: b.tempLlegada, TARIFA: b.tarifa,
    FECHAFACTURA: b.fechaFactura, FECHAPAGOPROG: b.fechaPagoProg, FECHAPAGO: b.fechaPago
  };
  Object.keys(pairs).forEach(k => {
    if (map[k] && pairs[k] !== undefined) sh.getRange(row, map[k]).setValue(pairs[k]);
  });
}

function findRow_(sh, map, key) {
  if (!key) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const keys = sh.getRange(2, map['KEY'], last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === String(key)) return i + 2;
  }
  return 0;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}


function folderFor_(id, key) {
  var rootName = 'BABIA - Documentos de Embarques';
  var roots = DriveApp.getFoldersByName(rootName);
  var root = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);
  var name = (id || 'SIN-ID') + ' \u2014 ' + String(key).slice(0, 8);
  var subs = root.getFoldersByName(name);
  return subs.hasNext() ? subs.next() : root.createFolder(name);
}

function cartaHtml_(g, extra) {
  var row = function (k, v) {
    return v ? '<tr><td style="padding:7px 14px;font-size:10px;letter-spacing:1px;text-transform:uppercase;color:#5C6B60;white-space:nowrap;vertical-align:top">' + k + '</td><td style="padding:7px 14px;font-size:13px;font-weight:bold;color:#17241D">' + v + '</td></tr>' : '';
  };
  var notas = [g('ACCIONES'), extra].filter(function (x) { return x; }).join('<br>');
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:0 auto;border:1px solid #DFE8E0;border-radius:12px;overflow:hidden">' +
    '<div style="background:#123B22;padding:22px 26px;color:#ffffff">' +
      '<div style="font-size:22px;font-weight:bold;letter-spacing:3px">BABIA</div>' +
      '<div style="font-size:11px;letter-spacing:2px;color:#A8C5AF;margin-top:3px">CARTA DE INSTRUCCIONES &middot; EMBARQUE ' + g('ID') + '</div>' +
    '</div>' +
    '<div style="padding:22px 26px;background:#ffffff">' +
      '<p style="font-size:13px;color:#17241D">Estimado transportista <b>' + g('TRANSPORTISTA') + '</b>:</p>' +
      '<p style="font-size:13px;color:#17241D">Por medio de la presente se giran las instrucciones del siguiente embarque. Favor de <b>confirmar de recibido</b> respondiendo a este correo.</p>' +
      '<table style="width:100%;border-collapse:collapse;background:#F7FAF6;border-radius:8px">' +
        row('Referencia / PO', g('ID')) +
        row('Producto', g('PRODUCTO')) +
        row('Temperatura de transporte', g('TEMP') ? g('TEMP') + ' &deg;F continuos' : '') +
        row('Cajas', g('CAJAS')) +
        row('Peso', g('PESO') ? g('PESO') + ' kg' : '') +
        row('Origen (carga)', g('ORIGEN')) +
        row('Cita de carga', g('CITACARGA')) +
        row('D\u00eda de carga', g('DIACARGA')) +
        row('Destino (entrega)', g('DESTINO')) +
        row('D\u00eda de descarga', g('DIADESCARGA')) +
        row('Fecha compromiso de entrega', g('FECHAENTREGA')) +
      '</table>' +
      (notas ? '<p style="font-size:12px;color:#17241D;background:#FFF7E6;border:1px solid #F0C36D;border-radius:8px;padding:10px 14px"><b>Instrucciones adicionales:</b><br>' + notas + '</p>' : '') +
      '<p style="font-size:12px;color:#5C6B60">Es responsabilidad del transportista mantener la cadena de fr\u00edo durante todo el trayecto, reportar de inmediato cualquier incidencia y entregar el POD firmado y sellado al concluir la entrega.</p>' +
      '<p style="font-size:13px;color:#17241D;margin-top:18px">Atentamente,<br><b>BABIA &middot; Log\u00edstica</b></p>' +
    '</div>' +
  '</div>';
}
