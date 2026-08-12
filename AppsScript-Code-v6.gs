/**
 * Procesador de contracargos → Google Sheet  (v6)
 * = sincronización (agrega/actualiza, inserción agrupada, marca "¿Volvió a pagar?")  +  monitoreo diario vía Metabase.
 * Pegá TODO este código en Extensiones > Apps Script de tu planilla. Es la versión final.
 */

/* ===================== CONFIGURACIÓN ===================== */
const CONFIG = {
  TAB_NAME: 'Base',   // pestaña de seguimiento
  TOKEN: ''           // opcional, debe coincidir con el "Token" de la página si lo usás
};

// --- Monitoreo diario (completar estos 4 valores) ---
const MONITOR = {
  METABASE_URL: 'https://TU-METABASE.com',   // <-- URL base de tu Metabase (sin /question)
  VENTAS_QUESTION_ID: 0,                       // <-- ID de la pregunta "Ventas Online" (número que aparece en /question/123)
  ALERT_EMAIL: 'adm@fu.do',                    // <-- a quién se avisa
  LOOKBACK_DAYS: 120,                          // solo revisa contracargos de los últimos N días (para no recorrer todo)
  MARCAR_EN_SHEET: true                        // si true, marca "¿Volvió a pagar?" = Sí y anota el pago nuevo
};
// La API key NO va acá: se guarda en Configuración del proyecto > Propiedades del script,
// con el nombre  METABASE_API_KEY  (así no queda expuesta en el código).

const REF_HEADERS = ['referencia desconocimiento', 'referencia', 'reference'];
const STATUS_HEADERS = ['status', 'estado dlocal', 'estado'];

/* ===================== SINCRONIZACIÓN (desde la página) ===================== */
function doPost(e) {
  try {
    var body;
    if (e && e.parameter && e.parameter.data) body = JSON.parse(e.parameter.data);
    else if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
    else body = {};

    if (CONFIG.TOKEN && body.token !== CONFIG.TOKEN) return json({ ok: false, error: 'Token inválido' });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CONFIG.TAB_NAME) || ss.getSheets()[0];
    var lastRow = Math.max(sh.getLastRow(), 1);
    var lastCol = sh.getLastColumn();
    var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = values[0].map(function (h) { return String(h).trim().toLowerCase(); });

    var refCol = firstIndex(headers, REF_HEADERS);
    var statusCol = firstIndex(headers, STATUS_HEADERS);
    var volvioCol = headers.indexOf('¿volvió a pagar?');
    var comentCol = headers.indexOf('comentarios');
    if (refCol < 0) return json({ ok: false, error: 'No encontré la columna de referencia en la fila 1.' });

    var refToRow = {}, lastRefRow = 1;
    for (var r = 1; r < values.length; r++) {
      var ref = String(values[r][refCol]).trim();
      if (ref) { refToRow[ref] = r + 1; lastRefRow = r + 1; }
    }

    var items = body.items || (body.rows ? body.rows.map(function (row) {
      return { ref: row['Referencia desconocimiento'] || '', status: row['Status'] || '', canInsert: true, row: row };
    }) : []);

    var added = 0, updated = 0, unchanged = 0, skipped = 0, volvieron = 0, toInsert = [];
    items.forEach(function (it) {
      var ref = String(it.ref || '').trim();
      var existingRow = ref ? refToRow[ref] : null;
      if (existingRow) {
        if (statusCol >= 0 && it.status) {
          var cur = String(values[existingRow - 1][statusCol]).trim();
          if (cur !== String(it.status).trim()) {
            sh.getRange(existingRow, statusCol + 1).setValue(it.status);
            values[existingRow - 1][statusCol] = it.status; updated++;
          } else unchanged++;
        } else unchanged++;
        // marcar "¿Volvió a pagar?" aunque el estado no haya cambiado
        if (it.volvio && volvioCol >= 0) {
          var curV = String(values[existingRow - 1][volvioCol]).trim().toLowerCase();
          if (curV !== 'sí' && curV !== 'si') {
            sh.getRange(existingRow, volvioCol + 1).setValue('Sí');
            if (comentCol >= 0) {
              var nota = 'Volvió a pagar ' + it.volvio.iso + ' (Payment_ID ' + it.volvio.pid + ')';
              var prev = String(values[existingRow - 1][comentCol] || '');
              sh.getRange(existingRow, comentCol + 1).setValue(prev ? (prev + ' | ' + nota) : nota);
            }
            volvieron++;
          }
        }
      } else {
        if (it.canInsert && it.row) { toInsert.push(buildLine(headers, it.row)); added++; if (it.volvio) volvieron++; }
        else skipped++;
      }
    });
    if (toInsert.length) {
      sh.insertRowsAfter(lastRefRow, toInsert.length);
      sh.getRange(lastRefRow + 1, 1, toInsert.length, headers.length).setValues(toInsert);
    }
    return json({ ok: true, added: added, updated: updated, unchanged: unchanged, skipped: skipped, volvieronAPagar: volvieron, tab: sh.getName() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}
function doGet() { return json({ ok: true, msg: 'Endpoint del procesador de contracargos activo.' }); }

/* ===================== MONITOREO DIARIO ("volvió a pagar") ===================== */
function monitoreoDiario() {
  var key = PropertiesService.getScriptProperties().getProperty('METABASE_API_KEY');
  if (!key) throw new Error('Falta la propiedad METABASE_API_KEY (Configuración del proyecto > Propiedades del script).');
  if (!MONITOR.VENTAS_QUESTION_ID || MONITOR.METABASE_URL.indexOf('TU-METABASE') >= 0)
    throw new Error('Completá MONITOR.METABASE_URL y MONITOR.VENTAS_QUESTION_ID.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.TAB_NAME) || ss.getSheets()[0];
  var vals = sh.getDataRange().getValues();
  var H = vals[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var iRef = firstIndex(H, REF_HEADERS);
  var iLink = H.indexOf('link dash');
  var iNom = H.indexOf('nombre de la cuenta');
  var iFin = H.indexOf('fecha fin');
  var iIni = H.indexOf('fecha inicio');
  var iFecha = H.indexOf('fecha');
  var iVolvio = H.indexOf('¿volvió a pagar?');
  var iComent = H.indexOf('comentarios');

  // pagos por cuenta desde Metabase (Ventas = solo aprobados)
  var ventas = fetchVentas();
  var pays = {};
  ventas.forEach(function (row) {
    var acc = String(pick(row, ['Account_ID', 'account_id', 'ID cuenta', 'id_cuenta']) || '').replace('.0', '').trim();
    var d = parseFecha(pick(row, ['Payment_Date', 'payment_date', 'Fecha de acreditación del pago']));
    if (!acc || !d) return;
    (pays[acc] = pays[acc] || []).push({
      date: d,
      pid: pick(row, ['Payment_ID', 'payment_id']),
      rec: pick(row, ['Payment_Receipt_Number', 'payment_receipt_number'])
    });
  });

  var hoy = new Date();
  var limite = new Date(hoy.getTime() - MONITOR.LOOKBACK_DAYS * 86400000);
  var alertas = [];

  for (var r = 1; r < vals.length; r++) {
    var ref = iRef >= 0 ? String(vals[r][iRef]).trim() : '';
    if (!ref) continue;
    if (iVolvio >= 0 && String(vals[r][iVolvio]).trim().toLowerCase() === 'sí') continue; // ya marcado

    var acc = '';
    if (iLink >= 0) { var m = /accounts\/(\d+)/.exec(String(vals[r][iLink])); if (m) acc = m[1]; }
    if (!acc) continue;

    // fecha ancla del pago desconocido
    var ancla = parseFecha(vals[r][iFin]) || parseFecha(vals[r][iIni]) || parseFecha(vals[r][iFecha]);
    if (!ancla || ancla < limite) continue;

    var lista = pays[acc] || [];
    var nuevo = null;
    lista.forEach(function (p) { if (p.date > ancla && (!nuevo || p.date > nuevo.date)) nuevo = p; });
    if (nuevo) {
      var nombre = iNom >= 0 ? vals[r][iNom] : '';
      alertas.push({ row: r + 1, ref: ref, acc: acc, nombre: nombre, fecha: nuevo.date, pid: nuevo.pid, rec: nuevo.rec });
      if (MONITOR.MARCAR_EN_SHEET) {
        if (iVolvio >= 0) sh.getRange(r + 1, iVolvio + 1).setValue('Sí');
        if (iComent >= 0) {
          var nota = 'Nuevo pago ' + fmtFecha(nuevo.date) + ' (Payment_ID ' + nuevo.pid + ', receipt ' + nuevo.rec + ')';
          var prev = String(vals[r][iComent] || '');
          sh.getRange(r + 1, iComent + 1).setValue(prev ? (prev + ' | ' + nota) : nota);
        }
      }
    }
  }

  if (alertas.length) enviarMail(alertas);
  return { revisados: vals.length - 1, alertas: alertas.length };
}

function fetchVentas() {
  var key = PropertiesService.getScriptProperties().getProperty('METABASE_API_KEY');
  var url = MONITOR.METABASE_URL.replace(/\/+$/, '') + '/api/card/' + MONITOR.VENTAS_QUESTION_ID + '/query/json';
  var res = UrlFetchApp.fetch(url, { method: 'post', headers: { 'x-api-key': key }, muteHttpExceptions: true });
  var code = res.getResponseCode();
  if (code !== 200) throw new Error('Metabase respondió ' + code + ': ' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText()); // array de objetos {columna: valor}
}

function enviarMail(alertas) {
  var html = '<p>Se detectaron <b>' + alertas.length + '</b> cuenta(s) con contracargo que <b>volvieron a pagar</b>:</p><ul>';
  alertas.forEach(function (a) {
    html += '<li><b>' + a.acc + '</b> ' + (a.nombre ? '(' + a.nombre + ') ' : '') +
      '— nuevo pago el ' + fmtFecha(a.fecha) + ' · Payment_ID ' + a.pid + ' · receipt ' + a.rec +
      ' · ref ' + a.ref + '</li>';
  });
  html += '</ul><p>Revisá si el pago nuevo quedó bien asignado.</p>';
  MailApp.sendEmail({ to: MONITOR.ALERT_EMAIL, subject: 'Contracargos: ' + alertas.length + ' cuenta(s) volvieron a pagar', htmlBody: html });
}

/** Ejecutar UNA vez para instalar el trigger diario (ej. crearTriggerDiario(9) = 9 AM). */
function crearTriggerDiario(hora) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'monitoreoDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('monitoreoDiario').timeBased().everyDays(1).atHour(hora || 9).create();
}

/* ===================== HELPERS ===================== */
function pick(obj, names) {
  for (var i = 0; i < names.length; i++) if (obj[names[i]] !== undefined && obj[names[i]] !== null && obj[names[i]] !== '') return obj[names[i]];
  // match case-insensitive
  var keys = Object.keys(obj);
  for (var j = 0; j < names.length; j++) {
    var want = names[j].toLowerCase();
    for (var k = 0; k < keys.length; k++) if (keys[k].toLowerCase() === want) return obj[keys[k]];
  }
  return undefined;
}
function parseFecha(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  var s = String(v).trim();
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);              // ISO
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(s);                // dd/mm/yyyy
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  var d = new Date(s);                                     // "July 07, 2026", etc.
  return isNaN(d.getTime()) ? null : d;
}
function fmtFecha(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function buildLine(headers, obj) {
  var lower = {};
  Object.keys(obj).forEach(function (k) { lower[k.trim().toLowerCase()] = obj[k]; });
  return headers.map(function (h) { return (h in lower) ? lower[h] : ''; });
}
function firstIndex(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) { var idx = headers.indexOf(candidates[i]); if (idx >= 0) return idx; }
  return -1;
}
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
