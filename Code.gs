/**
 * Backend del panel de Asfalto en Frío.
 * Lee/escribe en la planilla "Asfalto" y guarda fotos de remitos
 * en la carpeta de Drive "Remitos Asfalto".
 *
 * Desplegar como Web App:
 *   Ejecutar como: yo (Marcos)
 *   Quien tiene acceso: cualquier persona
 * Copiar la URL /exec y pegarla en index.html (CONFIG.API_URL).
 */

// === CONFIG (editar después de crear Sheet y carpeta) ===
const SHEET_ID  = 'PEGAR_AQUI_EL_ID_DE_LA_PLANILLA';
const FOLDER_ID = 'PEGAR_AQUI_EL_ID_DE_LA_CARPETA_REMITOS';

// === NOMBRES DE HOJAS ===
const SH_PEDIDOS    = 'PEDIDOS';
const SH_PRODUCCION = 'PRODUCCION';
const SH_STOCK      = 'STOCK';

// === HTTP ===
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'all';
    let data;
    if (action === 'pendientes')      data = { pendientes: getPendientes() };
    else if (action === 'entregados') data = { entregados: getEntregados() };
    else if (action === 'stock')      data = { stock: getStock() };
    else                              data = getAll();
    return jsonResponse({ ok: true, data });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    let result;
    if (action === 'entregar')      result = entregarPedido(body);
    else if (action === 'producir') result = registrarProduccion(body);
    else throw new Error('Acción desconocida: ' + action);
    return jsonResponse({ ok: true, data: result });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// === LECTURA ===
function getAll() {
  return {
    pendientes: getPendientes(),
    entregados: getEntregados(),
    stock:      getStock()
  };
}

function getPendientes() {
  const rows = readSheet(SH_PEDIDOS);
  return rows
    .filter(r => String(r.estado || '').toLowerCase() === 'pendiente')
    .map(r => ({
      id: r.id,
      fecha_carga: fmt(r.fecha_carga),
      cliente: r.cliente,
      cantidad_bolsas: Number(r.cantidad_bolsas) || 0,
      observacion_pedido: r.observacion_pedido || ''
    }));
}

function getEntregados() {
  const rows = readSheet(SH_PEDIDOS);
  const hoyIso = new Date().toISOString().slice(0, 10);
  return rows
    .filter(r => String(r.estado || '').toLowerCase() === 'entregado')
    .map(r => ({
      id: r.id,
      cliente: r.cliente,
      cantidad_bolsas: Number(r.cantidad_bolsas) || 0,
      observacion_pedido: r.observacion_pedido || '',
      fecha_entrega: fmt(r.fecha_entrega),
      observacion_entrega: r.observacion_entrega || '',
      link_remito: r.link_remito || '',
      _fecha_iso: r.fecha_entrega ? new Date(r.fecha_entrega).toISOString().slice(0, 10) : ''
    }))
    // mostrar entregados de hoy primero, después los de días anteriores; máximo 50
    .sort((a, b) => (b._fecha_iso || '').localeCompare(a._fecha_iso || ''))
    .slice(0, 50)
    .map(r => { delete r._fecha_iso; return r; });
}

function getStock() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_STOCK);
  if (!sh) throw new Error('Falta hoja STOCK');
  // Layout esperado de STOCK (columna A label, columna B valor):
  // A1 "stock_inicial"   B1 <numero>
  // A2 "total_producido" B2 =SUM(PRODUCCION!B:B)
  // A3 "total_vendido"   B3 =SUMIF(PEDIDOS!F:F, "entregado", PEDIDOS!D:D)
  // A4 "stock_actual"    B4 =B1+B2-B3
  const values = sh.getRange('A1:B4').getValues();
  const stock = {};
  values.forEach(([k, v]) => { if (k) stock[k] = Number(v) || 0; });
  return stock;
}

// === ESCRITURA ===
function entregarPedido(body) {
  const id          = String(body.id || '').trim();
  const observacion = String(body.observacion || '').trim();
  const fotoB64     = body.foto_base64 || '';
  const mime        = body.foto_mime || 'image/jpeg';
  if (!id) throw new Error('Falta id');

  const ss  = SpreadsheetApp.openById(SHEET_ID);
  const sh  = ss.getSheetByName(SH_PEDIDOS);
  if (!sh) throw new Error('Falta hoja PEDIDOS');
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const colId        = headers.indexOf('id');
  const colEstado    = headers.indexOf('estado');
  const colFEntrega  = headers.indexOf('fecha_entrega');
  const colObsE      = headers.indexOf('observacion_entrega');
  const colLink      = headers.indexOf('link_remito');
  const colCliente   = headers.indexOf('cliente');
  if (colId < 0 || colEstado < 0) throw new Error('PEDIDOS sin columnas requeridas');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colId]) === id) {
      let linkRemito = '';
      if (fotoB64) {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const cliente = colCliente >= 0 ? String(data[i][colCliente] || '').replace(/[^a-z0-9]+/gi, '_') : '';
        const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
        const name = `remito_${id}_${cliente}_${stamp}.jpg`;
        const blob = Utilities.newBlob(Utilities.base64Decode(fotoB64), mime, name);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        linkRemito = file.getUrl();
      }
      const row = i + 1;
      sh.getRange(row, colEstado    + 1).setValue('entregado');
      if (colFEntrega >= 0) sh.getRange(row, colFEntrega + 1).setValue(new Date());
      if (colObsE     >= 0) sh.getRange(row, colObsE     + 1).setValue(observacion);
      if (colLink     >= 0) sh.getRange(row, colLink     + 1).setValue(linkRemito);
      return { id, link_remito: linkRemito };
    }
  }
  throw new Error('Pedido no encontrado: ' + id);
}

function registrarProduccion(body) {
  const cantidad = Number(body.cantidad);
  const operario = String(body.operario || '').trim();
  if (!cantidad || cantidad <= 0) throw new Error('Cantidad inválida');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_PRODUCCION);
  if (!sh) throw new Error('Falta hoja PRODUCCION');
  sh.appendRow([new Date(), cantidad, operario]);
  return { fecha: new Date().toISOString(), cantidad, operario };
}

// === UTILIDADES ===
function readSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Falta hoja ' + name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1)
    .filter(row => row.some(c => c !== '' && c !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function fmt(d) {
  if (!d) return '';
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return String(d);
}

// === HELPER: ejecutar una vez para crear las hojas con headers correctos ===
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ensure = (name, headers) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
  };
  ensure(SH_PEDIDOS, [
    'id', 'fecha_carga', 'cliente', 'cantidad_bolsas',
    'observacion_pedido', 'estado', 'fecha_entrega',
    'observacion_entrega', 'link_remito'
  ]);
  ensure(SH_PRODUCCION, ['fecha', 'bolsas_producidas', 'operario']);
  let stock = ss.getSheetByName(SH_STOCK);
  if (!stock) stock = ss.insertSheet(SH_STOCK);
  if (stock.getLastRow() === 0) {
    stock.getRange('A1:B4').setValues([
      ['stock_inicial',   0],
      ['total_producido', '=SUM(PRODUCCION!B2:B)'],
      ['total_vendido',   '=SUMIF(PEDIDOS!F2:F, "entregado", PEDIDOS!D2:D)'],
      ['stock_actual',    '=B1+B2-B3']
    ]);
  }
}
