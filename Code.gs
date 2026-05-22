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
  // Calculamos todo en JS y escribimos B2/B3/B4 en la hoja
  // (así no depende de fórmulas locale-dependent de Sheets)
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const stockSh = ss.getSheetByName(SH_STOCK);
  if (!stockSh) throw new Error('Falta hoja STOCK');

  const inicial = Number(stockSh.getRange('B1').getValue()) || 0;

  // Sumar PRODUCCION!B (bolsas_producidas)
  const prodSh = ss.getSheetByName(SH_PRODUCCION);
  let producido = 0;
  if (prodSh && prodSh.getLastRow() >= 2) {
    const v = prodSh.getRange(2, 2, prodSh.getLastRow() - 1, 1).getValues();
    v.forEach(([n]) => { producido += Number(n) || 0; });
  }

  // Sumar cantidad_bolsas de PEDIDOS con estado=entregado
  // Columnas PEDIDOS: A=id B=fecha_carga C=cliente D=cantidad_bolsas E=obs_pedido F=estado
  const pedSh = ss.getSheetByName(SH_PEDIDOS);
  let vendido = 0;
  if (pedSh && pedSh.getLastRow() >= 2) {
    const v = pedSh.getRange(2, 1, pedSh.getLastRow() - 1, 6).getValues();
    v.forEach(row => {
      if (String(row[5]).toLowerCase() === 'entregado') {
        vendido += Number(row[3]) || 0;
      }
    });
  }

  const actual = inicial + producido - vendido;

  // Reflejar en la hoja para que se pueda ver desde Drive
  stockSh.getRange('B2:B4').setValues([[producido], [vendido], [actual]]);

  return {
    stock_inicial:   inicial,
    total_producido: producido,
    total_vendido:   vendido,
    stock_actual:    actual
  };
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
      ['total_producido', 0],
      ['total_vendido',   0],
      ['stock_actual',    0]
    ]);
  }
}
