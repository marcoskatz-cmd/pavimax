/**
 * Backend del panel PAVIMAX.
 * Lee/escribe en la planilla "PAVIMAX" y guarda fotos de remitos en Drive.
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
const SH_CLIENTES   = 'CLIENTES';
const SH_PRODUCTO   = 'PRODUCTO';
const SH_GANANCIAS  = 'GANANCIAS';

// === HTTP ===
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'all';
    let data;
    if (action === 'pendientes')      data = { pendientes: getPendientes() };
    else if (action === 'entregados') data = { entregados: getEntregados() };
    else if (action === 'stock')      data = { stock: getStock() };
    else if (action === 'ganancias')  data = { ganancias: getGanancias() };
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
    stock:      getStock(),
    ganancias:  getGanancias()
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
    .sort((a, b) => (b._fecha_iso || '').localeCompare(a._fecha_iso || ''))
    .slice(0, 50)
    .map(r => { delete r._fecha_iso; return r; });
}

function getStock() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const stockSh = ss.getSheetByName(SH_STOCK);
  if (!stockSh) throw new Error('Falta hoja STOCK');

  const inicial = Number(stockSh.getRange('B1').getValue()) || 0;
  const producido = sumProducido_();
  const vendido   = sumVendido_();
  const actual    = inicial + producido - vendido;

  stockSh.getRange('B2:B4').setValues([[producido], [vendido], [actual]]);

  return {
    stock_inicial:   inicial,
    total_producido: producido,
    total_vendido:   vendido,
    stock_actual:    actual
  };
}

function getGanancias() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const prodSh = ss.getSheetByName(SH_PRODUCTO);
  const ganSh  = ss.getSheetByName(SH_GANANCIAS);
  if (!prodSh || !ganSh) {
    return { error: 'Faltan hojas PRODUCTO o GANANCIAS — corré initSheets() una vez' };
  }

  // Costos por bolsa: B2..B5 (materia prima, mano de obra, envase, otros)
  const costos = prodSh.getRange('B2:B5').getValues();
  let costoUnit = 0;
  costos.forEach(([n]) => { costoUnit += Number(n) || 0; });
  prodSh.getRange('B6').setValue(costoUnit);  // total costo / bolsa
  const precioUnit = Number(prodSh.getRange('B7').getValue()) || 0;

  const producido = sumProducido_();
  const vendido   = sumVendido_();

  const costoTotal = producido * costoUnit;
  const ingresos   = vendido   * precioUnit;
  const ganancia   = ingresos - costoTotal;
  const margen     = ingresos > 0 ? Math.round(ganancia / ingresos * 1000) / 10 : 0;

  ganSh.getRange('B2:B7').setValues([
    [producido], [vendido], [costoTotal], [ingresos], [ganancia], [margen]
  ]);

  return {
    bolsas_producidas: producido,
    bolsas_vendidas:   vendido,
    costo_unitario:    costoUnit,
    precio_unitario:   precioUnit,
    costo_total:       costoTotal,
    ingresos:          ingresos,
    ganancia_bruta:    ganancia,
    margen_pct:        margen
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

// === SIMPLE TRIGGER ===
// Auto-completa id, fecha_carga y estado cuando alguien escribe en la columna
// cliente (C) de una fila nueva de PEDIDOS. NO requiere instalar trigger — se
// dispara solo al editar la planilla.
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SH_PEDIDOS) return;
    if (e.range.getColumn() !== 3) return; // C = cliente
    const row = e.range.getRow();
    if (row < 2) return;
    if (e.value === undefined || e.value === '') return;

    const idCell     = sh.getRange(row, 1);
    const fechaCell  = sh.getRange(row, 2);
    const estadoCell = sh.getRange(row, 6);

    if (!idCell.getValue()) {
      // Próximo id = max + 1 sobre todos los ids existentes
      const last = sh.getLastRow();
      let maxId = 0;
      if (last >= 2) {
        const ids = sh.getRange(2, 1, last - 1, 1).getValues();
        ids.forEach(([v]) => {
          const n = Number(v);
          if (!isNaN(n) && n > maxId) maxId = n;
        });
      }
      idCell.setValue(maxId + 1);
    }
    if (!fechaCell.getValue()) fechaCell.setValue(new Date());
    if (!estadoCell.getValue()) estadoCell.setValue('pendiente');
  } catch (_) { /* silenciado */ }
}

// === HELPERS ===
function sumProducido_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const prodSh = ss.getSheetByName(SH_PRODUCCION);
  if (!prodSh || prodSh.getLastRow() < 2) return 0;
  let total = 0;
  prodSh.getRange(2, 2, prodSh.getLastRow() - 1, 1).getValues()
    .forEach(([n]) => { total += Number(n) || 0; });
  return total;
}

function sumVendido_() {
  // Columnas PEDIDOS: A=id B=fecha_carga C=cliente D=cantidad_bolsas E=obs_pedido F=estado
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const pedSh = ss.getSheetByName(SH_PEDIDOS);
  if (!pedSh || pedSh.getLastRow() < 2) return 0;
  let total = 0;
  pedSh.getRange(2, 1, pedSh.getLastRow() - 1, 6).getValues()
    .forEach(row => {
      if (String(row[5]).toLowerCase() === 'entregado') {
        total += Number(row[3]) || 0;
      }
    });
  return total;
}

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

// === INIT: ejecutar una vez para crear/actualizar las hojas ===
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ensure = (name, headers) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    return sh;
  };

  // PEDIDOS, PRODUCCION
  ensure(SH_PEDIDOS, [
    'id', 'fecha_carga', 'cliente', 'cantidad_bolsas',
    'observacion_pedido', 'estado', 'fecha_entrega',
    'observacion_entrega', 'link_remito'
  ]);
  ensure(SH_PRODUCCION, ['fecha', 'bolsas_producidas', 'operario']);

  // STOCK
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

  // CLIENTES
  ensure(SH_CLIENTES, ['nombre', 'telefono', 'notas']);

  // PRODUCTO (costos por bolsa + precio venta)
  let prod = ss.getSheetByName(SH_PRODUCTO);
  if (!prod) prod = ss.insertSheet(SH_PRODUCTO);
  if (prod.getLastRow() === 0) {
    prod.getRange('A1:B7').setValues([
      ['concepto',         'valor por bolsa ($)'],
      ['materia_prima',    0],
      ['mano_de_obra',     0],
      ['envase',           0],
      ['otros',            0],
      ['costo_total',      0],   // calculado por getGanancias()
      ['precio_venta',     0]
    ]);
  }

  // GANANCIAS (totales)
  let gan = ss.getSheetByName(SH_GANANCIAS);
  if (!gan) gan = ss.insertSheet(SH_GANANCIAS);
  if (gan.getLastRow() === 0) {
    gan.getRange('A1:B7').setValues([
      ['concepto',          'valor'],
      ['bolsas_producidas', 0],
      ['bolsas_vendidas',   0],
      ['costo_total',       0],
      ['ingresos',          0],
      ['ganancia_bruta',    0],
      ['margen_%',          0]
    ]);
  }

  applyClientesValidation_();
}

// Aplica el dropdown de clientes a la columna C (cliente) de PEDIDOS,
// usando como fuente CLIENTES!A2:A. Permite tipear nombres nuevos (allowInvalid).
function applyClientesValidation_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const pedidos  = ss.getSheetByName(SH_PEDIDOS);
  const clientes = ss.getSheetByName(SH_CLIENTES);
  if (!pedidos || !clientes) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(clientes.getRange('A2:A'), true)
    .setAllowInvalid(true)
    .build();
  const maxRows = Math.max(pedidos.getMaxRows() - 1, 1000);
  pedidos.getRange(2, 3, maxRows, 1).setDataValidation(rule);
}
