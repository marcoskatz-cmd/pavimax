/**
 * Backend del panel PAVIMAX.
 * Lee/escribe en la planilla "PAVIMAX" y guarda fotos de remitos en Drive.
 *
 * Desplegar como Web App:
 *   Ejecutar como: yo (Marcos)
 *   Quien tiene acceso: cualquier persona
 * Copiar la URL /exec y pegarla en index.html (CONFIG.API_URL).
 */

// === CONFIG ===
const SHEET_ID  = '1r3LxaldzRNRA3GMenSbOfPk1nXZkjQvo9BJOa01dqv0';
const FOLDER_ID = '12mjpJkxDplBqp4XPnrgm9uz2gWWe2f-z';

// === NOMBRES DE HOJAS ===
const SH_PEDIDOS    = 'PEDIDOS';
const SH_PRODUCCION = 'PRODUCCION';
const SH_STOCK      = 'STOCK';
const SH_CLIENTES   = 'CLIENTES';
const SH_PRODUCTO   = 'PRODUCTO';
const SH_GANANCIAS  = 'GANANCIAS';

// === CAPACIDAD ===
const CAPACIDAD_DIARIA = 120;  // bolsas que se pueden producir por día
const HORIZONTE_DIAS   = 14;   // ventana hacia adelante para calcular capacidad

// === ESTILO ===
const COLOR_BRAND       = '#157f3d';
const COLOR_BRAND_DARK  = '#0f5e2d';
const COLOR_BRAND_SOFT  = '#e8f5ec';
const COLOR_HEADER_TXT  = '#ffffff';
const COLOR_LABEL_BG    = '#f5f6f8';
const COLOR_HILITE_BG   = '#d1fae5';
const COLOR_HILITE_TXT  = '#0f5e2d';
const COLOR_BORDER      = '#d1d5db';
const COLOR_PEND_BG     = '#fef3c7';
const COLOR_PEND_TXT    = '#92400e';
const COLOR_ENTR_BG     = '#d1fae5';
const COLOR_ENTR_TXT    = '#065f46';

// === HTTP ===
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'all';
    let data;
    if (action === 'pendientes')      data = { pendientes: getPendientes() };
    else if (action === 'entregados') data = { entregados: getEntregados() };
    else if (action === 'stock')      data = { stock: getStock() };
    else if (action === 'ganancias')  data = { ganancias: getGanancias() };
    else if (action === 'produccion') data = { produccion: getProduccion() };
    else if (action === 'capacidad')  data = { capacidad: getCapacidad() };
    else if (action === 'clientes')   data = { clientes: getClientes() };
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
    if (action === 'entregar')                    result = entregarPedido(body);
    else if (action === 'producir')               result = registrarProduccion(body);
    else if (action === 'cargarPedido')           result = cargarPedido(body);
    else if (action === 'addCliente')             result = addCliente(body);
    else if (action === 'actualizarFechaEntrega') result = actualizarFechaEntrega(body);
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
    ganancias:  getGanancias(),
    produccion: getProduccion(),
    capacidad:  getCapacidad(),
    clientes:   getClientes()
  };
}

// Últimas 5 cargas de producción, más recientes primero
function getProduccion() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_PRODUCCION);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  return values
    .filter(row => row[0])
    .map(([fecha, cantidad]) => ({
      fecha: fmt(fecha),
      _ts: fecha instanceof Date ? fecha.getTime() : 0,
      cantidad: Number(cantidad) || 0
    }))
    .sort((a, b) => b._ts - a._ts)
    .slice(0, 5)
    .map(r => { delete r._ts; return r; });
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
      fecha_entrega_solicitada: fmtDate_(r.fecha_entrega_solicitada),
      fecha_entrega_iso: toIsoDate_(r.fecha_entrega_solicitada),
      observacion_pedido: r.observacion_pedido || ''
    }))
    .sort((a, b) => (a.fecha_entrega_iso || '9999').localeCompare(b.fecha_entrega_iso || '9999'));
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

// STOCK layout: row 1 = header ("Concepto", "Bolsas"), row 2..5 = data
function getStock() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const stockSh = ss.getSheetByName(SH_STOCK);
  if (!stockSh) throw new Error('Falta hoja STOCK');

  const inicial   = Number(stockSh.getRange('B2').getValue()) || 0;
  const producido = sumProducido_();
  const vendido   = sumVendido_();
  const actual    = inicial + producido - vendido;

  stockSh.getRange('B3:B5').setValues([[producido], [vendido], [actual]]);

  return {
    stock_inicial:   inicial,
    total_producido: producido,
    total_vendido:   vendido,
    stock_actual:    actual
  };
}

// PRODUCTO layout: row 1 header, rows 2-5 costos, row 6 costo total (calc), row 7 precio
// GANANCIAS layout: row 1 header, rows 2-7 datos
function getGanancias() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const prodSh = ss.getSheetByName(SH_PRODUCTO);
  const ganSh  = ss.getSheetByName(SH_GANANCIAS);
  if (!prodSh || !ganSh) {
    return { error: 'Faltan hojas PRODUCTO o GANANCIAS — corré initSheets() una vez' };
  }

  const costos = prodSh.getRange('B2:B5').getValues();
  let costoUnit = 0;
  costos.forEach(([n]) => { costoUnit += Number(n) || 0; });
  prodSh.getRange('B6').setValue(costoUnit);
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

function cargarPedido(body) {
  const cliente = String(body.cliente || '').trim();
  const cantidad = Number(body.cantidad);
  const fechaStr = String(body.fecha_entrega || '').trim();
  const obs = String(body.observacion || '').trim();
  if (!cliente)  throw new Error('Falta cliente');
  if (!cantidad || cantidad <= 0) throw new Error('Cantidad inválida');
  if (!fechaStr) throw new Error('Falta fecha de entrega');
  const fechaEntrega = parseIsoDate_(fechaStr);

  // Auto-agregar cliente a CLIENTES si no existía (silencioso)
  try { addCliente({ nombre: cliente }); } catch (_) {}

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_PEDIDOS);
  if (!sh) throw new Error('Falta hoja PEDIDOS');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  // Próximo id
  let maxId = 0;
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
      .forEach(([v]) => {
        const n = Number(v);
        if (!isNaN(n) && n > maxId) maxId = n;
      });
  }
  const id = maxId + 1;

  // Construir fila siguiendo el orden real de columnas del sheet
  const valMap = {
    'id': id,
    'fecha_carga': new Date(),
    'cliente': cliente,
    'cantidad_bolsas': cantidad,
    'fecha_entrega_solicitada': fechaEntrega,
    'observacion_pedido': obs,
    'estado': 'pendiente',
    'fecha_entrega': '',
    'observacion_entrega': '',
    'link_remito': ''
  };
  const row = headers.map(h => valMap[h] !== undefined ? valMap[h] : '');
  sh.appendRow(row);

  return { id };
}

// === CAPACIDAD ===
// Devuelve un breakdown día por día de capacidad acumulada vs. bolsas comprometidas
function getCapacidad() {
  const stockObj = getStock();
  const stockActual = stockObj.stock_actual;

  // Pedidos pendientes con fecha objetivo
  const pendientes = readSheet(SH_PEDIDOS)
    .filter(r => String(r.estado || '').toLowerCase() === 'pendiente')
    .map(r => ({
      fechaIso: toIsoDate_(r.fecha_entrega_solicitada),
      cantidad: Number(r.cantidad_bolsas) || 0
    }))
    .filter(p => p.fechaIso && p.cantidad > 0);

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const horizonte = [];
  for (let i = 0; i <= HORIZONTE_DIAS; i++) {
    const d = new Date(hoy);
    d.setDate(hoy.getDate() + i);
    const iso = toIsoDate_(d);
    const capacidadAcum = stockActual + CAPACIDAD_DIARIA * i;
    const committedAcum = pendientes
      .filter(p => p.fechaIso <= iso)
      .reduce((s, p) => s + p.cantidad, 0);
    horizonte.push({
      fecha: iso,
      dia_offset: i,
      capacidad_acum: capacidadAcum,
      committed_acum: committedAcum,
      disponible: capacidadAcum - committedAcum
    });
  }
  return {
    capacidad_diaria: CAPACIDAD_DIARIA,
    stock_actual:     stockActual,
    horizonte:        horizonte
  };
}

function addCliente(body) {
  const nombre = String(body.nombre || '').trim();
  if (!nombre) throw new Error('Falta nombre');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_CLIENTES);
  if (!sh) throw new Error('Falta hoja CLIENTES');
  // Existe (case-insensitive)?
  if (sh.getLastRow() >= 2) {
    const existing = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const yaEsta = existing.some(([n]) =>
      String(n || '').trim().toLowerCase() === nombre.toLowerCase());
    if (yaEsta) return { nombre, existed: true };
  }
  sh.appendRow([nombre, '', '']);
  return { nombre, existed: false };
}

function actualizarFechaEntrega(body) {
  const id = String(body.id || '').trim();
  const fechaStr = String(body.fecha_entrega || '').trim();
  if (!id) throw new Error('Falta id');
  if (!fechaStr) throw new Error('Falta fecha');
  const fecha = parseIsoDate_(fechaStr);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_PEDIDOS);
  if (!sh) throw new Error('Falta hoja PEDIDOS');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const colId    = headers.indexOf('id') + 1;
  const colFecha = headers.indexOf('fecha_entrega_solicitada') + 1;
  if (colId < 1 || colFecha < 1) throw new Error('Faltan columnas id o fecha_entrega_solicitada');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('No hay pedidos');
  const ids = sh.getRange(2, colId, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      sh.getRange(2 + i, colFecha).setValue(fecha);
      return { id, fecha_entrega: fechaStr };
    }
  }
  throw new Error('Pedido no encontrado: ' + id);
}

function getClientes() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_CLIENTES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(([n]) => String(n || '').trim())
    .filter(n => n);
}

// === SIMPLE TRIGGER ===
// Auto-completa id, fecha_carga y estado cuando alguien escribe en la columna
// cliente (C) de una fila nueva de PEDIDOS. NO requiere instalar trigger.
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SH_PEDIDOS) return;
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(h => String(h).trim());
    const colCliente = headers.indexOf('cliente') + 1;
    const colId      = headers.indexOf('id') + 1;
    const colFecha   = headers.indexOf('fecha_carga') + 1;
    const colEstado  = headers.indexOf('estado') + 1;
    if (colCliente < 1) return;
    if (e.range.getColumn() !== colCliente) return;
    const row = e.range.getRow();
    if (row < 2) return;
    if (e.value === undefined || e.value === '') return;

    if (colId > 0) {
      const idCell = sh.getRange(row, colId);
      if (!idCell.getValue()) {
        const last = sh.getLastRow();
        let maxId = 0;
        if (last >= 2) {
          sh.getRange(2, colId, last - 1, 1).getValues().forEach(([v]) => {
            const n = Number(v);
            if (!isNaN(n) && n > maxId) maxId = n;
          });
        }
        idCell.setValue(maxId + 1);
      }
    }
    if (colFecha  > 0 && !sh.getRange(row, colFecha ).getValue()) sh.getRange(row, colFecha ).setValue(new Date());
    if (colEstado > 0 && !sh.getRange(row, colEstado).getValue()) sh.getRange(row, colEstado).setValue('pendiente');
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
  // Por headers — inmune a cambios de orden de columnas
  const rows = readSheet(SH_PEDIDOS);
  let total = 0;
  rows.forEach(r => {
    if (String(r.estado || '').toLowerCase() === 'entregado') {
      total += Number(r.cantidad_bolsas) || 0;
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

function fmtDate_(d) {
  if (!d) return '';
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(d);
}

function toIsoDate_(d) {
  if (!d) return '';
  const date = d instanceof Date ? d : new Date(d);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function parseIsoDate_(s) {
  const parts = String(s).split('-');
  if (parts.length !== 3) throw new Error('Fecha inválida: ' + s);
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

// =============================================================
// INIT + MIGRACIÓN + FORMATO VISUAL
// Ejecutá `initSheets` UNA VEZ después de pegar este archivo. Es
// idempotente: respeta los datos existentes y migra los layouts
// viejos (snake_case sin header) al nuevo (header + labels pretty).
// =============================================================
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // --- PEDIDOS ---
  let ped = ss.getSheetByName(SH_PEDIDOS);
  if (!ped) {
    ped = ss.insertSheet(SH_PEDIDOS);
    ped.appendRow([
      'id', 'fecha_carga', 'cliente', 'cantidad_bolsas',
      'fecha_entrega_solicitada',
      'observacion_pedido', 'estado', 'fecha_entrega',
      'observacion_entrega', 'link_remito'
    ]);
  } else {
    // Migración: insertar fecha_entrega_solicitada después de cantidad_bolsas si no existe
    const headers = ped.getRange(1, 1, 1, ped.getLastColumn()).getValues()[0]
      .map(h => String(h).trim());
    if (headers.indexOf('fecha_entrega_solicitada') === -1) {
      const colCantidad = headers.indexOf('cantidad_bolsas') + 1;
      if (colCantidad > 0) {
        ped.insertColumnAfter(colCantidad);
        ped.getRange(1, colCantidad + 1).setValue('fecha_entrega_solicitada');
      }
    }
  }

  // --- PRODUCCION ---
  let prodLog = ss.getSheetByName(SH_PRODUCCION);
  if (!prodLog) {
    prodLog = ss.insertSheet(SH_PRODUCCION);
    prodLog.appendRow(['fecha', 'bolsas_producidas', 'operario']);
  }

  // --- CLIENTES ---
  let cli = ss.getSheetByName(SH_CLIENTES);
  if (!cli) {
    cli = ss.insertSheet(SH_CLIENTES);
    cli.appendRow(['nombre', 'telefono', 'notas']);
  }

  // --- STOCK (con migración de layout viejo) ---
  let stock = ss.getSheetByName(SH_STOCK);
  if (!stock) {
    stock = ss.insertSheet(SH_STOCK);
    stock.getRange('A1:B5').setValues([
      ['Concepto',        'Bolsas'],
      ['Stock inicial',   0],
      ['Total producido', 0],
      ['Total vendido',   0],
      ['Stock actual',    0]
    ]);
  } else {
    // Migrar si A1 contiene un label en vez de header
    const a1 = String(stock.getRange('A1').getValue()).toLowerCase().trim();
    if (a1 === 'stock_inicial' || a1 === '') {
      stock.insertRowBefore(1);
    }
    stock.getRange('A1:B1').setValues([['Concepto', 'Bolsas']]);
    stock.getRange('A2').setValue('Stock inicial');
    stock.getRange('A3').setValue('Total producido');
    stock.getRange('A4').setValue('Total vendido');
    stock.getRange('A5').setValue('Stock actual');
  }

  // --- PRODUCTO ---
  let prod = ss.getSheetByName(SH_PRODUCTO);
  if (!prod) {
    prod = ss.insertSheet(SH_PRODUCTO);
    prod.getRange('A1:B7').setValues([
      ['Concepto',         'Valor por bolsa ($)'],
      ['Materia prima',    0],
      ['Mano de obra',     0],
      ['Envase',           0],
      ['Otros',            0],
      ['Costo total',      0],
      ['Precio de venta',  0]
    ]);
  } else {
    // Pretty labels (en caso de versión vieja con snake_case)
    prod.getRange('A1:B1').setValues([['Concepto', 'Valor por bolsa ($)']]);
    const pretty = ['Materia prima', 'Mano de obra', 'Envase', 'Otros', 'Costo total', 'Precio de venta'];
    pretty.forEach((label, i) => prod.getRange(2 + i, 1).setValue(label));
  }

  // --- GANANCIAS ---
  let gan = ss.getSheetByName(SH_GANANCIAS);
  if (!gan) {
    gan = ss.insertSheet(SH_GANANCIAS);
    gan.getRange('A1:B7').setValues([
      ['Concepto',           'Valor'],
      ['Bolsas producidas',  0],
      ['Bolsas vendidas',    0],
      ['Costo total ($)',    0],
      ['Ingresos ($)',       0],
      ['Ganancia bruta ($)', 0],
      ['Margen (%)',         0]
    ]);
  } else {
    gan.getRange('A1:B1').setValues([['Concepto', 'Valor']]);
    const pretty = ['Bolsas producidas','Bolsas vendidas','Costo total ($)','Ingresos ($)','Ganancia bruta ($)','Margen (%)'];
    pretty.forEach((label, i) => gan.getRange(2 + i, 1).setValue(label));
  }

  // Reordenar las hojas (de izq a der: las dashboards primero)
  reorderSheets_(ss, [SH_STOCK, SH_GANANCIAS, SH_PEDIDOS, SH_PRODUCCION, SH_CLIENTES, SH_PRODUCTO]);

  applyClientesValidation_();
  applyFormatting_();
}

function reorderSheets_(ss, order) {
  order.forEach((name, idx) => {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(idx + 1); }
  });
}

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

// === FORMATO VISUAL ===
function applyFormatting_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // PEDIDOS
  const ped = ss.getSheetByName(SH_PEDIDOS);
  if (ped) {
    const headers = ped.getRange(1, 1, 1, ped.getLastColumn()).getValues()[0]
      .map(h => String(h).trim());
    const colOf = name => headers.indexOf(name) + 1;
    const colId        = colOf('id');
    const colFCarga    = colOf('fecha_carga');
    const colCant      = colOf('cantidad_bolsas');
    const colFSolic    = colOf('fecha_entrega_solicitada');
    const colEstado    = colOf('estado');
    const colFEntrega  = colOf('fecha_entrega');
    const totalCols    = headers.length;

    // Anchos: armar dinámicamente
    const widths = headers.map(h => ({
      'id': 60, 'fecha_carga': 150, 'cliente': 200, 'cantidad_bolsas': 130,
      'fecha_entrega_solicitada': 150, 'observacion_pedido': 220, 'estado': 110,
      'fecha_entrega': 150, 'observacion_entrega': 220, 'link_remito': 200
    })[h] || 140);
    const numberCols = [colCant].filter(c => c > 0);
    const dateCols   = [colFCarga, colFSolic, colFEntrega].filter(c => c > 0);

    formatList_(ped, {
      cols: totalCols, widths,
      headerHeight: 36, rowHeight: 28,
      numberCols, dateCols
    });
    // Fecha solicitada usa solo fecha (sin hora)
    if (colFSolic > 0) {
      ped.getRange(2, colFSolic, ped.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    }
    // Conditional formatting en estado
    if (colEstado > 0) {
      const estadoRange = ped.getRange(2, colEstado, ped.getMaxRows() - 1, 1);
      const rulePend = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('pendiente')
        .setBackground(COLOR_PEND_BG).setFontColor(COLOR_PEND_TXT).setBold(true)
        .setRanges([estadoRange]).build();
      const ruleEntr = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('entregado')
        .setBackground(COLOR_ENTR_BG).setFontColor(COLOR_ENTR_TXT).setBold(true)
        .setRanges([estadoRange]).build();
      ped.setConditionalFormatRules([rulePend, ruleEntr]);
      ped.getRange(2, colEstado, ped.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    }
    if (colId   > 0) ped.getRange(2, colId,   ped.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    if (colCant > 0) ped.getRange(2, colCant, ped.getMaxRows() - 1, 1).setHorizontalAlignment('right');
  }

  // PRODUCCION
  const prodLog = ss.getSheetByName(SH_PRODUCCION);
  if (prodLog) {
    formatList_(prodLog, {
      cols: 3,
      widths: [180, 180, 220],
      headerHeight: 36, rowHeight: 28,
      numberCols: [2], dateCols: [1]
    });
    prodLog.getRange(2, 2, prodLog.getMaxRows() - 1, 1).setHorizontalAlignment('right');
  }

  // CLIENTES
  const cli = ss.getSheetByName(SH_CLIENTES);
  if (cli) {
    formatList_(cli, {
      cols: 3,
      widths: [240, 160, 320],
      headerHeight: 36, rowHeight: 28
    });
  }

  // STOCK (dashboard, 5 filas, valor entero)
  const stock = ss.getSheetByName(SH_STOCK);
  if (stock) {
    formatDashboard_(stock, {
      rows: 5,
      valueFormat: '#,##0" bolsas"',
      highlightRow: 5,
      colWidths: [260, 220]
    });
  }

  // PRODUCTO (dashboard, 7 filas, valor en $)
  const prod = ss.getSheetByName(SH_PRODUCTO);
  if (prod) {
    formatDashboard_(prod, {
      rows: 7,
      valueFormat: '"$" #,##0',
      highlightRow: 6,
      colWidths: [260, 220]
    });
    // Filas 2..5 son inputs del usuario: bg distinto
    prod.getRange('A2:B5').setBackground('#fffbeb'); // amarillo suave = editar
    // Fila 7 (precio de venta) también es input
    prod.getRange('A7:B7').setBackground('#fffbeb');
  }

  // GANANCIAS (dashboard, 7 filas, formato mixto)
  const gan = ss.getSheetByName(SH_GANANCIAS);
  if (gan) {
    formatDashboard_(gan, {
      rows: 7,
      valueFormat: '"$" #,##0',
      highlightRow: 6,
      colWidths: [260, 240],
      customFormats: { 2: '#,##0" bolsas"', 3: '#,##0" bolsas"', 7: '0.0"%"' }
    });
  }

  // Apagar grillas en todas
  [SH_PEDIDOS, SH_PRODUCCION, SH_CLIENTES, SH_STOCK, SH_PRODUCTO, SH_GANANCIAS].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s) s.setHiddenGridlines(true);
  });
}

function formatList_(sh, opt) {
  const cols = opt.cols;

  // Header
  const hdr = sh.getRange(1, 1, 1, cols);
  hdr.setBackground(COLOR_BRAND)
     .setFontColor(COLOR_HEADER_TXT)
     .setFontWeight('bold')
     .setFontSize(11)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
  sh.setRowHeight(1, opt.headerHeight || 32);
  sh.setFrozenRows(1);

  // Anchos de columna
  if (opt.widths) opt.widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Alturas de fila por defecto
  if (opt.rowHeight) {
    const last = sh.getMaxRows();
    if (last > 1) {
      try { sh.setRowHeights(2, last - 1, opt.rowHeight); } catch (_) {}
    }
  }

  // Banding (filas alternadas) — limpia las anteriores primero
  sh.getBandings().forEach(b => b.remove());
  const bandRange = sh.getRange(1, 1, sh.getMaxRows(), cols);
  const banding = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banding.setHeaderRowColor(COLOR_BRAND).setFirstRowColor('#ffffff').setSecondRowColor('#f9fafb');

  // Formatos
  if (opt.numberCols) {
    opt.numberCols.forEach(c => sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0'));
  }
  if (opt.dateCols) {
    opt.dateCols.forEach(c => sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy HH:mm'));
  }

  // Wrap en columnas de texto largas — solo aplicar a cells; mejora legibilidad
  sh.getRange(2, 1, sh.getMaxRows() - 1, cols).setVerticalAlignment('middle');
}

function formatDashboard_(sh, opt) {
  // Header
  const hdr = sh.getRange(1, 1, 1, 2);
  hdr.setBackground(COLOR_BRAND)
     .setFontColor(COLOR_HEADER_TXT)
     .setFontWeight('bold')
     .setFontSize(12)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
  sh.setRowHeight(1, 36);
  sh.setFrozenRows(1);

  // Anchos
  const w = opt.colWidths || [240, 200];
  sh.setColumnWidth(1, w[0]);
  sh.setColumnWidth(2, w[1]);

  // Labels (col A) bg suave + bold
  const labels = sh.getRange(2, 1, opt.rows - 1, 1);
  labels.setBackground(COLOR_LABEL_BG)
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle')
        .setFontSize(11);

  // Values (col B)
  const values = sh.getRange(2, 2, opt.rows - 1, 1);
  values.setBackground('#ffffff')
        .setHorizontalAlignment('right')
        .setVerticalAlignment('middle')
        .setFontSize(12);
  if (opt.valueFormat) values.setNumberFormat(opt.valueFormat);
  if (opt.customFormats) {
    Object.keys(opt.customFormats).forEach(row => {
      sh.getRange(Number(row), 2).setNumberFormat(opt.customFormats[row]);
    });
  }

  // Filas más altas
  for (let r = 2; r <= opt.rows; r++) sh.setRowHeight(r, 32);

  // Fila destacada (total)
  if (opt.highlightRow) {
    const r = sh.getRange(opt.highlightRow, 1, 1, 2);
    r.setBackground(COLOR_HILITE_BG)
     .setFontColor(COLOR_HILITE_TXT)
     .setFontWeight('bold')
     .setFontSize(14);
    sh.setRowHeight(opt.highlightRow, 42);
  }

  // Bordes
  sh.getRange(1, 1, opt.rows, 2)
    .setBorder(true, true, true, true, true, true, COLOR_BORDER, SpreadsheetApp.BorderStyle.SOLID);
}
