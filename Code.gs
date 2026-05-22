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

// === SIMPLE TRIGGER ===
// Auto-completa id, fecha_carga y estado cuando alguien escribe en la columna
// cliente (C) de una fila nueva de PEDIDOS. NO requiere instalar trigger.
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SH_PEDIDOS) return;
    if (e.range.getColumn() !== 3) return;
    const row = e.range.getRow();
    if (row < 2) return;
    if (e.value === undefined || e.value === '') return;

    const idCell     = sh.getRange(row, 1);
    const fechaCell  = sh.getRange(row, 2);
    const estadoCell = sh.getRange(row, 6);

    if (!idCell.getValue()) {
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
      'observacion_pedido', 'estado', 'fecha_entrega',
      'observacion_entrega', 'link_remito'
    ]);
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
    formatList_(ped, {
      cols: 9,
      widths: [60, 150, 200, 130, 220, 110, 150, 220, 200],
      headerHeight: 36,
      rowHeight: 28,
      numberCols: [4],
      dateCols: [2, 7]
    });
    // Conditional formatting en estado (col F)
    const estadoRange = ped.getRange(2, 6, ped.getMaxRows() - 1, 1);
    const rulePend = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('pendiente')
      .setBackground(COLOR_PEND_BG).setFontColor(COLOR_PEND_TXT).setBold(true)
      .setRanges([estadoRange]).build();
    const ruleEntr = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('entregado')
      .setBackground(COLOR_ENTR_BG).setFontColor(COLOR_ENTR_TXT).setBold(true)
      .setRanges([estadoRange]).build();
    ped.setConditionalFormatRules([rulePend, ruleEntr]);
    // Centrar estado e id
    ped.getRange(2, 1, ped.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    ped.getRange(2, 6, ped.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    ped.getRange(2, 4, ped.getMaxRows() - 1, 1).setHorizontalAlignment('right');
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
