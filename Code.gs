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
const SH_USUARIOS   = 'USUARIOS';
const SH_PARCIALES  = 'ENTREGAS_PARCIALES';
const SH_CONFIG     = 'CONFIG';

// === DEFAULTS ===
const CAPACIDAD_DIARIA_DEFAULT = 3000;  // kg/día si CONFIG no tiene valor (≈120 bolsas)
const HORIZONTE_DIAS           = 14;   // ventana hacia adelante para calcular capacidad
const USUARIOS_SEED            = ['Julia Katz', 'Agustin De Gregorio', 'Gerardo Guerrero'];

// === PRODUCTOS (multi-producto: material en kg, emulsión en botellas sin stock/producción) ===
const PRODUCTOS = ['bolsas_25kg', 'bigbag_1000kg', 'emul_1l', 'emul_5l', 'emul_20l'];
const KG_POR = { bolsas_25kg: 25, bigbag_1000kg: 1000, emul_1l: 0, emul_5l: 0, emul_20l: 0 };
const PARC_HEADERS = [
  'id_parcial', 'id_pedido', 'fecha_planeada',
  'bolsas_25kg', 'bigbag_1000kg', 'emul_1l', 'emul_5l', 'emul_20l',
  'estado', 'fecha_entrega', 'observacion_entrega', 'link_remito', 'usuario_entrega'
];

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
const COLOR_PARC_BG     = '#dbeafe';
const COLOR_PARC_TXT    = '#1e40af';

// === HTTP ===
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'all';
    let data;
    if (action === 'pendientes')           data = { pendientes: getPendientes() };
    else if (action === 'entregados')      data = { entregados: getEntregados() };
    else if (action === 'stock')           data = { stock: getStock() };
    else if (action === 'ganancias')       data = { ganancias: getGanancias() };
    else if (action === 'produccion')      data = { produccion: getProduccion() };
    else if (action === 'capacidad')       data = { capacidad: getCapacidad() };
    else if (action === 'clientes')        data = { clientes: getClientes() };
    else if (action === 'usuarios')        data = { usuarios: getUsuarios() };
    else if (action === 'capacidadDiaria') data = { capacidad_diaria: getCapacidadDiaria() };
    else                                   data = getAll();
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
    else if (action === 'addUsuario')             result = addUsuario(body);
    else if (action === 'actualizarFechaEntrega') result = actualizarFechaEntrega(body);
    else if (action === 'editarParciales')        result = editarParciales(body);
    else if (action === 'setCapacidadDiaria')     result = setCapacidadDiaria(body);
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
    pendientes:        getPendientes(),
    entregados:        getEntregados(),
    stock:             getStock(),
    ganancias:         getGanancias(),
    produccion:        getProduccion(),
    capacidad:         getCapacidad(),
    clientes:          getClientes(),
    usuarios:          getUsuarios(),
    capacidad_diaria:  getCapacidadDiaria()
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
  const pedidos   = readSheet(SH_PEDIDOS);
  const parciales = readSheet(SH_PARCIALES);
  const out = [];
  pedidos.forEach(p => {
    const total = Number(p.cantidad_bolsas) || 0;
    const misParciales = parciales.filter(pa => String(pa.id_pedido) === String(p.id));
    const entregadas = misParciales.filter(pa => String(pa.estado || '').toLowerCase() === 'entregada');
    const pendientesP = misParciales
      .filter(pa => String(pa.estado || '').toLowerCase() === 'pendiente')
      .sort((a, b) => (toIsoDate_(a.fecha_planeada) || '9999').localeCompare(toIsoDate_(b.fecha_planeada) || '9999'));
    const entregado = entregadas.reduce((s, pa) => s + (Number(pa.cantidad_entregada) || 0), 0);
    const saldo = total - entregado;
    if (saldo <= 0) return;  // ya está entregado completo

    const proxima = pendientesP[0];
    const estado = entregado > 0 ? 'parcial' : 'pendiente';
    out.push({
      id: p.id,
      fecha_carga: fmt(p.fecha_carga),
      usuario_carga: p.usuario_carga || '',
      cliente: p.cliente,
      cantidad_bolsas: total,
      cantidad_entregada: entregado,
      saldo: saldo,
      observacion_pedido: p.observacion_pedido || '',
      estado: estado,
      proxima_fecha: proxima ? fmtDate_(proxima.fecha_planeada) : '',
      proxima_fecha_iso: proxima ? toIsoDate_(proxima.fecha_planeada) : '',
      proxima_cantidad: proxima ? (Number(proxima.cantidad_planeada) || 0) : 0,
      parciales_pendientes: pendientesP.map(pa => ({
        id_parcial: pa.id_parcial,
        fecha: fmtDate_(pa.fecha_planeada),
        fecha_iso: toIsoDate_(pa.fecha_planeada),
        cantidad: Number(pa.cantidad_planeada) || 0
      })),
      // Compat con clientes viejos
      fecha_entrega_solicitada: proxima ? fmtDate_(proxima.fecha_planeada) : '',
      fecha_entrega_iso:        proxima ? toIsoDate_(proxima.fecha_planeada) : ''
    });
  });
  return out.sort((a, b) =>
    (a.proxima_fecha_iso || '9999').localeCompare(b.proxima_fecha_iso || '9999'));
}

function getEntregados() {
  const pedidos   = readSheet(SH_PEDIDOS);
  const parciales = readSheet(SH_PARCIALES);
  const out = [];
  pedidos.forEach(p => {
    const total = Number(p.cantidad_bolsas) || 0;
    const misParciales = parciales.filter(pa => String(pa.id_pedido) === String(p.id));
    if (misParciales.length === 0) return;
    const entregadas = misParciales.filter(pa => String(pa.estado || '').toLowerCase() === 'entregada');
    const entregado = entregadas.reduce((s, pa) => s + (Number(pa.cantidad_entregada) || 0), 0);
    if (entregado < total || entregado === 0) return; // no completamente entregado
    // Ordenar entregas más recientes primero
    const entregasOrd = entregadas
      .map(pa => ({
        fecha:      fmt(pa.fecha_entrega),
        fecha_iso:  toIsoDate_(pa.fecha_entrega),
        cantidad:   Number(pa.cantidad_entregada) || 0,
        observacion: pa.observacion_entrega || '',
        link_remito: pa.link_remito || '',
        usuario_entrega: pa.usuario_entrega || ''
      }))
      .sort((a, b) => (b.fecha_iso || '').localeCompare(a.fecha_iso || ''));
    const ultima = entregasOrd[0] || {};
    out.push({
      id: p.id,
      cliente: p.cliente,
      cantidad_bolsas: total,
      observacion_pedido: p.observacion_pedido || '',
      fecha_entrega: ultima.fecha || '',
      observacion_entrega: ultima.observacion || '',
      link_remito: ultima.link_remito || '',
      usuario_entrega: ultima.usuario_entrega || '',
      entregas: entregasOrd,
      _fecha_iso: ultima.fecha_iso || ''
    });
  });
  return out
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
/**
 * Entrega de un pedido (puede ser parcial).
 * body: { id, cantidad, observacion, foto_base64, foto_mime, usuario_entrega }
 *
 * Algoritmo:
 *  - Tomar parciales pendientes ordenadas por fecha.
 *  - La primera pendiente se "convierte" en una entrega con cantidad real.
 *  - Si entregó MENOS que la planeada → partir: queda nueva parcial pendiente con el resto.
 *  - Si entregó MÁS que la planeada → consumir parciales futuras (reducir / eliminar).
 *  - Saldo total = cantidad_bolsas - sum(entregadas). Si 0 → estado='entregado', si <total → 'parcial'.
 */
function entregarPedido(body) {
  const id           = String(body.id || '').trim();
  const cantidad     = Number(body.cantidad);
  const observacion  = String(body.observacion || '').trim();
  const fotoB64      = body.foto_base64 || '';
  const mime         = body.foto_mime || 'image/jpeg';
  const usuarioEntr  = String(body.usuario_entrega || '').trim();
  if (!id)           throw new Error('Falta id');
  if (!cantidad || cantidad <= 0) throw new Error('Cantidad inválida');
  if (!usuarioEntr)  throw new Error('Falta usuario_entrega');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const shPed  = ss.getSheetByName(SH_PEDIDOS);
  const shParc = ss.getSheetByName(SH_PARCIALES);
  if (!shPed)  throw new Error('Falta hoja PEDIDOS');
  if (!shParc) throw new Error('Falta hoja ENTREGAS_PARCIALES — corré initSheets()');

  // Cargar pedido
  const pedHeaders = shPed.getRange(1, 1, 1, shPed.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const colId      = pedHeaders.indexOf('id') + 1;
  const colCliente = pedHeaders.indexOf('cliente') + 1;
  const colCant    = pedHeaders.indexOf('cantidad_bolsas') + 1;
  const colEstado  = pedHeaders.indexOf('estado') + 1;
  if (colId < 1) throw new Error('PEDIDOS sin columna id');

  let rowPed = -1, total = 0, cliente = '';
  const lastP = shPed.getLastRow();
  if (lastP >= 2) {
    const data = shPed.getRange(2, 1, lastP - 1, shPed.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][colId - 1]) === id) {
        rowPed   = i + 2;
        total    = Number(data[i][colCant - 1]) || 0;
        cliente  = String(data[i][colCliente - 1] || '');
        break;
      }
    }
  }
  if (rowPed < 0) throw new Error('Pedido no encontrado: ' + id);

  // Cargar parciales del pedido
  const parcHeaders = shParc.getRange(1, 1, 1, shParc.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const cI = name => parcHeaders.indexOf(name) + 1;
  const cIdParc      = cI('id_parcial');
  const cIdPed       = cI('id_pedido');
  const cFechaP      = cI('fecha_planeada');
  const cCantP       = cI('cantidad_planeada');
  const cEstP        = cI('estado');
  const cFechaE      = cI('fecha_entrega');
  const cCantE       = cI('cantidad_entregada');
  const cObsE        = cI('observacion_entrega');
  const cLink        = cI('link_remito');
  const cUserE       = cI('usuario_entrega');
  if (cIdParc < 1 || cIdPed < 1 || cFechaP < 1 || cCantP < 1 || cEstP < 1)
    throw new Error('ENTREGAS_PARCIALES sin columnas requeridas');

  const lastPa = shParc.getLastRow();
  const allRows = lastPa >= 2 ? shParc.getRange(2, 1, lastPa - 1, shParc.getLastColumn()).getValues() : [];
  // Calcular entregado total y separar pendientes
  let entregado = 0;
  const pendientesIdx = []; // {row, cantidad, fechaIso}
  let maxIdParc = 0;
  allRows.forEach((r, i) => {
    const ipa = Number(r[cIdParc - 1]);
    if (!isNaN(ipa) && ipa > maxIdParc) maxIdParc = ipa;
    if (String(r[cIdPed - 1]) !== id) return;
    const est = String(r[cEstP - 1] || '').toLowerCase();
    if (est === 'entregada') {
      entregado += Number(r[cCantE - 1]) || 0;
    } else if (est === 'pendiente') {
      pendientesIdx.push({
        row: i + 2,
        cantidad: Number(r[cCantP - 1]) || 0,
        fechaIso: toIsoDate_(r[cFechaP - 1])
      });
    }
  });
  pendientesIdx.sort((a, b) => (a.fechaIso || '9999').localeCompare(b.fechaIso || '9999'));

  const saldo = total - entregado;
  if (cantidad > saldo) throw new Error('No podés entregar ' + cantidad + ': el saldo es ' + saldo);
  if (pendientesIdx.length === 0) throw new Error('No hay parciales pendientes (saldo: ' + saldo + ')');

  // Subir foto si vino
  let linkRemito = '';
  if (fotoB64) {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const cliSlug = cliente.replace(/[^a-z0-9]+/gi, '_');
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const name = `remito_${id}_${cliSlug}_${stamp}.jpg`;
    const blob = Utilities.newBlob(Utilities.base64Decode(fotoB64), mime, name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    linkRemito = file.getUrl();
  }

  // Convertir la primera pendiente en entregada con la cantidad real.
  const primera = pendientesIdx[0];
  const ahora = new Date();
  shParc.getRange(primera.row, cEstP).setValue('entregada');
  shParc.getRange(primera.row, cCantE).setValue(cantidad);
  shParc.getRange(primera.row, cFechaE).setValue(ahora);
  if (cObsE  > 0) shParc.getRange(primera.row, cObsE).setValue(observacion);
  if (cLink  > 0) shParc.getRange(primera.row, cLink).setValue(linkRemito);
  if (cUserE > 0) shParc.getRange(primera.row, cUserE).setValue(usuarioEntr);

  let diff = cantidad - primera.cantidad;
  if (diff < 0) {
    // Entregó menos: queda nueva parcial pendiente con el resto, misma fecha
    const fechaCell = shParc.getRange(primera.row, cFechaP).getValue();
    const nuevaId = maxIdParc + 1;
    const newRow = parcHeaders.map(h => {
      if (h === 'id_parcial')        return nuevaId;
      if (h === 'id_pedido')         return id;
      if (h === 'fecha_planeada')    return fechaCell;
      if (h === 'cantidad_planeada') return -diff;
      if (h === 'estado')            return 'pendiente';
      return '';
    });
    shParc.appendRow(newRow);
  } else if (diff > 0) {
    // Entregó más: consumir pendientes siguientes
    const filasABorrar = [];
    for (let i = 1; i < pendientesIdx.length && diff > 0; i++) {
      const p = pendientesIdx[i];
      if (p.cantidad <= diff) {
        filasABorrar.push(p.row);
        diff -= p.cantidad;
      } else {
        shParc.getRange(p.row, cCantP).setValue(p.cantidad - diff);
        diff = 0;
      }
    }
    // Borrar de abajo hacia arriba para no romper índices
    filasABorrar.sort((a, b) => b - a).forEach(r => shParc.deleteRow(r));
  }

  // Recalcular estado del pedido en la hoja PEDIDOS
  if (colEstado > 0) {
    const nuevoEntregado = entregado + cantidad;
    const nuevoEstado = nuevoEntregado >= total ? 'entregado' :
                        (nuevoEntregado > 0 ? 'parcial' : 'pendiente');
    shPed.getRange(rowPed, colEstado).setValue(nuevoEstado);
  }

  SpreadsheetApp.flush();
  return { id, cantidad_entregada: cantidad, link_remito: linkRemito, saldo: total - entregado - cantidad };
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

/**
 * Cargar pedido con usuario que carga y opcionalmente parciales.
 * body: {
 *   cliente, cantidad, observacion, usuario_carga,
 *   parciales: [{ fecha_entrega: 'YYYY-MM-DD', cantidad: N }, ...]
 * }
 * Si parciales no viene o está vacío, usa body.fecha_entrega como única parcial.
 */
function cargarPedido(body) {
  const cliente      = String(body.cliente || '').trim();
  const cantidad     = Number(body.cantidad);
  const obs          = String(body.observacion || '').trim();
  const usuarioCarga = String(body.usuario_carga || '').trim();
  if (!cliente)  throw new Error('Falta cliente');
  if (!cantidad || cantidad <= 0) throw new Error('Cantidad inválida');
  if (!usuarioCarga) throw new Error('Falta usuario_carga');

  // Resolver parciales — la fecha es opcional (pedidos "a favor" sin fecha pautada)
  let parciales = Array.isArray(body.parciales) ? body.parciales.slice() : [];
  if (parciales.length === 0) {
    const fechaStr = String(body.fecha_entrega || '').trim();
    parciales = [{ fecha_entrega: fechaStr, cantidad: cantidad }];
  }
  const parsedParciales = parciales.map(p => {
    const f = String(p.fecha_entrega || '').trim();
    const c = Number(p.cantidad);
    if (!c || c <= 0) throw new Error('Parcial con cantidad inválida');
    return { fecha: f ? parseIsoDate_(f) : null, cantidad: c };
  });
  const sumaP = parsedParciales.reduce((s, p) => s + p.cantidad, 0);
  if (sumaP !== cantidad) {
    throw new Error('La suma de parciales (' + sumaP + ') no coincide con el total (' + cantidad + ')');
  }

  // Auto-agregar cliente/usuario si no existían (silencioso)
  try { addCliente({ nombre: cliente }); } catch (_) {}
  try { addUsuario({ nombre: usuarioCarga }); } catch (_) {}

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const shPed  = ss.getSheetByName(SH_PEDIDOS);
  const shParc = ss.getSheetByName(SH_PARCIALES);
  if (!shPed)  throw new Error('Falta hoja PEDIDOS');
  if (!shParc) throw new Error('Falta hoja ENTREGAS_PARCIALES — corré initSheets()');

  const headers = shPed.getRange(1, 1, 1, shPed.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  // Próximo id de pedido
  let maxId = 0;
  if (shPed.getLastRow() >= 2) {
    shPed.getRange(2, 1, shPed.getLastRow() - 1, 1).getValues()
      .forEach(([v]) => {
        const n = Number(v);
        if (!isNaN(n) && n > maxId) maxId = n;
      });
  }
  const id = maxId + 1;

  // Construir fila del pedido. fecha_entrega_solicitada = fecha de la primera parcial con fecha (compat).
  const conFecha = parsedParciales.filter(p => p.fecha);
  const fechaPrimera = conFecha.length
    ? conFecha.slice().sort((a, b) => a.fecha.getTime() - b.fecha.getTime())[0].fecha
    : '';
  const valMap = {
    'id': id,
    'fecha_carga': new Date(),
    'usuario_carga': usuarioCarga,
    'cliente': cliente,
    'cantidad_bolsas': cantidad,
    'fecha_entrega_solicitada': fechaPrimera,
    'observacion_pedido': obs,
    'estado': 'pendiente',
    'fecha_entrega': '',
    'observacion_entrega': '',
    'link_remito': ''
  };
  const row = headers.map(h => valMap[h] !== undefined ? valMap[h] : '');
  shPed.appendRow(row);

  // Crear parciales
  const parcHeaders = shParc.getRange(1, 1, 1, shParc.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  let maxIdParc = 0;
  if (shParc.getLastRow() >= 2) {
    shParc.getRange(2, 1, shParc.getLastRow() - 1, 1).getValues()
      .forEach(([v]) => {
        const n = Number(v);
        if (!isNaN(n) && n > maxIdParc) maxIdParc = n;
      });
  }
  parsedParciales.forEach((p, idx) => {
    const idParc = maxIdParc + 1 + idx;
    const r = parcHeaders.map(h => {
      if (h === 'id_parcial')        return idParc;
      if (h === 'id_pedido')         return id;
      if (h === 'fecha_planeada')    return p.fecha || '';
      if (h === 'cantidad_planeada') return p.cantidad;
      if (h === 'estado')            return 'pendiente';
      return '';
    });
    shParc.appendRow(r);
  });

  return { id, parciales: parsedParciales.length };
}

// === CAPACIDAD ===
// Devuelve un breakdown día por día de capacidad acumulada vs. bolsas comprometidas
// (suma de parciales pendientes).
function getCapacidad() {
  const stockObj = getStock();
  const stockActual = stockObj.stock_actual;
  const capDiaria = getCapacidadDiaria();

  const parciales = readSheet(SH_PARCIALES)
    .filter(pa => String(pa.estado || '').toLowerCase() === 'pendiente')
    .map(pa => ({
      fechaIso: toIsoDate_(pa.fecha_planeada),
      cantidad: Number(pa.cantidad_planeada) || 0
    }))
    .filter(p => p.fechaIso && p.cantidad > 0);

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const horizonte = [];
  for (let i = 0; i <= HORIZONTE_DIAS; i++) {
    const d = new Date(hoy);
    d.setDate(hoy.getDate() + i);
    const iso = toIsoDate_(d);
    const capacidadAcum = stockActual + capDiaria * i;
    const committedAcum = parciales
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
    capacidad_diaria: capDiaria,
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
  if (sh.getLastRow() >= 2) {
    const existing = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const yaEsta = existing.some(([n]) =>
      String(n || '').trim().toLowerCase() === nombre.toLowerCase());
    if (yaEsta) return { nombre, existed: true };
  }
  sh.appendRow([nombre, '', '']);
  return { nombre, existed: false };
}

function addUsuario(body) {
  const nombre = String(body.nombre || '').trim();
  if (!nombre) throw new Error('Falta nombre');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_USUARIOS);
  if (!sh) throw new Error('Falta hoja USUARIOS — corré initSheets()');
  if (sh.getLastRow() >= 2) {
    const existing = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const yaEsta = existing.some(([n]) =>
      String(n || '').trim().toLowerCase() === nombre.toLowerCase());
    if (yaEsta) return { nombre, existed: true };
  }
  sh.appendRow([nombre, 'si']);
  return { nombre, existed: false };
}

/**
 * Cambiar fecha de una parcial pendiente (caso UI sencillo: 1 sola parcial).
 * body: { id (id_pedido), fecha_entrega }
 * Solo funciona si el pedido tiene exactamente 1 parcial pendiente.
 * Para casos con varias, usar editarParciales.
 */
function actualizarFechaEntrega(body) {
  const id       = String(body.id || '').trim();
  const fechaStr = String(body.fecha_entrega || '').trim();
  if (!id) throw new Error('Falta id');
  if (!fechaStr) throw new Error('Falta fecha');
  const fecha = parseIsoDate_(fechaStr);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const shParc = ss.getSheetByName(SH_PARCIALES);
  if (!shParc) throw new Error('Falta hoja ENTREGAS_PARCIALES');
  const headers = shParc.getRange(1, 1, 1, shParc.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const cIdPed  = headers.indexOf('id_pedido') + 1;
  const cFechaP = headers.indexOf('fecha_planeada') + 1;
  const cEstP   = headers.indexOf('estado') + 1;
  if (cIdPed < 1 || cFechaP < 1 || cEstP < 1) throw new Error('ENTREGAS_PARCIALES sin columnas');

  const lastRow = shParc.getLastRow();
  if (lastRow < 2) throw new Error('No hay parciales');
  const data = shParc.getRange(2, 1, lastRow - 1, shParc.getLastColumn()).getValues();
  const filas = [];
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][cIdPed - 1]) === id &&
        String(data[i][cEstP - 1] || '').toLowerCase() === 'pendiente') {
      filas.push(i + 2);
    }
  }
  if (filas.length === 0) throw new Error('No hay parciales pendientes para el pedido ' + id);
  if (filas.length > 1) throw new Error('El pedido tiene varias parciales — usá editarParciales');

  shParc.getRange(filas[0], cFechaP).setValue(fecha);
  return { id, fecha_entrega: fechaStr };
}

/**
 * Reemplazar todas las parciales pendientes de un pedido.
 * body: { id (id_pedido), parciales: [{ fecha_entrega, cantidad }, ...] }
 * La suma debe ser igual al saldo actual del pedido.
 */
function editarParciales(body) {
  const id        = String(body.id || '').trim();
  const incoming  = Array.isArray(body.parciales) ? body.parciales : [];
  if (!id) throw new Error('Falta id');
  if (incoming.length === 0) throw new Error('Lista de parciales vacía');

  const parsed = incoming.map(p => {
    const f = String(p.fecha_entrega || '').trim();
    const c = Number(p.cantidad);
    if (!c || c <= 0) throw new Error('Parcial con cantidad inválida');
    return { fecha: f ? parseIsoDate_(f) : null, cantidad: c };
  });
  const sumaNueva = parsed.reduce((s, p) => s + p.cantidad, 0);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const shPed  = ss.getSheetByName(SH_PEDIDOS);
  const shParc = ss.getSheetByName(SH_PARCIALES);
  if (!shPed || !shParc) throw new Error('Faltan hojas');

  // Resolver total del pedido
  const pedHeaders = shPed.getRange(1, 1, 1, shPed.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const colId   = pedHeaders.indexOf('id') + 1;
  const colCant = pedHeaders.indexOf('cantidad_bolsas') + 1;
  if (colId < 1 || colCant < 1) throw new Error('PEDIDOS sin columnas');
  const lastP = shPed.getLastRow();
  let total = 0;
  if (lastP >= 2) {
    const d = shPed.getRange(2, 1, lastP - 1, shPed.getLastColumn()).getValues();
    for (let i = 0; i < d.length; i++) {
      if (String(d[i][colId - 1]) === id) { total = Number(d[i][colCant - 1]) || 0; break; }
    }
  }
  if (!total) throw new Error('Pedido no encontrado: ' + id);

  // Calcular entregado y borrar pendientes actuales
  const parcHeaders = shParc.getRange(1, 1, 1, shParc.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const cI = name => parcHeaders.indexOf(name) + 1;
  const cIdParc = cI('id_parcial');
  const cIdPed  = cI('id_pedido');
  const cFechaP = cI('fecha_planeada');
  const cCantP  = cI('cantidad_planeada');
  const cEstP   = cI('estado');
  const cCantE  = cI('cantidad_entregada');

  const lastPa = shParc.getLastRow();
  const rows = lastPa >= 2 ? shParc.getRange(2, 1, lastPa - 1, shParc.getLastColumn()).getValues() : [];
  let entregado = 0;
  let maxIdParc = 0;
  const filasABorrar = [];
  rows.forEach((r, i) => {
    const ipa = Number(r[cIdParc - 1]);
    if (!isNaN(ipa) && ipa > maxIdParc) maxIdParc = ipa;
    if (String(r[cIdPed - 1]) !== id) return;
    if (String(r[cEstP - 1] || '').toLowerCase() === 'entregada') {
      entregado += Number(r[cCantE - 1]) || 0;
    } else {
      filasABorrar.push(i + 2);
    }
  });
  const saldo = total - entregado;
  if (sumaNueva !== saldo) {
    throw new Error('La suma de parciales (' + sumaNueva + ') no coincide con el saldo (' + saldo + ')');
  }

  // Borrar pendientes existentes (de abajo hacia arriba)
  filasABorrar.sort((a, b) => b - a).forEach(r => shParc.deleteRow(r));

  // Insertar nuevas pendientes
  parsed.forEach((p, idx) => {
    const idParc = maxIdParc + 1 + idx;
    const r = parcHeaders.map(h => {
      if (h === 'id_parcial')        return idParc;
      if (h === 'id_pedido')         return id;
      if (h === 'fecha_planeada')    return p.fecha || '';
      if (h === 'cantidad_planeada') return p.cantidad;
      if (h === 'estado')            return 'pendiente';
      return '';
    });
    shParc.appendRow(r);
  });

  SpreadsheetApp.flush();
  return { id, parciales: parsed.length, saldo };
}

function getClientes() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_CLIENTES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(([n]) => String(n || '').trim())
    .filter(n => n);
}

function getUsuarios() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_USUARIOS);
  if (!sh || sh.getLastRow() < 2) return [];
  const cols = sh.getLastColumn();
  return sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(cols, 2)).getValues()
    .filter(r => String(r[0] || '').trim())
    .filter(r => {
      const activo = String(r[1] || 'si').trim().toLowerCase();
      return activo !== 'no';
    })
    .map(r => String(r[0]).trim());
}

function getCapacidadDiaria() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_CONFIG);
  if (!sh || sh.getLastRow() < 2) return CAPACIDAD_DIARIA_DEFAULT;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === 'capacidad_diaria') {
      const n = Number(data[i][1]);
      return (isNaN(n) || n <= 0) ? CAPACIDAD_DIARIA_DEFAULT : n;
    }
  }
  return CAPACIDAD_DIARIA_DEFAULT;
}

function getConfigVal_(ss, clave) {
  const sh = ss.getSheetByName(SH_CONFIG);
  if (!sh || sh.getLastRow() < 2) return '';
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === clave) return data[i][1];
  }
  return '';
}
function setConfigVal_(ss, clave, valor) {
  const sh = ss.getSheetByName(SH_CONFIG);
  if (!sh) return;
  const last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === clave) { sh.getRange(i + 2, 2).setValue(valor); return; }
    }
  }
  sh.appendRow([clave, valor]);
}

function setCapacidadDiaria(body) {
  const n = Number(body.capacidad_diaria);
  if (!n || n <= 0) throw new Error('Capacidad inválida');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_CONFIG);
  if (!sh) throw new Error('Falta hoja CONFIG — corré initSheets()');
  // Buscar fila existente
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === 'capacidad_diaria') {
        sh.getRange(i + 2, 2).setValue(n);
        return { capacidad_diaria: n };
      }
    }
  }
  sh.appendRow(['capacidad_diaria', n]);
  return { capacidad_diaria: n };
}

// === SIMPLE TRIGGER ===
// Auto-completa id, fecha_carga y estado cuando alguien escribe en la columna
// cliente de una fila nueva de PEDIDOS. NO requiere instalar trigger.
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
function materialKg_(obj) {
  return (Number(obj.bolsas_25kg) || 0) * 25 + (Number(obj.bigbag_1000kg) || 0) * 1000;
}

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
  // Suma de cantidades entregadas en ENTREGAS_PARCIALES.
  const parciales = readSheet(SH_PARCIALES);
  let total = 0;
  parciales.forEach(p => {
    if (String(p.estado || '').toLowerCase() === 'entregada') {
      total += Number(p.cantidad_entregada) || 0;
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
// viejos al nuevo (USUARIOS, ENTREGAS_PARCIALES, CONFIG, etc.).
// =============================================================
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // --- PEDIDOS ---
  let ped = ss.getSheetByName(SH_PEDIDOS);
  if (!ped) {
    ped = ss.insertSheet(SH_PEDIDOS);
    ped.appendRow([
      'id', 'fecha_carga', 'usuario_carga', 'cliente', 'cantidad_bolsas',
      'fecha_entrega_solicitada',
      'observacion_pedido', 'estado', 'fecha_entrega',
      'observacion_entrega', 'link_remito'
    ]);
  } else {
    let headers = ped.getRange(1, 1, 1, ped.getLastColumn()).getValues()[0]
      .map(h => String(h).trim());
    // Migración: insertar fecha_entrega_solicitada si falta
    if (headers.indexOf('fecha_entrega_solicitada') === -1) {
      const colCantidad = headers.indexOf('cantidad_bolsas') + 1;
      if (colCantidad > 0) {
        ped.insertColumnAfter(colCantidad);
        ped.getRange(1, colCantidad + 1).setValue('fecha_entrega_solicitada');
        headers = ped.getRange(1, 1, 1, ped.getLastColumn()).getValues()[0]
          .map(h => String(h).trim());
      }
    }
    // Migración: insertar usuario_carga después de fecha_carga si falta
    if (headers.indexOf('usuario_carga') === -1) {
      const colFCarga = headers.indexOf('fecha_carga') + 1;
      const insertAfter = colFCarga > 0 ? colFCarga : 1;
      ped.insertColumnAfter(insertAfter);
      ped.getRange(1, insertAfter + 1).setValue('usuario_carga');
    }
  }

  // --- PRODUCCION ---
  let prodLog = ss.getSheetByName(SH_PRODUCCION);
  const PROD_HEADERS = ['fecha', 'bolsas_25kg', 'bigbag_1000kg', 'kg_total', 'operario'];
  if (!prodLog) {
    prodLog = ss.insertSheet(SH_PRODUCCION);
    prodLog.appendRow(PROD_HEADERS);
  } else {
    const h = prodLog.getRange(1, 1, 1, prodLog.getLastColumn()).getValues()[0].map(x => String(x).trim());
    if (h.indexOf('kg_total') === -1) {
      const rows = prodLog.getLastRow() >= 2
        ? prodLog.getRange(2, 1, prodLog.getLastRow() - 1, prodLog.getLastColumn()).getValues() : [];
      const iF = h.indexOf('fecha');
      const iB = h.indexOf('bolsas_producidas') >= 0 ? h.indexOf('bolsas_producidas') : h.indexOf('bolsas_25kg');
      const iO = h.indexOf('operario');
      const migrated = rows.map(r => {
        const bolsas = Number(r[iB]) || 0;
        return [iF >= 0 ? r[iF] : '', bolsas, 0, bolsas * 25, iO >= 0 ? r[iO] : ''];
      });
      prodLog.clear();
      prodLog.appendRow(PROD_HEADERS);
      if (migrated.length) prodLog.getRange(2, 1, migrated.length, PROD_HEADERS.length).setValues(migrated);
    }
  }

  // --- CLIENTES ---
  let cli = ss.getSheetByName(SH_CLIENTES);
  if (!cli) {
    cli = ss.insertSheet(SH_CLIENTES);
    cli.appendRow(['nombre', 'telefono', 'notas']);
  }

  // --- USUARIOS ---
  let usu = ss.getSheetByName(SH_USUARIOS);
  if (!usu) {
    usu = ss.insertSheet(SH_USUARIOS);
    usu.appendRow(['nombre', 'activo']);
  }
  // Sembrar usuarios default si faltan
  {
    const existing = (usu.getLastRow() >= 2)
      ? usu.getRange(2, 1, usu.getLastRow() - 1, 1).getValues().map(([n]) => String(n || '').trim().toLowerCase())
      : [];
    USUARIOS_SEED.forEach(name => {
      if (!existing.includes(name.toLowerCase())) {
        usu.appendRow([name, 'si']);
      }
    });
  }

  // --- CONFIG ---
  let cfg = ss.getSheetByName(SH_CONFIG);
  if (!cfg) {
    cfg = ss.insertSheet(SH_CONFIG);
    cfg.appendRow(['clave', 'valor']);
    cfg.appendRow(['capacidad_diaria', CAPACIDAD_DIARIA_DEFAULT]);
    cfg.appendRow(['capacidad_unidad', 'kg']);
  } else {
    // Garantizar header y fila capacidad_diaria
    const h = cfg.getRange(1, 1, 1, 2).getValues()[0].map(x => String(x).trim().toLowerCase());
    if (h[0] !== 'clave' || h[1] !== 'valor') {
      cfg.getRange(1, 1, 1, 2).setValues([['clave', 'valor']]);
    }
    let has = false;
    if (cfg.getLastRow() >= 2) {
      const data = cfg.getRange(2, 1, cfg.getLastRow() - 1, 2).getValues();
      has = data.some(r => String(r[0] || '').trim() === 'capacidad_diaria');
    }
    if (!has) {
      cfg.appendRow(['capacidad_diaria', CAPACIDAD_DIARIA_DEFAULT]);
      setConfigVal_(ss, 'capacidad_unidad', 'kg');
    }
  }

  // Migrar capacidad_diaria de bolsas → kg una sola vez
  if (getConfigVal_(ss, 'capacidad_unidad') !== 'kg') {
    const capViejo = Number(getConfigVal_(ss, 'capacidad_diaria')) || CAPACIDAD_DIARIA_DEFAULT;
    setConfigVal_(ss, 'capacidad_diaria', capViejo * 25);
    setConfigVal_(ss, 'capacidad_unidad', 'kg');
  }

  // --- ENTREGAS_PARCIALES ---
  let parc = ss.getSheetByName(SH_PARCIALES);
  if (!parc) {
    parc = ss.insertSheet(SH_PARCIALES);
    parc.appendRow(PARC_HEADERS);
    migrarPedidosAParciales_(ss);
  } else {
    const hdr = parc.getRange(1, 1, 1, parc.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (hdr.indexOf('cantidad_planeada') !== -1 && hdr.indexOf('bolsas_25kg') === -1) {
      // Migrar layout viejo (una sola cantidad) → multiproducto
      const rows = parc.getLastRow() >= 2
        ? parc.getRange(2, 1, parc.getLastRow() - 1, parc.getLastColumn()).getValues() : [];
      const idx = name => hdr.indexOf(name);
      const migrated = rows.map(r => {
        const o = {};
        PARC_HEADERS.forEach(h => { o[h] = ''; });
        o.id_parcial          = r[idx('id_parcial')];
        o.id_pedido           = r[idx('id_pedido')];
        o.fecha_planeada      = r[idx('fecha_planeada')];
        o.bolsas_25kg         = Number(r[idx('cantidad_planeada')]) || 0;
        o.bigbag_1000kg = o.emul_1l = o.emul_5l = o.emul_20l = 0;
        o.estado              = r[idx('estado')];
        o.fecha_entrega       = idx('fecha_entrega')      >= 0 ? r[idx('fecha_entrega')]      : '';
        o.observacion_entrega = idx('observacion_entrega')>= 0 ? r[idx('observacion_entrega')]: '';
        o.link_remito          = idx('link_remito')        >= 0 ? r[idx('link_remito')]        : '';
        o.usuario_entrega     = idx('usuario_entrega')    >= 0 ? r[idx('usuario_entrega')]    : '';
        return PARC_HEADERS.map(h => o[h]);
      });
      parc.clear();
      parc.appendRow(PARC_HEADERS);
      if (migrated.length) parc.getRange(2, 1, migrated.length, PARC_HEADERS.length).setValues(migrated);
    } else {
      parc.getRange(1, 1, 1, PARC_HEADERS.length).setValues([PARC_HEADERS]);
    }
    migrarPedidosAParciales_(ss);
  }

  // --- STOCK (con migración de layout viejo) ---
  let stock = ss.getSheetByName(SH_STOCK);
  if (!stock) {
    stock = ss.insertSheet(SH_STOCK);
    stock.getRange('A1:B5').setValues([
      ['Concepto',        'Kg'],
      ['Stock inicial',   0],
      ['Total producido', 0],
      ['Total vendido',   0],
      ['Stock actual',    0]
    ]);
    setConfigVal_(ss, 'stock_unidad', 'kg');
  } else {
    const a1 = String(stock.getRange('A1').getValue()).toLowerCase().trim();
    if (a1 === 'stock_inicial' || a1 === '') {
      stock.insertRowBefore(1);
    }
    stock.getRange('A2').setValue('Stock inicial');
    stock.getRange('A3').setValue('Total producido');
    stock.getRange('A4').setValue('Total vendido');
    stock.getRange('A5').setValue('Stock actual');
    // Migrar stock inicial de bolsas → kg una sola vez
    if (getConfigVal_(ss, 'stock_unidad') !== 'kg') {
      const ini = Number(stock.getRange('B2').getValue()) || 0;
      stock.getRange('B2').setValue(ini * 25);
      setConfigVal_(ss, 'stock_unidad', 'kg');
    }
    stock.getRange('A1:B1').setValues([['Concepto', 'Kg']]);
  }

  // --- PRODUCTO ---
  let prod = ss.getSheetByName(SH_PRODUCTO);
  const PROD_ROWS = [
    ['producto', 'unidad', 'costo', 'precio'],
    ['bolsas_25kg',   'bolsa (25kg)',     0, 0],
    ['bigbag_1000kg', 'big bag (1000kg)', 0, 0],
    ['emul_1l',       'botella 1L',       0, 0],
    ['emul_5l',       'botella 5L',       0, 0],
    ['emul_20l',      'botella 20L',      0, 0]
  ];
  if (!prod) {
    prod = ss.insertSheet(SH_PRODUCTO);
    prod.getRange(1, 1, PROD_ROWS.length, 4).setValues(PROD_ROWS);
  } else if (String(prod.getRange('A1').getValue()).trim().toLowerCase() !== 'producto') {
    // Layout viejo (costos de bolsa): preservar costo total (B6) y precio (B7) para bolsa
    const costoBolsa  = Number(prod.getRange('B6').getValue()) || 0;
    const precioBolsa = Number(prod.getRange('B7').getValue()) || 0;
    prod.clear();
    const rows = PROD_ROWS.map(r => r.slice());
    rows[1][2] = costoBolsa; rows[1][3] = precioBolsa;  // fila bolsas_25kg
    prod.getRange(1, 1, rows.length, 4).setValues(rows);
  }

  // --- GANANCIAS ---
  let gan = ss.getSheetByName(SH_GANANCIAS);
  const GAN_ROWS = [
    ['Concepto', 'Valor'],
    ['Material producido (kg)', 0],
    ['Material vendido (kg)',   0],
    ['Ingresos ($)',            0],
    ['Costo total ($)',         0],
    ['Ganancia bruta ($)',      0],
    ['Margen (%)',              0]
  ];
  if (!gan) { gan = ss.insertSheet(SH_GANANCIAS); }
  gan.getRange(1, 1, GAN_ROWS.length, 2).setValues(GAN_ROWS);

  reorderSheets_(ss, [SH_STOCK, SH_GANANCIAS, SH_PEDIDOS, SH_PARCIALES, SH_PRODUCCION, SH_USUARIOS, SH_CLIENTES, SH_PRODUCTO, SH_CONFIG]);

  applyClientesValidation_();
  applyUsuariosValidation_();
  applyFormatting_();
}

/**
 * Por cada pedido que NO tenga ninguna parcial en ENTREGAS_PARCIALES,
 * genera una parcial con la fecha_entrega_solicitada y la cantidad_bolsas.
 * Si el pedido está 'entregado', la parcial queda como entregada con la fecha_entrega.
 */
function migrarPedidosAParciales_(ss) {
  const shPed  = ss.getSheetByName(SH_PEDIDOS);
  const shParc = ss.getSheetByName(SH_PARCIALES);
  if (!shPed || !shParc) return;
  if (shPed.getLastRow() < 2) return;

  const pedHeaders = shPed.getRange(1, 1, 1, shPed.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const parcHeaders = shParc.getRange(1, 1, 1, shParc.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  const cP = name => pedHeaders.indexOf(name);
  const c_id        = cP('id');
  const c_cant      = cP('cantidad_bolsas');
  const c_fsolic    = cP('fecha_entrega_solicitada');
  const c_estado    = cP('estado');
  const c_fentrega  = cP('fecha_entrega');
  const c_obsE      = cP('observacion_entrega');
  const c_link      = cP('link_remito');

  // IDs de pedidos con parciales existentes
  const idsConParcial = new Set();
  let maxIdParc = 0;
  if (shParc.getLastRow() >= 2) {
    const pa = shParc.getRange(2, 1, shParc.getLastRow() - 1, shParc.getLastColumn()).getValues();
    const cIdParc = parcHeaders.indexOf('id_parcial');
    const cIdPed  = parcHeaders.indexOf('id_pedido');
    pa.forEach(r => {
      const ipa = Number(r[cIdParc]);
      if (!isNaN(ipa) && ipa > maxIdParc) maxIdParc = ipa;
      const ip = r[cIdPed];
      if (ip !== '' && ip !== null) idsConParcial.add(String(ip));
    });
  }

  const peds = shPed.getRange(2, 1, shPed.getLastRow() - 1, shPed.getLastColumn()).getValues();
  let added = 0;
  peds.forEach(r => {
    const idPed = r[c_id];
    if (idPed === '' || idPed === null) return;
    if (idsConParcial.has(String(idPed))) return;
    const cantidad = Number(r[c_cant]) || 0;
    if (cantidad <= 0) return;
    const fechaP   = r[c_fsolic] || r[c_fentrega] || new Date();
    const estPed   = String(r[c_estado] || '').toLowerCase();
    maxIdParc += 1;
    const newId = maxIdParc;
    if (estPed === 'entregado') {
      const fechaE = r[c_fentrega] || new Date();
      const obsE   = c_obsE >= 0 ? (r[c_obsE] || '') : '';
      const link   = c_link >= 0 ? (r[c_link] || '') : '';
      const row = parcHeaders.map(h => {
        if (h === 'id_parcial')         return newId;
        if (h === 'id_pedido')          return idPed;
        if (h === 'fecha_planeada')     return fechaP;
        if (h === 'bolsas_25kg')        return cantidad;
        if (h === 'estado')             return 'entregada';
        if (h === 'fecha_entrega')      return fechaE;
        if (h === 'observacion_entrega')return obsE;
        if (h === 'link_remito')        return link;
        return '';
      });
      shParc.appendRow(row);
    } else {
      const row = parcHeaders.map(h => {
        if (h === 'id_parcial')        return newId;
        if (h === 'id_pedido')         return idPed;
        if (h === 'fecha_planeada')    return fechaP;
        if (h === 'bolsas_25kg')       return cantidad;
        if (h === 'estado')            return 'pendiente';
        return '';
      });
      shParc.appendRow(row);
    }
    added += 1;
  });
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
  const headers = pedidos.getRange(1, 1, 1, pedidos.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const colCliente = headers.indexOf('cliente') + 1;
  if (colCliente < 1) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(clientes.getRange('A2:A'), true)
    .setAllowInvalid(true)
    .build();
  const maxRows = Math.max(pedidos.getMaxRows() - 1, 1000);
  pedidos.getRange(2, colCliente, maxRows, 1).setDataValidation(rule);
}

function applyUsuariosValidation_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const pedidos = ss.getSheetByName(SH_PEDIDOS);
  const usuarios = ss.getSheetByName(SH_USUARIOS);
  if (!pedidos || !usuarios) return;
  const headers = pedidos.getRange(1, 1, 1, pedidos.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const colUsu = headers.indexOf('usuario_carga') + 1;
  if (colUsu < 1) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(usuarios.getRange('A2:A'), true)
    .setAllowInvalid(true)
    .build();
  const maxRows = Math.max(pedidos.getMaxRows() - 1, 1000);
  pedidos.getRange(2, colUsu, maxRows, 1).setDataValidation(rule);
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

    const widths = headers.map(h => ({
      'id': 60, 'fecha_carga': 150, 'usuario_carga': 150, 'cliente': 200, 'cantidad_bolsas': 130,
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
    if (colFSolic > 0) {
      ped.getRange(2, colFSolic, ped.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    }
    if (colEstado > 0) {
      const estadoRange = ped.getRange(2, colEstado, ped.getMaxRows() - 1, 1);
      const rulePend = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('pendiente')
        .setBackground(COLOR_PEND_BG).setFontColor(COLOR_PEND_TXT).setBold(true)
        .setRanges([estadoRange]).build();
      const ruleParc = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('parcial')
        .setBackground(COLOR_PARC_BG).setFontColor(COLOR_PARC_TXT).setBold(true)
        .setRanges([estadoRange]).build();
      const ruleEntr = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('entregado')
        .setBackground(COLOR_ENTR_BG).setFontColor(COLOR_ENTR_TXT).setBold(true)
        .setRanges([estadoRange]).build();
      ped.setConditionalFormatRules([rulePend, ruleParc, ruleEntr]);
      ped.getRange(2, colEstado, ped.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    }
    if (colId   > 0) ped.getRange(2, colId,   ped.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    if (colCant > 0) ped.getRange(2, colCant, ped.getMaxRows() - 1, 1).setHorizontalAlignment('right');
  }

  // ENTREGAS_PARCIALES
  // PARC_HEADERS: 1 id_parcial, 2 id_pedido, 3 fecha_planeada, 4-8 productos (bolsas_25kg..emul_20l),
  // 9 estado, 10 fecha_entrega, 11 observacion_entrega, 12 link_remito, 13 usuario_entrega
  const parc = ss.getSheetByName(SH_PARCIALES);
  if (parc) {
    const widths = [80, 80, 130, 110, 120, 90, 90, 90, 110, 130, 200, 200, 150];
    formatList_(parc, {
      cols: PARC_HEADERS.length, widths,
      headerHeight: 36, rowHeight: 26,
      numberCols: [4, 5, 6, 7, 8],
      dateCols: [3, 10]
    });
    parc.getRange(2, 3, parc.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    parc.getRange(2, 10, parc.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy HH:mm');
    parc.getRange(2, 1, parc.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    parc.getRange(2, 2, parc.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    parc.getRange(2, 9, parc.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    const estR = parc.getRange(2, 9, parc.getMaxRows() - 1, 1);
    const rulePend = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('pendiente')
      .setBackground(COLOR_PEND_BG).setFontColor(COLOR_PEND_TXT).setBold(true)
      .setRanges([estR]).build();
    const ruleEntr = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('entregada')
      .setBackground(COLOR_ENTR_BG).setFontColor(COLOR_ENTR_TXT).setBold(true)
      .setRanges([estR]).build();
    parc.setConditionalFormatRules([rulePend, ruleEntr]);
  }

  // PRODUCCION
  const prodLog = ss.getSheetByName(SH_PRODUCCION);
  if (prodLog) {
    formatList_(prodLog, {
      cols: 5,
      widths: [180, 130, 130, 130, 220],
      headerHeight: 36, rowHeight: 28,
      numberCols: [2, 3, 4], dateCols: [1]
    });
    prodLog.getRange(2, 2, prodLog.getMaxRows() - 1, 3).setHorizontalAlignment('right');
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

  // USUARIOS
  const usu = ss.getSheetByName(SH_USUARIOS);
  if (usu) {
    formatList_(usu, {
      cols: 2,
      widths: [240, 120],
      headerHeight: 36, rowHeight: 28
    });
    usu.getRange(2, 2, usu.getMaxRows() - 1, 1).setHorizontalAlignment('center');
  }

  // CONFIG
  const cfg = ss.getSheetByName(SH_CONFIG);
  if (cfg) {
    formatList_(cfg, {
      cols: 2,
      widths: [220, 180],
      headerHeight: 36, rowHeight: 28
    });
    cfg.getRange(2, 2, cfg.getMaxRows() - 1, 1).setHorizontalAlignment('right');
  }

  // STOCK
  const stock = ss.getSheetByName(SH_STOCK);
  if (stock) {
    formatDashboard_(stock, {
      rows: 5,
      valueFormat: '#,##0" kg"',
      highlightRow: 5,
      colWidths: [260, 220]
    });
  }

  // PRODUCTO — ahora es tabla (1 fila por producto), no dashboard
  const prod = ss.getSheetByName(SH_PRODUCTO);
  if (prod) {
    formatList_(prod, {
      cols: 4,
      widths: [160, 200, 130, 130],
      headerHeight: 36, rowHeight: 30,
      numberCols: [3, 4]
    });
  }

  // GANANCIAS
  const gan = ss.getSheetByName(SH_GANANCIAS);
  if (gan) {
    formatDashboard_(gan, {
      rows: 7,
      valueFormat: '"$" #,##0',
      highlightRow: 6,
      colWidths: [260, 240],
      customFormats: { 2: '#,##0" kg"', 3: '#,##0" kg"', 7: '0.0"%"' }
    });
  }

  [SH_PEDIDOS, SH_PARCIALES, SH_PRODUCCION, SH_CLIENTES, SH_USUARIOS, SH_CONFIG, SH_STOCK, SH_PRODUCTO, SH_GANANCIAS].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s) s.setHiddenGridlines(true);
  });
}

function formatList_(sh, opt) {
  const cols = opt.cols;
  const hdr = sh.getRange(1, 1, 1, cols);
  hdr.setBackground(COLOR_BRAND)
     .setFontColor(COLOR_HEADER_TXT)
     .setFontWeight('bold')
     .setFontSize(11)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
  sh.setRowHeight(1, opt.headerHeight || 32);
  sh.setFrozenRows(1);
  if (opt.widths) opt.widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  if (opt.rowHeight) {
    const last = sh.getMaxRows();
    if (last > 1) {
      try { sh.setRowHeights(2, last - 1, opt.rowHeight); } catch (_) {}
    }
  }
  sh.getBandings().forEach(b => b.remove());
  const bandRange = sh.getRange(1, 1, sh.getMaxRows(), cols);
  const banding = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banding.setHeaderRowColor(COLOR_BRAND).setFirstRowColor('#ffffff').setSecondRowColor('#f9fafb');
  if (opt.numberCols) {
    opt.numberCols.forEach(c => sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0'));
  }
  if (opt.dateCols) {
    opt.dateCols.forEach(c => sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy HH:mm'));
  }
  sh.getRange(2, 1, sh.getMaxRows() - 1, cols).setVerticalAlignment('middle');
}

function formatDashboard_(sh, opt) {
  const hdr = sh.getRange(1, 1, 1, 2);
  hdr.setBackground(COLOR_BRAND)
     .setFontColor(COLOR_HEADER_TXT)
     .setFontWeight('bold')
     .setFontSize(12)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
  sh.setRowHeight(1, 36);
  sh.setFrozenRows(1);
  const w = opt.colWidths || [240, 200];
  sh.setColumnWidth(1, w[0]);
  sh.setColumnWidth(2, w[1]);
  const labels = sh.getRange(2, 1, opt.rows - 1, 1);
  labels.setBackground(COLOR_LABEL_BG)
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle')
        .setFontSize(11);
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
  for (let r = 2; r <= opt.rows; r++) sh.setRowHeight(r, 32);
  if (opt.highlightRow) {
    const r = sh.getRange(opt.highlightRow, 1, 1, 2);
    r.setBackground(COLOR_HILITE_BG)
     .setFontColor(COLOR_HILITE_TXT)
     .setFontWeight('bold')
     .setFontSize(14);
    sh.setRowHeight(opt.highlightRow, 42);
  }
  sh.getRange(1, 1, opt.rows, 2)
    .setBorder(true, true, true, true, true, true, COLOR_BORDER, SpreadsheetApp.BorderStyle.SOLID);
}
