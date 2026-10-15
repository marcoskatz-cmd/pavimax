/**
 * IngecoAvisos.gs — pegar tal cual en cualquier proyecto de Apps Script de INGECO.
 * Requiere la propiedad del script CLAVE_<MODULO> (la entrega Marcos). Nunca lanza: si falla, loguea y sigue.
 *
 * notificarAppIngeco_('PAVIMAX', {
 *   destinatarios: [{ con_modulo: 'PAVIMAX_OP' }, { admin: true }],   // legajos, {area}, {obra}, {sector}, {email}, {emails:[]}, {con_modulo}, {admin:true}, {todos:true}
 *   tipo: 'aviso_general', prioridad: 'normal',                        // 'critica' = push insistente
 *   titulo: 'Pedido nuevo #12 de Vialidad', cuerpo: '20 big bags para el martes',
 *   url_destino: 'https://...',                                         // opcional, abre el ítem
 *   clave_unica: 'pedido-12', dias_unica: 30                            // opcional: no repetir el mismo aviso
 * });
 */
var APP_INGECO_BACKEND_ = 'https://script.google.com/macros/s/AKfycbx9NN1FfOQjUjfymLXrMyd2QM6l96HP0lxY6RC7_xKsbDlO0PnwdHbaMtag54VYFor0/exec';

function notificarAppIngeco_(modulo, aviso) {
  try {
    var clave = PropertiesService.getScriptProperties().getProperty('CLAVE_' + String(modulo).toUpperCase());
    if (!clave) { console.warn('notificarAppIngeco_: falta CLAVE_' + modulo); return { ok: false, error: 'sin clave' }; }
    var body = { accion: 'notificar', clave_sistema: clave, modulo: String(modulo).toUpperCase() };
    for (var k in aviso) body[k] = aviso[k];
    var r = UrlFetchApp.fetch(APP_INGECO_BACKEND_, { method: 'post', contentType: 'text/plain;charset=utf-8', payload: JSON.stringify(body), muteHttpExceptions: true });
    var j = {}; try { j = JSON.parse(r.getContentText()); } catch (e) { }
    if (!j.ok) console.warn('notificarAppIngeco_: ' + r.getContentText().slice(0, 200));
    return j;
  } catch (e) { console.warn('notificarAppIngeco_: ' + e); return { ok: false, error: String(e) }; }
}
