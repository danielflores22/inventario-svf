/**
 * Sincronización de la app "Toma de inventario" con Google Sheets.
 * Pega este código en Extensiones > Apps Script de tu hoja de Google,
 * cambia la CLAVE y publícalo como aplicación web (ver instrucciones).
 */
const CLAVE = 'CAMBIA-ESTA-CLAVE';           // la misma clave se escribe en la app
const CARPETA_FOTOS = 'Inventario - fotos no listados';

const H_CONTEOS = ['id', 'k', 'ubicacion', 'codigo', 'descripcion', 'serie', 'conteo', 'recuento', 'serie_leida', 'observacion', 'pareja', 'ts'];
const H_NOLIST = ['id', 'ubicacion', 'descripcion', 'codigo_caja', 'serie', 'cantidad', 'foto', 'pareja', 'ts', 'vinculado_a'];
const H_UBIC = ['ubicacion', 'estado', 'pareja', 'cerrada', 'reabierta', 't'];

function hoja_(nombre, cab) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, sh.getMaxRows(), cab.length).setNumberFormat('@'); // todo como texto: evita que Sheets cambie fechas o códigos
    sh.appendRow(cab); sh.setFrozenRows(1);
  }
  return sh;
}
function leer_(sh) {
  const v = sh.getDataRange().getValues();
  const cab = v.shift();
  return { cab: cab, filas: v };
}
function respuesta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function carpeta_() {
  const it = DriveApp.getFoldersByName(CARPETA_FOTOS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_FOTOS);
}
function archivoPaquete_() {
  const it = carpeta_().getFilesByName('paquete.json');
  return it.hasNext() ? it.next() : null;
}

/* Inserta o actualiza filas por la primera columna; gana el registro con ts/t más reciente */
function upsert_(sh, cab, registros, colTs) {
  if (!registros.length) return 0;
  const d = leer_(sh);
  const idx = {};
  d.filas.forEach((f, i) => idx[String(f[0])] = i);
  const its = cab.indexOf(colTs);
  let n = 0;
  const nuevas = [];
  registros.forEach(r => {
    const fila = cab.map(c => r[c] === undefined || r[c] === null ? '' : r[c]);
    const i = idx[String(fila[0])];
    if (i === undefined) { nuevas.push(fila); idx[String(fila[0])] = -1; n++; }
    else if (i >= 0 && String(fila[its]) >= String(d.filas[i][its])) { sh.getRange(i + 2, 1, 1, cab.length).setNumberFormat('@').setValues([fila.map(v => String(v))]); n++; }
  });
  if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, cab.length).setNumberFormat('@').setValues(nuevas.map(f => f.map(v => String(v))));
  return n;
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.clave !== CLAVE) return respuesta_({ ok: false, error: 'Clave incorrecta' });
  if (p.accion === 'paquete') {
    const f = archivoPaquete_();
    return respuesta_({ ok: true, paquete: f ? JSON.parse(f.getBlob().getDataAsString()) : null });
  }
  if (p.accion === 'ubicaciones') {
    const out = { ok: true, locs: {} };
    const u = leer_(hoja_('Ubicaciones', H_UBIC));
    u.filas.forEach(f => {
      const r = {}; u.cab.forEach((h, i) => r[h] = f[i]);
      out.locs[r.ubicacion] = { status: String(r.estado), pair: String(r.pareja), closedAt: String(r.cerrada) || null, reopenedAt: String(r.reabierta) || null };
    });
    return respuesta_(out);
  }
  if (p.accion === 'traer') {
    const since = p.desde || '';
    const out = { ok: true, ahora: new Date().toISOString(), counts: {}, extras: [], locs: {} };
    const c = leer_(hoja_('Conteos', H_CONTEOS));
    c.filas.forEach(f => {
      const r = {}; c.cab.forEach((h, i) => r[h] = f[i]);
      if (String(r.ts) <= since) return;
      out.counts[r.id] = { k: r.k, q1: r.conteo === '' ? null : Number(r.conteo), q2: r.recuento === '' ? null : Number(r.recuento), lotRead: String(r.serie_leida || ''), obs: String(r.observacion || ''), pair: String(r.pareja || ''), ts: String(r.ts) };
    });
    const x = leer_(hoja_('NoListados', H_NOLIST));
    x.filas.forEach(f => {
      const r = {}; x.cab.forEach((h, i) => r[h] = f[i]);
      if (String(r.ts) <= since) return;
      out.extras.push({ id: String(r.id), loc: String(r.ubicacion), name: String(r.descripcion), code: String(r.codigo_caja), lot: String(r.serie), qty: Number(r.cantidad), photo: String(r.foto || '') || null, pair: String(r.pareja), ts: String(r.ts) });
    });
    const u = leer_(hoja_('Ubicaciones', H_UBIC));
    u.filas.forEach(f => {
      const r = {}; u.cab.forEach((h, i) => r[h] = f[i]);
      out.locs[r.ubicacion] = { status: String(r.estado), pair: String(r.pareja), closedAt: String(r.cerrada) || null, reopenedAt: String(r.reabierta) || null };
    });
    return respuesta_(out);
  }
  return respuesta_({ ok: true, mensaje: 'Conexión correcta' });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return respuesta_({ ok: false, error: 'Datos no válidos' }); }
  if (d.clave !== CLAVE) return respuesta_({ ok: false, error: 'Clave incorrecta' });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (d.accion === 'paquete') {
      const txt = JSON.stringify(d.paquete);
      const f = archivoPaquete_();
      if (f) f.setContent(txt); else carpeta_().createFile('paquete.json', txt, 'application/json');
      return respuesta_({ ok: true, lineas: d.paquete.lines.length });
    }
    if (d.accion === 'subir') {
      const conteos = Object.keys(d.counts || {}).map(id => {
        const c = d.counts[id], l = c.l || {};
        return { id: id, k: c.k, ubicacion: l.loc, codigo: l.code, descripcion: l.name, serie: l.lot, conteo: c.q1, recuento: c.q2, serie_leida: c.lotRead, observacion: c.obs, pareja: c.pair || d.pair, ts: c.ts };
      });
      const nc = upsert_(hoja_('Conteos', H_CONTEOS), H_CONTEOS, conteos, 'ts');
      const fotos = {};
      const extras = (d.extras || []).map(x => {
        let foto = x.photoUrl || '';
        if (x.photo && String(x.photo).indexOf('data:image') === 0) {
          const b = Utilities.newBlob(Utilities.base64Decode(x.photo.split(',')[1]), 'image/jpeg', x.loc.replace(/\W+/g, '_') + '_' + x.id + '.jpg');
          const file = carpeta_().createFile(b);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          foto = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
          fotos[x.id] = foto;
        }
        return { id: x.id, ubicacion: x.loc, descripcion: x.name, codigo_caja: x.code, serie: x.lot, cantidad: x.qty, foto: foto, pareja: x.pair || d.pair, ts: x.ts, vinculado_a: x.match ? (x.match.code + ' ' + x.match.name) : '' };
      });
      const nx = upsert_(hoja_('NoListados', H_NOLIST), H_NOLIST, extras, 'ts');
      const locs = Object.keys(d.locs || {}).map(l => {
        const v = d.locs[l];
        const t = [v.closedAt || '', v.reopenedAt || '', v.t || ''].sort().pop();
        return { ubicacion: l, estado: v.status, pareja: v.pair || '', cerrada: v.closedAt || '', reabierta: v.reopenedAt || '', t: t };
      }).filter(r => r.t);
      const nl = upsert_(hoja_('Ubicaciones', H_UBIC), H_UBIC, locs, 't');
      return respuesta_({ ok: true, conteos: nc, nolistados: nx, ubicaciones: nl, fotos: fotos });
    }
    return respuesta_({ ok: false, error: 'Acción desconocida' });
  } finally {
    lock.releaseLock();
  }
}
