/**
 * API de Santino — Google Apps Script
 * Sirve para "Mi dinero" y después para "Mi día" (cada petición trae el campo "app").
 *
 * Qué hace:
 *  - Guarda los datos en un Google Sheet "<App> - datos" (pestañas: movimientos, suscripciones, sugeridas).
 *  - Escribe "<app>-respaldo.json" en la carpeta de Drive "Datos apps" (ahí lo leen tus agentes).
 *  - Guarda una copia por día en "Datos apps/respaldos".
 *  - Una vez al día busca en tu Gmail cobros de suscripciones y los deja como "sugeridas".
 *  - Mi día: pone tus empleados y tu rutina fija en el calendario "Mi día" de Google Calendar.
 *
 * Primera vez: elige la función "instalar" arriba y toca ▶ Ejecutar. En "Registro de ejecución" sale tu clave.
 */

const CARPETA = "Datos apps";
const APPS = {
  "mi-dinero": { hoja: "Mi dinero - datos", archivo: "mi-dinero-respaldo.json" },
  "mi-dia":    { hoja: "Mi día - datos",    archivo: "mi-dia-respaldo.json" }
};
const BUSQUEDA_GMAIL = 'newer_than:60d (recibo OR receipt OR suscripción OR suscripcion OR renovación OR renovacion OR trial OR "your plan" OR factura OR invoice)';
const COLS_SUG = ["id", "servicio", "monto", "moneda", "frecuencia", "fecha", "prueba", "asunto", "de", "estado"];

/* ============ Instalación (correr 1 vez desde el editor) ============ */
function instalar() {
  const p = PropertiesService.getScriptProperties();
  let clave = p.getProperty("CLAVE");
  if (!clave) {
    clave = Utilities.getUuid().replace(/-/g, "").slice(0, 24);
    p.setProperty("CLAVE", clave);
  }
  carpeta_();
  hoja_("mi-dinero");
  // Activador: 1 vez al día, 7 am
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === "buscarSuscripciones"; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("buscarSuscripciones").timeBased().everyDays(1).atHour(7).create();
  buscarSuscripciones();
  Logger.log("✅ Listo. Tu clave secreta es:  " + clave);
  Logger.log("Cópiala y pégala en la app (Ajustes → Nube). No la compartas.");
}

/** Por si un día quieres cambiar la clave: corre esta función y usa la nueva en la app. */
function nuevaClave() {
  PropertiesService.getScriptProperties().deleteProperty("CLAVE");
  instalar();
}

/* ============ API ============ */
function doGet() {
  return out_({ ok: true, mensaje: "La API de Santino está funcionando." });
}

function doPost(e) {
  try {
    const q = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const clave = PropertiesService.getScriptProperties().getProperty("CLAVE");
    if (!clave || q.clave !== clave) return out_({ ok: false, error: "Clave incorrecta" });
    const app = String(q.app || "").toLowerCase().replace(/[^a-z0-9-]/g, "");
    if (!app) return out_({ ok: false, error: "Falta el campo app" });

    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      if (q.accion === "leer") {
        const r = leerJson_(app);
        return out_({ ok: true, estado: r ? r.datos : null, actualizado: r ? r.actualizado : null });
      }
      if (q.accion === "guardar") {
        const err = guardar_(app, q);
        if (err) return out_({ ok: false, error: err });
      } else if (q.accion === "buscar") {
        // Búsqueda manual en Gmail (ya tenemos el candado, por eso va la versión sin candado)
        if (app === "mi-dinero") buscarSuscripciones_();
      } else if (q.accion !== "sugeridas") {
        return out_({ ok: false, error: "Acción desconocida: " + q.accion });
      }
      const res = { ok: true };
      if (app === "mi-dia") res.calendario = PropertiesService.getScriptProperties().getProperty("CAL_ESTADO") || "";
      if (app === "mi-dinero") {
        decisiones_(q.decisiones);
        res.sugeridas = sugeridas_();
      }
      return out_(res);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ============ Guardar ============ */
function guardar_(app, q) {
  // Protección: un celular vacío (app recién instalada) no borra lo que ya hay en la nube.
  const movsNuevos = q.estado && Array.isArray(q.estado.movs) ? q.estado.movs.length : -1;
  // Mi día avisa con "vacio": true cuando el celular no tiene nada anotado
  const vacio = app === "mi-dinero" ? movsNuevos === 0 : q.vacio === true;
  if (vacio && !q.forzar) {
    const prev = leerJson_(app);
    const d = prev && prev.datos;
    const habia = d && (app === "mi-dinero"
      ? Array.isArray(d.movs) && d.movs.length > 0
      : (d.log && Object.keys(d.log).length > 1) || (Array.isArray(d.events) && d.events.length > 0));
    if (habia) {
      return "La nube tiene datos y este celular está vacío. Toca \"Restaurar desde la nube\".";
    }
  }
  const ss = SpreadsheetApp.openById(hoja_(app));
  const tablas = q.tablas || {};
  const reemplazar = q.reemplazar || [];
  Object.keys(tablas).forEach(function (nombre) {
    const filas = (tablas[nombre] || []).filter(function (r) { return r && r.id; });
    const sh = pestana_(ss, nombre);
    if (reemplazar.indexOf(nombre) >= 0) reemplazarTabla_(sh, filas);
    else if (filas.length) upsert_(sh, filas);
  });
  const borrados = q.borrados || {};
  Object.keys(borrados).forEach(function (nombre) {
    const ids = borrados[nombre] || [];
    if (ids.length) borrar_(pestana_(ss, nombre), ids);
  });
  if (q.estado) escribirJson_(app, q);
  // Mi día: empleados y rutina fija → calendario "Mi día" de Google (aviso 5 min antes)
  if (app === "mi-dia" && q.estado) {
    try { sincronizarCalendario_(q.estado); PropertiesService.getScriptProperties().setProperty("CAL_ESTADO", "ok"); }
    catch (err) { PropertiesService.getScriptProperties().setProperty("CAL_ESTADO", "Error: " + String(err && err.message || err)); }
  }
  return null;
}

/* ============ Google Calendar (Mi día) ============ */
const CAL_NOMBRE = "Mi día";
const DIAS_CAL = ["SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"];

/** Córrela 1 vez desde el editor para dar permiso de Calendar y sincronizar con lo último que subió la app. */
function sincronizarCalendario() {
  const r = leerJson_("mi-dia");
  if (!r || !r.datos) { Logger.log("Todavía no hay datos de Mi día en la nube. Abre la app con la nube conectada."); return; }
  sincronizarCalendario_(r.datos);
  PropertiesService.getScriptProperties().setProperty("CAL_ESTADO", "ok");
  Logger.log("✅ Listo. Revisa el calendario \"" + CAL_NOMBRE + "\" en Google Calendar.");
}

/** Crea solo lo que cambió: compara cada evento con lo que ya subió (huella) y reemplaza esa serie. */
function sincronizarCalendario_(estado) {
  const tz = Session.getScriptTimeZone();
  const p = PropertiesService.getScriptProperties();
  const antes = JSON.parse(p.getProperty("CAL_MAPA") || "{}");
  const items = itemsCalendario_(estado, tz);
  const ahora = {};
  let cal = null;
  const C = function () { return cal || (cal = calendarioMiDia_(tz)); };
  items.forEach(function (it) {
    const huella = huella_(it);
    const viejo = antes[it.key];
    if (viejo && viejo.h === huella) { ahora[it.key] = viejo; return; }
    if (viejo) borrarSerie_(C(), viejo.id);
    ahora[it.key] = { id: crearSerie_(C(), it, tz).getId(), h: huella };
  });
  Object.keys(antes).forEach(function (k) { if (!ahora[k]) borrarSerie_(C(), antes[k].id); });
  // Primera vez: quita los que se pusieron a mano en tu calendario principal con el mismo nombre
  if (!p.getProperty("CAL_LIMPIO")) { quitarCopias_(items); p.setProperty("CAL_LIMPIO", "1"); }
  p.setProperty("CAL_MAPA", JSON.stringify(ahora));
}

function itemsCalendario_(estado, tz) {
  const hoy = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const out = [];
  (estado.staff || []).forEach(function (x) {
    if (!x.days || !x.days.length || !x.t) return;
    const hm = String(x.t).split(":"), s = (+hm[0]) * 60 + (+hm[1] || 0);
    out.push({
      key: "staff-" + x.id, titulo: x.em + " " + x.n + " · " + x.what, dias: x.days.slice().sort(),
      s: s, e: Math.min(s + 15, 1439), fecha: primerDia_(x.days, hoy, tz), ex: [],
      desc: "Pídele: " + x.p + "\n\nAbrir Claude: https://claude.ai/new?q=" + encodeURIComponent(x.p)
    });
  });
  (estado.events || []).forEach(function (ev) {
    if (!ev.rep || !ev.rep.length) return;
    const desde = ev.from && ev.from > hoy ? ev.from : hoy, f = primerDia_(ev.rep, desde, tz);
    out.push({
      key: "ev-" + ev.id, titulo: ev.n, dias: ev.rep.slice().sort(), s: ev.s, e: ev.e, fecha: f,
      ex: Object.keys(ev.ex || {}).filter(function (d) { return d >= f; }).sort(), desc: "De tu semana en Mi día"
    });
  });
  return out;
}

function primerDia_(dias, desde, tz) {
  const base = Utilities.parseDate(desde + " 12:00", tz, "yyyy-MM-dd HH:mm").getTime();
  for (let i = 0; i < 7; i++) {
    const d = new Date(base + i * 864e5);
    if (dias.indexOf(Number(Utilities.formatDate(d, tz, "u")) % 7) >= 0) return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }
  return desde;
}

const hhmm_ = function (m) { return ("0" + Math.floor(m / 60)).slice(-2) + ":" + ("0" + (m % 60)).slice(-2); };

function huella_(it) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([it.titulo, it.dias, it.s, it.e, it.ex, it.desc]), Utilities.Charset.UTF_8);
  return b.map(function (x) { return ((x + 256) % 256).toString(16); }).join("");
}

function calendarioMiDia_(tz) {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty("CAL_ID");
  if (id) { const c = CalendarApp.getCalendarById(id); if (c) return c; }
  const c = CalendarApp.createCalendar(CAL_NOMBRE, { timeZone: tz, summary: "Tus empleados y tu rutina fija. Lo actualiza la app Mi día." });
  p.setProperty("CAL_ID", c.getId());
  p.deleteProperty("CAL_MAPA");
  return c;
}

function crearSerie_(cal, it, tz) {
  const f = "yyyy-MM-dd HH:mm";
  const ini = Utilities.parseDate(it.fecha + " " + hhmm_(it.s), tz, f);
  const fin = Utilities.parseDate(it.fecha + " " + hhmm_(it.e), tz, f);
  const rec = CalendarApp.newRecurrence();
  rec.addWeeklyRule().onlyOnWeekdays(it.dias.map(function (d) { return CalendarApp.Weekday[DIAS_CAL[d]]; }));
  it.ex.forEach(function (d) { rec.addDateExclusion(Utilities.parseDate(d + " " + hhmm_(it.s), tz, f)); });
  const serie = cal.createEventSeries(it.titulo, ini, fin, rec, { description: it.desc });
  serie.removeAllReminders();
  serie.addPopupReminder(5);
  return serie;
}

function borrarSerie_(cal, id) {
  try { const s = cal.getEventSeriesById(id); if (s) s.deleteEventSeries(); } catch (e) { /* ya no existe */ }
}

function quitarCopias_(items) {
  const titulos = {};
  items.forEach(function (it) { titulos[it.titulo] = true; });
  const cal = CalendarApp.getDefaultCalendar(), hechas = {};
  cal.getEvents(new Date(), new Date(Date.now() + 8 * 864e5)).forEach(function (ev) {
    if (!titulos[ev.getTitle()] || !ev.isRecurringEvent()) return;
    try {
      const s = ev.getEventSeries(), id = s.getId();
      if (!hechas[id]) { hechas[id] = true; s.deleteEventSeries(); }
    } catch (e) { /* nada */ }
  });
}

function pestana_(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}

/** Encabezados actuales + los nuevos que traigan las filas. "id" siempre en la columna A. */
function encabezados_(sh, filas) {
  let h = [];
  if (sh.getLastRow() > 0 && sh.getLastColumn() > 0) {
    h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String).filter(function (x) { return x; });
  }
  if (!h.length) h = ["id"];
  filas.forEach(function (r) {
    Object.keys(r).forEach(function (k) { if (h.indexOf(k) < 0) h.push(k); });
  });
  sh.getRange(1, 1, 1, h.length).setValues([h]);
  sh.setFrozenRows(1);
  return h;
}

function fila_(h, r) {
  return h.map(function (k) { return r[k] === undefined || r[k] === null ? "" : r[k]; });
}

/** Inserta o reemplaza por id: nunca duplica. */
function upsert_(sh, filas) {
  const h = encabezados_(sh, filas);
  const n = sh.getLastRow() - 1;
  const ids = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const pos = {};
  ids.forEach(function (id, i) { pos[id] = i + 2; });
  const nuevas = [];
  filas.forEach(function (r) {
    const v = fila_(h, r);
    const id = String(r.id);
    if (pos[id]) sh.getRange(pos[id], 1, 1, h.length).setValues([v]);
    else { nuevas.push(v); pos[id] = -1; }
  });
  if (nuevas.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, nuevas.length, 1).setNumberFormat("@");
    sh.getRange(start, 1, nuevas.length, h.length).setValues(nuevas);
  }
}

function reemplazarTabla_(sh, filas) {
  sh.clearContents();
  const h = encabezados_(sh, filas);
  if (filas.length) {
    sh.getRange(2, 1, filas.length, 1).setNumberFormat("@");
    sh.getRange(2, 1, filas.length, h.length).setValues(filas.map(function (r) { return fila_(h, r); }));
  }
}

function borrar_(sh, ids) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return;
  const set = {};
  ids.forEach(function (id) { set[String(id)] = true; });
  const col = sh.getRange(2, 1, n, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) {
    if (set[String(col[i][0])]) sh.deleteRow(i + 2);
  }
}

/* ============ Drive: archivo para tus agentes ============ */
function carpeta_() {
  const it = DriveApp.getFoldersByName(CARPETA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA);
}

function subcarpeta_(padre, nombre) {
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

function nombreArchivo_(app) {
  return APPS[app] ? APPS[app].archivo : app + "-respaldo.json";
}

function escribir_(carpeta, nombre, texto) {
  const it = carpeta.getFilesByName(nombre);
  if (it.hasNext()) it.next().setContent(texto);
  else carpeta.createFile(nombre, texto, "application/json");
}

function escribirJson_(app, q) {
  const tz = Session.getScriptTimeZone();
  const doc = {
    app: app,
    actualizado: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ssXXX"),
    explicacion: app === "mi-dia"
      ? "resumen = lo listo para leer (hoy, semana, empleados, horarios libres de la semana que viene, última revisión del domingo). datos = todo lo de la app (log = cada día: done, sec, mood, grat, sleep; events = semana; staff = empleados)."
      : "resumen = números listos para leer. datos = todo lo que hay en la app (movs = movimientos; settings.membresias = suscripciones).",
    resumen: q.resumen || null,
    datos: q.estado
  };
  const texto = JSON.stringify(doc, null, 1);
  const c = carpeta_();
  escribir_(c, nombreArchivo_(app), texto);
  // Copia del día (si algo sale mal, se puede volver a una fecha anterior)
  const hoy = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  escribir_(subcarpeta_(c, "respaldos"), app + "-" + hoy + ".json", texto);
}

function leerJson_(app) {
  const it = carpeta_().getFilesByName(nombreArchivo_(app));
  if (!it.hasNext()) return null;
  try { return JSON.parse(it.next().getBlob().getDataAsString("UTF-8")); } catch (e) { return null; }
}

/* ============ Google Sheet ============ */
function hoja_(app) {
  const p = PropertiesService.getScriptProperties();
  const key = "HOJA_" + app;
  const id = p.getProperty(key);
  if (id) {
    try { SpreadsheetApp.openById(id); return id; } catch (e) { /* la borraron: se crea otra */ }
  }
  const nombre = APPS[app] ? APPS[app].hoja : app + " - datos";
  const ss = SpreadsheetApp.create(nombre);
  const hojas = ss.getSheets();
  hojas[0].setName(app === "mi-dinero" ? "movimientos" : "registros");
  if (app === "mi-dinero") { ss.insertSheet("suscripciones"); ss.insertSheet("sugeridas"); }
  try { DriveApp.getFileById(ss.getId()).moveTo(carpeta_()); } catch (e) { /* se queda en Mi unidad */ }
  p.setProperty(key, ss.getId());
  return ss.getId();
}

/* ============ Gmail → suscripciones sugeridas ============ */
/** Corre sola 1 vez al día (activador). También puedes correrla a mano. Usa el mismo candado que la API. */
function buscarSuscripciones() {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { buscarSuscripciones_(); } finally { lock.releaseLock(); }
}

/** La búsqueda en sí (sin candado: quien la llame ya lo tiene). */
function buscarSuscripciones_() {
  const ss = SpreadsheetApp.openById(hoja_("mi-dinero"));
  const sh = pestana_(ss, "sugeridas");
  if (sh.getLastRow() === 0) { sh.getRange(1, 1, 1, COLS_SUG.length).setValues([COLS_SUG]); sh.setFrozenRows(1); }
  const n = sh.getLastRow() - 1;
  const vistos = {};
  if (n > 0) sh.getRange(2, 1, n, 1).getValues().forEach(function (r) { vistos[String(r[0])] = true; });

  const limite = new Date(Date.now() - 60 * 864e5);
  const nuevas = [];
  GmailApp.search(BUSQUEDA_GMAIL, 0, 150).forEach(function (t) {
    t.getMessages().forEach(function (msg) {
      const id = msg.getId();
      if (vistos[id] || msg.getDate() < limite) return;
      vistos[id] = true;
      const s = extraer_(msg);
      if (s) nuevas.push(COLS_SUG.map(function (k) { return k === "id" ? id : k === "estado" ? "nueva" : (s[k] === undefined ? "" : s[k]); }));
    });
  });
  if (nuevas.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, nuevas.length, COLS_SUG.length).setNumberFormat("@");
    sh.getRange(start, 1, nuevas.length, COLS_SUG.length).setValues(nuevas);
  }
  Logger.log("Correos nuevos con cobro: " + nuevas.length);
}

/** Saca servicio, monto, moneda, frecuencia y fecha de un correo. */
function extraer_(msg) {
  const asunto = msg.getSubject() || "";
  const cuerpo = (msg.getPlainBody() || "").slice(0, 8000);
  const texto = asunto + "\n" + cuerpo;
  const de = msg.getFrom() || "";
  const d = dinero_(texto);
  const low = texto.toLowerCase();
  const prueba = /(free trial|trial|prueba gratis|periodo de prueba|período de prueba)/.test(low);
  if (!d && !prueba) return null;
  return {
    servicio: servicio_(de),
    monto: d ? d.monto : 0,
    moneda: d ? d.moneda : "",
    frecuencia: frecuencia_(low),
    fecha: Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    prueba: prueba ? "sí" : "no",
    asunto: asunto.slice(0, 120),
    de: de.slice(0, 120)
  };
}

function servicio_(de) {
  let nombre = (de.match(/^\s*"?([^"<]+?)"?\s*</) || [])[1] || "";
  const dominio = ((de.match(/@([\w-]+)\./) || [])[1] || "").replace(/^(mail|email|e|info|billing)$/i, "");
  nombre = nombre.replace(/\b(no-?reply|noreply|billing|receipts?|recibos?|facturaci[oó]n|team|equipo|support|soporte|notifications?|account|cuenta)\b/ig, "")
    .replace(/[|•·\-–—,:]+/g, " ").replace(/\s+/g, " ").trim();
  if (!nombre) nombre = dominio ? dominio.charAt(0).toUpperCase() + dominio.slice(1) : "Servicio";
  return nombre.slice(0, 40);
}

function frecuencia_(low) {
  if (/(anual|annual|yearly|per year|\/\s?(yr|year|año)|al año|por año|cada año|12 meses|12 months)/.test(low)) return "anual";
  if (/(semanal|weekly|per week|\/\s?(wk|week|semana)|cada semana)/.test(low)) return "semanal";
  return "mensual";
}

/** Primer monto con moneda que aparezca. "$" solo se deja como "$" (puede ser USD o MXN; tú lo confirmas). */
function dinero_(t) {
  const num = "(\\d{1,3}(?:[.,]\\d{3})+(?:[.,]\\d{1,2})?|\\d+(?:[.,]\\d{1,2})?)";
  const re = new RegExp("(US\\$|USD|MXN|MX\\$|EUR|€|\\$)\\s?" + num + "|" + num + "\\s?(USD|MXN|EUR|pesos|d[oó]lares)", "i");
  const m = t.match(re);
  if (!m) return null;
  let sym = (m[1] || m[4] || "").toUpperCase();
  if (sym === "$") { // "$20.00 USD" → USD
    const tras = (t.slice(m.index + m[0].length).match(/^\s*(USD|MXN|MXP)\b/i) || [])[1];
    if (tras) sym = tras.toUpperCase().replace("MXP", "MXN");
  }
  let n = m[2] || m[3];
  n = n.replace(/[.,](?=\d{3}(\D|$))/g, "").replace(",", ".");
  const monto = Math.round(parseFloat(n) * 100) / 100;
  if (!(monto >= 0)) return null;
  const moneda = /USD|US\$|D[OÓ]LARES/.test(sym) ? "USD" : /EUR|€/.test(sym) ? "EUR" : /MXN|MX\$|PESOS/.test(sym) ? "MXN" : "$";
  return { monto: monto, moneda: moneda };
}

/** Lo que tocaste en la app: Aceptar → aceptada, Ignorar → ignorada. */
function decisiones_(dec) {
  if (!dec || typeof dec !== "object" || !Object.keys(dec).length) return;
  const sh = SpreadsheetApp.openById(hoja_("mi-dinero")).getSheetByName("sugeridas");
  if (!sh || sh.getLastRow() < 2) return;
  const iEst = COLS_SUG.indexOf("estado") + 1;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_SUG.length).getValues();
  rows.forEach(function (r, i) {
    const v = dec[String(r[0])];
    const est = v === "ok" ? "aceptada" : v === "no" ? "ignorada" : null;
    if (est && r[iEst - 1] !== est) sh.getRange(i + 2, iEst).setValue(est);
  });
}

/** Sugerencias pendientes: la más reciente por servicio. Si ya decidiste (aceptar/ignorar) un recibo de un servicio,
 *  se ocultan los pendientes de ese servicio hasta esa fecha; los recibos NUEVOS (más recientes) sí aparecen. */
function sugeridas_() {
  const sh = SpreadsheetApp.openById(hoja_("mi-dinero")).getSheetByName("sugeridas");
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_SUG.length).getDisplayValues().map(function (r) {
    const o = {};
    COLS_SUG.forEach(function (k, i) { o[k] = r[i]; });
    o.monto = parseFloat(String(o.monto).replace(",", ".")) || 0;
    return o;
  });
  const key = function (o) { return String(o.servicio).toLowerCase().trim(); };
  const decididos = {}; // servicio → fecha del recibo decidido más reciente
  rows.forEach(function (o) {
    if (o.estado === "nueva") return;
    const k = key(o);
    if (decididos[k] === undefined || o.fecha > decididos[k]) decididos[k] = o.fecha;
  });
  const porServicio = {};
  rows.forEach(function (o) {
    if (o.estado !== "nueva") return;
    const k = key(o);
    if (decididos[k] !== undefined && o.fecha <= decididos[k]) return;
    if (!porServicio[k] || o.fecha > porServicio[k].fecha) porServicio[k] = o;
  });
  return Object.keys(porServicio).map(function (k) { return porServicio[k]; })
    .sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; })
    .slice(0, 30);
}
