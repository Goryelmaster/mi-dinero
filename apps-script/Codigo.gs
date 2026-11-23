/**
 * API de Santino — Google Apps Script
 * Sirve para "Mi dinero" y después para "Mi día" (cada petición trae el campo "app").
 *
 * Qué hace:
 *  - Guarda los datos en un Google Sheet "<App> - datos" (pestañas: movimientos, suscripciones, sugeridas).
 *  - Escribe "<app>-respaldo.json" en la carpeta de Drive "Datos apps" (ahí lo leen tus agentes).
 *  - Guarda una copia por día en "Datos apps/respaldos".
 *  - Una vez al día busca en tu Gmail cobros de suscripciones y los deja como "sugeridas".
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
      } else if (q.accion !== "sugeridas") {
        return out_({ ok: false, error: "Acción desconocida: " + q.accion });
      }
      const res = { ok: true };
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
  return null;
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
/** Corre sola 1 vez al día (activador). También puedes correrla a mano. */
function buscarSuscripciones() {
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

/** Sugerencias pendientes: la más reciente por servicio, y nada de un servicio que ya aceptaste o ignoraste. */
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
  const decididos = {};
  rows.forEach(function (o) { if (o.estado !== "nueva") decididos[key(o)] = true; });
  const porServicio = {};
  rows.forEach(function (o) {
    if (o.estado !== "nueva" || decididos[key(o)]) return;
    const k = key(o);
    if (!porServicio[k] || o.fecha > porServicio[k].fecha) porServicio[k] = o;
  });
  return Object.keys(porServicio).map(function (k) { return porServicio[k]; })
    .sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; })
    .slice(0, 30);
}
