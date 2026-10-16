// FIT-ALL → esta hoja de cálculo.
// Pega este código en la hoja: Extensiones → Apps Script, y publícalo como aplicación web.
// Escribe los entrenos en las pestañas «Sem 1», «Sem 2»…, los pesos y la cintura en «Medidas»
// y la dieta y el agua de cada día en la pestaña «Dieta (FIT-ALL)», que crea si no existe.
// No toca las fórmulas (Peso máx, Volumen, Progreso…): solo las celdas que rellenas tú.
// Además guarda una copia de todos tus datos en la pestaña oculta «FIT-ALL datos» para que
// tus móviles estén siempre iguales. No toques esa pestaña.

function doGet() {
  return salida({ ok: true, app: "FIT-ALL", version: 2, hoja: SpreadsheetApp.getActiveSpreadsheet().getName() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var d = JSON.parse(e.postData.contents);
    // Solo acepta datos de tu app: la primera conexión guarda su clave y después la exige.
    var props = PropertiesService.getScriptProperties();
    var clave = props.getProperty("CLAVE_FITALL");
    if (!clave) { props.setProperty("CLAVE_FITALL", d.clave); }
    else if (clave !== d.clave) { return salida({ ok: false, error: "clave" }); }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (d.accion === "sync") return salida(sincronizar(ss, d));
    if (d.accion === "metodo") return salida(prepararMetodo(ss, d));
    var res = { ok: true, hoja: ss.getName(), sesiones: 0, medidas: 0, dieta: 0, noEncontrados: [] };
    (d.sesiones || []).forEach(function (s) { escribirSesion(ss, s, res); });
    (d.medidas || []).forEach(function (m) { escribirMedidas(ss, m, res); });
    (d.dieta || []).forEach(function (x) { escribirDieta(ss, x, res); });
    return salida(res);
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function norm(s) {
  return String(s == null ? "" : s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim();
}

function fechaHoja(iso) {
  var p = iso.split("-");
  return new Date(+p[0], +p[1] - 1, +p[2], 12, 0, 0);
}

function escribirSesion(ss, s, res) {
  var sh = ss.getSheetByName("Sem " + s.semana);
  if (!sh) { res.noEncontrados.push("Pestaña «Sem " + s.semana + "»"); return; }
  var vals = sh.getDataRange().getValues();

  // Fila de cabecera: «Ejercicio, Objetivo, Descanso, S1 kg, S1 reps…»
  var cab = -1;
  for (var r = 0; r < vals.length; r++) { if (norm(vals[r][0]) === "ejercicio") { cab = r; break; } }
  if (cab < 0) { res.noEncontrados.push("Cabecera de «Sem " + s.semana + "»"); return; }
  var cols = vals[cab].map(norm);
  function col(nombre) {
    var i = cols.indexOf(norm(nombre));
    if (i < 0 && ALIAS[nombre]) i = cols.indexOf(norm(ALIAS[nombre]));
    return i;
  }

  // Bloque de la sesión: empieza en «1. LUNES TARDE — …» y acaba en la siguiente sesión.
  var ini = -1, re = new RegExp("^" + s.n + "\\.\\s");
  for (var r2 = cab + 1; r2 < vals.length; r2++) { if (re.test(String(vals[r2][0]).trim())) { ini = r2; break; } }
  if (ini < 0) { res.noEncontrados.push("Sesión " + s.n + " (" + s.nombre + ") en «Sem " + s.semana + "»"); return; }
  var fin = vals.length;
  for (var r3 = ini + 1; r3 < vals.length; r3++) {
    var a = String(vals[r3][0]).trim();
    if (/^\d+\.\s/.test(a) || /^calentamiento/i.test(a)) { fin = r3; break; }
  }

  // Fecha, en la celda que hay justo después de «Fecha:».
  for (var c = 0; c < vals[ini].length; c++) {
    if (norm(vals[ini][c]) === "fecha:") { sh.getRange(ini + 1, c + 2).setValue(fechaHoja(s.fecha)); break; }
  }

  s.ejercicios.forEach(function (e) {
    var fila = -1;
    for (var r4 = ini + 1; r4 < fin; r4++) { if (norm(vals[r4][0]) === norm(e.nombre)) { fila = r4; break; } }
    if (fila < 0) { res.noEncontrados.push(e.nombre + " (" + s.fecha + ")"); return; }
    for (var k = 1; k <= 4; k++) {
      var ck = col("S" + k + " kg"), cr = col("S" + k + " reps"), se = e.series[k - 1];
      if (ck >= 0) sh.getRange(fila + 1, ck + 1).setValue(se && se.kg != null ? se.kg : "");
      if (cr >= 0) sh.getRange(fila + 1, cr + 1).setValue(se && se.reps != null ? se.reps : "");
    }
    var cRir = col("RIR"), cNot = col("Notas");
    if (cRir >= 0 && e.rir != null) sh.getRange(fila + 1, cRir + 1).setValue(e.rir);
    if (cNot >= 0 && e.notas) sh.getRange(fila + 1, cNot + 1).setValue(e.notas);
  });
  res.sesiones++;
}

function escribirMedidas(ss, m, res) {
  var sh = ss.getSheetByName("Medidas");
  if (!sh) { res.noEncontrados.push("Pestaña «Medidas»"); return; }
  var vals = sh.getDataRange().getValues();
  var cab = -1;
  for (var r = 0; r < vals.length; r++) { if (norm(vals[r][0]) === "semana") { cab = r; break; } }
  if (cab < 0) { res.noEncontrados.push("Cabecera de «Medidas»"); return; }
  var cols = vals[cab].map(norm);
  function col(nombre) { return cols.indexOf(norm(nombre)); }
  var fila = -1;
  for (var r2 = cab + 1; r2 < vals.length; r2++) { if (norm(vals[r2][0]) === norm("Sem " + m.semana)) { fila = r2; break; } }
  if (fila < 0) { res.noEncontrados.push("«Sem " + m.semana + "» en Medidas"); return; }
  var cF = col("Fecha inicio");
  if (cF >= 0 && !vals[fila][cF]) sh.getRange(fila + 1, cF + 1).setValue(fechaHoja(m.inicio));
  for (var k = 1; k <= 4; k++) {
    var c = col("Peso " + k + " (kg)");
    if (c >= 0) sh.getRange(fila + 1, c + 1).setValue(m.pesos[k - 1] != null ? m.pesos[k - 1] : "");
  }
  var cc = col("Cintura (cm)");
  if (cc >= 0 && m.cintura != null) sh.getRange(fila + 1, cc + 1).setValue(m.cintura);
  res.medidas++;
}

function escribirDieta(ss, x, res) {
  var nombre = "Dieta (FIT-ALL)";
  var sh = ss.getSheetByName(nombre);
  var cab = ["Fecha", "kcal", "Proteína (g)", "Hidratos (g)", "Grasa (g)", "Agua (L)", "Objetivo kcal", "Objetivo proteína (g)"];
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange("A:A").setNumberFormat("dd/mm/yyyy");
  }
  var tz = ss.getSpreadsheetTimeZone();
  var n = Math.max(1, sh.getLastRow());
  var fechas = sh.getRange(1, 1, n, 1).getValues().map(function (r) {
    return r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "yyyy-MM-dd") : String(r[0]);
  });
  var fila = fechas.indexOf(x.fecha) + 1;
  var v = [fechaHoja(x.fecha), x.kcal, x.p, x.c, x.g, x.agua, x.objKcal, x.objP];
  if (fila <= 1) {
    // Fila nueva en su sitio por fecha.
    var pos = n + 1;
    for (var i = 1; i < fechas.length; i++) { if (fechas[i] > x.fecha) { pos = i + 1; break; } }
    if (pos <= n) sh.insertRowBefore(pos);
    sh.getRange(pos, 1, 1, v.length).setValues([v]);
  } else {
    sh.getRange(fila, 1, 1, v.length).setValues([v]);
  }
  res.dieta++;
}

// ---------- sincronización entre móviles ----------
// Cada dato de la app (una sesión, un día de dieta, una medida…) es una fila: ruta, orden, hora y el contenido.
// Gana siempre la versión más reciente de cada dato.
var TROZO = 45000;
function sincronizar(ss, d) {
  var nombre = "FIT-ALL datos";
  var sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, 4).setValues([["Ruta", "Orden", "Hora", "Datos (no tocar)"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  var props = PropertiesService.getScriptProperties();
  var seq = Number(props.getProperty("FITALL_SEQ") || 0);
  var n = sh.getLastRow(), ancho = Math.max(4, sh.getLastColumn());
  var filas = n > 1 ? sh.getRange(2, 1, n - 1, ancho).getValues() : [];
  var idx = {};
  filas.forEach(function (f, i) { idx[f[0]] = i; });
  var cambiado = false;
  (d.cambios || []).forEach(function (c) {
    if (!c || !c.path) return;
    var i = idx[c.path];
    if (i !== undefined && Number(filas[i][2]) > Number(c.t)) return; // aquí hay una versión más nueva
    seq++;
    var json = c.data == null ? "" : JSON.stringify(c.data);
    var fila = [c.path, seq, Number(c.t)];
    for (var k = 0; k < json.length || k === 0; k += TROZO) fila.push(json.slice(k, k + TROZO));
    if (i !== undefined) filas[i] = fila; else { idx[c.path] = filas.length; filas.push(fila); }
    cambiado = true;
  });
  if (cambiado) {
    ancho = Math.max(4, filas.reduce(function (a, f) { return Math.max(a, f.length); }, 0));
    var datos = filas.map(function (f) { var r = f.slice(0, ancho); while (r.length < ancho) r.push(""); return r; });
    sh.getRange(2, 1, datos.length, ancho).setNumberFormat("@").setValues(datos);
    props.setProperty("FITALL_SEQ", String(seq));
  }
  var desde = Number(d.desde || 0), docs = [];
  filas.forEach(function (f) {
    if (Number(f[1]) <= desde) return;
    var json = f.slice(3).join("");
    docs.push({ path: f[0], t: Number(f[2]), data: json ? JSON.parse(json) : null });
  });
  return { ok: true, hoja: ss.getName(), seq: seq, docs: docs };
}

// ---------- método de alta intensidad: 1 serie de aproximación + 1-3 efectivas ----------
// Las columnas S1-S4 pasan a llamarse Aprox, E1, E2 y E3 en las semanas con el método.
var ALIAS = { "S1 kg": "Aprox kg", "S1 reps": "Aprox reps", "S2 kg": "E1 kg", "S2 reps": "E1 reps",
              "S3 kg": "E2 kg", "S3 reps": "E2 reps", "S4 kg": "E3 kg", "S4 reps": "E3 reps" };
function letra(i) { var s = ""; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
function prepararMetodo(ss, d) {
  var hechas = 0;
  Object.keys(d.semanas || {}).forEach(function (n) {
    var sh = ss.getSheetByName("Sem " + n);
    if (!sh) return;
    var info = d.semanas[n];
    var vals = sh.getDataRange().getValues();
    var cab = -1;
    for (var r = 0; r < vals.length; r++) { if (norm(vals[r][0]) === "ejercicio") { cab = r; break; } }
    if (cab < 0) return;
    // Cabeceras
    var c = {};
    Object.keys(ALIAS).forEach(function (viejo) {
      var i = vals[cab].map(norm).indexOf(norm(viejo));
      if (i < 0) i = vals[cab].map(norm).indexOf(norm(ALIAS[viejo]));
      if (i >= 0) { c[viejo] = i; sh.getRange(cab + 1, i + 1).setValue(ALIAS[viejo]); }
    });
    var cObj = vals[cab].map(norm).indexOf("objetivo");
    var cVol = -1;
    vals[cab].forEach(function (v, i) { if (norm(v).indexOf("volumen") === 0) cVol = i; });
    // Texto de la fase (fila «Nº de semana»)
    for (var r2 = 0; r2 < cab; r2++) {
      if (/^n\S*\s*de\s*semana/i.test(String(vals[r2][0]).trim()) && info.fase) { sh.getRange(r2 + 1, 3).setValue(info.fase); break; }
    }
    // Objetivo y volumen (solo efectivas) en cada ejercicio
    var obj = {};
    Object.keys(info.objetivos || {}).forEach(function (k) { obj[norm(k)] = info.objetivos[k]; });
    var pares = [["S2 kg", "S2 reps"], ["S3 kg", "S3 reps"], ["S4 kg", "S4 reps"]].filter(function (p) { return c[p[0]] !== undefined && c[p[1]] !== undefined; });
    for (var f = cab + 1; f < vals.length; f++) {
      var nombre = norm(vals[f][0]);
      if (!nombre || /^\d+\.\s/.test(String(vals[f][0]).trim()) || !obj[nombre]) continue;
      if (cObj >= 0) sh.getRange(f + 1, cObj + 1).setValue(obj[nombre]);
      if (cVol >= 0 && pares.length) {
        var fila = f + 1;
        var suma = pares.map(function (p) { return "N(" + letra(c[p[0]]) + fila + ")*N(" + letra(c[p[1]]) + fila + ")"; }).join("+");
        var rango = letra(c[pares[0][0]]) + fila + ":" + letra(c[pares[pares.length - 1][1]]) + fila;
        sh.getRange(fila, cVol + 1).setFormula("=IF(COUNT(" + rango + ")=0,\"\"," + suma + ")");
      }
    }
    hechas++;
  });
  // Explicación en «Cómo usar» (una sola vez)
  var guia = ss.getSheetByName("Cómo usar");
  if (guia) {
    var texto = guia.getDataRange().getValues().map(function (r) { return String(r[0]); }).join("\n");
    if (texto.indexOf("MÉTODO DE ALTA INTENSIDAD") < 0) {
      var fila0 = guia.getLastRow() + 2;
      var lineas = [
        ["MÉTODO DE ALTA INTENSIDAD (desde la semana " + d.desde + ")"],
        ["Antes de cada ejercicio, 1 serie de aproximación (columna Aprox): ligera, unos 6 reps con ~60 % del peso, sin cansarte. No cuenta para el volumen."],
        ["Después, 1-3 series efectivas (E1, E2, E3) al fallo o a 1 rep del fallo. En básicos con barra (sentadilla, peso muerto, press de banca, press militar, remo), para 1 rep antes de que se rompa la técnica."],
        ["Progresión: cuando en la primera serie efectiva llegues al máximo de repeticiones, sube el peso. FIT-ALL calcula los pesos de la aproximación y de las efectivas."],
        ["Ciclo: semana 1 a 1 rep del fallo · semanas 2 y 3 al fallo · semana 4 de descarga (1 efectiva y nada al fallo)."]
      ];
      guia.getRange(fila0, 1, lineas.length, 1).setValues(lineas);
      guia.getRange(fila0, 1).setFontWeight("bold");
    }
  }
  return { ok: true, metodo: hechas };
}
