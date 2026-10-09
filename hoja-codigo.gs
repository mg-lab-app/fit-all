// FIT-ALL → esta hoja de cálculo.
// Pega este código en la hoja: Extensiones → Apps Script, y publícalo como aplicación web.
// Escribe los entrenos en las pestañas «Sem 1», «Sem 2»…, los pesos y la cintura en «Medidas»
// y la dieta y el agua de cada día en la pestaña «Dieta (FIT-ALL)», que crea si no existe.
// No toca las fórmulas (Peso máx, Volumen, Progreso…): solo las celdas que rellenas tú.

function doGet() {
  return salida({ ok: true, app: "FIT-ALL", hoja: SpreadsheetApp.getActiveSpreadsheet().getName() });
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
  function col(nombre) { return cols.indexOf(norm(nombre)); }

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
