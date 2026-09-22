function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  let result;
  try {
    if (action === 'list') {
      result = { ok: true, data: obtenerMovimientos() };
    } else if (action === 'config') {
      result = { ok: true, data: leerConfig() };
    } else {
      result = { ok: false, error: 'Acción no soportada: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return responderJSON(result);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'add') {
      guardarMovimiento(body.mov);
      result = { ok: true };
    } else if (body.action === 'update') {
      actualizarMovimiento(body.mov);
      result = { ok: true };
    } else if (body.action === 'delete') {
      eliminarMovimiento(body.id);
      result = { ok: true };
    } else if (body.action === 'set-config') {
      guardarConfigValor(body.key, body.value);
      result = { ok: true };
    } else if (body.action === 'set-config-batch') {
      guardarConfigValores(body.valores);
      result = { ok: true };
    } else if (body.action === 'reset-all') {
      reiniciarTodo();
      result = { ok: true };
    } else {
      result = { ok: false, error: 'Acción no soportada: ' + body.action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return responderJSON(result);
}

function responderJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Movimientos');
  if (!sheet) {
    sheet = ss.insertSheet('Movimientos');
    sheet.appendRow(['id', 'monto', 'fecha', 'nota', 'tipo', 'categoria', 'cuentaId', 'origenRecurrenteId', 'xpOtorgado']);
  } else {
    if (sheet.getRange(1, 7).getValue() === '') {
      // Hoja creada antes de que existiera "Mis tarjetas": se le agrega la
      // columna sin tocar las filas que ya tenía.
      sheet.getRange(1, 7).setValue('cuentaId');
    }
    if (sheet.getRange(1, 8).getValue() === '') {
      // Hoja creada antes de "Transacciones recurrentes": misma lógica, se
      // agrega la columna nueva sin afectar las filas existentes.
      sheet.getRange(1, 8).setValue('origenRecurrenteId');
    }
    if (sheet.getRange(1, 9).getValue() === '') {
      // Hoja creada antes de que se guardara cuánto XP otorgó cada
      // movimiento (para poder quitarlo si se borra) — misma lógica.
      sheet.getRange(1, 9).setValue('xpOtorgado');
    }
  }
  return sheet;
}

// Google Sheets auto-convierte texto tipo "2026-08-29" a un valor Date interno.
// Esa mezcla de tipos (string vs Date) es lo que rompe el render en el cliente
// al recargar. Aquí siempre se devuelve una fecha como texto plano.
function normalizarFecha(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd');
  }
  return String(valor || '');
}

function obtenerMovimientos() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // sin encabezado
  return rows
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      id: String(r[0]),
      monto: Number(r[1]) || 0,
      fecha: normalizarFecha(r[2]),
      nota: String(r[3] || ''),
      tipo: String(r[4] || '').trim().toLowerCase(),
      categoria: String(r[5] || '').trim().toLowerCase(),
      cuentaId: String(r[6] || ''),
      origenRecurrenteId: String(r[7] || ''),
      xpOtorgado: r[8] === '' || r[8] === null || r[8] === undefined ? undefined : Number(r[8])
    }))
    .reverse();
}

// Cuando el usuario borra o edita varios movimientos muy seguido, el cliente
// dispara varias peticiones casi al mismo tiempo, y Apps Script puede
// ejecutar dos doPost en paralelo. Sin bloqueo, dos ejecuciones leían la
// hoja con las mismas filas antes de que la otra borrara/insertara una, y al
// actuar sobre un número de fila que ya había cambiado, terminaban borrando
// la fila equivocada o dejando la fila objetivo intacta sin avisar del
// error — de ahí movimientos que "no se borraban bien" al hacerlo rápido.
function guardarMovimiento(mov) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // En una conexión inestable, el cliente puede recibir un error de red
      // aunque esta escritura ya haya llegado y se haya guardado — en ese
      // caso reintenta el mismo "add" más tarde (ver encolar/procesarCola en
      // el cliente). Sin esta verificación, ese reintento duplicaba el
      // movimiento entero.
      if (String(data[i][0]) === String(mov.id)) {
        return true; // ya existe, no lo vuelvas a agregar
      }
    }
    sheet.appendRow([String(mov.id), Number(mov.monto), mov.fecha, mov.nota || '', mov.tipo, mov.categoria || '', mov.cuentaId || '', mov.origenRecurrenteId || '', mov.xpOtorgado === undefined ? '' : Number(mov.xpOtorgado)]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function actualizarMovimiento(mov) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(mov.id)) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          String(mov.id), Number(mov.monto), mov.fecha, mov.nota || '', mov.tipo, mov.categoria || '', mov.cuentaId || (String(data[i][6] || '')), mov.origenRecurrenteId || (String(data[i][7] || '')), mov.xpOtorgado === undefined ? (data[i][8] === '' ? '' : Number(data[i][8])) : Number(mov.xpOtorgado)
        ]]);
        return true;
      }
    }
    throw new Error('Movimiento no encontrado');
  } finally {
    lock.releaseLock();
  }
}

function eliminarMovimiento(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ---- Config: meta de ahorro, cierres de mes ya preguntados, presupuestos por
// categoría editados a mano, y datos de perfil. Todo en pares clave/valor. ----

function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.appendRow(['clave', 'valor']);
  }
  return sheet;
}

function leerConfig() {
  const sheet = getConfigSheet();
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const clave = data[i][0];
    let valor = data[i][1];
    if (!clave) continue;
    // Un valor tipo "2026-09-03" (por ejemplo gatoUltimoDiaEvaluado) que se
    // guardó antes de forzar la celda a texto plano puede haber quedado
    // convertido por Sheets a una fecha interna; se normaliza de vuelta al
    // mismo formato yyyy-MM-dd que usa el cliente, si no, la comparación de
    // fechas en el cliente nunca coincide y el día se re-evalúa sin fin.
    if (valor instanceof Date) {
      valor = normalizarFecha(valor);
    } else {
      try { valor = JSON.parse(valor); } catch (e) { /* se queda como texto/numero */ }
    }
    config[clave] = valor;
  }
  return config;
}

// El cliente puede disparar varias llamadas set-config casi al mismo tiempo
// (una por cada clave que cambia). Sin bloqueo, dos ejecuciones podían leer
// la hoja antes de que la otra escribiera y una de las dos escrituras se
// perdía en silencio (por eso config.gatoUltimoDiaEvaluado a veces no
// quedaba guardado y el día se volvía a evaluar en la siguiente sesión).
function guardarConfigValor(clave, valor) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    escribirValorConfig(getConfigSheet(), clave, valor);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// Registrar un gasto/ingreso normalmente cambia varias claves de config a la
// vez (el XP, el registro de XP, a veces el nivel...). Antes cada una viajaba
// como su propia petición a Apps Script (que de por sí no responde al
// instante), así que el aviso de "guardando cambios pendientes..." tardaba
// en desaparecer. Esto guarda un lote entero bajo un solo bloqueo y una
// sola lectura de la hoja, en vez de una ida y vuelta por clave.
function guardarConfigValores(valores) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getConfigSheet();
    Object.keys(valores).forEach(function(clave) {
      escribirValorConfig(sheet, clave, valores[clave]);
    });
    return true;
  } finally {
    lock.releaseLock();
  }
}

// Borra todas las filas de datos de Movimientos y de Config (deja los
// encabezados intactos). Toma ambos locks porque toca las dos hojas y no
// queremos que un guardarMovimiento/guardarConfigValor a medias se cruce
// con el borrado.
function reiniciarTodo() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const hojaMov = getSheet();
    if (hojaMov.getLastRow() > 1) {
      hojaMov.deleteRows(2, hojaMov.getLastRow() - 1);
    }
    const hojaConfig = getConfigSheet();
    if (hojaConfig.getLastRow() > 1) {
      hojaConfig.deleteRows(2, hojaConfig.getLastRow() - 1);
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

// Debe llamarse ya con el LockService tomado (guardarConfigValor y
// guardarConfigValores lo hacen antes de invocarla).
function escribirValorConfig(sheet, clave, valor) {
  const data = sheet.getDataRange().getValues();
  const texto = (typeof valor === 'object') ? JSON.stringify(valor) : valor;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clave) {
      const celda = sheet.getRange(i + 1, 2);
      // Forzar texto plano: si no, Sheets convierte solo un valor como
      // "2026-09-03" en una fecha interna, y al leerlo de vuelta ya no
      // coincide con el string que espera el cliente.
      celda.setNumberFormat('@').setValue(texto);
      return;
    }
  }
  sheet.appendRow([clave, texto]);
  sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@').setValue(texto);
}
