// ===================================================
// ALLTANSA Board v12 — Apps Script Backend
// Google Sheet: "Hoja 1" (pedidos) + "Inventario" + "Maquinaria" + "Mantenimientos" + "Cotizaciones" + "Facturas"
// Cuenta: alltansacloud@gmail.com
// ===================================================

const SHEET_NAME   = 'Hoja 1';
const INV_SHEET    = 'Inventario';
const MOV_SHEET    = 'Movimientos';
const MAQ_SHEET    = 'Maquinaria';
const MANT_SHEET   = 'Mantenimientos';
const MAQUILA_SHEET = 'Maquila';
const COT_SHEET    = 'Cotizaciones';
const FACT_SHEET   = 'Facturas';
const LOG_SHEET    = 'LogContab';
const LOG_COT_SHEET = 'LogCotizaciones';
const PRECIOS_AIRE_SHEET          = 'ListaPrecios_Aire';
const PRECIOS_DIESEL_SHEET        = 'ListaPrecios_Diesel';
const PRECIOS_ESTACIONARIOS_SHEET = 'ListaPrecios_Estacionarios';
const PRECIOS_CARBURACION_SHEET   = 'ListaPrecios_Carburacion';
const PRECIOS_AIRE_BUMA_SHEET     = 'ListaPrecios_Aire_BUMA';
const PRECIOS_CLEVELAND_SHEET     = 'ListaPrecios_Cleveland';

// ===================================================
// SEGURIDAD: TOKEN DE ACCESO
// ===================================================
const API_SECRET = 'segurid@d-ALLTANSA';
function tokenValido(e) {
  const tk = (e.parameter && e.parameter._tk) || '';
  return tk === API_SECRET;
}

// ===================================================
// CACHÉ DE LECTURAS PESADAS (getBootstrap / getFacturas)
// ===================================================
// Con varios usuarios entrando casi al mismo tiempo (ej. inicio de jornada),
// cada uno disparaba una lectura completa del Sheet — compitiendo por el
// mismo cupo de ejecuciones simultáneas de la cuenta (gratuita, compartida
// por todos). Mientras el caché esté vigente, solo la primera persona paga
// el costo real de leer el Sheet; el resto recibe la misma respuesta ya
// lista, sin volver a abrir el Spreadsheet.
//
// TTL corto (30s) para no mostrar datos viejos por mucho rato, y se invalida
// de inmediato en cualquier guardado/borrado relevante (ver llamadas a
// cacheInvalidar en el router) — así nadie tiene que esperar a que expire
// solo para ver su propio cambio reflejado.
const CACHE_TTL_SEGUNDOS  = 30;
const CACHE_KEY_BOOTSTRAP = 'bootstrap_v1'; // orders + inventario + clientes
const CACHE_KEY_FACTURAS  = 'facturas_v1';  // facturas + abonos

function cacheGet(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; } // caché no disponible/corrupto -> se sigue como si no hubiera caché
}

function cachePut(key, value) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), CACHE_TTL_SEGUNDOS);
  } catch(e) {
    // Límite de ~100KB por valor u otro error de CacheService — no es
    // crítico, el sistema sigue funcionando igual, solo sin el beneficio
    // del caché esta vez en particular.
  }
}

function cacheInvalidar(key) {
  try { CacheService.getScriptCache().remove(key); } catch(e) {}
}

// ===================================================
// ROUTER PRINCIPAL
// ===================================================
function doGet(e) {
  if (!tokenValido(e)) {
    return ContentService
      .createTextOutput(JSON.stringify({error:'No autorizado'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const invSh = getOrCreateInvSheet(ss);
  const movSh = getOrCreateMovSheet(ss);
  const maqSh = getOrCreateSheet(ss, MAQ_SHEET);
  const mantSh = getOrCreateSheet(ss, MANT_SHEET);
  const maqlSh = getOrCreateSheet(ss, MAQUILA_SHEET);
  const factSh = getOrCreateSheet(ss, FACT_SHEET);
  const logSh  = getOrCreateSheet(ss, LOG_SHEET);
  let result;
  try {
    const action = e.parameter && e.parameter.action;
    if      (action === 'getAll')   result = getAllOrders(sheet);
    else if (action === 'save')     { result = saveOrder(sheet,   JSON.parse(decodeURIComponent(e.parameter.data))); cacheInvalidar(CACHE_KEY_BOOTSTRAP); }
    else if (action === 'delete')   { result = deleteOrder(sheet, e.parameter.id); cacheInvalidar(CACHE_KEY_BOOTSTRAP); }
    else if (action === 'getInv')   result = getAllInv(invSh);
    else if (action === 'saveInv')  { result = saveInvItem(invSh, JSON.parse(decodeURIComponent(e.parameter.data))); cacheInvalidar(CACHE_KEY_BOOTSTRAP); }
    else if (action === 'deleteInv') { result = deleteInvItem(invSh, e.parameter.id); cacheInvalidar(CACHE_KEY_BOOTSTRAP); }
    else if (action === 'logMov')   result = logMovimiento(movSh, JSON.parse(decodeURIComponent(e.parameter.data)));
    else if (action === 'getMov')   result = getMovimientos(movSh, e.parameter.modelo||null);
    else if (action === 'getPasswords') result = getPasswords(ss.getSheetByName('Usuarios'));
    // ── Bootstrap: combina getAll+getInv+getPasswords en 1 sola llamada.
    //    Reduce el número de viajes redondos a Apps Script (cada uno cuesta
    //    1-9s por el salto de redirección de Google) en la carga inicial.
    //    Cacheado 30s (ver CACHE_KEY_BOOTSTRAP) — ver bloque de caché arriba.
    else if (action === 'getBootstrap') result = getBootstrap(sheet, invSh, ss);
    // ── Mantenimiento ──────────────────────────────
    else if (action === 'getMaquinaria')      result = getAllMaquinaria(maqSh);
    else if (action === 'saveMaquinaria')     result = saveMaquinaria(maqSh, JSON.parse(decodeURIComponent(e.parameter.data)));
    else if (action === 'deleteMaquinaria')   result = deleteMaquinaria(maqSh, e.parameter.id);
    else if (action === 'getMantenimientos')  result = getAllMantenimientos(mantSh, e.parameter.equipo_id||null);
    else if (action === 'saveMantenimiento')  result = saveMantenimiento(mantSh, JSON.parse(decodeURIComponent(e.parameter.data)));
    else if (action === 'deleteMantenimiento') result = deleteMantenimiento(mantSh, e.parameter.id);
    else if (action === 'getAlertasMant')     result = getAlertasMantenimiento(maqSh, mantSh);
    // ── Maquila (servicios subcontratados) ─────
    else if (action === 'getMaquila')         result = getAllMaquila(maqlSh);
    else if (action === 'saveMaquila')        { result = saveMaquila(maqlSh, JSON.parse(e.parameter.data)); }
    else if (action === 'deleteMaquila')      result = deleteMaquila(maqlSh, e.parameter.id);
    else if (action === 'getFoliosLigero')    result = getFoliosLigero(sheet);
    // ── Cotizaciones ───────────────────────────
    else if (action === 'getCotizaciones')    result = getCotizaciones(ss);
    else if (action === 'getClientes')        result = getClientes(ss);
    else if (action === 'getLogCotizacion')   result = getLogCotizacion(ss, e.parameter.folio||'');
    // ── Contabilidad / Facturas ─────────────────
    else if (action === 'getFacturas')        result = getAllFacturas(factSh); // cacheado 30s, ver CACHE_KEY_FACTURAS
    else if (action === 'cerrarPedidoContab') { result = cerrarPedidoDesdeContab(sheet, logSh, e.parameter.id, e.parameter.by); cacheInvalidar(CACHE_KEY_BOOTSTRAP); }
    // ── Lista de precios (solo lectura — se alimenta directo del Sheet) ──
    else if (action === 'getPrecios')         result = getListaPrecios(ss);
    else                            result = { error: 'Acción no reconocida' };
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!tokenValido(e)) {
    return ContentService
      .createTextOutput(JSON.stringify({error:'No autorizado'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const action = e.parameter && e.parameter.action;
  let result;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (action === 'delete') {
      result = deleteOrder(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME), e.parameter.id);
    } else if (action === 'notificarAnticipo') {
      const od = JSON.parse(e.parameter.data || '{}');
      try { notificarLuzVerde(od); result = { ok: true }; }
      catch(eN) { result = { ok: false, error: eN.message }; }
    } else if (action === 'notificarIngenieria') {
      const od = JSON.parse(e.parameter.data || '{}');
      try { notificarIngAutorizada(od); result = { ok: true }; }
      catch(eN) { result = { ok: false, error: eN.message }; }
    } else if (action === 'saveCotizacion') {
      const cot = JSON.parse(e.parameter.data || '{}');
      const esNueva  = !cot.id;
      const antes = esNueva ? null : getCotizacionPorId(ss, cot.id);
      result = saveCotizacion(ss, cot);
      if (result.ok && !result.duplicado) {
        const logSh3  = getOrCreateSheet(ss, LOG_COT_SHEET);
        const usuario = cot.modificado_por || e.parameter.by || 'Sistema';
        const folio   = result.folio || cot.folio || '';
        if (esNueva) {
          logContab(logSh3, usuario, 'Creó cotización', folio, `Cliente: ${cot.cliente||''} — Estado: ${cot.estado||'Nueva'}`);
        } else {
          const detalleCampos = construirDiffCotizacion(antes, cot);
          // El diff de las notas del PDF lo calcula el frontend (ya sabe qué
          // texto se le mostró al usuario al abrir el editor, incluyendo
          // defaults) y llega listo en este campo transitorio — no se guarda
          // como columna, solo se usa para completar el detalle del log.
          const detallePdf = cot._diff_pdf_notas || '';
          const detalle = [detalleCampos, detallePdf].filter(Boolean).join('; ');
          // Solo se registra el movimiento si de verdad cambió algo auditado —
          // evita llenar el log con guardados que no tocaron nada relevante
          // (ej. reordenar partidas o solo tocar la bitácora).
          if (detalle) {
            logContab(logSh3, usuario, 'Editó cotización', folio, detalle);
          }
        }
      }
    } else if (action === 'saveCliente') {
      const cli = JSON.parse(e.parameter.data || '{}');
      result = saveCliente(ss, cli);
      cacheInvalidar(CACHE_KEY_BOOTSTRAP); // clientes es parte del payload de bootstrap
    } else if (action === 'saveFactura') {
      const factSh2 = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), FACT_SHEET);
      const logSh2  = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), LOG_SHEET);
      const fData = JSON.parse(e.parameter.data || '{}');
      result = saveFactura(factSh2, fData);
      if (result.ok) logContab(logSh2, e.parameter.by||'Sistema', fData.id?'editarFactura':'crearFactura', fData.folio_pedido||'', 'FAC-'+fData.num_factura+' $'+fData.monto_total);
      cacheInvalidar(CACHE_KEY_FACTURAS);
    } else if (action === 'deleteFactura') {
      const factSh2 = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), FACT_SHEET);
      const logSh2  = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), LOG_SHEET);
      result = deleteFactura(factSh2, e.parameter.id);
      if (result.ok) logContab(logSh2, e.parameter.by||'Sistema', 'eliminarFactura', e.parameter.folio||'', 'ID: '+e.parameter.id);
      cacheInvalidar(CACHE_KEY_FACTURAS);
    } else if (action === 'saveAbono') {
      const factSh2 = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), FACT_SHEET);
      const logSh2  = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), LOG_SHEET);
      const aData = JSON.parse(e.parameter.data || '{}');
      result = saveAbono(factSh2, aData);
      if (result.ok) logContab(logSh2, e.parameter.by||'Sistema', 'registrarAbono', aData.folio_pedido||'', '$'+aData.monto+' '+aData.forma_pago);
      cacheInvalidar(CACHE_KEY_FACTURAS);
    } else if (action === 'deleteAbono') {
      const factSh2 = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), FACT_SHEET);
      const logSh2  = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), LOG_SHEET);
      result = deleteAbono(factSh2, e.parameter.id);
      cacheInvalidar(CACHE_KEY_FACTURAS);
      if (result.ok) logContab(logSh2, e.parameter.by||'Sistema', 'eliminarAbono', e.parameter.folio||'', 'ID: '+e.parameter.id);
    } else {
      return doGet(e);
    }
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
const EMAILS = {
  'Ventas / Embarques': ['ventas@alltansa.com'],
  'Dirección':          ['agarza@alltansa.com', 'emilio.alvarado@alltansa.com'],
  'Contabilidad':       ['contabilidad@alltansa.com'],
  'Ingeniería de ventas': ['ingenieria@alltansa.com'],
  'Producción':         ['ingenieria02@alltansa.com'],
  'Compras':            ['compras@alltansa.com'],
  'Calidad':            ['calidad@alltansa.com']
};

// Obtener lista de correos para una o varias áreas
function getEmails(...areas) {
  const set = new Set();
  areas.forEach(a => (EMAILS[a] || []).forEach(e => set.add(e)));
  return [...set];
}

// ===================================================
// PEDIDOS — CRUD
// ===================================================
function getAllOrders(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { orders: [] };
  const headers = data[0];
  const orders = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        const offset = val.getTimezoneOffset();
        val = new Date(val.getTime() - offset * 60000).toISOString().split('T')[0];
      }
      obj[h] = val;
    });
    return obj;
  });
  return { orders: orders.filter(o => o.id && o.id !== '') };
}

// ===================================================
// FOLIO ÚNICO — generado en servidor al crear pedido
// ===================================================
/**
 * Genera el siguiente folio disponible leyendo el Sheet en el momento del guardado.
 * Formato pedidos: A-YYMM-##  |  STK: STK-YYMM-##  |  Accesorios: AC-YYMM-##  |  Flete: FLE-YYMM-##
 * Al generarse aquí (servidor) se elimina la condición de carrera del frontend.
 */
function getNextFolio(sheet, tipo) {
  const now   = new Date();
  const local = Utilities.formatDate(now, 'America/Monterrey', 'yyMM');
  const t = String(tipo || '').toLowerCase();
  const prefixes = { stk: 'STK-', accesorios: 'AC-', flete: 'FLE-' };
  const prefix = (prefixes[t] || 'A-') + local + '-';

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const folioIdx = headers.indexOf('folio');
  if (folioIdx < 0) return prefix + '01';

  const nums = data.slice(1)
    .map(r => String(r[folioIdx] || ''))
    .filter(f => f.startsWith(prefix))
    .map(f => parseInt(f.replace(prefix, ''), 10))
    .filter(n => !isNaN(n));

  const next = nums.length > 0 ? Math.max(...nums) + 1 : 1;
  return prefix + String(next).padStart(2, '0');
}

function saveOrder(sheet, data) {
  // Truncar log
  if (data.log && data.log.length > 30) data.log = data.log.slice(-30);

  const isNewOrder = !data.id || String(data.id).trim() === '' ||
                     !data.folio || String(data.folio).trim() === '';
  const necesitaFolioNuevo = isNewOrder && (!data.folio || String(data.folio).trim() === '');

  // Candado SOLO para altas nuevas — evita que dos pedidos creados casi al
  // mismo tiempo lean el mismo "último folio" y calculen el mismo siguiente
  // número (getNextFolio + el appendRow de abajo quedan protegidos como una
  // sola operación). Las ediciones de un pedido ya existente no compiten por
  // ningún folio nuevo, así que no adquieren el candado — siguen igual de
  // rápidas que antes.
  let lock = null;
  if (necesitaFolioNuevo) {
    lock = LockService.getScriptLock();
    lock.waitLock(10000); // hasta 10s esperando su turno; si truena aquí, el error ya se ve reflejado al usuario en vez de quedarse colgado
  }

  try {
    // Generar folio en servidor si el pedido es nuevo y no trae folio
    if (necesitaFolioNuevo) {
      data.folio = getNextFolio(sheet, data.tipo);
      // Si el id también viene vacío, usarlo como id
      if (!data.id || String(data.id).trim() === '') {
        data.id = data.folio;
      }
    }

    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    // Agregar columnas nuevas que estén en data pero no en headers
    const dataKeys = Object.keys(data);
    dataKeys.forEach(key => {
      if (!headers.includes(key)) {
        headers.push(key);
        // Escribir el nuevo header en la hoja
        sheet.getRange(1, headers.length).setValue(key);
      }
    });

    const row = headers.map(h => {
      const val = data[h];
      if (val === undefined || val === null) return '';
      if (typeof val === 'object') return JSON.stringify(val);
      return val;
    });

    const ids      = allData.slice(1).map(r => r[headers.indexOf('id')]);
    const rowIndex = ids.indexOf(data.id);

    if (rowIndex >= 0) {
      sheet.getRange(rowIndex + 2, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    // ── Notificaciones de handoff al crear pedido nuevo ────────────
    try {
      const esNuevo = rowIndex < 0;
      if (esNuevo && String(data.tipo || 'pedido') !== 'stk') {
        notificarPedidoNuevo(data);
        const cond = String(data.condicion_pago || '').toLowerCase();
        const pctA = Number(data.pct_anticipo || 0);
        if (cond.includes('crédito') || cond.includes('credito') || pctA === 0) {
          notificarLuzVerde(data);
        }
      }
    } catch(eNotif) {
      Logger.log('Error en notificaciones pedido nuevo: ' + eNotif.message);
    }

    return { success: true, folio: data.folio, id: data.id };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function deleteOrder(sheet, id) {
  const allData  = sheet.getDataRange().getValues();
  const headers  = allData[0];
  const idCol    = headers.indexOf('id');
  const folioCol = headers.indexOf('folio');
  const ids      = allData.slice(1).map(r => r[idCol]);
  const rowIndex = ids.indexOf(id);
  if (rowIndex >= 0) {
    // Capturar el folio ANTES de borrar la fila — antes esta línea usaba
    // 'data.folio'/'data.id', pero 'data' no existe en esta función (sus
    // parámetros son sheet, id). Eso truena con "data is not defined" cada
    // vez que un borrado SÍ tiene éxito, y el navegador recibía un error
    // aunque la fila ya se hubiera eliminado del Sheet.
    const folioEliminado = folioCol >= 0 ? allData[rowIndex + 1][folioCol] : '';
    sheet.deleteRow(rowIndex + 2);
    return { success: true, folio: folioEliminado, id: id };
  }
  return { error: 'No encontrado' };
}

// ===================================================
// INVENTARIO — CRUD
// ===================================================
function getOrCreateInvSheet(ss) {
  let sh = ss.getSheetByName(INV_SHEET);
  if (!sh) {
    sh = ss.insertSheet(INV_SHEET);
    // Encabezados
    sh.getRange(1, 1, 1, 8).setValues([[
      'id', 'modelo', 'stock_actual', 'stock_minimo',
      'precio_referencia', 'ultima_entrada', 'ultima_salida', 'notas'
    ]]);
  }
  return sh;
}

function getAllInv(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { items: [] };
  const headers = data[0];
  const items = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        const offset = val.getTimezoneOffset();
        val = new Date(val.getTime() - offset * 60000).toISOString().split('T')[0];
      }
      obj[h] = val;
    });
    return obj;
  }).filter(o => o.id && o.id !== '');
  return { items };
}

// Lee una hoja de lista de precios genéricamente — cualquier columna que Alex agregue
// o quite en el Sheet se refleja sola aquí, sin tocar código.
function leerHojaPrecios(sheet, categoria) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = { categoria };
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        const offset = val.getTimezoneOffset();
        val = new Date(val.getTime() - offset * 60000).toISOString().split('T')[0];
      }
      obj[h] = val;
    });
    return obj;
  }).filter(o => o.modelo && String(o.modelo).trim() !== '');
}

function getListaPrecios(ss) {
  const shAire          = ss.getSheetByName(PRECIOS_AIRE_SHEET);
  const shDiesel        = ss.getSheetByName(PRECIOS_DIESEL_SHEET);
  const shEstacionarios = ss.getSheetByName(PRECIOS_ESTACIONARIOS_SHEET);
  const shCarburacion   = ss.getSheetByName(PRECIOS_CARBURACION_SHEET);
  const shAireBuma      = ss.getSheetByName(PRECIOS_AIRE_BUMA_SHEET);
  const shCleveland     = ss.getSheetByName(PRECIOS_CLEVELAND_SHEET);
  const items = [
    ...leerHojaPrecios(shAire, 'Aire'),
    ...leerHojaPrecios(shDiesel, 'Diésel'),
    ...leerHojaPrecios(shEstacionarios, 'Estacionarios'),
    ...leerHojaPrecios(shCarburacion, 'Carburación'),
    ...leerHojaPrecios(shAireBuma, 'Aire — BUMA'),
    ...leerHojaPrecios(shCleveland, 'Aire — Cleveland').map(procesarItemCleveland)
  ];
  return { items };
}

// La lista de Cleveland es casi toda de aire, salvo el modelo MIL235G (de
// gasolina). En vez de pedirle a Alex que capture categoría/orientación a
// mano por renglón, se detecta solo a partir del código del modelo:
// termina en "G" -> gasolina; empieza con "V" después del número -> vertical;
// cualquier otro caso (incluido "H") -> horizontal (default).
function procesarItemCleveland(item) {
  const m = String(item.modelo || '').toUpperCase();
  const match = m.match(/^MIL\d+([A-Z].*)?$/);
  const sufijo = match ? (match[1] || '') : '';
  const esGasolina = sufijo.charAt(0) === 'G';
  const esVertical = sufijo.charAt(0) === 'V';
  item.categoria  = esGasolina ? 'Gasolina — Cleveland' : 'Aire — Cleveland';
  item.orientacion = esVertical ? 'Vertical' : 'Horizontal';
  item.fluido = esGasolina ? 'Gasolina' : 'Aire';
  return item;
}

function saveInvItem(sheet, data) {
  const allData  = sheet.getDataRange().getValues();
  const headers  = allData[0];

  // Agregar columnas nuevas que estén en data pero no en headers (igual que saveOrder)
  const dataKeys = Object.keys(data);
  dataKeys.forEach(key => {
    if (!headers.includes(key)) {
      headers.push(key);
      sheet.getRange(1, headers.length).setValue(key);
    }
  });

  const row = headers.map(h => {
    const val = data[h];
    if (val === undefined || val === null) return '';
    return val;
  });
  const ids      = allData.slice(1).map(r => r[headers.indexOf('id')]);
  const rowIndex = ids.indexOf(data.id);
  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 2, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true, id: data.id };
}

function deleteInvItem(sheet, id) {
  const allData  = sheet.getDataRange().getValues();
  const headers  = allData[0];
  const ids      = allData.slice(1).map(r => r[headers.indexOf('id')]);
  const rowIndex = ids.indexOf(id);
  if (rowIndex >= 0) {
    sheet.deleteRow(rowIndex + 2);
    return { success: true, id: id };
  }
  return { error: 'Modelo no encontrado' };
}

// ===================================================
// MOVIMIENTOS — HISTORIAL COMPLETO
// ===================================================
function getOrCreateMovSheet(ss) {
  let sh = ss.getSheetByName(MOV_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MOV_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['fecha','tipo','modelo','qty','folio','usuario']]);
  }
  return sh;
}

function logMovimiento(sheet, data) {
  Logger.log('logMov recibido: ' + JSON.stringify(data));
  // data: { fecha, tipo, modelo, qty, folio, usuario }
  sheet.appendRow([
    data.fecha || '',
    data.tipo  || '',
    data.modelo|| '',
    data.qty   || 0,
    data.folio || '',
    data.usuario|| ''
  ]);
  return { success: true, folio: data.folio, id: data.id };
}

function getMovimientos(sheet, modeloFiltro) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { movimientos: [] };
  const headers = data[0];
  let rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
  if (modeloFiltro) {
    rows = rows.filter(r => r.modelo === modeloFiltro);
  }
  // Ordenar más reciente primero
  rows.reverse();
  return { movimientos: rows };
}

function getPasswords(sheet) {
  if (!sheet) return { passwords: {} };
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { passwords: {} };
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const areaIdx = headers.indexOf('area');
  const passIdx = headers.indexOf('password');
  if (areaIdx < 0 || passIdx < 0) return { passwords: {} };
  const passwords = {};
  data.slice(1).forEach(row => {
    const area = String(row[areaIdx] || '').trim();
    const pass = String(row[passIdx] || '').trim();
    if (area && pass) passwords[area] = pass;
  });
  return { passwords };
}

function getBootstrap(sheet, invSh, ss) {
  const cacheado = cacheGet(CACHE_KEY_BOOTSTRAP);
  if (cacheado) return cacheado;

  const ordersResult = getAllOrders(sheet);
  const invResult    = getAllInv(invSh);
  // Catálogo de clientes — mismo que usa Cotizaciones. Se incluye aquí para
  // que el Board de Pedidos también pueda ofrecer autocompletado al dar de
  // alta un pedido nuevo, sin necesitar una llamada aparte al servidor.
  let clientesResult = { clientes: [] };
  try { clientesResult = getClientes(ss); } catch(e) { /* si falla, el Board sigue funcionando sin autocompletado */ }
  const resultado = {
    orders: ordersResult.orders || [],
    items: invResult.items || [],
    clientes: clientesResult.clientes || []
  };
  cachePut(CACHE_KEY_BOOTSTRAP, resultado);
  return resultado;
}




// ===================================================
// REPORTE DIARIO CONSOLIDADO
// Un solo correo por día a todas las áreas,
// con sección personalizada por área.
// ===================================================

function enviarReporteDiario() {
  const notifKey = 'reporte_diario_' + today();
  if (wasSentToday(notifKey)) {
    Logger.log('Reporte diario ya enviado hoy.');
    return;
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const result = getAllOrders(sheet);
  const orders = (result.orders || []).filter(o => o.id);

  // Cargar alertas de mantenimiento
  const maqSh  = getOrCreateSheet(ss, MAQ_SHEET);
  const mantSh = getOrCreateSheet(ss, MANT_SHEET);
  const alertasMant = getAlertasMantenimiento(maqSh, mantSh).alertas || [];

  // Parsear campos JSON
  orders.forEach(o => {
    if (typeof o.stages === 'string')      { try { o.stages = JSON.parse(o.stages); }           catch(e){ o.stages = {}; } }
    if (typeof o.prod_stages === 'string') { try { o.prod_stages = JSON.parse(o.prod_stages); } catch(e){ o.prod_stages = {}; } }
  });

  // Solo pedidos activos (no cerrados, no cancelados)
  const activos = orders.filter(o => {
    if (o.cancelado) return false;
    if (o.stages && o.stages.cerrado && o.stages.cerrado.done) return false;
    return true;
  });

  // ── Datos por sección ──────────────────────────────────────────────

  // Helper: calidad liberada (campo doble)
  function calLiberadaFn(o) {
    return (o.prod_stages && o.prod_stages.lib_calidad && o.prod_stages.lib_calidad.done) ||
           (o.calidad && o.calidad.toLowerCase() === 'liberado');
  }

  // PRODUCCIÓN: pedidos con orden activa, en fabricación (calidad NO liberada, no embarcados)
  const enProduccion = activos.filter(o =>
    o.tipo !== 'stk' &&
    !(o.from_stock && (parseInt(o.qty_produccion) || 0) === 0) &&
    o.stages && o.stages.orden_prod && o.stages.orden_prod.done &&
    !(o.prod_stages && o.prod_stages.embarque && o.prod_stages.embarque.done) &&
    !calLiberadaFn(o)  // excluir si calidad ya liberó — esos son de Ventas
  );

  // PRODUCCIÓN: órdenes nuevas ≤48 hrs
  const ordenesNuevas = activos.filter(o =>
    o.stages && o.stages.orden_prod && o.stages.orden_prod.done &&
    daysSinceStr(o.stages.orden_prod.date) !== null &&
    daysSinceStr(o.stages.orden_prod.date) <= 2
  );

  // PRODUCCIÓN: material recibido hoy
  const materialHoy = activos.filter(o => {
    const placasHoy = daysSinceStr(o.compras_placas_fecha) === 0 && o.compras_placas === 'Recibido';
    const accHoy    = daysSinceStr(o.compras_accesorios_fecha) === 0 && o.compras_accesorios === 'Recibido';
    return placasHoy || accHoy;
  });

  // VENTAS: próximos a vencer sin embarcar
  const proxVencer = activos.filter(o => {
    if (!o.deadline || o.tipo === 'stk') return false;
    const d = daysUntilStr(o.deadline);
    if (d === null || d > 5) return false;
    if (o.prod_stages && o.prod_stages.embarque && o.prod_stages.embarque.done) return false;
    return true;
  });

  // CALIDAD: sandblast hecho, calidad NO liberada, no embarcado
  const pendCalidad = activos.filter(o => {
    if (o.tipo === 'stk') return false;
    if (o.from_stock && (parseInt(o.qty_produccion) || 0) === 0) return false;
    if (!o.prod_stages) return false;
    if (!(o.prod_stages.sandblast && o.prod_stages.sandblast.done)) return false;
    if (o.prod_stages.embarque && o.prod_stages.embarque.done) return false;
    return !calLiberadaFn(o);
  });

  // VENTAS: listos para embarcar (calidad ok + finiquito ok)
  const listosEmbarque = activos.filter(o => {
    if (o.tipo === 'stk') return false;
    if (o.prod_stages && o.prod_stages.embarque && o.prod_stages.embarque.done) return false;
    const fin    = o.finiquito || 'Pendiente';
    const finOk  = fin === 'Pagado' || fin === 'Crédito autorizado' || fin === 'Cartera vencida';
    const fromStock = !!(o.from_stock && (parseInt(o.qty_produccion) || 0) === 0);
    return fromStock ? finOk : (calLiberadaFn(o) && finOk);
  });

  // VENTAS: pendientes de embarque — calidad liberada pero finiquito pendiente
  const pendEmbarque = activos.filter(o => {
    if (o.tipo === 'stk') return false;
    if (o.prod_stages && o.prod_stages.embarque && o.prod_stages.embarque.done) return false;
    if (!calLiberadaFn(o)) return false;
    const fin   = o.finiquito || 'Pendiente';
    const finOk = fin === 'Pagado' || fin === 'Crédito autorizado' || fin === 'Cartera vencida';
    return !finOk;
  });

  // VENTAS: embarcados ayer
  const embarcadosAyer = activos.filter(o =>
    o.prod_stages && o.prod_stages.embarque && o.prod_stages.embarque.done &&
    daysSinceStr(o.prod_stages.embarque.date) === 1
  );

  // INGENIERÍA: pendientes de autorizar
  const pendIngenieria = activos.filter(o =>
    o.tipo !== 'stk' &&
    o.stages && (
      !(o.stages.ingenieria && o.stages.ingenieria.done) ||
      !(o.stages.requisicion && o.stages.requisicion.done)
    ) &&
    (o.anticipo_status === 'Anticipo recibido' || o.anticipo_status === 'Sin anticipo')
  );

  // COMPRAS: material pendiente con orden activa — excluir from_stock y ya embarcados
  const pendCompras = activos.filter(o =>
    o.tipo !== 'stk' &&
    !(o.from_stock && (parseInt(o.qty_produccion) || 0) === 0) && // excluir from_stock
    !(o.prod_stages && o.prod_stages.embarque && o.prod_stages.embarque.done) && // no embarcado
    o.stages && o.stages.orden_prod && o.stages.orden_prod.done && (
      (o.compras_placas || 'Pendiente') !== 'Recibido' ||
      (o.compras_accesorios || 'Pendiente') !== 'Recibido'
    )
  );

  // CONTABILIDAD: anticipos sin confirmar que bloquean
  const pendAnticipo = activos.filter(o =>
    o.tipo !== 'stk' &&
    (o.anticipo_status || 'Pendiente') === 'Pendiente' &&
    (parseInt(o.pct_anticipo) || 0) > 0
  );

  // CONTABILIDAD: pedidos pendientes de finiquito (embarcados O listos para embarcar con cal. liberada)
  const pendFiniquito = activos.filter(o => {
    if (o.tipo === 'stk') return false;
    const fin = o.finiquito || 'Pendiente';
    if (fin === 'Pagado') return false;
    // Ya embarcado y sin cobrar
    if (o.prod_stages && o.prod_stages.embarque && o.prod_stages.embarque.done) return true;
    // Calidad liberada, listo para embarcar, finiquito pendiente
    if (calLiberadaFn(o)) return true;
    // From_stock con finiquito pendiente
    if (o.from_stock && (parseInt(o.qty_produccion) || 0) === 0) return true;
    return false;
  });

  // DIRECCIÓN: bloqueados >3 días
  const bloqueados = activos.filter(o =>
    (o.bloqueo || 'Ninguno') !== 'Ninguno' &&
    o.lastProdUpdate && daysSinceStr(o.lastProdUpdate) > 3
  );

  // DIRECCIÓN: cancelados ayer
  const canceladosAyer = orders.filter(o =>
    o.cancelado && daysSinceStr(o.cancelado_date) === 1
  );

  // Si no hay nada relevante, no enviar
  const hayContenido = [enProduccion, ordenesNuevas, materialHoy, proxVencer, pendCalidad,
    listosEmbarque, pendEmbarque, pendIngenieria, pendCompras, pendAnticipo,
    pendFiniquito, bloqueados, alertasMant].some(arr => arr.length > 0);

  if (!hayContenido) {
    Logger.log('Sin contenido relevante para el reporte de hoy.');
    markSent(notifKey);
    return;
  }

  // Fecha formateada
  const fecha = new Date().toLocaleDateString('es-MX', {
    weekday:'long', day:'2-digit', month:'long', year:'numeric'
  });
  const fechaCap = fecha.charAt(0).toUpperCase() + fecha.slice(1);

  const html = buildReporteDiario(fechaCap, {
    enProduccion, ordenesNuevas, materialHoy, proxVencer, pendCalidad,
    listosEmbarque, embarcadosAyer, pendEmbarque, pendIngenieria,
    pendCompras, pendAnticipo, pendFiniquito,
    bloqueados, canceladosAyer, alertasMant,
    totalActivos:    activos.filter(o => o.tipo !== 'stk').length,
    totalRiesgo:     proxVencer.length,
    totalBloqueados: bloqueados.length,
    totalCancelados: canceladosAyer.length
  });

  const todos = getEmails('Ventas / Embarques','Dirección','Contabilidad',
                          'Ingeniería de ventas','Producción','Compras','Calidad','Mantenimiento');
  sendEmail(todos, 'Reporte Diario ALLTANSA Board — ' + fechaCap, html);
  markSent(notifKey);
  Logger.log('✅ Reporte diario enviado a: ' + todos.join(', '));
}

// ── Construcción HTML del reporte ─────────────────────────────────────────

function buildReporteDiario(fecha, d) {
  const URL_BOARD = 'https://alltansacloud.github.io/alltansa-board/';

  function tabla(headers, rows) {
    if (!rows || rows.length === 0)
      return '<p style="color:#6B7A99;font-size:12px;font-style:italic">Sin elementos para reportar hoy.</p>';
    const ths = headers.map(h =>
      '<th style="background:#1B2A4A;color:#fff;padding:6px 10px;font-size:11px;text-align:left;white-space:nowrap">' + h + '</th>'
    ).join('');
    const trs = rows.map((r, i) =>
      '<tr style="background:' + (i % 2 === 0 ? '#F8F9FC' : '#fff') + '">' +
      r.map(c => '<td style="padding:6px 10px;font-size:12px;color:#1B2A4A;border-bottom:1px solid #E8EAF0">' + (c !== null && c !== undefined ? c : '—') + '</td>').join('') +
      '</tr>'
    ).join('');
    return '<table style="width:100%;border-collapse:collapse;margin-bottom:8px"><thead><tr>' + ths + '</tr></thead><tbody>' + trs + '</tbody></table>';
  }

  function seccion(emoji, titulo, color, contenido) {
    return '<div style="margin-bottom:20px;border-left:4px solid ' + color + ';padding-left:12px">' +
      '<div style="font-size:14px;font-weight:700;color:' + color + ';margin-bottom:8px">' + emoji + '&nbsp;' + titulo + '</div>' +
      contenido + '</div>';
  }

  function kpiBox(label, valor, color) {
    return '<div style="background:#F8F9FC;border-radius:8px;padding:10px 16px;min-width:110px;text-align:center;display:inline-block;margin:4px">' +
      '<div style="font-size:22px;font-weight:700;color:' + color + '">' + valor + '</div>' +
      '<div style="font-size:10px;color:#6B7A99;margin-top:2px">' + label + '</div>' +
      '</div>';
  }

  // ── PRODUCCIÓN ──────────────────────────────────────────────────────
  const secProd = seccion('', 'PRODUCCIÓN', '#34A853',
    (d.enProduccion.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Pedidos en producción:</p>' +
      tabla(['Folio','Cliente','Modelo','Entrega','Etapa actual','Días rest.'],
        d.enProduccion.map(o => {
          const dias = daysUntilStr(o.deadline);
          const sem  = dias === null ? '—' : dias < 0 ? 'VENCIDO ' + Math.abs(dias) + 'd' : dias === 0 ? 'Vence hoy' : dias + 'd';
          return [o.folio, o.client, o.desc, formatDate(o.deadline), getEtapaActual(o), sem];
        })) : '') +
    (d.ordenesNuevas.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Ordenes nuevas (ultimas 48 hrs):</p>' +
      tabla(['Folio','Cliente','Modelo','Entrega'],
        d.ordenesNuevas.map(o => [o.folio, o.client, o.desc, formatDate(o.deadline)])) : '') +
    (d.materialHoy.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Material recibido hoy — listos para iniciar:</p>' +
      tabla(['Folio','Cliente','Modelo','Material'],
        d.materialHoy.map(o => {
          const p = o.compras_placas === 'Recibido' && daysSinceStr(o.compras_placas_fecha) === 0;
          const a = o.compras_accesorios === 'Recibido' && daysSinceStr(o.compras_accesorios_fecha) === 0;
          return [o.folio, o.client, o.desc, [p?'Placas':'', a?'Accesorios':''].filter(Boolean).join(' y ')];
        })) : '') +
    (d.enProduccion.length === 0 && d.ordenesNuevas.length === 0 && d.materialHoy.length === 0 ?
      '<p style="color:#6B7A99;font-size:12px;font-style:italic">Sin novedades para produccion hoy.</p>' : '')
  );

  // ── CALIDAD ─────────────────────────────────────────────────────────
  const secCal = seccion('', 'CALIDAD', '#9C27B0',
    d.pendCalidad.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Pendientes de liberación (Sandblast completado):</p>' +
      tabla(['Folio','Cliente','Modelo','Sandblast completado'],
        d.pendCalidad.map(o => [o.folio, o.client, o.desc,
          formatDate(o.prod_stages && o.prod_stages.sandblast ? o.prod_stages.sandblast.date : null)])) :
      '<p style="color:#6B7A99;font-size:12px;font-style:italic">Sin pendientes de liberación hoy.</p>'
  );

  // ── VENTAS / EMBARQUES ───────────────────────────────────────────────
  const secVentas = seccion('', 'VENTAS / EMBARQUES', '#1a73e8',
    (d.listosEmbarque.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Listos para embarcar:</p>' +
      tabla(['Folio','Cliente','Modelo','Condicion'],
        d.listosEmbarque.map(o => [o.folio, o.client, o.desc, o.condicion_pago || 'Contado'])) : '') +
    (d.pendEmbarque.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Pendientes de embarque (finiquito pendiente):</p>' +
      tabla(['Folio','Cliente','Modelo','Condicion','Finiquito'],
        d.pendEmbarque.map(o => [o.folio, o.client, o.desc,
          o.condicion_pago || 'Contado', o.finiquito || 'Pendiente'])) : '') +
    (d.proxVencer.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Pedidos vencidos o proximos a vencer:</p>' +
      tabla(['Folio','Cliente','Entrega','Dias rest.'],
        d.proxVencer.map(o => {
          const dias = daysUntilStr(o.deadline);
          const sem  = dias === null ? '—' : dias < 0 ? 'VENCIDO ' + Math.abs(dias) + 'd' : dias === 0 ? 'Vence hoy' : dias + 'd';
          return [o.folio, o.client, formatDate(o.deadline), sem];
        })) : '') +
    (d.embarcadosAyer.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Embarcados ayer:</p>' +
      tabla(['Folio','Cliente','Modelo'],
        d.embarcadosAyer.map(o => [o.folio, o.client, o.desc])) : '') +
    (d.listosEmbarque.length === 0 && d.pendEmbarque.length === 0 && d.proxVencer.length === 0 && d.embarcadosAyer.length === 0 ?
      '<p style="color:#6B7A99;font-size:12px;font-style:italic">Sin novedades para ventas y embarques hoy.</p>' : '')
  );

  // ── INGENIERÍA DE VENTAS ─────────────────────────────────────────────
  const secIng = seccion('', 'INGENIERÍA DE VENTAS', '#E37400',
    d.pendIngenieria.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Pendientes de autorización:</p>' +
      tabla(['Folio','Cliente','Modelo','Pendiente'],
        d.pendIngenieria.map(o => {
          const ing  = o.stages && o.stages.ingenieria  && o.stages.ingenieria.done;
          const req  = o.stages && o.stages.requisicion && o.stages.requisicion.done;
          const pend = (!ing && !req) ? 'Ingeniería + Requisición' : !ing ? 'Ingeniería' : 'Requisición';
          return [o.folio, o.client, o.desc, pend];
        })) :
      '<p style="color:#6B7A99;font-size:12px;font-style:italic">Sin pendientes de autorización hoy.</p>'
  );

  // ── COMPRAS ──────────────────────────────────────────────────────────
  const secCompras = seccion('', 'COMPRAS', '#F4511E',
    d.pendCompras.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Material pendiente:</p>' +
      tabla(['Folio','Cliente','Modelo','Placas','Accesorios'],
        d.pendCompras.map(o => [o.folio, o.client, o.desc,
          o.compras_placas || 'Pendiente', o.compras_accesorios || 'Pendiente'])) :
      '<p style="color:#6B7A99;font-size:12px;font-style:italic">Sin material pendiente hoy.</p>'
  );

  // ── CONTABILIDAD ─────────────────────────────────────────────────────
  const secContab = seccion('', 'CONTABILIDAD', '#00796B',
    (d.pendAnticipo.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Anticipos sin confirmar (bloquean producción):</p>' +
      tabla(['Folio','Cliente','Modelo'],
        d.pendAnticipo.map(o => [o.folio, o.client, o.desc])) : '') +
    (d.pendFiniquito.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Embarques pendientes de confirmar finiquito:</p>' +
      tabla(['Folio','Cliente','Modelo','Días desde embarque'],
        d.pendFiniquito.map(o => {
          const dias = daysSinceStr(o.prod_stages && o.prod_stages.embarque ? o.prod_stages.embarque.date : null);
          return [o.folio, o.client, o.desc, dias !== null ? dias + 'd' : '—'];
        })) : '') +
    (d.pendAnticipo.length === 0 && d.pendFiniquito.length === 0 ?
      '<p style="color:#6B7A99;font-size:12px;font-style:italic">Sin pendientes de contabilidad hoy.</p>' : '')
  );

  // ── DIRECCIÓN ────────────────────────────────────────────────────────
  const secDir = seccion('', 'DIRECCIÓN', '#1B2A4A',
    '<div style="margin-bottom:12px">' +
      kpiBox('Pedidos activos',    d.totalActivos,    '#1B2A4A') +
      kpiBox('En riesgo/vencidos', d.totalRiesgo,     d.totalRiesgo > 0     ? '#C0391B' : '#34A853') +
      kpiBox('Bloqueados >3 días', d.totalBloqueados, d.totalBloqueados > 0 ? '#E37400' : '#34A853') +
      kpiBox('Cancelados ayer',    d.totalCancelados, d.totalCancelados > 0 ? '#9E9E9E' : '#34A853') +
    '</div>' +
    (d.bloqueados.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Pedidos bloqueados >3 días:</p>' +
      tabla(['Folio','Cliente','Modelo','Causa','Días bloqueado'],
        d.bloqueados.map(o => [o.folio, o.client, o.desc, o.bloqueo || '—',
          o.lastProdUpdate ? daysSinceStr(o.lastProdUpdate) + 'd' : '—'])) : '') +
    (d.canceladosAyer.length > 0 ?
      '<p style="font-size:11px;font-weight:600;color:#5F6368;margin:4px 0">Cancelaciones de ayer:</p>' +
      tabla(['Folio','Cliente','Modelo','Motivo'],
        d.canceladosAyer.map(o => [o.folio, o.client, o.desc, o.cancelado_motivo || '—'])) : '')
  );

  // Sección Mantenimiento — solo aparece si hay alertas
  const secMant = (d.alertasMant && d.alertasMant.length > 0) ? seccion('🔧', 'MANTENIMIENTO', '#E37400',
    '<p style="font-size:11px;color:#5F6368;margin:0 0 8px 0">Equipos con mantenimiento próximo o vencido:</p>' +
    tabla(['Equipo','Área','Próximo mant.','Estado'],
      d.alertasMant.map(a => [
        a.equipo_nombre,
        a.area || '—',
        a.proxima_fecha || '—',
        a.dias < 0
          ? '🔴 Vencido hace ' + Math.abs(a.dias) + ' días'
          : '🟡 En ' + a.dias + ' días'
      ]))
  ) : '';

  return `<!DOCTYPE html>
<html><body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#F4F5F7;margin:0;padding:20px">
<div style="max-width:680px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.1)">
  <div style="background:#1B2A4A;padding:18px 24px;display:flex;align-items:center;gap:12px">
    <div style="color:#fff;font-size:18px;font-weight:700;letter-spacing:1px">ALLTANSA</div>
    <div style="color:rgba(255,255,255,.5);font-size:12px">Board de Pedidos</div>
    <div style="margin-left:auto;color:rgba(255,255,255,.7);font-size:12px">Reporte Diario</div>
  </div>
  <div style="background:#E8F0FE;padding:10px 24px;font-size:13px;font-weight:600;color:#1a73e8">
    ${fecha}
  </div>
  <div style="padding:20px 24px">
    ${secProd}
    ${secCal}
    ${secVentas}
    ${secIng}
    ${secCompras}
    ${secContab}
    ${secDir}
    ${secMant}
    <div style="margin-top:24px;text-align:center">
      <a href="${URL_BOARD}" style="background:#C0391B;color:#fff;padding:10px 28px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:600">
        Ver Board completo →
      </a>
    </div>
  </div>
  <div style="background:#F8F9FC;padding:10px 24px;font-size:11px;color:#6B7A99;text-align:center">
    Reporte automático diario · ALLTANSA Board · ${new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'})}
  </div>
</div>
</body></html>`;
}

function getEtapaActual(o) {
  if (!o.prod_stages) return 'Sin iniciar';
  const etapas = ['recepcion_mat','corte','rolado_tapas','ensamble','soldadura','prueba_hidro','sandblast','lib_calidad','embarque'];
  const labels  = ['Recep. mat.','Corte','Rolado/tapas','Ensamble','Soldadura','Prueba hidro.','Sandblast','Lib. Calidad','Embarque'];
  for (let i = 0; i < etapas.length; i++) {
    if (!o.prod_stages[etapas[i]] || !o.prod_stages[etapas[i]].done) return labels[i];
  }
  return 'Terminado';
}

// ===================================================
// HELPERS DE NOTIFICACIÓN
// ===================================================

function sendEmail(destinations, subject, htmlBody) {
  if (!destinations || destinations.length === 0) return;
  try {
    // Un solo envío: primer destinatario como TO, resto como BCC
    const to  = destinations[0];
    const bcc = destinations.slice(1).join(',');
    GmailApp.sendEmail(to, subject, '', {
      htmlBody: htmlBody,
      name: 'ALLTANSA Board',
      bcc: bcc || ''
    });
    Logger.log('✅ Correo enviado a: ' + destinations.join(', '));
  } catch(e) {
    Logger.log('Error enviando correo: ' + e.message);
  }
}

function buildEmailBody(titulo, campos) {
  const rows = Object.entries(campos).map(([k, v]) =>
    '<tr><td style="padding:6px 12px;font-weight:600;color:#6B7A99;font-size:13px;white-space:nowrap">' + k + '</td>' +
    '<td style="padding:6px 12px;font-size:13px;color:#1B2A4A">' + v + '</td></tr>'
  ).join('');
  return `<!DOCTYPE html>
<html><body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#F4F5F7;margin:0;padding:20px">
<div style="max-width:540px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.1)">
  <div style="background:#1B2A4A;padding:18px 24px;display:flex;align-items:center;gap:12px">
    <div style="color:#fff;font-size:18px;font-weight:700;letter-spacing:1px">ALLTANSA</div>
    <div style="color:rgba(255,255,255,.5);font-size:12px">Board de Pedidos</div>
  </div>
  <div style="padding:20px 24px">
    <div style="font-size:16px;font-weight:700;color:#1B2A4A;margin-bottom:16px">${titulo}</div>
    <table style="width:100%;border-collapse:collapse;background:#F8F9FC;border-radius:8px;overflow:hidden">
      ${rows}
    </table>
    <div style="margin-top:20px;text-align:center">
      <a href="https://alltansacloud.github.io/alltansa-board/"
         style="background:#C0391B;color:#fff;padding:10px 24px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:600">
        Ver en el Board →
      </a>
    </div>
  </div>
  <div style="background:#F8F9FC;padding:10px 24px;font-size:11px;color:#6B7A99;text-align:center">
    Mensaje automático de ALLTANSA Board · ${new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'})}
  </div>
</div>
</body></html>`;
}

// ===================================================
// DEDUPLICACIÓN: evitar re-envío del mismo correo
// Usa PropertiesService para guardar qué se envió hoy
// ===================================================
const PROPS = PropertiesService.getScriptProperties();

function wasSentToday(key) {
  const val = PROPS.getProperty(key);
  if (!val) return false;
  return val === today();
}

function markSent(key) {
  PROPS.setProperty(key, today());
}

// Limpiar propiedades antiguas (llamar mensualmente si hay muchas)
function cleanOldProps() {
  const allProps = PROPS.getProperties();
  const t = today();
  Object.keys(allProps).forEach(k => {
    if (allProps[k] !== t) PROPS.deleteProperty(k);
  });
}

// Limpiar solo la deduplicación del reporte diario para poder probar de nuevo hoy
function resetReporteDiario() {
  PROPS.deleteProperty('reporte_diario_' + today());
  Logger.log('✅ Deduplicación limpiada — puedes ejecutar enviarReporteDiario de nuevo');
}

// ===================================================
// HELPERS DE FECHA
// ===================================================
function today() {
  const d = new Date();
  return d.getFullYear() + '-' +
    String(d.getMonth() + 1).padStart(2, '0') + '-' +
    String(d.getDate()).padStart(2, '0');
}

function parseDate(str) {
  if (!str) return null;
  if (typeof str === 'string') {
    const s = str.substring(0, 10);
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      const p = s.split('-');
      return new Date(+p[0], +p[1] - 1, +p[2]);
    }
  }
  if (typeof str === 'number' && str > 0) {
    return new Date(Math.round((str - 25569) * 86400 * 1000));
  }
  return null;
}

function formatDate(str) {
  const d = parseDate(str);
  if (!d) return '—';
  return d.toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric' });
}

function daysSinceStr(str) {
  const d = parseDate(str);
  if (!d) return null;
  const t = new Date(); t.setHours(0, 0, 0, 0); d.setHours(0, 0, 0, 0);
  return Math.round((t - d) / 86400000);
}

function daysUntilStr(str) {
  const d = parseDate(str);
  if (!d) return null;
  const t = new Date(); t.setHours(0, 0, 0, 0); d.setHours(0, 0, 0, 0);
  return Math.round((d - t) / 86400000);
}

// ===================================================
// MANTENIMIENTO — HELPERS DE HOJAS
// ===================================================

// Helper genérico: obtener o crear hoja con headers mínimos
function getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === MAQ_SHEET) {
      sh.getRange(1, 1, 1, 9).setValues([[
        'id','nombre','marca','modelo','num_serie','area','fecha_adquisicion','estatus','notas'
      ]]);
    } else if (name === MANT_SHEET) {
      sh.getRange(1, 1, 1, 16).setValues([[
        'id','equipo_id','equipo_nombre','tipo','fecha','descripcion',
        'proveedor','tecnico','costo','observaciones',
        'proxima_fecha','periodicidad','causa_falla','genera_preventivo',
        'createdAt','createdBy'
      ]]);
    } else if (name === MAQUILA_SHEET) {
      sh.getRange(1, 1, 1, 15).setValues([[
        'id','folio_pedido','servicio','proveedor','oc_alltansa',
        'fecha_envio','fecha_recogido','factura','fecha_pago','comentarios',
        'notas','creado','modificado_por','cerrado','fecha_cierre'
      ]]);
    }
  }
  return sh;
}

// ===================================================
// MANTENIMIENTO — MAQUINARIA
// ===================================================

function getAllMaquinaria(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { items: [] };
  const headers = data[0];
  const items = data.slice(1)
    .filter(row => row[0] && String(row[0]).trim() !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (val instanceof Date) {
          const y = val.getFullYear();
          const m = String(val.getMonth()+1).padStart(2,'0');
          const d = String(val.getDate()).padStart(2,'0');
          val = y+'-'+m+'-'+d;
        }
        obj[h] = (val === null || val === undefined) ? '' : val;
      });
      return obj;
    });
  return { items };
}

function saveMaquinaria(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];

  // Crear columnas nuevas si no existen
  Object.keys(data).forEach(key => {
    if (!headers.includes(key)) {
      headers.push(key);
      sheet.getRange(1, headers.length).setValue(key);
    }
  });

  const row = headers.map(h => {
    const val = data[h];
    if (val === undefined || val === null) return '';
    if (typeof val === 'object') return JSON.stringify(val);
    return val;
  });

  const ids = allData.slice(1).map(r => String(r[headers.indexOf('id')]));
  const rowIndex = ids.indexOf(String(data.id));

  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 2, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true, folio: data.folio, id: data.id };
}

function deleteMaquinaria(sheet, id) {
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const ids = allData.slice(1).map(r => String(r[headers.indexOf('id')]));
  const rowIndex = ids.indexOf(String(id));
  if (rowIndex >= 0) {
    sheet.deleteRow(rowIndex + 2);
    return { success: true, folio: data.folio, id: data.id };
  }
  return { error: 'Equipo no encontrado' };
}

// ===================================================
// MANTENIMIENTO — REGISTROS
// ===================================================

function getAllMantenimientos(sheet, equipoId) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { items: [] };
  const headers = data[0];
  let items = data.slice(1)
    .filter(row => row[0] && String(row[0]).trim() !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (val instanceof Date) {
          const y = val.getFullYear();
          const m = String(val.getMonth()+1).padStart(2,'0');
          const d = String(val.getDate()).padStart(2,'0');
          val = y+'-'+m+'-'+d;
        }
        obj[h] = (val === null || val === undefined) ? '' : val;
      });
      return obj;
    });
  if (equipoId) {
    items = items.filter(i => String(i.equipo_id) === String(equipoId));
  }
  // Ordenar por fecha descendente
  items.sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
  return { items };
}

function saveMantenimiento(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];

  Object.keys(data).forEach(key => {
    if (!headers.includes(key)) {
      headers.push(key);
      sheet.getRange(1, headers.length).setValue(key);
    }
  });

  const row = headers.map(h => {
    const val = data[h];
    if (val === undefined || val === null) return '';
    if (typeof val === 'object') return JSON.stringify(val);
    return val;
  });

  const ids = allData.slice(1).map(r => String(r[headers.indexOf('id')]));
  const rowIndex = ids.indexOf(String(data.id));

  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 2, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true, folio: data.folio, id: data.id };
}

function deleteMantenimiento(sheet, id) {
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const ids = allData.slice(1).map(r => String(r[headers.indexOf('id')]));
  const rowIndex = ids.indexOf(String(id));
  if (rowIndex >= 0) {
    sheet.deleteRow(rowIndex + 2);
    return { success: true, folio: data.folio, id: data.id };
  }
  return { error: 'Registro no encontrado' };
}

// ===================================================
// MAQUILA — control de servicios subcontratados (rolado, formado de tapas)
// Módulo separado de index.html a propósito: solo Dirección lo usa y no
// necesita cargar el payload pesado de pedidos (partidas, bitácora, log).
// ===================================================

function getAllMaquila(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { items: [] };
  const headers = data[0];
  const items = data.slice(1)
    .filter(row => row[0] && String(row[0]).trim() !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (val instanceof Date) {
          const y = val.getFullYear();
          const m = String(val.getMonth()+1).padStart(2,'0');
          const d = String(val.getDate()).padStart(2,'0');
          val = y+'-'+m+'-'+d;
        }
        obj[h] = (val === null || val === undefined) ? '' : val;
      });
      return obj;
    });
  // Más recientes primero (por fecha de envío)
  items.sort((a, b) => String(b.fecha_envio||'').localeCompare(String(a.fecha_envio||'')));
  return { items };
}

function saveMaquila(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];

  // Auto-crear columnas nuevas si faltan — mismo patrón que el resto del sistema.
  Object.keys(data).forEach(key => {
    if (!headers.includes(key)) {
      headers.push(key);
      sheet.getRange(1, headers.length).setValue(key);
    }
  });

  const map = {};
  headers.forEach((h, i) => { map[h] = i; });

  const esNuevo = !data.id;
  if (esNuevo) {
    data.id = 'MAQL-' + Date.now();
    if (!data.creado) data.creado = today();
  }

  const ids = allData.slice(1).map(r => String(r[map['id']]));
  const rowIndex = ids.indexOf(String(data.id));

  if (rowIndex >= 0) {
    // Editar: escribir SOLO las columnas presentes en "data", por nombre —
    // no se pisa el resto de la fila (permite guardados parciales por celda).
    const sheetRow = rowIndex + 2;
    Object.keys(data).forEach(key => {
      if (map[key] !== undefined) sheet.getRange(sheetRow, map[key] + 1).setValue(data[key]);
    });
  } else {
    const row = headers.map(h => {
      const val = data[h];
      if (val === undefined || val === null) return '';
      if (typeof val === 'object') return JSON.stringify(val);
      return val;
    });
    sheet.appendRow(row);
  }
  return { success: true, id: data.id };
}

function deleteMaquila(sheet, id) {
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idCol = headers.indexOf('id');
  const ids = allData.slice(1).map(r => String(r[idCol]));
  const rowIndex = ids.indexOf(String(id));
  if (rowIndex >= 0) {
    sheet.deleteRow(rowIndex + 2);
    return { success: true, id: id };
  }
  return { error: 'Registro no encontrado' };
}

// Endpoint ligero para autocompletar el folio de pedido dentro de Maquila
// SIN cargar el payload completo de pedidos (partidas, bitácora, log, los
// ~40 campos que sí carga index.html). Solo regresa folio + cliente + modelo.
// OJO: en Hoja 1 las columnas se llaman "client" y "desc" (no "cliente"/
// "modelo" — esos nombres solo existen en la hoja de Cotizaciones). Se
// traducen aquí para que el frontend de Maquila use nombres claros.
function getFoliosLigero(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { folios: [] };
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const folioIdx   = headers.indexOf('folio');
  const clienteIdx = headers.indexOf('client');
  const modeloIdx  = headers.indexOf('desc');
  if (folioIdx < 0) return { folios: [] };
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const folios = data
    .map(r => ({
      folio:  r[folioIdx],
      cliente: clienteIdx >= 0 ? r[clienteIdx] : '',
      modelo:  modeloIdx  >= 0 ? r[modeloIdx]  : ''
    }))
    .filter(f => f.folio && String(f.folio).trim() !== '')
    .reverse(); // folios más recientes primero (se agregan al final de la hoja)
  return { folios };
}

// ===================================================
// MANTENIMIENTO — ALERTAS
// Devuelve equipos con próximo mantenimiento vencido
// o en los próximos 15 días
// ===================================================

function getAlertasMantenimiento(maqSh, mantSh) {
  const maqRes  = getAllMaquinaria(maqSh);
  const mantRes = getAllMantenimientos(mantSh, null);
  const equipos = (maqRes.items  || []).filter(e => e.estatus !== 'Baja');
  const mantos  = mantRes.items || [];

  const hoy    = new Date();
  const hoyStr = hoy.getFullYear()+'-'+String(hoy.getMonth()+1).padStart(2,'0')+'-'+String(hoy.getDate()).padStart(2,'0');

  const alertas = [];
  equipos.forEach(eq => {
    // Buscar el preventivo con proxima_fecha más reciente para este equipo
    const prevs = mantos
      .filter(m => String(m.equipo_id) === String(eq.id) && m.tipo === 'Preventivo' && m.proxima_fecha)
      .sort((a, b) => (b.proxima_fecha || '').localeCompare(a.proxima_fecha || ''));

    if (prevs.length === 0) return; // sin preventivo programado — no genera alerta

    const proxFecha = prevs[0].proxima_fecha;
    const diff = Math.round((new Date(proxFecha) - hoy) / (1000*60*60*24));

    if (diff <= 15) {
      alertas.push({
        equipo_id:    eq.id,
        equipo_nombre: eq.nombre,
        area:         eq.area,
        proxima_fecha: proxFecha,
        dias:         diff, // negativo = vencido
        nivel:        diff < 0 ? 'vencido' : 'proximo'
      });
    }
  });

  // Ordenar: vencidos primero, luego por días
  alertas.sort((a, b) => a.dias - b.dias);
  return { alertas };
}


// ===================================================
// NOTIFICACIONES DE HANDOFF — FLUJO DE ETAPAS
// ===================================================

function buildOrderInfo(o) {
  const tipo    = o.tipo === 'stk' ? 'Orden STK' : 'Pedido';
  const folio   = o.folio || o.id || '—';
  const cliente = o.tipo === 'stk' ? 'Almacén interno' : (o.client || '—');
  const modelo  = o.desc || '—';
  const qty     = o.qty ? String(o.qty) + ' uds' : '—';
  const total   = (o.qty && o.value)
    ? '$' + (Number(o.qty)*Number(o.value)).toLocaleString('es-MX',{minimumFractionDigits:2}) + ' antes IVA'
    : (o.value ? '$' + Number(o.value).toLocaleString('es-MX',{minimumFractionDigits:2}) + ' antes IVA' : '—');
  const cond    = o.condicion_pago || '—';
  const entrega = o.deadline ? formatDate(o.deadline) : 'Sin fecha';
  const pctA    = o.pct_anticipo ? o.pct_anticipo + '%' : 'Sin anticipo';
  const createdBy = (o.log && o.log[0]) ? o.log[0].by : '—';
  return { tipo, folio, cliente, modelo, qty, 'Valor total': total,
           'Condición de pago': cond, 'Entrega comprometida': entrega,
           'Anticipo solicitado': pctA, 'Registrado por': createdBy };
}

function notificarPedidoNuevo(o) {
  const key = 'ntf_nuevo_' + (o.folio || o.id);
  if (wasSentToday(key)) return;
  const folio = o.folio || o.id;
  const dest  = getEmails('Contabilidad', 'Ingeniería de ventas');
  dest.push('agarza@alltansa.com');
  const pctA  = Number(o.pct_anticipo || 0);
  const cond  = String(o.condicion_pago || '').toLowerCase();
  const esCredito = cond.includes('crédito') || cond.includes('credito');
  const accionContab = (pctA > 0 && !esCredito)
    ? 'Confirmar recepción del anticipo (' + pctA + '%) para liberar ingeniería'
    : 'Sin anticipo requerido — pedido liberado automáticamente';
  const body = buildEmailBody('Nuevo pedido registrado en el Board', {
    ...buildOrderInfo(o),
    'Acción Contabilidad':         accionContab,
    'Acción Ingeniería de ventas': 'Pedido en espera de luz verde financiera para iniciar'
  });
  sendEmail(dest, 'Nuevo pedido registrado: ' + folio + ' — ' + (o.client || o.desc || ''), body);
  markSent(key);
  Logger.log('notificarPedidoNuevo enviado: ' + folio);
}

function notificarLuzVerde(o) {
  const key = 'ntf_luzv_' + (o.folio || o.id);
  if (wasSentToday(key)) return;
  const folio = o.folio || o.id;
  const dest  = getEmails('Ingeniería de ventas');
  dest.push('agarza@alltansa.com');
  const cond  = String(o.condicion_pago || '').toLowerCase();
  const esCredito = cond.includes('crédito') || cond.includes('credito');
  const motivo = esCredito
    ? 'Pedido de crédito — liberado automáticamente'
    : (Number(o.pct_anticipo || 0) === 0
        ? 'Contado con pago al final — liberado automáticamente'
        : 'Anticipo recibido (' + o.pct_anticipo + '%) confirmado por Contabilidad');
  const body = buildEmailBody('Pedido liberado — puedes iniciar ingeniería', {
    ...buildOrderInfo(o),
    'Motivo de liberación': motivo,
    'Acción requerida':     'Autorizar etapa de Ingeniería y Requisición en el Board'
  });
  sendEmail(dest, 'LUZ VERDE - Ya puedes arrancar: ' + folio + ' — ' + (o.client || o.desc || ''), body);
  markSent(key);
  Logger.log('notificarLuzVerde enviado: ' + folio);
}

function notificarIngAutorizada(o) {
  const key = 'ntf_ing_' + (o.folio || o.id);
  if (wasSentToday(key)) return;
  const folio  = o.folio || o.id;
  const dest   = getEmails('Compras', 'Producción');
  dest.push('agarza@alltansa.com');
  let stgs = o.stages;
  if (typeof stgs === 'string') { try { stgs = JSON.parse(stgs); } catch(e){ stgs = {}; } }
  const reqDone = stgs && stgs.requisicion && stgs.requisicion.done;
  const body = buildEmailBody('Ingeniería autorizada — pedido en camino a producción', {
    ...buildOrderInfo(o),
    'Estado actual':     reqDone
      ? 'Ingeniería y Requisición autorizadas — Orden de producción generada'
      : 'Ingeniería autorizada — Requisición pendiente',
    'Acción Compras':    reqDone
      ? 'Iniciar gestión de materiales cuando aparezca en tu Board'
      : 'Aviso previo — aparecerá en tu Board cuando se autorice la Requisición',
    'Acción Producción': reqDone
      ? 'Pedido disponible en tu Board para iniciar planificación'
      : 'Aviso previo — aparecerá en tu Board cuando se genere la Orden de Producción'
  });
  sendEmail(dest, 'Pedido en camino: ' + folio + ' — ' + (o.client || o.desc || ''), body);
  markSent(key);
  Logger.log('notificarIngAutorizada enviado: ' + folio);
}

// ===================================================
// SETUP AUTOMÁTICO DE TRIGGER
// Ejecuta esta función UNA VEZ manualmente desde
// Apps Script para crear el trigger diario.
// ===================================================
function setupDailyTrigger() {
  // Eliminar triggers existentes para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkAndNotify' ||
        t.getHandlerFunction() === 'enviarReporteDiario') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Crear trigger diario a las 7:00 AM hora de Monterrey
  ScriptApp.newTrigger('enviarReporteDiario')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .inTimezone('America/Monterrey')
    .create();
  Logger.log('Trigger diario configurado: enviarReporteDiario a las 7:00 AM Monterrey');
}

// ===================================================
// COTIZACIONES — CRUD COMPLETO
// Hoja: "Cotizaciones" (se crea automáticamente)
// ===================================================

const COT_COLS = [
  // Identificación
  'id','folio','fecha','cliente','cliente_id','contacto','contacto_info','origen','responsable',
  // Requerimientos técnicos (etapa 1)
  'capacidad_l','fluido','presion','material','orientacion','norma',
  'conexiones','tiene_ing','ref_ing','notas_req',
  // Estimación interna (etapa 2)
  'ing_responsable','fecha_comp','estatus_est','notas_est',
  // Partidas y campos legacy
  'partidas','servicios_extra','tipo','cantidad','capacidad','descripcion',
  // Comercial (etapa 3)
  'moneda','monto','tipo_cambio','cond_pago','vigencia','fecha_envio','notas',
  // Notas del PDF (Incluye/No incluye/comerciales A-G/especiales) — antes eran
  // efímeras (solo vivían en el PDF generado); ahora se guardan de verdad para
  // que no se reseteen a los valores por defecto cada vez que se abre el editor.
  'pdf_incluye','pdf_no_incluye','notas_comerciales_pdf','notas_tecnicas_pdf','notas_especiales_pdf',
  // Señal dedicada de "ya se generó un PDF alguna vez" — separada de
  // notas_comerciales_pdf a propósito, porque ese campo se toca en CUALQUIER
  // guardado (aunque sea con un arreglo vacío), y usarlo como señal de
  // "primera generación" generaba falsos positivos de revisión.
  'pdf_generado',
  // Número de revisión — sube solo cuando cambia algo comercial (partidas,
  // precio, notas comerciales) y el usuario confirma que es una revisión
  // nueva y no un simple ajuste menor.
  'revision',
  // Seguimiento (etapa 4)
  'estado','ultimo_contacto','proxima_accion',
  'oc_vinculada','fecha_cierre','motivo_perdida','competidor','comentarios','folio_pedido',
  // Meta
  'etapa_actual','modificado_por','modificado_en',
  // Idempotencia — ID único generado en el cliente por cada cotización NUEVA
  // que se abre. Viaja igual en reintentos de la misma sesión de guardado
  // (ej. se cayó la señal y el usuario le da "Guardar" otra vez sin saber
  // si ya se había guardado). Sirve para detectar ese reintento y NO crear
  // una fila duplicada — ver bloque de idempotencia en saveCotizacion().
  'client_op_id'
];

function getCotSheet(ss) {
  let sh = ss.getSheetByName(COT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(COT_SHEET);
    sh.getRange(1, 1, 1, COT_COLS.length).setValues([COT_COLS]);
    sh.getRange(1, 1, 1, COT_COLS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  COT_COLS.forEach(col => {
    if (!headers.includes(col)) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(col);
    }
  });
  return sh;
}

function getCotHeaderMap(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const map = {};
  headers.forEach((h, i) => { map[h] = i; });
  return map;
}

function dateToLocalCot(val) {
  if (!(val instanceof Date)) return val;
  const tz   = 'America/Monterrey';
  const opts = { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' };
  const parts = new Intl.DateTimeFormat('es-MX', opts).formatToParts(val);
  const d = {};
  parts.forEach(p => { d[p.type] = p.value; });
  return d.year + '-' + d.month + '-' + d.day;
}

function getCotizaciones(ss) {
  try {
    const sh      = getCotSheet(ss);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { cotizaciones: [] };
    const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    const map  = getCotHeaderMap(sh);
    const cotizaciones = data
      .filter(row => row[map['id']] !== '' && row[map['id']] !== null)
      .map(row => {
        const obj = {};
        COT_COLS.forEach(col => {
          let v = map[col] !== undefined ? row[map[col]] : '';
          if (col === 'revision') {
            obj[col] = String(revisionNumeroSeguro(v));
            return;
          }
          if (v instanceof Date) v = dateToLocalCot(v);
          obj[col] = (v === null || v === undefined) ? '' : String(v);
        });
        return obj;
      });
    return { cotizaciones };
  } catch(e) {
    return { cotizaciones: [], error: e.message };
  }
}

// La columna "revision" a veces hereda por accidente formato de Fecha en el
// Sheet (copiar/pegar filas, autofill arrastrando desde una columna de fecha
// vecina, etc.). Cuando eso pasa, getValues() regresa un objeto Date en vez
// del número que en realidad es, y ese Date terminaba mostrándose tal cual
// como "Rev. 1905-03-12" en el pipeline — confuso y sin sentido de negocio.
// Esta función siempre regresa un número de revisión limpio:
// - Si el valor ya es un número normal, se usa tal cual.
// - Si es un objeto Date, se recupera el número de serie que Sheets usa
//   internamente (días desde el 30-dic-1899) — así, si el valor real SÍ era
//   un número de revisión válido (ej. 1, 2, 3) que solo se veía raro por el
//   formato de celda, se recupera correctamente.
// - Si ese número recuperado no es un conteo de revisión plausible (fuera de
//   0–50), se asume que la celda quedó con basura y se regresa 0 en vez de
//   propagar un número sin sentido.
function revisionNumeroSeguro(v) {
  if (v instanceof Date) {
    const epoch = Date.UTC(1899, 11, 30);
    const soloFecha = Date.UTC(v.getFullYear(), v.getMonth(), v.getDate());
    const serial = Math.round((soloFecha - epoch) / 86400000);
    return (serial >= 0 && serial <= 50) ? serial : 0;
  }
  const n = parseInt(v, 10);
  return (isNaN(n) || n < 0 || n > 50) ? 0 : n;
}

// Registro completo ANTES de sobreescribir — necesario para poder comparar
// campo por campo y saber exactamente qué cambió en cada guardado.
function getCotizacionPorId(ss, id) {
  if (!id) return null;
  try {
    const sh = getCotSheet(ss);
    const map = getCotHeaderMap(sh);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return null;
    const idCol = (map['id'] || 0) + 1;
    const ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues().flat().map(String);
    const rowIdx = ids.indexOf(String(id));
    if (rowIdx === -1) return null;
    const row = sh.getRange(rowIdx + 2, 1, 1, sh.getLastColumn()).getValues()[0];
    const obj = {};
    COT_COLS.forEach(col => {
      let v = map[col] !== undefined ? row[map[col]] : '';
      if (v instanceof Date) v = dateToLocalCot(v);
      obj[col] = (v === null || v === undefined) ? '' : String(v);
    });
    return obj;
  } catch(e) { return null; }
}

// Campos que vale la pena auditar campo por campo en el log de movimientos,
// con su etiqueta legible. Los que NO están aquí (partidas, conexiones,
// comentarios/bitácora, metadatos) se guardan igual pero no generan detalle
// de diff — o porque ya tienen su propio registro (bitácora manual), o
// porque un diff de JSON crudo no aporta nada legible.
const COT_DIFF_FIELDS = {
  cliente:        'Cliente',
  contacto:       'Contacto',
  contacto_info:  'Contacto (info)',
  origen:         'Origen del lead',
  responsable:    'Responsable',
  capacidad_l:    'Capacidad (L)',
  fluido:         'Fluido',
  presion:        'Presión',
  material:       'Material',
  orientacion:    'Orientación',
  norma:          'Norma',
  notas_req:      'Notas de requisición',
  ing_responsable:'Responsable de ingeniería',
  estatus_est:    'Estatus de estimación',
  notas_est:      'Notas de estimación',
  moneda:         'Moneda',
  monto:          'Monto',
  cond_pago:      'Condición de pago',
  vigencia:       'Vigencia',
  notas:          'Notas comerciales (formulario)',
  revision:       'Revisión',
  estado:         'Estado',
  proxima_accion: 'Próxima acción',
  motivo_perdida: 'Motivo de pérdida',
  competidor:     'Competidor'
};

function truncarTexto_(s, max) {
  s = String(s == null ? '' : s);
  return s.length > max ? s.slice(0, max) + '…' : s;
}

// Compara el registro ANTES vs los datos nuevos y arma la lista de cambios
// legibles. Devuelve '' si no hubo ningún cambio en los campos auditados
// (para no generar ruido en el log).
//
// Las notas del PDF (incluye/no incluye/comerciales/especiales) NO se comparan
// aquí — ese diff lo arma el frontend (construirDiffPdfNotas en cotizaciones.html)
// y llega ya calculado en cot._diff_pdf_notas. Razón: la primera vez que se
// abre el editor de notas de una cotización, no hay nada guardado todavía y
// el texto que se ve en pantalla es un default calculado — comparar contra
// la columna vacía del Sheet reporta TODO como "agregado" en vez de mostrar
// solo la línea que la persona realmente cambió. Solo el frontend sabe qué
// texto se le mostró de verdad al usuario al abrir el editor.
function construirDiffCotizacion(antes, despues) {
  if (!antes) return ''; // cotización nueva — no hay "antes" que comparar
  const cambios = [];
  Object.keys(COT_DIFF_FIELDS).forEach(campo => {
    if (despues[campo] === undefined) return; // el frontend no mandó este campo en este guardado
    const label = COT_DIFF_FIELDS[campo];
    const antesLeg   = String(antes[campo] == null ? '' : antes[campo]);
    const despuesLeg = String(despues[campo] == null ? '' : despues[campo]);
    if (antesLeg === despuesLeg) return;
    cambios.push(`${label}: '${truncarTexto_(antesLeg, 100)}' -> '${truncarTexto_(despuesLeg, 100)}'`);
  });
  return cambios.join('; ');
}

function getLogCotizacion(ss, folio) {
  try {
    const sh = getOrCreateSheet(ss, LOG_COT_SHEET);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { movimientos: [] };
    const data = sh.getRange(2, 1, lastRow - 1, 5).getValues();
    const movimientos = data
      .filter(row => !folio || String(row[3]) === String(folio))
      .map(row => ({
        timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
        usuario:   row[1],
        accion:    row[2],
        folio:     row[3],
        detalle:   row[4]
      }));
    return { movimientos };
  } catch(e) { return { movimientos: [], error: e.message }; }
}

function saveCotizacion(ss, cot) {
  try {
    if (!cot) return { ok: false, error: 'Datos vacios' };
    const sh      = getCotSheet(ss);
    const map     = getCotHeaderMap(sh);
    const lastRow = sh.getLastRow();
    const keys    = Object.keys(cot).filter(k => k !== 'id');
    const soloComentario = keys.length === 1 && keys[0] === 'comentarios';

    // ── Idempotencia (protección contra duplicados por pérdida de señal) ──
    // Si este guardado llega SIN id (el cliente cree que es una cotización
    // nueva) pero trae un client_op_id que YA existe en el Sheet, es un
    // reintento de un guardado anterior cuya respuesta se perdió antes de
    // llegar al navegador (el servidor sí la procesó). En vez de insertar
    // otra fila, se regresa el folio/id que ya se había creado — el
    // guardado se vuelve seguro de reintentar cuantas veces haga falta.
    const opId = cot.client_op_id ? String(cot.client_op_id).trim() : '';
    if ((!cot.id || String(cot.id).trim() === '') && opId &&
        map['client_op_id'] !== undefined && lastRow >= 2) {
      const opCol    = map['client_op_id'] + 1;
      const ops      = sh.getRange(2, opCol, lastRow - 1, 1).getValues().flat().map(String);
      const dupIdx   = ops.indexOf(opId);
      if (dupIdx !== -1) {
        const dupRow = dupIdx + 2;
        const idExist    = String(sh.getRange(dupRow, (map['id'] || 0) + 1).getValue() || '');
        const folioExist = String(sh.getRange(dupRow, (map['folio'] || 0) + 1).getValue() || '');
        return { ok: true, id: idExist, folio: folioExist, duplicado: true };
      }
    }

    if (cot.id && String(cot.id).trim() !== '') {
      const idCol  = (map['id'] || 0) + 1;
      const ids    = lastRow > 1
        ? sh.getRange(2, idCol, lastRow - 1, 1).getValues().flat().map(String)
        : [];
      const rowIdx = ids.indexOf(String(cot.id));
      if (rowIdx === -1) return { ok: false, error: 'ID no encontrado' };
      const sheetRow = rowIdx + 2;
      if (soloComentario) {
        sh.getRange(sheetRow, (map['comentarios'] || 0) + 1).setValue(cot.comentarios);
        sh.getRange(sheetRow, (map['modificado_en'] || 0) + 1)
          .setValue(Utilities.formatDate(new Date(), 'America/Monterrey', 'yyyy-MM-dd'));
      } else {
        // ── Protección de folio ──────────────────────────────────────
        // Si este guardado (actualización) llega con folio vacío, NUNCA se
        // debe pisar el folio que ya existe en el Sheet — antes pasaba que
        // un segundo guardado dentro de la misma sesión del modal (ej. dar
        // "Generar PDF" y luego "Guardar" otra vez) mandaba folio='' porque
        // el campo del formulario nunca se refrescaba con el folio recién
        // asignado, y esta rama lo escribía tal cual, borrando el folio.
        // Ahora: folio vacío entrante -> se conserva el de la hoja, y si la
        // hoja tampoco tiene folio (caso raro/legacy), se genera uno nuevo
        // en vez de dejarlo en blanco.
        const folioActual = String(sh.getRange(sheetRow, (map['folio'] || 0) + 1).getValue() || '').trim();
        if (!cot.folio || String(cot.folio).trim() === '') {
          cot.folio = folioActual || generarFolioCot(sh, map, lastRow);
        }
        COT_COLS.forEach(col => {
          if (col === 'id') return;
          if (cot[col] !== undefined && map[col] !== undefined) {
            sh.getRange(sheetRow, map[col] + 1).setValue(cot[col]);
          }
        });
      }
    } else {
      cot.id = String(Date.now());
      if (!cot.folio || cot.folio.trim() === '') {
        cot.folio = generarFolioCot(sh, map, lastRow);
      }
      // Escribir por NOMBRE de columna (usando el mapa de encabezados), no por
      // posición fija. Esto es crítico: el orden de COT_COLS en el código ya no
      // coincide con el orden físico real de columnas en el Sheet (los campos
      // nuevos de las 4 etapas se agregaron al final, no en su posición "lógica").
      // Escribir posicionalmente causaba que los valores cayeran en columnas
      // con el encabezado equivocado. Ahora cada valor va a SU columna, sin
      // importar en qué orden físico esté.
      const totalCols = sh.getLastColumn();
      const newRow = new Array(totalCols).fill('');
      COT_COLS.forEach(col => {
        if (map[col] !== undefined && cot[col] !== undefined) {
          newRow[map[col]] = cot[col];
        }
      });
      sh.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    }
    return { ok: true, folio: cot.folio, id: cot.id };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function generarFolioCot(sh, map, lastRow) {
  const tz   = 'America/Monterrey';
  const now  = new Date();
  const yy   = Utilities.formatDate(now, tz, 'yy');
  const mm   = Utilities.formatDate(now, tz, 'MM');
  const pref = 'COT-' + yy + mm + '-';

  let next = 1;
  if (lastRow >= 2) {
    const folioCol = (map['folio'] || 0) + 1;
    const folios   = sh.getRange(2, folioCol, lastRow - 1, 1).getValues().flat().map(String);
    const nums     = folios
      .filter(f => f.startsWith(pref))
      .map(f => parseInt(f.replace(pref, '')) || 0);
    if (nums.length > 0) next = Math.max(...nums) + 1;
  }

  // PISO TEMPORAL — solo julio 2026 (COT-2607-xx): ya se usaron folios 001-019
  // fuera del sistema, por lo que el primer folio automático debe ser 020.
  // Este bloque deja de aplicar solo por sí mismo al cambiar de mes/año
  // (el prefijo ya no coincidirá). Se puede eliminar después de julio 2026.
  if (pref === 'COT-2607-') {
    next = Math.max(next, 20);
  }

  return pref + String(next).padStart(3, '0');
}

// ===================================================
// CLIENTES — CRUD
// Hoja: "Clientes" (se crea automáticamente)
// ===================================================

const CLI_SHEET = 'Clientes';
const CLI_COLS  = ['id','empresa','contacto','ciudad','email','telefono','estatus','origen','notas','creado','modificado_por'];

function getCliSheet(ss) {
  let sh = ss.getSheetByName(CLI_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CLI_SHEET);
    sh.getRange(1,1,1,CLI_COLS.length).setValues([CLI_COLS]);
    sh.getRange(1,1,1,CLI_COLS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  CLI_COLS.forEach(col => { if (!headers.includes(col)) sh.getRange(1, sh.getLastColumn()+1).setValue(col); });
  return sh;
}

function getCliHeaderMap(sh) {
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const m = {}; h.forEach((v,i) => { m[v]=i; }); return m;
}

function getClientes(ss) {
  try {
    const sh = getCliSheet(ss);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { clientes: [] };
    const data = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
    const map  = getCliHeaderMap(sh);
    const clientes = data
      .filter(row => row[map['id']] !== '' && row[map['id']] !== null)
      .map(row => {
        const obj = {};
        CLI_COLS.forEach(col => {
          let v = map[col] !== undefined ? row[map[col]] : '';
          if (v instanceof Date) v = dateToLocalCot(v);
          obj[col] = (v === null || v === undefined) ? '' : String(v);
        });
        return obj;
      });
    return { clientes };
  } catch(e) { return { clientes: [], error: e.message }; }
}

function saveCliente(ss, cli) {
  try {
    if (!cli) return { ok: false, error: 'Datos vacios' };
    const sh  = getCliSheet(ss);
    const map = getCliHeaderMap(sh);
    const lastRow = sh.getLastRow();
    if (cli.id && String(cli.id).trim() !== '') {
      const ids = lastRow > 1 ? sh.getRange(2,(map['id']||0)+1,lastRow-1,1).getValues().flat().map(String) : [];
      const rowIdx = ids.indexOf(String(cli.id));
      if (rowIdx === -1) return { ok: false, error: 'ID no encontrado' };
      CLI_COLS.forEach(col => { if (col==='id') return; if (cli[col]!==undefined && map[col]!==undefined) sh.getRange(rowIdx+2, map[col]+1).setValue(cli[col]); });
    } else {
      cli.id = String(Date.now());
      const totalCols = sh.getLastColumn();
      const newRow = new Array(totalCols).fill('');
      CLI_COLS.forEach(col => { if (map[col] !== undefined && cli[col] !== undefined) newRow[map[col]] = cli[col]; });
      sh.getRange(lastRow+1,1,1,newRow.length).setValues([newRow]);
    }
    return { ok: true, id: cli.id };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ===================================================
// CONTABILIDAD — FACTURAS Y ABONOS
// ===================================================

const FACT_COLS = [
  'id','tipo','folio_pedido','cliente','num_factura','fecha_factura',
  'monto_sin_iva','iva','monto_total','tipo_factura','pct_pedido',
  'dias_credito','fecha_vencimiento','estatus_iva','fecha_pago_iva',
  'notas','creado_por','createdAt'
];

const ABONO_COLS = [
  'id','folio_pedido','fecha_abono','monto','referencia','forma_pago','notas','creado_por','createdAt'
];

function initFactSheet(sh) {
  if (sh.getLastRow() === 0) {
    const headers = FACT_COLS.concat(ABONO_COLS.map(c => 'ab_'+c));
    sh.getRange(1,1,1,FACT_COLS.length).setValues([FACT_COLS]);
    sh.getRange(1,1,1,FACT_COLS.length).setFontWeight('bold');
  }
  FACT_COLS.forEach(col => {
    const headers = sh.getRange(1,1,1,sh.getLastColumn()>0?sh.getLastColumn():1).getValues()[0];
    if (!headers.includes(col)) sh.getRange(1, sh.getLastColumn()+1).setValue(col);
  });
}

function initAbonoSheet(ss) {
  let sh = ss.getSheetByName('Abonos');
  if (!sh) { sh = ss.insertSheet('Abonos'); }
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,ABONO_COLS.length).setValues([ABONO_COLS]);
    sh.getRange(1,1,1,ABONO_COLS.length).setFontWeight('bold');
  }
  ABONO_COLS.forEach(col => {
    const headers = sh.getRange(1,1,1,sh.getLastColumn()>0?sh.getLastColumn():1).getValues()[0];
    if (!headers.includes(col)) sh.getRange(1, sh.getLastColumn()+1).setValue(col);
  });
  return sh;
}

function getAllFacturas(sh) {
  try {
    const cacheado = cacheGet(CACHE_KEY_FACTURAS);
    if (cacheado) return cacheado;

    initFactSheet(sh);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abonoSh = initAbonoSheet(ss);
    const lastRow = sh.getLastRow();
    const facturas = [];
    if (lastRow > 1) {
      const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      const data = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
      data.forEach(row => {
        const obj = {};
        headers.forEach((h,i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
        facturas.push(obj);
      });
    }
    // Abonos
    const abonos = [];
    const aLastRow = abonoSh.getLastRow();
    if (aLastRow > 1) {
      const aHeaders = abonoSh.getRange(1,1,1,abonoSh.getLastColumn()).getValues()[0];
      const aData = abonoSh.getRange(2,1,aLastRow-1,abonoSh.getLastColumn()).getValues();
      aData.forEach(row => {
        const obj = {};
        aHeaders.forEach((h,i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
        abonos.push(obj);
      });
    }
    const resultado = { facturas, abonos };
    cachePut(CACHE_KEY_FACTURAS, resultado);
    return resultado;
  } catch(e) { return { error: e.message }; }
}

function saveFactura(sh, data) {
  try {
    initFactSheet(sh);
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h,i) => { map[h] = i; });
    const now = new Date().toISOString();
    if (!data.id) data.id = 'FAC-' + Date.now();
    data.createdAt = data.createdAt || now;
    // buscar fila existente
    const lastRow = sh.getLastRow();
    let rowIdx = -1;
    if (lastRow > 1) {
      const ids = sh.getRange(2,1,lastRow-1,1).getValues();
      ids.forEach((r,i) => { if (r[0] === data.id) rowIdx = i; });
    }
    if (rowIdx >= 0) {
      // Editar: escribir SOLO las columnas presentes en data, por nombre.
      const sheetRow = rowIdx + 2;
      FACT_COLS.forEach(col => {
        if (data[col] !== undefined && map[col] !== undefined) {
          sh.getRange(sheetRow, map[col] + 1).setValue(data[col]);
        }
      });
    } else {
      // Fila nueva: por nombre de columna, no por posición fija (ver nota en
      // saveCotizacion sobre por qué el orden de *_COLS puede no coincidir
      // con el orden físico real de columnas en la hoja).
      const totalCols = sh.getLastColumn();
      const newRow = new Array(totalCols).fill('');
      FACT_COLS.forEach(col => {
        if (map[col] !== undefined && data[col] !== undefined) {
          newRow[map[col]] = data[col];
        }
      });
      sh.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    }
    return { ok: true, id: data.id };
  } catch(e) { return { ok: false, error: e.message }; }
}

function deleteFactura(sh, id) {
  try {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'No hay registros' };
    const ids = sh.getRange(2,1,lastRow-1,1).getValues();
    let rowIdx = -1;
    ids.forEach((r,i) => { if (r[0] === id) rowIdx = i; });
    if (rowIdx < 0) return { ok: false, error: 'No encontrado' };
    sh.deleteRow(rowIdx+2);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.message }; }
}

function saveAbono(factSh, data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = initAbonoSheet(ss);
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h,i) => { map[h] = i; });
    const now = new Date().toISOString();
    if (!data.id) data.id = 'AB-' + Date.now();
    data.createdAt = data.createdAt || now;
    const lastRow = sh.getLastRow();
    let rowIdx = -1;
    if (lastRow > 1) {
      const ids = sh.getRange(2,1,lastRow-1,1).getValues();
      ids.forEach((r,i) => { if (r[0] === data.id) rowIdx = i; });
    }
    if (rowIdx >= 0) {
      // Editar: escribir SOLO las columnas presentes en data, por nombre.
      const sheetRow = rowIdx + 2;
      ABONO_COLS.forEach(col => {
        if (data[col] !== undefined && map[col] !== undefined) {
          sh.getRange(sheetRow, map[col] + 1).setValue(data[col]);
        }
      });
    } else {
      // Fila nueva: por nombre de columna, no por posición fija.
      const totalCols = sh.getLastColumn();
      const newRow = new Array(totalCols).fill('');
      ABONO_COLS.forEach(col => {
        if (map[col] !== undefined && data[col] !== undefined) {
          newRow[map[col]] = data[col];
        }
      });
      sh.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    }
    return { ok: true, id: data.id };
  } catch(e) { return { ok: false, error: e.message }; }
}

function deleteAbono(factSh, id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = initAbonoSheet(ss);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'No hay registros' };
    const ids = sh.getRange(2,1,lastRow-1,1).getValues();
    let rowIdx = -1;
    ids.forEach((r,i) => { if (r[0] === id) rowIdx = i; });
    if (rowIdx < 0) return { ok: false, error: 'No encontrado' };
    sh.deleteRow(rowIdx+2);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ===================================================
// LOG DE AUDITORÍA — CONTABILIDAD
// ===================================================

function logContab(sh, usuario, accion, folio, detalle) {
  try {
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,5).setValues([['timestamp','usuario','accion','folio_pedido','detalle']]);
      sh.getRange(1,1,1,5).setFontWeight('bold');
    }
    sh.appendRow([new Date().toISOString(), usuario, accion, folio, detalle]);
  } catch(e) { /* no interrumpir flujo principal */ }
}

// ===================================================
// CIERRE DE PEDIDO DESDE CONTABILIDAD
// ===================================================

function cerrarPedidoDesdeContab(sheet, logSh, orderId, by) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'No hay pedidos' };
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const idCol = headers.indexOf('id');
    const stagesCol = headers.indexOf('stages');
    const finiqCol = headers.indexOf('finiquito');
    const finiqDateCol = headers.indexOf('finiquito_date');
    const logCol = headers.indexOf('log');
    const prodStagesCol = headers.indexOf('prod_stages');
    if (idCol < 0 || stagesCol < 0) return { ok: false, error: 'Columnas no encontradas' };

    const data = sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues();
    let rowIdx = -1;
    data.forEach((row, i) => { if (row[idCol] === orderId) rowIdx = i; });
    if (rowIdx < 0) return { ok: false, error: 'Pedido no encontrado' };

    const row = data[rowIdx];
    let stages = {};
    try { stages = JSON.parse(row[stagesCol] || '{}'); } catch(e) {}
    stages.cerrado = stages.cerrado || { done: false, date: null, by: null };

    if (stages.cerrado.done) return { ok: true, msg: 'Ya estaba cerrado' };

    // Candado de embarque: liquidar el saldo en Contabilidad no debe cerrar el
    // pedido si el embarque todavía no está confirmado — mismo requisito que ya
    // exige intentarCierreAutomatico() en el frontend (index.html) para el cierre
    // desde Producción/Admin. Antes este cierre solo miraba el saldo, dejando
    // pedidos marcados "Cerrado" con embarque pendiente (bug real detectado jul-2026).
    let prodStages = {};
    if (prodStagesCol >= 0) { try { prodStages = JSON.parse(row[prodStagesCol] || '{}'); } catch(e) {} }
    const embarqueDone = !!(prodStages && prodStages.embarque && prodStages.embarque.done);
    if (!embarqueDone) return { ok: false, error: 'Saldo liquidado, pero el embarque aún no está confirmado. No se cierra automáticamente.' };

    const now = new Date();
    const dateStr = now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0')+'-'+String(now.getDate()).padStart(2,'0');
    stages.cerrado = { done: true, date: dateStr, by: by };

    // Actualizar stages en el sheet
    const stagesRange = sheet.getRange(rowIdx+2, stagesCol+1);
    stagesRange.setValue(JSON.stringify(stages));

    // Actualizar finiquito
    if (finiqCol >= 0) sheet.getRange(rowIdx+2, finiqCol+1).setValue('Pagado');
    if (finiqDateCol >= 0) sheet.getRange(rowIdx+2, finiqDateCol+1).setValue(dateStr);

    // Agregar log
    if (logCol >= 0) {
      let log = [];
      try { log = JSON.parse(row[logCol] || '[]'); } catch(e) {}
      const ts = now.toLocaleDateString('es-MX', {day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
      log.push({ by: by, msg: 'Pedido cerrado automáticamente — saldo liquidado en Contabilidad', ts });
      if (log.length > 30) log = log.slice(-30);
      sheet.getRange(rowIdx+2, logCol+1).setValue(JSON.stringify(log));
    }

    logContab(logSh, by, 'cierrePedido', row[headers.indexOf('folio')] || orderId, 'Cierre automático por liquidación');
    return { ok: true };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ===================================================
// ALERTA DIARIA — FACTURAS VENCIDAS (trigger diario)
// ===================================================

function alertarFacturasVencidas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const factSh = getOrCreateSheet(ss, FACT_SHEET);
    const abonoSh = ss.getSheetByName('Abonos');
    if (!factSh || factSh.getLastRow() < 2) return;

    const headers = factSh.getRange(1,1,1,factSh.getLastColumn()).getValues()[0];
    const data = factSh.getRange(2,1,factSh.getLastRow()-1,factSh.getLastColumn()).getValues();

    // Calcular abonos por folio
    const abonosPorFolio = {};
    if (abonoSh && abonoSh.getLastRow() > 1) {
      const aHeaders = abonoSh.getRange(1,1,1,abonoSh.getLastColumn()).getValues()[0];
      const aData = abonoSh.getRange(2,1,abonoSh.getLastRow()-1,abonoSh.getLastColumn()).getValues();
      const aFolioCol = aHeaders.indexOf('folio_pedido');
      const aMontoCol = aHeaders.indexOf('monto');
      aData.forEach(row => {
        const f = row[aFolioCol]; const m = parseFloat(row[aMontoCol])||0;
        abonosPorFolio[f] = (abonosPorFolio[f]||0) + m;
      });
    }

    const hoy = new Date(); hoy.setHours(0,0,0,0);
    const vencidas = [];

    const folioCol = headers.indexOf('folio_pedido');
    const clienteCol = headers.indexOf('cliente');
    const numFactCol = headers.indexOf('num_factura');
    const totalCol = headers.indexOf('monto_total');
    const vencCol = headers.indexOf('fecha_vencimiento');

    data.forEach(row => {
      const vencStr = row[vencCol];
      if (!vencStr) return;
      const vencDate = new Date(vencStr); vencDate.setHours(0,0,0,0);
      if (vencDate >= hoy) return;
      const folio = row[folioCol];
      const cobrado = abonosPorFolio[folio] || 0;
      const total = parseFloat(row[totalCol]) || 0;
      const saldo = total - cobrado;
      if (saldo <= 0.01) return;
      const diasVencido = Math.floor((hoy - vencDate) / 86400000);
      vencidas.push({ folio, cliente: row[clienteCol], factura: row[numFactCol], saldo, diasVencido });
    });

    if (vencidas.length === 0) return;

    const tabla = vencidas.map(v =>
      `• ${v.folio} — ${v.cliente} — FAC-${v.factura} — Saldo: $${v.saldo.toLocaleString('es-MX',{minimumFractionDigits:2})} — Vencida hace ${v.diasVencido} días`
    ).join('\n');

    const asunto = `ALLTANSA: ${vencidas.length} factura(s) vencida(s) sin cobrar`;
    const cuerpo = `Estimados,\n\nSe detectaron las siguientes facturas vencidas con saldo pendiente:\n\n${tabla}\n\nFavor de dar seguimiento a la cobranza.\n\nALLTANSA Board — Módulo de Contabilidad`;

    const destinatarios = ['agarza@alltansa.com', 'contabilidad@alltansa.com'];
    destinatarios.forEach(email => {
      try { MailApp.sendEmail(email, asunto, cuerpo); } catch(e) {}
    });
  } catch(e) { Logger.log('Error alertarFacturasVencidas: ' + e.message); }
}
