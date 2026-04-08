// ============================================================
// NORTE VERDE · Sistema de Stock · Apps Script
// Planilla: una sola hoja de cálculo con múltiples pestañas
// ============================================================

// ── CONFIGURACIÓN ───────────────────────────────────────────
const CONFIG = {
  SHEETS: {
    MAESTRO:    'MAESTRO_SKU',
    MOVIMIENTOS:'MOVIMIENTOS',
    STOCK:      'STOCK_CALCULADO',
    VENTAS:     'VENTAS_MERCADITO',
  },
  NODOS: ['BODEGA','BARRA','MERCADITO'],
  TIPOS: ['RECEPCION','TRASPASO','VENTA','AJUSTE','DEVOLUCION'],
  CATEGORIAS: ['VINO','PISCO','ACEITE_OLIVA','OTRO'],
};

// ── SETUP INICIAL ────────────────────────────────────────────
// Ejecutar UNA VEZ para crear todas las hojas con sus cabeceras

function setupPlanilla() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _crearHoja(ss, CONFIG.SHEETS.MAESTRO, [
    'SKU_ID','NOMBRE','CATEGORIA','PRODUCTOR','VALLE',
    'PRECIO_COSTO','PRECIO_VENTA','UNIDAD','ACTIVO'
  ]);

  _crearHoja(ss, CONFIG.SHEETS.MOVIMIENTOS, [
    'MOV_ID','TIMESTAMP','TIPO_MOV','SKU_ID','CANTIDAD',
    'ORIGEN','DESTINO','PRECIO_UNITARIO','USUARIO','NOTAS'
  ]);

  _crearHoja(ss, CONFIG.SHEETS.STOCK, [
    'SKU_ID','NOMBRE','CATEGORIA','PRODUCTOR','VALLE',
    'STOCK_BODEGA','STOCK_BARRA','STOCK_MERCADITO',
    'EN_TRANSITO','TOTAL_SISTEMA','ACTUALIZADO'
  ]);

  _crearHoja(ss, CONFIG.SHEETS.VENTAS, [
    'TIMESTAMP','SKU_ID','NOMBRE','CATEGORIA',
    'CANTIDAD','PRECIO_UNITARIO','TOTAL','USUARIO'
  ]);

  SpreadsheetApp.getUi().alert('✅ Planilla Norte Verde lista. Crea los formularios con crearFormularios().');
}

function _crearHoja(ss, nombre, cabeceras) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
  } else {
    hoja.clear();
  }
  hoja.getRange(1, 1, 1, cabeceras.length)
    .setValues([cabeceras])
    .setFontWeight('bold')
    .setBackground('#1a3a2a')
    .setFontColor('#ffffff');
  hoja.setFrozenRows(1);
  return hoja;
}

// ── FORMULARIOS ──────────────────────────────────────────────
// Crea los tres formularios y guarda sus IDs en PropertiesService

function crearFormularios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  // 1. Formulario de RECEPCIÓN en bodega
  const fRecep = FormApp.create('NV · Recepción Bodega');
  fRecep.setDescription('Norte Verde · Registra mercadería que ingresa a Bodega Central');
  fRecep.addListItem()
    .setTitle('SKU')
    .setChoiceValues(_getSKUs())
    .setRequired(true);
  fRecep.addTextItem().setTitle('Cantidad (botellas)').setRequired(true);
  fRecep.addListItem()
    .setTitle('Origen')
    .setChoiceValues(['EXTERNO', ...(_getTransitos())])
    .setRequired(true);
  fRecep.addTextItem().setTitle('Proveedor / notas').setRequired(false);
  fRecep.addTextItem().setTitle('Tu nombre').setRequired(true);
  props.setProperty('FORM_RECEPCION_ID', fRecep.getId());

  // 2. Formulario de TRASPASO entre nodos
  const fTrasp = FormApp.create('NV · Traspaso entre puntos');
  fTrasp.setDescription('Norte Verde · Mueve botellas entre Bodega, Barra y Mercadito');
  fTrasp.addListItem()
    .setTitle('SKU')
    .setChoiceValues(_getSKUs())
    .setRequired(true);
  fTrasp.addTextItem().setTitle('Cantidad (botellas)').setRequired(true);
  fTrasp.addListItem()
    .setTitle('Desde')
    .setChoiceValues(['BODEGA','BARRA','MERCADITO'])
    .setRequired(true);
  fTrasp.addListItem()
    .setTitle('Hacia')
    .setChoiceValues(['BODEGA','BARRA','MERCADITO'])
    .setRequired(true);
  fTrasp.addTextItem().setTitle('Tu nombre').setRequired(true);
  fTrasp.addTextItem().setTitle('Notas').setRequired(false);
  props.setProperty('FORM_TRASPASO_ID', fTrasp.getId());

  // 3. Formulario de VENTA en Mercadito
  const fVenta = FormApp.create('NV · Venta Mercadito');
  fVenta.setDescription('Norte Verde · Registro de venta en El Mercadito');
  fVenta.addListItem()
    .setTitle('Producto')
    .setChoiceValues(_getSKUsConPrecio())
    .setRequired(true);
  fVenta.addTextItem().setTitle('Cantidad').setRequired(true);
  fVenta.addTextItem().setTitle('Precio cobrado ($)').setRequired(true);
  fVenta.addTextItem().setTitle('Tu nombre').setRequired(true);
  props.setProperty('FORM_VENTA_ID', fVenta.getId());

  // Guarda URLs
  const urls = {
    recepcion: fRecep.getPublishedUrl(),
    traspaso:  fTrasp.getPublishedUrl(),
    venta:     fVenta.getPublishedUrl(),
    recepcionEdit: fRecep.getEditUrl(),
    traspasoEdit:  fTrasp.getEditUrl(),
    ventaEdit:     fVenta.getEditUrl(),
  };
  props.setProperty('FORM_URLS', JSON.stringify(urls));

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '✅ Formularios creados\n\n' +
    'Recepción:\n' + urls.recepcion + '\n\n' +
    'Traspaso:\n' + urls.traspaso + '\n\n' +
    'Venta Mercadito:\n' + urls.venta + '\n\n' +
    'Guarda estas URLs. Instala triggers con instalarTriggers().'
  );
}

function _getSKUs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.MAESTRO);
  if (!hoja) return ['(sin SKUs)'];
  const datos = hoja.getDataRange().getValues();
  const filas = datos.slice(1).filter(r => {
    if (!r[0]) return false; // fila vacía
    const activo = String(r[8]).trim().toUpperCase();
    return activo === 'TRUE' || activo === 'VERDADERO' || activo === '1' || r[8] === true;
  });
  if (filas.length === 0) return ['(sin SKUs activos)'];
  return filas.map(r => r[0] + ' · ' + r[1]);
}

function _getSKUsConPrecio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.MAESTRO);
  if (!hoja) return ['(sin SKUs)'];
  const datos = hoja.getDataRange().getValues();
  const filas = datos.slice(1).filter(r => {
    if (!r[0]) return false;
    const activo = String(r[8]).trim().toUpperCase();
    return activo === 'TRUE' || activo === 'VERDADERO' || activo === '1' || r[8] === true;
  });
  if (filas.length === 0) return ['(sin SKUs activos)'];
  return filas.map(r => r[0] + ' · ' + r[1] + ' $' + r[6]);
}

function _getTransitos() {
  // Devuelve nodos TRANSITO_xxx que existen en movimientos
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.MOVIMIENTOS);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getRange(2, 6, hoja.getLastRow()-1, 1).getValues();
  const transitos = [...new Set(datos.flat().filter(v => String(v).startsWith('TRANSITO_')))];
  return transitos;
}

// ── TRIGGERS ─────────────────────────────────────────────────

function instalarTriggers() {
  // Borra triggers anteriores para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  const props = PropertiesService.getScriptProperties();
  const recepId = props.getProperty('FORM_RECEPCION_ID');
  const traspId = props.getProperty('FORM_TRASPASO_ID');
  const ventaId = props.getProperty('FORM_VENTA_ID');

  if (recepId) {
    ScriptApp.newTrigger('onRecepcion')
      .forForm(recepId).onFormSubmit().create();
  }
  if (traspId) {
    ScriptApp.newTrigger('onTraspaso')
      .forForm(traspId).onFormSubmit().create();
  }
  if (ventaId) {
    ScriptApp.newTrigger('onVenta')
      .forForm(ventaId).onFormSubmit().create();
  }

  // Trigger de recálculo automático cada 5 minutos
  ScriptApp.newTrigger('recalcularStock')
    .timeBased().everyMinutes(5).create();

  SpreadsheetApp.getUi().alert('✅ Triggers instalados. El sistema está activo.');
}

// ── HANDLERS DE FORMULARIOS ──────────────────────────────────

function onRecepcion(e) {
  try {
    const r = e.response.getItemResponses();
    const skuRaw   = r[0].getResponse(); // "NV-VT-001 · Nombre"
    const skuId    = skuRaw.split(' · ')[0].trim();
    const cantidad = parseInt(r[1].getResponse(), 10);
    const origen   = r[2].getResponse();
    const notas    = r[3] ? r[3].getResponse() : '';
    const usuario  = r[4] ? r[4].getResponse() : 'sistema';

    // Si origen es EXTERNO → recepción directa a bodega
    // Si origen es TRANSITO_xxx → cierre de tránsito → bodega
    const destino = 'BODEGA';
    const tipo    = origen.startsWith('TRANSITO_') ? 'TRASPASO' : 'RECEPCION';

    _registrarMovimiento(tipo, skuId, cantidad, origen, destino, 0, usuario, notas);
    recalcularStock();
  } catch(err) {
    Logger.log('Error onRecepcion: ' + err.message);
  }
}

function onTraspaso(e) {
  try {
    const r = e.response.getItemResponses();
    const skuId    = r[0].getResponse().split(' · ')[0].trim();
    const cantidad = parseInt(r[1].getResponse(), 10);
    const origen   = r[2].getResponse();
    const destino  = r[3].getResponse();
    const usuario  = r[4].getResponse();
    const notas    = r[5] ? r[5].getResponse() : '';

    if (origen === destino) {
      Logger.log('Traspaso ignorado: origen = destino = ' + origen);
      return;
    }

    _registrarMovimiento('TRASPASO', skuId, cantidad, origen, destino, 0, usuario, notas);
    recalcularStock();
  } catch(err) {
    Logger.log('Error onTraspaso: ' + err.message);
  }
}

function onVenta(e) {
  try {
    const r = e.response.getItemResponses();
    const skuRaw   = r[0].getResponse();
    const skuId    = skuRaw.split(' · ')[0].trim();
    const cantidad = parseInt(r[1].getResponse(), 10);
    const precio   = parseFloat(r[2].getResponse().replace(/[$.,\s]/g,''));
    const usuario  = r[3].getResponse();

    _registrarMovimiento('VENTA', skuId, cantidad, 'MERCADITO', 'EXTERNO', precio, usuario, '');
    _registrarVentaMercadito(skuId, cantidad, precio, usuario);
    recalcularStock();
  } catch(err) {
    Logger.log('Error onVenta: ' + err.message);
  }
}

// ── ESCRITURA EN MOVIMIENTOS ─────────────────────────────────

function _registrarMovimiento(tipo, skuId, cantidad, origen, destino, precio, usuario, notas) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.MOVIMIENTOS);
  const ts    = new Date();
  const movId = Utilities.formatDate(ts, 'America/Santiago', 'yyyyMMddHHmmss') +
                Math.random().toString(36).substr(2,4).toUpperCase();

  hoja.appendRow([
    movId, ts, tipo, skuId, cantidad,
    origen, destino, precio || '', usuario, notas || ''
  ]);
}

function _registrarVentaMercadito(skuId, cantidad, precio, usuario) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hMov  = ss.getSheetByName(CONFIG.SHEETS.MOVIMIENTOS);
  const hVent = ss.getSheetByName(CONFIG.SHEETS.VENTAS);
  const maestro = _getMaestroMap();
  const sku     = maestro[skuId] || {};
  const ts      = new Date();

  hVent.appendRow([
    ts, skuId, sku.nombre || '', sku.categoria || '',
    cantidad, precio, cantidad * precio, usuario
  ]);
}

function _getMaestroMap() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return {};
  const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, 9).getValues();
  const map = {};
  datos.forEach(r => {
    if (r[0]) map[r[0]] = {
      nombre: r[1], categoria: r[2], productor: r[3],
      valle: r[4], precioCosto: r[5], precioVenta: r[6]
    };
  });
  return map;
}

// ── RECÁLCULO DE STOCK ───────────────────────────────────────
// Lee MOVIMIENTOS completo y recalcula STOCK_CALCULADO

function recalcularStock() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hMov    = ss.getSheetByName(CONFIG.SHEETS.MOVIMIENTOS);
  const hStock  = ss.getSheetByName(CONFIG.SHEETS.STOCK);
  const maestro = _getMaestroMap();

  if (!hMov || hMov.getLastRow() < 2) return;

  const movs = hMov.getRange(2, 1, hMov.getLastRow()-1, 10).getValues();

  // Acumuladores por SKU y nodo
  const acc = {}; // acc[skuId][nodo] = cantidad

  const _add = (skuId, nodo, delta) => {
    if (!acc[skuId]) acc[skuId] = { BODEGA:0, BARRA:0, MERCADITO:0, TRANSITO:0 };
    if (nodo === 'BODEGA' || nodo === 'BARRA' || nodo === 'MERCADITO') {
      acc[skuId][nodo] += delta;
    } else if (String(nodo).startsWith('TRANSITO_')) {
      acc[skuId]['TRANSITO'] += delta;
    }
  };

  movs.forEach(m => {
    const skuId   = m[3];
    const cant    = Number(m[4]) || 0;
    const origen  = String(m[5]);
    const destino = String(m[6]);
    if (!skuId || !cant) return;

    _add(skuId, destino, +cant);
    _add(skuId, origen,  -cant);
  });

  // Reconstruye hoja STOCK_CALCULADO
  const ts   = new Date();
  const rows = [];
  Object.keys(acc).sort().forEach(skuId => {
    const a   = acc[skuId];
    const sku = maestro[skuId] || {};
    const total = (a.BODEGA||0) + (a.BARRA||0) + (a.MERCADITO||0) + (a.TRANSITO||0);
    rows.push([
      skuId,
      sku.nombre    || '',
      sku.categoria || '',
      sku.productor || '',
      sku.valle     || '',
      Math.max(0, a.BODEGA    || 0),
      Math.max(0, a.BARRA     || 0),
      Math.max(0, a.MERCADITO || 0),
      Math.max(0, a.TRANSITO  || 0),
      Math.max(0, total),
      ts
    ]);
  });

  // Limpia y escribe de una vez (eficiente)
  if (hStock.getLastRow() > 1) {
    hStock.getRange(2, 1, hStock.getLastRow()-1, 11).clearContent();
  }
  if (rows.length > 0) {
    hStock.getRange(2, 1, rows.length, 11).setValues(rows);
  }

  Logger.log('Stock recalculado: ' + rows.length + ' SKUs · ' + ts);
}

// ── REGISTRO MANUAL DE TRÁNSITO ──────────────────────────────
// Para declarar mercadería que salió del proveedor pero no llegó
// Ejecutar desde el menú o llamar directamente

function registrarTransito(skuId, cantidad, nombreProveedor, usuario, notas) {
  const origenTransito = 'TRANSITO_' + nombreProveedor.replace(/\s/g,'_').toUpperCase();
  _registrarMovimiento('RECEPCION', skuId, cantidad, 'EXTERNO', origenTransito, 0, usuario || 'admin', notas || '');
  recalcularStock();
  Logger.log('Tránsito registrado: ' + cantidad + ' ' + skuId + ' → ' + origenTransito);
}

// ── EXPORTAR URLs de FORMULARIOS ─────────────────────────────

function verURLsFormularios() {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty('FORM_URLS');
  if (!raw) {
    SpreadsheetApp.getUi().alert('No se encontraron formularios. Ejecuta crearFormularios() primero.');
    return;
  }
  const urls = JSON.parse(raw);
  SpreadsheetApp.getUi().alert(
    'RECEPCIÓN BODEGA:\n' + urls.recepcion +
    '\n\nTRASPASO:\n' + urls.traspaso +
    '\n\nVENTA MERCADITO:\n' + urls.venta
  );
}

// ── MENÚ PERSONALIZADO ───────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🍷 Norte Verde')
    .addItem('1. Setup inicial (primera vez)', 'setupPlanilla')
    .addItem('2. Crear formularios', 'crearFormularios')
    .addItem('3. Instalar triggers', 'instalarTriggers')
    .addSeparator()
    .addItem('Recalcular stock ahora', 'recalcularStock')
    .addItem('Ver URLs de formularios', 'verURLsFormularios')
    .addToUi();
}

// ── API ENDPOINT PARA EL DASHBOARD ───────────────────────────
// Deploy como Web App (ejecutar como yo, acceso a cualquiera)

function doGet(e) {
  const accion = e.parameter.accion || 'stock';

  let data;
  if (accion === 'stock') {
    data = _getStockData();
  } else if (accion === 'ventas') {
    data = _getVentasData();
  } else if (accion === 'movimientos') {
    data = _getMovimientosRecientes();
  } else {
    data = { error: 'accion no reconocida' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function _getStockData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.STOCK);
  if (!hoja || hoja.getLastRow() < 2) return { items: [], actualizado: new Date() };

  const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, 11).getValues();
  const items = datos
    .filter(r => r[0])
    .map(r => ({
      sku:        r[0],
      nombre:     r[1],
      categoria:  r[2],
      productor:  r[3],
      valle:      r[4],
      bodega:     r[5],
      barra:      r[6],
      mercadito:  r[7],
      transito:   r[8],
      total:      r[9],
      actualizado:r[10],
    }));

  return { items, actualizado: new Date() };
}

function _getVentasData() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.VENTAS);
  if (!hoja || hoja.getLastRow() < 2) return { ventas: [], totalVentas: 0 };

  const datos  = hoja.getRange(2, 1, hoja.getLastRow()-1, 8).getValues();
  const ventas = datos
    .filter(r => r[0])
    .map(r => ({
      timestamp:      r[0],
      sku:            r[1],
      nombre:         r[2],
      categoria:      r[3],
      cantidad:       r[4],
      precioUnitario: r[5],
      total:          r[6],
      usuario:        r[7],
    }));

  const totalVentas = ventas.reduce((s, v) => s + (v.total || 0), 0);
  return { ventas, totalVentas };
}

function _getMovimientosRecientes() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.MOVIMIENTOS);
  if (!hoja || hoja.getLastRow() < 2) return { movimientos: [] };

  const lastRow = hoja.getLastRow();
  const desde   = Math.max(2, lastRow - 49); // últimos 50
  const datos   = hoja.getRange(desde, 1, lastRow - desde + 1, 10).getValues();

  const movimientos = datos
    .filter(r => r[0])
    .reverse()
    .map(r => ({
      id:       r[0],
      timestamp:r[1],
      tipo:     r[2],
      sku:      r[3],
      cantidad: r[4],
      origen:   r[5],
      destino:  r[6],
      precio:   r[7],
      usuario:  r[8],
      notas:    r[9],
    }));

  return { movimientos };
}
