// LAYOUT REZAGOS ARBA01
// Apps Script Web App - sirve el HTML y expone funciones para google.script.run
//
// COMO DEPLOYAR:
// 1. En el editor de Apps Script del archivo "BD Layout 2026":
//    - Crear un nuevo archivo HTML llamado "layout-rezagos-arba01"
//    - Pegar el contenido del HTML en ese archivo
// 2. Reemplazar el contenido de macros.gs con este archivo
// 3. Guardar (Ctrl+S)
// 4. Implementar > Gestionar implementaciones > editar (lapiz)
//    - Ejecutar como: Yo
//    - Acceso: Cualquier usuario de Mercadolibre SRL
//    - Guardar e Implementar
// 5. Abrir la URL del deploy en el navegador - esa es la app

var SHEET_NAME = 'Almacenamiento';

// Indices de columnas (base 0)
// A=0 Marca temporal, B=1 PALLET, C=2 DESTINO, D=3 Responsable, F=5 Dimensiones
var COL_TS        = 0;
var COL_PALLET    = 1;
var COL_POSITION  = 2;
var COL_OPERATOR  = 3;
var COL_DIMENSION = 5;

// ─────────────────────────────────────────────────────────────
// Sirve el HTML como aplicacion web
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('layout-rezagos-arba01')
    .setTitle('Layout Rezagos - ARBA01')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────────────────────────
// Lee la hoja Almacenamiento y devuelve el estado actual de cada posicion.
// Para cada posicion, toma el registro mas reciente (por Marca temporal).
// Llamado via google.script.run.getLayoutState()
function getLayoutState() {
  try {
    var sheet = getSheet();
    var data  = sheet.getDataRange().getValues();

    var posState = {};

    for (var i = 1; i < data.length; i++) {
      var row      = data[i];
      var tsRaw    = row[COL_TS];
      var pallet   = String(row[COL_PALLET]   || '').trim();
      var position = String(row[COL_POSITION]  || '').trim().toUpperCase();
      var operator = String(row[COL_OPERATOR]  || '').trim();
      var dim      = String(row[COL_DIMENSION] || '').trim();

      if (!position) continue;

      var ts = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);

      if (!posState[position] || ts > posState[position]._ts) {
        posState[position] = {
          pallet   : pallet,
          operator : operator,
          dimension: dim,
          timestamp: ts.toLocaleString('es-AR'),
          _ts      : ts
        };
      }
    }

    // Limpiar campo interno _ts antes de responder
    var keys = Object.keys(posState);
    for (var k = 0; k < keys.length; k++) {
      delete posState[keys[k]]._ts;
    }

    return { success: true, data: posState };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
// Registra un movimiento (almacenar o vaciar) en la hoja Almacenamiento.
// Llamado via google.script.run.registerMovement(data)
function registerMovement(data) {
  try {
    var sheet = getSheet();
    var ts    = data.timestamp ? new Date(data.timestamp) : new Date();

    var row = [];
    row[COL_TS]        = ts;
    row[COL_PALLET]    = data.pallet    || '';
    row[COL_POSITION]  = data.position  || '';
    row[COL_OPERATOR]  = data.operator  || '';
    row[COL_DIMENSION] = data.dimension || '';

    sheet.appendRow(row);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('No se encontro la hoja: ' + SHEET_NAME);
  return sh;
}
