/**
 * ============================================================
 *  EBS+ · GOOGLE APPS SCRIPT — PLANTILLA PARA SHEETS
 * ============================================================
 *
 *  INSTRUCCIONES:
 *  1. Abre Google Sheets (crea una hoja nueva para cada caso)
 *  2. Ve a Extensiones > Apps Script
 *  3. Borra el código que aparece y pega este archivo
 *  4. Arriba en el código, elige qué formulario maneja este
 *     script cambiando la variable TIPO_FORMULARIO
 *  5. Guarda y luego: Implementar > Nueva implementación
 *     · Tipo: Aplicación web
 *     · Ejecutar como: Yo (tu cuenta)
 *     · Quién tiene acceso: Cualquier persona
 *  6. Copia la URL /exec y pégala en config.js
 *
 *  Necesitas crear 3 scripts × N municipios.
 *  Por ejemplo para 2 municipios = 6 scripts en total.
 * ============================================================
 */

// ─────────────────────────────────────────────────────────────
//  AJUSTAR ESTO en cada script según corresponda:
// ─────────────────────────────────────────────────────────────

var TIPO_FORMULARIO = 'carto'; 
// Opciones:  'carto'  |  'recoleccion'  |  'macro'

var NOMBRE_MUNICIPIO = 'La Belleza';
// Escribe el nombre exacto del municipio (solo referencia/info)

var NOMBRE_HOJA = 'Datos';
// Nombre de la pestaña dentro del spreadsheet donde se guardan los datos.
// Si no existe, se crea automáticamente.

// ─────────────────────────────────────────────────────────────
//  (No editar abajo de esta línea salvo ajustes avanzados)
// ─────────────────────────────────────────────────────────────


// ── Definición de columnas por tipo de formulario ──────────
var COLUMNAS = {

  carto: [
    'id', 'fecha', 'municipio',
    'casa', 'familia', 'idFam',
    'coord', 'riesgo', 'conv',
    'obs', 'foto'
  ],

  recoleccion: [
    'id', 'fecha', 'municipio',
    'codigo', 'personas', 'estrato', 'vivienda',
    'servicios', 'salud', 'menores', 'mayores',
    'enfermedades', 'obs'
  ],

  macro: [
    'id', 'fecha', 'municipio',
    'fechaVisita', 'vereda', 'microterritorio', 'auxiliar',
    'novedad', 'familia', 'casa', 'telefono',
    'pyp', 'cargado', 'segTelf', 'segMedico',
    // Vivienda
    'viv_acueducto', 'viv_pozo', 'viv_sanitaria', 'viv_ducha',
    'viv_cocina', 'viv_residuos', 'viv_paredes', 'viv_piso',
    'viv_techo', 'viv_electricidad', 'viv_riesgoEntorno',
    'viv_nivelRiesgo', 'viv_ingresos', 'viv_sisben',
    // Integrantes (se expanden hasta MAX_INTEGRANTES columnas)
    // Se generan dinámicamente abajo
  ],
};

var MAX_INTEGRANTES = 15; // Número máximo de integrantes por familia en el macro


// ── Punto de entrada principal ──────────────────────────────

function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    guardarFila(datos);
    return respuesta(true, 'Guardado correctamente');
  } catch (err) {
    return respuesta(false, 'Error: ' + err.message);
  }
}

function doGet(e) {
  // Permite verificar que el script está activo desde el navegador
  return respuesta(true, 'EBS+ Script activo · ' + NOMBRE_MUNICIPIO + ' · ' + TIPO_FORMULARIO);
}


// ── Lógica de guardado ──────────────────────────────────────

function guardarFila(datos) {
  var sheet = obtenerHoja();
  var cols = obtenerColumnas();

  // Si la hoja está vacía, escribir encabezados
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(cols);
    sheet.getRange(1, 1, 1, cols.length)
      .setBackground('#005f73')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Construir la fila según el tipo
  var fila;
  if (TIPO_FORMULARIO === 'macro') {
    fila = construirFilaMacro(datos, cols);
  } else {
    fila = construirFilaSimple(datos, cols);
  }

  sheet.appendRow(fila);
}

function construirFilaSimple(datos, cols) {
  return cols.map(function(col) {
    var val = datos[col];
    if (val === null || val === undefined) return '';
    if (typeof val === 'object') return JSON.stringify(val);
    return val;
  });
}

function construirFilaMacro(datos, cols) {
  // Para el macro, los campos de vivienda y de integrantes
  // se "aplanan" en columnas separadas
  var viv = datos.vivienda || {};
  var ints = datos.integrantes || [];

  return cols.map(function(col) {
    // Campos normales
    if (datos[col] !== undefined) return datos[col] || '';

    // Campos de vivienda (prefijo viv_)
    if (col.startsWith('viv_')) {
      var campo = col.replace('viv_', '');
      return viv[campo] || '';
    }

    // Campos de integrantes (formato: int1_nombre, int2_edad, etc.)
    var matchInt = col.match(/^int(\d+)_(.+)$/);
    if (matchInt) {
      var idx = parseInt(matchInt[1]) - 1;
      var prop = matchInt[2];
      return (ints[idx] && ints[idx][prop]) ? ints[idx][prop] : '';
    }

    return '';
  });
}


// ── Columnas dinámicas para integrantes (MACRO) ─────────────

function obtenerColumnas() {
  var base = COLUMNAS[TIPO_FORMULARIO].slice();

  if (TIPO_FORMULARIO === 'macro') {
    var camposInt = [
      'nombre', 'tipodoc', 'numdoc', 'edad', 'cursovida',
      'genero', 'orientacion', 'migrante', 'enfermedades',
      'denovo', 'adherencia', 'discapacidad', 'tipodiscap',
      'certdiscap', 'cuidador', 'eps', 'alertas',
      'profesional', 'victimaca', 'victimavif', 'educativo', 'telefono'
    ];
    for (var i = 1; i <= MAX_INTEGRANTES; i++) {
      camposInt.forEach(function(c) {
        base.push('int' + i + '_' + c);
      });
    }
  }

  return base;
}


// ── Utilidades ──────────────────────────────────────────────

function obtenerHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA);
  }
  return hoja;
}

function respuesta(ok, mensaje) {
  var resultado = JSON.stringify({ ok: ok, mensaje: mensaje });
  return ContentService
    .createTextOutput(resultado)
    .setMimeType(ContentService.MimeType.JSON);
}
