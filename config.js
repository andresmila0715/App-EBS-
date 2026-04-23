/**
 * ============================================================
 *  EBS+ · ARCHIVO DE CONFIGURACIÓN CENTRAL
 * ============================================================
 *  Este es el único archivo que debes editar para:
 *   1. Agregar/cambiar municipios
 *   2. Agregar/cambiar usuarios y contraseñas
 *   3. Cambiar los links de Google Sheets
 *
 *  ⚠️  IMPORTANTE — CADA VEZ QUE EDITES ESTE ARCHIVO:
 *      Abre sw.js y cambia el número de versión del caché:
 *        CACHE_NAME = 'ebs-app-v7'  →  'ebs-app-v8'  (suma 1)
 *      Esto obliga a todos los dispositivos a descargar
 *      el config actualizado en la próxima apertura.
 * ============================================================
 */

// ─────────────────────────────────────────────────────────────
//  1. MUNICIPIOS
//     Lista de todos los municipios activos en la app.
//     El nombre aquí debe coincidir exactamente con el
//     campo "municipio" en los usuarios de abajo.
// ─────────────────────────────────────────────────────────────
const MUNIS = [
  'La Belleza',
  'Sucre',
  'Enciso',
  'Muestra',
  // 'Landázuri',   // <-- para agregar un nuevo municipio:
  //                //     1) descomenta esta línea
  //                //     2) agrega su usuario abajo
  //                //     3) agrega sus Sheets en SHEETS_URLS abajo
];


// ─────────────────────────────────────────────────────────────
//  2. USUARIOS
//     Formato de cada usuario:
//     'NOMBRE_EN_MAYUSCULAS': {
//       password: 'Contraseña123',
//       municipio: 'Nombre del Municipio',   // null si es BOSS
//       role: 'user'  // o 'boss' para ver todos los municipios
//     }
// ─────────────────────────────────────────────────────────────
const USERS = {
  // ── Usuarios por municipio ──────────────────────────────
  'LABELLEZA': {
    password: 'Labellez@2026',
    municipio: 'La Belleza',
    role: 'user'
  },
  'SUCRE': {
    password: 'Sucr32026',
    municipio: 'Sucre',
    role: 'user'
  },
  'ENCISO': {
    password: 'Encis02026',
    municipio: 'Enciso',
    role: 'user'
  },
    'MUESTRA': {
    password: 'MUESTRA',
    municipio: 'Muestra',
    role: 'user'
  },
  // Para agregar un nuevo usuario de municipio:
  // 'LANDAZURI': {
  //   password: 'Land@2026',
  //   municipio: 'Landázuri',
  //   role: 'user'
  // },

  // ── Usuario administrador (ve todos los municipios) ─────
  'BOSS': {
    password: 'Allfor1',
    municipio: null,
    role: 'boss'
  },
};


// ─────────────────────────────────────────────────────────────
//  3. URLS DE GOOGLE SHEETS (Apps Script Web App)
//
//     Cada municipio tiene 3 spreadsheets, uno por formulario:
//       - carto       → Información Cartografías
//       - recoleccion → Instrumento de Recolección de Datos
//       - macro       → Macro (caracterización familiar)
//
//     Para obtener una URL:
//       1. Abre el script en Google Sheets
//       2. Implementar > Nueva implementación
//       3. Tipo: Aplicación web
//       4. Ejecutar como: Yo
//       5. Quién tiene acceso: Cualquier persona
//       6. Copia la URL que termina en /exec
//
//     Si un formulario aún no tiene spreadsheet asignado,
//     deja el valor como null y los datos quedarán en cola
//     local hasta que se configure.
// ─────────────────────────────────────────────────────────────
const SHEETS_URLS = {

  'La Belleza': {
    carto:       'https://script.google.com/macros/s/REEMPLAZAR_LABELLEZA_CARTO/exec',
    recoleccion: 'https://script.google.com/macros/s/REEMPLAZAR_LABELLEZA_RECOLECCION/exec',
    macro:       'https://script.google.com/macros/s/AKfycbxIhQ29PNqvYy__USCB_fEdiSum3yF9DwybKY_u9_0T52RDODHJdx1GTqCqhEOBR8SP/exec',
  },

  'Sucre': {
    carto:       'https://script.google.com/macros/s/REEMPLAZAR_SUCRE_CARTO/exec',
    recoleccion: 'https://script.google.com/macros/s/REEMPLAZAR_SUCRE_RECOLECCION/exec',
    macro:       'https://script.google.com/macros/s/REEMPLAZAR_SUCRE_MACRO/exec',
  },

    'Enciso': {
    carto:       'https://script.google.com/macros/s/AKfycbwEqpgAikHpJNMGuJW1m8ZQEAdkFDEmd3GBpjg4DIgTvm5m9RBeXlfAhiTxt9QYiOtv/exec',
    recoleccion: 'https://script.google.com/macros/s/AKfycbzDyWxFW8Wh804SnyNwUx6rxzv9t4PRKM6eIEV7XKIYLCiVTuObNu_u7zkyW9LrqjOd/exec',
    macro:       'https://script.google.com/macros/s/AKfycbwpJwJ-RDjT5xv2E2UCLw2xcDAWewQO9heogU7TBK1dACTQlmX8AInZYg5reTCfWzMg/exec',
  },

    'Muestra': {
    carto:       'https://script.google.com/macros/s/REEMPLAZAR_SUCRE_CARTO/exec',
    recoleccion: 'https://script.google.com/macros/s/REEMPLAZAR_SUCRE_RECOLECCION/exec',
    macro:       'https://script.google.com/macros/s/REEMPLAZAR_SUCRE_MACRO/exec',
  },
  // Para agregar Landázuri:
  // 'Landázuri': {
  //   carto:       'https://script.google.com/macros/s/TU_URL_LANDAZURI_CARTO/exec',
  //   recoleccion: 'https://script.google.com/macros/s/TU_URL_LANDAZURI_RECOLECCION/exec',
  //   macro:       'https://script.google.com/macros/s/TU_URL_LANDAZURI_MACRO/exec',
  // },
};


// ─────────────────────────────────────────────────────────────
//  (No editar abajo de esta línea)
// ─────────────────────────────────────────────────────────────

/**
 * Devuelve la URL del Apps Script para un municipio y tipo de formulario.
 * Retorna null si no está configurada.
 */
function getSheetUrl(municipio, tipo) {
  return SHEETS_URLS[municipio]?.[tipo] ?? null;
}
