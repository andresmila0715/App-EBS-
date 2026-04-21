/**
 * ============================================================
 *  EBS+ · GOOGLE APPS SCRIPT
 * ============================================================
 *  1. Abre Google Sheets → Extensiones → Apps Script
 *  2. Borra todo y pega esto
 *  3. Implementar → Nueva implementación
 *     · Tipo: Aplicación web
 *     · Ejecutar como: Yo
 *     · Acceso: Cualquier persona
 *  4. Copia la URL /exec y pégala en config.js
 * ============================================================
 */

var TIPO_FORMULARIO  = 'recoleccion';  // 'carto' | 'recoleccion' | 'macro'
var NOMBRE_MUNICIPIO = 'Enciso';
var NOMBRE_HOJA      = 'Datos App';
var NOMBRE_CARPETA   = 'EBS+ Instrumentos';

// ── Punto de entrada ─────────────────────────────────────────
function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    guardarDatos(datos);
    return respuesta(true, 'Guardado correctamente');
  } catch (err) {
    return respuesta(false, 'Error: ' + err.message);
  }
}
function doGet() {
  return respuesta(true, 'EBS+ activo · ' + NOMBRE_MUNICIPIO + ' · ' + TIPO_FORMULARIO);
}
function guardarDatos(datos) {
  var sheet = obtenerHoja();
  if      (TIPO_FORMULARIO === 'carto')       guardarCarto(sheet, datos);
  else if (TIPO_FORMULARIO === 'recoleccion') guardarRecoleccion(sheet, datos);
  else if (TIPO_FORMULARIO === 'macro')       guardarMacro(sheet, datos);
}

// ── CARTOGRAFÍAS ─────────────────────────────────────────────
function guardarCarto(sheet, d) {
  initEncabezados(sheet, [
    'ID','Fecha','Municipio','Vereda','N° Casa','N° Familia','ID Familia',
    'Coordenada','Riesgo','Convenciones','Observaciones','Foto'
  ]);
  sheet.appendRow([
    d.id||'', d.fecha||'', d.municipio||'',
    d.vereda||'',
    d.casa||'', d.familia||'', d.idFam||'',
    d.coord||'', d.riesgo||'', d.conv||'',
    d.obs||'', d.foto ? '✅ Tiene foto' : 'Sin foto'
  ]);
}

// ════════════════════════════════════════════════════════════
//  RECOLECCIÓN
// ════════════════════════════════════════════════════════════
function guardarRecoleccion(sheet, d) {
  initEncabezados(sheet, [
    'ID','Fecha','Municipio','Vereda','Microterritorio','Teléfono',
    // Familiograma
    'FG Familia','FG Vereda','FG Jefe Nombre','FG Jefe Año','FG Jefe Estado','FG Jefe Enf',
    'FG Pareja Nombre','FG Pareja Año','FG Pareja Estado','FG Pareja Enf',
    'FG Relación Pareja','FG Hijos','FG Observaciones',
    // APGAR — 6 miembros × 9 campos
    'APG M1 Nombre','APG M1 Sexo','APG M1 Edad','APG M1 P1','APG M1 P2','APG M1 P3','APG M1 P4','APG M1 P5','APG M1 Total',
    'APG M2 Nombre','APG M2 Sexo','APG M2 Edad','APG M2 P1','APG M2 P2','APG M2 P3','APG M2 P4','APG M2 P5','APG M2 Total',
    'APG M3 Nombre','APG M3 Sexo','APG M3 Edad','APG M3 P1','APG M3 P2','APG M3 P3','APG M3 P4','APG M3 P5','APG M3 Total',
    'APG M4 Nombre','APG M4 Sexo','APG M4 Edad','APG M4 P1','APG M4 P2','APG M4 P3','APG M4 P4','APG M4 P5','APG M4 Total',
    'APG M5 Nombre','APG M5 Sexo','APG M5 Edad','APG M5 P1','APG M5 P2','APG M5 P3','APG M5 P4','APG M5 P5','APG M5 Total',
    'APG M6 Nombre','APG M6 Sexo','APG M6 Edad','APG M6 P1','APG M6 P2','APG M6 P3','APG M6 P4','APG M6 P5','APG M6 Total',
    'APG Análisis',
    // Zarit — 22 preguntas + resumen
    'Zarit P1','Zarit P2','Zarit P3','Zarit P4','Zarit P5','Zarit P6','Zarit P7',
    'Zarit P8','Zarit P9','Zarit P10','Zarit P11','Zarit P12','Zarit P13','Zarit P14',
    'Zarit P15','Zarit P16','Zarit P17','Zarit P18','Zarit P19','Zarit P20','Zarit P21','Zarit P22',
    'Zarit Total','Zarit Interpretación','Zarit Observaciones',
    // Ecomapa — 13 sistemas + obs
    'Eco Vías Acceso','Eco Fam. Extensa','Eco Prog. Social','Eco Justicia',
    'Eco Org. Comunit.','Eco Vecinos','Eco Espiritualidad','Eco Educación',
    'Eco Trabajo','Eco JAC','Eco Recreación','Eco Salud','Eco Rec. Económicos',
    'Eco Observaciones',
    'Doc URL'
  ]);

  var sub    = d.subData        || {};
  var fg     = sub.familiograma  || {};
  var apg    = sub.apgar         || {};
  var zar    = sub.zarit         || {};
  var eco    = sub.ecomapa       || {};
  var jefe   = fg.principal || fg.padre || {};
  var pareja = pickPareja(fg);

  var hijos = (fg.hijos || []).concat(
    (fg.miembros||[]).filter(function(m){ return m.parentesco === 'hijo'; })
  ).map(function(h){
    return (h.nombre||'?') + ' (' + (h.anio||'?') + '/' + (h.genero||'?') + ')';
  }).join(' | ');

  // APGAR — la app envía r.nombres (plural)
  var apgarMiembros = apg.miembros || [];
  var apgarCols = [];
  for (var am = 0; am < 6; am++) {
    var rm = apgarMiembros[am] || {};
    apgarCols = apgarCols.concat([
      rm.nombres || rm.nombre || '',
      rm.sexo  || '',
      rm.edad  || '',
      valOrEmpty(rm.p1), valOrEmpty(rm.p2), valOrEmpty(rm.p3),
      valOrEmpty(rm.p4), valOrEmpty(rm.p5),
      rm.total !== null && rm.total !== undefined ? rm.total : ''
    ]);
  }
  apgarCols.push(apg.analisis || '');

  // Zarit — 22 respuestas individuales
  var zaritResps  = zar.respuestas || [];
  var zaritCols   = [];
  for (var zi = 0; zi < 22; zi++) {
    var zv = zaritResps[zi];
    zaritCols.push(zv !== null && zv !== undefined ? zv : '');
  }
  var zaritPuntaje = zar.puntaje || 0;
  var zaritInterp  = zaritPuntaje <= 46 ? 'No sobrecarga'
                   : zaritPuntaje <= 55 ? 'Sobrecarga leve' : 'Sobrecarga intensa';
  zaritCols = zaritCols.concat([zaritPuntaje, zaritInterp, zar.obs || '']);

  // Ecomapa
  var ecoVals = eco.valores || eco || {};
  var ECO_IDS = ['eco-vias','eco-fam-ext','eco-programas','eco-justicia','eco-org-com',
                 'eco-vecinos','eco-espiritualidad','eco-educacion','eco-trabajo',
                 'eco-jac','eco-recreacion','eco-salud','eco-recursos'];
  var ecoCols = ECO_IDS.map(function(id){ return ecoVals[id] || ''; });
  ecoCols.push(eco.obs || '');

  // Crear documento
  var carpeta = obtenerCarpeta();
  var urlDoc  = '';
  try { urlDoc = crearDocInstrumentos(carpeta, d, fg, apg, zar, eco); } catch(e) { urlDoc = 'Error: ' + e.message; }

  var fila = [
    d.id||'', d.fecha||'', d.municipio||NOMBRE_MUNICIPIO,
    d.vereda||'', d.microterritorio||'', d.telefono||'',
    fg.familia||'',
    fg.vereda||'',
    jefe.nombre||'', jefe.anio||'', jefe.estado||'', jefe.enf||'',
    pareja.nombre||'', pareja.anio||'', pareja.estado||'', pareja.enf||'',
    fg.relacion || fg.relacionPareja || '',
    hijos, fg.obs||''
  ].concat(apgarCols).concat(zaritCols).concat(ecoCols).concat([urlDoc]);

  sheet.appendRow(fila);
}

// ════════════════════════════════════════════════════════════
//  DOC ÚNICO — 4 secciones en un solo documento
// ════════════════════════════════════════════════════════════
function crearDocInstrumentos(carpeta, d, fg, apg, zar, eco) {
  var jefe       = fg.principal || fg.padre || {};
  var jefeNombre = jefe.nombre  || 'Sin nombre';
  var famNombre  = fg.familia   || d.familia || d.id || 'Familia';
  var nombreDoc  = jefeNombre + ' - ' + famNombre;

  var doc  = DocumentApp.create(nombreDoc);
  var body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  // ── Encabezado de página ──────────────────────────────────
  var hdr = doc.addHeader();
  var hdrPar = hdr.appendParagraph(
    'EQUIPOS BÁSICOS EN SALUD  ·  E.S.E SAN MARTÍN  ·  MUNICIPIO DE ' +
    (d.municipio || NOMBRE_MUNICIPIO).toUpperCase()
  );
  hdrPar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  hdrPar.editAsText().setBold(true).setFontSize(9).setForegroundColor('#005f73');

  // ════════════════════════════════════════════════════════
  //  SECCIÓN 1: FAMILIOGRAMA
  // ════════════════════════════════════════════════════════
  var bannerFg = body.appendTable([['  FORMATO DE FAMILIOGRAMA · ECOMAPA · APGAR  (Resolución 2788)']]);
  bannerFg.getRow(0).getCell(0)
    .setBackgroundColor('#005f73')
    .editAsText().setBold(true).setFontSize(11).setForegroundColor('#ffffff');
  setTableBorders(bannerFg, '#005f73');
  salto(body);

  // — Tabla de identificación —
  var municipio   = d.municipio       || NOMBRE_MUNICIPIO;
  var vereda      = fg.vereda         || d.vereda        || '';
  var microterrit = d.microterritorio || '';
  var numFam      = d.familia         || d.id            || '';
  var tel         = d.telefono        || '';
  var fecha       = d.fecha           || '';

  var tblId = body.appendTable([
    ['Municipio',   municipio,   'Vereda / Sector', vereda],
    ['Microterritorio', microterrit, 'N° Familia', numFam],
    ['Jefe del Hogar', jefeNombre, 'Teléfono', tel],
    ['Fecha de Registro', fecha, '', '']
  ]);
  tblId.getRow(3).getCell(2).setColSpan ? null : null; // GAS no soporta colspan directamente
  setTableBorders(tblId, '#000000');
  // Estilo: celdas de etiqueta con fondo gris
  for (var ri = 0; ri < tblId.getNumRows(); ri++) {
    for (var ci = 0; ci < tblId.getRow(ri).getNumCells(); ci++) {
      var cell = tblId.getRow(ri).getCell(ci);
      if (ci % 2 === 0) {
        cell.setBackgroundColor('#d9d9d9');
        cell.editAsText().setBold(true).setFontSize(9);
      } else {
        cell.editAsText().setFontSize(9);
      }
    }
  }

  // — Datos del familiograma en tabla —
  var pareja = pickPareja(fg);
  var hijosList = (fg.hijos || []).concat(
    (fg.miembros||[]).filter(function(m){ return m.parentesco === 'hijo'; })
  );
  var hijosTexto = hijosList.length > 0
    ? hijosList.map(function(h){ return (h.nombre||'?')+' ('+(h.anio||'?')+')'; }).join(', ')
    : '—';

  salto(body);
  var tblMiembros = body.appendTable([
    ['INTEGRANTE','NOMBRE','AÑO NAC.','ESTADO DE SALUD','ENFERMEDAD'],
    ['Jefe del Hogar ★', jefeNombre, jefe.anio||'—', jefe.estado||'—', jefe.enf||'—'],
    ['Pareja', pareja.nombre||'—', pareja.anio||'—', pareja.estado||'—', pareja.enf||'—']
  ]);
  // Agregar hijos
  hijosList.forEach(function(h, idx) {
    tblMiembros.appendTableRow().appendTableCell('Hijo/a #'+(idx+1))
      .getParentRow().appendTableCell(h.nombre||'—')
      .getParentRow().appendTableCell(h.anio||'—')
      .getParentRow().appendTableCell(h.estado||'—')
      .getParentRow().appendTableCell(h.enf||'—');
  });
  // Otros miembros del familiograma
  (fg.miembros||[]).filter(function(m){ return m.parentesco !== 'hijo'; }).forEach(function(m) {
    var parLabel = { pareja:'Pareja', padre_jefe:'Padre del Jefe', madre_jefe:'Madre del Jefe',
      padre_pareja:'Padre de la Pareja', madre_pareja:'Madre de la Pareja' };
    tblMiembros.appendTableRow().appendTableCell(parLabel[m.parentesco] || m.parentesco || '—')
      .getParentRow().appendTableCell(m.nombre||'—')
      .getParentRow().appendTableCell(m.anio||'—')
      .getParentRow().appendTableCell(m.estado||'—')
      .getParentRow().appendTableCell(m.nota||'—');
  });
  styleDataTable(tblMiembros, true);
  // Marcar jefe con fondo dorado
  tblMiembros.getRow(1).getCell(0).setBackgroundColor('#fef3c7');
  tblMiembros.getRow(1).getCell(0).editAsText().setBold(true).setFontSize(9);

  if (fg.relacionPareja) {
    salto(body);
    body.appendParagraph('Tipo de relación de pareja: ' + fg.relacionPareja)
        .editAsText().setFontSize(9).setItalic(true);
  }

  salto(body);
  seccionTitulo(body, 'FAMILIOGRAMA', 13, true);
  insertarImagen(body, fg.imagenBase64||fg.imagen||'', 480, 320,
    '[Imagen del familiograma — abrir el acordeón antes de guardar para generarla]');
  salto(body);
  body.appendParagraph('Observaciones / Análisis').editAsText().setBold(true).setFontSize(10);
  cajaNota(body, fg.analisis || fg.obs || '', 90);

  // ════════════════════════════════════════════════════════
  //  SECCIÓN 2: ECOMAPA
  // ════════════════════════════════════════════════════════
  paginaNueva(body);
  var bannerEco = body.appendTable([['  ECOMAPA FAMILIAR']]);
  bannerEco.getRow(0).getCell(0)
    .setBackgroundColor('#0e7490')
    .editAsText().setBold(true).setFontSize(13).setForegroundColor('#ffffff');
  setTableBorders(bannerEco, '#0e7490');
  salto(body);

  // Encabezado de identificación resumido
  var tblEcoId = body.appendTable([
    ['Familia', famNombre, 'Vereda', vereda],
    ['Municipio', municipio, 'Fecha', fecha]
  ]);
  for (var re = 0; re < tblEcoId.getNumRows(); re++) {
    for (var ce = 0; ce < tblEcoId.getRow(re).getNumCells(); ce++) {
      var cellE = tblEcoId.getRow(re).getCell(ce);
      if (ce % 2 === 0) {
        cellE.setBackgroundColor('#e0f2fe');
        cellE.editAsText().setBold(true).setFontSize(9);
      } else {
        cellE.editAsText().setFontSize(9);
      }
    }
  }
  setTableBorders(tblEcoId, '#000000');
  salto(body);

  seccionTitulo(body, 'DIAGRAMA DEL ECOMAPA', 11, true);
  insertarImagen(body, eco.imagenBase64||eco.imagen||'', 480, 380,
    '[Imagen del ecomapa — abrir el acordeón antes de guardar para generarla]');
  salto(body);

  var ecoVals = eco.valores || eco || {};
  var ECO_CATS = [
    ['eco-vias','Vías de acceso'],['eco-fam-ext','Familia extensa'],
    ['eco-programas','Programas sociales y comunitarios'],
    ['eco-justicia','Protección en justicia'],
    ['eco-org-com','Organización comunitaria'],
    ['eco-vecinos','Vecinos'],['eco-espiritualidad','Espiritualidad'],
    ['eco-educacion','Educación'],['eco-trabajo','Trabajo'],
    ['eco-jac','Junta de acción comunal'],
    ['eco-recreacion','Recreación y esparcimiento'],
    ['eco-salud','Servicios de salud'],
    ['eco-recursos','Recursos económicos']
  ];
  var ECO_TX = {'-2':'No aplica','1':'Fuerte','2':'Estresante','3':'Moderada',
                '4':'Mod. estresante','5':'Débil','6':'Lev. estresante'};

  var ecoRows = [['Sistema / Recurso','Valor','Interpretación']].concat(
    ECO_CATS.map(function(p){
      var v = String(ecoVals[p[0]]||'');
      return [p[1], v||'—', v ? (ECO_TX[v]||v) : '—'];
    })
  );
  var tblEco = body.appendTable(ecoRows);
  styleDataTable(tblEco, true);
  if (eco.obs) {
    salto(body);
    body.appendParagraph('Observaciones:').editAsText().setBold(true).setFontSize(9);
    cajaNota(body, eco.obs, 50);
  }

  // ════════════════════════════════════════════════════════
  //  SECCIÓN 3: APGAR
  // ════════════════════════════════════════════════════════
  paginaNueva(body);
  var bannerApg = body.appendTable([['  APGAR FAMILIAR']]);
  bannerApg.getRow(0).getCell(0)
    .setBackgroundColor('#065f46')
    .editAsText().setBold(true).setFontSize(13).setForegroundColor('#ffffff');
  setTableBorders(bannerApg, '#065f46');
  salto(body);

  // Encabezado
  var tblApgId = body.appendTable([
    ['Familia', famNombre, 'Vereda', vereda],
    ['Municipio', municipio, 'Fecha', fecha]
  ]);
  for (var ra = 0; ra < tblApgId.getNumRows(); ra++) {
    for (var ca = 0; ca < tblApgId.getRow(ra).getNumCells(); ca++) {
      var cellA = tblApgId.getRow(ra).getCell(ca);
      if (ca % 2 === 0) {
        cellA.setBackgroundColor('#d1fae5');
        cellA.editAsText().setBold(true).setFontSize(9);
      } else {
        cellA.editAsText().setFontSize(9);
      }
    }
  }
  setTableBorders(tblApgId, '#000000');
  salto(body);

  var tblInterp = body.appendTable([[
    'Escala de puntuación por pregunta: 0=Nunca  1=Casi nunca  2=Algunas veces  3=Casi siempre  4=Siempre\n' +
    'Interpretación del total: Normal (21–25) · Disfunción leve (16–20) · Disfunción moderada (11–15) · Disfunción severa (0–10)'
  ]]);
  setTableBorders(tblInterp, '#065f46');
  tblInterp.getRow(0).getCell(0).setBackgroundColor('#ecfdf5').editAsText().setFontSize(9);
  salto(body);

  var apgarMbs = apg.miembros || [];
  var apgarHdrs = ['#','Nombres y Apellidos','Sexo','Edad','P1','P2','P3','P4','P5','Total','Interpretación'];
  var apgarRows = apgarMbs.length > 0
    ? apgarMbs.map(function(r, i){
        var tot = (r.total !== null && r.total !== undefined) ? parseInt(r.total) : null;
        var interp = tot !== null ? (
          tot >= 21 ? 'Normal' : tot >= 16 ? 'Dis. leve' : tot >= 11 ? 'Dis. moderada' : 'Dis. severa'
        ) : '—';
        return [String(i+1),
          r.nombres||r.nombre||'', r.sexo||'', String(r.edad||''),
          valOrEmpty(r.p1), valOrEmpty(r.p2), valOrEmpty(r.p3),
          valOrEmpty(r.p4), valOrEmpty(r.p5),
          tot !== null ? String(tot) : '—', interp];
      })
    : [['','','','','','','','','','',''],['','','','','','','','','','',''],
       ['','','','','','','','','','',''],['','','','','','','','','','','']];

  var tblApgar = body.appendTable([apgarHdrs].concat(apgarRows));
  styleDataTable(tblApgar, true);
  salto(body);
  body.appendParagraph('Análisis del APGAR Familiar').editAsText().setBold(true).setFontSize(10);
  cajaNota(body, apg.analisis || '', 80);

  // ════════════════════════════════════════════════════════
  //  SECCIÓN 4: ZARIT
  // ════════════════════════════════════════════════════════
  paginaNueva(body);
  var bannerZar = body.appendTable([['  INSTRUMENTO DE SOBRECARGA DEL CUIDADOR — ESCALA ZARIT']]);
  bannerZar.getRow(0).getCell(0)
    .setBackgroundColor('#7c3aed')
    .editAsText().setBold(true).setFontSize(13).setForegroundColor('#ffffff');
  setTableBorders(bannerZar, '#7c3aed');
  salto(body);

  // Encabezado identificación
  var tblZarId = body.appendTable([
    ['Familia', famNombre, 'Vereda', vereda],
    ['Municipio', municipio, 'Fecha', fecha]
  ]);
  for (var rz = 0; rz < tblZarId.getNumRows(); rz++) {
    for (var cz = 0; cz < tblZarId.getRow(rz).getNumCells(); cz++) {
      var cellZ = tblZarId.getRow(rz).getCell(cz);
      if (cz % 2 === 0) {
        cellZ.setBackgroundColor('#ede9fe');
        cellZ.editAsText().setBold(true).setFontSize(9);
      } else {
        cellZ.editAsText().setFontSize(9);
      }
    }
  }
  setTableBorders(tblZarId, '#000000');
  salto(body);

  body.appendParagraph(
    'Frecuencia: 0=Nunca · 1=Rara vez · 2=Algunas veces · 3=Bastantes veces · 4=Casi siempre  |  Puntuación total: 0–88'
  ).editAsText().setFontSize(9).setItalic(true);
  salto(body);

  var ZARIT_PREGS = [
    '¿Cree que su familiar solicita más ayuda de la que realmente necesita?',
    '¿Cree que debido al tiempo que dedica a su familiar ya no dispone de tiempo suficiente para usted?',
    '¿Se siente agobiado entre cuidar a su familiar y atender además otras responsabilidades en su trabajo o familia?',
    '¿Se siente avergonzado por la conducta de su familiar?',
    '¿Se siente enfadado cuando está cerca de su familiar?',
    '¿Piensa que su familiar afecta negativamente a su relación con otros miembros de su familia?',
    '¿Tiene miedo de lo que el futuro depare a su familiar?',
    '¿Cree que su familiar depende de usted?',
    '¿Se siente tenso cuando está cerca de su familiar?',
    '¿Cree que su salud se ha resentido por cuidar a su familiar?',
    '¿Cree que no tiene tanta intimidad como le gustaría debido a su familiar?',
    '¿Cree que su vida social se ha resentido por cuidar a su familiar?',
    '¿Se siente incómodo por desatender a sus amistades debido a su familiar?',
    '¿Cree que su familiar parece esperar que usted sea la persona que le cuide, como si usted fuera la única persona de quien depende?',
    '¿Cree que no tiene suficiente dinero para cuidar a su familiar además de sus otros gastos?',
    '¿Cree que será incapaz de cuidarle/a por mucho más tiempo?',
    '¿Siente que ha perdido el control de su vida desde la enfermedad de su familiar?',
    '¿Desearía poder dejar el cuidado de su familiar a otros?',
    '¿Se siente indeciso sobre qué hacer con su familiar?',
    '¿Cree que debería hacer más por su familiar?',
    '¿Cree que podría cuidar mejor de su familiar?',
    'Globalmente, ¿qué grado de carga experimenta por el hecho de cuidar a su familiar?'
  ];
  var zaritResps = zar.respuestas || [];
  var zaritRows2 = [['N°','PREGUNTA','0','1','2','3','4']].concat(
    ZARIT_PREGS.map(function(preg, i){
      var val = (zaritResps[i] !== null && zaritResps[i] !== undefined) ? String(zaritResps[i]) : '';
      return [String(i+1), preg,
        val==='0'?'✓':'', val==='1'?'✓':'', val==='2'?'✓':'',
        val==='3'?'✓':'', val==='4'?'✓':''];
    })
  );
  var tblZarit = body.appendTable(zaritRows2);
  styleZaritTable(tblZarit);
  salto(body);
  var puntaje = zar.puntaje || 0;
  var interp  = puntaje <= 46 ? 'No sobrecarga' : puntaje <= 55 ? 'Sobrecarga leve' : 'Sobrecarga intensa';

  var tblResult = body.appendTable([['Puntaje total: ' + puntaje + '     Interpretación: ' + interp]]);
  var bgRes = puntaje <= 46 ? '#d1fae5' : puntaje <= 55 ? '#fef3c7' : '#fee2e2';
  tblResult.getRow(0).getCell(0)
    .setBackgroundColor(bgRes)
    .editAsText().setBold(true).setFontSize(11).setForegroundColor('#1e1e1e');
  setTableBorders(tblResult, '#7c3aed');

  salto(body);
  body.appendParagraph('Interpretación de los resultados y observaciones').editAsText().setBold(true).setFontSize(10);
  cajaNota(body, zar.obs || '', 80);

  // ── Pie de página ─────────────────────────────────────────
  var ftr = doc.addFooter();
  var ftrPar = ftr.appendParagraph(
    'E.S.E San Martín  ·  Instrumentos de Recolección  ·  ' +
    famNombre + '  ·  ' + (d.fecha||'')
  );
  ftrPar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  ftrPar.editAsText().setFontSize(8).setForegroundColor('#888888');

  // Mover a carpeta (método robusto)
  moverACarpeta(doc, carpeta);
  doc.saveAndClose();
  return doc.getUrl();
}

// ── MACRO ────────────────────────────────────────────────────
function guardarMacro(sheet, d) {
  initEncabezados(sheet, [
    'ID','Fecha Registro','Municipio',
    'Fecha Visita','Vereda / Sector','Microterritorio','Auxiliar a Cargo',
    'Novedad','Familia','N° Casa','Teléfono',
    'Cargado','Seg. Tel. Laboratorios','Seguimiento Médico',
    'Acueducto','Pozo / Alcantarillado','Unidad Sanitaria','Ducha',
    'Cocina','Manejo de Residuos','Paredes','Piso','Techo',
    'Electricidad','Riesgo Entorno','Nivel Riesgo Vivienda',
    'Ingresos Económicos','SISBEN',
    'N° Integrante','Nombre','Tipo Doc','N° Doc','Edad',
    'Curso de Vida','Género','Orientación Sexual','Migrante',
    'Enfermedades','De Novo','Adherencia',
    'Discapacidad','Tipo Discapacidad','Cert. Discapacidad','Cuidador',
    'EPS','Alertas','Profesional Requerido',
    'Víctima Conflicto','Víctima VIF',
    'Nivel Educativo','Teléfono Integrante'
  ]);

  var viv  = d.vivienda    || {};
  var ints = d.integrantes || [];
  var base = [
    d.id||'', d.fecha||'', d.municipio||'',
    d.fechaVisita||'', d.vereda||'', d.microterritorio||'', d.auxiliar||'',
    d.novedad||'', d.familia||'', d.casa||'', d.telefono||'',
    d.cargado||'', d.segTelf||'', d.segMedico||'',
    viv.acueducto||'', viv.pozo||'', viv.sanitaria||'', viv.ducha||'',
    viv.cocina||'', viv.residuos||'', viv.paredes||'', viv.piso||'',
    viv.techo||'', viv.electricidad||'', viv.riesgoEntorno||'',
    viv.nivelRiesgo||'', viv.ingresos||'', viv.sisben||''
  ];
  if (ints.length === 0) {
    sheet.appendRow(base.concat(new Array(23).fill('')));
  } else {
    ints.forEach(function(int, idx) {
      sheet.appendRow(base.concat([
        idx+1, int.nombre||'', int.tipodoc||'', int.numdoc||'', int.edad||'',
        int.cursovida||'', int.genero||'', int.orientacion||'', int.migrante||'',
        int.enfermedades||'', int.denovo||'', int.adherencia||'',
        int.discapacidad||'', int.tipodiscap||'', int.certdiscap||'', int.cuidador||'',
        int.eps||'', int.alertas||'', int.profesional||'',
        int.victimaca||'', int.victimavif||'',
        int.educativo||'', int.telefono||''
      ]));
    });
  }
}

// ════════════════════════════════════════════════════════════
//  HELPERS DE DOCUMENTO
// ════════════════════════════════════════════════════════════

// Salto de página robusto — appendPageBreak no existe en GAS
function paginaNueva(body) {
  body.appendParagraph('').appendPageBreak();
}

function salto(body) { body.appendParagraph(''); }

function seccionTitulo(body, texto, size, color) {
  var p = body.appendParagraph(texto);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  var t = p.editAsText().setBold(true).setFontSize(size);
  if (color) t.setForegroundColor('#005f73');
}

function cajaNota(body, texto, minH) {
  var tbl = body.appendTable([[texto || '']]);
  tbl.getRow(0).getCell(0).setMinimumHeight(minH || 60);
  setTableBorders(tbl, '#000000');
  tbl.getRow(0).getCell(0).editAsText().setFontSize(10);
}

function nota(body, label, texto) {
  body.appendParagraph(label + texto).editAsText().setFontSize(10);
}

function estiloCeldas(tbl, size) {
  for (var r = 0; r < tbl.getNumRows(); r++)
    for (var c = 0; c < tbl.getRow(r).getNumCells(); c++)
      tbl.getRow(r).getCell(c).editAsText().setFontSize(size || 9);
}

function insertarImagen(body, base64, w, h, fallback) {
  if (base64) {
    try {
      var clean = base64.replace(/^data:image\/\w+;base64,/, '');
      var blob  = Utilities.newBlob(Utilities.base64Decode(clean), 'image/png', 'img.png');
      var par   = body.appendParagraph('');
      par.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      par.appendInlineImage(blob).setWidth(w).setHeight(h);
      return;
    } catch(e) { /* cae al fallback */ }
  }
  body.appendParagraph(fallback || '[Imagen no disponible]')
      .editAsText().setItalic(true).setForegroundColor('#888888').setFontSize(9);
}

function styleDataTable(tbl, hasHeader) {
  setTableBorders(tbl, '#cccccc');
  if (hasHeader) {
    var hr = tbl.getRow(0);
    for (var c = 0; c < hr.getNumCells(); c++) {
      hr.getCell(c).setBackgroundColor('#005f73');
      hr.getCell(c).editAsText().setBold(true).setFontSize(9).setForegroundColor('#ffffff');
    }
  }
  for (var r = hasHeader ? 1 : 0; r < tbl.getNumRows(); r++) {
    var bg = r % 2 === 0 ? '#f0f9ff' : '#ffffff';
    for (var c2 = 0; c2 < tbl.getRow(r).getNumCells(); c2++) {
      tbl.getRow(r).getCell(c2).setBackgroundColor(bg);
      tbl.getRow(r).getCell(c2).editAsText().setFontSize(9);
    }
  }
}

function styleZaritTable(tbl) {
  setTableBorders(tbl, '#b2ebf2');
  var hr = tbl.getRow(0);
  for (var c = 0; c < hr.getNumCells(); c++) {
    hr.getCell(c).setBackgroundColor('#0097a7');
    hr.getCell(c).editAsText().setBold(true).setFontSize(9).setForegroundColor('#ffffff');
  }
  for (var r = 1; r < tbl.getNumRows(); r++) {
    var bg  = r % 2 === 0 ? '#e0f7fa' : '#ffffff';
    var row = tbl.getRow(r);
    row.getCell(0).setBackgroundColor('#b2ebf2').editAsText().setBold(true).setFontSize(9);
    row.getCell(1).setBackgroundColor(bg).editAsText().setFontSize(8);
    for (var c2 = 2; c2 < row.getNumCells(); c2++) {
      row.getCell(c2).setBackgroundColor(bg).setMinimumHeight(18);
      row.getCell(c2).editAsText().setFontSize(9).setForegroundColor('#005f73');
    }
  }
}

function setTableBorders(tbl, color) {
  tbl.setBorderWidth(1);
  tbl.setBorderColor(color || '#cccccc');
}

// ════════════════════════════════════════════════════════════
//  UTILIDADES GENERALES
// ════════════════════════════════════════════════════════════

function valOrEmpty(v) {
  return (v !== null && v !== undefined && v !== '') ? String(v) : '';
}

function pickPareja(fg) {
  if (fg.pareja && (fg.pareja.nombre || fg.pareja.anio)) return fg.pareja;
  if (fg.madre  && (fg.madre.nombre  || fg.madre.anio))  return fg.madre;
  var par = (fg.miembros||[]).filter(function(m){ return m.parentesco === 'pareja'; })[0];
  return par || {};
}

function initEncabezados(sheet, cols) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(cols);
    var r = sheet.getRange(1, 1, 1, cols.length);
    r.setBackground('#005f73');
    r.setFontColor('#ffffff');
    r.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function obtenerHoja() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(NOMBRE_HOJA);
  if (!hoja) hoja = ss.insertSheet(NOMBRE_HOJA);
  return hoja;
}

function obtenerCarpeta() {
  var folders = DriveApp.getFoldersByName(NOMBRE_CARPETA);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(NOMBRE_CARPETA);
}

// Método robusto: elimina de TODOS los parents actuales, luego agrega a carpeta destino
function moverACarpeta(doc, carpeta) {
  var file    = DriveApp.getFileById(doc.getId());
  var parents = file.getParents();
  while (parents.hasNext()) {
    parents.next().removeFile(file);
  }
  carpeta.addFile(file);
}

function respuesta(ok, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, mensaje: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
