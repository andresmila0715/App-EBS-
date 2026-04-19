/**
 * ============================================================
 *  EBS+ · GOOGLE APPS SCRIPT
 * ============================================================
 *  INSTRUCCIONES:
 *  1. Abre Google Sheets
 *  2. Extensiones > Apps Script > borra todo y pega esto
 *  3. Cambia TIPO_FORMULARIO según corresponda
 *  4. Implementar > Nueva implementación
 *     · Tipo: Aplicación web
 *     · Ejecutar como: Yo
 *     · Quién tiene acceso: Cualquier persona
 *  5. Copia la URL /exec y pégala en config.js
 *
 *  DOCUMENTOS QUE SE CREAN AUTOMÁTICAMENTE:
 *  · Doc 1 – Familiograma (imagen del canvas + tabla de integrantes)
 *  · Doc 2 – Ecomapa + APGAR Familiar + Zarit (tablas con formato institucional)
 *
 *  CARPETA DE GOOGLE DRIVE:
 *  Cambia NOMBRE_CARPETA para organizar los documentos.
 *  Si la carpeta no existe se crea automáticamente.
 * ============================================================
 */

var TIPO_FORMULARIO  = 'recoleccion';  // 'carto' | 'recoleccion' | 'macro'
var NOMBRE_MUNICIPIO = 'La Belleza';
var NOMBRE_HOJA      = 'Datos';
var NOMBRE_CARPETA   = 'EBS+ Instrumentos';   // carpeta en Google Drive

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

function doGet(e) {
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
    'ID','Fecha','Municipio','N° Casa','N° Familia','ID Familia',
    'Coordenada','Riesgo','Convenciones','Observaciones','Foto'
  ]);
  sheet.appendRow([
    d.id||'', d.fecha||'', d.municipio||'',
    d.casa||'', d.familia||'', d.idFam||'',
    d.coord||'', d.riesgo||'', d.conv||'',
    d.obs||'', d.foto ? '✅ Tiene foto' : 'Sin foto'
  ]);
}

// ── RECOLECCIÓN ──────────────────────────────────────────────
function guardarRecoleccion(sheet, d) {
  initEncabezados(sheet, [
    'ID','Fecha','Municipio',
    'FG Jefe Nombre','FG Jefe Año','FG Jefe Estado','FG Jefe Enf',
    'FG Pareja Nombre','FG Pareja Año','FG Pareja Estado','FG Pareja Enf',
    'FG Relación Pareja','FG Hijos Resumen','FG Observaciones',
    'APGAR Resumen','Zarit Puntaje','Zarit Obs',
    'Eco Vías Acceso','Eco Fam. Extensa','Eco Prog. Social',
    'Eco Justicia','Eco Org. Comunit.','Eco Vecinos',
    'Eco Espiritualidad','Eco Educación','Eco Trabajo',
    'Eco JAC','Eco Recreación','Eco Salud','Eco Rec. Económicos',
    'Doc Familiograma URL','Doc Instrumentos URL'
  ]);

  var sub  = d.subData        || {};
  var fg   = sub.familiograma  || {};
  var apg  = sub.apgar         || {};
  var zar  = sub.zarit         || {};
  var eco  = sub.ecomapa       || {};
  var jefe   = fg.principal || fg.padre || {};
  var pareja = fg.pareja    || fg.madre || {};

  var hijos = (fg.hijos||[]).map(function(h){
    return (h.nombre||'?')+'('+(h.anio||'?')+','+(h.genero==='F'?'F':'M')+')';
  }).join(' | ');

  var apgarMiembros = apg.miembros || apg.rows || [];
  var apgarResumen = Array.isArray(apgarMiembros) && apgarMiembros.length
    ? apgarMiembros.map(function(r){ return (r.nombre||'?')+':'+(r.total||'?'); }).join(' | ')
    : (apg.obs || '');

  var ecoVals = eco.valores || eco || {};

  // Crear documentos Google Docs
  var carpeta = obtenerCarpeta();
  var titulo  = (d.familia || d.id || 'Registro') + ' · ' + (d.fecha || '');
  var urlFamiliograma = crearDocFamiliograma(carpeta, titulo, d, fg);
  var urlInstrumentos = crearDocInstrumentos(carpeta, titulo, d, apg, zar, eco);

  sheet.appendRow([
    d.id||'', d.fecha||'', d.municipio||'',
    jefe.nombre||'',   jefe.anio||'',   jefe.estado||'',   jefe.enf||'',
    pareja.nombre||'', pareja.anio||'', pareja.estado||'', pareja.enf||'',
    fg.relacion||fg.relacionPareja||'', hijos, fg.obs||'',
    apgarResumen, zar.puntaje||0, zar.obs||'',
    ecoVals['eco-vias']||'',           ecoVals['eco-fam-ext']||'',
    ecoVals['eco-programas']||'',      ecoVals['eco-justicia']||'',
    ecoVals['eco-org-com']||'',        ecoVals['eco-vecinos']||'',
    ecoVals['eco-espiritualidad']||'', ecoVals['eco-educacion']||'',
    ecoVals['eco-trabajo']||'',        ecoVals['eco-jac']||'',
    ecoVals['eco-recreacion']||'',     ecoVals['eco-salud']||'',
    ecoVals['eco-recursos']||'',
    urlFamiliograma, urlInstrumentos
  ]);
}

// ════════════════════════════════════════════════════════════
//  DOC 1 — FAMILIOGRAMA  (formato institucional ESE San Martín)
// ════════════════════════════════════════════════════════════
function crearDocFamiliograma(carpeta, titulo, d, fg) {
  var doc  = DocumentApp.create('Familiograma · ' + titulo);
  var body = doc.getBody();
  body.setMarginTop(28).setMarginBottom(28).setMarginLeft(50).setMarginRight(50);

  // ── Encabezado institucional (header del doc) ─────────────
  var header = doc.addHeader();
  var hPar   = header.appendParagraph(
    'EQUIPOS BÁSICOS EN SALUD\nE.S.E SAN MARTÍN\nMUNICIPIO DE ' +
    (d.municipio || NOMBRE_MUNICIPIO).toUpperCase()
  );
  hPar.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  hPar.editAsText().setBold(true).setItalic(true).setFontSize(10);

  // ── Título del formato ────────────────────────────────────
  var titPar = body.appendParagraph(
    'FORMATO DE FAMILIOGRAMA ECOMAPA, APGAR (LINEAMIENTOS RESOLUCIÓN 2788)'
  );
  titPar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titPar.editAsText().setBold(true).setFontSize(11);

  // ── Tabla: DATOS DE IDENTIFICACIÓN ───────────────────────
  var tblIdHdr = body.appendTable([['DATOS DE IDENTIFICACIÓN']]);
  tblIdHdr.getRow(0).getCell(0).setBackgroundColor('#d9d9d9');
  tblIdHdr.getRow(0).getCell(0).editAsText().setBold(true).setFontSize(9);
  setTableBorders(tblIdHdr, '#000000');

  var municipio    = d.municipio       || NOMBRE_MUNICIPIO;
  var barrio       = fg.barrio         || d.vereda        || '';
  var direccion    = fg.direccion      || d.direccion     || '';
  var territorio   = fg.territorio     || d.territorio    || '';
  var microterrit  = d.microterritorio || '';
  var numFamilia   = d.familia         || d.id            || '';
  var jefe         = fg.principal || fg.padre || {};
  var jefeNombre   = jefe.nombre || '';
  var visitaNombre = fg.quienRecibe || jefeNombre;
  var telefono     = d.telefono || '';

  var tblId = body.appendTable([
    ['Municipio: ' + municipio,
     'Barrio o Vereda: ' + barrio,
     'Dirección: ' + direccion],
    ['Territorio: ' + territorio,
     'Microterritorio: ' + microterrit,
     'Número De Familia: ' + numFamilia],
    ['Nombre del jefe del Hogar: ' + jefeNombre, '', ''],
    ['Nombre de quien recibe la visita:', visitaNombre, ''],
    ['N° teléfono contacto familiar:', telefono, '']
  ]);
  setTableBorders(tblId, '#000000');
  for (var r = 0; r < tblId.getNumRows(); r++) {
    for (var c = 0; c < tblId.getRow(r).getNumCells(); c++) {
      tblId.getRow(r).getCell(c).editAsText().setFontSize(9);
    }
  }
  body.appendParagraph('');

  // ── Título FAMILIOGRAMA ───────────────────────────────────
  var fgTit = body.appendParagraph('FAMILIOGRAMA');
  fgTit.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  fgTit.editAsText().setBold(true).setFontSize(14);
  body.appendParagraph('');

  // ── Imagen del canvas ─────────────────────────────────────
  var imgData = fg.imagenBase64 || fg.imagen || '';
  if (imgData) {
    try {
      var b64  = imgData.replace(/^data:image\/\w+;base64,/, '');
      var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', 'familiograma.png');
      var ipar = body.appendParagraph('');
      ipar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      ipar.appendInlineImage(blob).setWidth(480).setHeight(340);
    } catch(e) {
      body.appendParagraph('[Imagen del familiograma no disponible — abrir el acordeón antes de guardar]')
          .editAsText().setItalic(true).setForegroundColor('#888888').setFontSize(9);
    }
  } else {
    body.appendParagraph('[Imagen del familiograma no disponible — abrir el acordeón antes de guardar]')
        .editAsText().setItalic(true).setForegroundColor('#888888').setFontSize(9);
  }
  body.appendParagraph('');

  // ── Leyenda de símbolos + Tipo de familia ─────────────────
  // Replicando la tabla del formato original de 6 columnas
  var tblLey = body.appendTable([[
    '□ Hombre\n⊠ Muerte\n○ Mujer\n★ Caso índice',
    '── Matrimonio\n══ Separación\n┄┄ Unión Libre\n╪ Divorcio',
    '≈≈ Rel. Esporádica\n○─□ Hermanos\n○ ○ Gemelas\n▲ Aborto / Embarazo',
    '░░ Armonía\n══ Amistad Fuerte\n━━ Fusión\n∿∿ Hostil',
    '▓▓ Estresante\n▓▓ Muy Estresante\n·-·- Rompimiento\n- - Distancia',
    'TIPO DE FAMILIA:\n1. Nuclear biparental\n2. Nuclear monoparental\n3. Extenso biparental\n4. Extenso\n5. Compuesto biparental\n6. Compuesto monoparental\n7. Unipersonal'
  ]]);
  setTableBorders(tblLey, '#000000');
  for (var c2 = 0; c2 < 6; c2++) {
    tblLey.getRow(0).getCell(c2).editAsText().setFontSize(8);
    if (c2 === 5) tblLey.getRow(0).getCell(c2).editAsText().setBold(true);
  }
  body.appendParagraph('');

  // ── ANÁLISIS ──────────────────────────────────────────────
  body.appendParagraph('ANÁLISIS').editAsText().setBold(true).setFontSize(10);
  var analisis = fg.analisis || fg.obs || '';
  var tblAn = body.appendTable([[analisis]]);
  tblAn.getRow(0).getCell(0).setMinimumHeight(90);
  setTableBorders(tblAn, '#000000');
  tblAn.getRow(0).getCell(0).editAsText().setFontSize(10);

  // ── Pie de página ─────────────────────────────────────────
  var footer = doc.addFooter();
  footer.appendParagraph(
    'E.S.E San Martín  ·  Familiograma  ·  Resolución 2788  ·  Fecha: ' + (d.fecha||'')
  ).setAlignment(DocumentApp.HorizontalAlignment.CENTER)
   .editAsText().setFontSize(8).setForegroundColor('#888888');

  moverACarpeta(doc, carpeta);
  doc.saveAndClose();
  return doc.getUrl();
}

// ════════════════════════════════════════════════════════════
//  DOC 2 — ECOMAPA + APGAR FAMILIAR + ZARIT
// ════════════════════════════════════════════════════════════
function crearDocInstrumentos(carpeta, titulo, d, apg, zar, eco) {
  var doc  = DocumentApp.create('Instrumentos · ' + titulo);
  var body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  var header = doc.addHeader();
  var hPar   = header.appendParagraph('E.S.E SAN MARTIN');
  hPar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  hPar.editAsText().setBold(true).setFontSize(11);

  // ──────────────────────────────────────────────────────────
  //  SECCIÓN 1 — ECOMAPA
  // ──────────────────────────────────────────────────────────
  var titEco = body.appendParagraph('ECOMAPA');
  titEco.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titEco.editAsText().setBold(true).setFontSize(18).setForegroundColor('#005f73');
  body.appendParagraph('');

  // Imagen del ecomapa
  var ecoImg = eco.imagenBase64 || eco.imagen || '';
  if (ecoImg) {
    try {
      var b64e = ecoImg.replace(/^data:image\/\w+;base64,/, '');
      var blbe = Utilities.newBlob(Utilities.base64Decode(b64e), 'image/png', 'ecomapa.png');
      var epar = body.appendParagraph('');
      epar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      epar.appendInlineImage(blbe).setWidth(460).setHeight(380);
    } catch(e) {
      body.appendParagraph('[Imagen del ecomapa no disponible]')
          .editAsText().setItalic(true).setForegroundColor('#888888');
    }
  }
  body.appendParagraph('');

  // Tabla de valores del ecomapa
  var ecoVals = eco.valores || eco || {};
  var ECO_CATS = [
    ['eco-vias',           'Vías de acceso'],
    ['eco-fam-ext',        'Familia extensa'],
    ['eco-programas',      'Programas sociales y comunitarios'],
    ['eco-justicia',       'Protección en justicia'],
    ['eco-org-com',        'Organización comunitaria'],
    ['eco-vecinos',        'Vecinos'],
    ['eco-espiritualidad', 'Espiritualidad'],
    ['eco-educacion',      'Educación'],
    ['eco-trabajo',        'Trabajo'],
    ['eco-jac',            'Junta de acción comunal'],
    ['eco-recreacion',     'Recreación y esparcimiento'],
    ['eco-salud',          'Servicios de salud'],
    ['eco-recursos',       'Recursos económicos']
  ];
  var ECO_TEXTO = {
    '-2':'No aplica','1':'Fuerte','2':'Estresante',
    '3':'Moderada','4':'Mod. estresante','5':'Débil','6':'Lev. estresante'
  };
  var ecoRows = [['Sistema / Recurso', 'Valor', 'Interpretación']].concat(
    ECO_CATS.map(function(p){
      var v = String(ecoVals[p[0]] || '');
      return [p[1], v||'—', v ? (ECO_TEXTO[v]||v) : '—'];
    })
  );
  var tblEco = body.appendTable(ecoRows);
  styleDataTable(tblEco, true);

  if (eco.obs) {
    body.appendParagraph('');
    body.appendParagraph('Observaciones del Ecomapa:').editAsText().setBold(true).setFontSize(10);
    body.appendParagraph(eco.obs).editAsText().setFontSize(10);
  }
  body.appendParagraph('');
  body.appendParagraph('−2=No aplica  ·  1=Fuerte  ·  2=Estresante  ·  3=Moderada  ·  4=Mod. estresante  ·  5=Débil  ·  6=Lev. estresante')
      .editAsText().setFontSize(8).setItalic(true).setForegroundColor('#555555');

  // ──────────────────────────────────────────────────────────
  //  SECCIÓN 2 — APGAR FAMILIAR
  // ──────────────────────────────────────────────────────────
  body.appendPageBreak();
  var titApgar = body.appendParagraph('APGAR FAMILIAR');
  titApgar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titApgar.editAsText().setBold(true).setFontSize(18).setForegroundColor('#005f73');
  body.appendParagraph('');

  // Bloque de interpretación (igual al formato institucional)
  var tblInterp = body.appendTable([[
    'Interpretación del puntaje:\n' +
    '• Función familiar normal: 21–25 pts\n' +
    '• Función familiar leve: 16–20 pts\n' +
    '• Disfunción moderada: 11–15 pts\n' +
    '• Disfunción severa: 0–10 pts'
  ]]);
  setTableBorders(tblInterp, '#005f73');
  tblInterp.getRow(0).getCell(0).editAsText().setFontSize(9);
  body.appendParagraph('');

  // Tabla de miembros con puntuaciones
  var apgarMiembros = apg.miembros || apg.rows || [];
  var apgarHdrs = ['COD','Nombres y Apellidos','Sexo','Edad','P1','P2','P3','P4','P5','Total'];
  var apgarRows = [];
  if (Array.isArray(apgarMiembros) && apgarMiembros.length > 0) {
    apgarMiembros.forEach(function(r, i){
      apgarRows.push([
        String(i+1), r.nombre||'', r.sexo||'', r.edad||'',
        String(r.p1||r.preg1||''), String(r.p2||r.preg2||''),
        String(r.p3||r.preg3||''), String(r.p4||r.preg4||''),
        String(r.p5||r.preg5||''), String(r.total||'')
      ]);
    });
  } else {
    for (var i = 0; i < 6; i++) apgarRows.push(['','','','','','','','','','']);
  }
  var tblApgar = body.appendTable([apgarHdrs].concat(apgarRows));
  styleDataTable(tblApgar, true);

  body.appendParagraph('');
  body.appendParagraph('0=NUNCA  1=CASI NUNCA  2=ALGUNAS VECES  3=CASI SIEMPRE  4=SIEMPRE')
      .editAsText().setFontSize(8).setBold(true);
  body.appendParagraph('');
  body.appendParagraph('ANÁLISIS').editAsText().setBold(true).setFontSize(10);
  var tblAn = body.appendTable([[apg.obs || apg.analisis || '']]);
  tblAn.getRow(0).getCell(0).setMinimumHeight(60);
  setTableBorders(tblAn, '#cccccc');

  // ──────────────────────────────────────────────────────────
  //  SECCIÓN 3 — INSTRUMENTO ZARIT
  // ──────────────────────────────────────────────────────────
  body.appendPageBreak();
  var titZarit = body.appendParagraph('APLICACIÓN DEL INSTRUMENTO ZARIT');
  titZarit.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titZarit.editAsText().setBold(true).setFontSize(14).setForegroundColor('#005f73');
  body.appendParagraph('');
  body.appendParagraph(
    'A continuación se presenta una lista de afirmaciones. Indique la frecuencia: ' +
    '0=Nunca · 1=Rara vez · 2=Algunas veces · 3=Bastantes veces · 4=Casi siempre. ' +
    'La puntuación global oscila entre 0 y 88 puntos.'
  ).editAsText().setFontSize(9).setItalic(true);
  body.appendParagraph('');

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
  var zaritResps = zar.respuestas || zar.items || [];
  var zaritRows  = [['N°','PREGUNTA','0','1','2','3','4']].concat(
    ZARIT_PREGS.map(function(preg, i){
      var val = (zaritResps[i] !== undefined) ? String(zaritResps[i]) : '';
      return [String(i+1), preg,
        val==='0'?'✓':'', val==='1'?'✓':'', val==='2'?'✓':'',
        val==='3'?'✓':'', val==='4'?'✓':''];
    })
  );
  var tblZarit = body.appendTable(zaritRows);
  styleZaritTable(tblZarit);

  body.appendParagraph('');
  var puntaje   = zar.puntaje || 0;
  var interp    = puntaje <= 46 ? 'No sobrecarga' : puntaje <= 55 ? 'Sobrecarga leve' : 'Sobrecarga intensa';
  body.appendParagraph('Puntaje total: ' + puntaje + '    Interpretación: ' + interp)
      .editAsText().setBold(true).setFontSize(10);
  body.appendParagraph('');
  body.appendParagraph('INTERPRETACIÓN DE LOS RESULTADOS').editAsText().setBold(true).setFontSize(10);
  var tblIntZ = body.appendTable([[zar.obs || zar.interpretacion || '']]);
  tblIntZ.getRow(0).getCell(0).setMinimumHeight(60);
  setTableBorders(tblIntZ, '#cccccc');

  // Pie de página
  var footer = doc.addFooter();
  footer.appendParagraph('E.S.E San Martín  ·  Ecomapa + APGAR + Zarit  ·  ' + (d.fecha||''))
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .editAsText().setFontSize(8).setForegroundColor('#888888');

  moverACarpeta(doc, carpeta);
  doc.saveAndClose();
  return doc.getUrl();
}

// ── MACRO — una fila por integrante ─────────────────────────
function guardarMacro(sheet, d) {
  initEncabezados(sheet, [
    'ID','Fecha Registro','Municipio',
    'Fecha Visita','Vereda / Sector','Microterritorio','Auxiliar a Cargo',
    'Novedad','Familia','N° Casa','Teléfono',
    'Cargado','Seguimiento Telefónico Laboratorios','Seguimiento Médico',
    'Acueducto','Pozo / Alcantarillado','Unidad Sanitaria','Ducha',
    'Cocina','Manejo de Residuos','Paredes','Piso','Techo',
    'Electricidad','Riesgo en el Entorno','Nivel de Riesgo Vivienda',
    'Ingresos Económicos','SISBEN',
    'N° Integrante','Nombre Integrante','Tipo Documento','N° Documento','Edad',
    'Curso de Vida','Género','Orientación Sexual','Migrante',
    'Enfermedades Detectadas','De Novo','Adherencia',
    'Discapacidad','Tipo Discapacidad','Certificación Discapacidad','Cuidador',
    'EPS','Alertas','Profesional Requerido',
    'Víctima Conflicto Armado','Víctima Violencia Intrafamiliar',
    'Nivel Educativo','Contacto Integrante'
  ]);
  var viv  = d.vivienda   || {};
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
        idx+1,
        int.nombre||'', int.tipodoc||'', int.numdoc||'', int.edad||'',
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
//  ESTILOS DE TABLAS
// ════════════════════════════════════════════════════════════
function styleIdentTable(tbl) {
  setTableBorders(tbl, '#cccccc');
  for (var r = 0; r < tbl.getNumRows(); r++) {
    var row = tbl.getRow(r);
    row.getCell(0).editAsText().setBold(true).setFontSize(9).setForegroundColor('#005f73');
    if (row.getNumCells() > 1) row.getCell(1).editAsText().setFontSize(9);
  }
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
    var bg = (r % 2 === 0) ? '#f0f9ff' : '#ffffff';
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

function moverACarpeta(doc, carpeta) {
  var file = DriveApp.getFileById(doc.getId());
  carpeta.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
}

function respuesta(ok, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, mensaje: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
