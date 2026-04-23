/**
 * ============================================================
 *  EBS+ · GOOGLE APPS SCRIPT
 * ============================================================
 */

var TIPO_FORMULARIO  = 'recoleccion';
var NOMBRE_MUNICIPIO = 'Enciso';
var NOMBRE_HOJA      = 'Datos App';
var NOMBRE_CARPETA   = 'EBS+ Instrumentos';

function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    guardarDatos(datos);
    return respuesta(true, 'Guardado correctamente');
  } catch (err) {
    Logger.log('Error en doPost: ' + err.message);
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
    d.obs||'', d.foto ? 'Tiene foto' : 'Sin foto'
  ]);
}

function guardarRecoleccion(sheet, d) {
  initEncabezados(sheet, [
    'ID','Fecha','Municipio','Hospital','Vereda','Microterritorio','Teléfono Jefe Hogar',
    'FG Familia','FG Vereda','FG Jefe Nombre','FG Jefe Año','FG Jefe Estado','FG Jefe Enf',
    'FG Pareja Nombre','FG Pareja Año','FG Pareja Estado','FG Pareja Enf',
    'FG Relación Pareja','FG Hijos','FG Observaciones',
    'APG M1 Nombre','APG M1 Sexo','APG M1 Edad','APG M1 P1','APG M1 P2','APG M1 P3','APG M1 P4','APG M1 P5','APG M1 Total',
    'APG M2 Nombre','APG M2 Sexo','APG M2 Edad','APG M2 P1','APG M2 P2','APG M2 P3','APG M2 P4','APG M2 P5','APG M2 Total',
    'APG M3 Nombre','APG M3 Sexo','APG M3 Edad','APG M3 P1','APG M3 P2','APG M3 P3','APG M3 P4','APG M3 P5','APG M3 Total',
    'APG M4 Nombre','APG M4 Sexo','APG M4 Edad','APG M4 P1','APG M4 P2','APG M4 P3','APG M4 P4','APG M4 P5','APG M4 Total',
    'APG M5 Nombre','APG M5 Sexo','APG M5 Edad','APG M5 P1','APG M5 P2','APG M5 P3','APG M5 P4','APG M5 P5','APG M5 Total',
    'APG M6 Nombre','APG M6 Sexo','APG M6 Edad','APG M6 P1','APG M6 P2','APG M6 P3','APG M6 P4','APG M6 P5','APG M6 Total',
    'APG Análisis',
    'Zarit P1','Zarit P2','Zarit P3','Zarit P4','Zarit P5','Zarit P6','Zarit P7',
    'Zarit P8','Zarit P9','Zarit P10','Zarit P11','Zarit P12','Zarit P13','Zarit P14',
    'Zarit P15','Zarit P16','Zarit P17','Zarit P18','Zarit P19','Zarit P20','Zarit P21','Zarit P22',
    'Zarit Total','Zarit Interpretación','Zarit Observaciones',
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
  var jefe   = fg.principal || {};
  var pareja = pickPareja(fg);

  // ✅ CORRECCIÓN: Evitar duplicados en hijos
  var todosMiembros = fg.miembros || [];
  var hijosUnicos = [];
  var vistos = {};
  
  todosMiembros.forEach(function(m) {
    if (m.parentesco === 'hijo' && !vistos[m.nombre]) {
      hijosUnicos.push(m);
      vistos[m.nombre] = true;
    }
  });
  
  var hijos = hijosUnicos.map(function(h){
    return (h.nombre||'?') + ' (' + (h.anio||'?') + '/' + (h.genero||'?') + ')';
  }).join(' | ');

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

  var ecoVals = eco.valores || eco || {};
  var ECO_IDS = ['eco-vias','eco-fam-ext','eco-programas','eco-justicia','eco-org-com',
                 'eco-vecinos','eco-espiritualidad','eco-educacion','eco-trabajo',
                 'eco-jac','eco-recreacion','eco-salud','eco-recursos'];
  var ecoCols = ECO_IDS.map(function(id){ return ecoVals[id] || ''; });
  ecoCols.push(eco.obs || '');

  var carpeta = obtenerCarpeta();
  var urlDoc  = '';
  try { 
    urlDoc = crearDocInstrumentos(carpeta, d, fg, apg, zar, eco); 
    Logger.log('Documento creado: ' + urlDoc);
  } catch(e) { 
    Logger.log('Error creando documento: ' + e.message);
    urlDoc = 'Error: ' + e.message; 
  }

  var fila = [
    d.id||'', d.fecha||'', d.municipio||NOMBRE_MUNICIPIO,
    d.hospital||'', d.vereda||'', d.microterritorio||'', d.telefono||'',
    fg.familia||'',
    fg.vereda||'',
    jefe.nombre||'', jefe.anio||'', jefe.estado||'', jefe.enf||'',
    pareja.nombre||'', pareja.anio||'', pareja.estado||'', pareja.enf||'',
    fg.relacion || fg.relacionPareja || '',
    hijos, fg.obs||''
  ].concat(apgarCols).concat(zaritCols).concat(ecoCols).concat([urlDoc]);

  sheet.appendRow(fila);
}

function crearDocInstrumentos(carpeta, d, fg, apg, zar, eco) {
  var jefe       = fg.principal || {};
  var jefeNombre = jefe.nombre  || 'SinNombre';
  var famNombre  = fg.familia   || d.familia || d.id || 'Familia';
  
  var nombreDoc  = limpiarNombre(jefeNombre) + ' - ' + limpiarNombre(famNombre);

  var doc  = DocumentApp.create(nombreDoc);
  var body = doc.getBody();
  
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);

  try {
    // ENCABEZADO INSTITUCIONAL
    var nombreHospital = d.hospital || 'E.S.E SAN MARTÍN';
    var headerTable = body.appendTable([
      ['', nombreHospital, ''],
      ['EQUIPOS BÁSICOS DE SALUD - EBS+', 'MUNICIPIO DE ' + (d.municipio || NOMBRE_MUNICIPIO).toUpperCase(), 'Instrumento de Recolección']
    ]);

    // Celda izquierda fila 0: Logo Hospital
    var celdaLogoHosp = headerTable.getRow(0).getCell(0);
    celdaLogoHosp.setBackgroundColor('#ffffff');
    if (d.logoHospital && d.logoHospital.length > 100) {
      try {
        var b64Hosp = d.logoHospital.indexOf(',') !== -1 ? d.logoHospital.split(',')[1] : d.logoHospital;
        var blobHosp = Utilities.newBlob(Utilities.base64Decode(b64Hosp), 'image/png', 'logo_hospital.png');
        var imgHosp = body.getImages ? null : null; // placeholder
        celdaLogoHosp.insertImage(0, blobHosp);
      } catch(eImg) {
        celdaLogoHosp.editAsText().setText('LOGO\nHOSPITAL').setFontSize(7).setForegroundColor('#005f73');
      }
    } else {
      celdaLogoHosp.editAsText().setText('').setFontSize(8);
    }

    // Celda central fila 0: Nombre del hospital
    headerTable.getRow(0).getCell(1).setBackgroundColor('#005f73');
    headerTable.getRow(0).getCell(1).editAsText().setBold(true).setFontSize(14).setForegroundColor('#ffffff');

    // Celda derecha fila 0: Logo EBS
    var celdaLogoEbs = headerTable.getRow(0).getCell(2);
    celdaLogoEbs.setBackgroundColor('#ffffff');
    if (d.logoEbs && d.logoEbs.length > 100) {
      try {
        var b64Ebs = d.logoEbs.indexOf(',') !== -1 ? d.logoEbs.split(',')[1] : d.logoEbs;
        var blobEbs = Utilities.newBlob(Utilities.base64Decode(b64Ebs), 'image/png', 'logo_ebs.png');
        celdaLogoEbs.insertImage(0, blobEbs);
      } catch(eImg2) {
        celdaLogoEbs.editAsText().setText('LOGO\nEBS').setFontSize(7).setForegroundColor('#005f73');
      }
    } else {
      celdaLogoEbs.editAsText().setText('').setFontSize(8);
    }
    
    headerTable.getRow(1).getCell(0).setBackgroundColor('#0a9396');
    headerTable.getRow(1).getCell(0).editAsText().setBold(true).setFontSize(10).setForegroundColor('#ffffff');
    
    headerTable.getRow(1).getCell(1).setBackgroundColor('#0a9396');
    headerTable.getRow(1).getCell(1).editAsText().setBold(true).setFontSize(10).setForegroundColor('#ffffff');
    
    headerTable.getRow(1).getCell(2).setBackgroundColor('#0a9396');
    headerTable.getRow(1).getCell(2).editAsText().setBold(true).setFontSize(10).setForegroundColor('#ffffff');
    
    headerTable.setBorderWidth(1);
    headerTable.setBorderColor('#005f73');
    
    body.appendParagraph('');

    // DATOS DE IDENTIFICACIÓN
    var identTitle = body.appendParagraph('DATOS DE IDENTIFICACIÓN');
    identTitle.editAsText().setBold(true).setFontSize(12).setForegroundColor('#005f73');
    body.appendParagraph('');

    var identTable = body.appendTable([
      ['Familia:', famNombre, 'Vereda/Sector:', fg.vereda || d.vereda || ''],
      ['Municipio:', d.municipio || NOMBRE_MUNICIPIO, 'Microterritorio:', d.microterritorio || ''],
      ['Jefe del Hogar:', jefe.nombre || '', 'Teléfono:', d.telefono || jefe.telefono || ''],
      ['Fecha de Registro:', d.fecha || '', 'N° Familia:', d.familia || d.id || '']
    ]);
    
    identTable.setBorderWidth(1);
    identTable.setBorderColor('#005f73');
    
    for (var ri = 0; ri < identTable.getNumRows(); ri++) {
      for (var ci = 0; ci < identTable.getRow(ri).getNumCells(); ci++) {
        var cell = identTable.getRow(ri).getCell(ci);
        cell.setPaddingBottom(5);
        cell.setPaddingTop(5);
        cell.setPaddingLeft(7);
        cell.setPaddingRight(7);
        if (ci % 2 === 0) {
          cell.setBackgroundColor('#e0f2fe');
          cell.editAsText().setBold(true).setFontSize(9);
        } else {
          cell.editAsText().setFontSize(9);
        }
      }
    }
    
    body.appendParagraph('');

    // FAMILIOGRAMA
    body.appendPageBreak();
    
    var section1Title = body.appendParagraph('FAMILIOGRAMA FAMILIAR');
    section1Title.editAsText().setBold(true).setFontSize(13).setForegroundColor('#005f73');
    body.appendParagraph('');

    var miembrosTableData = [['INTEGRANTE', 'NOMBRE', 'AÑO NAC.', 'ESTADO DE SALUD', 'ENFERMEDAD/NOTA']];
    
    // Jefe del hogar
    if (jefe.nombre) {
      miembrosTableData.push([
        'Jefe del Hogar',
        jefe.nombre || '—',
        jefe.anio || '—',
        jefe.estado || '—',
        jefe.enf || '—'
      ]);
    }
    
    // Pareja (sin duplicar)
    var pareja = pickPareja(fg);
    if (pareja.nombre && !esMiembroRepetido(pareja.nombre, miembrosTableData)) {
      miembrosTableData.push([
        'Pareja',
        pareja.nombre || '—',
        pareja.anio || '—',
        pareja.estado || '—',
        pareja.enf || '—'
      ]);
    }
    
    // ✅ CORRECCIÓN: Procesar miembros sin duplicados
    var todosMiembros = fg.miembros || [];
    var miembrosProcesados = {};
    
    todosMiembros.forEach(function(m) {
      // Evitar duplicados por nombre
      if (!m.nombre || miembrosProcesados[m.nombre]) return;
      miembrosProcesados[m.nombre] = true;
      
      var etiqueta = obtenerEtiquetaParentesco(m.parentesco);
      miembrosTableData.push([
        etiqueta,
        m.nombre || '—',
        m.anio || '—',
        m.estado || '—',
        m.nota || '—'
      ]);
    });

    var miembrosTable = body.appendTable(miembrosTableData);
    styleProfessionalTable(miembrosTable, true);
    
    if (fg.relacionPareja || fg.relacion) {
      body.appendParagraph('');
      var relPar = body.appendParagraph('Tipo de relación de pareja: ' + (fg.relacionPareja || fg.relacion));
      relPar.editAsText().setFontSize(9).setItalic(true);
    }

    body.appendParagraph('');
    body.appendParagraph('DIAGRAMA DEL FAMILIOGRAMA').editAsText().setBold(true).setFontSize(10).setForegroundColor('#005f73');
    
    insertarImagenRedimensionada(body, fg.imagenBase64||fg.imagen||'', 450, 350, '[Imagen del familiograma no disponible]');
    
    if (fg.obs || fg.analisis) {
      body.appendParagraph('');
      body.appendParagraph('OBSERVACIONES / ANÁLISIS').editAsText().setBold(true).setFontSize(10).setForegroundColor('#005f73');
      cajaNota(body, fg.analisis || fg.obs || '', 80);
    }

    // ECOMAPA
    body.appendPageBreak();
    
    var section2Title = body.appendParagraph('ECOMAPA FAMILIAR');
    section2Title.editAsText().setBold(true).setFontSize(13).setForegroundColor('#0e7490');
    body.appendParagraph('');

    body.appendParagraph('DIAGRAMA DEL ECOMAPA').editAsText().setBold(true).setFontSize(10).setForegroundColor('#0e7490');
    insertarImagenRedimensionada(body, eco.imagenBase64||eco.imagen||'', 450, 400, '[Imagen del ecomapa no disponible]');

    body.appendParagraph('');
    body.appendParagraph('VALORACIÓN DE SISTEMAS Y RECURSOS').editAsText().setBold(true).setFontSize(10).setForegroundColor('#0e7490');
    
    var ECO_CATS = [
      ['eco-vias','Vías de acceso'],
      ['eco-fam-ext','Familia extensa'],
      ['eco-programas','Programas sociales y comunitarios'],
      ['eco-justicia','Protección en justicia'],
      ['eco-org-com','Organización comunitaria'],
      ['eco-vecinos','Vecinos'],
      ['eco-espiritualidad','Espiritualidad'],
      ['eco-educacion','Educación'],
      ['eco-trabajo','Trabajo'],
      ['eco-jac','Junta de acción comunal'],
      ['eco-recreacion','Recreación y esparcimiento'],
      ['eco-salud','Servicios de salud'],
      ['eco-recursos','Recursos económicos']
    ];
    
    var ECO_TX = {
      '-2':'No aplica',
      '1':'Fuerte',
      '2':'Estresante',
      '3':'Moderada',
      '4':'Mod. estresante',
      '5':'Débil',
      '6':'Lev. estresante'
    };

    var ecoTableData = [['SISTEMA / RECURSO', 'VALOR', 'INTERPRETACIÓN']];
    var ecoVals = eco.valores || eco || {};
    
    ECO_CATS.forEach(function(p){
      var v = String(ecoVals[p[0]]||'');
      ecoTableData.push([p[1], v||'—', v ? (ECO_TX[v]||v) : '—']);
    });

    var ecoTable = body.appendTable(ecoTableData);
    styleProfessionalTable(ecoTable, true);
    
    if (eco.obs) {
      body.appendParagraph('');
      body.appendParagraph('OBSERVACIONES').editAsText().setBold(true).setFontSize(10).setForegroundColor('#0e7490');
      cajaNota(body, eco.obs, 60);
    }

    // APGAR
    body.appendPageBreak();
    
    var section3Title = body.appendParagraph('APGAR FAMILIAR');
    section3Title.editAsText().setBold(true).setFontSize(13).setForegroundColor('#065f46');
    body.appendParagraph('');

    var refBox = body.appendTable([[
      'INTERPRETACIÓN DEL PUNTAJE:\n\n' +
      'Función familiar normal: 21-25 puntos\n' +
      'Disfunción leve: 16-20 puntos\n' +
      'Disfunción moderada: 11-15 puntos\n' +
      'Disfunción severa: 0-10 puntos\n\n' +
      'ESCALA: 0=Nunca, 1=Casi nunca, 2=Algunas veces, 3=Casi siempre, 4=Siempre, 5=Siempre'
    ]]);
    refBox.getRow(0).getCell(0).setBackgroundColor('#ecfdf5');
    refBox.getRow(0).getCell(0).editAsText().setFontSize(9);
    refBox.setBorderWidth(1);
    refBox.setBorderColor('#065f46');
    body.appendParagraph('');

    body.appendParagraph('PREGUNTAS DE VALORACIÓN:').editAsText().setBold(true).setFontSize(10).setForegroundColor('#065f46');
    var preguntasList = body.appendParagraph('');
    preguntasList.appendText('1. ¿Se siente satisfecho con la ayuda que recibe de su familia cuando tiene algún problema o necesidad?\n');
    preguntasList.appendText('2. ¿Se siente satisfecho con la forma en que su familia habla de las cosas y comparten sus problemas?\n');
    preguntasList.appendText('3. ¿Se siente satisfecho con la forma en que su familia acepta y apoya sus deseos de emprender nuevas actividades?\n');
    preguntasList.appendText('4. ¿Se siente satisfecho con la forma en que su familia expresa afecto y responde a sus emociones?\n');
    preguntasList.appendText('5. ¿Se siente satisfecho con la manera como comparten en su familia el tiempo, espacios o dinero?');
    preguntasList.editAsText().setFontSize(8).setForegroundColor('#64748b');
    body.appendParagraph('');

    var apgarMbs = apg.miembros || [];
    var apgarHdrs = ['#', 'NOMBRES Y APELLIDOS', 'SEXO', 'EDAD', 'P.1', 'P.2', 'P.3', 'P.4', 'P.5', 'TOTAL', 'FUNCIÓN'];
    var apgarRows = [apgarHdrs];
    
    if (apgarMbs.length > 0) {
      apgarMbs.forEach(function(r, i){
        var tot = (r.total !== null && r.total !== undefined) ? parseInt(r.total) : null;
        var interp = tot !== null ? (
          tot >= 21 ? 'Normal' : tot >= 16 ? 'Dis. leve' : tot >= 11 ? 'Dis. moderada' : 'Dis. severa'
        ) : '—';
        apgarRows.push([
          String(i+1),
          r.nombres||r.nombre||'',
          r.sexo||'',
          String(r.edad||''),
          valOrEmpty(r.p1), valOrEmpty(r.p2), valOrEmpty(r.p3),
          valOrEmpty(r.p4), valOrEmpty(r.p5),
          tot !== null ? String(tot) : '—',
          interp
        ]);
      });
    } else {
      apgarRows.push(['', '', '', '', '', '', '', '', '', '', '']);
    }

    var apgarTable = body.appendTable(apgarRows);
    styleProfessionalTable(apgarTable, true);
    
    if (apg.analisis) {
      body.appendParagraph('');
      body.appendParagraph('ANÁLISIS DEL APGAR FAMILIAR').editAsText().setBold(true).setFontSize(10).setForegroundColor('#065f46');
      cajaNota(body, apg.analisis, 80);
    }

    // ZARIT
    body.appendPageBreak();
    
    var section4Title = body.appendParagraph('ESCALA DE ZARIT');
    section4Title.editAsText().setBold(true).setFontSize(13).setForegroundColor('#7c3aed');
    body.appendParagraph('');
    
    var zaritSub = body.appendParagraph('Instrumento de valoración de sobrecarga del cuidador');
    zaritSub.editAsText().setItalic(true).setFontSize(9).setForegroundColor('#64748b');
    body.appendParagraph('');

    var zaritRefBox = body.appendTable([[
      'INTERPRETACIÓN DE RESULTADOS:\n\n' +
      '0-20 puntos: Carga ausente o leve\n' +
      '21-40 puntos: Carga leve a moderada\n' +
      '41-60 puntos: Carga moderada a relevante\n' +
      '61-88 puntos: Carga severa - Valorar apoyo\n\n' +
      'ESCALA: 0=Nunca, 1=Rara vez, 2=Algunas veces, 3=Bastantes veces, 4=Casi siempre'
    ]]);
    zaritRefBox.getRow(0).getCell(0).setBackgroundColor('#ede9fe');
    zaritRefBox.getRow(0).getCell(0).editAsText().setFontSize(9);
    zaritRefBox.setBorderWidth(1);
    zaritRefBox.setBorderColor('#7c3aed');
    body.appendParagraph('');

    var ZARIT_PREGS = [
      '¿Cree que su familiar solicita más ayuda de la que realmente necesita?',
      '¿Cree que debido al tiempo que dedica a su familiar ya no dispone de tiempo suficiente para usted?',
      '¿Se siente agobiado entre cuidar a su familiar y atender además otras responsabilidades?',
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
      '¿Cree que su familiar parece esperar que usted sea la única persona que le cuide?',
      '¿Cree que no tiene suficiente dinero para cuidar a su familiar además de sus otros gastos?',
      '¿Cree que será incapaz de cuidarle/a por mucho más tiempo?',
      '¿Siente que ha perdido el control de su vida desde la enfermedad de su familiar?',
      '¿Desearía poder dejar el cuidado de su familiar a otros?',
      '¿Se siente indeciso sobre qué hacer con su familiar?',
      '¿Cree que debería hacer más por su familiar?',
      '¿Cree que podría cuidar mejor de su familiar?',
      'Globalmente, ¿qué grado de carga experimenta por el hecho de cuidar a su familiar?'
    ];

    var zaritTableData = [['N°', 'PREGUNTA', '0', '1', '2', '3', '4']];
    var zaritResps = zar.respuestas || [];
    
    ZARIT_PREGS.forEach(function(preg, i){
      var val = (zaritResps[i] !== null && zaritResps[i] !== undefined) ? String(zaritResps[i]) : '';
      zaritTableData.push([
        String(i+1),
        preg,
        val==='0'?'X':'',
        val==='1'?'X':'',
        val==='2'?'X':'',
        val==='3'?'X':'',
        val==='4'?'X':''
      ]);
    });

    var zaritTable = body.appendTable(zaritTableData);
    styleZaritTable(zaritTable);
    
    body.appendParagraph('');
    var puntaje = zar.puntaje || 0;
    var interp  = puntaje <= 20 ? 'Carga ausente o leve (0-20)' 
                : puntaje <= 40 ? 'Carga leve a moderada (21-40)'
                : puntaje <= 60 ? 'Carga moderada a relevante (41-60)'
                : 'Carga severa (61-88) - Valorar apoyo';

    var resultTable = body.appendTable([['PUNTAJE TOTAL: ' + puntaje + ' / 88     |     INTERPRETACIÓN: ' + interp]]);
    var bgRes = puntaje <= 20 ? '#dcfce7' 
                : puntaje <= 40 ? '#fef3c7' 
                : puntaje <= 60 ? '#fed7aa' 
                : '#fee2e2';
    resultTable.getRow(0).getCell(0)
      .setBackgroundColor(bgRes)
      .editAsText().setBold(true).setFontSize(11).setForegroundColor('#1e1e1e');
    resultTable.setBorderWidth(1);
    resultTable.setBorderColor('#7c3aed');

    if (zar.obs) {
      body.appendParagraph('');
      body.appendParagraph('OBSERVACIONES').editAsText().setBold(true).setFontSize(10).setForegroundColor('#7c3aed');
      cajaNota(body, zar.obs, 80);
    }

    // PIE DE PÁGINA
    body.appendPageBreak();
    
    var footerTable = body.appendTable([
      [(d.hospital || 'E.S.E San Martín') + ' - Equipos Básicos de Salud - EBS+', 'Fecha: ' + (d.fecha||'')],
      ['Municipio: ' + (d.municipio || NOMBRE_MUNICIPIO), 'Documento generado automáticamente']
    ]);
    
    footerTable.setBorderWidth(1);
    footerTable.setBorderColor('#e2e8f0');
    
    for (var fi = 0; fi < footerTable.getNumRows(); fi++) {
      for (var fc = 0; fc < footerTable.getRow(fi).getNumCells(); fc++) {
        var fCell = footerTable.getRow(fi).getCell(fc);
        fCell.setBackgroundColor('#f8fafc');
        fCell.setPaddingBottom(5);
        fCell.setPaddingTop(5);
        fCell.setPaddingLeft(7);
        fCell.setPaddingRight(7);
        fCell.editAsText().setFontSize(8).setForegroundColor('#64748b');
      }
    }

    moverACarpeta(doc, carpeta);
    doc.saveAndClose();
    
    Logger.log('Documento finalizado: ' + doc.getUrl());
    return doc.getUrl();
    
  } catch (e) {
    Logger.log('ERROR en crearDocInstrumentos: ' + e.message + ' | ' + e.stack);
    body.appendParagraph('ERROR AL GENERAR EL DOCUMENTO: ' + e.message);
    moverACarpeta(doc, carpeta);
    doc.saveAndClose();
    return doc.getUrl();
  }
}

// ════════════════════════════════════════════════════════════
//  FUNCIONES AUXILIARES
// ════════════════════════════════════════════════════════════

// ✅ NUEVA: Verificar si un miembro ya está en la tabla
function esMiembroRepetido(nombre, tablaData) {
  for (var i = 1; i < tablaData.length; i++) {
    if (tablaData[i][1] === nombre) return true;
  }
  return false;
}

// ✅ NUEVA: Obtener etiqueta legible para parentesco
function obtenerEtiquetaParentesco(parentesco) {
  var etiquetas = {
    'pareja': 'Pareja',
    'padre_jefe': 'Padre del Jefe',
    'madre_jefe': 'Madre del Jefe',
    'padre_pareja': 'Padre de la Pareja',
    'madre_pareja': 'Madre de la Pareja',
    'hijo': 'Hijo/a',
    'hermano': 'Hermano/a',
    'otro': 'Otro familiar'
  };
  return etiquetas[parentesco] || parentesco || 'Familiar';
}

function limpiarNombre(nombre) {
  if (!nombre) return 'SinNombre';
  return nombre
    .toString()
    .replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ\s-]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 50);
}

function insertarImagenRedimensionada(body, base64, w, h, fallback) {
  if (!base64 || base64.length < 100) {
    body.appendParagraph(fallback)
        .editAsText().setItalic(true).setForegroundColor('#888888').setFontSize(9);
    return;
  }
  
  try {
    var clean = base64.replace(/^data:image\/\w+;base64,/, '');
    
    if (clean.length > 700000) {
      Logger.log('Imagen demasiado grande, omitiendo: ' + clean.length + ' bytes');
      body.appendParagraph('Imagen muy grande (omitida para evitar errores)')
          .editAsText().setItalic(true).setForegroundColor('#888888').setFontSize(9);
      return;
    }
    
    var blob  = Utilities.newBlob(Utilities.base64Decode(clean), 'image/png', 'img.png');
    var par   = body.appendParagraph('');
    par.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    par.appendInlineImage(blob).setWidth(w).setHeight(h);
    
  } catch(e) {
    Logger.log('Error insertando imagen: ' + e.message);
    body.appendParagraph(fallback)
        .editAsText().setItalic(true).setForegroundColor('#888888').setFontSize(9);
  }
}

function styleProfessionalTable(tbl, hasHeader) {
  tbl.setBorderWidth(1);
  tbl.setBorderColor('#005f73');
  
  if (hasHeader) {
    var hr = tbl.getRow(0);
    for (var c = 0; c < hr.getNumCells(); c++) {
      hr.getCell(c).setBackgroundColor('#005f73');
      hr.getCell(c).editAsText().setBold(true).setFontSize(9).setForegroundColor('#ffffff');
      hr.getCell(c).setPaddingBottom(6);
      hr.getCell(c).setPaddingTop(6);
    }
  }
  
  for (var r = hasHeader ? 1 : 0; r < tbl.getNumRows(); r++) {
    var bg = r % 2 === 0 ? '#f0f9ff' : '#ffffff';
    for (var c2 = 0; c2 < tbl.getRow(r).getNumCells(); c2++) {
      var cell = tbl.getRow(r).getCell(c2);
      cell.setBackgroundColor(bg);
      cell.editAsText().setFontSize(9);
      cell.setPaddingBottom(4);
      cell.setPaddingTop(4);
      cell.setPaddingLeft(6);
      cell.setPaddingRight(6);
    }
  }
}

function styleZaritTable(tbl) {
  tbl.setBorderWidth(1);
  tbl.setBorderColor('#7c3aed');
  
  var hr = tbl.getRow(0);
  for (var c = 0; c < hr.getNumCells(); c++) {
    hr.getCell(c).setBackgroundColor('#7c3aed');
    hr.getCell(c).editAsText().setBold(true).setFontSize(8).setForegroundColor('#ffffff');
  }
  
  for (var r = 1; r < tbl.getNumRows(); r++) {
    var bg  = r % 2 === 0 ? '#f5f3ff' : '#ffffff';
    var row = tbl.getRow(r);
    
    row.getCell(0).setBackgroundColor('#ede9fe').editAsText().setBold(true).setFontSize(8);
    row.getCell(1).setBackgroundColor(bg).editAsText().setFontSize(8);
    
    for (var c2 = 2; c2 < row.getNumCells(); c2++) {
      var cell = row.getCell(c2);
      cell.setBackgroundColor(bg);
      cell.editAsText().setFontSize(9).setForegroundColor('#7c3aed');
    }
  }
}

function cajaNota(body, texto, minH) {
  var tbl = body.appendTable([[texto || '']]);
  tbl.setBorderWidth(1);
  tbl.setBorderColor('#005f73');
  var cell = tbl.getRow(0).getCell(0);
  cell.setBackgroundColor('#f8fafc');
  cell.editAsText().setFontSize(9);
  cell.setPaddingBottom(8);
  cell.setPaddingTop(8);
  cell.setPaddingLeft(10);
  cell.setPaddingRight(10);
  
  if (minH > 60) {
    for (var i = 0; i < 3; i++) {
      body.appendParagraph(' ');
    }
  }
}

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
  if (folders.hasNext()) {
    Logger.log('Carpeta encontrada: ' + NOMBRE_CARPETA);
    return folders.next();
  }
  Logger.log('Creando carpeta: ' + NOMBRE_CARPETA);
  return DriveApp.createFolder(NOMBRE_CARPETA);
}

function moverACarpeta(doc, carpeta) {
  try {
    var file = DriveApp.getFileById(doc.getId());
    var parents = file.getParents();
    
    while (parents.hasNext()) {
      var parent = parents.next();
      if (parent.getName() !== 'Mi Unidad' && parent.getName() !== 'My Drive') {
        parent.removeFile(file);
      }
    }
    
    carpeta.addFile(file);
    Logger.log('Documento movido a carpeta: ' + carpeta.getName());
  } catch (e) {
    Logger.log('Error moviendo documento: ' + e.message);
  }
}

function respuesta(ok, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, mensaje: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
