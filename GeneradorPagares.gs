/*************************************************************************************
 * GENERADOR DE PAGARÉS — FINANCIERA CUALLI  ·  v1.1
 * -------------------------------------------------------------------------------------
 * Extensión del proyecto "Generador de Cuadro de Dispersiones".
 *
 * Reutiliza helpers existentes del proyecto:
 *   inferirTrato_, inferirArticulo_, formatearNombrePersona_, formatearNombreEmpresa_,
 *   cmToPt, toast_, openLink_, showError_, NUM_TEXTO (custom function)
 *
 * Cambios v1.1:
 *   [FIX] Bug "Document is closed" al construir header/footer
 *   [NUEVO] Preview modal HTML con vista previa completa del pagaré
 *   [NUEVO] Log de pagarés generados en hoja oculta _LOG_PAGARES
 *   [NUEVO] Regenerar: detectar pagaré existente y ofrecer sobrescribir
 *************************************************************************************/


// =====================================================================================
// CONFIGURACIÓN
// =====================================================================================
const CONFIG_PAGARES = {
  FOLDER_ID_PAGARES: '1t79xliGsBgTb_Zh6K8h9OlLwIvKCZs0A',
  HOJAS_VALIDAS: ['PAGARE_1', 'PAGARE_2'],
  HOJA_LOG: '_LOG_PAGARES',
  DOMICILIO_CUALLI: 'Boulevard Manuel Ávila Camacho número 184, piso 7 "A", ' +
                    'Colonia Reforma Social, Alcaldía Miguel Hidalgo, ' +
                    'Ciudad de México, C.P. 11650',
  COLOR_AMARILLO: '#fbb818',
  COLOR_AMARILLO_TENUE: '#FEF9E7',
  COLOR_GRIS:     '#515151'
};

const CELDAS_PAGARE = {
  REGIMEN:          'C8',
  TIPO_FIRMA:       'E8',
  NUM_AVALES:       'C10',
  CODIGO_PLANTILLA: 'E10',

  ID_OPORTUNIDAD:   'C14',
  MONEDA:           'E14',
  MONTO_NUM:        'C16',
  TASA_NUM:         'E16',
  MONTO_LETRA:      'C18',
  TASA_LETRA:       'C20',
  FECHA_FIRMA:      'C22',
  FECHA_LETRA:      'E22',
  PLAZO_NUM:        'C24',
  PLAZO_UNIDAD:     'E24',
  FECHA_VENC:       'C26',
  FECHA_VENC_LETRA: 'E26',


  NOMBRE_SUSC:      'C26',
  TRATO_SUSC_INFO:  'C27',
  DOMI_SUSC:        'C28',
  REP_LEGAL_1:      'C30',
  TRATO_RL1_INFO:   'C31',
  REP_LEGAL_2:      'C32',
  TRATO_RL2_INFO:   'C33',

  AVALES_START_ROW: 37
};


// =====================================================================================
// MENÚ
// =====================================================================================



// =====================================================================================
// BOOTSTRAP — DISEÑO "DASHBOARD PREMIUM" (FORMATO CONDICIONAL INTELIGENTE)
// =====================================================================================
function crearHojasPagare() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const hojaActiva = ss.getActiveSheet();
  const nombreActual = hojaActiva.getName();

  // 1. Selector de alcance con botones nativos
  const res = ui.alert('🛠️ Mantenimiento del Sistema',
    '¿Qué alcance deseas para la restauración?\n\n' +
    '• [SÍ]: Regenerar TODAS las hojas (' + CONFIG_PAGARES.HOJAS_VALIDAS.join(', ') + ').\n' +
    '• [NO]: Regenerar SOLO la hoja actual: "' + nombreActual + '".\n' +
    '• ❌[CANCELAR]: No realizar cambios.',
    ui.ButtonSet.YES_NO_CANCEL);

  if (res === ui.Button.CANCEL) return;

  if (res === ui.Button.YES) {
    // --- LÓGICA PARA TODAS LAS HOJAS ---
    const confirmAll = ui.alert('⚠️ Confirmación Total', 
      'Esto borrará los datos de TODAS las hojas de pagarés. ¿Continuar?', 
      ui.ButtonSet.YES_NO);
    if (confirmAll !== ui.Button.YES) return;

    CONFIG_PAGARES.HOJAS_VALIDAS.forEach(nombre => {
      let sh = ss.getSheetByName(nombre);
      if (sh) {
        let idx = sh.getIndex();
        ss.deleteSheet(sh);
        sh = ss.insertSheet(nombre, idx - 1);
      } else {
        sh = ss.insertSheet(nombre);
      }
      construirHojaPagare_(sh);
    });
    sincronizarListaOportunidadesC15();
    obtenerOcreaHojaLog_();
    ui.alert('✨ Sistema Restaurado', 'Se han regenerado todas las interfaces con éxito.', ui.ButtonSet.OK);

  } else {
    // --- LÓGICA PARA HOJA ACTUAL ---
    if (!CONFIG_PAGARES.HOJAS_VALIDAS.includes(nombreActual)) {
      ui.alert('❌ Error de Ubicación', 
        'No estás en una hoja de pagarés válida. Sitúate en PAGARE_1 o PAGARE_2.', 
        ui.ButtonSet.OK);
      return;
    }

    let idx = hojaActiva.getIndex();
    ss.deleteSheet(hojaActiva);
    let nuevaHoja = ss.insertSheet(nombreActual, idx - 1);
    construirHojaPagare_(nuevaHoja);
    sincronizarListaOportunidadesC15();
    obtenerOcreaHojaLog_();
    ui.alert('✨ Hoja Actualizada', `La hoja "${nombreActual}" ha sido regenerada individualmente.`, ui.ButtonSet.OK);
  }
}
  // ═══════════════════════════════════════════════════════════════════════════
  // SET HOJA
  // ═══════════════════════════════════════════════════════════════════════════
function construirHojaPagare_(sh) {
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, 100, 26).setBackground('#E6E6E6').setFontFamily('Arial');

  // Anchos de columna
  const anchos = [25, 190, 310, 190, 310, 25];
  anchos.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 1 — BANNER PRINCIPAL
  // ═══════════════════════════════════════════════════════════════════════════
sh.setRowHeight(1, 48); // Forzamos el alto de la fila

  // Celda B1: Contenedor Blanco para el Logo
  sh.getRange('B1').setBackground('#FFFFFF')
    .setBorder(true, true, true, false, null, null, '#FDB913', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Celdas C1 a E1: Título con Fondo Gris
  sh.getRange('C1:E1').merge()
    .setValue('SISTEMA GENERADOR DE PAGARÉS 📝')
    .setBackground('#515151').setFontColor('#FDB913').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, false, true, false, null, null, '#FDB913', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Inyección y ajuste exacto del Logo
  const urlLogo = "https://cualli.mx/wp-content/uploads/2022/07/cualli-bl@3x.png";
  const logo = sh.insertImage(urlLogo, 2, 1); // Insertar en Columna 2 (B), Fila 1
  
  // Ajuste matemático para que mida lo mismo que la fila (48px)
  const altoFila = 48;
  const anchoOriginal = logo.getWidth();
  const altoOriginal = logo.getHeight();
  
  // Calculamos el nuevo ancho manteniendo la proporción para que el alto sea 48
  const nuevoAlto = altoFila - 4; // Restamos 4px de margen para que no toque los bordes
  const nuevoAncho = (nuevoAlto * anchoOriginal) / altoOriginal;

  logo.setHeight(nuevoAlto).setWidth(nuevoAncho);
  
  // Centramos el logo dentro de la columna B (que mide 190px en tu config)
  const anchoColumnaB = 190;
  const offsetX = (anchoColumnaB - nuevoAncho) / 2;
  const offsetY = 2; // Pequeño margen superior
  logo.setAnchorCellXOffset(offsetX).setAnchorCellYOffset(offsetY);
  // ═══════════════════════════════════════════════════════════════════════════
  // FILAS 2-7 — GUÍA DE USO (INLINE) — Estilo dashboard ejecutivo
  // ═══════════════════════════════════════════════════════════════════════════

  // Fila 2: Título de la guía
  sh.getRange('B2:E2').merge()
    .setValue('GUÍA DE USO Y RECOMENDACIONES')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor('#515151').setBackground('#FFFFFF')
    .setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBorder(false, false, true, false, null, null, '#FDB913',
               SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(2, 28);

  // Items de la guía (filas 3, 4, 5)
  const items = [
    ['🔒  PROTECCIÓN', 'El documento está protegido; solo las secciones en fblanco son editables.'],
    ['📍  DIRECCIONES', 'Escribir tal cual deben aparecer (mayúsculas, acentos, puntos, etc.).'],
    ['📋  PEGADO', 'Si copia datos, use "Pegar solo valores" para mantener el diseño.']
  ];
  items.forEach((item, idx) => {
    const filaItem = 3 + idx;
    sh.getRange('B' + filaItem)
      .setValue(item[0])
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setFontColor('#515151').setBackground('#FFFFFF')
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
    sh.getRange('C' + filaItem + ':E' + filaItem).merge()
      .setValue(item[1])
      .setFontFamily('Arial').setFontSize(9).setFontWeight('normal')
      .setFontColor('#7F8C8D').setBackground('#FFFFFF')
      .setVerticalAlignment('middle');
    sh.setRowHeight(filaItem, 22);
  });

  // Fila 6: separador vacío
  sh.setRowHeight(6, 12);
  // Fila 7: separador vacío
  sh.setRowHeight(7, 8);

  // ═══════════════════════════════════════════════════════════════════════════
  // FILAS 8-12 — SECCIÓN I — CONFIGURACIÓN GENERAL
  // ═══════════════════════════════════════════════════════════════════════════

  // Fila 8: Header sección I
  secHeader_(sh, 8, 'I', 'CONFIGURACIÓN GENERAL');

  // Fila 9: Régimen y Tipo de Firma
  lblVal_(sh, 9, 'Régimen', 'Persona Física', 'Tipo de Firma', 'Autógrafa');
  estiloInput_(sh, ['C9', 'E9']);
  valList_(sh, 'C9', ['Persona Física', 'Persona Moral']);
  valList_(sh, 'E9', ['Autógrafa', 'Digital']);

  // Fila 10: vacía (separador visual)
  sh.setRowHeight(10, 8);

  // Fila 11: No. de Avales y Código de Documento
  lblVal_(sh, 11, 'No. de Avales', 0, 'Código de Documento', null);
  estiloInput_(sh, 'C11');
  valList_(sh, 'C11', [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);

  sh.getRange('E11').setFormula(
    '=IF(OR($C$9="",$E$9="")," ",' +
    '"C_Pagaré_"&IF($C$9="Persona Física","PF","PM")&' +
    'IF($C$11>0,"A","")&IF($E$9="Digital","_D",""))'
  );
  estiloResultado_(sh, 'E11');

  // Fila 12: vacía (separador visual)
  sh.setRowHeight(12, 10);

  // ═══════════════════════════════════════════════════════════════════════════
  // FILAS 13-28 — SECCIÓN II — CONDICIONES FINANCIERAS
  // ═══════════════════════════════════════════════════════════════════════════

  // Fila 13: Header sección II
  secHeader_(sh, 13, 'II', 'CONDICIONES FINANCIERAS');

  // Fila 14: vacía (separador visual)
  sh.setRowHeight(14, 8);

  // Fila 15: ID Oportunidad y Divisa
  lblVal_(sh, 15, 'ID Oportunidad', '', 'Divisa', 'MXN');
  estiloInput_(sh, ['C15', 'E15']);
  valList_(sh, 'E15', ['MXN', 'USD']);

  // Fila 16: vacía
  sh.setRowHeight(16, 8);

  // Fila 17: Monto Principal y Tasa Ordinaria
  lblVal_(sh, 17, 'Monto Principal', 0, 'Tasa Ordinaria', 0);
  estiloInput_(sh, ['C17', 'E17']);
  sh.getRange('C17').setNumberFormat('"$ "#,##0.00');
  sh.getRange('E17').setNumberFormat('0.00%');

  // Fila 18: vacía
  sh.setRowHeight(18, 8);

  // Fila 19: Monto en Letra
  sh.getRange('B19').setValue('Monto en Letra')
    .setFontWeight('normal').setFontColor('#7F8C8D').setFontSize(11)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setFontFamily('Arial');
  sh.getRange('C19:E19').merge()
    .setFormula('=IF($C$17=0," ",IFERROR(NUM_TEXTO($C$17, $E$15)," "))');
  estiloResultado_(sh, 'C19:E19');

  // Fila 20: vacía
  sh.setRowHeight(20, 8);

  // Fila 21: Tasa en Letra
  sh.getRange('B21').setValue('Tasa en Letra')
    .setFontWeight('normal').setFontColor('#7F8C8D').setFontSize(11)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setFontFamily('Arial');
  sh.getRange('C21:E21').merge()
    .setFormula('=IF($E$17=0," ",TASA_TEXTO($E$17))');
  estiloResultado_(sh, 'C21:E21');

  // Fila 22: vacía
  sh.setRowHeight(22, 8);

  // Fila 23: Fecha de Firma y Fecha en Letra
  lblVal_(sh, 23, 'Fecha de Firma', '', 'Fecha en Letra', null);
  estiloInput_(sh, 'C23');
  sh.getRange('C23').setNumberFormat('dd/mm/yyyy');
  sh.getRange('E23').setFormula('=IF($C$23="","",FECHA_LETRA($C$23))');
  estiloResultado_(sh, 'E23');

  // Fila 24: vacía
  sh.setRowHeight(24, 8);

  // Fila 25: Plazo y Unidad
  lblVal_(sh, 25, 'Plazo', 0, 'Unidad', 'Meses');
  estiloInput_(sh, ['C25', 'E25']);
  sh.getRange('C25').setNumberFormat('0');
  valList_(sh, 'E25', ['Meses', 'Días']);
  sh.getRange('C25').setNote('Número entero. Usa la unidad adecuada en la celda de la derecha.');

  // Fila 26: vacía
  sh.setRowHeight(26, 8);

  // Fila 27: Fecha de Vencimiento y Vencimiento en Letra
  sh.getRange('B27').setValue('Fecha de Vencimiento')
    .setFontWeight('normal').setFontColor('#7F8C8D').setFontSize(11)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setFontFamily('Arial');
  sh.getRange('C27').setFormula(
    '=IF(OR($C$23="",$C$25=""),"",' +
    'IF($E$25="Meses",EDATE($C$23,$C$25),$C$23+$C$25))'
  );
  estiloResultado_(sh, 'C27');
  sh.getRange('C27').setNumberFormat('dd/mm/yyyy');

  sh.getRange('D27').setValue('Vencimiento en Letra')
    .setFontWeight('normal').setFontColor('#7F8C8D').setFontSize(11)
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontFamily('Arial');
  sh.getRange('E27').setFormula('=IF($C$27="","",FECHA_LETRA($C$27))');
  estiloResultado_(sh, 'E27');

  // Fila 28: vacía
  sh.setRowHeight(28, 10);

  // ═══════════════════════════════════════════════════════════════════════════
  // FILAS 29-37 — SECCIÓN III — SUSCRIPTOR
  // ═══════════════════════════════════════════════════════════════════════════

  // Fila 29: Header sección III
  secHeader_(sh, 29, 'III', 'IDENTIFICACIÓN DEL SUSCRIPTOR');

  // Fila 30: vacía
  sh.setRowHeight(30, 8);

  // Fila 31: Nombre / Razón Social
  sh.getRange('B31').setValue('Nombre / Razón Social')
    .setFontWeight('normal').setFontColor('#7F8C8D').setFontSize(11)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setFontFamily('Arial');
  sh.getRange('C31:E31').merge();
  estiloInput_(sh, 'C31:E31');

  // Fila 32: ↳ Tratamiento
  sh.getRange('B32').setValue('↳ Tratamiento')
    .setFontColor('#7F8C8D').setFontSize(9).setFontWeight('normal')
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontFamily('Arial');
  sh.getRange('C32:E32').merge()
    .setFormula('=IF($C$31="","",INFERIR_TRATO($C$31))');
  estiloResultado_(sh, 'C32:E32');

  // Fila 33: Domicilio
  sh.getRange('B33').setValue('Domicilio Completo')
    .setFontWeight('normal').setFontColor('#7F8C8D').setFontSize(11)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setFontFamily('Arial');
  sh.getRange('C33:E33').merge();
  estiloInput_(sh, 'C33:E33');

  // Fila 34: vacía
  sh.setRowHeight(34, 8);

// Fila 35: Representante 1 y 2
  lblVal_(sh, 35, 'Representante 1', '', 'Representante 2', '');
  estiloInput_(sh, ['C35', 'E35']);

  // Fila 36: Tratamientos inferidos de RL1 y RL2
  sh.getRange('B36').setValue('↳ Tratamiento')
    .setFontColor('#7F8C8D').setFontSize(9).setFontWeight('normal')
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontFamily('Arial');
    
  sh.getRange('D36').setValue('↳ Tratamiento')
    .setFontColor('#7F8C8D').setFontSize(9).setFontWeight('normal')
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontFamily('Arial');

  sh.getRange('C36').setFormula('=IF($C$35="","",INFERIR_TRATO($C$35))');
  sh.getRange('E36').setFormula('=IF($E$35="","",INFERIR_TRATO($E$35))');
  estiloResultado_(sh, ['C36', 'E36']);
  // Fila 37: vacía
  sh.setRowHeight(37, 10);

  // ═══════════════════════════════════════════════════════════════════════════
  // FILAS 38+ — SECCIÓN IV — AVALES (arranca en 39)
  // ═══════════════════════════════════════════════════════════════════════════

  // Fila 38: Header sección IV
  secHeader_(sh, 38, 'IV', 'RELACIÓN DE AVALES (hasta 10)');

  let currentRow = 39;
  let reglasFormato = [];

  // Regla: representantes se apagan si no es PM
  let reglaReps = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C$9<>"Persona Moral"')
    .setBackground('#F5F5F5').setFontColor('#BDC3C7')
    .setRanges([sh.getRange('B35:E36')]).build();
  reglasFormato.push(reglaReps);

  for (let i = 1; i <= 10; i++) {
    // Cabecera del bloque del aval
    sh.getRange(currentRow, 2, 1, 4).merge()
      .setValue('   AVAL ' + i)
      .setBackground('#F5F5F5')
      .setFontColor('#515151')
      .setFontWeight('bold')
      .setFontSize(9)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setFontFamily('Arial')
      .setBorder(false, false, true, false, null, null, '#FDB913',
                 SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(currentRow, 24);

    // Tratamiento + Nombre
    sh.getRange(currentRow + 1, 2).setValue('Tratamiento')
      .setFontWeight('normal').setFontColor('#7F8C8D').setFontSize(11)
      .setHorizontalAlignment('right')
      .setVerticalAlignment('middle').setFontFamily('Arial');
    sh.getRange('C' + (currentRow + 1))
      .setFormula('=IF(E' + (currentRow + 1) + '="","",INFERIR_TRATO(E' + (currentRow + 1) + '))');
    estiloResultado_(sh, 'C' + (currentRow + 1));

    sh.getRange(currentRow + 1, 4).setValue('Nombre Completo')
      .setFontWeight('normal').setFontColor('#7F8C8D').setFontSize(11)
      .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontFamily('Arial');
    estiloInput_(sh, 'E' + (currentRow + 1));
    sh.setRowHeight(currentRow + 1, 30);

    // Domicilio
    sh.getRange(currentRow + 2, 2).setValue('Domicilio Particular')
      .setFontWeight('normal').setFontColor('#7F8C8D').setFontSize(11)
      .setHorizontalAlignment('right')
      .setVerticalAlignment('middle').setFontFamily('Arial');
    sh.getRange('C' + (currentRow + 2) + ':E' + (currentRow + 2)).merge();
    estiloInput_(sh, 'C' + (currentRow + 2) + ':E' + (currentRow + 2));
    sh.setRowHeight(currentRow + 2, 30);

    // Regla: apagar bloque cuando C11 (No. Avales) < i
    let reglaAval = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C$11<' + i)
      .setBackground('#F5F5F5').setFontColor('#BDC3C7')
      .setRanges([sh.getRange(currentRow, 2, 3, 4)]).build();
    reglasFormato.push(reglaAval);

    currentRow += 4;
  }

  sh.setConditionalFormatRules(reglasFormato);

  // ═══════════════════════════════════════════════════════════════════════════
  // LEYENDA Y PIE DE PÁGINA
  // ═══════════════════════════════════════════════════════════════════════════
  sh.getRange(currentRow, 2, 1, 4).merge()
    .setValue('↳ Las direcciones repetidas se agruparán automáticamente en el pagaré')
    .setFontColor('#7F8C8D').setFontSize(9).setFontStyle('italic')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(currentRow, 22);
  currentRow += 2;

  sh.getRange(currentRow, 2, 1, 4).merge()
    .setValue('Sistema de Gestión de Pagarés Financiera Cualli V.1')
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontColor('#FFFFFF').setBackground('#515151')
    .setWrap(true);
  sh.setRowHeight(currentRow, 55);
  currentRow++;

  sh.getRange(currentRow, 2, 1, 4).merge().setBackground('#FDB913');
  sh.setRowHeight(currentRow, 3);
  currentRow++;

  sh.getRange(currentRow, 2, 1, 4).merge()
    .setValue('© 2026 Financiera Cualli. Todos los derechos reservados.')
    .setFontFamily('Arial').setFontSize(9).setFontWeight('normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontColor('#7F8C8D').setBackground('#FFFFFF');
  sh.setRowHeight(currentRow, 25);

  sh.setFrozenRows(1);
  aplicarProteccionesPagare_(sh);
  sincronizarListaOportunidadesC15();
}
/**
 *  panel de instruccionesssss
 */

function escribirSeccionInstrucciones(sh, fila) {
  const AM = CONFIG_PAGARES.COLOR_AMARILLO;
  const GR = CONFIG_PAGARES.COLOR_GRIS;
  const AT = CONFIG_PAGARES.COLOR_AMARILLO_TENUE;

  // Título de la sección
  sh.getRange(fila, 2, 1, 4).merge()
    .setValue('📋  GUÍA DE USO Y RECOMENDACIONES')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(AM).setBackground(GR)
    .setVerticalAlignment('middle');
  
  // Borde superior de la sección
  sh.getRange(fila, 2, 1, 4).setBorder(true, true, false, true, false, false, AM, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(fila, 28);
  fila++;

  const items = [
    ['🔒  PROTECCIÓN', 'El documento está protegido; solo las secciones en amarillo son editables.'],
    ['📍  DIRECCIONES', 'Escribir tal cual deben aparecer (mayúsculas, acentos, puntos, etc.).'],
    ['📋  PEGADO', 'Si copia datos, use "Pegar solo valores" para mantener el diseño.']
  ];

  items.forEach(item => {
    // Etiqueta (Col B)
    sh.getRange(fila, 2).setValue(item[0])
      .setFontFamily('Arial').setFontSize(9)
      .setFontWeight('bold') // SE CORRIGIÓ AQUÍ: de .setBold(true) a .setFontWeight('bold')
      .setFontColor(GR).setBackground(AT)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
    
    // Descripción (Cols C-E)
    sh.getRange(fila, 3, 1, 3).merge().setValue(item[1])
      .setFontFamily('Arial').setFontSize(9)
      .setFontColor(GR).setBackground(AT)
      .setVerticalAlignment('middle');

    sh.getRange(fila, 2, 1, 4).setBorder(false, true, false, true, false, false, AM, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(fila, 22);
    fila++;
  });

  sh.getRange(fila - 1, 2, 1, 4).setBorder(false, true, true, true, false, false, AM, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(fila, 15); 
  fila++;

  return fila;
}
/**
 *  Panel para Cargar Info SalesForce
 */



/**
 *  Header de sección con numeración romana separada.
 */
function secHeader_(sh, fila, romano, texto) {
  const rango = sh.getRange('B' + fila + ':E' + fila);
  rango.merge()
    .setValue('   ' + romano + '  ·  ' + texto)
    .setBackground('#515151')
    .setFontColor('#FDB913')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setFontFamily('Arial')
    .setBorder(false, false, true, false, null, null, '#FDB913',
               SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(fila, 28);
}


/**
 * Fila tipo Label-Value-Label-Value.
 */
function lblVal_(sh, fila, l1, v1, l2, v2) {
  sh.getRange('B' + fila)
    .setValue(l1)
    .setFontWeight('normal')
    .setFontColor('#7F8C8D')
    .setFontSize(11)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle')
    .setFontFamily('Arial');

  if (v1 !== null) sh.getRange('C' + fila).setValue(v1);

  // Etiqueta derecha (col D)
  sh.getRange('D' + fila)
    .setValue(l2)
    .setFontWeight('normal')
    .setFontColor('#7F8C8D')
    .setFontSize(11)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle')
    .setFontFamily('Arial');

  if (v2 !== null) sh.getRange('E' + fila).setValue(v2);

  sh.setRowHeight(fila, 28);
}


/**
 * Estilo de campo de entrada (lo que el usuario captura).
 */
function estiloInput_(sh, rangoStr) {
  const rangos = Array.isArray(rangoStr) ? rangoStr : [rangoStr];
  rangos.forEach(r => {
    sh.getRange(r)
      .setBackground('#FFFFFF')
      .setFontColor('#1B2631')
      .setFontWeight('bold')
      .setFontStyle('normal')
      .setFontFamily('Arial')
      .setFontSize(11)
      .setBorder(true, true, true, true, null, null, '#E5E7EB',
                 SpreadsheetApp.BorderStyle.SOLID);
  });
}

/**
 *  Estilo de campo de resultado (fórmulas).
 */
function estiloResultado_(sh, rangoStr) {
  const rangos = Array.isArray(rangoStr) ? rangoStr : [rangoStr];
  rangos.forEach(r => {
    sh.getRange(r)
      .setBackground('#d1e6f7')
      .setFontColor('#1F4E79')
      .setFontWeight('bold')
      .setFontStyle('italic')
      .setFontFamily('Arial')
      .setFontSize(11)
      .setBorder(true, true, true, true, null, null, '#D6E4F0',
                 SpreadsheetApp.BorderStyle.SOLID);
  });
}


function valList_(sh, celda, opciones) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones.map(String), true)
    .setAllowInvalid(false).build();
  sh.getRange(celda).setDataValidation(rule);
}


// =====================================================================================
// LECTURA Y VALIDACIÓN
// =====================================================================================
function leerDatosPagare_(sh) {
  const numAvales = Number(sh.getRange('C11').getValue()) || 0;
  const avalesArr = [];

  let currentRow = 39;
  for (let i = 1; i <= numAvales; i++) {
    avalesArr.push({
      trato:  String(sh.getRange('C' + (currentRow + 1)).getValue() || '').trim(),
      nombre: String(sh.getRange('E' + (currentRow + 1)).getValue() || '').trim(),
      domi:   String(sh.getRange('C' + (currentRow + 2)).getValue() || '').trim()
    });
    currentRow += 4;
  }

  return {
    regimen:        sh.getRange('C9').getValue(),
    tipoFirma:      sh.getRange('E9').getValue(),
    numAvales:      numAvales,
    codigo:         sh.getRange('E11').getValue(),
    idOp:           sh.getRange('C15').getValue(),
    moneda:         sh.getRange('E15').getValue(),
    montoNum:       Number(sh.getRange('C17').getValue()),
    tasaNum:        Number(sh.getRange('E17').getValue()),
    montoLetra:     sh.getRange('C19').getValue(),
    tasaLetra:      sh.getRange('C21').getValue(),
    fechaFirma:     sh.getRange('C23').getValue(),
    fechaLetra:     sh.getRange('E23').getValue(),
    plazoNum:       Number(sh.getRange('C25').getValue()) || 0,
    plazoUnidad:    String(sh.getRange('E25').getValue() || 'meses').trim().toLowerCase(),
    fechaVenc:      sh.getRange('C27').getValue(),
    fechaVencLetra: String(sh.getRange('E27').getValue() || '').trim(),
    nombreSusc:     sh.getRange('C31').getValue(),
    domiSusc:       sh.getRange('C33').getValue(),
    rl1:            sh.getRange('C35').getValue(),
    rl2:            sh.getRange('E35').getValue(),
    avales:         avalesArr
  };
}


function validarDatosPagare_(d) {
  const errs = [];
  if (!d.regimen)   errs.push('Falta Régimen (I).');
  if (!d.tipoFirma) errs.push('Falta Tipo de firma (I).');
  if (d.numAvales < 0 || d.numAvales > 10 || isNaN(d.numAvales))
                    errs.push('No. de avales debe ser 0–10 (I).');
  if (!d.idOp)      errs.push('Falta ID Oportunidad (II).');
  if (!d.moneda)    errs.push('Falta Moneda (II).');
  if (!d.montoNum || d.montoNum <= 0) errs.push('Monto inválido (II).');
  if (!d.tasaNum  || d.tasaNum  <= 0) errs.push('Tasa inválida (II).');
  if (!d.fechaFirma) errs.push('Falta Fecha de Firma (II).');
  if (!d.plazoNum || d.plazoNum <= 0) errs.push('Falta Plazo o es inválido (II).');
  if (!d.fechaVenc) errs.push('No se pudo calcular la Fecha de Vencimiento (II). Verifica Fecha de Firma y Plazo.');
  if (!d.nombreSusc) errs.push('Falta Nombre del Suscriptor (III).');
  if (!d.domiSusc)   errs.push('Falta Domicilio del Suscriptor (III).');

  if (d.regimen === 'Persona Moral' && !d.rl1) {
    errs.push('Falta Representante Legal 1 (III). Requerido en PM.');
  }
  for (let i = 0; i < d.numAvales; i++) {
    const a = d.avales[i];
    if (!a.nombre) errs.push('Aval ' + (i+1) + ': falta nombre (IV).');
    if (!a.domi)   errs.push('Aval ' + (i+1) + ': falta domicilio (IV).');
    if (!a.trato)  errs.push('Aval ' + (i+1) + ': falta tratamiento (IV).');
  }
  return errs;
}



// =====================================================================================
// ENDPOINTS PÚBLICOS (invocados desde el sidebar)
// =====================================================================================
function validarDatosDesdeSidebar() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (CONFIG_PAGARES.HOJAS_VALIDAS.indexOf(sh.getName()) === -1) {
    throw new Error('Debes estar posicionado en la hoja PAGARE_1 o PAGARE_2.');
  }
  const d = leerDatosPagare_(sh);
  const errs = validarDatosPagare_(d);
  if (errs.length === 0) {
    return { ok: true, codigo: d.codigo, numAvales: d.numAvales };
  }
  return { ok: false, errs: errs };
}

function verificarPagareExistenteDesdeSidebar() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (CONFIG_PAGARES.HOJAS_VALIDAS.indexOf(sh.getName()) === -1) {
    throw new Error('Debes estar en PAGARE_1 o PAGARE_2.');
  }
  const d = leerDatosPagare_(sh);
  if (!d.idOp) return null;
  return buscarPagareEnLog_(d.idOp);
}

function generarPagareDesdeSidebar(opciones) {
  const sh = SpreadsheetApp.getActiveSheet();
  opciones = opciones || {};
  const sobrescribir = opciones.sobrescribir === true;

  if (CONFIG_PAGARES.HOJAS_VALIDAS.indexOf(sh.getName()) === -1) {
    throw new Error('Debes estar posicionado en la hoja PAGARE_1 o PAGARE_2.');
  }
  if (!CONFIG_PAGARES.FOLDER_ID_PAGARES) {
    throw new Error('Falta configurar la carpeta destino.');
  }

  const d = leerDatosPagare_(sh);
  const errs = validarDatosPagare_(d);
  if (errs.length > 0) {
    throw new Error('Datos incompletos:\n\n' + errs.join('\n'));
  }

  if (sobrescribir) {
    const existente = buscarPagareEnLog_(d.idOp);
    if (existente && existente.docId) {
      try { DriveApp.getFileById(existente.docId).setTrashed(true); } catch (e) {}
      marcarEnLog_(existente.fila, 'SOBRESCRITO');
    }
  }

  const docUrl = renderizarPagare_(d);
  const mDocId = docUrl.match(/\/d\/([^\/]+)/);
  const docId = mDocId ? mDocId[1] : '';
  registrarEnLog_(d, docId, docUrl, sobrescribir ? 'REGENERADO' : 'NUEVO');

  toast_('✅ Pagaré generado correctamente', 3);
  openLink_(docUrl, 'Abriendo el pagaré generado…');
  return docUrl;
}

function previewPagareDesdeSidebar() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (CONFIG_PAGARES.HOJAS_VALIDAS.indexOf(sh.getName()) === -1) {
    throw new Error('Debes estar en PAGARE_1 o PAGARE_2.');
  }
  const d = leerDatosPagare_(sh);
  const errs = validarDatosPagare_(d);
  if (errs.length > 0) {
    throw new Error('Datos incompletos:\n\n' + errs.join('\n'));
  }
  mostrarModalPreviewPagare_(d);
  return true;
}


// =====================================================================================
// MOTOR DE RENDERIZADO A GOOGLE DOCS
// =====================================================================================
function renderizarPagare_(d) {
  const folder = DriveApp.getFolderById(CONFIG_PAGARES.FOLDER_ID_PAGARES);
  const tz = Session.getScriptTimeZone() || 'America/Mexico_City';
  const fecha = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const docName = d.idOp + '_' + d.codigo + '_' + fecha;

  const doc = DocumentApp.create(docName);
  const docId = doc.getId();

  const file = DriveApp.getFileById(docId);
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const body = doc.getBody();
  body.setPageHeight(792); // 11 pulgadas (Tamaño Carta)
  body.setPageWidth(612);  // 8.5 pulgadas (Tamaño Carta)
  body.setMarginTop(cmToPt(2.54));
  body.setMarginBottom(cmToPt(2.54));
  body.setMarginLeft(cmToPt(2.54));
  body.setMarginRight(cmToPt(2.54));
  body.clear();

  // [FIX v1.1] Header y footer se construyen SIN saveAndClose intermedio
  construirHeaderPagare_(doc, d);
  construirFooterPagare_(doc, d);

  renderizarCuerpoPagare_(body, d);

  // ÚNICO saveAndClose, al final de todo
  doc.saveAndClose();
  return 'https://docs.google.com/document/d/' + docId + '/edit';
}

function renderizarCuerpoPagare_(body, d) {
  const esPM      = d.regimen === 'Persona Moral';
  const esDigital = d.tipoFirma === 'Digital';
  const tieneAval = d.numAvales > 0;

  // === CÁLCULO DE FIRMANTES Y PÁGINAS ===
  const numFirmantesSusc = esPM ? (d.rl2 ? 2 : 1) : 1;
  const totalFirmantes = numFirmantesSusc + d.numAvales;
  const esUnaPagina = (!esDigital && totalFirmantes <= 2);
  const textoPaginas = esUnaPagina ? '1 (una)' : '2 (dos)';
  const palabraPagina = esUnaPagina ? 'página' : 'páginas';

  let textoRL = '';
  if (esPM) {
    const t1 = inferirTrato_(d.rl1);
    const a1 = inferirArticulo_(t1);
    textoRL = a1 + ' ' + t1.toUpperCase() + ' ' + formatearNombrePersona_(d.rl1);
    if (d.rl2) {
      const t2 = inferirTrato_(d.rl2);
      const a2 = inferirArticulo_(t2);
      textoRL += ' y ' + a2 + ' ' + t2 + ' ' + formatearNombrePersona_(d.rl2);
    }
  }

  const suscFormateado = esPM
    ? formatearNombreEmpresa_(d.nombreSusc).replace(/,$/, '')
    : formatearNombrePersona_(d.nombreSusc);
  const queSuscribe = esPM ? 'la que suscribe' :
    (inferirTrato_(d.nombreSusc) === 'Sra.' ? 'la que suscribe' : 'el que suscribe');
  const renunciaSomete = tieneAval
    ? 'el Suscriptor y el Aval y/o Avales se someten'
    : 'el Suscriptor se somete';
  const renuncianFuero = tieneAval
    ? 'el Suscriptor y el Aval y/o Avales renuncian'
    : 'el Suscriptor renuncia';
  const renuncianPresenta = tieneAval
    ? 'el Suscriptor y el Aval y/o Avales renuncian'
    : 'el Suscriptor renuncia';
  const prometenPagar = tieneAval
    ? 'El Suscriptor y el Aval y/o Avales prometen'
    : 'El Suscriptor promete';
  const tengaOLleguen = tieneAval ? 'tengan o lleguen' : 'tenga o llegue';
  const sufijoMoneda = d.moneda === 'USD'
    ? 'moneda de curso legal de los Estados Unidos de América'
    : 'moneda de curso legal de los Estados Unidos Mexicanos';
  const codigoMoneda = d.moneda === 'USD' ? 'USD' : 'M.N.';
  const montoFormateado = '$' + Number(d.montoNum).toLocaleString('en-US',
    {minimumFractionDigits: 2, maximumFractionDigits: 2});
  const tasaPct = (Number(d.tasaNum) * 100).toFixed(2).replace(/\.00$/, '');
  const fechaLetra = d.fechaLetra || fechaEnLetra_(d.fechaFirma);

  const tit = body.appendParagraph('PAGARÉ');
  tit.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
     .setSpacingBefore(6).setSpacingAfter(12);
  tit.editAsText().setBold(true).setFontSize(9).setFontFamily('Arial');
  const txtSuma = 'SUMA PRINCIPAL: ' + montoFormateado + ' ' + codigoMoneda;
  const sum = body.appendParagraph(txtSuma);
  sum.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(10);
  const tSum = sum.editAsText();
  tSum.setFontFamily('Arial').setFontSize(9).setBold(true);
  const fechaVencLetra = d.fechaVencLetra || fechaEnLetra_(d.fechaVenc);

  const partes1 = [
    '1. Por el presente Pagaré y por valor recibido, ' + queSuscribe + ' ',
    {bold: true, text: suscFormateado},
    ' (el "Suscriptor"), ',
    esPM ? 'representada por ' : '',
    esPM ? {bold: true, text: textoRL} : '',
    esPM ? ', ' : '',
    'promete incondicionalmente pagar a la orden de Financiera Cualli, Sociedad Anónima Promotora de Inversión de Capital Variable, Sociedad Financiera de Objeto Múltiple, Entidad No Regulada, en el domicilio de dicha sociedad ubicado en ' + CONFIG_PAGARES.DOMICILIO_CUALLI + ', el día ',
    {bold: true, text: fechaVencLetra},
    ', la suma principal de ',
    {bold: true, text: montoFormateado + ' (' + (d.montoLetra || '') + ', ' + sufijoMoneda + ')'},
    '.'
  ];
  agregarParrafoConFormato_(body, partes1);

  const partes2 = [
    '2. La suma principal amparada por este Pagaré devengará intereses ordinarios a partir de esta fecha y hasta el pago total del mismo a una tasa de interés anual de ',
    {bold: true, text: tasaPct + '% (' + (d.tasaLetra || '') + ')'},
    '. Ante el incumplimiento en el pago de la suma principal e intereses derivados de este Pagaré en lugar de la tasa de interés ordinaria referida anteriormente este Pagaré devengará intereses moratorios sobre su saldo insoluto a una tasa de interés igual a la tasa de interés anual ordinaria mencionada al principio de este párrafo multiplicada por 2 (dos).'
  ];
  agregarParrafoConFormato_(body, partes2);

  parrJust_(body,
    '3. Los intereses se computarán sobre la base de días naturales efectivamente transcurridos ' +
    '(desde el día siguiente en que el Suscriptor reciba el Préstamo, hasta e incluyendo, la fecha ' +
    'en que proceda el pago), considerando un año de trescientos sesenta (360) días.');
  parrJust_(body,
    '4. El Suscriptor tendrá el derecho de pagar anticipadamente todo o parte del principal y/o ' +
    'los intereses generados sobre el principal de este Pagaré, debiéndose hacer la anotación ' +
    'respectiva en este Pagaré de las cantidades efectivamente pagadas por el Suscriptor.');
  parrJust_(body,
    '5. Los pagos de las sumas debidas conforme a este Pagaré, deberán ser hechos por el Suscriptor ' +
    'libres y sin ninguna deducción por concepto de impuestos ya sean presentes o futuros, tributos, ' +
    'contribuciones, deducciones, cargos, retenciones, recargos, multas, sanciones y cualesquier ' +
    'otras cargas fiscales de cualquier clase.');
  parrJust_(body,
    '6. Este Pagaré estará sujeto y se interpretará conforme a las leyes de los Estados Unidos Mexicanos.');
  
  let numCl = 7;
  if (esDigital) {
    parrJust_(body,
      numCl + '. La Firma Electrónica tiene el mismo valor probatorio que la firma autógrafa tiene, ' +
      'por lo que el uso y aplicación de la misma en la Plataforma se sujetará a la legislación aplicable.');
    numCl++;
  }

  parrJust_(body,
    numCl + '. Para la ejecución y cumplimiento de este Pagaré y para el requerimiento judicial de ' +
    'pago de las cantidades adeudadas conforme al mismo, ' + renunciaSomete + ' expresa e ' +
    'irrevocablemente a la jurisdicción de los tribunales competentes en la Ciudad de México. ' +
    'Mediante la suscripción y entrega de este Pagaré ' + renuncianFuero + ' irrevocablemente a ' +
    'cualquier otro fuero al que ' + tengaOLleguen + ' a tener derecho, en virtud de su domicilio ' +
    '(presente o futuro) o por cualquier otra razón.');
  numCl++;
  
  renderizarBloqueDomicilios_(body, d, numCl, tieneAval);
  numCl++;

  parrJust_(body,
    numCl + '. Por el presente Pagaré ' + renuncianPresenta + ' a cualquier diligencia de ' +
    'presentación, requerimiento o protesto. La omisión o retraso del tenedor del presente Pagaré ' +
    'en el ejercicio de cualquiera de sus derechos conforme a este Pagaré en ningún caso ' +
    'constituirá una renuncia a dichos derechos.');
  numCl++;

  parrJust_(body,
    numCl + '. ' + prometenPagar + ' incondicional e irrevocablemente pagar los costos y gastos ' +
    'que impliquen el cobro de este Pagaré incluyendo, sin limitación alguna, los honorarios de los ' +
    'abogados que intervengan en el cobro, en caso de incumplimiento en el pago de este Pagaré.');
  numCl++;

  if (esDigital) {
    parrJust_(body,
      numCl + '. El presente Pagaré se firma con Firma Electrónica Avanzada, la Firma Electrónica ' +
      'Avanzada tiene el mismo valor probatorio que la firma autógrafa, por lo que el uso y ' +
      'aplicación de la misma en la Plataforma se sujetará a la legislación aplicable, se considera ' +
      'el lugar y fecha de firma la relacionada en la constancia emitida por el prestador de ' +
      'servicios de certificación.');
  } else {
    // TEXTO DINÁMICO DE PÁGINAS INYECTADO AQUÍ
    agregarParrafoConFormato_(body, [
      numCl + '. El presente Pagaré consiste de ',
      {bold: true, text: textoPaginas},
      ' ' + palabraPagina + ' y se suscribe y entrega en la Ciudad de México, el día ',
      {bold: true, text: fechaLetra},
      '.'
    ]);
  }

  body.appendParagraph('').setSpacingAfter(esUnaPagina ? 5 : 20);
  
  construirTablaFirmas_(body, d, esPM, tieneAval, esUnaPagina);
}

// =====================================================================================
// HELPERS DE RENDERIZADO (todos SIN saveAndClose)
// =====================================================================================
function parrJust_(body, texto) {
  const p = body.appendParagraph(texto);
  p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
   .setSpacingAfter(6).setSpacingBefore(2).setLineSpacing(1.15);
  p.editAsText().setFontFamily('Arial').setFontSize(8).setBold(false);
  return p;
}

function agregarParrafoConFormato_(body, partes) {
  const p = body.appendParagraph('');
  p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
   .setSpacingAfter(6).setSpacingBefore(2).setLineSpacing(1.15);
  
  let fullText = "";
  let boldRanges = [];
  
  // 1. Unimos todo el texto y guardamos las posiciones exactas
  partes.forEach(parte => {
    if (!parte) return;
    if (typeof parte === 'string') {
      if (parte.length > 0) fullText += parte;
    } else {
      const txt = parte.text;
      if (txt && txt.length > 0) {
        const start = fullText.length;
        fullText += txt;
        if (parte.bold) boldRanges.push({start: start, end: fullText.length - 1});
      }
    }
  });
  
  // 2. Insertamos TODO el texto plano con tamaño 8 y quitamos negritas a todo
  const t = p.editAsText();
  t.setText(fullText);
  t.setFontFamily('Arial').setFontSize(8).setBold(false);
  
  // 3. Aplicamos las negritas SOLO a los rangos que guardamos matemáticamente
  boldRanges.forEach(r => {
    t.setBold(r.start, r.end, true);
  });
}

function agregarParrafoDomicilio_(body, etiqueta, domicilio) {
  const p = body.appendParagraph('');
  p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
   .setSpacingAfter(4).setSpacingBefore(2).setLineSpacing(1.15);
  const t = p.editAsText();
  t.appendText(etiqueta);
  t.appendText(domicilio + '.');
  t.setFontFamily('Arial').setFontSize(8);
  t.setBold(0, etiqueta.length - 1, true);   // ETIQUETA EN NEGRITAS
  t.setBold(etiqueta.length, etiqueta.length + domicilio.length, false);
}

/**
 * [FIX v1.1] Construye header sin cerrar el documento.
 * El flujo principal en renderizarPagare_ hace el único saveAndClose al final.
 */
function construirHeaderPagare_(doc, d) {
  let header = doc.getHeader();
  if (!header) header = doc.addHeader();
  header.clear();

  const p = header.appendParagraph(d.codigo + '_V1.0');
  p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.0);
  p.editAsText().setFontFamily('Courier New').setFontSize(8).setForegroundColor('#666666');
}

function construirFooterPagare_(doc, d) {
  let footer = doc.getFooter();
  if (footer) footer.clear();
}
function construirTablaFirmas_(body, d, esPM, tieneAval, esUnaPagina = false) {
  const lineaFirma = '_______________________________';
  const columnas = [];
  if (esPM) {
    columnas.push(construirColFirmaRL_(d.nombreSusc, d.rl1));
    if (d.rl2) columnas.push(construirColFirmaRL_(d.nombreSusc, d.rl2));
  } else {
    columnas.push(construirColFirmaPF_(d.nombreSusc));
  }

  for (let i = 0; i < d.numAvales; i++) {
    const a = d.avales[i];
    columnas.push({
      header: '"EL AVAL"',
      linea: lineaFirma,
      empresa: null,
      nombre: a.trato.toUpperCase() + ' ' + formatearNombrePersona_(a.nombre),
      rol: 'POR PROPIO DERECHO'
    });
  }

  const esDigital = d.tipoFirma === 'Digital';

  if (esDigital) {
    construirListaFirmasDigital_(body, columnas, esPM);
  } else {
    const filas = distribuirFirmas_(columnas, esPM, d.rl2 ? 2 : 1);
    filas.forEach((fila, idxFila) => {
      construirFilaFirmas_(body, fila, esPM, esUnaPagina);
      if (idxFila < filas.length - 1) {
        const sep = body.appendParagraph('\u00A0');
        // Reducir la separación entre filas si queremos comprimir
        sep.setSpacingBefore(esUnaPagina ? 5 : 14).setSpacingAfter(0);
        sep.editAsText().setFontSize(1);
      }
    });
  }
}



/**
 * Construye UNA fila de firmas como tabla independiente.
 * Si fila tiene 1 firmante: tabla de 3 celdas [padding | firma | padding] para centrar.
 * Si fila tiene 2 firmantes: tabla de 2 celdas, cada una con su firma.
 */
function construirFilaFirmas_(body, fila, esPM, esUnaPagina = false) {
  const tbl = body.appendTable();
  tbl.setBorderWidth(0);

  const esSola = fila.length === 1;
  const row = tbl.appendTableRow();

  if (esSola) {
    row.appendTableCell('').setWidth(cmToPt(4.5));
    const celda = row.appendTableCell();
    llenarCeldaFirma_(celda, fila[0], esPM, esUnaPagina);
    row.appendTableCell('').setWidth(cmToPt(4.5));
  } else {
    fila.forEach(col => {
      const celda = row.appendTableCell();
      llenarCeldaFirma_(celda, col, esPM, esUnaPagina);
    });
  }
}

/**
 * Llena una celda con el bloque completo de una firma:
 *   header ("EL SUSCRIPTOR" / "EL AVAL N")
 *   [línea en blanco]
 *   _______________________________
 *   [empresa si aplica]
 *   nombre del firmante
 *   rol ("POR PROPIO DERECHO" / "REPRESENTANTE LEGAL")
 */
function llenarCeldaFirma_(celda, col, esPM, esUnaPagina = false) {
  let pHeader;
  if (celda.getNumChildren() > 0 &&
      celda.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
    pHeader = celda.getChild(0).asParagraph();
    pHeader.setText(col.header);
  } else {
    pHeader = celda.appendParagraph(col.header);
  }
  estiloParrafoFirma_(pHeader, 9, true);
  
  // 2. Aire dinámico (Comprimido a 30 si es una sola página, normal de 60 si son dos)
  const pAire = celda.appendParagraph('\u00A0');
  estiloParrafoFirma_(pAire, 9, false); 
  pAire.setSpacingBefore(esUnaPagina ? 30 : 60); 

  // 3. Línea de firma
  const pLinea = celda.appendParagraph(col.linea);
  estiloParrafoFirma_(pLinea, 10, false);
  
  if (esPM && col.empresa) {
    const pEmpresa = celda.appendParagraph(col.empresa);
    estiloParrafoFirma_(pEmpresa, 9, true);
    const t = pEmpresa.editAsText();
    const m = col.empresa.match(/^"([^"]+)"/);
    if (m) {
      const nom = m[1];
      const idx = col.empresa.indexOf(nom);
      if (idx !== -1) t.setBold(idx, idx + nom.length - 1, true);
    }
  }

  const pNombre = celda.appendParagraph(col.nombre);
  estiloParrafoFirma_(pNombre, 9, true);

  const pRol = celda.appendParagraph(col.rol);
  estiloParrafoFirma_(pRol, 9, true);
}


function estiloParrafoFirma_(p, fontSize, bold) {
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
   .setLineSpacing(1.15).setSpacingBefore(0).setSpacingAfter(0);
  const t = p.editAsText();
  t.setFontFamily('Arial').setFontSize(fontSize).setBold(bold);
}




function construirColFirmaRL_(nombreEmpresa, nombreRL) {
  const empresaFmt = formatearNombreEmpresa_(nombreEmpresa).replace(/,$/, '');
  const t = inferirTrato_(nombreRL).toUpperCase();
  return {
    header: '"EL SUSCRIPTOR"',
    linea: '_______________________________',
    empresa: empresaFmt,
    nombre: t + ' ' + formatearNombrePersona_(nombreRL),
    rol: 'REPRESENTANTE LEGAL'
  };
}

function construirColFirmaPF_(nombreSusc) {
  const trato = inferirTrato_(nombreSusc).toUpperCase();
  return {
    header: '"EL SUSCRIPTOR"',
    linea: '_______________________________',
    empresa: null,
    nombre: trato + ' ' + formatearNombrePersona_(nombreSusc),
    rol: 'POR PROPIO DERECHO'
  };
}

function estiloFirma_(celda, centrar, fontSize, bold) {
  celda.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(8).setPaddingRight(8);
  const n = celda.getNumChildren();
  for (let i = 0; i < n; i++) {
    const el = celda.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      p.setAlignment(centrar ? DocumentApp.HorizontalAlignment.CENTER
                             : DocumentApp.HorizontalAlignment.LEFT)
       .setLineSpacing(1.15).setSpacingBefore(0).setSpacingAfter(0);
      const t = p.editAsText();
      t.setFontFamily('Arial').setFontSize(8).setBold(bold); // Tamaño forzado a 8
    }
  }
}



function numToLetra_(n) {
  const pal = ['cero','una','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez'];
  return pal[n] || String(n);
}

function fechaEnLetra_(fecha) {
  if (!fecha) return '';
  const d = fecha instanceof Date ? fecha : new Date(fecha);
  const meses = ['enero','febrero','marzo','abril','mayo','junio',
                 'julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return d.getDate() + ' de ' + meses[d.getMonth()] + ' de ' + d.getFullYear();
}


// =====================================================================================
// CUSTOM FUNCTIONS PARA CELDAS
// =====================================================================================
/** @customfunction */
function FECHA_LETRA(fecha) {
  if (!fecha) return '';
  return fechaEnLetra_(fecha);
}

/** @customfunction */
function TASA_TEXTO(decimal) {
  if (decimal === '' || decimal === null || decimal === undefined) return '';
  
  const pct = Number(decimal) * 100;
  const entero = Math.floor(pct);
  // Extraemos los decimales del porcentaje (ej. 29.50% -> 50)
  const centesimas = Math.round((pct - entero) * 100); 
  
  let txt = numeroAPalabras_(entero);
  
  if (centesimas > 0) {
    let txtDecimal = numeroAPalabras_(centesimas);
    // Si el decimal es menor a 10 (ej. 29.05%), agregamos un "CERO" para que diga "PUNTO CERO CINCO"
    if (centesimas < 10) {
      txtDecimal = "cero " + txtDecimal;
    }
    txt += " punto " + txtDecimal;
  }
  
  txt += " por ciento";
  
    return txt.charAt(0).toUpperCase() + txt.slice(1).toLowerCase();
}
/** @customfunction */
function INFERIR_TRATO(nombre) {
  if (!nombre) return '';
  return inferirTrato_(nombre);
}

function numeroAPalabras_(n) {
  if (n === 0) return 'cero';
  const u = ['','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve',
             'diez','once','doce','trece','catorce','quince','dieciséis','diecisiete',
             'dieciocho','diecinueve','veinte','veintiuno','veintidós','veintitrés',
             'veinticuatro','veinticinco','veintiséis','veintisiete','veintiocho','veintinueve'];
  const d = ['','','','treinta','cuarenta','cincuenta','sesenta','setenta','ochenta','noventa'];
  const c = ['','ciento','doscientos','trescientos','cuatrocientos','quinientos',
             'seiscientos','setecientos','ochocientos','novecientos'];
  if (n < 30) return u[n];
  if (n < 100) { const dd = Math.floor(n/10), uu = n%10; return d[dd] + (uu ? ' y ' + u[uu] : ''); }
  if (n === 100) return 'cien';
  const cc = Math.floor(n/100), r = n%100;
  return c[cc] + (r ? ' ' + numeroAPalabras_(r) : '');
}


// =====================================================================================
// LOG DE PAGARÉS GENERADOS (hoja oculta _LOG_PAGARES)
// =====================================================================================
function obtenerOcreaHojaLog_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG_PAGARES.HOJA_LOG);
  if (sh) return sh;

  sh = ss.insertSheet(CONFIG_PAGARES.HOJA_LOG);
  sh.getRange('A1:H1').setValues([[
    'Timestamp', 'Usuario', 'ID Oportunidad', 'Plantilla',
    'Suscriptor', 'Doc ID', 'URL', 'Estatus'
  ]]).setFontWeight('bold').setBackground('#515151').setFontColor('#FDB913');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 140);
  sh.setColumnWidth(5, 280);
  sh.setColumnWidth(6, 240);
  sh.setColumnWidth(7, 300);
  sh.setColumnWidth(8, 100);
  sh.hideSheet();
  return sh;
}

function registrarEnLog_(d, docId, url, estatus) {
  const sh = obtenerOcreaHojaLog_();
  const tz = Session.getScriptTimeZone() || 'America/Mexico_City';
  const ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  let usuario = '(desconocido)';
  try { usuario = Session.getActiveUser().getEmail() || '(desconocido)'; } catch(e) {}
  sh.appendRow([ts, usuario, d.idOp, d.codigo, d.nombreSusc, docId, url, estatus || 'NUEVO']);
}

function marcarEnLog_(fila, estatus) {
  const sh = obtenerOcreaHojaLog_();
  sh.getRange(fila, 8).setValue(estatus);
}

function buscarPagareEnLog_(idOp) {
  const sh = obtenerOcreaHojaLog_();
  if (sh.getLastRow() < 2) return null;

  const datos = sh.getRange(1, 1, sh.getLastRow(), 8).getValues();
  for (let i = datos.length - 1; i >= 1; i--) {
    const fila = datos[i];
    const rowIdOp = String(fila[2] || '').trim();
    const rowEstatus = String(fila[7] || '').trim();
    if (rowIdOp === idOp.trim() && rowEstatus !== 'SOBRESCRITO') {
      return {
        fila: i + 1,
        idOp: rowIdOp,
        timestamp: String(fila[0] || ''),
        plantilla: String(fila[3] || ''),
        suscriptor: String(fila[4] || ''),
        docId: String(fila[5] || ''),
        url: String(fila[6] || ''),
        estatus: rowEstatus
      };
    }
  }
  return null;
}

function abrirLogPagares() {
  const sh = obtenerOcreaHojaLog_();
  sh.showSheet();
  SpreadsheetApp.getActive().setActiveSheet(sh);
  toast_('📋 Log de pagarés abierto', 2);
}


// =====================================================================================
// MODAL DE PREVIEW (vista previa antes de generar)
// =====================================================================================
function mostrarModalPreviewPagare_(d) {
  const htmlContenido = construirHTMLPreview_(d);
  const wrapper = `
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Vista previa del Pagaré</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      margin: 0; padding: 20px;
      background: #f5f5f5;
      display: flex; flex-direction: column; align-items: center;
    }
    .toolbar {
      position: sticky; top: 0; z-index: 10;
      background: #fff; padding: 12px 16px;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,.1);
      margin-bottom: 20px;
      display: flex; gap: 10px; align-items: center;
      width: 21cm;
    }
    .btn {
      padding: 8px 14px;
      border: none; border-radius: 6px;
      background: #2563eb; color: white;
      cursor: pointer; font-weight: 600;
      font-family: Arial, sans-serif;
    }
    .btn:hover { background: #1d4ed8; }
    .btn-alt { background: #6b7280; }
    .btn-alt:hover { background: #4b5563; }
    .hint { color: #666; font-size: 12px; margin-left: auto; }

    .doc-preview {
      background: white;
      width: 21cm;
      min-height: 29.7cm;
      padding: 2cm 2.5cm;
      box-shadow: 0 4px 20px rgba(0,0,0,.15);
      font-family: Arial, sans-serif;
      font-size: 9pt;
      line-height: 1.3;
      position: relative;
    }
    .doc-header-mini {
      text-align: right;
      font-family: 'Courier New', monospace;
      font-size: 8pt;
      color: #666;
      border-bottom: 1px solid #ddd;
      padding-bottom: 8px;
      margin-bottom: 20px;
    }
    .doc-titulo {
      text-align: center; font-size: 11pt; font-weight: bold;
      margin: 20px 0 15px;
    }
    .doc-suma { margin-bottom: 15px; font-weight: bold; font-size: 10pt; text-align: center; }
    .doc-clausulas p {
      text-align: justify; margin: 8px 0;
    }

    /* === FIRMAS con distribución 2×N === */
    .firmas-wrap {
      display: flex; flex-direction: column;
      margin-top: 50px; gap: 40px;
    }
    .firmas-fila {
      display: flex; justify-content: center;
      gap: 60px;
    }
    .firmas-fila-par .firma-col {
      flex: 1 1 45%; max-width: 45%;
    }
    .firmas-fila-sola .firma-col {
      flex: 0 0 45%; max-width: 45%;
    }
    .firma-col {
      text-align: center; font-size: 9pt;
      font-weight: bold; /* Sincroniza con el estilo del documento */
    }
    .firma-header { font-weight: bold; margin-bottom: 4px; }
    .firma-linea-aire { height: 80px; }
    .firma-linea { font-weight: normal; font-size: 10pt; margin-bottom: 4px; }
    .firma-empresa { margin: 2px 0; }
    .firma-empresa { font-weight: bold; }
    .firma-nombre { margin: 2px 0; font-weight: bold; }
    .firma-rol { margin: 2px 0; font-weight: bold; }

    .firmas-digital {
      margin-top: 50px;
      display: flex; flex-direction: column;
      gap: 18px;
      align-items: flex-start;
    }
    .firma-digital-bloque {
      text-align: left;
      font-size: 9pt;
      line-height: 1.4;
    }
    .firma-digital-header,
    .firma-digital-empresa,
    .firma-digital-nombre,
    .firma-digital-rol {
      font-weight: bold;
      margin: 0;
    }


    .watermark {
      position: absolute;
      top: 40%; left: 50%;
      transform: translate(-50%, -50%) rotate(-30deg);
      font-size: 80pt; color: rgba(251, 184, 24, 0.15);
      font-weight: bold; z-index: 0;
      pointer-events: none;
      font-family: Arial, sans-serif;
      white-space: nowrap;
    }
    .doc-preview > *:not(.watermark) { position: relative; z-index: 1; }

    @media print {
      body { background: white; padding: 0; }
      .toolbar { display: none; }
      .doc-preview { box-shadow: none; width: auto; }
    }
  </style>
</head>
<body>
  <div class="toolbar">
    <button class="btn" onclick="window.print()">🖨 Imprimir / PDF</button>
    <button class="btn btn-alt" onclick="google.script.host.close()">Cerrar</button>
    <span class="hint">Vista previa — revisa antes de generar el Doc final</span>
  </div>

  <div class="doc-preview">
    <div class="watermark">BORRADOR</div>
    ${htmlContenido}
  </div>
</body>
</html>`;
  const html = HtmlService.createHtmlOutput(wrapper).setWidth(960).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vista previa del Pagaré');
}
function construirHTMLPreview_(d) {
  const esc = (s) => !s ? '' : s.toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');

  const esPM      = d.regimen === 'Persona Moral';
  const esDigital = d.tipoFirma === 'Digital';
  const tieneAval = d.numAvales > 0;

  // === CÁLCULO DE FIRMANTES Y PÁGINAS ===
  const numFirmantesSusc = esPM ? (d.rl2 ? 2 : 1) : 1;
  const totalFirmantes = numFirmantesSusc + d.numAvales;
  const esUnaPagina = (!esDigital && totalFirmantes <= 2);
  const textoPaginas = esUnaPagina ? '1 (una)' : '2 (dos)';
  const palabraPagina = esUnaPagina ? 'página' : 'páginas';

  let textoRL = '';
  if (esPM) {
    const t1 = inferirTrato_(d.rl1);
    const a1 = inferirArticulo_(t1);
    textoRL = a1 + ' <strong>' + t1.toUpperCase() + ' ' + esc(formatearNombrePersona_(d.rl1)) + '</strong>';
    if (d.rl2) {
      const t2 = inferirTrato_(d.rl2);
      const a2 = inferirArticulo_(t2);
      textoRL += ' y ' + a2 + ' <strong>' + t2 + ' ' + esc(formatearNombrePersona_(d.rl2)) + '</strong>';
    }
  }

  const suscFormateado = esPM
    ? formatearNombreEmpresa_(d.nombreSusc).replace(/,$/, '')
    : formatearNombrePersona_(d.nombreSusc);
  const queSuscribe = esPM ? 'la que suscribe' :
    (inferirTrato_(d.nombreSusc) === 'Sra.' ? 'la que suscribe' : 'el que suscribe');
  const renunciaSomete = tieneAval ? 'el Suscriptor y el Aval y/o Avales se someten' : 'el Suscriptor se somete';
  const renuncianFuero = tieneAval ? 'el Suscriptor y el Aval y/o Avales renuncian' : 'el Suscriptor renuncia';
  const renuncianPresenta = tieneAval ? 'el Suscriptor y el Aval y/o Avales renuncian' : 'el Suscriptor renuncia';
  const prometenPagar = tieneAval ? 'El Suscriptor y el Aval y/o Avales prometen' : 'El Suscriptor promete';
  const tengaOLleguen = tieneAval ? 'tengan o lleguen' : 'tenga o llegue';
  const sufijoMoneda = d.moneda === 'USD' ? 'moneda de curso legal de los Estados Unidos de América' : 'moneda de curso legal de los Estados Unidos Mexicanos';
  const codigoMoneda = d.moneda === 'USD' ? 'USD' : 'M.N.';
  const montoFormateado = '$' + Number(d.montoNum).toLocaleString('en-US',
    {minimumFractionDigits: 2, maximumFractionDigits: 2});
  const tasaPct = (Number(d.tasaNum) * 100).toFixed(2).replace(/\.00$/, '');
  const fechaLetra = d.fechaLetra || fechaEnLetra_(d.fechaFirma);
  const domiCualli = CONFIG_PAGARES.DOMICILIO_CUALLI;

  let numCl = 7;
  const fechaVencLetra = d.fechaVencLetra || fechaEnLetra_(d.fechaVenc);
  
  const cl1 = `1. Por el presente Pagaré y por valor recibido, ${queSuscribe} <strong>${esc(suscFormateado)}</strong> (el "Suscriptor"), ${esPM ? 'representada por ' + textoRL + ', ' : ''}promete incondicionalmente pagar a la orden de <strong>Financiera Cualli, Sociedad Anónima Promotora de Inversión de Capital Variable, Sociedad Financiera de Objeto Múltiple, Entidad No Regulada</strong>, en el domicilio de dicha sociedad ubicado en ${domiCualli}, <strong>el día ${fechaVencLetra}</strong>, la suma principal de <strong>${montoFormateado} (${esc(d.montoLetra || '')} ${sufijoMoneda})</strong>.`;
  const cl2 = `2. La suma principal amparada por este Pagaré devengará intereses ordinarios a partir de esta fecha y hasta el pago total del mismo a una tasa de interés anual de <strong>${tasaPct}% (${esc(d.tasaLetra || '')})</strong>. Ante el incumplimiento en el pago de la suma principal e intereses derivados de este Pagaré en lugar de la tasa de interés ordinaria referida anteriormente este Pagaré devengará intereses moratorios sobre su saldo insoluto a una tasa de interés igual a la tasa de interés anual ordinaria mencionada al principio de este párrafo multiplicada por 2 (dos).`;
  
  let clausulasHTML = `<p>${cl1}</p><p>${cl2}</p>`;
  clausulasHTML += `<p>3. Los intereses se computarán sobre la base de días naturales efectivamente transcurridos (desde el día siguiente en que el Suscriptor reciba el Préstamo, hasta e incluyendo, la fecha en que proceda el pago), considerando un año de trescientos sesenta (360) días.</p>`;
  clausulasHTML += `<p>4. El Suscriptor tendrá el derecho de pagar anticipadamente todo o parte del principal y/o los intereses generados sobre el principal de este Pagaré, debiéndose hacer la anotación respectiva en este Pagaré de las cantidades efectivamente pagadas por el Suscriptor.</p>`;
  clausulasHTML += `<p>5. Los pagos de las sumas debidas conforme a este Pagaré, deberán ser hechos por el Suscriptor libres y sin ninguna deducción por concepto de impuestos ya sean presentes o futuros, tributos, contribuciones, deducciones, cargos, retenciones, recargos, multas, sanciones y cualesquier otras cargas fiscales de cualquier clase.</p>`;
  clausulasHTML += `<p>6. Este Pagaré estará sujeto y se interpretará conforme a las leyes de los Estados Unidos Mexicanos.</p>`;
  
  if (esDigital) {
    clausulasHTML += `<p>${numCl}. La Firma Electrónica tiene el mismo valor probatorio que la firma autógrafa tiene, por lo que el uso y aplicación de la misma en la Plataforma se sujetará a la legislación aplicable.</p>`;
    numCl++;
  }

  clausulasHTML += `<p>${numCl}. Para la ejecución y cumplimiento de este Pagaré y para el requerimiento judicial de pago de las cantidades adeudadas conforme al mismo, ${renunciaSomete} expresa e irrevocablemente a la jurisdicción de los tribunales competentes en la Ciudad de México. Mediante la suscripción y entrega de este Pagaré ${renuncianFuero} irrevocablemente a cualquier otro fuero al que ${tengaOLleguen} a tener derecho, en virtud de su domicilio (presente o futuro) o por cualquier otra razón.</p>`;
  numCl++;

  // BLOQUE DE DOMICILIOS AGRUPADOS
  const grupos = agruparPorDomicilio_(d);
  if (d.numAvales === 0 && grupos.length === 1) {
    const domiFmt = formatearDomicilio_(grupos[0].domicilio);
    clausulasHTML += `<p>${numCl}. El Suscriptor designa como domicilio para requerimiento judicial de pago, el siguiente: <strong>${esc(domiFmt)}</strong>.</p>`;
  } else {
    const domIntro = `${numCl}. El Suscriptor y el Aval designan como sus domicilios para requerimiento judicial de pago, los siguientes:`;
    clausulasHTML += `<p>${domIntro}</p>`;
    grupos.forEach(g => {
      const domiFmt = formatearDomicilio_(g.domicilio);
      clausulasHTML += `<p><strong>${esc(g.etiqueta)}:</strong> ${esc(domiFmt)}.</p>`;
    });
  }
  numCl++;

  clausulasHTML += `<p>${numCl}. Por el presente Pagaré ${renuncianPresenta} a cualquier diligencia de presentación, requerimiento o protesto. La omisión o retraso del tenedor del presente Pagaré en el ejercicio de cualquiera de sus derechos conforme a este Pagaré en ningún caso constituirá una renuncia a dichos derechos.</p>`;
  numCl++;

  clausulasHTML += `<p>${numCl}. ${prometenPagar} incondicional e irrevocablemente pagar los costos y gastos que impliquen el cobro de este Pagaré incluyendo, sin limitación alguna, los honorarios de los abogados que intervengan en el cobro, en caso de incumplimiento en el pago de este Pagaré.</p>`;
  numCl++;

  if (esDigital) {
    clausulasHTML += `<p>${numCl}. El presente Pagaré se firma con Firma Electrónica Avanzada, la Firma Electrónica Avanzada tiene el mismo valor probatorio que la firma autógrafa, por lo que el uso y aplicación de la misma en la Plataforma se sujetará a la legislación aplicable, se considera el lugar y fecha de firma la relacionada en la constancia emitida por el prestador de servicios de certificación.</p>`;
  } else {
    clausulasHTML += `<p>${numCl}. El presente Pagaré consiste de <strong>${textoPaginas}</strong> ${palabraPagina} y se suscribe y entrega en la Ciudad de México, el día <strong>${fechaLetra}</strong>.</p>`;
  }

  // TABLA DE FIRMAS
  const columnas = [];
  if (esPM) {
    const empresaBold = resaltarEmpresa_(suscFormateado);
    columnas.push({
      header: '"EL SUSCRIPTOR"',
      empresa: empresaBold,
      nombre: `<strong>${esc(inferirTrato_(d.rl1)).toUpperCase()} ${esc(formatearNombrePersona_(d.rl1))}</strong>`,
      rol: 'REPRESENTANTE LEGAL'
    });
    if (d.rl2) {
      columnas.push({
        header: '"EL SUSCRIPTOR"',
        empresa: empresaBold,
        nombre: `<strong>${esc(inferirTrato_(d.rl2)).toUpperCase()} ${esc(formatearNombrePersona_(d.rl2))}</strong>`,
        rol: 'REPRESENTANTE LEGAL'
      });
    }
  } else {
    columnas.push({
      header: '"EL SUSCRIPTOR"',
      empresa: null,
      nombre: `<strong>${esc(inferirTrato_(d.nombreSusc)).toUpperCase()} ${esc(formatearNombrePersona_(d.nombreSusc))}</strong>`,
      rol: 'POR PROPIO DERECHO'
    });
  }

  for (let i = 0; i < d.numAvales; i++) {
    const a = d.avales[i];
    columnas.push({
      header: '"EL AVAL"',
      empresa: null,
      nombre: `<strong>${esc(a.trato).toUpperCase()} ${esc(formatearNombrePersona_(a.nombre))}</strong>`,
      rol: 'POR PROPIO DERECHO'
    });
  }

  let firmasHTML = '';
  if (esDigital) {
    firmasHTML = '<div class="firmas-digital">' + columnas.map(c => `
      <div class="firma-digital-bloque">
        <div class="firma-digital-header">${c.header}</div>
        ${c.empresa ? '<div class="firma-digital-empresa">' + c.empresa + '</div>' : ''}
        <div class="firma-digital-nombre">${c.nombre}</div>
        <div class="firma-digital-rol">${c.rol}</div>
      </div>
    `).join('') + '</div>';
  } else {
    const filas = distribuirFirmas_(columnas, esPM, d.rl2 ? 2 : 1);
    firmasHTML = filas.map(fila => {
      const esSola = fila.length === 1;
      const colsHTML = fila.map(c => `
        <div class="firma-col">
          <div class="firma-header">${c.header}</div>
          <div class="firma-linea-aire" style="height:${esUnaPagina ? '50px' : '80px'};">&nbsp;</div>
          <div class="firma-linea">_______________________________</div>
          ${c.empresa ? '<div class="firma-empresa">' + c.empresa + '</div>' : ''}
          <div class="firma-nombre">${c.nombre}</div>
          <div class="firma-rol">${c.rol}</div>
        </div>
      `).join('');
      if (esSola) {
        return `<div class="firmas-fila firmas-fila-sola">${colsHTML}</div>`;
      }
      return `<div class="firmas-fila firmas-fila-par">${colsHTML}</div>`;
    }).join('');
  }

  return `
    <div class="doc-header-mini">
      ${esc(d.codigo)}_V1.0
    </div>
    <h1 class="doc-titulo">PAGARÉ</h1>
    <p class="doc-suma">SUMA PRINCIPAL: ${montoFormateado} ${codigoMoneda}</p>
    <div class="doc-clausulas">${clausulasHTML}</div>
    <div class="firmas-wrap">${firmasHTML}</div>
  `;
}

function renderizarBloqueDomicilios_(body, d, numCl, tieneAval) {
  const grupos = agruparPorDomicilio_(d);

  if (d.numAvales === 0 && grupos.length === 1) {
    const domiFormateado = formatearDomicilio_(grupos[0].domicilio);

    const p = body.appendParagraph('');
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
     .setSpacingAfter(6).setSpacingBefore(2).setLineSpacing(1.15);
    const t = p.editAsText();

    const textoIntro = numCl + '. El Suscriptor designa como domicilio para ' +
                       'requerimiento judicial de pago, el siguiente: ';
    t.appendText(textoIntro);
    t.appendText(domiFormateado + '.');
    t.setFontFamily('Arial').setFontSize(8);
    t.setBold(textoIntro.length, textoIntro.length + domiFormateado.length, true);
    return;
  }

  const domIntro = numCl + '. El Suscriptor y el Aval designan como sus domicilios ' +
                   'para requerimiento judicial de pago, los siguientes:';
  parrJust_(body, domIntro);

  grupos.forEach(g => {
    const domiFormateado = formatearDomicilio_(g.domicilio);
    agregarParrafoDomicilio_(body, g.etiqueta + ': ', domiFormateado);
  });
}

/**
 * Normaliza una dirección para comparar: minúsculas, sin acentos, espacios colapsados.
 * Así "Calle Juárez 45" = "calle juarez 45" = "CALLE  JUAREZ 45".
 */
function normalizarDomicilio_(dir) {
  if (!dir) return '';
  return String(dir).trim()
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')  // quitar acentos
    .replace(/\s+/g, ' ')                              // colapsar espacios
    .replace(/[.,;]+/g, '')                            // quitar puntuación suave
    .trim();
}

/**
 * Agrupa suscriptor + avales por domicilio.
 * Devuelve [{etiqueta, domicilio}, ...] en el orden en que aparecen.
 */
function agruparPorDomicilio_(d) {
  const personas = [];

  // Suscriptor con su nombre formateado (Title Case si PF, razón social si PM)
  const esPM = d.regimen === 'Persona Moral';
  const nombreSuscParaEtiqueta = esPM
    ? formatearNombreEmpresa_(d.nombreSusc).replace(/,$/, '')
    : formatearNombrePersonaTC_(d.nombreSusc);

  personas.push({
    tipo: 'suscriptor',
    nombre: nombreSuscParaEtiqueta,
    numAval: null,
    domiOrig: d.domiSusc,
    domiNorm: normalizarDomicilio_(d.domiSusc)
  });

  for (let i = 0; i < d.numAvales; i++) {
    const a = d.avales[i];
    personas.push({
      tipo: 'aval',
      nombre: formatearNombrePersonaTC_(a.nombre),
      numAval: i + 1,
      domiOrig: a.domi,
      domiNorm: normalizarDomicilio_(a.domi)
    });
  }

  const mapa = new Map();
  const ordenClaves = [];
  personas.forEach(p => {
    if (!p.domiNorm) return;
    if (!mapa.has(p.domiNorm)) {
      mapa.set(p.domiNorm, { domicilio: p.domiOrig, miembros: [] });
      ordenClaves.push(p.domiNorm);
    }
    mapa.get(p.domiNorm).miembros.push(p);
  });

  const grupos = ordenClaves.map(k => mapa.get(k));

  // Calcular cuántos grupos contienen al menos un aval. Si hay más de uno, las
  // etiquetas de los grupos con avales necesitan nombres para distinguirse.
  const gruposConAval = grupos.filter(g =>
    g.miembros.some(m => m.tipo === 'aval')
  );
  const necesitaNombres = gruposConAval.length > 1;

  return grupos.map(g => ({
    etiqueta: construirEtiquetaGrupo_(g.miembros, necesitaNombres),
    domicilio: g.domicilio
  }));
}

/**
 * Genera la etiqueta legal según quiénes comparten el domicilio:
 *
 *   - solo suscriptor:                 "EL SUSCRIPTOR"
 *   - solo 1 aval:                     "EL AVAL (NOMBRE)"
 *   - solo N avales (N≥2):             "LOS AVALES (N1 y N2)" / "LOS AVALES (N1, N2 y N3)"
 *   - suscriptor + 1 aval:             "EL SUSCRIPTOR Y EL AVAL (NOMBRE)"
 *   - suscriptor + N avales (N≥2):     "EL SUSCRIPTOR Y LOS AVALES (N1 y N2)"
 */
function construirEtiquetaGrupo_(miembros, necesitaNombres) {
  const susc = miembros.find(m => m.tipo === 'suscriptor');
  const avales = miembros.filter(m => m.tipo === 'aval');
  const numAv = avales.length;

  // Solo suscriptor (sin avales) — siempre sin nombre
  if (susc && numAv === 0) {
    return 'El Suscriptor';
  }

  // Solo aval(es), sin suscriptor
  if (!susc && numAv >= 1) {
    if (necesitaNombres) {
      return 'El Aval (' + unirNombres_(avales.map(a => a.nombre)) + ')';
    }
    return 'El Aval';
  }

  // Suscriptor + aval(es) en el mismo grupo
  if (susc && numAv >= 1) {
    if (necesitaNombres) {
      const todosLosNombres = [susc.nombre].concat(avales.map(a => a.nombre));
      return 'El Suscriptor y el Aval (' + unirNombres_(todosLosNombres) + ')';
    }
    return 'El Suscriptor y el Aval';
  }

  return 'El Suscriptor';
}


/**
 * Une nombres con "y" natural en español:
 *   ['A']           -> 'A'
 *   ['A', 'B']      -> 'A y B'
 *   ['A', 'B', 'C'] -> 'A, B y C'
 */
function unirNombres_(nombres) {
  if (nombres.length === 0) return '';
  if (nombres.length === 1) return nombres[0];
  if (nombres.length === 2) return nombres[0] + ' y ' + nombres[1];
  const ultimos = nombres.slice(-1)[0];
  const previos = nombres.slice(0, -1);
  return previos.join(', ') + ' y ' + ultimos;
}

function formatearDomicilio_(texto) {
  if (!texto) return '';

  // Palabras que se mantienen en minúsculas (si no van al inicio de la frase)
  const enMinuscula = new Set([
    'de', 'del', 'la', 'las', 'el', 'los', 'y', 'o', 'u',
    'en', 'a', 'al', 'con', 'por', 'para', 'sin', 'sobre',
    'entre', 'bajo', 'hacia', 'hasta', 'desde', 'e', 'ni',
    'y/o', 'o/y'
  ]);

  // Palabras que siempre van en MAYÚSCULAS (siglas y abreviaciones técnicas)
  const enMayuscula = new Set([
    'BIS', 'TER', 'QTR',
    'C.P.', 'CP', 'S/N', 'SN',
    'MZ', 'LT', 'ESQ',
    'S.A.', 'SA', 'S.C.', 'SC', 'S.R.L.', 'SRL',
    'N°', 'NO.', 'NUM.',
    'DF', 'CDMX', 'EDOMEX', 'NL', 'BC', 'BCS'
  ]);

  // Separar por palabras manteniendo delimitadores (incluidas las comillas, que
  // resetean el contador de "primera palabra" para que dentro de "Rincón del Bosque"
  // la R vaya en mayúscula).
  const partes = texto.split(/(\s+|,|;|\.(?![0-9])|"|'|«|»)/);

  let primeraPalabraReal = true;

  const resultado = partes.map(parte => {
    if (/^\s+$/.test(parte)) return parte;

    // Comas, puntos, comillas → la siguiente palabra se trata como inicio de frase
    if (parte === ',' || parte === ';' || parte === '.' ||
        parte === '"' || parte === "'" || parte === '«' || parte === '»') {
      primeraPalabraReal = true;
      return parte;
    }
    if (parte === '') return parte;

    // Número puro → dejar como viene (ej. códigos postales)
    if (/^[0-9]+$/.test(parte)) {
      primeraPalabraReal = false;
      return parte;
    }

    // Alfanumérico mixto como "1BISB" o "21A" → capitalización especial
    if (/[0-9]/.test(parte) && /[A-Za-zÁÉÍÓÚáéíóúÑñ]/.test(parte)) {
      primeraPalabraReal = false;
      return capitalizarAlfanumerico_(parte, enMayuscula);
    }

    const parteUpper = parte.toUpperCase();
    const parteLower = parte.toLowerCase();

    // Sigla conocida (BIS, C.P., S/N, etc.) → siempre mayúsculas
    if (enMayuscula.has(parteUpper)) {
      primeraPalabraReal = false;
      return parteUpper;
    }

    // Letra sola (A, B, C dentro de direcciones) → siempre mayúscula
    if (parte.length === 1 && /[A-Za-zÁÉÍÓÚáéíóúÑñ]/.test(parte)) {
      primeraPalabraReal = false;
      return parteUpper;
    }

    // Preposición/artículo/conjunción → minúsculas (salvo al inicio)
    if (enMinuscula.has(parteLower) && !primeraPalabraReal) {
      return parteLower;
    }

    // Caso general → Title Case
    primeraPalabraReal = false;
    return capitalizarPalabra_(parte);
  });

  return resultado.join('');
}


/**
 * Capitaliza solo la primera letra y mantiene el resto en minúsculas.
 * Respeta caracteres especiales del español (ñ, á, é, etc.).
 */
function capitalizarPalabra_(p) {
  if (!p) return p;
  return p.charAt(0).toUpperCase() + p.slice(1).toLowerCase();
}

/**
 * Capitaliza palabras alfanuméricas tipo "1BISB" o "21A".
 * Separa letras y números; las letras se ponen en mayúsculas si son siglas conocidas,
 * si no, Title Case.
 */
function capitalizarAlfanumerico_(p, enMayuscula) {
  // Separar en grupos de letras y números
  const tokens = p.match(/[0-9]+|[A-Za-zÁÉÍÓÚáéíóúÑñ]+/g) || [p];
  return tokens.map(t => {
    if (/^[0-9]+$/.test(t)) return t;
    const upper = t.toUpperCase();
    if (enMayuscula.has(upper)) return upper;
    // Si es una sola letra, dejar en mayúscula (tipo "B", "A")
    if (t.length === 1) return upper;
    return capitalizarPalabra_(t);
  }).join('');
}


/**
 * Distribuye las columnas de firmantes en filas lógicas de máximo 2.
 * Regla:
 *   - PM con 2 RL: primera fila = [RL1, RL2]
 *   - PM con 1 RL: primera fila = [RL1] (solo, centrado)
 *   - PF: primera fila = [Suscriptor] (solo, centrado)
 *   - Avales: 2 por fila; si al final sobra 1, va solo centrado
 */
function distribuirFirmas_(columnas, esPM, numRL) {
  const filas = [];
  let i = 0;

  // Caso especial PM con 2 RL: ambos RL siempre juntos en la primera fila
  if (esPM && numRL === 2) {
    filas.push([columnas[0], columnas[1]]);
    i = 2;
  }
  // Caso PM con 1 RL o PF: el Suscriptor NO va solo, se empareja con quien siga
  // (ya no hay regla especial aquí — la primera fila se llena como cualquier otra)

  // De aquí en adelante, agrupar de 2 en 2 y el impar al final se centra
  while (i < columnas.length) {
    if (i + 1 < columnas.length) {
      filas.push([columnas[i], columnas[i + 1]]);
      i += 2;
    } else {
      filas.push([columnas[i]]);
      i += 1;
    }
  }

  return filas;
}


function resaltarEmpresa_(texto) {
  if (!texto) return '';
  
  const esc = (s) => s.toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');

  return `<strong>${esc(texto)}</strong>`;
}
function formatearNombrePersonaTC_(nombre) {
  if (!nombre) return '';

  // Preposiciones y artículos que van en minúsculas (salvo al inicio)
  const enMinuscula = new Set([
    'de', 'del', 'la', 'las', 'el', 'los', 'y', 'e',
    'da', 'do', 'das', 'dos',  // portugués, por si acaso
    'di', 'della', 'von', 'van' // europeos comunes
  ]);

  const palabras = nombre.trim().split(/\s+/);
  if (palabras.length === 0) return '';

  return palabras.map((palabra, idx) => {
    if (!palabra) return palabra;

    const lower = palabra.toLowerCase();

    // Primera palabra siempre va capitalizada
    if (idx === 0) {
      return palabra.charAt(0).toUpperCase() + palabra.slice(1).toLowerCase();
    }

    // Preposición → minúscula
    if (enMinuscula.has(lower)) {
      return lower;
    }

    // Palabra normal → Title Case
    return palabra.charAt(0).toUpperCase() + palabra.slice(1).toLowerCase();
  }).join(' ');
}


// =====================================================================================
// [B2] FIRMAS DIGITALES — AGREGAR AL FINAL del archivo
// -------------------------------------------------------------------------------------
// Construye la lista vertical de firmantes (formato pagaré digital con FEA).
// Cada firmante consta de 3-4 párrafos alineados a la izquierda:
//   "EL SUSCRIPTOR" / "EL AVAL"
//   [Razón social — solo si PM]
//   SR. NOMBRE COMPLETO
//   POR PROPIO DERECHO / REPRESENTANTE LEGAL
// Entre firmantes: párrafo separador con espaciado.
// =====================================================================================

function construirListaFirmasDigital_(body, columnas, esPM) {
  columnas.forEach((col, idx) => {
    // Párrafo 1: header (entre comillas)
    const pHeader = body.appendParagraph(col.header);
    pHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT)
           .setLineSpacing(1.15).setSpacingBefore(0).setSpacingAfter(0);
    pHeader.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true);

    // Párrafo opcional: razón social (solo si PM tiene empresa)
    if (esPM && col.empresa) {
      const pEmpresa = body.appendParagraph(col.empresa);
      pEmpresa.setAlignment(DocumentApp.HorizontalAlignment.LEFT)
              .setLineSpacing(1.15).setSpacingBefore(0).setSpacingAfter(0);
      pEmpresa.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true);
    }

    // Párrafo 2: tratamiento + nombre
    const pNombre = body.appendParagraph(col.nombre);
    pNombre.setAlignment(DocumentApp.HorizontalAlignment.LEFT)
           .setLineSpacing(1.15).setSpacingBefore(0).setSpacingAfter(0);
    pNombre.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true);

    // Párrafo 3: rol
    const pRol = body.appendParagraph(col.rol);
    pRol.setAlignment(DocumentApp.HorizontalAlignment.LEFT)
        .setLineSpacing(1.15).setSpacingBefore(0).setSpacingAfter(0);
    pRol.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true);

    // Separador entre firmantes (no después del último)
    if (idx < columnas.length - 1) {
      const sep = body.appendParagraph('\u00A0');
      sep.setSpacingBefore(8).setSpacingAfter(8);
      sep.editAsText().setFontSize(9);
    }
  });
}


// =====================================================================================
// MODAL DE CONFIRMACIÓN DE SOBRESCRITURA (Fuera del Sidebar)
// =====================================================================================
function mostrarConfirmacionSobrescribir(existente) {
  const htmlStr = construirHTMLModalConfirmacion_(existente);
  const html = HtmlService.createHtmlOutput(htmlStr)
    .setWidth(450)
    .setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, '⚠️ Acción Requerida');
}

// =====================================================================================
// SIDEBAR
// =====================================================================================
function mostrarSidebarPagares() {
  const html = HtmlService.createHtmlOutput(SIDEBAR_HTML_PAGARES_).setTitle('📄 Pagarés Cualli').setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

const SIDEBAR_HTML_PAGARES_ = `
<!DOCTYPE html>
<html lang="es">
<head>
  <base target="_top">
  <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root { --comic-ink: #000; --comic-yellow: #fbb818; --comic-paper: #fff; }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Poppins', Arial, sans-serif;
      background: linear-gradient(135deg, #f5f7fa 0%, #e4e8f0 100%);
      padding: 0; min-height: 100vh;
      display: flex; align-items: center; justify-content: center;
    }
    .container {
      background: var(--comic-paper);
      border: 3px solid var(--comic-ink);
      border-radius: 16px;
      box-shadow: 5px 5px 0 var(--comic-ink), 0 10px 25px rgba(0,0,0,.1);
      padding: 24px; width: 100%; max-width: 380px; margin: 5px;
      position: relative; /* Clave para el overlay */
      overflow: hidden;
    }
    h2 {
      display: flex; align-items: center; gap: 12px;
      color: var(--comic-ink); margin-bottom: 25px;
      font-weight: 700; font-size: 1.1rem;
      padding-bottom: 15px; border-bottom: 2px dashed rgba(0,0,0,.1);
      text-transform: uppercase; letter-spacing: .5px;
    }
    .title-icon {
      width: 55px; height: 55px; background: var(--comic-paper);
      padding: 3px; border-radius: 5px; border: 1px solid var(--comic-ink);
      box-shadow: 3px 3px 0 var(--comic-ink);
    }
    .btn-group { display: flex; flex-direction: column; gap: 15px; margin: 20px 0; }
    .btn {
      display: flex; align-items: flex-start; gap: 12px;
      padding: 20px 20px 20px 5px;
      border: 3px solid var(--comic-ink); border-radius: 16px; cursor: pointer;
      font-weight: 700; font-size: 14px;
      transition: all .3s cubic-bezier(.25,.8,.25,1);
      text-align: left; color: var(--comic-ink);
      box-shadow: 4px 4px 0 var(--comic-ink);
      text-transform: uppercase; letter-spacing: .8px;
      width: 100%; min-height: 90px;
    }
    .btn-content { display: flex; flex-direction: column; flex: 1; }
    .btn-title { font-weight: 800; font-size: 13px; margin-bottom: 4px; line-height: 1.2; }
    .btn-desc { font-size: 10px; opacity: .9; font-weight: 600; line-height: 1.3; }
    .btn-icon {
      width: 38px; height: 38px; flex-shrink: 0; display: flex;
      align-items: center; justify-content: center; border-radius: 10px;
      background: #fff; border: 3px solid var(--comic-ink);
      box-shadow: 1px 1px 0 var(--comic-ink);
    }
    .btn-icon img { width: 28px; height: 28px; object-fit: contain; }
    .btn:hover { transform: translate(-3px,-3px); box-shadow: 6px 6px 0 var(--comic-ink); }
    .btn:active { transform: translate(1px,1px); box-shadow: 2px 2px 0 var(--comic-ink); }
    
    /* Colores de Botones */
    .btn-validar { background: linear-gradient(145deg,#2196f3,#1976d2); border-color: #1565c0; }
    .btn-preview { background: linear-gradient(145deg,#9c27b0,#7b1fa2); border-color: #6a1b9a; }
    .btn-generar { background: linear-gradient(145deg,#4caf50,#43a047); border-color: #2e7d32; }
    .btn-guardar { background: linear-gradient(145deg,#f39c12,#d35400); border-color: #e67e22; }
    .btn-cargar  { background: linear-gradient(145deg,#fbb818,#FFD46F); border-color: #E8A300; }
    .btn-limpiar { background: linear-gradient(145deg,#e74c3c,#c0392b); border-color: #c0392b; }
    .btn-maint   { background: linear-gradient(145deg,#7f8c8d,#34495e); border-color: #2c3e50; }
    .btn .btn-title, .btn .btn-desc { color: white; text-shadow: 1px 1px 0 rgba(0,0,0,.3); }

    .footer { margin-top: 20px; text-align: center; font-size: 11px; color: #666; padding-top: 15px; border-top: 1px dashed rgba(0,0,0,.1); }

    /* ======================================================================
       SISTEMA DE OVERLAY INTERNO (Bloquea interacciones en el Sidebar)
       ====================================================================== */
    .app-overlay {
      position: absolute; top: 0; left: 0; width: 100%; height: 100%;
      background: rgba(255, 255, 255, 0.85); backdrop-filter: blur(5px);
      z-index: 1000; display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      opacity: 0; visibility: hidden; transition: all 0.2s ease;
      padding: 20px;
    }
    .app-overlay.active { opacity: 1; visibility: visible; }
    
    .overlay-card {
      background: var(--comic-paper); border: 3px solid var(--comic-ink); border-radius: 12px;
      box-shadow: 6px 6px 0 var(--comic-ink); padding: 25px 20px; text-align: center; width: 100%;
      transform: translateY(20px); transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
      display: flex; flex-direction: column; align-items: center;
    }
    .app-overlay.active .overlay-card { transform: translateY(0); }
    
    .overlay-icon { font-size: 45px; margin-bottom: 10px; line-height: 1; }
    .overlay-title { font-size: 16px; font-weight: 800; text-transform: uppercase; margin-bottom: 8px; color: var(--comic-ink); }
    .overlay-desc { font-size: 12px; color: #444; font-weight: 500; line-height: 1.4; margin-bottom: 15px; width: 100%; max-height: 150px; overflow-y: auto; text-align: center;}
    .overlay-desc.left-align { text-align: left; background: #f5f5f5; padding: 10px; border-radius: 6px; border: 1px dashed #ccc;}
    
    .overlay-actions { display: flex; gap: 10px; width: 100%; }
    .btn-small {
      padding: 12px 10px; border: 2px solid var(--comic-ink); border-radius: 8px;
      font-weight: 700; font-size: 12px; cursor: pointer; text-transform: uppercase;
      flex: 1; transition: all 0.2s; box-shadow: 2px 2px 0 var(--comic-ink); font-family: 'Poppins';
    }
    .btn-small:hover { transform: translate(-2px, -2px); box-shadow: 4px 4px 0 var(--comic-ink); }

    /* Spinner Animado */
    .spinner-giant {
      width: 45px; height: 45px; border: 5px solid #e3e3e3; border-top: 5px solid var(--comic-yellow);
      border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto;
    }
    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
  </style>
</head>
<body>
  <div class="container">
    <h2>
      <img src="https://cualli.mx/wp-content/uploads/2022/07/cualli-bl@3x.png" class="title-icon">
      Pagarés · Cualli
    </h2>
    
    <div class="btn-group">
      <button class="btn btn-validar" onclick="validar()">
        <div class="btn-icon"><img src="https://cdn-icons-png.flaticon.com/512/6135/6135181.png"></div>
        <div class="btn-content">
          <div class="btn-title">Valida los Datos</div>
          <div class="btn-desc">Revisa que la hoja esté completa</div>
        </div>
      </button>
      
      <button class="btn btn-preview" onclick="preview()">
        <div class="btn-icon"><img src="https://cdn-icons-png.flaticon.com/512/1011/1011884.png"></div>
        <div class="btn-content">
          <div class="btn-title">Previsualiza el Pagaré</div>
          <div class="btn-desc">Vista previa antes de generar</div>
        </div>
      </button>
      
      <button class="btn btn-generar" onclick="generar()">
        <div class="btn-icon"><img src="https://cdn-icons-png.flaticon.com/512/7035/7035543.png"></div>
        <div class="btn-content">
          <div class="btn-title">Genera el Pagaré</div>
          <div class="btn-desc">Crea el Google Doc en Drive</div>
        </div>
      </button>
      
      <button class="btn btn-guardar" onclick="guardarRegistro()">
        <div class="btn-icon"><img src="https://cdn-icons-png.freepik.com/512/11352/11352452.png?ga=GA1.1.229237360.1777314028" style="filter: grayscale(0%); opacity: 0.7;"></div>
        <div class="btn-content">
          <div class="btn-title">Guardar Registro</div>
          <div class="btn-desc">Guarda tu progreso en la base</div>
        </div>
      </button>

      <button class="btn btn-cargar" onclick="cargarRegistro()">
        <div class="btn-icon"><img src="https://cdn-icons-png.flaticon.com/512/3097/3097412.png" style="filter: grayscale(0%); opacity: 0.7;"></div>
        <div class="btn-content">
          <div class="btn-title">Cargar Registro</div>
          <div class="btn-desc">Digita el ID en C15 para recuperar</div>
        </div>
      </button>
      
      <button class="btn btn-limpiar" onclick="limpiar()">
        <div class="btn-icon"><img src="https://cdn-icons-png.freepik.com/512/8389/8389871.png?ga=GA1.1.229237360.1777314028" style="filter: grayscale(0%); opacity: 0.7;"></div>
        <div class="btn-content">
          <div class="btn-title">Limpiar Formulario</div>
          <div class="btn-desc">Borra todos los campos capturados</div>
        </div>
      </button>

       <button class="btn btn-maint" onclick="mantenimiento()">
        <div class="btn-icon"><img src="https://cdn-icons-png.flaticon.com/512/3524/3524659.png" style="filter: grayscale(0%); opacity: 0.7;"></div>
        <div class="btn-content">
          <div class="btn-title">Restaurar Sistema</div>
          <div class="btn-desc">Regenerar hojas del módulo</div>
        </div>
      </button>
    </div>

    <div class="footer">© ${new Date().getFullYear()} Financiera Cualli.</div>

    <div id="app-overlay" class="app-overlay">
      <div class="overlay-card">
        <div id="ov-icon" class="overlay-icon"></div>
        <h3 id="ov-title" class="overlay-title"></h3>
        <div id="ov-desc" class="overlay-desc"></div>
        <div id="ov-actions" class="overlay-actions"></div>
      </div>
    </div>

  </div>

  <script>
    // =========================================================
    // MOTOR DEL OVERLAY (Control de Interfaz Centralizado)
    // =========================================================
    const Overlay = {
      el: document.getElementById('app-overlay'),
      icon: document.getElementById('ov-icon'),
      title: document.getElementById('ov-title'),
      desc: document.getElementById('ov-desc'),
      actions: document.getElementById('ov-actions'),
      timeout: null,

      show: function() { this.el.classList.add('active'); },
      hide: function() { this.el.classList.remove('active'); clearTimeout(this.timeout); },

      // Estado 1: Procesando (Spinner sin botones)
      loading: function(actionText) {
        clearTimeout(this.timeout);
        this.icon.innerHTML = '<div class="spinner-giant"></div>';
        this.title.innerText = 'PROCESANDO...';
        this.title.style.color = 'var(--comic-yellow)';
        this.desc.innerText = actionText || 'Por favor espera un momento';
        this.desc.className = 'overlay-desc';
        this.actions.innerHTML = '';
        this.show();
      },

      // Estado 2: Éxito (Check verde + Botón Entendido + Auto-Cierre opcional)
      success: function(title, msg, autoClose = false) {
        clearTimeout(this.timeout);
        this.icon.innerHTML = '✅';
        this.title.innerText = title;
        this.title.style.color = '#27ae60';
        this.desc.innerHTML = msg.replace(/\\n/g, '<br>');
        this.desc.className = 'overlay-desc';
        this.actions.innerHTML = '<button class="btn-small" style="background:#27ae60;color:white;" onclick="Overlay.hide()">ENTENDIDO</button>';
        this.show();
        if (autoClose) this.timeout = setTimeout(() => this.hide(), 5000);
      },

      // Estado 3: Error (Tache rojo + Detalles justificados + Botón Entendido)
      error: function(title, msg) {
        clearTimeout(this.timeout);
        this.icon.innerHTML = '❌';
        this.title.innerText = title;
        this.title.style.color = '#e74c3c';
        this.desc.innerHTML = msg.replace(/\\n/g, '<br>');
        this.desc.className = 'overlay-desc left-align'; // Justificado para leer errores
        this.actions.innerHTML = '<button class="btn-small" style="background:#e74c3c;color:white;" onclick="Overlay.hide()">ENTENDIDO</button>';
        this.show();
      },

      // Estado 4: Confirmación (Icono de alerta + 2 botones)
      confirm: function(icon, title, msg, btnConfirmText, jsFunctionString) {
        clearTimeout(this.timeout);
        this.icon.innerHTML = icon;
        this.title.innerText = title;
        this.title.style.color = '#e74c3c';
        this.desc.innerHTML = msg.replace(/\\n/g, '<br>');
        this.desc.className = 'overlay-desc';
        this.actions.innerHTML = \`
          <button class="btn-small" style="background:#eaeded;color:#333;" onclick="Overlay.hide()">CANCELAR</button>
          <button class="btn-small" style="background:#e74c3c;color:white;" onclick="\${jsFunctionString}">\${btnConfirmText}</button>
        \`;
        this.show();
      }
    };


    // =========================================================
    // FUNCIONES CONECTADAS AL BACKEND
    // =========================================================
    function validar() {
      Overlay.loading('Validando los datos capturados...');
      google.script.run
        .withSuccessHandler(function(res){
          if (res.ok) {
            Overlay.success('VALIDACIÓN EXITOSA', 'El formulario está completo.\\nPlantilla: ' + res.codigo, true);
          } else {
            Overlay.error('FALTAN DATOS', 'Por favor completa lo siguiente:\\n\\n' + res.errs.join('\\n'));
          }
        })
        .withFailureHandler(function(err){
          Overlay.error('ERROR', err.message || err);
        })
        .validarDatosDesdeSidebar();
    }

    function preview() {
      Overlay.loading('Generando vista previa...');
      google.script.run
        .withSuccessHandler(function(){
          Overlay.success('LISTO', 'Revisa la vista previa en el centro de tu pantalla.', true);
        })
        .withFailureHandler(function(err){
          Overlay.error('ERROR', err.message || err);
        })
        .previewPagareDesdeSidebar();
    }

    function generar() {
      Overlay.loading('Verificando el historial...');
      google.script.run
        .withSuccessHandler(function(existente){
          if (existente) {
            Overlay.confirm('📜', 'PAGARÉ EXISTENTE', 'Ya existe un pagaré generado para este ID Oportunidad.\\n\\n¿Deseas sobrescribirlo con los datos actuales?', 'SOBRESCRIBIR', 'ejecutarGeneracion(true)');
          } else {
            ejecutarGeneracion(false);
          }
        })
        .withFailureHandler(function(err){
          Overlay.error('ERROR DE RED', err.message || err);
        })
        .verificarPagareExistenteDesdeSidebar();
    }
    
    function ejecutarGeneracion(sobrescribir) {
      Overlay.loading('Generando el documento en Google Docs...');
      google.script.run
        .withSuccessHandler(function(url){
          Overlay.success('¡PAGARÉ LISTO! 🚀', 'El documento se generó exitosamente y se abrirá en una nueva pestaña.', true);
        })
        .withFailureHandler(function(err){
          Overlay.error('ERROR AL GENERAR', err.message || err);
        })
        .generarPagareDesdeSidebar({sobrescribir: sobrescribir});
    }

    function guardarRegistro() {
      Overlay.loading('Guardando información en la Base de Datos...');
      google.script.run
        .withSuccessHandler(function(){
          Overlay.success('GUARDADO 💾', 'El progreso del pagaré se ha registrado exitosamente.', true);
        })
        .withFailureHandler(function(err){
          Overlay.error('ERROR AL GUARDAR', err.message || err);
        })
        .guardarBorradorDesdeSidebar();
    }

    function cargarRegistro() {
      Overlay.loading('Buscando ID Oportunidad...');
      google.script.run
        .withSuccessHandler(function(){
          Overlay.success('CARGA EXITOSA 📂', 'Los datos han sido recuperados y plasmados en el formulario.', true);
        })
        .withFailureHandler(function(err){
          Overlay.error('ERROR AL CARGAR', err.message || err);
        })
        .cargarBorradorDesdeSidebar();
    }

    function limpiar() {
      Overlay.confirm('🧹', 'LIMPIAR FORMULARIO', '¿Estás seguro de borrar todos los datos capturados?\\nEsta acción no se puede deshacer.', 'SÍ, LIMPIAR', 'ejecutarLimpieza()');
    }

    function ejecutarLimpieza() {
      Overlay.loading('Restableciendo el formulario...');
      google.script.run
        .withSuccessHandler(function(){
          Overlay.success('FORMULARIO LIMPIO ✨', 'Listo para una nueva captura.', true);
        })
        .withFailureHandler(function(err){
          Overlay.error('ERROR', err.message || err);
        })
        .limpiarFormularioDesdeSidebar();
    }

    // --- NUEVA FUNCIÓN PARA EL BOTÓN DE MANTENIMIENTO ---
    function mantenimiento() {
      Overlay.loading('Preparando restauración...');
      
      // Hacemos una llamada rápida al backend para asegurarnos que está activo,
      // y si responde bien, ocultamos el overlay y llamamos a tu función nativa con sus propios menús.
      google.script.run
        .withSuccessHandler(function() {
          Overlay.hide();
          google.script.run.crearHojasPagare(); 
        })
        .withFailureHandler(function(err) {
          Overlay.error('ERROR', err.message || err);
        })
        .validarDatosDesdeSidebar(); 
    }
  </script>
</body>
</html>
`;


function construirHTMLModalConfirmacion_(existente) {
  // 1. Formatear la Fecha
  let fechaLegible = "Fecha no disponible";
  if (existente.timestamp) {
    const fechaObj = new Date(existente.timestamp);
    const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    
    const dia = fechaObj.getDate();
    const mes = meses[fechaObj.getMonth()];
    const anio = fechaObj.getFullYear();
    let horas = fechaObj.getHours();
    const minutos = fechaObj.getMinutes().toString().padStart(2, '0');
    const ampm = horas >= 12 ? 'PM' : 'AM';
    horas = horas % 12;
    horas = horas ? horas : 12; 
    
    fechaLegible = `${dia} de ${mes} de ${anio} a las ${horas}:${minutos} ${ampm}`;
  }

  // 2. "Traducir" el código de la plantilla
  let plantillaDesc = existente.plantilla || "Estándar";
  if (plantillaDesc.includes('Pagaré')) {
    let detalles = [];
    if (plantillaDesc.includes('PM')) detalles.push('Persona Moral');
    if (plantillaDesc.includes('PF')) detalles.push('Persona Física');
    if (plantillaDesc.includes('A')) detalles.push('con Avales');
    if (plantillaDesc.includes('_D')) detalles.push('Firma Digital');
    plantillaDesc = detalles.join(' ');
  }

  // 3. Capturar el ID Oportunidad
  const idOportunidad = existente.idOp || "ID Desconocido";

  return `
  <!DOCTYPE html>
  <html>
  <head>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
      :root { --comic-ink: #000; --comic-yellow: #fbb818; --comic-paper: #fff; }
      body { 
        font-family: 'Poppins', Arial, sans-serif; 
        background: #f5f7fa; padding: 20px; margin: 0;
        display: flex; flex-direction: column; align-items: center; justify-content: center;
      }
      .modal-container {
        background: var(--comic-paper);
        border: 3px solid var(--comic-ink);
        border-radius: 12px;
        box-shadow: 4px 4px 0 var(--comic-ink);
        padding: 20px; width: 100%; box-sizing: border-box; text-align: center;
      }
      .icon { font-size: 35px; margin-bottom: 5px; }
      h2 { color: var(--comic-ink); margin: 0 0 10px 0; font-weight: 700; font-size: 18px; text-transform: uppercase; }
      
      .info-box {
        background: #fff9e6;
        border: 2px dashed #f1c40f;
        border-radius: 8px;
        padding: 15px; text-align: left;
        font-size: 13px; color: #333; margin-bottom: 20px;
        line-height: 1.6;
      }
      .info-label { font-weight: 700; color: #7f8c8d; text-transform: uppercase; font-size: 10px; display: block; margin-bottom: -2px; }
      .info-value { font-weight: 600; color: #2c3e50; display: block; margin-bottom: 8px; }
      .id-destacado { font-family: monospace; font-size: 14px; color: #d35400; font-weight: 700; letter-spacing: 0.5px; }
      
      .btn-group { display: flex; gap: 15px; justify-content: center; }
      .btn {
        padding: 12px 15px; border: 2px solid var(--comic-ink); border-radius: 8px;
        font-weight: 700; font-size: 13px; font-family: 'Poppins'; cursor: pointer;
        transition: all 0.2s; text-transform: uppercase; width: 100%;
      }
      .btn-cancel { background: #eaeded; color: #333; box-shadow: 2px 2px 0 var(--comic-ink); }
      .btn-confirm { background: #e74c3c; color: white; box-shadow: 2px 2px 0 var(--comic-ink); }
      .btn:hover { transform: translate(-1px, -1px); box-shadow: 3px 3px 0 var(--comic-ink); }
      
      .loader { display: none; font-weight: 700; font-size: 13px; color: #27ae60; margin-top: 15px; background: #e8f5e9; padding: 10px; border-radius: 8px; border: 2px dashed #27ae60;}
    </style>
  </head>
  <body>
    <div class="modal-container">
      <div class="icon">📜</div>
      <h2>Pagaré ya generado</h2>
      <p style="font-size: 12px; color: #666;">Se encontró un documento previo para esta oportunidad:</p>
      
      <div class="info-box">
        <span class="info-label">ID Oportunidad</span>
        <span class="info-value id-destacado">${idOportunidad}</span>

        <span class="info-label">Emitido el</span>
        <span class="info-value">${fechaLegible}</span>
        
        <span class="info-label">Tipo de Plantilla</span>
        <span class="info-value">${plantillaDesc}</span>
      </div>
      
      <p style="font-size: 12px; font-weight: 600; margin-bottom: 20px;">¿Deseas reemplazarlo con la información actual?</p>
      
      <div class="btn-group" id="botones">
        <button class="btn btn-cancel" onclick="google.script.host.close()">No, volver</button>
        <button class="btn btn-confirm" onclick="sobrescribir()">Sí, sobrescribir</button>
      </div>
      
      <div id="loader" class="loader">⏳ Procesando cambio...</div>
    </div>

    <script>
      function sobrescribir() {
        document.getElementById('botones').style.display = 'none';
        document.getElementById('loader').style.display = 'block';
        google.script.run
          .withSuccessHandler(function() {
            document.getElementById('loader').innerHTML = '✅ ¡Actualizado con éxito!';
            setTimeout(function() { google.script.host.close(); }, 2000);
          })
          .generarPagareDesdeSidebar({sobrescribir: true});
      }
    </script>
  </body>
  </html>`;
}

function aplicarProteccionesPagare_(sh) {
  // 1. Limpiar protecciones previas para evitar duplicados [cite: 496]
  const protecciones = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protecciones.forEach(p => p.remove());

  // 2. Proteger la hoja entera. Nadie podrá editar nada por defecto [cite: 497]
  const proteccion = sh.protect().setDescription('Blindaje Total Interfaz Pagarés');

  // 3. Definir EXCLUSIVAMENTE los rangos editables (Sección I a III) [cite: 498]
  const celdasEditables = [
    'C9', 'E9',       // Régimen, Tipo Firma [cite: 26]
    'C11',            // No. de Avales [cite: 28]
    'C15', 'E15',     // ID Oportunidad, Divisa [cite: 33]
    'C17', 'E17',     // Monto Principal, Tasa [cite: 34]
    'C23',            // Fecha de Firma [cite: 41]
    'C25', 'E25',     // Plazo, Unidad (E25 es el selector azul) [cite: 42, 43]
    'C31:E31',        // Nombre Suscriptor [cite: 49]
    'C33:E33',        // Domicilio Suscriptor [cite: 51]
    'C35', 'E35',      // Representante 1 y 2 [cite: 52]
    
    // RANGOS DEL PANEL DE REFERENCIA (Columna H y Botón G17)
    'H9',             // Entrada de ID Oportunidad en el Panel
    'H10:H16',        // Resultados de consulta (permitir selección para copiar)
    'G17'             // Celda del Botón "Actualizar"
  ];

  // 4. Agregar dinámicamente los campos de los 10 avales (Sección IV) [cite: 499]
  // Los avales van de la fila 39 a la 78 aprox. [cite: 506]
  let currentRow = 39;
  for (let i = 1; i <= 10; i++) {
    celdasEditables.push('E' + (currentRow + 1));             // Nombre Aval [cite: 500, 506]
    celdasEditables.push('C' + (currentRow + 2) + ':E' + (currentRow + 2)); // Domicilio Aval [cite: 501, 507]
    currentRow += 4;
  }

  // 5. Convertir notaciones A1 a objetos de rango y desbloquearlos 
  const rangosDesbloqueados = celdasEditables.map(a1 => sh.getRange(a1));
  proteccion.setUnprotectedRanges(rangosDesbloqueados);
  
  // 6. Restricción adicional: Solo tú y el propietario pueden editar el resto
  // (Opcional: puedes añadir correos específicos aquí)
  proteccion.removeEditors(proteccion.getEditors());
  if (proteccion.canDomainEdit()) {
    proteccion.setDomainEdit(false);
  }
}

function handleRestaurarFormatoPagares_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (CONFIG_PAGARES.HOJAS_VALIDAS.indexOf(sh.getName()) === -1) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const a1 = e.range.getA1Notation();
  const GRIS_BORDE = '#E5E7EB';

  let tipoCelda = null;
  let columnasACombinar = 1; // Por defecto es una sola celda

  // 1. Mapeo de celdas simples
  const inputsSimples = ['C9', 'E9', 'C11', 'C15', 'E15', 'C17', 'E17', 'C23', 'C25', 'C35', 'E35'];
  
  if (inputsSimples.includes(a1)) {
    tipoCelda = 'BLANCO';
  } else if (a1 === 'H9') {
    tipoCelda = 'AMARILLO';
  } else if (a1 === 'E25') {
    tipoCelda = 'AZUL';
  } 
  // 2. Mapeo de campos COMBINADOS (C, D y E)
  else if (a1 === 'C31' || a1 === 'C33') {
    tipoCelda = 'BLANCO';
    columnasACombinar = 3; // C, D y E
  } 
  // 3. Resultados del Panel (H10:H16)
  else if (col === 8 && row >= 10 && row <= 16) {
    tipoCelda = 'BLANCO_SMALL';
  }
  // 4. Campos dinámicos de Avales
  else if (row >= 40 && row <= 78) {
    if (col === 5 && (row - 40) % 4 === 0) {
      tipoCelda = 'BLANCO'; // Nombre Aval (Simple)
    }
    if (col === 3 && (row - 41) % 4 === 0) {
      tipoCelda = 'BLANCO'; // Domicilio Aval (Combinado)
      columnasACombinar = 3;
    }
  }

  // 5. Aplicar la "Auto-Curación" de formato y combinación
  if (tipoCelda) {
    // Definimos el rango real (si debe ser 1 celda o 3 combinadas)
    const target = sh.getRange(row, col, 1, columnasACombinar);

    // Si el pegado rompió la combinación, la forzamos de nuevo
    if (columnasACombinar > 1) {
      target.breakApart(); // Limpiamos restos por si acaso
      target.merge();      // Volvemos a combinar C-D-E
    }

    let colorFondo = '#FFFFFF';
    let colorTexto = '#1B2631';
    let tamanoFuente = 11;
    let alineacion = (columnasACombinar > 1) ? 'left' : 'left';

    if (tipoCelda === 'AMARILLO') colorFondo = CONFIG_PAGARES.COLOR_AMARILLO_TENUE;
    if (tipoCelda === 'AZUL') {
      colorFondo = '#2E4053';
      colorTexto = '#FFFFFF';
      tamanoFuente = 9;
      alineacion = 'center';
    }
    if (tipoCelda === 'BLANCO_SMALL') tamanoFuente = 9;

    // Aplicar el blindaje visual completo
    target.setBackground(colorFondo)
          .setFontColor(colorTexto)
          .setFontWeight('bold')
          .setFontFamily('Arial')
          .setFontSize(tamanoFuente)
          .setHorizontalAlignment(alineacion)
          .setVerticalAlignment('middle')
          .setFontStyle('normal')
          .setBorder(true, true, true, true, null, null, GRIS_BORDE, SpreadsheetApp.BorderStyle.SOLID);
  }
}

// =====================================================================================
// SISTEMA DE GUARDADO / CARGA (JSON PAYLOAD)
// =====================================================================================

function actualizarCabeceraLog_() {
  const sh = obtenerOcreaHojaLog_();
  // Agregamos la columna 9 para el paquete JSON
  sh.getRange('A1:I1').setValues([[
    'Timestamp', 'Usuario', 'ID Oportunidad', 'Plantilla',
    'Suscriptor', 'Doc ID', 'URL', 'Estatus', 'Payload JSON'
  ]]).setFontWeight('bold').setBackground('#515151').setFontColor('#FDB913');
}

function guardarBorradorDesdeSidebar() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (CONFIG_PAGARES.HOJAS_VALIDAS.indexOf(sh.getName()) === -1) {
    throw new Error('Debes estar en PAGARE_1 o PAGARE_2 para guardar.');
  }

  const d = leerDatosPagare_(sh); // Usamos tu función existente
  if (!d.idOp) {
    throw new Error('Debes escribir al menos el ID Oportunidad (C15) para guardar el registro.');
  }

  actualizarCabeceraLog_();
  const logSh = obtenerOcreaHojaLog_();
  const jsonPayload = JSON.stringify(d); // Empaquetamos todo
  const tz = Session.getScriptTimeZone() || 'America/Mexico_City';
  const ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail() || 'Usuario';

  const existente = buscarPagareEnLog_(d.idOp);

  if (existente && existente.fila) {
    // Actualizar registro existente
    logSh.getRange(existente.fila, 1).setValue(ts);
    logSh.getRange(existente.fila, 4).setValue(d.codigo);
    logSh.getRange(existente.fila, 5).setValue(d.nombreSusc);
    logSh.getRange(existente.fila, 8).setValue('BORRADOR / ACTUALIZADO');
    logSh.getRange(existente.fila, 9).setValue(jsonPayload);
  } else {
    // Crear nuevo registro
    logSh.appendRow([ts, user, d.idOp, d.codigo, d.nombreSusc, '', '', 'BORRADOR', jsonPayload]);
  }
  
  return true;
}

function cargarBorradorDesdeSidebar() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (CONFIG_PAGARES.HOJAS_VALIDAS.indexOf(sh.getName()) === -1) {
    throw new Error('Debes estar en PAGARE_1 o PAGARE_2 para cargar.');
  }

  // Leemos el ID desde C15 (o desde el panel si prefieres)
  const idBuscado = sh.getRange('C15').getValue().toString().trim();
  if (!idBuscado) {
    throw new Error('Escribe el ID Oportunidad en la celda C15 y presiona Cargar.');
  }

  const logSh = obtenerOcreaHojaLog_();
  const datos = logSh.getDataRange().getValues();
  let jsonEncontrado = null;

  // Buscar de abajo hacia arriba para traer el más reciente
  for (let i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][2]).trim() === idBuscado) {
      jsonEncontrado = datos[i][8]; // Columna I
      break;
    }
  }

  if (!jsonEncontrado) {
    throw new Error(`No existe ningún registro guardado para el ID: ${idBuscado}`);
  }

  // Desempaquetar y pintar en la hoja
  const d = JSON.parse(jsonEncontrado);

  sh.getRange('C9').setValue(d.regimen || '');
  sh.getRange('E9').setValue(d.tipoFirma || '');
  sh.getRange('C11').setValue(d.numAvales || 0);
  sh.getRange('C15').setValue(d.idOp || '');
  sh.getRange('E15').setValue(d.moneda || 'MXN');
  sh.getRange('C17').setValue(d.montoNum || '');
  sh.getRange('E17').setValue(d.tasaNum || '');
  let fechaFirmaLimpia = d.fechaFirma || '';
  if (fechaFirmaLimpia && typeof fechaFirmaLimpia === 'string' && fechaFirmaLimpia.includes('T')) {
    fechaFirmaLimpia = new Date(fechaFirmaLimpia);
  }
  sh.getRange('C23').setValue(fechaFirmaLimpia);
  sh.getRange('C25').setValue(d.plazoNum || '');
  let unidadGuardada = String(d.plazoUnidad || 'Meses').toLowerCase().trim();
  let unidadValida = (unidadGuardada === 'días' || unidadGuardada === 'dias') ? 'Días' : 'Meses';
  sh.getRange('E25').setValue(unidadValida);
  sh.getRange('C31').setValue(d.nombreSusc || '');
  sh.getRange('C33').setValue(d.domiSusc || '');
  sh.getRange('C35').setValue(d.rl1 || '');
  sh.getRange('E35').setValue(d.rl2 || '');

  // Limpiar Avales actuales primero
  let currentRow = 39;
  for (let i = 1; i <= 10; i++) {
    sh.getRange('E' + (currentRow + 1)).clearContent();
    sh.getRange('C' + (currentRow + 2)).clearContent();
    currentRow += 4;
  }

  // Poblar Avales guardados
  currentRow = 39;
  if (d.avales && d.avales.length > 0) {
    for (let i = 0; i < d.numAvales; i++) {
      sh.getRange('E' + (currentRow + 1)).setValue(d.avales[i].nombre || '');
      sh.getRange('C' + (currentRow + 2)).setValue(d.avales[i].domi || '');
      currentRow += 4;
    }
  }

  return true;
}

function limpiarFormularioDesdeSidebar() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (CONFIG_PAGARES.HOJAS_VALIDAS.indexOf(sh.getName()) === -1) {
    throw new Error('Debes estar en PAGARE_1 o PAGARE_2 para limpiar.');
  }

  // 1. Limpiar celdas de texto simple
  const celdasTexto = [
    'C15', 'C17', 'E17', 'C23', 'C31', 'C33', 'C35', 'E35', 'H9'
  ];
  celdasTexto.forEach(a1 => sh.getRange(a1).clearContent());

  // 2. Restablecer valores por defecto (Listas desplegables y números)
  sh.getRange('C9').setValue('Persona Física');
  sh.getRange('E9').setValue('Autógrafa');
  sh.getRange('C11').setValue(0);
  sh.getRange('E15').setValue('MXN');
  sh.getRange('C25').clearContent();



  // 4. Limpiar los 10 Avales
  let currentRow = 39;
  for (let i = 1; i <= 10; i++) {
    sh.getRange('E' + (currentRow + 1)).clearContent(); // Nombre
    sh.getRange('C' + (currentRow + 2)).clearContent(); // Domicilio
    currentRow += 4;
  }

  return true;
}

function sincronizarListaOportunidadesC15() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shBD = ss.getSheetByName('BD');
  
  if (!shBD) {
    console.error('No se encontró la hoja "BD"');
    return;
  }

  // 1. Obtener todos los datos de la base de datos
  const datosBD = shBD.getDataRange().getValues();
  if (datosBD.length < 2) return; // Si solo hay encabezados, salir

  const headers = datosBD[0];
  const idxID = headers.indexOf('Nombre de la oportunidad');
  
  if (idxID === -1) {
    console.error('No se encontró la columna "Nombre de la oportunidad"');
    return;
  }

  // 2. Extraer los IDs únicos y eliminar celdas vacías
  const listaIds = datosBD.slice(1)
    .map(fila => String(fila[idxID]).trim())
    .filter(id => id !== "");

  // 3. Crear la regla de validación de datos (el menú desplegable)
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(listaIds, true)
    .setAllowInvalid(false) // No permite escribir IDs que no existan
    .setHelpText('Selecciona un ID Oportunidad válido de la lista.')
    .build();

  // 4. Aplicar la regla a la celda C15 en cada hoja de pagaré válida
  CONFIG_PAGARES.HOJAS_VALIDAS.forEach(nombre => {
    let hoja = ss.getSheetByName(nombre);
    if (hoja) {
      hoja.getRange('C15').setDataValidation(regla);
    }
  });

  console.log('✅ Sincronización de lista C15 completada.');
}
