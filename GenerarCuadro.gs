function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActual = ss.getActiveSheet().getName();
  sincronizarListaOportunidadesDatos();
  
  // 1. Lógica para el menú de Cuadro de Transferencias
  let menuTitulo = '💰 Cuadro de Transferencias';
  if (hojaActual.toUpperCase() === 'DATOS_2') {
    menuTitulo = '💰 Cuadro de Transferencias (Hoja 2)';
  }
  
  ui.createMenu(menuTitulo)
    .addItem('🛠️ Abrir Herramientas', 'mostrarSidebarCuadroDispersion')
    .addToUi();

  // 2. Invocamos la creación del menú de Pagarés (la función que ya tienes en tu script)
  crearMenuPagares_();
}

/**
 * Menu para la herramienta de pagares
 */
function crearMenuPagares_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 Pagarés')
    .addItem('🛠️ Abrir Herramientas de Pagarés', 'mostrarSidebarPagares')
    .addToUi();
}


const CONFIG = {
  SHEET_NAME: 'Datos',
  START_ROW: 29,
  COLUMNAS: {
    NUM: 1,
    FECHA: 2,
    CONCEPTO: 3,
    MONTO: 4,
    COMENT: 6,
    BANCO: 7,
    CLABE: 8,
    BENEFI: 9
  },
  CELDAS: {
    ID_OPORT: 'B3',
    ACREDITADO: 'B5',
    REGIMEN: 'B7',
    CLIENTE: 'B17',
    REP1: 'B23',   
    REP2: 'B25',
    MONEDA: 'E16'    
  },
  OUTPUT_FOLDER_ID: null,
  TEMPLATE_DOC_ID: '',
  COLOR_HEADER: '#b7b7b7',
  COLOR_SUBHDR: '#b7b7b7',
  ALTURA: {
    TABLA_BASE: 120,
    EXTRA_CADA_50_CH: 12,
    MIN_INICIO: 90,
    UTIL_PAG: 500,
    AUTORIZACION_APROX: 180
  },
  MARGEN_SEGURIDAD: 24,
  MARGENES: { top: 1.20, bottom: 2.44, left: 3.20, right: 3.00 },
  DOCUMENTO: {
    PAGE_HEIGHT_CM: 21.0,
    PAGE_WIDTH_CM: 29.7
  },
ESTILOS: {
FUENTE_BASE: 'Arial',
    TAMANO_TEXTO_NORMAL: 10,
    TAMANO_TEXTO_PEQUEÑO: 10,
    TAMANO_TEXTO_HEADER: 10,
    TAMANO_TEXTO_FOOTER: 7.5,
    TAMANO_ACELERANDO: 7.5,
    TAMANO_DIRECCION: 6.8
  },
  COLORES: {
    AMARILLO: '#fbb818',
    GRIS_OSCURO: '#525352',
    GRIS_DIRECCION: '#515151'
  },
  LOGO: {
  URL: 'https://cualli.mx/wp-content/uploads/2022/07/cualli-bl@3x.png',
  WIDTH: 75, 
  HEIGHT: 75
},
};

function leerDatos() {
  const ss = SpreadsheetApp.getActive();
   const sh = HOJA_ACTIVA_CONFIG.obtenerHojaDatos();
  if (!sh) throw new Error(`No se encuentra la hoja: ${CONFIG.SHEET_NAME}`);

  const criticalCells = [CONFIG.CELDAS.ID_OPORT, CONFIG.CELDAS.ACREDITADO, CONFIG.CELDAS.REGIMEN];
  criticalCells.forEach(cell => {
    if (!sh.getRange(cell).getDisplayValue().trim()) {
      throw new Error(`Celda importante vacía, debes colocar un ID en: ${cell}`);
    }
  });

  const tz  = Session.getScriptTimeZone() || 'America/Mexico_City';
  const hoy = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');

  const get = (a1) => sh.getRange(a1).getDisplayValue();

  const regimen = get(CONFIG.CELDAS.REGIMEN) || '';
  const sufijo  = regimen.toLowerCase().includes('moral') ? 'PM' : 'PF';
  const footer  = `C_Cuadro_de_Transferencias_${sufijo}_V1.0`;

  const meta = {
    id:         get(CONFIG.CELDAS.ID_OPORT),
    cliente:    get(CONFIG.CELDAS.CLIENTE) || '',
    fechaHoy:   hoy,
    acreditado: get(CONFIG.CELDAS.ACREDITADO),
    regimen:    regimen,
    footer:     footer,
    rep1:       get(CONFIG.CELDAS.REP1) || '',  
    rep2:       get(CONFIG.CELDAS.REP2) || '',
    moneda:     get(CONFIG.CELDAS.MONEDA) || 'MXN'    
  };

  const lastRow = sh.getLastRow();
  const datos   = [];

  if (lastRow >= CONFIG.START_ROW) {
    const rango   = sh.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 10);
    const valores = rango.getDisplayValues();

    for (let i = 0; i < valores.length; i++) {
      const fila = valores[i];
      if (fila[0] || fila[2] || fila[3] || fila[5]) {
        datos.push({
          num:          fila[CONFIG.COLUMNAS.NUM - 1]      || '',
          fecha:        fila[CONFIG.COLUMNAS.FECHA - 1]    || '',
          concepto:     fila[CONFIG.COLUMNAS.CONCEPTO - 1] || '',
          monto:        fila[CONFIG.COLUMNAS.MONTO - 1]    || '',
          comentarios:  fila[CONFIG.COLUMNAS.COMENT - 1]   || '',
          banco:        fila[CONFIG.COLUMNAS.BANCO - 1]    || '',
          clabe:        fila[CONFIG.COLUMNAS.CLABE - 1]    || '',
          beneficiario: fila[CONFIG.COLUMNAS.BENEFI - 1]   || ''
        });
      }
    }
  }

  if (datos.length === 0) {
    throw new Error(`No se encontraron datos desde la fila ${CONFIG.START_ROW} para generar el cuadro, deberas capturar al menos una solicitud de dispersion`);
  }

  const info = determinarTipoPersonaYNombre(meta);
  validarDatosFirmaPersonaMoral(meta, info);

  return { meta, datos };
}

function cmToPt(cm) { return cm * 28.3464567; }

function limpiarTexto_(texto) {
  if (!texto) return '';
  return texto.toString()
    .replace(/\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function formatMontoMillonesMX_(valor, moneda = 'MXN') {
  if (valor == null || valor === '') return '';
  let s = String(valor).trim();
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  s = s
    .replace(/\$/g, '')
    .replace(/USD/ig, '') 
    .replace(/\s+/g, '')
    .replace(/,/g, '')
    .replace(/\u2019/g, '')  
    .replace(/'/g, '');    

  let num = Number(s);
  if (!isFinite(num)) return String(valor);
  if (neg) num = -Math.abs(num);

  const abs  = Math.abs(num);
  
  // Condicional para decidir el símbolo según lo que diga la celda
  const esUSD = (moneda.toString().trim().toUpperCase() === 'USD');
  const prefijo = esUSD ? 'USD ' : '$';
  
  const base = prefijo + abs.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const MILLON = 1000000;
  
  // Opcional: solo mantener el apóstrofe para millones si no es USD
  const conApo = (abs >= MILLON && !esUSD) ? base.replace(',', '\u02BC') : base;
  return (num < 0 ? '-' : '') + conApo;
}


function _estimarAlturaTabla_(d) {
  const chars = ('' + (d.concepto||'')).length +
                ('' + (d.comentarios||'')).length +
                ('' + (d.beneficiario||'')).length;
  return CONFIG.ALTURA.TABLA_BASE + Math.ceil(chars / 50) * CONFIG.ALTURA.EXTRA_CADA_50_CH;
}

function _saltosInteligentesAntesDeTabla_(body, alturaAcumulada, alturaTablaEstimada) {
  const espacioRestante = CONFIG.ALTURA.UTIL_PAG - alturaAcumulada;

  if (espacioRestante < (alturaTablaEstimada + CONFIG.MARGEN_SEGURIDAD)) {
    body.appendPageBreak();
    return 0;
  }

  if (espacioRestante < CONFIG.ALTURA.MIN_INICIO) {
    body.appendPageBreak();
    return 0;
  }

  return alturaAcumulada;
}

function _saltosAntesDeAutorizacion_(body, alturaAcumulada, alturaBloque) {
  const espacioRestante = CONFIG.ALTURA.UTIL_PAG - alturaAcumulada;
  if (espacioRestante < alturaBloque) {
    body.appendPageBreak();
    return 0;
  }
  return alturaAcumulada;
}

function crearDocumento_(nombre) {
  const folder = CONFIG.OUTPUT_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID)
    : DriveApp.getRootFolder();

  if (CONFIG.TEMPLATE_DOC_ID && CONFIG.TEMPLATE_DOC_ID.trim()) {
    const plantilla = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
    return plantilla.makeCopy(nombre, folder);
  }

  const doc     = DocumentApp.create(nombre);
  const archivo = DriveApp.getFileById(doc.getId());
  folder.addFile(archivo);
  DriveApp.getRootFolder().removeFile(archivo);
  return archivo;
}

function configurarDocumento_(docId) {
  const doc  = DocumentApp.openById(docId);
  const body = doc.getBody();
  body.setPageHeight(cmToPt(CONFIG.DOCUMENTO.PAGE_HEIGHT_CM));
  body.setPageWidth(cmToPt(CONFIG.DOCUMENTO.PAGE_WIDTH_CM));
  doc.saveAndClose();
}

function setSectionMarginsCm(docId, { top, right, bottom, left }) {
  const docJson = Docs.Documents.get(docId);
  const content = (docJson.body && docJson.body.content) || [];
  const startIndex = 1;
  const endIndex   = content.length ? content[content.length - 1].endIndex : 1;

  Docs.Documents.batchUpdate({
    requests: [{
      updateSectionStyle: {
        range: { segmentId: null, startIndex, endIndex },
        sectionStyle: {
          marginTop:    { magnitude: cmToPt(top),    unit: 'PT' },
          marginRight:  { magnitude: cmToPt(right),  unit: 'PT' },
          marginBottom: { magnitude: cmToPt(bottom), unit: 'PT' },
          marginLeft:   { magnitude: cmToPt(left),   unit: 'PT' }
        },
        fields: 'marginTop,marginRight,marginBottom,marginLeft'
      }
    }]
  }, docId);
}
function setHeaderFooterMarginsCm_(docId, headerCm, footerCm) {
  Docs.Documents.batchUpdate({
    requests: [{
      updateDocumentStyle: {
        documentStyle: {
          marginHeader: { magnitude: cmToPt(headerCm), unit: 'PT' },
          marginFooter: { magnitude: cmToPt(footerCm), unit: 'PT' }
        },
        fields: 'marginHeader,marginFooter'
      }
    }]
  }, docId);
}

function configurarHeaderFooter_(docId, meta) {
  const doc = DocumentApp.openById(docId);
  
  // --- ENCABEZADO ---
  let header = doc.getHeader();
  if(!header) header = doc.addHeader();
  header.clear();

  const tableH = header.appendTable();
  tableH.setBorderWidth(0);              
  const rowH = tableH.appendTableRow();

  const c0 = rowH.appendTableCell();
  c0.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
  
  try {
const blob = UrlFetchApp.fetch(CONFIG.LOGO.URL, { followRedirects: true }).getBlob();
    const img  = c0.appendImage(blob);
    img.setWidth(CONFIG.LOGO.WIDTH);
    img.setHeight(CONFIG.LOGO.HEIGHT);

   
    const p0 = c0.getChild(0).asParagraph();
    p0.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.0)
      .setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  } catch (e) {
    c0.setText('');
  }


  const c1 = rowH.appendTableCell();
  c1.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
  const pId = c1.appendParagraph(`ID Operación: ${meta.id || ''}`);
  pId.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  pId.editAsText().setBold(true).setFontSize(10);

  const pCli = c1.appendParagraph(`Cliente: ${meta.cliente || ''}`);
  pCli.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  pCli.editAsText().setBold(true).setFontSize(10);

  const pFe = c1.appendParagraph(`Fecha: ${meta.fechaHoy}`);
  pFe.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  pFe.editAsText().setBold(true).setFontSize(10);

  // --- PIE DE PÁGINA ---
  let footer = doc.getFooter();
  if(!footer) footer = doc.addFooter();
  footer.clear();

  const tableF = footer.appendTable();
  tableF.setBorderWidth(0);
  const rowF = tableF.appendTableRow();
  const cLeft = rowF.appendTableCell();
  
  const p1 = cLeft.appendParagraph("acelerandoportunidades");
  const t1 = p1.editAsText();
  t1.setFontSize(CONFIG.ESTILOS.TAMANO_ACELERANDO).setFontFamily('Arial');
  t1.setForegroundColor(0, 8, CONFIG.COLORES.AMARILLO);
  t1.setForegroundColor(9, 21, CONFIG.COLORES.GRIS_OSCURO);

  const dirText = "Ciudad de México. Torre del Árbol Blvd. Manuel Ávila Camacho 184, Piso 7 Colonia Reforma Social.CP 11650.+(52) 55 8117 1700.";
  const p2 = cLeft.appendParagraph(dirText);
  p2.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.0);
  const t2 = p2.editAsText();
  t2.setFontSize(CONFIG.ESTILOS.TAMANO_DIRECCION).setFontFamily('Arial');
  t2.setForegroundColor(CONFIG.COLORES.GRIS_DIRECCION);
  t2.setForegroundColor(0, 16, CONFIG.COLORES.AMARILLO);

  const cRight = rowF.appendTableCell();
  const p3 = cRight.appendParagraph(meta.footer);
  p3.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  p3.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.0);
  p3.editAsText().setFontFamily('Courier New').setFontSize(7.5).setForegroundColor('#666666');

  tableF.setColumnWidth(0, 460); 
  tableF.setColumnWidth(1, 190);


  doc.saveAndClose(); 

}


function _getLastTableStartIndex_(docId) {
  const docJson = Docs.Documents.get(docId);
  const content = (docJson.body && docJson.body.content) || [];
  for (let i = content.length - 1; i >= 0; i--) {
    if (content[i].table) return content[i].startIndex;
  }
  return null;
}

function _applyColumnWidthsToLastTable_(docId, widthsPtsByColIndex) {
  const startIndex = _getLastTableStartIndex_(docId);
  if (startIndex == null) return;

  const requests = Object.keys(widthsPtsByColIndex).map(k => {
    const colIndex = Number(k);
    const widthPts = widthsPtsByColIndex[k];
    return {
      updateTableColumnProperties: {
        tableStartLocation: { index: startIndex },
        columnIndices: [colIndex],
        tableColumnProperties: {
          width: { magnitude: widthPts, unit: 'PT' },
          widthType: 'FIXED_WIDTH'
        },
        fields: 'width,widthType'
      }
    };
  });

  Docs.Documents.batchUpdate({ requests }, docId);
}

function _mergeCellsInLastTable_(docId, rowIndex, colIndex, rowSpan, colSpan) {
  const startIndex = _getLastTableStartIndex_(docId);
  if (startIndex == null) return;

  Docs.Documents.batchUpdate({
    requests: [{
      mergeTableCells: {
        tableRange: {
          tableCellLocation: {
            tableStartLocation: { index: startIndex },
            rowIndex: rowIndex,
            columnIndex: colIndex
          },
          rowSpan: rowSpan,
          columnSpan: colSpan
        }
      }
    }]
  }, docId);
}

function _listTableStartIndices_(docId) {
  const docJson = Docs.Documents.get(docId);
  const content = (docJson.body && docJson.body.content) || [];
  const starts = [];
  for (let i = 0; i < content.length; i++) {
    if (content[i].table) starts.push(content[i].startIndex);
  }
  return starts;
}

function _compactarEspaciosEntreTablas_(docId) {
  const doc  = DocumentApp.openById(docId);
  const body = doc.getBody();
  const n = body.getNumChildren();

  for (let i = 0; i < n - 1; i++) {
    const a = body.getChild(i);
    const b = body.getChild(i + 1);

    if (a.getType() === DocumentApp.ElementType.TABLE &&
        b.getType() === DocumentApp.ElementType.PARAGRAPH) {

      const p = b.asParagraph();
      const txt = p.getText ? p.getText() : '';

      p.setSpacingBefore(0);
      p.setSpacingAfter(0);
      p.setLineSpacing(1.0);

      const t = p.editAsText();
      if (t.getText().length === 0) t.appendText('\u00A0');
      t.setFontSize(0, t.getText().length - 1, 1);
    }
  }

  for (let i = body.getNumChildren() - 2; i >= 0; i--) {
    const cur = body.getChild(i);
    const nxt = body.getChild(i + 1);
    if (cur.getType() === DocumentApp.ElementType.PARAGRAPH &&
        nxt.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const tc = cur.asParagraph().getText().trim();
      const tn = nxt.asParagraph().getText().trim();
      if (!tc && !tn) {
        body.removeChild(nxt);
      }
    }
  }

  doc.saveAndClose();
}

function _aplicarEstiloBaseCelda_(celda, alineacion = DocumentApp.HorizontalAlignment.CENTER, fontSize = 10) {
  celda.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);

  celda.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
  celda.setPaddingTop(0.1 * 28.3464567)  
       .setPaddingBottom(0.1 * 28.3464567)
       .setPaddingLeft(0.1 * 28.3464567)
       .setPaddingRight(0.1 * 28.3464567);  

  const n = celda.getNumChildren();
  for (let i = 0; i < n; i++) {
    const el = celda.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      p.setAlignment(alineacion);
      p.setLineSpacing(1.10);
      p.setSpacingBefore(0);
      p.setSpacingAfter(0);
      p.setIndentFirstLine(0);
      p.setIndentStart(0);
      p.setIndentEnd(0);

      if (p.editAsText) {
        const t = p.editAsText();
        t.setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);
        t.setFontSize(fontSize);
        t.setBold(false);
      }
    }
  }
}

function configurarCeldaNormal_(celda, alineacion) {
   celda.setPaddingTop(0.1 * 28.3464567)
       .setPaddingBottom(0.1 * 28.3464567)
       .setPaddingLeft(0.1 * 28.3464567)
       .setPaddingRight(0.1 * 28.3464567);
  _aplicarEstiloBaseCelda_(celda, alineacion, CONFIG.ESTILOS.TAMANO_TEXTO_NORMAL);
}

function configurarCeldaCentrada_(celda) {
   celda.setPaddingTop(0.1 * 28.3464567)
       .setPaddingBottom(0.1 * 28.3464567)
       .setPaddingLeft(0.1 * 28.3464567)
       .setPaddingRight(0.1 * 28.3464567);
  _aplicarEstiloBaseCelda_(celda, DocumentApp.HorizontalAlignment.CENTER, CONFIG.ESTILOS.TAMANO_TEXTO_PEQUEÑO);
}

function configurarCeldaVacia_(celda) {
  celda.setBackgroundColor('#ffffff');
  celda.setPaddingTop(0.1 * 28.3464567)
       .setPaddingBottom(0.1 * 28.3464567)
       .setPaddingLeft(0.1 * 28.3464567)
       .setPaddingRight(0.1 * 28.3464567);
  _aplicarEstiloBaseCelda_(celda);
}

function crearCeldaEncabezado_(fila, texto) {
  const celda = fila.appendTableCell(texto);
  celda.setBackgroundColor(CONFIG.COLOR_HEADER);
  celda.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
  celda.setPaddingTop(0.1 * 28.3464567)  
       .setPaddingBottom(0.1 * 28.3464567)
       .setPaddingLeft(0.1 * 28.3464567)
       .setPaddingRight(0.1 * 28.3464567);

  const parrafo = celda.getChild(0).asParagraph();
  parrafo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  parrafo.setLineSpacing(1.10);
  parrafo.setSpacingBefore(0);
  parrafo.setSpacingAfter(0);
  parrafo.setIndentFirstLine(0);
  parrafo.setIndentStart(0);
  parrafo.setIndentEnd(0);

  const t = celda.editAsText();
  t.setBold(true);
  t.setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);
  t.setFontSize(CONFIG.ESTILOS.TAMANO_TEXTO_HEADER);

  return celda;
}

function crearCeldaSubencabezado_(fila, texto) {
  const celda = fila.appendTableCell(texto);
  celda.setBackgroundColor(CONFIG.COLOR_SUBHDR);
  celda.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
  celda.setPaddingTop(0.1 * 28.3464567)   
       .setPaddingBottom(0.1 * 28.3464567)
       .setPaddingLeft(0.1 * 28.3464567)
       .setPaddingRight(0.1 * 28.3464567);

  const parrafo = celda.getChild(0).asParagraph();
  parrafo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  parrafo.setLineSpacing(1.10);
  parrafo.setSpacingBefore(0);
  parrafo.setSpacingAfter(0);
  parrafo.setIndentFirstLine(0);
  parrafo.setIndentStart(0);
  parrafo.setIndentEnd(0);

  const t = celda.editAsText();
  t.setBold(true);
  t.setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);
  t.setFontSize(CONFIG.ESTILOS.TAMANO_TEXTO_HEADER);

  return celda;
}

function etiquetaClabe_(valor) {
  const s = (valor || '').toString();
  if (!s.trim()) return 'CLABE';

  if (/\bRETENID[OA]S?\b/i.test(s)) return 'CLABE';

  const soloDigitos = s.replace(/[\s\-.]/g, '').match(/\d+/g)?.join('') || '';

  if (/\d{19,}/.test(soloDigitos)) return 'LÍNEA DE CAPTURA';

  const tieneNumero = /\d/.test(s);
  const tieneLetra  = /[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]/.test(s);
  if (tieneNumero && tieneLetra) return 'LÍNEA DE CAPTURA';

  return 'CLABE';
}

function insertarContenidoPrincipal_(docId, datos, meta) {
  let doc  = DocumentApp.openById(docId);
  let body = doc.getBody();

  body.appendParagraph('').setSpacingAfter(1);

  let alturaAcumulada = 1;

  for (let i = 0; i < datos.length; i++) {
    const d = datos[i];

    const alturaEstimada = _estimarAlturaTabla_(d);
    alturaAcumulada = _saltosInteligentesAntesDeTabla_(body, alturaAcumulada, alturaEstimada);

    const tabla = body.appendTable();

    const filaEnc = tabla.appendTableRow();
    crearCeldaEncabezado_(filaEnc, 'No. de Solicitud');
    crearCeldaEncabezado_(filaEnc, 'Fecha de Solicitud');
    crearCeldaEncabezado_(filaEnc, 'Concepto');
    crearCeldaEncabezado_(filaEnc, 'Monto');
    crearCeldaEncabezado_(filaEnc, 'Comentarios');

    const filaDatos = tabla.appendTableRow();
    const cNum = filaDatos.appendTableCell(d.num || '');
    configurarCeldaCentrada_(cNum);
    const tNum = cNum.editAsText();
    if (tNum) tNum.setBold(true);
    configurarCeldaCentrada_(filaDatos.appendTableCell(d.fecha || ''));
    configurarCeldaNormal_(filaDatos.appendTableCell(limpiarTexto_(d.concepto) || ''), DocumentApp.HorizontalAlignment.CENTER);
    const montoFmt = formatMontoMillonesMX_(d.monto, meta.moneda);
    const cMonto = filaDatos.appendTableCell('');
    configurarCeldaNormal_(cMonto, DocumentApp.HorizontalAlignment.CENTER);
    cMonto.editAsText().setText(montoFmt);                
    configurarCeldaNormal_(filaDatos.appendTableCell(limpiarTexto_(d.comentarios) || ''), DocumentApp.HorizontalAlignment.CENTER);

    const filaSub = tabla.appendTableRow();
    configurarCeldaVacia_(filaSub.appendTableCell(''));
    configurarCeldaVacia_(filaSub.appendTableCell(''));
    crearCeldaSubencabezado_(filaSub, 'Banco');
    crearCeldaSubencabezado_(filaSub, etiquetaClabe_(d.clabe));
    crearCeldaSubencabezado_(filaSub, 'BENEFICIARIO');

    const filaBco = tabla.appendTableRow();
    configurarCeldaVacia_(filaBco.appendTableCell(''));
    configurarCeldaVacia_(filaBco.appendTableCell(''));
    configurarCeldaCentrada_(filaBco.appendTableCell(limpiarTexto_(d.banco) || ''));
    configurarCeldaCentrada_(filaBco.appendTableCell(limpiarTexto_(d.clabe) || ''));
    configurarCeldaCentrada_(filaBco.appendTableCell(limpiarTexto_(d.beneficiario) || ''));

    alturaAcumulada += alturaEstimada;
  }

  doc.saveAndClose();

  const starts = _listTableStartIndices_(docId);
  if (starts.length) {
    const req = [];
    for (const s of starts) {
      req.push(
        { mergeTableCells: { tableRange: { tableCellLocation: { tableStartLocation: { index: s }, rowIndex: 1, columnIndex: 0 }, rowSpan: 3, columnSpan: 1 } } },
        { mergeTableCells: { tableRange: { tableCellLocation: { tableStartLocation: { index: s }, rowIndex: 1, columnIndex: 1 }, rowSpan: 3, columnSpan: 1 } } },
        { updateTableCellStyle: {
            tableRange: { tableCellLocation: { tableStartLocation: { index: s }, rowIndex: 0, columnIndex: 0 }, rowSpan: 4, columnSpan: 5 },
            tableCellStyle: { contentAlignment: 'MIDDLE' },
            fields: 'contentAlignment'
        } },
        { updateTableColumnProperties: {
            tableStartLocation: { index: s }, columnIndices: [0], //No
            tableColumnProperties: { width: { magnitude: 66.5, unit: 'PT' }, widthType: 'FIXED_WIDTH' },
            fields: 'width,widthType'
        } },
        { updateTableColumnProperties: {
            tableStartLocation: { index: s }, columnIndices: [1], //Fecha
            tableColumnProperties: { width: { magnitude: 80, unit: 'PT' }, widthType: 'FIXED_WIDTH' },
            fields: 'width,widthType'
        } },
        { updateTableColumnProperties: {
        tableStartLocation: { index: s }, columnIndices: [2], // Concepto
        tableColumnProperties: { width: { magnitude: cmToPt(5.5), unit: 'PT' }, widthType: 'FIXED_WIDTH' },
        fields: 'width,widthType'
        } },
        { updateTableColumnProperties: {
        tableStartLocation: { index: s }, columnIndices: [3], // Monto
        tableColumnProperties: { width: { magnitude: cmToPt(5.0), unit: 'PT' }, widthType: 'FIXED_WIDTH' },
        fields: 'width,widthType'
        } },
        { updateTableColumnProperties: {
            tableStartLocation: { index: s }, columnIndices: [4], //Comentarios
            tableColumnProperties: { width: { magnitude: cmToPt(7.541), unit: 'PT' }, widthType: 'FIXED_WIDTH' },
            fields: 'width,widthType'
        } }
      );
    }
    Docs.Documents.batchUpdate({ requests: req }, docId);
  }

  _compactarEspaciosEntreTablas_(docId);

  doc  = DocumentApp.openById(docId);
  body = doc.getBody();
  alturaAcumulada = _saltosAntesDeAutorizacion_(body, 0, CONFIG.ALTURA.AUTORIZACION_APROX);
  insertarAutorizacionContinuo_(body, meta);
  doc.saveAndClose();
}

function determinarTipoPersonaYNombre(meta) {
  const regimen = (meta.regimen || '').toLowerCase();
  const esPersonaMoral = regimen.includes('moral');

  const nombreAcreditado = meta.acreditado || '';

  const rep1Raw = (meta.rep1 || '').trim();
  const rep2Raw = (meta.rep2 || '').trim();
  
  const rep1 = (rep1Raw && rep1Raw.toUpperCase() !== 'N/A') ? rep1Raw : '';
  const rep2 = (rep2Raw && rep2Raw.toUpperCase() !== 'N/A') ? rep2Raw : '';

  return {
    esPersonaMoral,
    nombreAcreditado,
    nombreRepresentante1: rep1,
    nombreRepresentante2: rep2
  };
}

function validarDatosFirmaPersonaMoral(meta, info) {
  if (info.esPersonaMoral) {
    // Obtener valores limpios sin N/A
    const rep1Raw = (meta.rep1 || '').trim();
    const rep2Raw = (meta.rep2 || '').trim();
    
    const rep1 = (rep1Raw && rep1Raw.toUpperCase() !== 'N/A') ? rep1Raw : '';
    const rep2 = (rep2Raw && rep2Raw.toUpperCase() !== 'N/A') ? rep2Raw : '';
    
    if (!rep1 && !rep2) {
      throw new Error(
        '❌ PERSONA MORAL SIN REPRESENTANTES\n\n' +
        'Para generar el cuadro de transferencias se requiere al menos un representante legal.\n\n' +
        '📝 Favor de capturar al menos un nombre en:\n' +
        '• Celda B23 (Representante 1) - OBLIGATORIO\n' +  
        '• Celda B25 (Representante 2) - OPCIONAL\n\n' +
        '⚠️ Nota: Los valores "N/A" se consideran como vacíos.\n'
      );
    }
   }
  
  return true;
}

function extraerNombreDeID(id) {
  if (!id) return 'Representante Legal';
  const ultimoGuion = id.lastIndexOf('-');
  if (ultimoGuion !== -1 && ultimoGuion < id.length - 1) {
    const nombre = id.substring(ultimoGuion + 1).trim();
    if (nombre.includes(' ') || nombre.split(' ').length >= 2) return nombre;
  }
  return 'Representante Legal';
}

function inferirTrato_(nombre) {
  const nombreLimpio = (nombre || '').trim();
  if (!nombreLimpio) return 'Sr.';

  const lower = nombreLimpio.toLowerCase();

 const femeninos = [
  'maría','maria', 'mari', 'carmen', 'maria del carmen', 'carina','luz','isabel','ana','elena','patricia','gabriela',
  'verónica','veronica','andrea','elizabeth','elizabet','lizbeth','lizeth',
  'lizabeth','diana','claudia','rosa','teresa','beatriz','alejandra',
  'guadalupe','josefina','francisca','concepción','adriana','alicia','amalia',
  'amelia','angélica','angelica','antonia','araceli','aurora','bárbara','barbara',
  'bertha','blanca','carolina','catalina','cecilia','celia','christian','cindy',
  'citlali','clara','cristina', 'danna', 'daniela','débora','debora','delia','dolores',
  'dorotea','edith','elisa','elsa','emilia','enriqueta','esperanza','estela',
  'eugenia','eva','evelyn', 'evelin','fabiola','fatima','flor','florencia','frida',
  'gema','gloria','graciela','griselda','helena','hilda','ilse','ines','irene',
  'irma','isabela','ivette','jacquelin','jessica','jimena','joana','jocelyn',
  'johana','josefa','juana','julia','karen', 'karelia','karina','karla','katherine','katia',
  'kimberly','laura','leticia','lidia','liliana','linda','lorena','lourdes',
  'lucia','luisa','lupita','lydia','magdalena','marcela','margarita','mariana',
  'maricela', 'marvelys','marisol','marta','martina','mayra','melissa','mercedes','michel',
  'miriam','monica','monique','naomi','natalia','nayeli','nelly','nicole',
  'noemi','norma','olga','olivia','paola','paulina','pilar','ramona','raquel',
  'rebeca','regina','renata','rita','rocio','romina', 'roberta', 'rosario','rosaura',
  'rubi','ruth','sabrina','salma','samanta','sandra','sara','silvia','simone',
  'soledad','sonia','sophia','susana','tania','tatiana','valentina','valeria',
  'vanessa','vera','victoria','virginia','vivian','wendy','ximena','yadira',
  'yanet','yaquelin','yazmin', 'jazmin','yesenia','yolanda','zoe','zulema'
];

const masculinos = [
  'josé','jose','carlos','juan','pedro','luis','miguel','antonio','francisco',
  'david','alejandro','manuel','javier','rafael','ricardo','eduardo','alfonso',
  'ernesto','fernando','rodrigo','sergio','abraham','adrian','alberto','alex',
  'alexander','alfredo','alonso','andrés','andres','angel','arturo','augusto',
  'benito','benjamín','benjamin','bernardo','brandon','cesar','cristian',
  'damian','daniel','dario','diego','eduardo','elias','emilio', 'emiliano','enrique',
  'erick','esteban','federico','felix','felipe','fernando','francisco',
  'gabriel','gerardo','gonzalo','gregorio','guadalupe','guillermo','gustavo',
  'hector','hugo','ignacio','isaac','ismael','ivan','jacobo', 'jair', 'jaime','javier',
  'jerónimo','jeromino','jesus','joaquin','jorge','jose','juan','julian',
  'julio','kevin', 'lamberto', 'leonardo','leonel','leonardo','lorenzo','lucas','marcelo',
  'marco', 'marcos', 'mario','martin','mateo','matias','mauricio','maximiliano','michael',
  'moises','narciso','nestor','nicolas','octavio','omar','orlando','oscar',
  'pablo','patricio','ramiro','ramon','raul','renato','rene','roberto',
  'rodolfo','rolando','roman','ruben','salvador','samuel','san','santiago',
  'saul','sebastian','simon','teodoro','timoteo','tomas','ulises','victor',
  'vicente','walter','wilson','xavier','yahir'
];

  const primera = lower.split(/\s+/)[0];
  if (femeninos.includes(primera))  return 'Sra.';
  if (masculinos.includes(primera)) return 'Sr.';
  return (/[aá]$/i.test(primera)) ? 'Sra.' : 'Sr.';
}

function formatearNombreEmpresa_(nombre) {
  if (!nombre) return '';

  let nombreFormateado = nombre.toString().toUpperCase().trim();

  nombreFormateado = nombreFormateado.replace(/\s{2,}/g, ' ');

  nombreFormateado = nombreFormateado
    // SAPI DE CV SOFOM con puntos en todas las abreviaturas
    .replace(/\bS\.?A\.?\s*P\.?I\.?\s*DE\s*C\.?V\.?\s*S\.?O\.?F\.?O\.?M\s*E\.?N\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA PROMOTORA DE INVERSIÓN DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA')
    .replace(/\bS\.?A\.?\s*P\.?I\.?\s*DE\s*C\.?V\.?\s*S\.?O\.?F\.?O\.?M\s*E\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA PROMOTORA DE INVERSIÓN DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA')
    
    // SA DE CV SOFOM con puntos en todas las abreviaturas
    .replace(/\bS\.?A\.?\s*DE\s*C\.?V\.?\s*S\.?O\.?F\.?O\.?M\s*E\.?N\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA')
    .replace(/\bS\.?A\.?\s*DE\s*C\.?V\.?\s*S\.?O\.?F\.?O\.?M\s*E\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA')
    
    // SAB DE CV SOFOM con puntos en todas las abreviaturas
    .replace(/\bS\.?A\.?\s*B\.?\s*DE\s*C\.?V\.?\s*S\.?O\.?F\.?O\.?M\s*E\.?N\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA BURSATIL DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA')
    .replace(/\bS\.?A\.?\s*B\.?\s*DE\s*C\.?V\.?\s*S\.?O\.?F\.?O\.?M\s*E\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA BURSATIL DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA')
    
    // S DE RL DE CV SOFOM con puntos en todas las abreviaturas
    .replace(/\bS\.?\s*DE\s*R\.?L\.?\s*DE\s*C\.?V\.?\s*S\.?O\.?F\.?O\.?M\s*E\.?R\.?\b/gi, 'SOCIEDAD DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA')
    .replace(/\bS\.?\s*DE\s*R\.?L\.?\s*DE\s*C\.?V\.?\s*S\.?O\.?F\.?O\.?M\s*E\.?N\.?R\.?\b/gi, 'SOCIEDAD DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA')
    
    // También mantener las versiones sin puntos (por compatibilidad)
    .replace(/\bS\.?A\.?\s*P\.?I\.?\s*DE\s*C\.?V\.?\s*SOFOM\s*E\.?N\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA PROMOTORA DE INVERSIÓN DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA')
    .replace(/\bS\.?A\.?\s*P\.?I\.?\s*DE\s*C\.?V\.?\s*SOFOM\s*E\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA PROMOTORA DE INVERSIÓN DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA')
    .replace(/\bS\.?A\.?\s*DE\s*C\.?V\.?\s*SOFOM\s*E\.?N\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA')
    .replace(/\bS\.?A\.?\s*DE\s*C\.?V\.?\s*SOFOM\s*E\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA')
    .replace(/\bS\.?A\.?\s*B\.?\s*DE\s*C\.?V\.?\s*SOFOM\s*E\.?N\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA BURSATIL DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA')
    .replace(/\bS\.?A\.?\s*B\.?\s*DE\s*C\.?V\.?\s*SOFOM\s*E\.?R\.?\b/gi, 'SOCIEDAD ANÓNIMA BURSATIL DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA')
    .replace(/\bS\.?\s*DE\s*R\.?L\.?\s*DE\s*C\.?V\.?\s*SOFOM\s*E\.?R\.?\b/gi, 'SOCIEDAD DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA')
    .replace(/\bS\.?\s*DE\s*R\.?L\.?\s*DE\s*C\.?V\.?\s*SOFOM\s*E\.?N\.?R\.?\b/gi, 'SOCIEDAD DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA')
    
    // Tipos societarios básicos sin SOFOM - con puntos opcionales en todas partes
    .replace(/\bS\.?A\.?\s*B\.?\s*DE\s*C\.?V\.?\b/gi, 'SOCIEDAD ANÓNIMA BURSATIL DE CAPITAL VARIABLE')
    .replace(/\bS\.?A\.?\s*P\.?I\.?\s*DE\s*C\.?V\.?\b/gi, 'SOCIEDAD ANÓNIMA PROMOTORA DE INVERSIÓN DE CAPITAL VARIABLE')
    .replace(/\bS\.?A\.?\s*DE\s*C\.?V\.?\b/gi, 'SOCIEDAD ANÓNIMA DE CAPITAL VARIABLE')
    .replace(/\bS\.?\s*DE\s*R\.?L\.?\s*DE\s*C\.?V\.?\b/gi, 'SOCIEDAD DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE')
    .replace(/\bS\.?\s*DE\s*R\.?L\.?\b/gi, 'SOCIEDAD DE RESPONSABILIDAD LIMITADA')
    .replace(/\bA\.?C\.?\b/gi, 'ASOCIACIÓN CIVIL')
    .replace(/\bS\.?A\.?\b/gi, 'SOCIEDAD ANÓNIMA')
    .replace(/\bS\.?C\.?\b/gi, 'SOCIEDAD CIVIL')
    .replace(/\bR\.?L\.?\b/gi, 'RESPONSABILIDAD LIMITADA')
    .replace(/\bC\.?V\.?\b/gi, 'CAPITAL VARIABLE');

  const tiposSocietarios = [
    // SOFOM ENR
    'SOCIEDAD ANÓNIMA PROMOTORA DE INVERSIÓN DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA',
    'SOCIEDAD ANÓNIMA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA',
    'SOCIEDAD ANÓNIMA BURSATIL DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA',
    'SOCIEDAD DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA',
    
    // SOFOM ER
    'SOCIEDAD ANÓNIMA PROMOTORA DE INVERSIÓN DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA',
    'SOCIEDAD ANÓNIMA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA',
    'SOCIEDAD ANÓNIMA BURSATIL DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA',
    'SOCIEDAD DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD REGULADA',
    
    // Tipos básicos
    'SOCIEDAD ANÓNIMA BURSATIL DE CAPITAL VARIABLE',
    'SOCIEDAD ANÓNIMA PROMOTORA DE INVERSIÓN DE CAPITAL VARIABLE',
    'SOCIEDAD ANÓNIMA DE CAPITAL VARIABLE',
    'SOCIEDAD DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE',
    'SOCIEDAD DE RESPONSABILIDAD LIMITADA',
    'ASOCIACIÓN CIVIL',
    'SOCIEDAD ANÓNIMA',
    'SOCIEDAD CIVIL'
  ];

  let nombreComercial = nombreFormateado;
  let tipoSocietario = '';

  for (const tipo of tiposSocietarios) {
    if (nombreFormateado.includes(tipo)) {
      tipoSocietario = tipo;
      nombreComercial = nombreFormateado.replace(tipo, '').trim();
      break;
    }
  }

  nombreComercial = nombreComercial
    .replace(/,\s*\./g, '')
    .replace(/^\s*[,\.]+\s*/g, '')
    .replace(/\s*[,\.]+\s*$/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();

  if (!tipoSocietario) {
    return `"${nombreComercial || nombreFormateado.replace(/\s*[,\.]+\s*$/g, '').trim()}",`;
  }

  return `"${nombreComercial}", ${tipoSocietario},`;
}

function inferirArticulo_(trato) {
  return trato === 'Sra.' ? 'la' : 'el';
}

function formatearNombrePersona_(nombre) {
  if (!nombre) return '';
  return nombre.toString().toUpperCase();
}

function insertarAutorizacionContinuo_(body, meta) {
  const info = determinarTipoPersonaYNombre(meta);
  const acreditadoFormateado = formatearNombreEmpresa_(info.nombreAcreditado);
  const acreditadoPersonaFisica = formatearNombrePersona_(meta.acreditado);

  // Usar los valores ya filtrados de info (sin N/A)
  const rep1Nombre = formatearNombrePersona_(info.nombreRepresentante1 || '');
  const rep2Nombre = formatearNombrePersona_(info.nombreRepresentante2 || '');

  body.appendParagraph('').setSpacingAfter(1);

  let textoAutorizacion = '';
  
  if (info.esPersonaMoral) {
    const t1 = inferirTrato_(rep1Nombre);
    const a1 = inferirArticulo_(t1);
    let reps = `${a1} ${t1} ${rep1Nombre}`;
    if (rep2Nombre) {
      const t2 = inferirTrato_(rep2Nombre);
      const a2 = inferirArticulo_(t2);
      reps += ` y ${a2} ${t2} ${rep2Nombre}`;
    }

    textoAutorizacion =
      `${acreditadoFormateado} en su carácter de "El Acreditado" y/o "El Cliente", ` +
      `representado por ${reps}, con número de cliente indicado ` +
      `en la parte superior derecha del presente, manifiesto bajo protesta de decir verdad y autorizo ` +
      `a FINANCIERA CUALLI, S.A.P.I. DE C.V., SOFOM E.N.R., en su carácter de "El Acreditante" y/o ` +
      `"La Financiera", para que lleve a cabo las dispersiones de las solicitudes que se indican en ` +
      `el cuadro reflejado en la parte superior, por lo que en este acto, AUTORIZO a la Financiera ` +
      `para que se depositen los recursos en las cuentas indicadas, manifestado que dicho capital ` +
      `será descontado de la línea de crédito solicitada y se tomarán como disposiciones realizadas ` +
      `a la misma línea, por lo anterior en este momento me doy por enterado, acepto y autorizo.`;

    

  } else {
    const trato = inferirTrato_(meta.acreditado);
    const tratoMay = trato.toUpperCase();
    textoAutorizacion =
      `Yo, ${acreditadoPersonaFisica || '_____________________'} en mi carácter de "Acreditado y/o Cliente", ` +
      `por propio derecho y con número de cliente indicado en la parte superior derecha del ` +
      `presente, manifiesto bajo protesta de decir verdad y autorizo a FINANCIERA CUALLI, S.A.P.I. DE C.V., ` +
      `SOFOM E.N.R., en su carácter de "Acreditante y/o La Financiera", para que lleve a cabo las ` +
      `dispersiones de las solicitudes que se indican en el cuadro reflejado en la parte superior, ` +
      `por lo que en este acto, AUTORIZO a la Financiera para que se depositen los recursos en las ` +
      `cuentas indicadas, manifestado que dicho capital será descontado de la línea de crédito ` +
      `solicitada y se tomarán como disposiciones realizadas a la misma línea, por lo anterior en este ` +
      `momento me doy por enterado, acepto y autorizo.`;

    textoFirma = `${tratoMay} ${acreditadoPersonaFisica || ''}\nPOR PROPIO DERECHO`;
  }

  const pAut = body.appendParagraph(textoAutorizacion);
  pAut.setLineSpacing(1.10).setSpacingBefore(0).setSpacingAfter(8)
      .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  const tAut = pAut.editAsText();
  tAut.setFontSize(CONFIG.ESTILOS.TAMANO_TEXTO_NORMAL).setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);

  if (info.esPersonaMoral) {
    const mc = acreditadoFormateado.match(/^"([^"]+)"/);
    if (mc) {
      const nombreCom = mc[1]; const iA = textoAutorizacion.indexOf(`"${nombreCom}"`);
      if (iA !== -1) tAut.setBold(iA + 1, iA + nombreCom.length, true);
    }
    const mt = acreditadoFormateado.match(/"([^"]+)",\s*([^,]+(?:, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD (?:NO )?REGULADA)?),/);
    if (mt) {
      const tipo = mt[2]; const iT = textoAutorizacion.indexOf(tipo);
      if (iT !== -1) tAut.setBold(iT, iT + tipo.length - 1, true);
    }
    if (rep1Nombre) { const i1 = textoAutorizacion.indexOf(rep1Nombre); if (i1 !== -1) tAut.setBold(i1, i1 + rep1Nombre.length - 1, true); }
    if (rep2Nombre) { const i2 = textoAutorizacion.indexOf(rep2Nombre); if (i2 !== -1) tAut.setBold(i2, i2 + rep2Nombre.length - 1, true); }
  } else if (acreditadoPersonaFisica) {
    const iF = textoAutorizacion.indexOf(acreditadoPersonaFisica);
    if (iF !== -1) tAut.setBold(iF, iF + acreditadoPersonaFisica.length - 1, true);
  }
    
  const razonSocial = "FINANCIERA CUALLI, S.A.P.I. DE C.V., SOFOM E.N.R.";
  const iFinanciera = textoAutorizacion.indexOf(razonSocial);
  if (iFinanciera !== -1) {
    tAut.setBold(iFinanciera, iFinanciera + razonSocial.length - 1, true);
  }

  const lineaFirma = '____________________________________________________'; 

  if (info.esPersonaMoral && rep2Nombre) {
    const table = body.appendTable();
    
    table.setBorderWidth(0); 
    
    const filaLineas = table.appendTableRow();
    const celdaLinea1 = filaLineas.appendTableCell(lineaFirma);
    const celdaLinea2 = filaLineas.appendTableCell(lineaFirma);
    
    const nombreSinComa = acreditadoFormateado.replace(/,$/, '');
    const filaEmpresa = table.appendTableRow();
    const celdaEmpresa1 = filaEmpresa.appendTableCell(nombreSinComa);
    const celdaEmpresa2 = filaEmpresa.appendTableCell(nombreSinComa);
    
    const filaNombres = table.appendTableRow();
    const celdaNombre1 = filaNombres.appendTableCell(`${inferirTrato_(rep1Nombre).toUpperCase()} ${rep1Nombre}`);
    const celdaNombre2 = filaNombres.appendTableCell(`${inferirTrato_(rep2Nombre).toUpperCase()} ${rep2Nombre}`);
    
    const filaCargos = table.appendTableRow();
    const celdaCargo1 = filaCargos.appendTableCell('REPRESENTANTE LEGAL');
    const celdaCargo2 = filaCargos.appendTableCell('REPRESENTANTE LEGAL');
    
    const celdas = [
      celdaLinea1, celdaLinea2,
      celdaEmpresa1, celdaEmpresa2, 
      celdaNombre1, celdaNombre2,
      celdaCargo1, celdaCargo2
    ];
    
    celdas.forEach(celda => {
      celda.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(20);
      const paragraph = celda.getChild(0).asParagraph();
      paragraph.setSpacingBefore(0).setSpacingAfter(0);
      paragraph.setLineSpacing(1.15); 
      const text = paragraph.editAsText();
      text.setFontSize(CONFIG.ESTILOS.TAMANO_TEXTO_NORMAL)
          .setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);
    });
    
    [celdaLinea1, celdaLinea2].forEach(celda => {
      const text = celda.getChild(0).asParagraph().editAsText();
      text.setBold(true);
      text.setFontSize(11); 
      celda.setPaddingTop(30); 
    });
    
    [celdaCargo1, celdaCargo2].forEach(celda => {
      const text = celda.getChild(0).asParagraph().editAsText();
      text.setBold(true);
    });
    
    [celdaEmpresa1, celdaEmpresa2].forEach(celda => {
      const text = celda.getChild(0).asParagraph().editAsText();
      const textoCompleto = text.getText();
      
      const m1 = textoCompleto.match(/^"([^"]+)"/);
      if (m1) {
        const nom = m1[1];
        const i = textoCompleto.indexOf(nom);
        if (i !== -1) text.setBold(i, i + nom.length - 1, true);
      }
      
      const m2 = textoCompleto.match(/"([^"]+)",\s*([^,]+(?:, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD (?:NO )?REGULADA)?)/);
      if (m2) {
        const tipo = m2[2];
        const iT = textoCompleto.indexOf(tipo);
        if (iT !== -1) text.setBold(iT, iT + tipo.length - 1, true);
      }
    });
    
    const textNombre1 = celdaNombre1.getChild(0).asParagraph().editAsText();
    const textoNombre1 = textNombre1.getText();
    const trato1 = inferirTrato_(rep1Nombre).toUpperCase();
    const iTrato1 = textoNombre1.indexOf(trato1);
    if (iTrato1 !== -1) textNombre1.setBold(iTrato1, iTrato1 + trato1.length - 1, true);
    const iNombre1 = textoNombre1.indexOf(rep1Nombre);
    if (iNombre1 !== -1) textNombre1.setBold(iNombre1, iNombre1 + rep1Nombre.length - 1, true);
    
    const textNombre2 = celdaNombre2.getChild(0).asParagraph().editAsText();
    const textoNombre2 = textNombre2.getText();
    const trato2 = inferirTrato_(rep2Nombre).toUpperCase();
    const iTrato2 = textoNombre2.indexOf(trato2);
    if (iTrato2 !== -1) textNombre2.setBold(iTrato2, iTrato2 + trato2.length - 1, true);
    const iNombre2 = textoNombre2.indexOf(rep2Nombre);
    if (iNombre2 !== -1) textNombre2.setBold(iNombre2, iNombre2 + rep2Nombre.length - 1, true);

  } else {
    const table = body.appendTable();
    table.setBorderWidth(0);
    
    const filaLinea = table.appendTableRow();
    const celdaLinea = filaLinea.appendTableCell(lineaFirma);
    celdaLinea.setPaddingTop(30).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
    const pLinea = celdaLinea.getChild(0).asParagraph();
    pLinea.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    pLinea.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.15);
    const tLinea = pLinea.editAsText();
    tLinea.setBold(true);
    tLinea.setFontSize(11);
    tLinea.setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);

    if (info.esPersonaMoral) {
      const t1May = inferirTrato_(rep1Nombre).toUpperCase();
      const nombreSinComa = acreditadoFormateado.replace(/,$/, '');
      
      const filaEmpresa = table.appendTableRow();
      const celdaEmpresa = filaEmpresa.appendTableCell(nombreSinComa);
      celdaEmpresa.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
      celdaEmpresa.setWidth(400); // Ancho fijo para forzar el salto de línea
      const pEmpresa = celdaEmpresa.getChild(0).asParagraph();
      pEmpresa.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      pEmpresa.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.15);
      const tEmpresa = pEmpresa.editAsText();
      tEmpresa.setFontSize(CONFIG.ESTILOS.TAMANO_TEXTO_NORMAL);
      tEmpresa.setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);
      
      const m1 = nombreSinComa.match(/^"([^"]+)"/);
      if (m1) {
        const nom = m1[1];
        const i = nombreSinComa.indexOf(nom);
        if (i !== -1) tEmpresa.setBold(i, i + nom.length - 1, true);
      }
      const m2 = nombreSinComa.match(/"([^"]+)",\s*([^,]+(?:, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD (?:NO )?REGULADA)?)/);
      if (m2) {
        const tipo = m2[2];
        const iT = nombreSinComa.indexOf(tipo);
        if (iT !== -1) tEmpresa.setBold(iT, iT + tipo.length - 1, true);
      }
      
      const filaNombre = table.appendTableRow();
      const celdaNombre = filaNombre.appendTableCell(`${t1May} ${rep1Nombre}`);
      celdaNombre.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
      const pNombre = celdaNombre.getChild(0).asParagraph();
      pNombre.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      pNombre.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.15);
      const tNombre = pNombre.editAsText();
      tNombre.setFontSize(CONFIG.ESTILOS.TAMANO_TEXTO_NORMAL);
      tNombre.setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);
      
      const iTrato = 0;
      tNombre.setBold(iTrato, t1May.length - 1, true);
      const iNombre = t1May.length + 1;
      tNombre.setBold(iNombre, iNombre + rep1Nombre.length - 1, true);
      
      const filaCargo = table.appendTableRow();
      const celdaCargo = filaCargo.appendTableCell('REPRESENTANTE LEGAL');
      celdaCargo.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
      const pCargo = celdaCargo.getChild(0).asParagraph();
      pCargo.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      pCargo.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.15);
      const tCargo = pCargo.editAsText();
      tCargo.setFontSize(CONFIG.ESTILOS.TAMANO_TEXTO_NORMAL);
      tCargo.setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);
      tCargo.setBold(true);
      
    } else {
      const tratoMay = inferirTrato_(meta.acreditado).toUpperCase();
      
      const filaNombre = table.appendTableRow();
      const celdaNombre = filaNombre.appendTableCell(`${tratoMay} ${acreditadoPersonaFisica || ''}`);
      celdaNombre.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
      const pNombre = celdaNombre.getChild(0).asParagraph();
      pNombre.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      pNombre.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.15);
      const tNombre = pNombre.editAsText();
      tNombre.setFontSize(CONFIG.ESTILOS.TAMANO_TEXTO_NORMAL);
      tNombre.setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);
      
      const iTrato = 0;
      tNombre.setBold(iTrato, tratoMay.length - 1, true);
      if (acreditadoPersonaFisica) {
        const iNombre = tratoMay.length + 1;
        tNombre.setBold(iNombre, iNombre + acreditadoPersonaFisica.length - 1, true);
      }
      
      const filaCargo = table.appendTableRow();
      const celdaCargo = filaCargo.appendTableCell('POR PROPIO DERECHO');
      celdaCargo.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
      const pCargo = celdaCargo.getChild(0).asParagraph();
      pCargo.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      pCargo.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.15);
      const tCargo = pCargo.editAsText();
      tCargo.setFontSize(CONFIG.ESTILOS.TAMANO_TEXTO_NORMAL);
      tCargo.setFontFamily(CONFIG.ESTILOS.FUENTE_BASE);
      tCargo.setBold(true);
    }
  }
}


function generarDOCS_NATIVO() {
  try {
    const { meta, datos } = leerDatos();
    
    if (!meta.id) {
      throw new Error('ID de operación no encontrado. Verifique la celda B3.');
    }
    
    const nombreDoc = `CuadroTransferencias_${meta.id}_${Utilities.formatDate(new Date(), 'GMT', 'yyyyMMdd')}`;

    const doc   = crearDocumento_(nombreDoc);
    const docId = doc.getId();
    const docUrl= doc.getUrl();

    configurarDocumento_(docId);
    setSectionMarginsCm(docId, CONFIG.MARGENES);
    setHeaderFooterMarginsCm_(docId, 0.01, 0.6);
    configurarHeaderFooter_(docId, meta);
    insertarContenidoPrincipal_(docId, datos, meta);

    toast_('✅ Documento generado', 3);
    openLink_(docUrl, 'Abriendo el documento generado…');
    return docUrl;

  } catch (error) {
    console.error('Error en generarDOCS_NATIVO:', error);
    showError_('Error al generar el documento', error.message, error.stack);
    throw error;
  }
}

function toast_(msg, secs) {
  SpreadsheetApp.getActive().toast(msg, 'Generador', secs || 4);
}

function openLink_(url, titulo) {
  const html = HtmlService.createHtmlOutput(
    `<html><body style="font-family:Arial,sans-serif">
       <p>${titulo ? titulo : '¡Listo! 🚀 Abriendo tu documento…'}</p>
       <script>
         try {
           window.open(${JSON.stringify(url)}, '_blank');
         } catch (e) {}
         setTimeout(function(){ google.script.host.close(); }, 200);
       </script>
     </body></html>`
  ).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModelessDialog(html, '¡Listo! 🚀 Abriendo tu documento…');
}

function showError_(title, message, stack) {
  const esc = s => (s||'').toString()
     .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
     .replace(/'/g,'&#39;').replace(/"/g,'&quot;');
  const html = HtmlService.createHtmlOutput(
    `<html><body style="font-family:Arial,sans-serif;padding:12px;max-width:600px">
       <h3 style="margin:0 0 6px;color:#b00020">❌ ${esc(title)}</h3>
       <div style="white-space:pre-wrap;font-size:12px;border:1px solid #eee;background:#fafafa;padding:10px;border-radius:6px">${esc(message)}</div>
       ${stack ? `<details style="margin-top:8px"><summary>Detalles</summary>
         <pre style="white-space:pre-wrap">${esc(stack)}</pre></details>` : ``}
       <button onclick="navigator.clipboard.writeText('${esc(message)}\\n\\n${esc(stack||'')}')"
               style="margin-top:10px;padding:6px 10px;border:1px solid #ddd;border-radius:6px;background:#fff;cursor:pointer">
         Copiar error
       </button>
       <button onclick="google.script.host.close()"
               style="margin-top:10px;margin-left:6px;padding:6px 10px;border:1px solid #ddd;border-radius:6px;background:#fff;cursor:pointer">
         Cerrar
       </button>
       <script>setTimeout(()=>{},0)</script>
     </body></html>`
  ).setWidth(480).setHeight(300);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Error');
}

function construirTablasHTML_DocLikeInline_(datos, meta) {
  const esc = (s) => !s ? '' : s.toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');

  const W0 = 76.5;
  const W1 = 86.5;
  const W2 = 140;
  const W3 = 140;
  const W4 = cmToPt(7.141);
  const TOTAL_PT = W0 + W1 + W2 + W3 + W4;

  const tdBase   = [
    'border:0.5pt solid #666',
    'padding:3px',
    'vertical-align:middle',
    'text-align:center',
    'font-family:Arial,sans-serif',
    'font-size:10pt',
    'white-space:normal',
    'word-break:break-word',
    'overflow-wrap:anywhere'
  ].join(';') + ';';

  const tdHdr    = tdBase + 'background:#b7b7b7;font-weight:bold;';
  const tdSubHdr = tdBase + 'background:#b7b7b7;font-weight:bold;';
  const tableSt  = [
    `width:${TOTAL_PT}pt`,
    'border-collapse:collapse',
    'margin:4px 0',
    'table-layout:fixed'
  ].join(';') + ';';

  return datos.map(d => {
    const montoFmt = formatMontoMillonesMX_(d.monto, meta ? meta.moneda : 'MXN'); 
    return `
    <table style="${tableSt}">
      <colgroup>
        <col style="width:${W0}pt">
        <col style="width:${W1}pt">
        <col style="width:${W2}pt">
        <col style="width:${W3}pt">
        <col style="width:${W4}pt">
      </colgroup>

      <tr>
        <td style="${tdHdr}">No. de Solicitud</td>
        <td style="${tdHdr}">Fecha de Solicitud</td>
        <td style="${tdHdr}">Concepto</td>
        <td style="${tdHdr}">Monto</td>
        <td style="${tdHdr}">Comentarios</td>
      </tr>

      <tr>
        <td rowspan="3" style="${tdBase}font-weight:bold;">${esc(d.num)}</td>
        <td rowspan="3" style="${tdBase}">${esc(d.fecha)}</td>
        <td style="${tdBase}">${esc(d.concepto)}</td>
         <td style="${tdBase}">${esc(montoFmt)}</td>
        <td style="${tdBase}">${esc(d.comentarios)}</td>
      </tr>

      <tr>
        <td style="${tdSubHdr}">Banco</td>
        <td style="${tdSubHdr}">${etiquetaClabe_(d.clabe)}</td>
        <td style="${tdSubHdr}">BENEFICIARIO</td>
      </tr>

      <tr>
        <td style="${tdBase}">${esc(d.banco)}</td>
        <td style="${tdBase}">${esc(d.clabe)}</td>
        <td style="${tdBase}">${esc(d.beneficiario)}</td>
      </tr>
    </table>
  `;
  }).join('\n');
}


function construirTablasHTML_Completo_(datos, meta) {
  const esc = (s) => !s ? '' : s.toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');

  const tdBase   = [
    'border:0.5pt solid #666',
    'padding:1.135pt',
    'vertical-align:middle',
    'text-align:center',
    'font-family:Arial,sans-serif',
    'font-size:10pt',
    'white-space:normal',
    'word-break:break-word',
    'overflow-wrap:anywhere',
    'height:30px'
  ].join(';') + ';';

  const tdHdr    = tdBase + 'background:#b7b7b7;font-weight:bold;font-size:10pt;padding:5px;';
  const tdSubHdr = tdBase + 'background:#b7b7b7;font-weight:bold;font-size:10pt;padding:5px;';
  
  const tableSt  = [
    'width:100%',
    'border-collapse:collapse',
    'margin:5px 0',
    'table-layout:fixed'
  ].join(';') + ';';

  return datos.map(d => {
    const montoFmt = formatMontoMillonesMX_(d.monto, meta ? meta.moneda : 'MXN');
    return `
    <table class="tabla-cuadro" style="${tableSt}">
      <colgroup>
        <col style="width:10%">   <!-- No. Solicitud -->
        <col style="width:10%">  <!-- Fecha -->
        <col style="width:20%">  <!-- Concepto -->
        <col style="width:20%">  <!-- Monto -->
        <col style="width:25%">  <!-- Comentarios -->
      </colgroup>

      <tr>
        <td style="${tdHdr}">No. de Solicitud</td>
        <td style="${tdHdr}">Fecha de Solicitud</td>
        <td style="${tdHdr}">Concepto</td>
        <td style="${tdHdr}">Monto</td>
        <td style="${tdHdr}">Comentarios</td>
      </tr>

      <tr>
        <td rowspan="3" style="${tdBase}font-weight:bold;">${esc(d.num)}</td>
        <td rowspan="3" style="${tdBase}">${esc(d.fecha)}</td>
        <td style="${tdBase}">${esc(d.concepto)}</td>
        <td style="${tdBase}">${esc(montoFmt)}</td>
        <td style="${tdBase}">${esc(d.comentarios)}</td>
      </tr>

      <tr>
        <td style="${tdSubHdr}">Banco</td>
        <td style="${tdSubHdr}">${etiquetaClabe_(d.clabe)}</td>
        <td style="${tdSubHdr}">BENEFICIARIO</td>
      </tr>

      <tr>
        <td style="${tdBase}">${esc(d.banco)}</td>
        <td style="${tdBase}">${esc(d.clabe)}</td>
        <td style="${tdBase}">${esc(d.beneficiario)}</td>
      </tr>
    </table>
  `;
  }).join('\n');
}


function previsualizarCuadrosModal() {
  try {
    const { meta, datos } = leerDatos();
    const tablasHTML = construirTablasHTML_DocLikeInline_(datos, meta);
    
    const saludo = obtenerSaludoAutomatico_();

    const html = HtmlService.createHtmlOutput(`
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Previsualización (con borrador)</title>
  <style>
    :root { --accent:#2563eb; --ok:#16a34a; --warn:#b45309; --err:#b00020; }
    body{font-family:Arial,sans-serif;margin:0; background: #ffffff;}
    .banner{
      position:sticky; top:0; z-index:9999;
      width:100%; padding:10px 14px;
      background:#111827; color:#fff;
      display:none; opacity:0; transform:translateY(-6px);
      transition:opacity .25s ease, transform .25s ease;
      font-size:13px;
    }
    .banner.ok { background: var(--ok); }
    .banner.warn { background: var(--warn); }
    .banner.err { background: var(--err); }
    .wrap{padding:12px;}
    .bar{position:sticky; top:0; background:#fff; padding:8px 0 12px; margin:0 0 12px; border-bottom:1px solid #eee;}
    
    .btn{
      padding:10px 16px; 
      border: none;
      border-radius:8px; 
      background: #2563eb;
      color: white;
      cursor:pointer;
      font-weight: 600;
      font-size: 14px;
      transition: all 0.2s ease;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    .btn:hover {
      background: #1d4ed8;
      transform: translateY(-1px);
      box-shadow: 0 2px 5px rgba(0,0,0,0.15);
    }
    
    .btn:active {
      transform: translateY(0);
      box-shadow: 0 1px 2px rgba(0,0,0,0.1);
    }
    
    .btn-secondary {
      background: #6b7280;
    }
    
    .btn-secondary:hover {
      background: #4b5563;
    }
    
    .btn + .btn{margin-left:8px}
    
    .btn-icon {
      width: 16px;
      height: 16px;
    }
    
    .hint{color:#666; font-size:12px; margin-left:8px}
    .viewport{max-height:75vh; overflow:auto; border:1px solid #eee; border-radius:8px; padding:12px; background:#fff}
    
    .page-title {
      display: flex;
      align-items: center;
      justify-content:center;
      gap: 8px;
      font-size: 18px;
      font-weight: bold;
      margin-bottom: 8px;
      color: #1f2937;
    }
    
    .lightning-icon {
      width: 20px;
      height: 20px;
      color: #f59e0b;
    }
    
    .footer {
      text-align: center;
      padding: 16px 0;
      margin-top: 20px;
      color: #6b7280;
      font-size: 12px;
      border-top: 1px solid #eee;
    }

    .title-logo {
      width: 32px;
      height: 32px;
      vertical-align: middle;
      margin-right: 12px;
    }
    
    .borrador {
      font-family: Arial, sans-serif;
      font-size: 14px;
      line-height: 1.5;
      color: #000000;
      margin-bottom: 20px;
      padding: 0 8px;
    }
    
    .borrador p {
      margin: 0 0 12px 0;
    }
    
    .controls-panel {
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 16px;
      margin-bottom: 16px;
    }
    
    .control-group {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 12px;
    }
    
    .control-label {
      font-weight: 600;
      color: #374151;
      font-size: 14px;
      min-width: 120px;
    }
    
    .select-control {
      padding: 8px 12px;
      border: 1px solid #d1d5db;
      border-radius: 6px;
      background: white;
      font-size: 14px;
      min-width: 140px;
      cursor: pointer;
    }
    
    .select-control:focus {
      outline: none;
      border-color: #2563eb;
      box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
    }
    
    .preview-section {
      background: white;
      border: 1px solid #e5e7eb;
      border-radius: 8px;
      padding: 16px;
      margin-top: 16px;
    }
    
    .preview-title {
      font-weight: 600;
      color: #374151;
      margin-bottom: 12px;
      font-size: 14px;
    }
  </style>
</head>
<body>
  <div id="banner" class="banner"></div>

  <div class="wrap">
    <div class="page-title">
      <img src="https://img.icons8.com/?size=100&id=63647&format=png&color=000000" alt="Checklist" class="title-logo">
      Previsualización del Cuadro para Gmail
    </div>
    
    <div class="bar">
      <button class="btn" onclick="copiarTodo()">
        <svg class="btn-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"></path>
        </svg>
        Copiar Todo
      </button>
      <button class="btn btn-secondary" onclick="abrirNueva()">
        <svg class="btn-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path>
        </svg>
        Abrir en pestaña nueva
      </button>
      <span class="hint">Se copiará el borrador completo con el cuadro</span>
    </div>

    <div class="controls-panel">
      <div class="control-group">
        <span class="control-label">Fecha de firma:</span>
        <select id="fechaFirma" class="select-control" onchange="actualizarBorrador()">
          <option value="hoy">Hoy</option>
          <option value="mañana">Mañana</option>
        </select>
      </div>
    </div>

    <div class="preview-section">
      <div class="preview-title">Vista previa del borrador:</div>
      <div id="content" class="viewport">
        <div id="borradorContent" class="borrador">
          <!-- El borrador se actualizará dinámicamente -->
        </div>
        ${tablasHTML}
      </div>
    </div>
    
    <div class="footer">
      © ${new Date().getFullYear()} Financiera Cualli. Todos los derechos reservados.
    </div>
  </div>

  <script>
    document.addEventListener('DOMContentLoaded', function() {
      actualizarBorrador();
    });

    function actualizarBorrador() {
      const fechaSeleccionada = document.getElementById('fechaFirma').value;
      const saludo = '${saludo}';
      
      let textoFecha;
      if (fechaSeleccionada === 'hoy') {
        textoFecha = 'el día de hoy';
      } else {
        textoFecha = 'el dia de mañana';
      }
      
      const borradorHTML = \`
        <p>Estimado equipo, \${saludo}</p>
        
        <p>Les comentamos que \${textoFecha} estaremos firmando la operación en referencia, favor de programar las siguientes dispersiones:</p>
      \`;
      
  const borradorDiv = document.getElementById('borradorContent');
  if (borradorDiv) {
    borradorDiv.innerHTML = borradorHTML;
  }

  // 4) Mensaje final DESPUÉS de las tablas
  const content = document.getElementById('content');
  if (!content) return;

  // Si ya existe un cierre anterior, lo quitamos para no duplicar
  const cierrePrevio = document.getElementById('cierreCorreo');
  if (cierrePrevio) {
    cierrePrevio.remove();
  }

  const cierre = document.createElement('p');
  cierre.id = 'cierreCorreo';
  cierre.innerHTML = 'Saludos,<br>Equipo Jurídico.';
  cierre.style.marginTop = '12px';
  cierre.style.fontFamily = 'Arial, sans-serif';
  cierre.style.fontSize = '14px';

  // Se inserta al final de #content → después de las tablas
  content.appendChild(cierre);
}

    function showBanner(msg, kind){
      const el = document.getElementById('banner');
      el.className = 'banner ' + (kind||'');
      el.textContent = msg;
      el.style.display = 'block';
      requestAnimationFrame(()=>{ el.style.opacity = '1'; el.style.transform = 'translateY(0)'; });
      clearTimeout(window.__bannerTimer);
      window.__bannerTimer = setTimeout(()=>{
        el.style.opacity = '0'; el.style.transform = 'translateY(-6px)';
        setTimeout(()=>{ el.style.display='none'; }, 250);
      }, 2200);
    }

    function seleccionarNodo(el){
      const range = document.createRange();
      range.selectNodeContents(el);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }

    async function copiarTodo(){
      const el = document.getElementById('content');
      const html = el.innerHTML;
      const type = 'text/html';

      if (navigator.clipboard && window.ClipboardItem) {
        try {
          const blob = new Blob([html], {type});
          await navigator.clipboard.write([ new ClipboardItem({[type]: blob}) ]);
          showBanner('Borrador y cuadro copiados. Pega en Gmail.', 'ok');
          return;
        } catch(e){ }
      }
      try {
        seleccionarNodo(el);
        const ok = document.execCommand('copy');
        showBanner(ok ? 'Contenido copiado (fallback). Pega en Gmail.' : 'No se pudo copiar automáticamente.', ok ? 'ok' : 'warn');
      } catch(e){
        showBanner('No se pudo copiar automáticamente. Copia manualmente.', 'err');
      }
    }

    function abrirNueva(){
      const html = document.documentElement.outerHTML;
      const w = window.open('about:blank','_blank');
      w.document.open(); w.document.write(html); w.document.close();
      showBanner('Abriendo en pestaña nueva…', 'ok');
    }
  </script>
</body>
</html>
    `).setWidth(1000).setHeight(750);

    SpreadsheetApp.getUi().showModalDialog(html, ' ');
  } catch (error) {
    console.error(error);
    showError_('Error en previsualización', error.message, error.stack);
    throw error;
  }
}

function obtenerSaludoAutomatico_() {
  const hora = new Date().getHours();
  
  if (hora >= 5 && hora < 12) {
    return "buenos días.";
  } else if (hora >= 12 && hora < 19) {
    return "buenas tardes.";
  } else {
    return "buenas noches.";
  }
}

function previsualizarDocumentoCompleto() {
  try {
    const { meta, datos } = leerDatos();
    const htmlCompleto = construirHTMLCompletoDocumento_(datos, meta);

    const html = HtmlService.createHtmlOutput(`
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Previsualización Completa del Documento</title>
  <style>
    :root { 
      --accent:#2563eb; --ok:#16a34a; --warn:#b45309; --err:#b00020;
      --color-header: #b7b7b7;
      --color-subhdr: #b7b7b7;
      --page-width: 29.7cm;
      --page-height: 21.0cm;
      --margin-top: 3.44cm;
      --margin-bottom: 3.44cm;
      --margin-left: 2.40cm;
      --margin-right: 2.44cm;
    }
    
    @page {
      size: A4 landscape;
      margin: 0;
    }
    
    body {
      font-family: Arial, sans-serif;
      margin: 0;
      background: #ffffff;
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 20px;
    }
    
    .banner{
      position: sticky; 
      top: 0; 
      z-index: 9999;
      width: 100%; 
      padding: 10px 14px;
      background: #111827; 
      color: #fff;
      display: none; 
      opacity: 0; 
      transform: translateY(-6px);
      transition: opacity .25s ease, transform .25s ease;
      font-size: 13px;
    }
    
    .banner.ok { background: var(--ok); }
    .banner.warn { background: var(--warn); }
    .banner.err { background: var(--err); }
    
    .document-container {
      width: var(--page-width);
      min-height: var(--page-height);
      background: white;
      box-shadow: 0 4px 12px rgba(0,0,0,0.15);
      margin-bottom: 20px;
      position: relative;
      padding: var(--margin-top) var(--margin-right) var(--margin-bottom) var(--margin-left);
      box-sizing: border-box;
    }
    
    .header {
      position: absolute;
      top: 0;
      left: var(--margin-left);
      right: var(--margin-right);
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding-top: 15px;
      padding-bottom: 15px; 
      border-bottom: 1px solid #ddd;
    }
    
    .logo-container {
      flex: 0 0 auto;
    }
    
    .logo {
      width: ${CONFIG.LOGO.WIDTH}px;
      height: ${CONFIG.LOGO.HEIGHT}px;
      object-fit: contain;
    }
    
    .header-info {
      flex: 1;
      text-align: right;
      font-size: 10pt;
      font-family: Arial, sans-serif;
    }
    
    .header-info p {
      margin: 2px 0;
      font-weight: bold;
    }
    
    .content {
      font-family: Arial, sans-serif;
      font-size: 10pt;
      line-height: 1.0;
    }
    
    .tabla-cuadro {
      width: 100%;
      border-collapse: collapse;
      margin: 15px 0;
      font-size: 9pt;
      table-layout: fixed;
    }
    
    .tabla-cuadro td {
      border: 0.5pt solid #666;
      padding: 3px;
      vertical-align: middle;
      text-align: center;
      font-family: Arial, sans-serif;
      white-space: normal;
      word-break: break-word;
      overflow-wrap: anywhere;
    }
    
    .celda-encabezado {
      background: var(--color-header);
      font-weight: bold;
      font-size: 10pt;
    }
    
    .celda-subencabezado {
      background: var(--color-subhdr);
      font-weight: bold;
      font-size: 10pt;
    }
    
    .texto-autorizacion {
      text-align: justify;
      line-height: 1.0;
      margin: 20px 0;
      font-size: 10pt;
    }
    
    .firma-section {
      margin-top: 40px;
    }
    
    .linea-firma {
      margin-bottom: 5px;
      font-weight: bold;
    }
    
    .texto-firma {
      font-size: 10pt;
      margin: 3px 0;
    }
    
     .footer {
      position: absolute;
      bottom: 0.6cm;
      left: var(--margin-left);
      right: var(--margin-right);
      height: 1.2cm;
      border-top: 1px solid #ddd;
      padding-top: 8px;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      font-family: Arial, sans-serif;
    }
    
    .footer-left {
      text-align: left;
    }
    
    .footer-right {
      text-align: right;
      font-family: 'Courier New', monospace;
      font-size: 7.5pt;
      color: #666666;
    }

    .acelerando {
      font-size: 7.5pt;
      font-weight: bold;
    }
    
    .direccion {
      font-size: 6.8pt;
      color: #515151;
      margin: 0;
    }


    .controls-bar {
      position: sticky;
      top: 0;
      background: #fff;
      padding: 12px;
      margin: 0 0 20px 0;
      border-bottom: 1px solid #eee;
      width: var(--page-width);
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      z-index: 1000;
    }
    
    .btn{
      padding: 10px 16px; 
      border: none;
      border-radius: 8px; 
      background: #2563eb;
      color: white;
      cursor: pointer;
      font-weight: 600;
      font-size: 14px;
      transition: all 0.2s ease;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
      margin-right: 8px;
    }
    
    .btn:hover {
      background: #1d4ed8;
      transform: translateY(-1px);
      box-shadow: 0 2px 5px rgba(0,0,0,0.15);
    }
    
    .btn-secondary {
      background: #45aad4;
    }
    
    .btn-secondary:hover {
      background: #4b5563;
    }
    
    .btn-icon {
      width: 16px;
      height: 16px;
    }
    
    .page-title {
      font-size: 18px;
      font-weight: bold;
      margin-bottom: 8px;
      color: #1f2937;
      text-align: center;
    }
    
    .hint {
      color: #666; 
      font-size: 12px; 
      margin-left: 8px;
    }

    .firma-doble {
      display: flex;
      justify-content: space-between;
      gap: 40px;
      margin-top: 20px;
    }

    .firma-columna {
      flex: 1;
    }

    .texto-firma strong {
      font-weight: bold;
    }

    .linea-firma {
    
      margin-bottom: 5px;
      font-weight: bold;
    }

    .title-logo {
    width: 40px;
    height: 40px;
    vertical-align: middle;
    margin-right: 12px;
  }

  </style>
</head>
<body>
  <div id="banner" class="banner"></div>

  <div class="wrap">
    <div class="page-title">
      <img src="https://img.icons8.com/?size=100&id=63647&format=png&color=000000" alt="Checklist" class="title-logo">
      Vista Previa del Documento Completo
    </div>

  <div class="controls-bar">
    <div style="display: flex; align-items: center; justify-content: space-between;">
      <button class="btn btn-secondary" onclick="abrirNuevaVentana()">
          <svg class="btn-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path>
          </svg>
          Abrir en nueva ventana
        </button>
        </div>
        </div>

  <div class="document-container">
    <!-- Header -->
    <div class="header">
      <div class="logo-container">
        <img src="${CONFIG.LOGO.URL}" alt="Logo" class="logo" onerror="this.style.display='none'">
      </div>
      <div class="header-info">
        <p>ID Operación: ${meta.id || ''}</p>
        <p>Cliente: ${meta.cliente || ''}</p>
        <p>Fecha: ${meta.fechaHoy}</p>
      </div>
    </div>

    <!-- Contenido principal -->
    <div class="content">
      ${htmlCompleto}
    </div>

    <!-- Footer -->
<div class="footer">
      <div class="footer-left">
        <div class="acelerando">
          <span style="color: #fbb818;">acelerando</span><span style="color: #525352;">oportunidades</span>
        </div>
        <p class="direccion">
          <span style="color: #fbb818;">Ciudad de México.</span> Torre del Árbol Blvd. Manuel Ávila Camacho 184, Piso 7 Colonia Reforma Social. CP 11650. +(52) 55 8117 1700
        </p>
      </div>
      <div class="footer-right">
        ${meta.footer}
      </div>
    </div>

  <script>
    function showBanner(msg, kind){
      const el = document.getElementById('banner');
      el.className = 'banner ' + (kind||'');
      el.textContent = msg;
      el.style.display = 'block';
      requestAnimationFrame(()=>{ el.style.opacity = '1'; el.style.transform = 'translateY(0)'; });
      clearTimeout(window.__bannerTimer);
      window.__bannerTimer = setTimeout(()=>{
        el.style.opacity = '0'; el.style.transform = 'translateY(-6px)';
        setTimeout(()=>{ el.style.display='none'; }, 250);
      }, 2200);
    }

   

    function abrirNuevaVentana() {
      const html = document.documentElement.outerHTML;
      const w = window.open('about:blank', '_blank', 'width=1200,height=800,scrollbars=yes');
      w.document.open();
      w.document.write(html);
      w.document.close();
      showBanner('Abriendo en nueva ventana…', 'ok');
    }

    window.addEventListener('load', function() {
      const content = document.querySelector('.content');
      const container = document.querySelector('.document-container');
      if (content && container) {
        const contentHeight = content.scrollHeight;
        const minHeight = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--page-height'));
        if (contentHeight > minHeight) {
          container.style.minHeight = contentHeight + 'px';
        }
      }
    });
  </script>
</body>
</html>
    `).setWidth(1200).setHeight(900);

    SpreadsheetApp.getUi().showModalDialog(html, ' ');
  } catch (error) {
    console.error(error);
    showError_('Error en previsualización completa', error.message, error.stack);
    throw error;
  }
}

function construirHTMLCompletoDocumento_(datos, meta) {
  const tablasHTML = construirTablasHTML_Completo_(datos, meta);
  const autorizacionHTML = construirHTMLAutorizacion_(meta);
  
  return `
    ${tablasHTML}
    ${autorizacionHTML}
  `;
}

function construirHTMLAutorizacion_(meta) {
  const info = determinarTipoPersonaYNombre(meta);
  const acreditadoFormateado = formatearNombreEmpresa_(info.nombreAcreditado);
  const acreditadoPersonaFisica = formatearNombrePersona_(meta.acreditado);

  // Usar los valores ya filtrados de info (sin N/A)
  const rep1Nombre = formatearNombrePersona_(info.nombreRepresentante1 || '');
  const rep2Nombre = formatearNombrePersona_(info.nombreRepresentante2 || '');

  let textoAutorizacion = '';
  let firmaHTML = '';

  if (info.esPersonaMoral) {
    const t1 = inferirTrato_(rep1Nombre);
    const a1 = inferirArticulo_(t1);
    let reps = `${a1} ${t1} ${rep1Nombre}`;
    if (rep2Nombre) {
      const t2 = inferirTrato_(rep2Nombre);
      const a2 = inferirArticulo_(t2);
      reps += ` y ${a2} ${t2} ${rep2Nombre}`;
    }

    textoAutorizacion =
      `${acreditadoFormateado} en su carácter de "El Acreditado" y/o "El Cliente", ` +
      `representado por ${reps}, con número de cliente indicado ` +
      `en la parte superior derecha del presente, manifiesto bajo protesta de decir verdad y autorizo ` +
      `a FINANCIERA CUALLI, S.A.P.I. DE C.V., SOFOM E.N.R., en su carácter de "El Acreditante" y/o ` +
      `"La Financiera", para que lleve a cabo las dispersiones de las solicitudes que se indican en ` +
      `el cuadro reflejado en la parte superior, por lo que en este acto, AUTORIZO a la Financiera ` +
      `para que se depositen los recursos en las cuentas indicadas, manifestado que dicho capital ` +
      `será descontado de la línea de crédito solicitada y se tomarán como disposiciones realizadas ` +
      `a la misma línea, por lo anterior en este momento me doy por enterado, acepto y autorizo.`;

    if (rep2Nombre) {
      const t1May = t1.toUpperCase();
      const t2May = inferirTrato_(rep2Nombre).toUpperCase();
      
      const nombreSinComa = acreditadoFormateado.replace(/,$/, '');
      let empresaConNegritas = nombreSinComa;
      
      const matchComillas = nombreSinComa.match(/^"([^"]+)"/);
      if (matchComillas) {
        const nombreCom = matchComillas[1];
        empresaConNegritas = empresaConNegritas.replace(
          `"${nombreCom}"`, 
          `"<strong>${nombreCom}</strong>"`
        );
      }
      
      const matchTipo = nombreSinComa.match(/"([^"]+)",\s*([^,]+(?:, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD (?:NO )?REGULADA)?)/);
      if (matchTipo) {
        const tipo = matchTipo[2];
        empresaConNegritas = empresaConNegritas.replace(
          tipo, 
          `<strong>${tipo}</strong>`
        );
      }

      firmaHTML = `
        <div class="firma-doble">
          <div class="firma-columna">
            <div class="linea-firma">___________________________________</div>
            <div class="texto-firma">${empresaConNegritas}</div>
            <div class="texto-firma"><strong>${t1May}</strong> <strong>${rep1Nombre}</strong></div>
            <div class="texto-firma"><strong>REPRESENTANTE LEGAL</strong></div>
          </div>
          <div class="firma-columna">
            <div class="linea-firma">___________________________________</div>
            <div class="texto-firma">${empresaConNegritas}</div>
            <div class="texto-firma"><strong>${t2May}</strong> <strong>${rep2Nombre}</strong></div>
            <div class="texto-firma"><strong>REPRESENTANTE LEGAL</strong></div>
          </div>
        </div>
      `;
    } else {
      const t1May = t1.toUpperCase();
      
      const nombreSinComa = acreditadoFormateado.replace(/,$/, '');
      let empresaConNegritas = nombreSinComa;
      
      const matchComillas = nombreSinComa.match(/^"([^"]+)"/);
      if (matchComillas) {
        const nombreCom = matchComillas[1];
        empresaConNegritas = empresaConNegritas.replace(
          `"${nombreCom}"`, 
          `"<strong>${nombreCom}</strong>"`
        );
      }
      
      const matchTipo = nombreSinComa.match(/"([^"]+)",\s*([^,]+(?:, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD (?:NO )?REGULADA)?)/);
      if (matchTipo) {
        const tipo = matchTipo[2];
        empresaConNegritas = empresaConNegritas.replace(
          tipo, 
          `<strong>${tipo}</strong>`
        );
      }

      firmaHTML = `
        <div class="linea-firma">_______________________________________</div>
        <div class="texto-firma">${empresaConNegritas}</div>
        <div class="texto-firma"><strong>${t1May}</strong> <strong>${rep1Nombre}</strong></div>
        <div class="texto-firma"><strong>REPRESENTANTE LEGAL</strong></div>
      `;
    }

  } else {
    const trato = inferirTrato_(meta.acreditado);
    const tratoMay = trato.toUpperCase();
    textoAutorizacion =
      `Yo, ${acreditadoPersonaFisica || '_____________________'} en mi carácter de "Acreditado y/o Cliente", ` +
      `por propio derecho y con número de cliente indicado en la parte superior derecha del ` +
      `presente, manifiesto bajo protesta de decir verdad y autorizo a FINANCIERA CUALLI, S.A.P.I. DE C.V., ` +
      `SOFOM E.N.R., en su carácter de "Acreditante y/o La Financiera", para que lleve a cabo las ` +
      `dispersiones de las solicitudes que se indican en el cuadro reflejado en la parte superior, ` +
      `por lo que en este acto, AUTORIZO a la Financiera para que se depositen los recursos en las ` +
      `cuentas indicadas, manifestado que dicho capital será descontado de la línea de crédito ` +
      `solicitada y se tomarán como disposiciones realizadas a la misma línea, por lo anterior en este ` +
      `momento me doy por enterado, acepto y autorizo.`;

    firmaHTML = `
      <div class="linea-firma">_________________________________________</div>
      <div class="texto-firma"><strong>${tratoMay}</strong> <strong>${acreditadoPersonaFisica || ''}</strong></div>
      <div class="texto-firma"><strong>POR PROPIO DERECHO</strong></div>
    `;
  }

  let textoAutorizacionConNegritas = textoAutorizacion;

  if (info.esPersonaMoral) {
    const matchComillas = acreditadoFormateado.match(/^"([^"]+)"/);
    if (matchComillas) {
      const nombreCom = matchComillas[1];
      const regexNombre = new RegExp(`"${nombreCom}"`, 'g');
      textoAutorizacionConNegritas = textoAutorizacionConNegritas.replace(
        regexNombre, 
        `"<strong>${nombreCom}</strong>"`
      );
    }

    const matchTipo = acreditadoFormateado.match(/"([^"]+)",\s*([^,]+(?:, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD (?:NO )?REGULADA)?),/);
    if (matchTipo) {
      const tipo = matchTipo[2];
      const regexTipo = new RegExp(tipo.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
      textoAutorizacionConNegritas = textoAutorizacionConNegritas.replace(
        regexTipo, 
        `<strong>${tipo}</strong>`
      );
    }

    if (rep1Nombre) {
      const regexRep1 = new RegExp(rep1Nombre.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
      textoAutorizacionConNegritas = textoAutorizacionConNegritas.replace(
        regexRep1, 
        `<strong>${rep1Nombre}</strong>`
      );
    }
    if (rep2Nombre) {
      const regexRep2 = new RegExp(rep2Nombre.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
      textoAutorizacionConNegritas = textoAutorizacionConNegritas.replace(
        regexRep2, 
        `<strong>${rep2Nombre}</strong>`
      );
    }
  } else if (acreditadoPersonaFisica) {
    const regexNombre = new RegExp(acreditadoPersonaFisica.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
    textoAutorizacionConNegritas = textoAutorizacionConNegritas.replace(
      regexNombre, 
      `<strong>${acreditadoPersonaFisica}</strong>`
    );
  }

  textoAutorizacionConNegritas = textoAutorizacionConNegritas.replace(
    /FINANCIERA CUALLI, S\.A\.P\.I\. DE C\.V\., SOFOM E\.N\.R\./g, 
    '<strong>FINANCIERA CUALLI, S.A.P.I. DE C.V., SOFOM E.N.R.</strong>'
  );

  return `
    <div class="texto-autorizacion">
      ${textoAutorizacionConNegritas}
    </div>
    <div class="firma-section">
      ${firmaHTML}
    </div>
  `;
}

function mostrarSidebarCuadroDispersion() {
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html lang="es">

<head>
    <base target="_top">
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {
            --comic-ink: #000;
            --comic-yellow: #ffcc00;
            --comic-orange: #ff9500;
            --comic-red: #ff3b30;
            --comic-blue: #007aff;
            --comic-green: #4cd964;
            --comic-paper: #fff;
            --comic-light: #f8f9fa;
            --comic-shadow: 3px 3px 0 var(--comic-ink);
            --comic-shadow-lg: 5px 5px 0 var(--comic-ink);
            --comic-border: 3px solid var(--comic-ink);
            --comic-radius: 12px;
            --comic-radius-lg: 16px;
            --transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: 'Poppins', Arial, sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #e4e8f0 100%);
            padding: 0;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .container {
            background: var(--comic-paper);
            border: var(--comic-border);
            border-radius: var(--comic-radius-lg);
            box-shadow: var(--comic-shadow-lg), 0 10px 25px rgba(0, 0, 0, 0.1);
            padding: 24px;
            width: 100%;
            max-width: 380px;
            transition: var(--transition);
            position: relative;
            overflow: hidden;
            margin: 5px;
        }

        .container::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 6px;
            border-radius: var(--comic-radius-lg) var(--comic-radius-lg) 0 0;
        }

        h2 {
            display: flex;
            align-items: center;
            gap: 12px;
            color: var(--comic-ink);
            margin-bottom: 25px;
            font-weight: 700;
            font-size: 1.1rem;
            padding-bottom: 15px;
            border-bottom: 2px dashed rgba(0, 0, 0, 0.1);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        h2 i {
            color: var(--comic-paper);
            background: var(--comic-blue);
            padding: 10px;
            border-radius: 10px;
            width: 42px;
            height: 42px;
            display: flex;
            align-items: center;
            justify-content: center;
            border: 2px solid var(--comic-ink);
            box-shadow: var(--comic-shadow);
        }

        .btn-group {
            display: flex;
            flex-direction: column;
            gap: 15px;
            margin: 20px 0;
        }

        .btn {
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 20px 20px 20px 5px;
            border: 3px solid var(--comic-ink);
            border-radius: 16px;
            cursor: pointer;
            font-weight: 700;
            font-size: 14px;
            transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
            position: relative;
            overflow: hidden;
            text-align: left;
            color: var(--comic-ink);
            box-shadow:
                4px 4px 0 var(--comic-ink),
                0 8px 20px rgba(0, 0, 0, 0.15);
            text-transform: uppercase;
            letter-spacing: 0.8px;
            width: 100%;
            min-height: 90px;
            height: auto;
            align-items: flex-start;
        }

        .btn-content {
            display: flex;
            flex-direction: column;
            flex: 1;
            position: relative;
            z-index: 2;
            min-width: 0;
            width: 100%;
        }

        .btn-title {
            font-weight: 800;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            margin-bottom: 4px;
            text-shadow: 1px 1px 0 rgba(255, 255, 255, 0.8);
            white-space: normal;
            overflow: visible;
            text-overflow: unset;
            width: 100%;
            line-height: 1.2;
            max-height: 2.4em;
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
        }

        .btn-desc {
            font-size: 10px;
            opacity: 0.9;
            font-weight: 600;
            letter-spacing: 0.4px;
            line-height: 1.3;
            white-space: normal;
            overflow: visible;
            text-overflow: unset;
            width: 100%;
            max-height: 2.6em;
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
        }

        .btn-icon {
            width: 38px;
            height: 38px;
            flex-shrink: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 10px;
            padding: 8x;
            background: rgba(255, 255, 255, 0.95);
            border: 3px solid var(--comic-ink);
            box-shadow:
                1px 1px 0 var(--comic-ink),
                inset 0 1px 1px rgba(255, 255, 255, 0.8);
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }

        .btn-icon::before {
            content: '';
            position: absolute;
            top: -50%;
            left: -50%;
            width: 200%;
            height: 200%;
            background: linear-gradient(45deg,
                    transparent,
                    rgba(255, 255, 255, 0.3),
                    transparent);
            transform: rotate(45deg);
            transition: all 0.6s ease;
        }

        .btn:hover .btn-icon::before {
            left: 50%;
            top: 50%;
        }

        .btn-icon img {
            width: 28px;
            height: 28px;
            object-fit: contain;
            transition: transform 0.3s ease;
            filter: drop-shadow(1px 1px 1px rgba(0, 0, 0, 0.3));
        }

        .btn:hover .btn-icon img {
            transform: scale(1.1);
        }

        .btn::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg,
                    transparent,
                    rgba(255, 255, 255, 0.6),
                    transparent);
            transition: left 0.6s ease;
        }

        .btn:hover::before {
            left: 100%;
        }

        .btn:hover {
            transform: translate(-3px, -3px);
            box-shadow:
                6px 6px 0 var(--comic-ink),
                0 12px 25px rgba(0, 0, 0, 0.2);
        }

        .btn:active {
            transform: translate(1px, 1px);
            box-shadow:
                2px 2px 0 var(--comic-ink),
                0 4px 15px rgba(0, 0, 0, 0.1);
        }

        .btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            transform: none !important;
            box-shadow:
                4px 4px 0 var(--comic-ink),
                0 8px 20px rgba(0, 0, 0, 0.15) !important;
        }

        .btn-docs {
            background: linear-gradient(145deg, #4caf50, #43a047);
            border-color: #2e7d32;
            position: relative;
        }

        .btn-docs::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(135deg,
                    rgba(255, 255, 255, 0.2) 0%,
                    transparent 50%,
                    rgba(0, 0, 0, 0.1) 100%);
            border-radius: 13px;
            pointer-events: none;
        }

        .btn-preview {
            background: linear-gradient(145deg, #2196f3, #1976d2);
            border-color: #1565c0;
            position: relative;
        }

        .btn-preview::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(135deg,
                    rgba(255, 255, 255, 0.2) 0%,
                    transparent 50%,
                    rgba(0, 0, 0, 0.1) 100%);
            border-radius: 13px;
            pointer-events: none;
        }

        .btn-complete {
            background: linear-gradient(145deg, #ff9800, #f57c00);
            border-color: #ef6c00;
            position: relative;
        }

        .btn-complete::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(135deg,
                    rgba(255, 255, 255, 0.2) 0%,
                    transparent 50%,
                    rgba(0, 0, 0, 0.1) 100%);
            border-radius: 13px;
            pointer-events: none;
        }

        .btn-docs .btn-title,
        .btn-docs .btn-desc,
        .btn-preview .btn-title,
        .btn-preview .btn-desc,
        .btn-complete .btn-title,
        .btn-complete .btn-desc {
            color: white;
            text-shadow: 2px 2px 0 rgba(0, 0, 0, 0.3);
        }

        .status {
            margin-top: 20px;
            padding: 16px;
            border-radius: var(--comic-radius);
            display: none;
            font-size: 13px;
            font-weight: 600;
            transition: var(--transition);
            animation: fadeIn 0.3s ease;
            border-left: 4px solid;
            box-shadow: var(--comic-shadow);
            border: 2px solid var(--comic-ink);
        }

        @keyframes fadeIn {
            from {
                opacity: 0;
                transform: translateY(10px);
            }

            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        .status.info {
            display: block;
            background: #e3f2fd;
            color: #1565c0;
            border-left-color: #2196f3;
        }

        .status.success {
            display: block;
            background: #e8f5e9;
            color: #2e7d32;
            border-left-color: #4caf50;
        }

        .status.warning {
            display: block;
            background: #fff3e0;
            color: #ef6c00;
            border-left-color: #ff9800;
        }

        .status.error {
            display: block;
            background: #ffebee;
            color: #c62828;
            border-left-color: #f44336;
        }

        .pulse {
            animation: pulse 1.5s infinite;
        }

        @keyframes pulse {
            0% {
                transform: scale(1);
            }

            50% {
                transform: scale(1.05);
            }

            100% {
                transform: scale(1);
            }
        }

        /* Efecto Ripple para botones */
        .ripple {
            position: absolute;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.7);
            transform: scale(0);
            animation: ripple-animation 0.6s linear;
            pointer-events: none;
        }

        @keyframes ripple-animation {
            to {
                transform: scale(4);
                opacity: 0;
            }
        }

        /* Spinner de carga mejorado */
        .loading {
            position: relative;
            color: transparent !important;
            background: linear-gradient(145deg,
                    rgba(0, 0, 0, 0.8),
                    rgba(0, 0, 0, 0.9)) !important;
            border-color: rgba(0, 0, 0, 0.95) !important;
            transform: none !important;
            box-shadow:
                2px 2px 0 rgba(0, 0, 0, 0.7),
                inset 0 0 30px rgba(0, 0, 0, 0.8) !important;
            transition: all 0.3s ease !important;
            pointer-events: none !important;
            cursor: not-allowed !important;
            opacity: 0.9 !important;
        }

        .loading .btn-content,
        .loading .btn-title,
        .loading .btn-desc {
            color: rgba(255, 255, 255, 0.3) !important;
            text-shadow: none !important;
        }

        .loading .btn-icon {
            background: rgba(0, 0, 0, 0.6) !important;
            border-color: rgba(0, 0, 0, 0.8) !important;
            box-shadow: none !important;
            opacity: 0.4 !important;
        }

        .loading::after {
            content: '';
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 32px;
            height: 32px;
            border: 4px solid rgba(255, 255, 255, 0.1);
            border-top: 5px solid #ffffff;
            border-radius: 50%;
            animation: spin 0.7s linear infinite;
            z-index: 100;
            filter: drop-shadow(0 0 3px rgba(255, 255, 255, 0.5));
        }

        @keyframes spin {
            0% {
                transform: translate(-50%, -50%) rotate(0deg);
            }

            100% {
                transform: translate(-50%, -50%) rotate(360deg);
            }
        }

        .comic-dots {
            position: absolute;
            top: -5px;
            left: -5px;
            right: -5px;
            bottom: -5px;
            border: 2px dashed rgba(0, 0, 0, 0.1);
            border-radius: calc(var(--comic-radius-lg) + 5px);
            pointer-events: none;
            z-index: -1;
        }

        .title-icon {
            width: 65px;
            height: 65px;
            background: var(--comic-paper);
            padding: 3px;
            border-radius: 5px;
            border: 1px solid var(--comic-ink);
            box-shadow: var(--comic-shadow);
        }

        .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 11px;
            color: #666;
            font-weight: 500;
            padding-top: 15px;
            border-top: 1px dashed rgba(0, 0, 0, 0.1);
        }
    </style>
</head>

<body>
    <div class="container">
        <div class="comic-dots"></div>
        <h2>
            <img src="https://cualli.mx/wp-content/uploads/2022/07/cualli-bl@3x.png" class="title-icon">
            C1C/Cuadro de Transferencias
        </h2>
        <div class="btn-group">
            <button class="btn btn-preview" onclick="ejecutarAccion('previsualizarCuadrosModal', this)">
                <div class="btn-icon">
                    <img src="https://cdn-icons-png.flaticon.com/512/6135/6135181.png" alt="Previsualizar">
                </div>
                <div class="btn-content">
                    <div class="btn-title">Previsualiza tu Cuadro</div>
                    <div class="btn-desc">Vista previa para copiar en email</div>
                </div>
            </button>
            <button class="btn btn-complete" onclick="ejecutarAccion('previsualizarDocumentoCompleto', this)">
                <div class="btn-icon">
                    <img src="https://cdn-icons-png.flaticon.com/512/1011/1011884.png" alt="Documento Completo">
                </div>
                <div class="btn-content">
                    <div class="btn-title">Previsualiza Completo</div>
                    <div class="btn-desc">Vista completa con autorización</div>
                </div>
            </button>
            <button class="btn btn-docs" onclick="ejecutarAccion('generarDOCS_NATIVO', this)">
                <div class="btn-icon">
                    <img src="https://cdn-icons-png.flaticon.com/512/7035/7035543.png" alt="Google Docs">
                </div>
                <div class="btn-content">
                    <div class="btn-title">Genera Cuadro Final</div>
                    <div class="btn-desc">Crear documento completo en Docs</div>
                </div>
            </button>
        </div>
        <div id="status" class="status"></div>
        <div class="footer">
            © ${new Date().getFullYear()} Financiera Cualli, todos los derechos reservados.
        </div>
    </div>
    <script>
        function createRipple(event) {
            const button = event.currentTarget;
            
            if (button.disabled) return;
            
            const circle = document.createElement("span");
            const diameter = Math.max(button.clientWidth, button.clientHeight);
            const radius = diameter / 2;

            circle.style.width  = diameter + 'px';
            circle.style.height = diameter + 'px';
            circle.style.left   = (event.clientX - button.getBoundingClientRect().left - radius) + 'px';
            circle.style.top    = (event.clientY - button.getBoundingClientRect().top  - radius) + 'px';
            circle.className = 'ripple';

               const ripple = button.getElementsByClassName("ripple")[0];
            if (ripple) {
                ripple.remove();
            }

            button.appendChild(circle);
        }

        document.querySelectorAll('.btn').forEach(button => {
            button.addEventListener('click', createRipple);
            
            button.addEventListener('mouseenter', function() {
                if (!this.disabled) {
                    this.style.transform = 'translate(-2px, -2px)';
                    this.style.boxShadow = '5px 5px 0 var(--comic-ink)';
                }
            });
            
            button.addEventListener('mouseleave', function() {
                if (!this.disabled && !this.classList.contains('loading')) {
                    this.style.transform = '';
                    this.style.boxShadow = '4px 4px 0 var(--comic-ink)';
                }
            });
        });

        function setStatus(kind, msg) {
            if (kind === 'error') {
                return; 
            }

            const el = document.getElementById('status');
            el.className = 'status ' + kind;
            el.textContent = msg || '';
            el.style.display = 'block';

            if (kind === 'success') {
                el.classList.add('pulse');
                setTimeout(() => el.classList.remove('pulse'), 1500);

                setTimeout(() => {
                    if (el.style.display !== 'none') {
                        el.style.display = 'none';
                    }
                }, 10000);
            }

            if (kind === 'info') {
                setTimeout(() => {
                    if (el.style.display !== 'none') {
                        el.style.display = 'none';
                    }
                }, 10000);
            }
        }

        function ejecutarAccion(accion, button) {
            button.disabled = true;
            button.classList.add('loading');
            setStatus('info', '⏳ Generando, espera...');

            google.script.run
                .withSuccessHandler(function(resultado) {
                    button.disabled = false;
                    button.classList.remove('loading');
                    setStatus('success', '✅ Se generó la acción exitosamente');
                })
                .withFailureHandler(function(error) {
                    button.disabled = false;
                    button.classList.remove('loading');
                    const statusEl = document.getElementById('status');
                    statusEl.style.display = 'none';
                })[accion]();
        }

        // Precargar imágenes
        window.addEventListener('load', function() {
            const imageUrls = [
                'https://cdn-icons-png.flaticon.com/512/7035/7035543.png',
                'https://cdn-icons-png.flaticon.com/512/6135/6135181.png',
                'https://cdn-icons-png.flaticon.com/512/1011/1011884.png'
            ];
            imageUrls.forEach(url => {
                const img = new Image();
                img.src = url;
            });
        });
    </script>
</body>

</html>

  `)
  .setTitle('🛠️ Herramientas Para Genera Cuadro')
  .setWidth(500);

  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Sincroniza la lista desplegable de IDs en B3 para las hojas Datos_1 y Datos_2
 */
function sincronizarListaOportunidadesDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shBD = ss.getSheetByName('BD');
  
  if (!shBD) {
    console.error('No se encontró la hoja "BD" para sincronizar la lista.');
    return;
  }

  // 1. Obtener los datos de la base de datos
  const datosBD = shBD.getDataRange().getValues();
  if (datosBD.length < 2) return; 

  const headers = datosBD[0];
  const idxID = headers.indexOf('Nombre de la oportunidad');
  
  if (idxID === -1) {
    console.error('No se encontró la columna "Nombre de la oportunidad" en la hoja BD');
    return;
  }

  // 2. Extraer IDs únicos y limpiar vacíos
  const listaIds = datosBD.slice(1)
    .map(fila => String(fila[idxID]).trim())
    .filter(id => id !== "");

  if (listaIds.length === 0) return;

  // 3. Crear la regla de validación
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(listaIds, true)
    .setAllowInvalid(false) 
    .setHelpText('Selecciona un ID Oportunidad válido de la lista.')
    .build();

  // 4. Aplicar a las hojas de Datos en la celda B3
  const hojasObjetivo = ['Datos_1', 'Datos_2'];
  
  hojasObjetivo.forEach(nombre => {
    let hoja = ss.getSheetByName(nombre);
    if (hoja) {
      hoja.getRange('B3').setDataValidation(regla);
    }
  });

  console.log('✅ Lista desplegable en B3 sincronizada para Datos_1 y Datos_2.');
}
