const IVA_MX = 0.16;                
const CURRENCY_FORMAT = '$#,##0.00';
const TARGET_SHEET   = ['DATOS_1', 'Datos_1', 'DATOS_2', 'Datos_2'];
const CELL_TIPO      = 'E5';   
const CELL_PORC      = 'E7';   
const CELL_BASE      = 'E9';   
const CELL_RESULTADO = 'F11';  
const CELL_SALIDA    = 'E11';  
const CELL_COPIA     = 'D29';  

const REPL_SHEET      = 'DATOS';  
const REPL_START_ROW  = 29;      
const REPL_COL_CONS   = 1;     
const REPL_COL_FECHA  = 2;       
const REPL_INPUT_CELL = 'B21';      
const REPL_FECHA_CELL = 'B19';     
const REPL_DATE_FORMAT= 'dd/mm/yyyy';

const FECHA_CONFIG = {
  SHEET_NAME: 'DATOS',
  CHECKBOX_CELL: 'I13',        
  FECHA_FIRMA_CELL: 'B19',    
  PLAZO_MESES_CELL: 'H11', 
  RESULTADO_CELL: 'H13'       
};

// Porcentajes
const PERCENT_FORMAT = '0.00%';
const PERCENT_CELLS = ['E7']; 
const SHOW_TOASTS = true;

function onEdit(e) {
  onEditHandler(e);
}

function onEditHandler(e) {
  try {
    if (!e || !e.range) return;

    handleFormatoB3_(e);
    handleFormatoPorcentaje_(e);
    handleReplicarConsecutivosFechas_(e);
    handleCalculoComisionCheckbox_(e);
    handleClabeDesdeCatalogo_(e);
    handleCalculoFechaVencimiento_(e);
    handleFormatoAllText_(e);
    handleRestaurarFormatoPagares_(e);

  } catch (err) {
    SpreadsheetApp.getActive().toast('Error en onEditHandler: ' + err, 'Error', 5);
  }
}

function handleFormatoPorcentaje_(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const sheetName = (sh.getName() || '').trim();
    if (!HOJA_ACTIVA_CONFIG.esHojaDatos(sheetName)) return;

    const editedA1 = e.range.getA1Notation();
    if (!PERCENT_CELLS.includes(editedA1)) return;

    const cell = sh.getRange(editedA1);
    const raw = cell.getValue();
    const disp = cell.getDisplayValue();

    const processed = procesarEntradaPorcentaje_(raw !== null && raw !== undefined ? raw : disp);
    if (processed !== null && isFinite(processed)) {
      cell.setValue(processed);
      cell.setNumberFormat(PERCENT_FORMAT);
      if (SHOW_TOASTS) toast_('✅ Porcentaje: ' + (processed * 100).toFixed(2) + '%', 'Formato', 2);
    }
  } catch (err) {
    console.error('handleFormatoPorcentaje_ error:', err);
  }
}

function procesarEntradaPorcentaje_(input) {
  if (input === null || input === undefined) return null;
  
  if (typeof input === 'number' && isFinite(input)) {
    if (input >= 0 && input <= 1) return input;
    if (input > 1 && input <= 100) return input / 100;
    if (input > 100) return input / 100;
    return input;
  }

  let s = String(input).trim();
  if (s === '') return null;

  const hadPercent = /%/.test(s);
  s = s.replace(/\s+/g, '').replace(/,/g, '.');
  const cleaned = s.replace(/[^0-9.\-]/g, '');
  if (cleaned === '' || cleaned === '-' || cleaned === '.') return null;

  const n = parseFloat(cleaned);
  if (!isFinite(n)) return null;

  if (!hadPercent && n <= 100) {
    return n <= 1 ? n : n / 100;
  }
  
  return n / 100;
}

function handleCalculoComisionCheckbox_(e) {
  if (!e || !e.range) return;
  
  const sh = e.range.getSheet();
  const sheetName = (sh.getName() || '').toLowerCase().trim();
  if (!HOJA_ACTIVA_CONFIG.esHojaDatos(sheetName)) return;

  
  const editedCell = e.range.getA1Notation();
  if (editedCell !== CELL_RESULTADO) return;
  
  const isChecked = e.value === true || 
                   (typeof e.value === 'string' && e.value.toUpperCase() === 'TRUE') ||
                   sh.getRange(CELL_RESULTADO).getValue() === true;
  
  const rOut = sh.getRange(CELL_SALIDA);
  
  if (!isChecked) {
    rOut.clearContent();
    sh.getRange(CELL_COPIA).clearContent();
    return;
  }
  
  try {
    const rangeValues = sh.getRange('E7:E9').getValues();
    const porcVal = rangeValues[0][0];
    const baseVal = rangeValues[2][0];
    
    const pct = procesarEntradaPorcentaje_(porcVal) || 0;
    const base = parseNumberOptimizado_(baseVal, sh.getRange(CELL_BASE).getDisplayValue());
    
    if (!isFinite(base) || base <= 0 || !isFinite(pct) || pct <= 0) {
      rOut.setNumberFormat('@').setValue('⚠️ Revisa E7/E9');
      sh.getRange(CELL_RESULTADO).setValue(false);
      return;
    }
    
   const total = Math.round(base * pct * (1 + IVA_MX) * 100) / 100;
    
    
    rOut.setNumberFormat(CURRENCY_FORMAT).setValue(total);
    sh.getRange(CELL_COPIA).setNumberFormat(CURRENCY_FORMAT).setValue(total);
    
    sh.getRange(CELL_RESULTADO).setValue(false);
    
    if (SHOW_TOASTS) {
      SpreadsheetApp.getActive().toast('✅ Comisión calculada', '', 2);
    }
    
  } catch (err) {
    rOut.setNumberFormat('@').setValue('Error en cálculo');
    sh.getRange(CELL_RESULTADO).setValue(false);
  }
}

function parseNumberOptimizado_(rawVal, dispVal) {
  if (typeof rawVal === 'number' && isFinite(rawVal)) return rawVal;
  
  let str = String(dispVal || '').trim();
  if (!str) return 0;
  
  str = str.replace(/[^\d,-.]/g, '');
  const isNegative = str.startsWith('(') && str.endsWith(')');
  
  if (isNegative) {
    str = '-' + str.replace(/[()]/g, '');
  }
  
  // Manejo decimal simple
  const parts = str.split(/[.,]/);
  if (parts.length > 1) {
    str = parts.slice(0, -1).join('') + '.' + parts[parts.length - 1];
  }
  
  return parseFloat(str) || 0;
}

function handleReplicarConsecutivosFechas_(e) {
  const sh = e.range.getSheet();
 if (!HOJA_ACTIVA_CONFIG.esHojaDatos(sh.getName())) return;

  const a1 = e.range.getA1Notation();
  if (a1 !== REPL_INPUT_CELL && a1 !== REPL_FECHA_CELL) return;

  const cantidad = parseInt(sh.getRange(REPL_INPUT_CELL).getValue(), 10);
  const fecha    = sh.getRange(REPL_FECHA_CELL).getValue();

  const maxRows   = sh.getMaxRows();
  const totalRows = Math.max(0, maxRows - REPL_START_ROW + 1);
  if (totalRows > 0) {
    sh.getRange(REPL_START_ROW, REPL_COL_CONS, totalRows, 1).clearContent();
    sh.getRange(REPL_START_ROW, REPL_COL_FECHA, totalRows, 1).clearContent();
  }
  
  if (!isNaN(cantidad) && cantidad > 0) {
    const consecutivos = Array.from({ length: cantidad }, (_, i) => [i + 1]);
    const fechas       = Array.from({ length: cantidad }, () => [fecha]);

    sh.getRange(REPL_START_ROW, REPL_COL_CONS, cantidad, 1)
      .setValues(consecutivos)
      .setHorizontalAlignment('right');

    sh.getRange(REPL_START_ROW, REPL_COL_FECHA, cantidad, 1)
      .setValues(fechas)
      .setNumberFormat(REPL_DATE_FORMAT)
      .setHorizontalAlignment('center');

    SpreadsheetApp.getActive().toast(`📋 ${cantidad} registros generados`, 'Consecutivos', 3);
  }
}

function aplicarFormatosAlAbrir() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojas = ['DATOS_1', 'Datos_1', 'Datos_2', 'DATOS_2'];
    
    hojas.forEach(nombreHoja => {
      const hoja = ss.getSheetByName(nombreHoja);
      if (hoja) {
        PERCENT_CELLS.forEach(celda => {
          const rango = hoja.getRange(celda);
          const valor = rango.getValue();
          
          if (typeof valor === 'number' && isFinite(valor)) {
            rango.setNumberFormat(PERCENT_FORMAT);
            
            if (valor > 1) {
              rango.setValue(valor / 100);
            }
          }
        });
      }
    });
    
    if (SHOW_TOASTS) {
      toast_('✅ Formatos de porcentaje verificados', 'Sistema', 3);
    }
  } catch (err) {
    console.error('Error en aplicarFormatosAlAbrir:', err);
  }
}

function toast_(msg, title, secs) {
  try {
    const m = String(msg == null ? '' : msg);
    const t = String(title == null ? '' : title); 
    const s = Math.max(1, Math.min(Number(secs == null ? 2 : secs), 10));
    SpreadsheetApp.getActive().toast(m, t, s);
  } catch (_) {}
}

function handleFormatoB3_(e) {
  const sh = e.range.getSheet();
  if (!sh || !HOJA_ACTIVA_CONFIG.esHojaDatos(sh.getName())) return false;


  const r = e.range;
  const incluyeB3 =
    r.getRow() <= 3 && 3 <= r.getLastRow() &&
    r.getColumn() <= 2 && 2 <= r.getLastColumn();

  if (!incluyeB3) return false;

  const b3 = sh.getRange('B3');


  const val = b3.getDisplayValue();
  b3.setNumberFormat('@');        
  b3.setValue(val);             


  b3.setFontFamily('Arial')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')   
    .setVerticalAlignment('middle')
    .setBackground('white');



  return true;
}

function handleCalculoFechaVencimiento_(e) {
  if (!e || !e.range) return;
  
  const sh = e.range.getSheet();
  const sheetName = (sh.getName() || '').toLowerCase().trim();
  if (!HOJA_ACTIVA_CONFIG.esHojaDatos(sheetName)) return;
  
  const editedCell = e.range.getA1Notation();
  if (editedCell !== FECHA_CONFIG.CHECKBOX_CELL) return;
  
  const isChecked = e.value === true || 
                   (typeof e.value === 'string' && e.value.toUpperCase() === 'TRUE') ||
                   sh.getRange(FECHA_CONFIG.CHECKBOX_CELL).getValue() === true;
  
  if (!isChecked) return; 
  
  try {
    const resultado = calcularFechaVencimiento();
    
    sh.getRange(FECHA_CONFIG.CHECKBOX_CELL).setValue(false);
    
    if (SHOW_TOASTS) {
      SpreadsheetApp.getActive().toast(`✅ Fecha calculada: ${resultado}`, 'Cálculo Completado', 4);
    }
    
  } catch (error) {
    console.error('Error en handleCalculoFechaVencimiento_:', error);
    sh.getRange(FECHA_CONFIG.CHECKBOX_CELL).setValue(false);
    
    SpreadsheetApp.getActive().toast('❌ Error calculando fecha', 'Error', 4);
  }
}


function handleFormatoAllText_(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (!sh) return;

    if (!HOJA_ACTIVA_CONFIG.esHojaDatos(sh.getName())) return;

    const edited = e.range;
    const editStartRow = edited.getRow();
    const editEndRow   = edited.getLastRow();
    const editStartCol = edited.getColumn();
    const editEndCol   = edited.getLastColumn();

    const MIN_ROW = 29;
    const MIN_COL = 3;  // C
    const MAX_COL = 9;  // I
    const COLUMNAS_EXCLUIDAS = [4, 5];

    const startRow = Math.max(MIN_ROW, editStartRow);
    const endRow   = Math.max(MIN_ROW, editEndRow); 
    const startCol = Math.max(MIN_COL, editStartCol);
    const endCol   = Math.min(MAX_COL, editEndCol);

    const numRows = endRow - startRow + 1;

    if (endRow < startRow || startCol > endCol) {
      return;
    }

    const rangoIncluyeExcluidas = COLUMNAS_EXCLUIDAS.some(col => 
      col >= startCol && col <= endCol
    );

    if (rangoIncluyeExcluidas) {
      const rangosAFormatear = [];
      
      // Parte antes de las columnas excluidas (solo C)
      if (startCol <= 3 && endCol >= 3) {
        rangosAFormatear.push({
          startCol: 3,
          endCol: 3,
          numCols: 1
        });
      }
      
      // Parte después de las columnas excluidas (F-I)
      if (endCol >= 6) {
        const afterStartCol = Math.max(6, startCol);
        rangosAFormatear.push({
          startCol: afterStartCol,
          endCol: endCol,
          numCols: endCol - afterStartCol + 1
        });
      }
      rangosAFormatear.forEach(rango => {
        const targetRange = sh.getRange(
          startRow, 
          rango.startCol, 
          numRows, 
          rango.numCols
        );
        
        targetRange.clearFormat();                 
        targetRange.setNumberFormat('@');          
        targetRange.setFontFamily('Arial')
                   .setFontSize(12)
                   .setHorizontalAlignment('center')
                   .setVerticalAlignment('middle')
                   .setBackground('white')
                   .setWrap(true)
                   .setBorder(false, false, false, false, false, false);
      });
    } else {
      const numRows = endRow - startRow + 1;
      const numCols = endCol - startCol + 1;
      const targetRange = sh.getRange(startRow, startCol, numRows, numCols);

      targetRange.clearFormat();                 
      targetRange.setNumberFormat('@');          
      targetRange.setFontFamily('Arial')
                 .setFontSize(12)
                 .setHorizontalAlignment('center')
                 .setVerticalAlignment('middle')
                 .setBackground('white')
                 .setWrap(true)
                 .setBorder(false, false, false, false, false, false);
    }

    SpreadsheetApp.flush();

  } catch (error) {
    console.error('Error en handleFormatoAllText_:', error);
  }
}

