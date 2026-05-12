function limpiarfilasDispersion() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
   const hoja = HOJA_ACTIVA_CONFIG.obtenerHojaDatos();

   if (!hoja) {
    SpreadsheetApp.getUi().alert('⚠️ Por favor, selecciona la hoja DATOS_1 o DATOS_2 antes de limpiar');
    return;
  }

  const FILA_INICIO = 29;           
  const CELDAS_ADICIONALES = ['B5','B7', 'B9', 'B11', 'B13', 'B15', 'B17', 'B19', 'E7', 'E11', 'E9', 'B21', 'B23', 'B25', 'H5', 'H7', 'H9', 'H11', 'H13', 'H15', 'H21', 'H23', 'J6', 'J8', 'J10', 'J12', 'J14', 'J20'];
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < FILA_INICIO) {
    if (CELDAS_ADICIONALES.length) {
      hoja.getRangeList(CELDAS_ADICIONALES).clearContent();
    }
    aplicarMenusDesplegables(hoja, FILA_INICIO, 1);
    SpreadsheetApp.getActive().toast("No hay datos que limpiar", "Limpieza", 3);
    return;
  }

  const numFilas = (ultimaFila - FILA_INICIO) + 1;

  hoja.getRange(FILA_INICIO, 1, numFilas, 4).clearContent();
  hoja.getRange(FILA_INICIO, 6, numFilas, 3).clearContent();
  hoja.getRange(FILA_INICIO, 9, numFilas, 1).clearContent();

  if (CELDAS_ADICIONALES.length) hoja.getRangeList(CELDAS_ADICIONALES).clearContent();

  hoja.getRange(FILA_INICIO, 1, numFilas, 4).setBackground('#ffffff');
  hoja.getRange(FILA_INICIO, 6, numFilas, 3).setBackground('#ffffff');
  hoja.getRange(FILA_INICIO, 9, numFilas, 1).setBackground('#ffffff');

 
  SpreadsheetApp.getActive().toast("🧹 ¡Limpieza exitosa!", "Limpieza", 3);
}

