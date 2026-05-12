function BuscarC() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = ss.getSheetByName("REGISTRO DISPERSION");
  const hojaDestino = HOJA_ACTIVA_CONFIG.obtenerHojaDatos();

  const idOportunidad = hojaDestino.getRange("B3").getValue();

  if (!idOportunidad) {
    SpreadsheetApp.getUi().alert("Por favor, ingrese un ID Oportunidad en la celda B3.");
    return;
  }

  const CONFIG_GENERAL = {
    ORIG_HEADER_ROW: 28,
    ORIG_DATA_START_ROW: 29,
    DEST_MIN_START_ROW: 2,
    MAX_FILAS_HOJA: 50 
  };

  const COLUMNAS_PRINCIPALES_INVERSO = [
    { origen: 1, destino: "B3" },
    { origen: 2, destino: "B5" },
    { origen: 3, destino: "B7" },
    { origen: 4, destino: "B23" },
    { origen: 5, destino: "B25" },
    { origen: 6, destino: "B9" },
    { origen: 7, destino: "B11" },
    { origen: 8, destino: "B17" },
    { origen: 9, destino: "B19" },
    { origen: 10, destino: "E5" },
    { origen: 11, destino: "E7" },
    { origen: 12, destino: "E9" }
  ];

  eliminarFilasExcedentes(hojaDestino, CONFIG_GENERAL.MAX_FILAS_HOJA);

  const ultimaFilaRealOrigen = hojaOrigen.getLastRow();
  const datosOrigen = ultimaFilaRealOrigen > 0 ? 
    hojaOrigen.getRange(1, 1, ultimaFilaRealOrigen, hojaOrigen.getLastColumn()).getValues() : [];

  const filasEncontradas = datosOrigen.filter(row => row[0] == idOportunidad);

  if (filasEncontradas.length === 0) {
    SpreadsheetApp.getUi().alert(`No se encontró información para el ID Oportunidad: ${idOportunidad}`);
    return;
  }

  const filaPrincipal = filasEncontradas[0];
  const datosPrincipales = filaPrincipal.slice(0, 14); 

  COLUMNAS_PRINCIPALES_INVERSO.forEach(config => {
    if (config.origen <= datosPrincipales.length) {
      const valor = datosPrincipales[config.origen - 1]; 
      hojaDestino.getRange(config.destino).setValue(valor);
    }
  });

  const filasALimpiar = 50; 
  hojaDestino.getRange(CONFIG_GENERAL.ORIG_DATA_START_ROW, 1, filasALimpiar, 4).clearContent();
  hojaDestino.getRange(CONFIG_GENERAL.ORIG_DATA_START_ROW, 6, filasALimpiar, 4).clearContent();

  if (filasEncontradas.length > 0) {
    const datosMP = filasEncontradas.map(fila => fila.slice(12, 16)); 
    const datosRU = filasEncontradas.map(fila => fila.slice(17, 21)); 

    const datosMPLimitados = datosMP.slice(0, 20);
    const datosRULimitados = datosRU.slice(0, 20);

    if (datosMPLimitados.length > 0 && datosMPLimitados[0].length > 0) {
      hojaDestino.getRange(29, 1, datosMPLimitados.length, datosMPLimitados[0].length)
        .setValues(datosMPLimitados);
    }
    
    if (datosRULimitados.length > 0 && datosRULimitados[0].length > 0) {
      hojaDestino.getRange(29, 6, datosRULimitados.length, datosRULimitados[0].length)
        .setValues(datosRULimitados);
    }

    hojaDestino.getRange("E3").setValue(filasEncontradas.length);
  }
  
  aplicarFormatosEspeciales(hojaDestino);
  eliminarFilasExcedentes(hojaDestino, CONFIG_GENERAL.MAX_FILAS_HOJA);

  // ⭐ FORZAR RECÁLCULO tocando solo celdas D con datos
  forzarRecalculoColumnaDConDatos(hojaDestino, filasEncontradas.length);

  try {
    if (typeof cargarDatosDesdeBase === 'function') {
      PropertiesService.getScriptProperties().setProperty('SUPRIMIR_TOAST_BD', 'true');
      cargarDatosDesdeBase();
      PropertiesService.getScriptProperties().deleteProperty('SUPRIMIR_TOAST_BD');
    }
  } catch (error) {
    PropertiesService.getScriptProperties().deleteProperty('SUPRIMIR_TOAST_BD');
    console.log("Error cargando BD:", error);
  }
}

// ⭐ NUEVA FUNCIÓN: Solo toca las celdas D con datos
function forzarRecalculoColumnaDConDatos(hoja, numFilas) {
  try {
    if (numFilas <= 0) return;
    
    // Leer columna D desde fila 29
    const filaInicio = 29;
    const cantidadFilas = Math.min(numFilas, 20); // Limitar a 20 filas máximo
    const rangoD = hoja.getRange(filaInicio, 4, cantidadFilas, 1); // Columna D (índice 4)
    const valoresD = rangoD.getValues();
    
    // Filtrar solo las celdas con datos
    let celdasConDatos = 0;
    for (let i = 0; i < valoresD.length; i++) {
      const valor = valoresD[i][0];
      
      // Solo procesar si tiene un valor válido (no vacío, no null, no undefined)
      if (valor !== null && valor !== undefined && valor !== "" && String(valor).trim() !== "") {
        celdasConDatos++;
      }
    }
    
    if (celdasConDatos === 0) {
      console.log('ℹ️ No hay datos en columna D para recalcular');
      return;
    }
    
    // Re-escribir los mismos valores para forzar recálculo
    rangoD.setValues(valoresD);
    SpreadsheetApp.flush();
    
    console.log(`✅ Recalculadas ${celdasConDatos} celdas con datos en columna D`);
    
  } catch (error) {
    console.log('⚠️ Error forzando recálculo en columna D:', error);
  }
}

function eliminarFilasExcedentes(hoja, maxFilas) {
  const filasActuales = hoja.getMaxRows();
  
  if (filasActuales > maxFilas) {
    const filasAEliminar = filasActuales - maxFilas;
    hoja.deleteRows(maxFilas + 1, filasAEliminar);
    console.log(`🗑️ Eliminadas ${filasAEliminar} filas excedentes. Total: ${maxFilas} filas`);
  }
}

function aplicarFormatosEspeciales(hoja) {
  try {
    const celdaPorcentaje = hoja.getRange("E7");
    const valorPorcentaje = celdaPorcentaje.getValue();
    
    if (typeof valorPorcentaje === 'number' && isFinite(valorPorcentaje)) {
      celdaPorcentaje.setNumberFormat('0.00%');
    }

    hoja.getRange("E9").setNumberFormat('$#,##0.00');
    hoja.getRange("E11").setNumberFormat('$#,##0.00');

    const rangoFechas = hoja.getRange("B26:B100");
    rangoFechas.setNumberFormat('dd/mm/yyyy');

  } catch (err) {
    console.log("Error aplicando formatos:", err);
  }
}
