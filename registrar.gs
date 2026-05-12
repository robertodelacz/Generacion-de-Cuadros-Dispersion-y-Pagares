

function RegistrarC() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const hojaOrigen = HOJA_ACTIVA_CONFIG.obtenerHojaDatos();
  const hojaDestino = ss.getSheetByName("REGISTRO DISPERSION");

  // ==================== CONFIGURACIÓN ====================
  const CONFIG = {
    ORIG_HEADER_ROW: 28,
    ORIG_DATA_START_ROW: 29,
    DEST_MIN_START_ROW: 2,
    COLOR_A: "#f7e9b0",
    COLOR_B: "#d9d8d7",
    PRINC_COLS_COUNT: 12,
    TOTAL_COLS: 21,
    ORIG_DETALLE_COLS: 10
  };

  const COLUMNAS_PRINCIPALES = [
    { origen: "B3", destino: 1 },
    { origen: "B5", destino: 2 },
    { origen: "B7", destino: 3 },
    { origen: "B23", destino: 4, defaultValue: "N/A" },
    { origen: "B25", destino: 5, defaultValue: "N/A" },
    { origen: "B9", destino: 6 },
    { origen: "B11", destino: 7 },
    { origen: "B17", destino: 8 },
    { origen: "B19", destino: 9 },
    { origen: "E5", destino: 10 },
    { origen: "E7", destino: 11 },
    { origen: "E9", destino: 12 }
  ];

  const COLUMNAS_DETALLE = [
    { origen: { start: 0, length: 8 }, destino: 13 },
    { origen: { start: 8, length: 1 }, destino: 21 }
  ];

  // ==================== FUNCIONES AUXILIARES ====================
  const obtenerColorAlternado = (hoja, filaInicio) => {
    if (filaInicio <= CONFIG.DEST_MIN_START_ROW) return CONFIG.COLOR_A;
    const colorAnterior = hoja.getRange(filaInicio - 1, 1).getBackground();
    return colorAnterior === CONFIG.COLOR_A ? CONFIG.COLOR_B : CONFIG.COLOR_A;
  };

  const getNextEmptyRow = (hoja) => {
    const lastRow = hoja.getLastRow();
    if (lastRow < CONFIG.DEST_MIN_START_ROW) return CONFIG.DEST_MIN_START_ROW;

    const numRows = lastRow - CONFIG.DEST_MIN_START_ROW + 1;
    const values = hoja.getRange(CONFIG.DEST_MIN_START_ROW, 1, numRows, CONFIG.TOTAL_COLS).getValues();

    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i].some(v => String(v).trim())) {
        return CONFIG.DEST_MIN_START_ROW + i + 1;
      }
    }
    return CONFIG.DEST_MIN_START_ROW;
  };

  const obtenerDatosDestino = () => {
    const lastRow = hojaDestino.getLastRow();
    if (lastRow < CONFIG.DEST_MIN_START_ROW) return { ids: [], colores: [] };

    const numRows = lastRow - CONFIG.DEST_MIN_START_ROW + 1;
    const ids = hojaDestino.getRange(CONFIG.DEST_MIN_START_ROW, 1, numRows, 1).getValues().flat();
    const colores = hojaDestino.getRange(CONFIG.DEST_MIN_START_ROW, 1, numRows, 1).getBackgrounds().flat();

    return { ids, colores };
  };

  const encontrarBloqueExistente = (idOportunidad, ids) => {
    const startIndex = ids.findIndex(id => id === idOportunidad);
    if (startIndex === -1) return null;

    let rowCount = 1;
    for (let j = startIndex + 1; j < ids.length; j++) {
      if (ids[j] === "–" || ids[j] === idOportunidad) {
        rowCount++;
      } else {
        break;
      }
    }

    return {
      startRow: CONFIG.DEST_MIN_START_ROW + startIndex,
      rowCount
    };
  };

  const leerDatosPrincipales = () => {
    const datos = Array(CONFIG.PRINC_COLS_COUNT).fill("");

    COLUMNAS_PRINCIPALES.forEach(col => {
      let valor = hojaOrigen.getRange(col.origen).getValue();
      if (col.defaultValue && !String(valor).trim()) {
        valor = col.defaultValue;
      }
      datos[col.destino - 1] = valor;
    });

    return datos;
  };

  const leerFilasDetalle = () => {
    const e3Valor = Number(hojaOrigen.getRange("E3").getValue()) || 0;
    const lastRowOrigen = hojaOrigen.getLastRow();
    const maxFilasPosibles = Math.max(0, lastRowOrigen - CONFIG.ORIG_HEADER_ROW);

    if (maxFilasPosibles === 0) return [];

    const raw = hojaOrigen
      .getRange(CONFIG.ORIG_DATA_START_ROW, 1, maxFilasPosibles, CONFIG.ORIG_DETALLE_COLS)
      .getValues();

    let filasDetalle;
    if (e3Valor > 0) {
      filasDetalle = raw.slice(0, Math.min(e3Valor, raw.length));
    } else {
      filasDetalle = [];
      for (const fila of raw) {
        if (!String(fila[0]).trim()) break;
        filasDetalle.push(fila);
      }
    }

    return filasDetalle.filter(r => r.join("").trim());
  };

  const ajustarFilasBloque = (filaInicio, filasActuales, filasNuevas) => {
    const diferencia = filasNuevas - filasActuales;

    if (diferencia > 0) {
      hojaDestino.insertRowsAfter(filaInicio + filasActuales - 1, diferencia);
    } else if (diferencia < 0) {
      for (let i = 0; i < Math.abs(diferencia); i++) {
        hojaDestino.deleteRow(filaInicio + filasNuevas);
      }
    }
  };

  const escribirDatosEnDestino = (filaInicio, datosPrincipales, filasDetalle, idOportunidad) => {
    const nFilasDetalle = filasDetalle.length;
    const totalFilas = Math.max(1, nFilasDetalle);

    // Limpiar rango destino
    hojaDestino.getRange(filaInicio, 1, totalFilas, CONFIG.TOTAL_COLS).clearContent();

    // Escribir datos principales
    hojaDestino.getRange(filaInicio, 1, 1, CONFIG.PRINC_COLS_COUNT).setValues([datosPrincipales]);

    if (nFilasDetalle > 0) {
      // Escribir columnas de detalle
      COLUMNAS_DETALLE.forEach(config => {
        const detalleData = filasDetalle.map(r =>
          r.slice(config.origen.start, config.origen.start + config.origen.length)
        );
        hojaDestino.getRange(filaInicio, config.destino, nFilasDetalle, config.origen.length).setValues(detalleData);
      });

      // Escribir guiones en filas adicionales
      if (nFilasDetalle > 1) {
        const guiones = Array.from({ length: nFilasDetalle - 1 }, () => {
          const row = Array(CONFIG.PRINC_COLS_COUNT).fill("–");
          row[0] = idOportunidad;
          return row;
        });
        hojaDestino.getRange(filaInicio + 1, 1, nFilasDetalle - 1, CONFIG.PRINC_COLS_COUNT).setValues(guiones);
      }
    }

    return totalFilas;
  };

  // ==================== LÓGICA PRINCIPAL ====================
  try {
    const datosPrincipales = leerDatosPrincipales();
    const idOportunidad = datosPrincipales[0];

    // Validación
    if (!idOportunidad) {
      ui.alert(
        "No se puede registrar",
        "El 'ID Oportunidad' (B3) está vacío.",
        ui.ButtonSet.OK
      );
      return;
    }

    const filasDetalle = leerFilasDetalle();
    const totalFilasBloque = Math.max(1, filasDetalle.length);

    // Obtener datos existentes en destino (una sola lectura)
    const { ids, colores } = obtenerDatosDestino();
    const bloqueExistente = encontrarBloqueExistente(idOportunidad, ids);

    let filaInicio;
    let colorBase;
    let esActualizacion = false;

    if (bloqueExistente) {
      const resp = ui.alert(
        "Información existente",
        `Ya existe un bloque para el ID Oportunidad ${idOportunidad}. ¿Deseas reemplazarlo?`,
        ui.ButtonSet.YES_NO
      );
      if (resp === ui.Button.NO) return;

      filaInicio = bloqueExistente.startRow;
      colorBase = colores[bloqueExistente.startRow - CONFIG.DEST_MIN_START_ROW];
      esActualizacion = true;

      ajustarFilasBloque(filaInicio, bloqueExistente.rowCount, totalFilasBloque);
    } else {
      filaInicio = getNextEmptyRow(hojaDestino);
      colorBase = obtenerColorAlternado(hojaDestino, filaInicio);
    }

    const filasEscritas = escribirDatosEnDestino(filaInicio, datosPrincipales, filasDetalle, idOportunidad);

    // Aplicar color al bloque
    hojaDestino.getRange(filaInicio, 1, filasEscritas, CONFIG.TOTAL_COLS).setBackground(colorBase);

    // Notificación de éxito
    const accion = esActualizacion ? "actualizado" : "capturado";
    SpreadsheetApp.getActive().toast(
      `✅ Registro ${accion} con éxito - ${filasEscritas} fila(s)`,
      "Completado",
      4
    );

  } catch (error) {
    console.error("Error en RegistrarC:", error);
    ui.alert(
      "Error inesperado",
      `Ocurrió un error al procesar el registro:\n${error.message}`,
      ui.ButtonSet.OK
    );
  }
}
