function cargarDatosDesdeBase() {

  
  const suprimirToast = PropertiesService.getScriptProperties().getProperty('SUPRIMIR_TOAST_BD') === 'true';

  const CONFIG = {
    HOJAS: {
      DATOS: 'Datos',
      BASE_DATOS: 'BD'
    },
    CELDAS: {
      NOMBRE_OPORTUNIDAD: 'B3',
      CAMPOS_MAPEO: {
        'Nombre de la cuenta': 'B5',
        'Régimen fiscal': 'B7',
        'Importe': ['B9', 'E9'],
        'Tipo': 'B15',
        'Porcentaje de comisión de Apertura': 'E7',
        'Tipo de gravamen': 'B11',
        'Producto': 'B13',
        'Importe de mediación': 'H5',
        'Costo estimado de Gastos Notariales': 'H7',
        'Tasa de interes Ordinaria Anual': 'H9',
        'Vigencia del Contrato (meses)': 'H11',
        'Banco': 'H21',
        'CLABE': 'H23'
      }
    },
    MENSAJES: {
      NOMBRE_VACIO: '⚠️ Escribe el "Nombre de la oportunidad" en B3.',
      SIN_DATOS: '⚠️ La hoja llamada BD no tiene datos.',
      COLUMNA_NO_ENCONTRADA: '⚠️ No se encontró la columna "Nombre de la oportunidad" en la hoja llamada BD.',
      OPORTUNIDAD_NO_ENCONTRADA: '❌ No se encontró coincidencia exacta en la base de datos.',
      COINCIDENCIA_SIMILAR: '⚠️ No hay coincidencia exacta, pero se encontró similar: "{similar}". Verifica el ID en B3.',
      MULTIPLES_COINCIDENCIAS: 'ℹ️ Hay {count} coincidencias exactas. Se usará la primera.',
      EXITO: '✅ Datos cargados correctamente.'
    },
    COLUMNAS: {
      NOMBRE_OPORTUNIDAD: 'Nombre de la oportunidad'
    }
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMenu = HOJA_ACTIVA_CONFIG.obtenerHojaDatos();
  const hojaBase = ss.getSheetByName(CONFIG.HOJAS.BASE_DATOS);

  // ⭐ CAMBIO: Usar _toast en lugar de _alert
  if (!hojaMenu) {
    _toast(`⚠️ No se encontró la hoja "${CONFIG.HOJAS.DATOS}"`, 'Error', 5);
    return;
  }
  if (!hojaBase) {
    _toast(`⚠️ No se encontró la hoja "${CONFIG.HOJAS.BASE_DATOS}"`, 'Error', 5);
    return;
  }

  const nombreBuscadoCrudo = String(hojaMenu.getRange(CONFIG.CELDAS.NOMBRE_OPORTUNIDAD).getValue() || '');
  const nombreBuscado = nombreBuscadoCrudo.trim();

  if (!nombreBuscado) {
    _toast(CONFIG.MENSAJES.NOMBRE_VACIO, 'Atención', 5);
    return;
  }

  const datos = hojaBase.getDataRange().getValues();
  if (datos.length < 2) {
    _toast(CONFIG.MENSAJES.SIN_DATOS, 'Error', 5);
    return;
  }

  const encabezados = datos[0].map(String);
  const filas = datos.slice(1);

  const idxNombreOportunidad = _findHeaderIndex(encabezados, CONFIG.COLUMNAS.NOMBRE_OPORTUNIDAD);
  if (idxNombreOportunidad === -1) {
    _toast(CONFIG.MENSAJES.COLUMNA_NO_ENCONTRADA, 'Error', 5);
    return;
  }

  const resultadoBusqueda = _buscarConValidacion(filas, idxNombreOportunidad, nombreBuscado);

  if (resultadoBusqueda.coincidenciasExactas.length === 0) {
    if (resultadoBusqueda.coincidenciaSimilar) {
      _toast(CONFIG.MENSAJES.COINCIDENCIA_SIMILAR.replace('{similar}', resultadoBusqueda.coincidenciaSimilar), 'Verificar ID', 8);
    } else {
      _toast(CONFIG.MENSAJES.OPORTUNIDAD_NO_ENCONTRADA, 'No encontrado', 5);
    }
    return;
  }

  if (resultadoBusqueda.coincidenciasExactas.length > 1) {
    _toast(CONFIG.MENSAJES.MULTIPLES_COINCIDENCIAS.replace('{count}', resultadoBusqueda.coincidenciasExactas.length), 'Múltiples coincidencias', 4);
  }

  const filaEncontrada = resultadoBusqueda.coincidenciasExactas[0].fila;

  Object.entries(CONFIG.CELDAS.CAMPOS_MAPEO).forEach(([headerBuscado, a1Destino]) => {
    const idx = _findHeaderIndex(encabezados, headerBuscado);
    if (idx !== -1) {
      const valor = filaEncontrada[idx];
      
      if (Array.isArray(a1Destino)) {
        a1Destino.forEach(celda => {
          hojaMenu.getRange(celda).setValue(valor);
        });
      } else {
        hojaMenu.getRange(a1Destino).setValue(valor);
      }
    }
  });

// ========== CARGAR DATOS DE CT (COLUMNAS G Y H) ==========
try {
  const hojaCT = ss.getSheetByName('CT');
  if (hojaCT) {
    const datosCT = hojaCT.getDataRange().getValues();
    const idBuscado = nombreBuscadoCrudo.trim();
    const columnaIdCT = 1;      // Columna B donde buscas el ID
    const columnaDatoG = 6;     // Columna G (índice 6)
    const columnaDatoH = 7;     // Columna H (índice 7)
    
    let datoG = null;
    let datoH = null;
    
    // Buscar en CT
    for (let i = 1; i < datosCT.length; i++) {
      const idEnCT = String(datosCT[i][columnaIdCT] || '').trim();
      if (idEnCT === idBuscado) {
        datoG = datosCT[i][columnaDatoG];
        datoH = datosCT[i][columnaDatoH];
        break;
      }
    }
    
    // Escribir dato G en hojaMenu H15
    if (datoG !== null && datoG !== undefined && datoG !== '') {
      hojaMenu.getRange('H15').setValue(datoG);
    }
    
    // Escribir dato H en hojaMenu H21 (la hoja activa actual)
    if (datoH !== null && datoH !== undefined && datoH !== '') {
      hojaMenu.getRange('H21').setValue(datoH);
      console.log(`✅ CT: Dato G="${datoG}" → H15, Dato H="${datoH}" → H21`);
    }
  }
} catch (error) {
  console.log('❌ Error en CT:', error.toString());
}

  // ========== BÚSQUEDA EN CONVENIOS ==========
  try {
    const hojaConvenios = ss.getSheetByName('CONVENIOS');
    
    if (hojaConvenios) {
      const valorB5Raw = hojaMenu.getRange('B5').getValue();
      const valorB5 = valorB5Raw == null ? '' : String(valorB5Raw).trim();
      
      if (valorB5) {
        const datosConvenios = hojaConvenios.getDataRange().getValues();
        
        if (datosConvenios && datosConvenios.length > 1) {
          const encabezadosConvenios = datosConvenios[0].map(h => 
            typeof h === 'string' ? h.trim().toLowerCase() : String(h).toLowerCase()
          );
          
          const idxNombreCuenta = encabezadosConvenios.indexOf('nombre de la cuenta');
          const idxFechaSolicitud = encabezadosConvenios.indexOf('fecha de solicitud de convenio');
          const idxNumConvenio = encabezadosConvenios.indexOf('convenio: número de convenio modificatorio');
          const idxMonto = encabezadosConvenios.indexOf('monto');
          const idxTasa = encabezadosConvenios.indexOf('tasa');
          const idxTipo = encabezadosConvenios.indexOf('tipo');
          const idxDescripcion = encabezadosConvenios.indexOf('descripción de convenio');
          
          if (idxNombreCuenta === -1 || idxFechaSolicitud === -1) {
            console.log('⚠️ No se encontraron las columnas requeridas en CONVENIOS');
          } else {
            const filasCoincidentes = datosConvenios.slice(1).filter(row => {
              const cell = row[idxNombreCuenta];
              return String(cell ?? '').trim() === valorB5;
            });
            
            if (filasCoincidentes.length > 0) {
              let filaMasReciente = filasCoincidentes[0];
              let fechaMasReciente = filasCoincidentes[0][idxFechaSolicitud];
              
              if (fechaMasReciente instanceof Date) {
                fechaMasReciente = fechaMasReciente.getTime();
              } else if (typeof fechaMasReciente === 'string') {
                fechaMasReciente = new Date(fechaMasReciente).getTime();
              } else {
                fechaMasReciente = 0;
              }
              
              for (let i = 1; i < filasCoincidentes.length; i++) {
                let fechaActual = filasCoincidentes[i][idxFechaSolicitud];
                
                if (fechaActual instanceof Date) {
                  fechaActual = fechaActual.getTime();
                } else if (typeof fechaActual === 'string') {
                  fechaActual = new Date(fechaActual).getTime();
                } else {
                  fechaActual = 0;
                }
                
                if (fechaActual > fechaMasReciente) {
                  fechaMasReciente = fechaActual;
                  filaMasReciente = filasCoincidentes[i];
                }
              }
              
              const mapeoConvenios = {
                'J6': idxNumConvenio !== -1 ? filaMasReciente[idxNumConvenio] : '',
                'J8': idxFechaSolicitud !== -1 ? filaMasReciente[idxFechaSolicitud] : '',
                'J10': idxMonto !== -1 ? filaMasReciente[idxMonto] : '',
                'J12': idxTasa !== -1 ? filaMasReciente[idxTasa] : '',
                'J14': idxTipo !== -1 ? filaMasReciente[idxTipo] : '',
                'J20': idxDescripcion !== -1 ? filaMasReciente[idxDescripcion] : ''
              };
              
              Object.entries(mapeoConvenios).forEach(([celda, valor]) => {
                hojaMenu.getRange(celda).setValue(valor ?? '');
              });
              
              console.log(`✅ Datos de CONVENIOS cargados correctamente (${filasCoincidentes.length} convenio(s) encontrado(s), se seleccionó el más reciente)`);
              
            } else {
              console.log('⚠️ No se encontraron convenios para: ' + valorB5);
              ['J6', 'J8', 'J10', 'J12', 'J14', 'J20'].forEach(celda => {
                hojaMenu.getRange(celda).setValue('');
              });
            }
          }
        }
      } else {
        console.log('⚠️ La celda B5 (Nombre de la cuenta) está vacía, no se buscará en CONVENIOS');
      }
    } else {
      console.log('⚠️ Hoja CONVENIOS no encontrada');
    }
  } catch (error) {
    console.log('Error en búsqueda de CONVENIOS:', error.toString());
  }  

  if (!suprimirToast) {
    _toast(CONFIG.MENSAJES.EXITO, 'Completado', 3);
  }
}

function _buscarConValidacion(filas, indiceColumna, textoBuscado) {
  const coincidenciasExactas = [];
  let mejorCoincidenciaSimilar = null;
  let mejorScore = 0;
  
  const textoBuscadoNormalizado = _normalizarEspacios(textoBuscado);
  
  for (let i = 0; i < filas.length; i++) {
    const valorCelda = String(filas[i][indiceColumna] || '').trim();
    if (!valorCelda) continue;
    
    const valorNormalizado = _normalizarEspacios(valorCelda);
    
    if (valorNormalizado === textoBuscadoNormalizado) {
      coincidenciasExactas.push({ idx: i, fila: filas[i] });
      continue;
    }
    
    const similitud = _calcularSimilitudEstricta(textoBuscadoNormalizado, valorNormalizado);
    if (similitud > 0.80 && similitud > mejorScore) {
      mejorScore = similitud;
      mejorCoincidenciaSimilar = valorCelda;
    }
  }
  
  return {
    coincidenciasExactas: coincidenciasExactas,
    coincidenciaSimilar: mejorCoincidenciaSimilar
  };
}

function _normalizarEspacios(texto) {
  return texto.trim().replace(/\s+/g, ' ');
}

function _calcularSimilitudEstricta(texto1, texto2) {
  if (texto1 === texto2) return 1.0;
  if (!texto1 || !texto2) return 0.0;
  
  const longitudMax = Math.max(texto1.length, texto2.length);
  const longitudMin = Math.min(texto1.length, texto2.length);
  
  let caracteresCoincidentes = 0;
  for (let i = 0; i < longitudMin; i++) {
    if (texto1[i] === texto2[i]) {
      caracteresCoincidentes++;
    }
  }
  
  const penalizacionLongitud = (longitudMax - longitudMin) / longitudMax;
  return (caracteresCoincidentes / longitudMax) * (1 - penalizacionLongitud * 0.5);
}

function _norm(s) {
  return s
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .toLowerCase()
    .trim();
}

function _findHeaderIndex(encabezados, nombreBuscado) {
  const target = _norm(nombreBuscado).replace(/\s+/g, ' ');
  for (let i = 0; i < encabezados.length; i++) {
    const cand = _norm(encabezados[i]).replace(/\s+/g, ' ');
    if (cand === target) return i;
  }
  return -1;
}

function _alert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (_) {
    Logger.log(String(msg));
  }
}

function _toast(msg, title, secs) {
  try {
    const m = String(msg == null ? '' : msg);
    const t = String(title == null ? '' : title); 
    const s = Math.max(1, Math.min(Number(secs == null ? 4 : secs), 10));
    SpreadsheetApp.getActiveSpreadsheet().toast(m, t, s);
  } catch (_) {
    Logger.log(String(msg));
  }
}
