function calcularFechaVencimiento() {
  const sheet = HOJA_ACTIVA_CONFIG.obtenerHojaDatos();
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠️ Por favor, selecciona la hoja DATOS_1 o DATOS_2 antes de calcular la fecha de vencimiento');
    return "❌ HOJA INVÁLIDA";
  }
  
  const fechaFirma = sheet.getRange(FECHA_CONFIG.FECHA_FIRMA_CELL).getValue();
  const plazoMeses = sheet.getRange(FECHA_CONFIG.PLAZO_MESES_CELL).getValue();
  
  let resultado = "SIN DATOS";

  if (fechaFirma && fechaFirma instanceof Date && !isNaN(fechaFirma.getTime())) {
    if (plazoMeses && !isNaN(plazoMeses) && plazoMeses > 0) {
      try {
        const vencimiento = agregarMesesConBisiestos(fechaFirma, parseInt(plazoMeses));
        resultado = Utilities.formatDate(vencimiento, Session.getScriptTimeZone(), "dd/MM/yyyy");
        
        const celdaResultado = sheet.getRange(FECHA_CONFIG.RESULTADO_CELL);
        celdaResultado.setValue(resultado);
        
        if (FECHA_CONFIG.PROTECT_RESULT) {
          celdaResultado.setNumberFormat('@STRING@');
        }
        
      } catch (error) {
        console.error('Error en cálculo de vencimiento:', error);
        resultado = "❌ ERROR CÁLCULO";
        sheet.getRange(FECHA_CONFIG.RESULTADO_CELL).setValue(resultado);
      }
    } else {
      resultado = "⚠️ PLAZO INVÁLIDO";
      sheet.getRange(FECHA_CONFIG.RESULTADO_CELL).setValue(resultado);
    }
  } else {
    resultado = "⚠️ FECHA INVÁLIDA";
    sheet.getRange(FECHA_CONFIG.RESULTADO_CELL).setValue(resultado);
  }
  
  return resultado;
}

function agregarMesesConBisiestos(fecha, meses) {
  const nuevaFecha = new Date(fecha);
  const mesOriginal = nuevaFecha.getMonth();
  const diaOriginal = nuevaFecha.getDate();
  
  const nuevoMes = (mesOriginal + meses) % 12;
  const anosAgregados = Math.floor((mesOriginal + meses) / 12);
  
  nuevaFecha.setFullYear(nuevaFecha.getFullYear() + anosAgregados);
  nuevaFecha.setMonth(nuevoMes);
  
  if (nuevaFecha.getDate() !== diaOriginal) {
    nuevaFecha.setDate(0);
  }
  
  return nuevaFecha;
}

function esBisiesto(ano) {
  return (ano % 4 === 0 && ano % 100 !== 0) || (ano % 400 === 0);
}
