const HOJA_ACTIVA_CONFIG = {
  obtenerHojaDatos: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getActiveSheet();
    const nombreHoja = hoja.getName();
    
    if (nombreHoja === 'DATOS_2' || nombreHoja === 'Datos_2' || nombreHoja === 'datos_2') {
      return ss.getSheetByName('DATOS_2') || ss.getSheetByName('Datos_2');
    }
    
    return ss.getSheetByName('DATOS_1') || ss.getSheetByName('Datos_1') || ss.getSheetByName('Datos');
  },
  
  esHojaDatos: function(nombreHoja) {
    if (!nombreHoja) return false;
    const nombre = nombreHoja.toLowerCase().trim();
    return nombre === 'datos_1' || nombre === 'datos_2' || 
           nombre === 'datos-1' || nombre === 'datos-2';
  }
};
