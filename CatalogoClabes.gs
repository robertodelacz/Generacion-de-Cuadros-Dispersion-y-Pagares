/********************** LOOKUP CLABE Y BANCO (SIN TRIGGERS) **********************/
const LK = {
  // ✅ ACTUALIZADO: Incluir todas las variantes de nombres de hojas
  HOJA_DATOS: ['DATOS_1', 'Datos_1', 'DATOS_2', 'Datos_2', 'DATOS', 'Datos'],
  HOJA_CATALOGO: 'Catalogo Cuentas',
  COL_OPCION: 9,   // I - Beneficiario
  COL_BANCO:  7,   // G - Banco
  COL_CLABE:  8,   // H - CLABE
  START_ROW: 29,   // ✅ Actualizado a 29 (en lugar de 26)
  CACHE_KEY: 'CAT_CUENTAS_V1',
  CACHE_SECS: 300
};

function _norm_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ')
    .trim()
    .toUpperCase();
}

function _getCatalogoMap_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(LK.CACHE_KEY);
  if (cached) { 
    try { 
      return new Map(JSON.parse(cached)); 
    } catch(_){} 
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(LK.HOJA_CATALOGO);
  if (!sh) throw new Error('No se encuentra la hoja "Catalogo Cuentas".');

  const last = sh.getLastRow();
  const map = new Map();
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last-1, 3).getDisplayValues();
    for (const [beneficiario, banco, clabe] of vals) {
      const k = _norm_(beneficiario);
      const v = {
        banco: String(banco || '').trim(),
        clabe: String(clabe || '').trim()
      };
      if (k) map.set(k, v);
    }
  }
  try { 
    cache.put(LK.CACHE_KEY, JSON.stringify([...map]), LK.CACHE_SECS); 
  } catch(_){}
  return map;
}

function buscarDatosPorBeneficiario(beneficiario) {
  const map = _getCatalogoMap_();
  const key = _norm_(beneficiario);
  if (!key) return { banco: '', clabe: '' };
  
  if (map.has(key)) return map.get(key);
  
  for (const [k, v] of map) {
    if (k.includes(key) || key.includes(k)) return v;
  }
  
  return { banco: '', clabe: '' };
}

function _esClabeValida_(clabe) {
  return /^\d{18}$/.test(String(clabe||'').trim());
}

function handleClabeDesdeCatalogo_(e) {
  try {
    if (!e || !e.range) return;
    
    const sh = e.range.getSheet();
    const name = (sh.getName()||'').trim();
    
    // ✅ CAMBIO: Usar función centralizada de validación
    if (!HOJA_ACTIVA_CONFIG.esHojaDatos(name)) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < LK.START_ROW) return;
    if (col !== LK.COL_OPCION) return;

    const beneficiario = (e.value != null ? e.value : e.range.getDisplayValue());
    const rngBanco = sh.getRange(row, LK.COL_BANCO);
    const rngClabe = sh.getRange(row, LK.COL_CLABE);

    // Si está vacío, no hacemos nada
    if (!beneficiario || String(beneficiario).trim() === '') {
      return;
    }

    const datos = buscarDatosPorBeneficiario(beneficiario);
    
    if (!datos.banco && !datos.clabe) {
      SpreadsheetApp.getActive().toast('ℹ️ No encontrado en catálogo', 'Info', 3);
      return;
    }

    if (datos.banco) {
      rngBanco.setValue(datos.banco);
    }

    if (datos.clabe) {
      rngClabe.setNumberFormat('@').setValue(datos.clabe);
      if (!_esClabeValida_(datos.clabe)) {
        SpreadsheetApp.getActive().toast('⚠️ CLABE no tiene 18 dígitos', 'Advertencia', 5);
      }
    }

  } catch (err) {
    console.error('Error en handleClabeDesdeCatalogo_:', err);
  }
}

function poblarBancoYClabeManual() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getActiveSheet();
    const name = (sheet.getName()||'').trim();
    
    // ✅ CAMBIO: Usar función centralizada de validación
    if (!HOJA_ACTIVA_CONFIG.esHojaDatos(name)) {
      SpreadsheetApp.getUi().alert('❌ Esta función solo funciona en las hojas DATOS_1 o DATOS_2');
      return;
    }

    const celdaActiva = sheet.getActiveCell();
    const fila = celdaActiva.getRow();
    const columna = celdaActiva.getColumn();

    if (fila < LK.START_ROW) {
      SpreadsheetApp.getUi().alert('⚠️ Debes seleccionar una celda en la fila ' + LK.START_ROW + ' o superior');
      return;
    }

    if (columna !== LK.COL_OPCION) {
      SpreadsheetApp.getUi().alert('⚠️ Debes seleccionar una celda en la columna I (Beneficiario)');
      return;
    }

    const beneficiario = celdaActiva.getDisplayValue();
    const rngBanco = sheet.getRange(fila, LK.COL_BANCO);
    const rngClabe = sheet.getRange(fila, LK.COL_CLABE);

    if (!beneficiario || String(beneficiario).trim() === '') {
      SpreadsheetApp.getActive().toast('ℹ️ Celda vacía - Datos conservados', 'Info', 3);
      return;
    }

    const datos = buscarDatosPorBeneficiario(beneficiario);
    
    if (!datos.banco && !datos.clabe) {
      SpreadsheetApp.getActive().toast('❌ No encontrado en catálogo', 'Error', 3);
      return;
    }

    if (datos.banco) {
      rngBanco.setValue(datos.banco);
    } else {
      SpreadsheetApp.getActive().toast('⚠️ Banco no encontrado en catálogo', 'Advertencia', 3);
    }

    if (datos.clabe) {
      rngClabe.setNumberFormat('@').setValue(datos.clabe);
      if (!_esClabeValida_(datos.clabe)) {
        SpreadsheetApp.getActive().toast('⚠️ CLABE no tiene 18 dígitos', 'Advertencia', 5);
      }
    } else {
      SpreadsheetApp.getActive().toast('⚠️ CLABE no encontrada en catálogo', 'Advertencia', 3);
    }

    SpreadsheetApp.getActive().toast('✅ Banco y CLABE actualizados', 'Éxito', 3);

  } catch (err) {
    console.error('Error en poblarBancoYClabeManual:', err);
    SpreadsheetApp.getUi().alert('❌ Error: ' + (err.message || err));
  }
}
