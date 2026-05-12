const CENTENAS = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"];
const DECENAS  = ["", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"];
const UNIDADES = ["", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"];
const ESPECIALES = ["DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE"];
const VEINTI = ["VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISÉIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE"];

function NUM_TEXTO(numero, monedaCelda) {
  return _numTextoCore(numero, monedaCelda);
}

function _numTextoCore(numero, monedaCelda) {
  if (numero === "" || numero === null || numero === undefined || numero === 0) {
    return "-";
  }

  const num = Number(numero);
  if (isNaN(num)) return "Error: No es un número válido";

  const parteEntera = Math.trunc(Math.abs(num));
  const centavos = Math.round((Math.abs(num) - parteEntera) * 100);
  const centavosTxt = (centavos < 10 ? "0" : "") + centavos + "/100";

  let moneda = "";
  if (Array.isArray(monedaCelda)) {
    moneda = (monedaCelda[0] && monedaCelda[0][0] != null) ? String(monedaCelda[0][0]) : "";
  } else {
    moneda = String(monedaCelda || "");
  }
  moneda = moneda.trim().toUpperCase();

  const esUSD = ["USD", "US", "DOLAR", "DÓLAR", "DOLARES", "DÓLARES"].includes(moneda);
  const CURRENCY = esUSD
    ? { singular: " DÓLAR", plural: " DÓLARES", de: " DE DÓLARES" }
    : { singular: " PESO",  plural: " PESOS",   de: " DE PESOS"  };

  let cadena;
  if (parteEntera === 0) {
    cadena = "CERO" + CURRENCY.plural + " " + centavosTxt;
  } else {
    cadena = convertirParteEntera(parteEntera);
    const esMillon = parteEntera >= 1000000;
    const esMillonExacto = esMillon && (parteEntera % 1000000 === 0);

    if (esMillonExacto) {
      cadena += CURRENCY.de;
    } else if (esMillon) {
      cadena += CURRENCY.plural;
    } else {
      cadena += (parteEntera === 1 ? CURRENCY.singular : CURRENCY.plural);
    }

    cadena = cadena + " " + centavosTxt;
  }

  // Capitalizar solo la primera letra y dejar el resto en minúsculas
  return cadena.charAt(0).toUpperCase() + cadena.slice(1).toLowerCase();
}


function convertirParteEntera(n) {
  if (n === 0) return "";
  const millones = Math.floor(n / 1000000);
  const restoMillones = n % 1000000;
  const miles = Math.floor(restoMillones / 1000);
  const cientos = restoMillones % 1000;

  let s = "";
  if (millones > 0) s += convierteCifra(millones, true) + (millones === 1 ? " MILLÓN" : " MILLONES");
  if (miles > 0) {
    if (s) s += " ";
    const milTexto = convierteCifra(miles, true);
    s += (milTexto === "UN" && miles === 1) ? "MIL" : (milTexto + " MIL");
  }
  if (cientos > 0) {
    if (s) s += " ";
    s += convierteCifra(cientos, false);
  }
  return s;
}

function convierteCifra(n, esMayor) {
  if (n === 0)   return "";
  if (n === 100) return "CIEN";

  const centena = Math.floor(n / 100);
  const resto = n % 100;
  const decena = Math.floor(resto / 10);
  const unidad = resto % 10;

  let partes = [];
  if (centena > 0) partes.push(CENTENAS[centena]);

  if (resto > 0) {
    if (decena === 1) {
      partes.push(ESPECIALES[unidad]);
    } else if (decena === 2 && unidad > 0) {
      partes.push(VEINTI[unidad]);
    } else {
      if (decena > 1) partes.push(DECENAS[decena]);
      if (unidad > 0) {
        if (decena > 2) partes.push("Y");
        partes.push(UNIDADES[unidad]);
      }
    }
  }
  return partes.join(" ");
}
