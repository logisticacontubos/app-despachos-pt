// lib/sheets.js
// Reemplaza lo que hacía Code.gs con SpreadsheetApp — pero hablando directo con la API de
// Google Sheets desde Vercel, en vez de pasar por la capa "Aplicación web" de Apps Script
// (que fue la que estuvo fallando).

const { google } = require('googleapis');

const SHEET_ID = process.env.GOOGLE_SHEET_ID;

let _sheetsClientCache = null;

async function getSheetsClient() {
  if (_sheetsClientCache) return _sheetsClientCache;
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const key = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');
  const auth = new google.auth.JWT(email, null, key, ['https://www.googleapis.com/auth/spreadsheets']);
  await auth.authorize();
  _sheetsClientCache = google.sheets({ version: 'v4', auth });
  return _sheetsClientCache;
}

// Lee una hoja completa y la devuelve como arreglo de objetos { Columna: valor, ... },
// igual que hacía sheetToObjects_ en Code.gs. Los valores llegan ya formateados como texto
// (fechas en dd/MM/yyyy, igual que se veían en el Sheet), gracias a que se piden con
// valueRenderOption FORMATTED_VALUE (el comportamiento por defecto de la API).
async function sheetToObjects(nombreHoja) {
  const sheets = await getSheetsClient();
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: nombreHoja
  });
  const data = resp.data.values || [];
  if (data.length === 0) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(fila => fila.join('') !== '')
    .map((fila, idx) => {
      const obj = { _rowIndex: idx + 2 }; // fila 1 es encabezado, igual que en Code.gs
      headers.forEach((h, i) => { obj[h] = fila[i] !== undefined ? fila[i] : ''; });
      return obj;
    });
}

// Agrega una o varias filas al final de la hoja, en un solo viaje a la API.
async function appendRows(nombreHoja, filas) {
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: nombreHoja,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: filas }
  });
}

// Actualiza una sola celda, dado el número de fila real (_rowIndex) y el nombre de la columna.
// Necesita los encabezados para saber en qué posición (A, B, C...) está esa columna.
async function updateCell(nombreHoja, rowIndex, nombreColumna, valor) {
  const sheets = await getSheetsClient();
  const headers = await getHeaders(nombreHoja);
  const colIdx = headers.indexOf(nombreColumna);
  if (colIdx === -1) throw new Error(`Columna "${nombreColumna}" no encontrada en "${nombreHoja}"`);
  const colLetra = columnaALetra(colIdx + 1);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${nombreHoja}!${colLetra}${rowIndex}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[valor]] }
  });
}

// Actualiza varias columnas de la misma fila en un solo viaje (más eficiente que updateCell varias veces).
async function updateRowFields(nombreHoja, rowIndex, campos) {
  const sheets = await getSheetsClient();
  const headers = await getHeaders(nombreHoja);
  const data = [];
  Object.keys(campos).forEach(nombreColumna => {
    const colIdx = headers.indexOf(nombreColumna);
    if (colIdx === -1) return;
    const colLetra = columnaALetra(colIdx + 1);
    data.push({ range: `${nombreHoja}!${colLetra}${rowIndex}`, values: [[campos[nombreColumna]]] });
  });
  if (data.length === 0) return;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { valueInputOption: 'USER_ENTERED', data }
  });
}

let _headersCache = {};
async function getHeaders(nombreHoja) {
  if (_headersCache[nombreHoja]) return _headersCache[nombreHoja];
  const sheets = await getSheetsClient();
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${nombreHoja}!1:1`
  });
  const headers = (resp.data.values || [[]])[0];
  _headersCache[nombreHoja] = headers;
  return headers;
}

function columnaALetra(num) {
  let letra = '';
  while (num > 0) {
    const resto = (num - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    num = Math.floor((num - 1) / 26);
  }
  return letra;
}

// Genera el siguiente ID de un prefijo dado (ej. "PED-0042"), buscando el número más alto ya
// usado — igual que generarId_ en Code.gs.
async function generarId(prefijo, nombreHoja) {
  const filas = await sheetToObjects(nombreHoja);
  let maxNum = 0;
  const patron = new RegExp('^' + prefijo + '-(\\d+)$');
  filas.forEach(f => {
    const idVal = String(f[Object.keys(f)[1]] || ''); // la primera columna de datos suele ser el ID
    const match = idVal.match(patron);
    if (match) {
      const n = parseInt(match[1], 10);
      if (n > maxNum) maxNum = n;
    }
  });
  return prefijo + '-' + String(maxNum + 1).padStart(4, '0');
}

// Formatea una fecha de hoy como "dd/MM/yyyy" (zona horaria de Colombia).
function hoyTexto() {
  const ahora = new Date();
  const bogota = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Bogota' }));
  const d = String(bogota.getDate()).padStart(2, '0');
  const m = String(bogota.getMonth() + 1).padStart(2, '0');
  const y = bogota.getFullYear();
  return `${d}/${m}/${y}`;
}

// Formatea la hora actual como "HH:mm:ss" (zona horaria de Colombia).
function horaTexto() {
  const ahora = new Date();
  const bogota = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Bogota' }));
  const h = String(bogota.getHours()).padStart(2, '0');
  const m = String(bogota.getMinutes()).padStart(2, '0');
  const s = String(bogota.getSeconds()).padStart(2, '0');
  return `${h}:${m}:${s}`;
}

module.exports = {
  sheetToObjects,
  appendRows,
  updateCell,
  updateRowFields,
  getHeaders,
  generarId,
  hoyTexto,
  horaTexto
};
