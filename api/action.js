// api/action.js
// Reemplaza el doPost/doGet de Code.gs. Corre en el mismo dominio que el index.html, así que no
// hay problema de CORS ni de la capa "Aplicación web" de Apps Script que estaba fallando.
//
// SEGUNDA TANDA — ya incluye: login, getPedidos, getHistorialPedidos, getEntregas,
// getNoProgramados, registrarEntrega, registrarNoProgramado, registrarRazonAtraso.
// Todavía faltan: cargarPlanilla, editarPedido, eliminarPedido, cerrarPedido, Alistamiento
// (llegan en la siguiente entrega).

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

let _headersCache = {};
async function getHeaders(nombreHoja) {
  if (_headersCache[nombreHoja]) return _headersCache[nombreHoja];
  const sheets = await getSheetsClient();
  const resp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${nombreHoja}!1:1` });
  const headers = (resp.data.values || [[]])[0];
  _headersCache[nombreHoja] = headers;
  return headers;
}

async function sheetToObjects(nombreHoja) {
  const sheets = await getSheetsClient();
  const resp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: nombreHoja });
  const data = resp.data.values || [];
  if (data.length === 0) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(fila => fila.join('') !== '')
    .map((fila, idx) => {
      const obj = { _rowIndex: idx + 2 };
      headers.forEach((h, i) => { obj[h] = fila[i] !== undefined ? fila[i] : ''; });
      return obj;
    });
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

async function appendRowObj(nombreHoja, objeto) {
  const headers = await getHeaders(nombreHoja);
  const fila = headers.map(h => (objeto[h] !== undefined ? objeto[h] : ''));
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: nombreHoja,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [fila] }
  });
}

async function updateRowFields(nombreHoja, rowIndex, campos) {
  const headers = await getHeaders(nombreHoja);
  const sheets = await getSheetsClient();
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

async function generarId(prefijo, nombreHoja, columnaId) {
  const filas = await sheetToObjects(nombreHoja);
  let maxNum = 0;
  const patron = new RegExp('^' + prefijo + '-(\\d+)$');
  filas.forEach(f => {
    const idVal = String(f[columnaId] || '');
    const match = idVal.match(patron);
    if (match) {
      const n = parseInt(match[1], 10);
      if (n > maxNum) maxNum = n;
    }
  });
  return prefijo + '-' + String(maxNum + 1).padStart(4, '0');
}

function hoyTexto() {
  const bogota = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Bogota' }));
  return `${String(bogota.getDate()).padStart(2, '0')}/${String(bogota.getMonth() + 1).padStart(2, '0')}/${bogota.getFullYear()}`;
}

// ---------- ENDPOINT PRINCIPAL ----------
module.exports = async (req, res) => {
  try {
    let body;
    if (req.method === 'GET') {
      body = req.query.data ? JSON.parse(req.query.data) : {};
    } else {
      body = req.body || {};
    }

    const accion = body.accion;
    let resultado;

    switch (accion) {
      case 'login':
        resultado = await login(body.pin);
        break;
      case 'getPedidos':
        resultado = await getPedidos(body.empresa);
        break;
      case 'getHistorialPedidos':
        resultado = await getHistorialPedidos(body.empresa);
        break;
      case 'getEntregas':
        resultado = await getEntregas(body.idPedido);
        break;
      case 'getNoProgramados':
        resultado = await getNoProgramados(body.empresa);
        break;
      case 'registrarEntrega':
        resultado = await registrarEntrega(body);
        break;
      case 'registrarNoProgramado':
        resultado = await registrarNoProgramado(body);
        break;
      case 'registrarRazonAtraso':
        resultado = await registrarRazonAtraso(body);
        break;
      default:
        resultado = { ok: false, error: 'Acción no reconocida o aún no migrada: ' + accion };
    }

    res.status(200).json(resultado);
  } catch (err) {
    res.status(200).json({ ok: false, error: err.message });
  }
};

// ---------- LOGIN ----------
async function login(pin) {
  const usuarios = await sheetToObjects('Usuarios');
  const usuario = usuarios.find(u =>
    String(u.PIN).trim() === String(pin).trim() &&
    (u.Activo === true || String(u.Activo).toUpperCase() === 'TRUE')
  );
  if (!usuario) return { ok: false, error: 'PIN incorrecto.' };
  return {
    ok: true,
    usuario: { nombre: usuario.Nombre, rol: usuario.Rol, empresaAsignada: usuario.Empresa_Asignada || null }
  };
}

// ---------- PEDIDOS ----------
async function getPedidos(empresa) {
  const todos = (await sheetToObjects('Pedidos')).filter(p => p.Empresa === empresa);
  const hoy = hoyTexto();
  const pedidos = todos.filter(p => {
    if (p.Estado !== 'Completado' && p.Estado !== 'Cerrado') return true;
    const fechaRef = p.Fecha_Ultima_Entrega || p.Fecha_Cierre || p.Fecha_Entrega;
    return fechaRef === hoy;
  });
  return { ok: true, pedidos };
}

async function getHistorialPedidos(empresa) {
  const todos = (await sheetToObjects('Pedidos')).filter(p => p.Empresa === empresa);
  const hoy = hoyTexto();
  const historial = todos.filter(p => {
    if (p.Estado !== 'Completado' && p.Estado !== 'Cerrado') return false;
    const fechaRef = p.Fecha_Ultima_Entrega || p.Fecha_Cierre || p.Fecha_Entrega;
    return fechaRef !== hoy;
  });
  return { ok: true, pedidos: historial };
}

// ---------- ENTREGAS (trazabilidad) ----------
async function getEntregas(idPedido) {
  const entregas = (await sheetToObjects('Entregas')).filter(e => e.ID_Pedido === idPedido);
  return { ok: true, entregas };
}

// Registra una entrega: agrega la fila en Entregas, y actualiza el pedido (cantidad entregada,
// estado, fecha última entrega) — igual que registrarEntrega_ + actualizarEstadoPedido_ en Code.gs.
async function registrarEntrega(body) {
  const idEntrega = await generarId('ENT', 'Entregas', 'ID_Entrega');

  let duracionMin = '';
  if (body.horaInicio && body.horaFin) {
    const inicio = new Date('1970-01-01T' + body.horaInicio + 'Z');
    const fin = new Date('1970-01-01T' + body.horaFin + 'Z');
    duracionMin = Math.max(1, Math.round((fin - inicio) / 60000));
  }

  await appendRowObj('Entregas', {
    ID_Entrega: idEntrega,
    ID_Pedido: body.idPedido,
    Fecha_Entrega: hoyTexto(),
    Cantidad_Despachada: body.cantidad,
    Transportador: body.transportador || '',
    Tipo_Novedad: body.tipoNovedad || 'Sin novedad',
    Detalle_Novedad: body.detalleNovedad || '',
    Hora_Inicio_Cargue: body.horaInicio || '',
    Hora_Fin_Cargue: body.horaFin || '',
    Duracion_Cargue_Min: duracionMin,
    Registrado_Por: body.registradoPor || ''
  });

  const pedidos = await sheetToObjects('Pedidos');
  const pedido = pedidos.find(p => p.ID_Pedido === body.idPedido);
  if (!pedido) return { ok: false, error: 'Pedido no encontrado.' };

  const totalNum = Number(pedido.Cantidad_Total);
  const esAbierta = isNaN(totalNum);
  const nuevaEntregada = (Number(pedido.Cantidad_Entregada) || 0) + Number(body.cantidad);
  const nuevoEstado = esAbierta ? 'Completado' : (nuevaEntregada >= totalNum ? 'Completado' : 'Parcial');

  // Fecha_Ultima_Entrega se actualiza SIEMPRE que se registre una entrega (completa o parcial) —
  // es la fecha en que se tocó el pedido por última vez, igual que en Code.gs.
  const campos = { Cantidad_Entregada: nuevaEntregada, Estado: nuevoEstado, Fecha_Ultima_Entrega: hoyTexto() };
  await updateRowFields('Pedidos', pedido._rowIndex, campos);

  return { ok: true };
}

// ---------- NO PROGRAMADOS ----------
async function getNoProgramados(empresa) {
  const noProgramados = (await sheetToObjects('NoProgramados')).filter(n => n.Empresa === empresa);
  return { ok: true, noProgramados };
}

async function registrarNoProgramado(body) {
  const id = await generarId('NOPROG', 'NoProgramados', 'ID_NoProgramado');
  await appendRowObj('NoProgramados', {
    ID_NoProgramado: id,
    Empresa: body.empresa,
    OC_Cliente: body.ocCliente,
    Razon_Social_Cliente: body.cliente,
    Referencia: body.referencia,
    Ciudad: body.ciudad,
    Cantidad_Despachada: body.cantidad,
    Transportador: body.transportador || '',
    Fecha_Registro: hoyTexto(),
    Registrado_Por: body.registradoPor || '',
    Estado_Vinculo: 'Pendiente',
    ID_Pedido_Vinculado: '',
    PVT: body.pvt || '',
    Autorizado_Por: body.autorizadoPor || ''
  });
  return { ok: true };
}

// ---------- RAZÓN DE ATRASO ----------
async function registrarRazonAtraso(body) {
  const id = await generarId('RAZ', 'Razones_Atraso', 'ID_Razon');
  await appendRowObj('Razones_Atraso', {
    ID_Razon: id,
    ID_Pedido: body.idPedido,
    Fecha_Programada_Original: body.fechaProgramadaOriginal || '',
    Razon: body.razon,
    Registrado_Por: body.registradoPor || '',
    Fecha_Registro: hoyTexto(),
    Nueva_Fecha_Programada: body.nuevaFechaProgramada || ''
  });
  return { ok: true };
}
