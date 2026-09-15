// api/action.js
// Reemplaza el doPost/doGet de Code.gs. Corre en el mismo dominio que el index.html, así que no
// hay problema de CORS ni de la capa "Aplicación web" de Apps Script que estaba fallando.
//
// TERCERA Y ÚLTIMA TANDA — ya incluye TODAS las acciones: login, getPedidos, getHistorialPedidos,
// getEntregas, getNoProgramados, registrarEntrega, registrarNoProgramado, registrarRazonAtraso,
// cargarPlanilla, editarPedido, eliminarPedido, cerrarPedido, getAlistamiento, registrarAlistamiento.

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

let _sheetIdsCache = null;
async function getSheetIdPorNombre(nombreHoja) {
  if (!_sheetIdsCache) {
    const sheets = await getSheetsClient();
    const resp = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    _sheetIdsCache = {};
    resp.data.sheets.forEach(s => { _sheetIdsCache[s.properties.title] = s.properties.sheetId; });
  }
  return _sheetIdsCache[nombreHoja];
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

async function appendRowsObj(nombreHoja, objetos) {
  if (objetos.length === 0) return;
  const headers = await getHeaders(nombreHoja);
  const filas = objetos.map(objeto => headers.map(h => (objeto[h] !== undefined ? objeto[h] : '')));
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: nombreHoja,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: filas }
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

async function eliminarFila(nombreHoja, rowIndex) {
  const sheetId = await getSheetIdPorNombre(nombreHoja);
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: rowIndex - 1, endIndex: rowIndex }
        }
      }]
    }
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

function parseFechaDDMMYYYY(texto) {
  if (!texto) return null;
  const partes = String(texto).split('/');
  if (partes.length !== 3) return null;
  const d = parseInt(partes[0], 10), m = parseInt(partes[1], 10), y = parseInt(partes[2], 10);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  return new Date(y, m - 1, d);
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
      case 'login': resultado = await login(body.pin); break;
      case 'getPedidos': resultado = await getPedidos(body.empresa); break;
      case 'getHistorialPedidos': resultado = await getHistorialPedidos(body.empresa); break;
      case 'getEntregas': resultado = await getEntregas(body.idPedido); break;
      case 'getNoProgramados': resultado = await getNoProgramados(body.empresa); break;
      case 'registrarEntrega': resultado = await registrarEntrega(body); break;
      case 'registrarNoProgramado': resultado = await registrarNoProgramado(body); break;
      case 'registrarRazonAtraso': resultado = await registrarRazonAtraso(body); break;
      case 'cargarPlanilla': resultado = await cargarPlanilla(body); break;
      case 'editarPedido': resultado = await editarPedido(body); break;
      case 'eliminarPedido': resultado = await eliminarPedido(body.idPedido); break;
      case 'cerrarPedido': resultado = await cerrarPedido(body); break;
      case 'getAlistamiento': resultado = await getAlistamiento(body.empresa, body.fecha); break;
      case 'registrarAlistamiento': resultado = await registrarAlistamiento(body); break;
      default:
        resultado = { ok: false, error: 'Acción no reconocida: ' + accion };
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
  return { ok: true, usuario: { nombre: usuario.Nombre, rol: usuario.Rol, empresaAsignada: usuario.Empresa_Asignada || null } };
}

// ---------- PEDIDOS ----------
async function limpiarAbiertasPapelesNacionalesVencidas() {
  const todos = await sheetToObjects('Pedidos');
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  for (const p of todos) {
    const cliente = String(p.Razon_Social_Cliente || '').toUpperCase();
    const total = String(p.Cantidad_Total || '').trim().toLowerCase();
    const esPapelesNacionales = cliente.indexOf('PAPELES NACIONALES') !== -1;
    const esCantidadAbierta = total === 'abierta';
    if (!esPapelesNacionales || !esCantidadAbierta || p.Estado === 'Completado') continue;
    if (String(p.Fecha_Entrega).trim().toLowerCase() === 'abierta') continue;
    const fechaPedido = parseFechaDDMMYYYY(p.Fecha_Entrega);
    if (!fechaPedido) continue;
    if (fechaPedido.getTime() < hoy.getTime()) {
      await eliminarFila('Pedidos', p._rowIndex);
    }
  }
}

async function getPedidos(empresa) {
  await limpiarAbiertasPapelesNacionalesVencidas();
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

async function editarPedido(body) {
  const pedidos = await sheetToObjects('Pedidos');
  const pedido = pedidos.find(p => p.ID_Pedido === body.idPedido);
  if (!pedido) return { ok: false, error: 'Pedido no encontrado.' };

  const campoAColumna = {
    ocCliente: 'OC_Cliente', pvt: 'PVT', cliente: 'Razon_Social_Cliente',
    referencia: 'Referencia', ciudad: 'Ciudad', fecha: 'Fecha_Entrega', cantidad: 'Cantidad_Total'
  };
  const campos = {};
  Object.keys(campoAColumna).forEach(campo => {
    if (body[campo] !== undefined && body[campo] !== '') campos[campoAColumna[campo]] = body[campo];
  });
  await updateRowFields('Pedidos', pedido._rowIndex, campos);
  return { ok: true };
}

async function eliminarPedido(idPedido) {
  const pedidos = await sheetToObjects('Pedidos');
  const pedido = pedidos.find(p => p.ID_Pedido === idPedido);
  if (!pedido) return { ok: false, error: 'Pedido no encontrado.' };
  const entregada = Number(pedido.Cantidad_Entregada) || 0;
  if (entregada > 0) {
    return { ok: false, error: 'Este pedido ya tiene entregas registradas — no se puede eliminar. Corrígelo en vez de borrarlo.' };
  }
  await eliminarFila('Pedidos', pedido._rowIndex);
  return { ok: true };
}

async function cerrarPedido(body) {
  const pedidos = await sheetToObjects('Pedidos');
  const pedido = pedidos.find(p => p.ID_Pedido === body.idPedido);
  if (!pedido) return { ok: false, error: 'Pedido no encontrado.' };
  await updateRowFields('Pedidos', pedido._rowIndex, {
    Estado: 'Cerrado',
    Razon_Cierre: body.razon || '',
    Cerrado_Por: body.cerradoPor || '',
    Fecha_Cierre: hoyTexto()
  });
  return { ok: true };
}

// ---------- ENTREGAS (trazabilidad) ----------
async function getEntregas(idPedido) {
  const entregas = (await sheetToObjects('Entregas')).filter(e => e.ID_Pedido === idPedido);
  return { ok: true, entregas };
}

async function registrarEntrega(body) {
  const idEntrega = await generarId('ENT', 'Entregas', 'ID_Entrega');
  let duracionMin = '';
  if (body.horaInicio && body.horaFin) {
    const inicio = new Date('1970-01-01T' + body.horaInicio + 'Z');
    const fin = new Date('1970-01-01T' + body.horaFin + 'Z');
    duracionMin = Math.max(1, Math.round((fin - inicio) / 60000));
  }
  await appendRowObj('Entregas', {
    ID_Entrega: idEntrega, ID_Pedido: body.idPedido, Fecha_Entrega: hoyTexto(),
    Cantidad_Despachada: body.cantidad, Transportador: body.transportador || '',
    Tipo_Novedad: body.tipoNovedad || 'Sin novedad', Detalle_Novedad: body.detalleNovedad || '',
    Hora_Inicio_Cargue: body.horaInicio || '', Hora_Fin_Cargue: body.horaFin || '',
    Duracion_Cargue_Min: duracionMin, Registrado_Por: body.registradoPor || ''
  });

  await aplicarEntregaAPedido(body.idPedido, body.cantidad);
  return { ok: true };
}

async function aplicarEntregaAPedido(idPedido, cantidadNueva) {
  const pedidos = await sheetToObjects('Pedidos');
  const pedido = pedidos.find(p => p.ID_Pedido === idPedido);
  if (!pedido) return;
  const totalNum = Number(pedido.Cantidad_Total);
  const esAbierta = isNaN(totalNum);
  const nuevaEntregada = (Number(pedido.Cantidad_Entregada) || 0) + Number(cantidadNueva);
  const nuevoEstado = esAbierta ? 'Completado' : (nuevaEntregada >= totalNum ? 'Completado' : 'Parcial');
  await updateRowFields('Pedidos', pedido._rowIndex, {
    Cantidad_Entregada: nuevaEntregada, Estado: nuevoEstado, Fecha_Ultima_Entrega: hoyTexto()
  });
}

// ---------- NO PROGRAMADOS ----------
async function getNoProgramados(empresa) {
  const noProgramados = (await sheetToObjects('NoProgramados')).filter(n => n.Empresa === empresa);
  return { ok: true, noProgramados };
}

async function registrarNoProgramado(body) {
  const id = await generarId('NOPROG', 'NoProgramados', 'ID_NoProgramado');
  await appendRowObj('NoProgramados', {
    ID_NoProgramado: id, Empresa: body.empresa, OC_Cliente: body.ocCliente,
    Razon_Social_Cliente: body.cliente, Referencia: body.referencia, Ciudad: body.ciudad,
    Cantidad_Despachada: body.cantidad, Transportador: body.transportador || '',
    Fecha_Registro: hoyTexto(), Registrado_Por: body.registradoPor || '',
    Estado_Vinculo: 'Pendiente', ID_Pedido_Vinculado: '', PVT: body.pvt || '',
    Autorizado_Por: body.autorizadoPor || ''
  });
  return { ok: true };
}

async function vincularNoProgramados(empresa) {
  const noProgramados = (await sheetToObjects('NoProgramados')).filter(n => n.Empresa === empresa && n.Estado_Vinculo === 'Pendiente');
  if (noProgramados.length === 0) return;
  const pedidos = await sheetToObjects('Pedidos');

  for (const np of noProgramados) {
    const pedidoCoincide = pedidos.find(p => p.Empresa === empresa && String(p.OC_Cliente) === String(np.OC_Cliente));
    if (pedidoCoincide) {
      await aplicarEntregaAPedido(pedidoCoincide.ID_Pedido, np.Cantidad_Despachada);
      await updateRowFields('NoProgramados', np._rowIndex, {
        Estado_Vinculo: 'Vinculado', ID_Pedido_Vinculado: pedidoCoincide.ID_Pedido
      });
    }
  }
}

// ---------- RAZÓN DE ATRASO ----------
async function registrarRazonAtraso(body) {
  const id = await generarId('RAZ', 'Razones_Atraso', 'ID_Razon');
  await appendRowObj('Razones_Atraso', {
    ID_Razon: id, ID_Pedido: body.idPedido, Fecha_Programada_Original: body.fechaProgramadaOriginal || '',
    Razon: body.razon, Registrado_Por: body.registradoPor || '', Fecha_Registro: hoyTexto(),
    Nueva_Fecha_Programada: body.nuevaFechaProgramada || ''
  });
  return { ok: true };
}

// ---------- CARGAR PLANILLA ----------
async function cargarPlanilla(body) {
  const empresa = body.empresa;
  const fechaEntregaDefault = body.fecha || '';
  const pedidosNuevos = body.pedidos;

  const existentes = await sheetToObjects('Pedidos');
  let maxNum = 0;
  const patronId = /^PED-(\d+)$/;
  existentes.forEach(p => {
    const match = String(p.ID_Pedido || '').match(patronId);
    if (match) { const n = parseInt(match[1], 10); if (n > maxNum) maxNum = n; }
  });

  let creados = 0, vinculados = 0;
  const filasNuevas = [];

  for (const p of pedidosNuevos) {
    const existente = existentes.find(e =>
      String(e.Empresa).trim() === String(empresa).trim() &&
      String(e.OC_Cliente).trim() === String(p.ocCliente).trim() &&
      String(e.PVT || '').trim() === String(p.pvt || '').trim() &&
      String(e.Razon_Social_Cliente).trim().toUpperCase() === String(p.cliente).trim().toUpperCase() &&
      String(e.Referencia).trim().toUpperCase() === String(p.referencia).trim().toUpperCase() &&
      e.Estado !== 'Completado'
    );

    if (existente) {
      const fechaNueva = p.fecha || fechaEntregaDefault;
      const campos = {};
      if (fechaNueva) campos.Fecha_Entrega = fechaNueva;
      if (p.notas) campos.Notas_Planilla = p.notas;
      await updateRowFields('Pedidos', existente._rowIndex, campos);
      vinculados++;
    } else {
      maxNum++;
      const id = 'PED-' + String(maxNum).padStart(4, '0');
      const esClienteRecoge = /cliente\s*recoge/i.test(p.notas || '') ? 'TRUE' : 'FALSE';
      filasNuevas.push({
        ID_Pedido: id, Empresa: empresa, OC_Cliente: p.ocCliente, PVT: p.pvt || '',
        Razon_Social_Cliente: p.cliente, Referencia: p.referencia, Ciudad: p.ciudad,
        Fecha_Entrega: p.fecha || fechaEntregaDefault, Cantidad_Total: p.cantidad, Cantidad_Entregada: 0,
        Estado: 'Pendiente', Fecha_Carga_Planilla: hoyTexto(), Cargado_Por: body.cargadoPor || '',
        Es_No_Programado: 'FALSO', Notas_Planilla: p.notas || '', Es_Cliente_Recoge: esClienteRecoge
      });
      creados++;
    }
  }

  await appendRowsObj('Pedidos', filasNuevas);
  await vincularNoProgramados(empresa);

  return { ok: true, pedidosCargados: creados, pedidosVinculados: vinculados };
}

// ---------- ALISTAMIENTO ----------
async function getAlistamiento(empresa, fecha) {
  const registros = (await sheetToObjects('Alistamiento')).filter(a => a.Empresa === empresa && a.Fecha_Revision === fecha);
  const porPedido = {};
  registros.forEach(a => { porPedido[a.ID_Pedido] = a; });
  return { ok: true, alistamiento: porPedido };
}

async function registrarAlistamiento(body) {
  const registros = await sheetToObjects('Alistamiento');
  const existente = registros.find(a => a.ID_Pedido === body.idPedido && a.Fecha_Revision === body.fecha);

  if (existente) {
    await updateRowFields('Alistamiento', existente._rowIndex, {
      Estado_Alistamiento: body.estado, Cantidad_Disponible: body.cantidadDisponible || '',
      Cantidad_Faltante: body.cantidadFaltante || '', Registrado_Por: body.registradoPor || ''
    });
    return { ok: true, actualizado: true };
  }

  const id = await generarId('ALT', 'Alistamiento', 'ID_Alistamiento');
  await appendRowObj('Alistamiento', {
    ID_Alistamiento: id, ID_Pedido: body.idPedido, Empresa: body.empresa, Fecha_Revision: body.fecha,
    Estado_Alistamiento: body.estado, Cantidad_Disponible: body.cantidadDisponible || '',
    Cantidad_Faltante: body.cantidadFaltante || '', Registrado_Por: body.registradoPor || ''
  });
  return { ok: true, actualizado: false };
}
