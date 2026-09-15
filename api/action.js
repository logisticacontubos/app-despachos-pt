// api/action.js
// TODO combinado en un solo archivo (login + sheets + rutas) para evitar problemas de Vercel
// encontrando módulos en carpetas separadas.

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

function hoyTexto() {
  const ahora = new Date();
  const bogota = new Date(ahora.toLocaleString('en-US', { timeZone: 'America/Bogota' }));
  const d = String(bogota.getDate()).padStart(2, '0');
  const m = String(bogota.getMonth() + 1).padStart(2, '0');
  const y = bogota.getFullYear();
  return `${d}/${m}/${y}`;
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
      default:
        resultado = { ok: false, error: 'Acción no reconocida o aún no migrada: ' + accion };
    }

    res.status(200).json(resultado);
  } catch (err) {
    res.status(200).json({ ok: false, error: err.message });
  }
};

async function login(pin) {
  const usuarios = await sheetToObjects('Usuarios');
  const usuario = usuarios.find(u =>
    String(u.PIN).trim() === String(pin).trim() &&
    (u.Activo === true || String(u.Activo).toUpperCase() === 'TRUE')
  );
  if (!usuario) return { ok: false, error: 'PIN incorrecto.' };
  return {
    ok: true,
    usuario: {
      nombre: usuario.Nombre,
      rol: usuario.Rol,
      empresaAsignada: usuario.Empresa_Asignada || null
    }
  };
}

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
