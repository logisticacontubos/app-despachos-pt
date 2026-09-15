// api/action.js
// Reemplaza el doPost/doGet de Code.gs. Como corre en el mismo dominio que el index.html
// (app-despachos-pt.vercel.app), no hay problema de CORS ni de la capa "Aplicación web" de
// Apps Script que estaba fallando — esto es una función normal de servidor.
//
// PRIMERA PARTE — implementado hasta ahora: login, getPedidos, getHistorialPedidos.
// El resto de acciones (registrarEntrega, editarPedido, etc.) llegan en la siguiente entrega.

const { sheetToObjects, hoyTexto } = require('../lib/sheets');

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
        resultado = { ok: false, error: 'Acción no reconocida o aún no migrada: ' + accion + ' (llega en la próxima entrega)' };
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
    usuario: {
      nombre: usuario.Nombre,
      rol: usuario.Rol,
      empresaAsignada: usuario.Empresa_Asignada || null
    }
  };
}

// ---------- PEDIDOS ----------
// Misma lógica que en Code.gs: "Inicio" solo trae lo activo + lo completado/cerrado de HOY.
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

// Historial completo (Completados/Cerrados de días anteriores) — se pide aparte, solo al abrir esa pestaña.
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
