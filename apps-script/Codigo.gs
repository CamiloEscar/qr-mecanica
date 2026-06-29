function obtenerSheetId() {
  return PropertiesService.getScriptProperties().getProperty('SHEET_ID');
}

function obtenerFotosFolderId() {
  return PropertiesService.getScriptProperties().getProperty('FOTOS_FOLDER_ID');
}

function obtenerQrFolderId() {
  return PropertiesService.getScriptProperties().getProperty('DRIVE_QR_FOLDER_ID');
}

function obtenerAdminKey() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
}

const SHEETS = {
  CLIENTES: 'Clientes',
  VEHICULOS: 'Vehiculos',
  SERVICIOS: 'Servicios',
  ACCESOS_LOG: 'AccesosLog'
};

const COLUMNAS = {
  CLIENTES: ['ID_Cliente', 'Nombre', 'Telefono', 'Email', 'Fecha_Alta', 'Consentimiento'],
  VEHICULOS: ['ID_Vehiculo', 'Token', 'Fecha_Alta', 'Activo', 'Patente', 'Marca', 'Modelo', 'Anio', 'Combustible', 'ID_Cliente', 'URL_QR_Impresa'],
  SERVICIOS: ['ID_Servicio', 'Fecha', 'ID_Vehiculo', 'Kilometraje', 'Descripcion', 'Repuestos', 'Proximo_Mantenimiento', 'Observaciones', 'Fotos_IDs', 'ID_Mecanico'],
  ACCESOS_LOG: ['Timestamp', 'Token_Suffix', 'Email_O_Anonimo']
};

const TOKEN_REGEX = /^[a-zA-Z0-9]{16}$/;
const CHARSET = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';

function doGet(e) {
  const accion = (e.parameter && e.parameter.accion) || '';
  if (accion === 'diag') return serveDiagnostico();
  if (accion === 'admin') return serveAdminPage((e.parameter && e.parameter.session) || '');
  if (accion === 'imprimir' && e.parameter.v) return serveQRPrintPage(e.parameter.v, (e.parameter && e.parameter.session) || '');
  if (accion === 'api') return apiMecano(e);
  const token = (e && e.parameter && e.parameter.t) || '';
  if (!TOKEN_REGEX.test(token)) {
    return serveError('El enlace no es válido. Verificá que esté completo.');
  }
  const vehiculo = buscarVehiculoPorToken(token);
  if (!vehiculo) {
    return serveError('No encontramos el vehículo o el enlace está inactivo.');
  }
  const servicios = listarServicios(vehiculo.ID_Vehiculo);
  logAcceso(token);
  return serveClientePage(vehiculo, servicios);
}

function doPost(e) {
  try {
    const adminKey = obtenerAdminKey();
    const provided = (e && e.parameter && e.parameter.adminKey) || '';
    if (!adminKey || provided !== adminKey) {
      return jsonOutput({ ok: false, error: 'No autorizado' });
    }
    const action = (e && e.parameter && e.parameter.action) || '';
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        payload = {};
      }
    }
    let resultado;
    switch (action) {
      case 'altaCliente':
        resultado = altaCliente(payload);
        break;
      case 'altaVehiculo':
        resultado = altaVehiculo(payload);
        break;
      case 'altaServicio':
        resultado = altaServicio(payload);
        break;
      default:
        resultado = { ok: false, error: 'Acción desconocida: ' + action };
    }
    return jsonOutput(resultado);
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function obtenerSheet(nombre) {
  return SpreadsheetApp.openById(obtenerSheetId()).getSheetByName(nombre);
}

function sheetAObjetos(nombreSheet) {
  const sheet = obtenerSheet(nombreSheet);
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return [];
  const headers = data[0];
  const objetos = [];
  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      const v = fila[j];
      obj[headers[j]] = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : v;
    }
    objetos.push(obj);
  }
  return objetos;
}

function generarTokenUnico() {
  const vehiculos = sheetAObjetos(SHEETS.VEHICULOS);
  const existentes = new Set();
  for (let i = 0; i < vehiculos.length; i++) {
    if (vehiculos[i].Token) existentes.add(vehiculos[i].Token);
  }
  while (true) {
    let token = '';
    for (let i = 0; i < 16; i++) {
      token += CHARSET.charAt(Math.floor(Math.random() * CHARSET.length));
    }
    if (!existentes.has(token)) return token;
  }
}

function buscarVehiculoPorToken(token) {
  const vehiculos = sheetAObjetos(SHEETS.VEHICULOS);
  for (let i = 0; i < vehiculos.length; i++) {
    const v = vehiculos[i];
    const activo = v.Activo === true || String(v.Activo).toUpperCase() === 'TRUE';
    if (v.Token === token && activo) {
      return v;
    }
  }
  return null;
}

function listarServicios(idVehiculo) {
  const servicios = sheetAObjetos(SHEETS.SERVICIOS);
  const filtrados = servicios.filter(function (s) { return s.ID_Vehiculo === idVehiculo; });
  filtrados.sort(function (a, b) {
    const fa = a.Fecha instanceof Date ? a.Fecha.getTime() : new Date(a.Fecha).getTime();
    const fb = b.Fecha instanceof Date ? b.Fecha.getTime() : new Date(b.Fecha).getTime();
    return fb - fa;
  });
  return filtrados;
}

function nextId(nombreSheet, prefijo) {
  const sheet = obtenerSheet(nombreSheet);
  const data = sheet.getDataRange().getValues();
  let max = 0;
  const patron = new RegExp('^' + prefijo + '-(\\d+)$');
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]);
    const match = id.match(patron);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > max) max = num;
    }
  }
  return prefijo + '-' + String(max + 1).padStart(4, '0');
}

function ahoraComoString() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function altaCliente(payload) {
  const sheet = obtenerSheet(SHEETS.CLIENTES);
  const id = nextId(SHEETS.CLIENTES, 'C');
  const fila = [
    id,
    payload.nombre || '',
    payload.telefono || '',
    payload.email || '',
    ahoraComoString(),
    payload.consentimiento || 'NO'
  ];
  sheet.appendRow(fila);
  return { ok: true, id: id };
}

function altaVehiculo(payload) {
  const sheet = obtenerSheet(SHEETS.VEHICULOS);
  const id = nextId(SHEETS.VEHICULOS, 'V');
  const token = generarTokenUnico();
  const fila = [
    id,
    token,
    ahoraComoString(),
    'TRUE',
    payload.patente || '',
    payload.marca || '',
    payload.modelo || '',
    payload.anio || '',
    payload.combustible || '',
    payload.idCliente || '',
    ''
  ];
  sheet.appendRow(fila);
  const urlQR = generarQRParaVehiculo(token, id);
  sheet.getRange(sheet.getLastRow(), 11).setValue(urlQR);
  return { ok: true, id: id, token: token, urlQR: urlQR };
}

function altaServicio(payload) {
  const sheet = obtenerSheet(SHEETS.SERVICIOS);
  const id = nextId(SHEETS.SERVICIOS, 'S');
  const fecha = payload.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fila = [
    id,
    fecha,
    payload.idVehiculo || '',
    payload.kilometraje || '',
    payload.descripcion || '',
    payload.repuestos || '',
    payload.proximoMantenimiento || '',
    payload.observaciones || '',
    payload.fotosIds || '',
    payload.idMecanico || ''
  ];
  sheet.appendRow(fila);
  return { ok: true, id: id };
}

function generarQRParaVehiculo(token, id) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const urlCompleta = webAppUrl + '?t=' + token;
  const qrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(urlCompleta) + '&size=400';
  const respuesta = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
  if (respuesta.getResponseCode() !== 200) {
    throw new Error('QuickChart no respondió 200: ' + respuesta.getResponseCode());
  }
  const blob = respuesta.getBlob();
  blob.setName('QR_' + id + '.png');
  const carpeta = DriveApp.getFolderById(obtenerQrFolderId());
  const archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return archivo.getUrl();
}

function logAcceso(token) {
  try {
    const sheet = obtenerSheet(SHEETS.ACCESOS_LOG);
    const sufijo = String(token).slice(-4);
    sheet.appendRow([ahoraComoString(), sufijo, 'anonimo']);
  } catch (err) {
  }
}

function serveClientePage(vehiculo, servicios) {
  const template = HtmlService.createTemplateFromFile('ClientePage');
  template.vehiculo = vehiculo;
  template.servicios = servicios;
  template.ultimo = servicios.length > 0 ? servicios[0] : null;
  return template
    .evaluate()
    .setTitle(vehiculo.Marca + ' ' + vehiculo.Modelo)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function serveError(mensaje) {
  const html = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Enlace no disponible</title>' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f8fafc;color:#0f172a;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;padding:20px}' +
    '.card{background:#ffffff;border-radius:12px;padding:32px;max-width:400px;text-align:center;box-shadow:0 4px 6px rgba(0,0,0,0.05);border:1px solid #e2e8f0}' +
    'h1{color:#dc2626;margin:0 0 16px;font-size:20px;font-weight:700}' +
    'p{margin:0;color:#64748b;font-size:14px;line-height:1.5}' +
    '</style></head><body>' +
    '<div class="card"><h1>Enlace no disponible</h1><p>' + mensaje + '</p></div>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('Enlace no disponible')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupInicial() {
  const props = PropertiesService.getScriptProperties();

  let sheetId = props.getProperty('SHEET_ID');
  if (!sheetId) {
    sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    props.setProperty('SHEET_ID', sheetId);
  }

  const ss = SpreadsheetApp.openById(sheetId);
  const hojasACrear = ['Clientes', 'Vehiculos', 'Servicios', 'AccesosLog'];
  const headersPorHoja = {
    'Clientes': COLUMNAS.CLIENTES,
    'Vehiculos': COLUMNAS.VEHICULOS,
    'Servicios': COLUMNAS.SERVICIOS,
    'AccesosLog': COLUMNAS.ACCESOS_LOG,
  };

  hojasACrear.forEach(nombre => {
    let hoja = ss.getSheetByName(nombre);
    if (!hoja) {
      hoja = ss.insertSheet(nombre);
      hoja.getRange(1, 1, 1, headersPorHoja[nombre].length)
        .setValues([headersPorHoja[nombre]]);
      hoja.getRange(1, 1, 1, headersPorHoja[nombre].length)
        .setFontWeight('bold')
        .setBackground('#e2e8f0');
      hoja.setFrozenRows(1);
    }
  });

  const hojaDefault = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
  if (hojaDefault && hojaDefault.getLastRow() <= 1 && ss.getSheets().length > 1) {
    ss.deleteSheet(hojaDefault);
  }

  let fotosFolderId = props.getProperty('FOTOS_FOLDER_ID');
  if (!fotosFolderId) {
    const carpeta = DriveApp.createFolder('QR_Mecanica_Fotos');
    fotosFolderId = carpeta.getId();
    props.setProperty('FOTOS_FOLDER_ID', fotosFolderId);
  }

  let qrFolderId = props.getProperty('DRIVE_QR_FOLDER_ID');
  if (!qrFolderId) {
    const carpeta = DriveApp.createFolder('QR_Mecanica_QRs');
    qrFolderId = carpeta.getId();
    props.setProperty('DRIVE_QR_FOLDER_ID', qrFolderId);
  }

  let adminKey = props.getProperty('ADMIN_KEY');
  if (!adminKey) {
    adminKey = generarTokenUnico();
    props.setProperty('ADMIN_KEY', adminKey);
  }

  let mecanicoPassword = props.getProperty('MECANICO_PASSWORD');
  if (!mecanicoPassword) {
    mecanicoPassword = generarTokenUnico();
    props.setProperty('MECANICO_PASSWORD', mecanicoPassword);
  }

  return {
    sheetId: sheetId,
    fotosFolderId: fotosFolderId,
    qrFolderId: qrFolderId,
    adminKey: adminKey,
    mecanicoPassword: mecanicoPassword,
    mensaje: 'Setup completado'
  };
}

function crearVehiculoDemo() {
  const props = PropertiesService.getScriptProperties();
  const adminKey = props.getProperty('ADMIN_KEY');

  const clienteDemo = altaCliente({
    nombre: 'Cliente Demo',
    telefono: '+5491100000000',
    email: 'demo@ejemplo.com',
    consentimiento: true
  });

  const clienteId = clienteDemo.id;

  const vehiculoDemo = altaVehiculo({
    idCliente: clienteId,
    patente: 'AB123CD',
    marca: 'Toyota',
    modelo: 'Corolla',
    anio: 2019,
    combustible: 'Nafta'
  });

  const servicioDemo = altaServicio({
    idVehiculo: vehiculoDemo.id,
    fecha: new Date(),
    kilometraje: 54320,
    descripcion: 'Cambio de aceite y filtro - Service de prueba',
    repuestos: 'Filtro Mann W712/75, Aceite 5W30 sintetico 4L',
    proximoMantenimiento: 'Diciembre 2026 o 60.000 km',
    observaciones: 'Vehiculo en excelente estado general',
    fotosIds: ''
  });

  const webAppUrl = ScriptApp.getService().getUrl();
  const urlCliente = webAppUrl + '?t=' + vehiculoDemo.token;

  return {
    mensaje: 'Demo creado exitosamente',
    adminKey: adminKey,
    clienteId: clienteId,
    vehiculoId: vehiculoDemo.id,
    vehiculoToken: vehiculoDemo.token,
    servicioId: servicioDemo.id,
    qrImageUrl: vehiculoDemo.urlQR,
    urlCliente: urlCliente
  };
}

function cargarDatosDemo() {
  const ss = SpreadsheetApp.openById(obtenerSheetId());
  const sheetClientes = ss.getSheetByName('Clientes');
  const sheetVehiculos = ss.getSheetByName('Vehiculos');
  const sheetServicios = ss.getSheetByName('Servicios');

  const dataClientes = sheetClientes.getDataRange().getValues();
  for (let i = 1; i < dataClientes.length; i++) {
    if (String(dataClientes[i][1] || '').trim() === 'Cliente Demo') {
      throw new Error('Ya existe un Cliente Demo en la base. Si querés volver a cargar los datos de prueba, primero eliminá ese cliente y sus vehículos asociados desde la planilla.');
    }
  }

  const hace = (dias) => {
    const fecha = new Date();
    fecha.setDate(fecha.getDate() - dias);
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  };

  const fechaAlta = ahoraComoString();

  const clientesData = [
    { nombre: 'Juan Perez', telefono: '+5491144551122', email: 'juan.perez@gmail.com' },
    { nombre: 'Maria Gonzalez', telefono: '+5491155663344', email: 'maria.gonzalez@hotmail.com' },
    { nombre: 'Carlos Rodriguez', telefono: '+5491166778899', email: '' },
    { nombre: 'Ana Martinez', telefono: '+5491177889900', email: 'ana.martinez@yahoo.com.ar' },
    { nombre: 'Pedro Lopez', telefono: '+5491188990011', email: 'pedro.lopez@outlook.com' }
  ];

  const idsClientes = {};
  clientesData.forEach(function (c) {
    const id = nextId('Clientes', 'C');
    idsClientes[c.nombre] = id;
    sheetClientes.appendRow([id, c.nombre, c.telefono, c.email, fechaAlta, true]);
  });

  const vehiculosData = [
    { patente: 'AB123CD', marca: 'Toyota', modelo: 'Corolla', anio: 2020, cliente: 'Juan Perez', combustible: 'Nafta' },
    { patente: 'AC456EF', marca: 'Ford', modelo: 'Focus', anio: 2018, cliente: 'Maria Gonzalez', combustible: 'Nafta' },
    { patente: 'AD789GH', marca: 'Volkswagen', modelo: 'Gol', anio: 2015, cliente: 'Carlos Rodriguez', combustible: 'GNC' },
    { patente: 'AE012IJ', marca: 'Chevrolet', modelo: 'Onix', anio: 2022, cliente: 'Ana Martinez', combustible: 'Nafta' },
    { patente: 'AF345KL', marca: 'Renault', modelo: 'Kwid', anio: 2021, cliente: 'Pedro Lopez', combustible: 'Nafta' },
    { patente: 'AG678MN', marca: 'Toyota', modelo: 'Hilux', anio: 2019, cliente: 'Juan Perez', combustible: 'Diesel' }
  ];

  const idsVehiculos = {};
  vehiculosData.forEach(function (v) {
    const id = nextId('Vehiculos', 'V');
    const token = generarTokenUnico();
    idsVehiculos[v.patente] = id;
    sheetVehiculos.appendRow([id, token, fechaAlta, true, v.patente, v.marca, v.modelo, v.anio, v.combustible, idsClientes[v.cliente], '']);
  });

  const serviciosData = [
    { patente: 'AB123CD', descripcion: 'Service completo', km: 35000, dias: 60, repuestos: 'Aceite 5W30 sintetico, filtro de aceite, filtro de aire', proximo: 'Diciembre 2026 o 45.000 km', observaciones: '' },
    { patente: 'AB123CD', descripcion: 'Cambio de aceite', km: 40000, dias: 30, repuestos: 'Aceite 5W30 4L, filtro de aceite', proximo: 'Octubre 2026 o 50.000 km', observaciones: '' },
    { patente: 'AC456EF', descripcion: 'Service completo', km: 60000, dias: 180, repuestos: 'Aceite, filtros multiples, revision general', proximo: '', observaciones: '' },
    { patente: 'AC456EF', descripcion: 'Cambio de aceite y filtro', km: 65000, dias: 90, repuestos: 'Aceite 5W30, filtro de aceite, filtro de aire', proximo: '', observaciones: '' },
    { patente: 'AD789GH', descripcion: 'Service completo', km: 90000, dias: 365, repuestos: 'Aceite, filtros, bujias, correa de distribucion', proximo: '', observaciones: 'Vehiculo con mucho uso urbano' },
    { patente: 'AD789GH', descripcion: 'Reparacion de frenos', km: 88000, dias: 200, repuestos: 'Pastillas delanteras y traseras, discos', proximo: '', observaciones: 'Cambio completo del sistema de frenos' },
    { patente: 'AE012IJ', descripcion: 'Service inicial', km: 15000, dias: 45, repuestos: 'Aceite, filtro de aceite', proximo: 'Octubre 2026 o 20.000 km', observaciones: '' },
    { patente: 'AF345KL', descripcion: 'Service inicial', km: 5000, dias: 15, repuestos: 'Aceite, filtro de aceite', proximo: 'Agosto 2026 o 10.000 km', observaciones: '' }
  ];

  const idsServicios = [];
  serviciosData.forEach(function (s) {
    const id = nextId('Servicios', 'S');
    const idVehiculo = idsVehiculos[s.patente];
    idsServicios.push(id);
    sheetServicios.appendRow([id, hace(s.dias), idVehiculo, s.km, s.descripcion, s.repuestos, s.proximo, s.observaciones, '', '']);
  });

  return {
    ok: true,
    mensaje: 'Datos demo cargados correctamente',
    clientesCreados: Object.keys(idsClientes).length,
    vehiculosCreados: Object.keys(idsVehiculos).length,
    serviciosCreados: idsServicios.length,
    clientes: idsClientes,
    vehiculos: idsVehiculos,
    servicios: idsServicios
  };
}

const SESION_TTL_SEGUNDOS = 24 * 60 * 60;

function obtenerMecanicoPassword() {
  return PropertiesService.getScriptProperties().getProperty('MECANICO_PASSWORD');
}

function loginMecanico(password) {
  const esperado = obtenerMecanicoPassword();
  if (!esperado) return { ok: false, error: 'Sistema no inicializado. Corré setupInicial.' };
  if (password !== esperado) return { ok: false, error: 'Contraseña incorrecta.' };
  const token = generarTokenUnico();
  const cache = CacheService.getScriptCache();
  cache.put('sesion_' + token, JSON.stringify({ ts: Date.now() }), SESION_TTL_SEGUNDOS);
  return { ok: true, token: token, expiraEn: SESION_TTL_SEGUNDOS };
}

function validarSesion(token) {
  if (!token || !/^[a-zA-Z0-9]{16}$/.test(token)) return false;
  const cache = CacheService.getScriptCache();
  const data = cache.get('sesion_' + token);
  return data !== null;
}

function cerrarSesion(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('sesion_' + token);
}

function serveAdminPage(sessionToken) {
  if (!validarSesion(sessionToken)) {
    return serveLoginPage();
  }
  const t = HtmlService.createTemplateFromFile('AdminPage');
  t.webappUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle('Panel del Taller')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function serveLoginPage() {
  const t = HtmlService.createTemplateFromFile('LoginPage');
  t.webappUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle('Acceso - Panel del Taller')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function serveQRPrintPage(idVehiculo, sessionToken) {
  if (!validarSesion(sessionToken)) return serveError('Sesión inválida');
  const vehiculo = buscarVehiculoPorId(idVehiculo);
  if (!vehiculo) return serveError('Vehículo no encontrado');
  const t = HtmlService.createTemplateFromFile('QRPrintPage');
  t.vehiculo = vehiculo;
  if (vehiculo.ID_Cliente) {
    const clientes = sheetAObjetos('Clientes');
    const cli = clientes.find(c => c.ID_Cliente === vehiculo.ID_Cliente);
    vehiculo.clienteNombre = cli ? cli.Nombre : '';
  }
  t.urlCliente = ScriptApp.getService().getUrl() + '?t=' + vehiculo.Token;
  return t.evaluate()
    .setTitle('QR - ' + vehiculo.Patente)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function buscarVehiculoPorId(id) {
  const objs = sheetAObjetos('Vehiculos');
  return objs.find(v => v.ID_Vehiculo === id) || null;
}

function apiMecano(e) {
  try {
    const action = e.parameter.action || '';
    const session = e.parameter.session || '';
    let payload = {};
    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch (err) {}
    }

    if (action === 'login') {
      const password = payload.password || e.parameter.password || '';
      return jsonOutput(apiLogin(password));
    }

    if (!validarSesion(session)) {
      return jsonOutput({ ok: false, error: 'Sesión inválida o expirada' });
    }

    let resultado;
    try {
      switch (action) {
        case 'logout':
          cerrarSesion(session);
          resultado = { ok: true };
          break;
        case 'buscarClientes':
          resultado = { ok: true, clientes: buscarClientes(payload.q || e.parameter.q || '') };
          break;
        case 'buscarVehiculos':
          resultado = { ok: true, vehiculos: buscarVehiculos(payload.q || e.parameter.q || '') };
          break;
        case 'historialVehiculo':
          resultado = historialVehiculo(e.parameter.idVehiculo || payload.idVehiculo);
          break;
        case 'altaCliente':
          resultado = altaClienteInterno(payload);
          break;
        case 'altaVehiculo':
          resultado = altaVehiculoInterno(payload);
          break;
        case 'altaServicio':
          resultado = altaServicioInterno(payload);
          break;
        case 'subirFoto':
          resultado = subirFoto(payload);
          break;
        case 'dashboard':
          resultado = obtenerDashboard();
          break;
        default:
          resultado = { ok: false, error: 'Acción desconocida: ' + action };
      }
    } catch (innerErr) {
      resultado = { ok: false, error: 'Error procesando ' + action + ': ' + (innerErr.message || innerErr) };
    }
    if (!resultado || typeof resultado !== 'object') {
      resultado = { ok: false, error: 'Respuesta inválida del servidor' };
    }
    return jsonOutput(resultado);
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function altaClienteInterno(p) {
  try {
    if (!p || !p.nombre) return { ok: false, error: 'Falta el nombre del cliente' };
    const id = nextId('Clientes', 'C');
    const sheet = SpreadsheetApp.openById(obtenerSheetId()).getSheetByName('Clientes');
    sheet.appendRow([id, p.nombre, p.telefono || '', p.email || '', new Date(), p.consentimiento === true]);
    return { ok: true, id: id };
  } catch (err) {
    return { ok: false, error: 'No se pudo crear el cliente: ' + (err && err.message ? err.message : err) };
  }
}

function altaVehiculoInterno(p) {
  try {
    if (!p || !p.patente || !p.marca || !p.modelo) {
      return { ok: false, error: 'Patente, marca y modelo son obligatorios' };
    }
    if (!p.idCliente) return { ok: false, error: 'Falta el cliente asociado' };
    const token = generarTokenUnico();
    const id = nextId('Vehiculos', 'V');
    const sheet = SpreadsheetApp.openById(obtenerSheetId()).getSheetByName('Vehiculos');
    sheet.appendRow([id, token, new Date(), true, (p.patente || '').toUpperCase(), p.marca, p.modelo, p.anio || '', p.combustible || '', p.idCliente, '']);
    return { ok: true, id: id, token: token, urlCliente: ScriptApp.getService().getUrl() + '?t=' + token };
  } catch (err) {
    return { ok: false, error: 'No se pudo crear el vehículo: ' + (err && err.message ? err.message : err) };
  }
}

function altaServicioInterno(p) {
  try {
    if (!p || !p.idVehiculo) return { ok: false, error: 'Falta el vehículo' };
    if (!p.descripcion) return { ok: false, error: 'La descripción es obligatoria' };
    const id = nextId('Servicios', 'S');
    const sheet = SpreadsheetApp.openById(obtenerSheetId()).getSheetByName('Servicios');
    const fecha = p.fecha ? new Date(p.fecha) : new Date();
    const fotos = Array.isArray(p.fotosIds) ? p.fotosIds.join(',') : (p.fotosIds || '');
    sheet.appendRow([id, fecha, p.idVehiculo, p.kilometraje || '', p.descripcion, p.repuestos || '', p.proximoMantenimiento || '', p.observaciones || '', fotos, '']);
    return { ok: true, id: id };
  } catch (err) {
    return { ok: false, error: 'No se pudo guardar el servicio: ' + (err && err.message ? err.message : err) };
  }
}

function esActivo(valor) {
  if (valor === true || valor === false) return valor === true;
  if (valor === null || valor === undefined) return true;
  const s = String(valor).trim().toUpperCase();
  if (s === '' || s === 'NULL' || s === 'UNDEFINED') return true;
  if (s === 'TRUE' || s === 'VERDADERO' || s === 'SI' || s === 'S' || s === '1' || s === 'ACTIVO') return true;
  if (s === 'FALSE' || s === 'FALSO' || s === 'NO' || s === 'N' || s === '0' || s === 'INACTIVO') return false;
  return true;
}

function normalizarQ(q) {
  if (q === null || q === undefined) return '';
  return String(q).trim();
}

function buscarClientes(q) {
  try {
    const clientes = sheetAObjetos('Clientes');
    const qs = normalizarQ(q);
    if (!qs) return clientes.slice(0, 20);
    const ql = qs.toLowerCase();
    return clientes.filter(c =>
      (c.Nombre || '').toString().toLowerCase().includes(ql) ||
      (c.Telefono || '').toString().toLowerCase().includes(ql) ||
      (c.ID_Cliente || '').toString().toLowerCase().includes(ql) ||
      (c.Email || '').toString().toLowerCase().includes(ql)
    ).slice(0, 20);
  } catch (err) {
    Logger.log('Error en buscarClientes: ' + (err && err.message ? err.message : err));
    return [];
  }
}

function buscarVehiculos(q) {
  try {
    const todos = sheetAObjetos('Vehiculos');
    const vehiculos = todos.filter(function (v) { return esActivo(v.Activo); });
    const clientes = sheetAObjetos('Clientes');
    const clienteMap = {};
    clientes.forEach(function (c) { clienteMap[c.ID_Cliente] = c.Nombre; });
    vehiculos.forEach(function (v) { v.clienteNombre = clienteMap[v.ID_Cliente] || ''; });
    const qs = normalizarQ(q);
    if (!qs) return vehiculos.slice(0, 30);
    const ql = qs.toLowerCase();
    return vehiculos.filter(function (v) {
      return (v.Patente || '').toString().toLowerCase().includes(ql) ||
        (v.Marca || '').toString().toLowerCase().includes(ql) ||
        (v.Modelo || '').toString().toLowerCase().includes(ql) ||
        (v.Combustible || '').toString().toLowerCase().includes(ql) ||
        (v.clienteNombre || '').toString().toLowerCase().includes(ql) ||
        (v.Token || '').toString().toLowerCase().includes(ql) ||
        (v.ID_Vehiculo || '').toString().toLowerCase().includes(ql);
    }).slice(0, 30);
  } catch (err) {
    Logger.log('Error en buscarVehiculos: ' + (err && err.message ? err.message : err));
    return [];
  }
}

function historialVehiculo(idVehiculo) {
  try {
    if (!idVehiculo) return { ok: false, error: 'Falta idVehiculo' };
    const vehiculo = buscarVehiculoPorId(idVehiculo);
    if (!vehiculo) return { ok: false, error: 'Vehículo no encontrado' };
    const servicios = listarServicios(idVehiculo);
    const cliente = sheetAObjetos('Clientes').find(c => c.ID_Cliente === vehiculo.ID_Cliente);
    return { ok: true, vehiculo: vehiculo, servicios: servicios, cliente: cliente || null };
  } catch (err) {
    return { ok: false, error: 'No se pudo cargar el historial: ' + (err && err.message ? err.message : err) };
  }
}

function subirFoto(p) {
  try {
    if (!p || !p.base64 || !p.nombreArchivo) return { ok: false, error: 'Faltan datos de la foto' };
    const folder = DriveApp.getFolderById(obtenerFotosFolderId());
    const bytes = Utilities.base64Decode(p.base64);
    const blob = Utilities.newBlob(bytes, p.mimeType || 'image/jpeg', p.nombreArchivo);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { ok: true, id: file.getId(), url: 'https://drive.google.com/uc?id=' + file.getId() };
  } catch (err) {
    return { ok: false, error: 'No se pudo subir la foto: ' + (err && err.message ? err.message : err) };
  }
}

function obtenerDashboard() {
  try {
    const clientes = sheetAObjetos('Clientes');
    const vehiculos = sheetAObjetos('Vehiculos');
    const servicios = sheetAObjetos('Servicios');
    const hace30Dias = Date.now() - 30 * 24 * 60 * 60 * 1000;
    const serviciosRecientes = servicios.filter(s => {
      const f = s.Fecha instanceof Date ? s.Fecha.getTime() : new Date(s.Fecha).getTime();
      return !isNaN(f) && f >= hace30Dias;
    }).length;
    return { ok: true, totalClientes: clientes.length, totalVehiculos: vehiculos.length, totalServicios: servicios.length, serviciosUltimoMes: serviciosRecientes };
  } catch (err) {
    return { ok: false, error: 'No se pudo cargar el dashboard: ' + (err && err.message ? err.message : err) };
  }
}

function apiLogin(password) {
  try {
    return loginMecanico(password) || { ok: false, error: 'No se pudo procesar el login' };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function apiAccion(action, payloadJson, session) {
  try {
    if (!validarSesion(session || '')) {
      return { ok: false, error: 'Sesión inválida o expirada. Recargá la página.' };
    }
    let payload = {};
    if (payloadJson) {
      try { payload = JSON.parse(payloadJson); } catch (err) {}
    }
    let resultado;
    try {
      switch (action) {
        case 'buscarClientes': resultado = { ok: true, clientes: buscarClientes(payload.q || '') }; break;
        case 'buscarVehiculos': resultado = { ok: true, vehiculos: buscarVehiculos(payload.q || '') }; break;
        case 'historialVehiculo': resultado = historialVehiculo(payload.idVehiculo || ''); break;
        case 'altaCliente': resultado = altaClienteInterno(payload); break;
        case 'altaVehiculo': resultado = altaVehiculoInterno(payload); break;
        case 'altaServicio': resultado = altaServicioInterno(payload); break;
        case 'subirFoto': resultado = subirFoto(payload); break;
        case 'dashboard': resultado = obtenerDashboard(); break;
        case 'logout': cerrarSesion(session); resultado = { ok: true }; break;
        default: resultado = { ok: false, error: 'Acción desconocida: ' + action };
      }
    } catch (innerErr) {
      resultado = { ok: false, error: 'Error procesando ' + action + ': ' + (innerErr.message || innerErr) };
    }
    if (!resultado || typeof resultado !== 'object') {
      resultado = { ok: false, error: 'Respuesta inválida del servidor' };
    }
    if (!('ok' in resultado)) {
      resultado = { ok: false, error: 'Respuesta sin estado válido' };
    }
    return resultado;
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function serveDiagnostico() {
  try {
    const props = PropertiesService.getScriptProperties();
    const sheetId = props.getProperty('SHEET_ID');

    let html = '<!doctype html><html><head><meta charset="utf-8"><title>Diagnóstico del Sistema</title><style>' +
      'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;background:#f8fafc;padding:24px;color:#0f172a;max-width:900px;margin:0 auto}' +
      'h1{color:#1e40af}' +
      'h2{color:#334155;margin-top:32px;border-bottom:2px solid #e2e8f0;padding-bottom:8px}' +
      '.card{background:#fff;padding:16px 20px;border-radius:12px;margin:12px 0;box-shadow:0 1px 3px rgba(0,0,0,0.05)}' +
      'table{border-collapse:collapse;width:100%;margin-top:8px;font-size:13px}' +
      'th,td{border:1px solid #e2e8f0;padding:6px 10px;text-align:left}' +
      'th{background:#f1f5f9}' +
      '.ok{color:#10b981;font-weight:600}' +
      '.err{color:#ef4444;font-weight:600}' +
      'pre{background:#1e293b;color:#f1f5f9;padding:12px;border-radius:8px;overflow-x:auto;font-size:12px}' +
      '</style></head><body>';

    html += '<h1>Diagnóstico del Sistema</h1>';

    html += '<div class="card"><h2>Backend</h2>';
    html += '<p><b>SHEET_ID:</b> ' + (sheetId || '<span class="err">NO CONFIGURADO</span>') + '</p>';
    html += '<p><b>FOTOS_FOLDER_ID:</b> ' + (props.getProperty('FOTOS_FOLDER_ID') || '<span class="err">NO CONFIGURADO</span>') + '</p>';
    html += '<p><b>DRIVE_QR_FOLDER_ID:</b> ' + (props.getProperty('DRIVE_QR_FOLDER_ID') || '<span class="err">NO CONFIGURADO</span>') + '</p>';
    html += '<p><b>ADMIN_KEY:</b> ' + ((props.getProperty('ADMIN_KEY') || '').substring(0, 4)) + '...</p>';
    html += '<p><b>MECANICO_PASSWORD:</b> ' + ((props.getProperty('MECANICO_PASSWORD') || '').substring(0, 4)) + '...</p>';
    html += '<p><b>Hora del servidor:</b> ' + new Date().toISOString() + '</p>';
    html += '</div>';

    html += '<div class="card"><h2>Hojas</h2>';
    try {
      const ss = SpreadsheetApp.openById(sheetId);
      const sheetNames = ['Clientes', 'Vehiculos', 'Servicios', 'AccesosLog'];
      for (let i = 0; i < sheetNames.length; i++) {
        const name = sheetNames[i];
        const sheet = ss.getSheetByName(name);
        if (!sheet) {
          html += '<h3>' + escapeHtml(name) + '</h3><p class="err">NO EXISTE</p>';
          continue;
        }
        const data = sheet.getDataRange().getValues();
        html += '<h3>' + escapeHtml(name) + ' (' + (data.length - 1) + ' filas)</h3>';
        if (data.length > 0) {
          html += '<table><tr>';
          for (let c = 0; c < data[0].length; c++) {
            html += '<th>' + escapeHtml(String(data[0][c])) + '</th>';
          }
          html += '</tr>';
          const lastN = Math.min(data.length, 6);
          for (let r = 1; r < lastN; r++) {
            html += '<tr>';
            for (let c = 0; c < data[r].length; c++) {
              const val = data[r][c];
              const display = val instanceof Date ? val.toISOString() : String(val);
              html += '<td>' + escapeHtml(display) + '</td>';
            }
            html += '</tr>';
          }
          html += '</table>';
        } else {
          html += '<p>Hoja vacía</p>';
        }
      }
    } catch (e) {
      html += '<p class="err">Error leyendo hojas: ' + escapeHtml(String(e && e.message ? e.message : e)) + '</p>';
    }
    html += '</div>';

    html += '<div class="card"><h2>Sesiones activas</h2>';
    try {
      const cache = CacheService.getScriptCache();
      const all = (cache.getAll && cache.getAll()) || {};
      const sessionKeys = Object.keys(all).filter(function (k) { return k.indexOf('sesion_') === 0; });
      html += '<p>' + sessionKeys.length + ' sesiones activas</p>';
      if (sessionKeys.length > 0) {
        html += '<ul>';
        const limite = Math.min(sessionKeys.length, 10);
        for (let i = 0; i < limite; i++) {
          html += '<li>' + escapeHtml(sessionKeys[i]) + '</li>';
        }
        html += '</ul>';
      }
    } catch (e) {
      html += '<p class="err">Error leyendo cache: ' + escapeHtml(String(e && e.message ? e.message : e)) + '</p>';
    }
    html += '</div>';

    html += '<div class="card"><h2>Test backend</h2>';
    try {
      const vehiculos = buscarVehiculos('');
      html += '<p>buscarVehiculos("") devolvió ' + vehiculos.length + ' vehículos</p>';
      if (vehiculos.length > 0) {
        html += '<ul>';
        const limite = Math.min(vehiculos.length, 3);
        for (let i = 0; i < limite; i++) {
          const v = vehiculos[i];
          html += '<li>' + escapeHtml(JSON.stringify({
            ID_Vehiculo: v.ID_Vehiculo,
            Patente: v.Patente,
            Marca: v.Marca,
            Modelo: v.Modelo,
            Activo: v.Activo
          })) + '</li>';
        }
        html += '</ul>';
      }
    } catch (e) {
      html += '<p class="err">Error en buscarVehiculos: ' + escapeHtml(String(e && e.message ? e.message : e)) + '</p>';
    }
    html += '</div>';

    html += '</body></html>';

    return HtmlService.createHtmlOutput(html)
      .setTitle('Diagnóstico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<h1>Error en diagnóstico</h1><pre>' + escapeHtml(String(err && err.message ? err.message : err)) + '</pre>'
    );
  }
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (m) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
  });
}

function normalizarDatos() {
  const ss = SpreadsheetApp.openById(obtenerSheetId());
  const reporte = { eliminados: { clientes: 0, vehiculos: 0, servicios: 0 }, idsMalformados: [] };

  const sheetClientes = ss.getSheetByName('Clientes');
  if (sheetClientes) {
    const dataClientes = sheetClientes.getDataRange().getValues();
    for (let i = dataClientes.length - 1; i >= 1; i--) {
      const fila = dataClientes[i];
      const id = String(fila[0] || '').trim();
      const nombre = String(fila[1] || '').trim();
      if (!nombre && !String(fila[2] || '').trim()) {
        sheetClientes.deleteRow(i + 1);
        reporte.eliminados.clientes++;
      }
    }
  }

  const sheetVehiculos = ss.getSheetByName('Vehiculos');
  if (sheetVehiculos) {
    const dataVehiculos = sheetVehiculos.getDataRange().getValues();
    for (let i = dataVehiculos.length - 1; i >= 1; i--) {
      const fila = dataVehiculos[i];
      const id = String(fila[0] || '').trim();
      if (!/^[CVS]-\d+$/.test(id)) {
        reporte.idsMalformados.push({ hoja: 'Vehiculos', id: id, fila: i + 1 });
        sheetVehiculos.deleteRow(i + 1);
        reporte.eliminados.vehiculos++;
      }
    }
  }

  const sheetServicios = ss.getSheetByName('Servicios');
  if (sheetServicios) {
    const dataServicios = sheetServicios.getDataRange().getValues();
    const vehiculosRestantes = sheetVehiculos.getDataRange().getValues().slice(1).map(r => String(r[0]).trim());
    for (let i = dataServicios.length - 1; i >= 1; i--) {
      const fila = dataServicios[i];
      const id = String(fila[0] || '').trim();
      const idVehiculo = String(fila[2] || '').trim();
      if (!/^[CVS]-\d+$/.test(id) || !vehiculosRestantes.includes(idVehiculo)) {
        reporte.idsMalformados.push({ hoja: 'Servicios', id: id, fila: i + 1 });
        sheetServicios.deleteRow(i + 1);
        reporte.eliminados.servicios++;
      }
    }
  }

  return reporte;
}

function migrarCombustible() {
  const ss = SpreadsheetApp.openById(obtenerSheetId());
  const sheet = ss.getSheetByName('Vehiculos');
  if (!sheet) return { ok: false, error: 'No existe la hoja Vehiculos' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const vinColumnIndex = headers.indexOf('VIN');
  const combustibleColumnIndex = headers.indexOf('Combustible');

  if (combustibleColumnIndex !== -1 && vinColumnIndex === -1) {
    return { ok: true, mensaje: 'La columna Combustible ya existe, nada que migrar', accion: 'ninguna' };
  }

  if (vinColumnIndex === -1 && combustibleColumnIndex === -1) {
    return { ok: false, error: 'No se encontro ninguna columna VIN o Combustible. Headers actuales: ' + JSON.stringify(headers) };
  }

  if (vinColumnIndex !== -1) {
    sheet.getRange(1, vinColumnIndex + 1).setValue('Combustible');
    return { ok: true, mensaje: 'Header VIN renombrado a Combustible', accion: 'header actualizado' };
  }

  return { ok: true, mensaje: 'Sin cambios necesarios' };
}

function cargarDatosDemoMasivos() {
  const ss = SpreadsheetApp.openById(obtenerSheetId());
  const sheetClientes = ss.getSheetByName('Clientes');
  const sheetVehiculos = ss.getSheetByName('Vehiculos');
  const sheetServicios = ss.getSheetByName('Servicios');

  const dataClientes = sheetClientes.getLastRow() > 1 ? sheetClientes.getDataRange().getValues().slice(1) : [];
  const mapaClientes = {};
  dataClientes.forEach(function (f) {
    const nombre = String(f[1] || '').trim();
    if (nombre) mapaClientes[nombre] = f[0];
  });

  const dataVehiculos = sheetVehiculos.getLastRow() > 1 ? sheetVehiculos.getDataRange().getValues().slice(1) : [];
  const mapaVehiculos = {};
  dataVehiculos.forEach(function (f) {
    const patente = String(f[4] || '').trim();
    if (patente) mapaVehiculos[patente] = { id: f[0], token: f[1] };
  });

  const dataServicios = sheetServicios.getLastRow() > 1 ? sheetServicios.getDataRange().getValues().slice(1) : [];
  const setServiciosExistentes = new Set();
  dataServicios.forEach(function (f) {
    const idV = String(f[2] || '').trim();
    const fecha = String(f[1] || '').trim();
    const desc = String(f[4] || '').trim();
    setServiciosExistentes.add(idV + '|' + fecha + '|' + desc);
  });

  const clientesData = [
    { nombre: 'Juan Perez', telefono: '+5491144551122', email: 'juan.perez@gmail.com', consentimiento: true },
    { nombre: 'Maria Gonzalez', telefono: '+5491155663344', email: 'maria.gonzalez@hotmail.com', consentimiento: true },
    { nombre: 'Carlos Rodriguez', telefono: '+5491166778899', email: '', consentimiento: true },
    { nombre: 'Ana Martinez', telefono: '+5491177889900', email: 'ana.martinez@yahoo.com.ar', consentimiento: true },
    { nombre: 'Pedro Lopez', telefono: '+5491188990011', email: 'pedro.lopez@outlook.com', consentimiento: true },
    { nombre: 'Lucia Fernandez', telefono: '+5493415556677', email: '', consentimiento: false },
    { nombre: 'Diego Sanchez', telefono: '+5491144112233', email: 'diego.sanchez@gmail.com', consentimiento: true },
    { nombre: 'Carolina Romero', telefono: '+5492215554433', email: '', consentimiento: false },
    { nombre: 'Sebastian Acosta', telefono: '+5493515559988', email: 'sebastian.acosta@hotmail.com', consentimiento: true },
    { nombre: 'Valeria Diaz', telefono: '+5491155443322', email: 'valeria.diaz@yahoo.com.ar', consentimiento: true },
    { nombre: 'Miguel Angel Torres', telefono: '+5492215557766', email: '', consentimiento: false },
    { nombre: 'Gabriela Suarez', telefono: '+5491144332211', email: 'gabriela.suarez@gmail.com', consentimiento: true },
    { nombre: 'Roberto Castro', telefono: '+5491144556677', email: '', consentimiento: false },
    { nombre: 'Patricia Morales', telefono: '+5493415554433', email: 'patricia.morales@hotmail.com', consentimiento: true },
    { nombre: 'Fernando Ortiz', telefono: '+5491144667788', email: 'fernando.ortiz@outlook.com', consentimiento: true },
    { nombre: 'Silvina Vega', telefono: '+5493515553322', email: '', consentimiento: false },
    { nombre: 'Hector Ramos', telefono: '+5491144778855', email: 'hector.ramos@gmail.com', consentimiento: true },
    { nombre: 'Norma Mendez', telefono: '+5492215559988', email: '', consentimiento: false },
    { nombre: 'Ariel Juarez', telefono: '+5491144889966', email: 'ariel.juarez@yahoo.com.ar', consentimiento: true },
    { nombre: 'Monica Pereira', telefono: '+5491144997755', email: '', consentimiento: false }
  ];

  const nombresDemo = clientesData.map(function (c) { return c.nombre; });
  const nombresCargados = nombresDemo.filter(function (n) { return mapaClientes[n]; });
  if (nombresCargados.length >= 18) {
    return {
      ok: false,
      yaCargado: true,
      mensaje: 'Los datos demo ya estan cargados (' + nombresCargados.length + ' de ' + nombresDemo.length + ' clientes demo presentes). Para recargar, primero elimina manualmente los clientes con esos nombres desde la hoja Clientes y los vehiculos asociados en Vehiculos.'
    };
  }

  const vehiculosData = [
    { patente: 'AB123CD', marca: 'Toyota', modelo: 'Corolla', anio: 2020, combustible: 'Nafta', cliente: 'Juan Perez' },
    { patente: 'AC456EF', marca: 'Ford', modelo: 'Focus', anio: 2018, combustible: 'Nafta', cliente: 'Maria Gonzalez' },
    { patente: 'AD789GH', marca: 'Volkswagen', modelo: 'Gol', anio: 2015, combustible: 'GNC', cliente: 'Carlos Rodriguez' },
    { patente: 'AE012IJ', marca: 'Chevrolet', modelo: 'Onix', anio: 2022, combustible: 'Nafta', cliente: 'Ana Martinez' },
    { patente: 'AF345KL', marca: 'Renault', modelo: 'Kwid', anio: 2021, combustible: 'Nafta', cliente: 'Pedro Lopez' },
    { patente: 'AG678MN', marca: 'Toyota', modelo: 'Hilux', anio: 2019, combustible: 'Diesel', cliente: 'Juan Perez' },
    { patente: 'AA111AA', marca: 'Ford', modelo: 'Ranger', anio: 2017, combustible: 'Diesel', cliente: 'Diego Sanchez' },
    { patente: 'AB222BB', marca: 'Fiat', modelo: 'Cronos', anio: 2023, combustible: 'Nafta', cliente: 'Carolina Romero' },
    { patente: 'AC333CC', marca: 'Peugeot', modelo: '208', anio: 2020, combustible: 'Nafta', cliente: 'Sebastian Acosta' },
    { patente: 'AD444DD', marca: 'Volkswagen', modelo: 'Amarok', anio: 2021, combustible: 'Diesel', cliente: 'Valeria Diaz' },
    { patente: 'AE555EE', marca: 'Renault', modelo: 'Duster', anio: 2018, combustible: 'Nafta + GNC', cliente: 'Miguel Angel Torres' },
    { patente: 'AF666FF', marca: 'Citroen', modelo: 'C3', anio: 2016, combustible: 'Nafta', cliente: 'Gabriela Suarez' },
    { patente: 'AG777GG', marca: 'Honda', modelo: 'HR-V', anio: 2022, combustible: 'Nafta', cliente: 'Roberto Castro' },
    { patente: 'AH888HH', marca: 'Toyota', modelo: 'Etios', anio: 2017, combustible: 'Nafta', cliente: 'Patricia Morales' },
    { patente: 'AI999II', marca: 'Ford', modelo: 'EcoSport', anio: 2019, combustible: 'Nafta', cliente: 'Fernando Ortiz' },
    { patente: 'AJ000JJ', marca: 'Chevrolet', modelo: 'Cruze', anio: 2021, combustible: 'Nafta', cliente: 'Silvina Vega' },
    { patente: 'AK111KK', marca: 'Nissan', modelo: 'Frontier', anio: 2018, combustible: 'Diesel', cliente: 'Hector Ramos' },
    { patente: 'AL222LL', marca: 'Renault', modelo: 'Sandero', anio: 2015, combustible: 'Nafta + GNC', cliente: 'Norma Mendez' },
    { patente: 'AM333MM', marca: 'Fiat', modelo: 'Palio', anio: 2014, combustible: 'Nafta', cliente: 'Ariel Juarez' },
    { patente: 'AN444NN', marca: 'Peugeot', modelo: '308', anio: 2017, combustible: 'Nafta', cliente: 'Monica Pereira' },
    { patente: 'AO555OO', marca: 'Volkswagen', modelo: 'Polo', anio: 2022, combustible: 'Nafta', cliente: 'Juan Perez' },
    { patente: 'AP666PP', marca: 'Jeep', modelo: 'Renegade', anio: 2020, combustible: 'Nafta', cliente: 'Maria Gonzalez' },
    { patente: 'AQ777QQ', marca: 'Toyota', modelo: 'Yaris', anio: 2023, combustible: 'Hibrido', cliente: 'Carlos Rodriguez' },
    { patente: 'AR888RR', marca: 'Chevrolet', modelo: 'Prisma', anio: 2018, combustible: 'Nafta + GNC', cliente: 'Ana Martinez' },
    { patente: 'AS999SS', marca: 'Hyundai', modelo: 'Tucson', anio: 2019, combustible: 'Nafta', cliente: 'Pedro Lopez' },
    { patente: 'AT000TT', marca: 'Renault', modelo: 'Logan', anio: 2016, combustible: 'Nafta + GNC', cliente: 'Lucia Fernandez' },
    { patente: 'AU111UU', marca: 'Ford', modelo: 'Ka', anio: 2019, combustible: 'Nafta', cliente: 'Diego Sanchez' },
    { patente: 'AV222VV', marca: 'Citroen', modelo: 'C4', anio: 2014, combustible: 'Nafta', cliente: 'Carolina Romero' },
    { patente: 'AW333WW', marca: 'Fiat', modelo: 'Mobi', anio: 2021, combustible: 'Nafta', cliente: 'Sebastian Acosta' },
    { patente: 'AX444XX', marca: 'Volkswagen', modelo: 'Vento', anio: 2020, combustible: 'Nafta', cliente: 'Valeria Diaz' }
  ];

  const hace = (dias) => {
    const fecha = new Date();
    fecha.setDate(fecha.getDate() - dias);
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  };

  const fechaAlta = ahoraComoString();

  let clientesCreados = 0;
  clientesData.forEach(function (c) {
    if (mapaClientes[c.nombre]) return;
    const id = nextId('Clientes', 'C');
    sheetClientes.appendRow([id, c.nombre, c.telefono, c.email, fechaAlta, c.consentimiento]);
    mapaClientes[c.nombre] = id;
    clientesCreados++;
  });

  let vehiculosCreados = 0;
  vehiculosData.forEach(function (v) {
    if (mapaVehiculos[v.patente]) return;
    const id = nextId('Vehiculos', 'V');
    const token = generarTokenUnico();
    sheetVehiculos.appendRow([id, token, fechaAlta, 'TRUE', v.patente, v.marca, v.modelo, v.anio, v.combustible, mapaClientes[v.cliente], '']);
    mapaVehiculos[v.patente] = { id: id, token: token };
    vehiculosCreados++;
  });

  const servAntiguo = [
    { desc: 'Cambio de aceite y filtro', rep: 'Aceite Shell Helix 5W40 4L, Filtro de aceite Mann W712/75, Filtro de aire Mann C25114', prox: 'Cada 10.000 km o 6 meses', obs: '' },
    { desc: 'Service completo 60.000 km', rep: 'Aceite Total Quartz 9000 5W40 5L, Filtro aceite Wix 51348, Filtro aire Bosch F026400004, Filtro habitaculo Wix WP9322, Bujias NGK LZKAR7A', prox: 'En 20.000 km', obs: 'Service segun manual del fabricante' },
    { desc: 'Cambio de pastillas de freno delanteras', rep: 'Pastillas Brembo P85020, Liquido de frenos DOT 4 1L', prox: 'Revision en 20.000 km', obs: '' },
    { desc: 'Cambio de discos y pastillas', rep: 'Discos de freno Fremax, Pastillas Bosch BP1100, Liquido DOT 4', prox: 'En 40.000 km', obs: 'Sistema de frenos en buen estado' },
    { desc: 'Reparacion de embrague', rep: 'Kit de embrague Valeo (disco, plato, ruleman), Liquido hidraulico', prox: '', obs: 'Embrague con desgaste normal por uso' },
    { desc: 'Cambio de kit de distribucion', rep: 'Kit distribucion Gates K015457XS (correa, rodillos, bomba de agua)', prox: 'En 80.000 km', obs: 'Se reemplazo tambien la bomba de agua' },
    { desc: 'Cambio de amortiguadores delanteros', rep: 'Amortiguadores Monroe G7308 x2, Cazoletas y bujes', prox: 'Revision en 50.000 km', obs: '' },
    { desc: 'Alineacion y balanceo', rep: '', prox: 'Cada 10.000 km', obs: 'Alineacion computarizada y balanceo de 4 ruedas' },
    { desc: 'Cambio de liquido de frenos', rep: 'Liquido de frenos DOT 4 sintetico 1L', prox: 'Cada 2 anos', obs: 'Purgado completo del sistema' },
    { desc: 'Reparacion de tren delantero', rep: 'Precaps, rotulas, barras, bujes (kit completo)', prox: 'Revision en 30.000 km', obs: 'Golpes detectados en inspeccion' },
    { desc: 'Cambio de bujias', rep: 'Bujias NGK BKR6E x4', prox: 'Cada 30.000 km', obs: '' },
    { desc: 'Service de GNC', rep: 'Regulacion de mezcla, cambio de filtros GNC, prueba de cilindros', prox: 'Cada 10.000 km o 1 ano', obs: 'Sistema GNC funcionando correctamente' },
    { desc: 'Cambio de filtro de aire y habitaculo', rep: 'Filtro de aire Mann C25114, Filtro habitaculo Bosch A8531', prox: 'Cada 20.000 km', obs: '' },
    { desc: 'Diagnostico electronico', rep: '', prox: '', obs: 'Escaneo con scanner OBD-II, sin codigos de error activos' },
    { desc: 'Reparacion de aire acondicionado', rep: 'Recarga de gas R134a 600g, filtro secador nuevo', prox: 'Carga de gas cada 2 anos', obs: 'Se verifico compresor y condensador' }
  ];

  const servReciente = [
    { desc: 'Cambio de aceite y filtro', rep: 'Aceite Mobil 1 ESP 5W30 4L, Filtro de aceite Wix 51348, Filtro de aire Mann', prox: 'Cada 10.000 km o 6 meses', obs: 'Vehiculo en buen estado general' },
    { desc: 'Service completo 100.000 km', rep: 'Aceite Shell Helix Ultra 5W40 5L, Filtros Mann completos, Bujias NGK Iridium, Correa accesorios Gates', prox: 'En 20.000 km', obs: 'Service mayor cumplido' },
    { desc: 'Cambio de pastillas de freno', rep: 'Pastillas Brembo P85025, Liquido DOT 4 sint 1L', prox: 'Revision en 25.000 km', obs: 'Desgaste normal de pastillas' },
    { desc: 'Cambio de discos y pastillas (ambos ejes)', rep: 'Discos Fremax delanteros y traseros, Pastillas Bosch, Liquido freno DOT 4', prox: 'En 50.000 km', obs: '' },
    { desc: 'Reparacion de embrague', rep: 'Kit embrague Sachs (disco, plato, ruleman, collarin hidraulico)', prox: '', obs: 'Embrague patinado, reparacion completa' },
    { desc: 'Cambio de kit de distribucion completo', rep: 'Kit distribucion INA, Correa, Tensor, Rodillos, Bomba de agua, Correa poly-v', prox: 'En 100.000 km', obs: 'Kit completo con bomba de agua incluida' },
    { desc: 'Cambio de amortiguadores (4 unidades)', rep: 'Amortiguadores Monroe Original x4, soportes y cazoletas', prox: 'Revision en 60.000 km', obs: 'Cambio de los 4 amortiguadores' },
    { desc: 'Alineacion, balanceo y rotacion de cubiertas', rep: '', prox: 'Cada 10.000 km', obs: 'Vehiculo alineado y cubiertas rotadas' },
    { desc: 'Cambio de liquido de frenos y purgado', rep: 'Liquido de frenos DOT 4 sintetico 1L', prox: 'En 2 anos', obs: 'Sistema purgado correctamente' },
    { desc: 'Reparacion completa de tren delantero', rep: 'Precaps, rotulas, extremos, bujes, barra estabilizadora', prox: 'En 40.000 km', obs: 'Tren delantero restaurado, geometria corregida' },
    { desc: 'Cambio de bujias y cables', rep: 'Bujias NGK Iridium ILZKR7B x4, Cables de bujia', prox: 'En 40.000 km', obs: '' },
    { desc: 'Service completo GNC', rep: 'Regulacion, filtros GNC nuevos, prueba de estanqueidad, cambio de juntas', prox: 'En 10.000 km', obs: 'Equipo GNC revisado y calibrado' },
    { desc: 'Cambio de filtro de aire, habitaculo y aceite', rep: 'Filtro aire Mann, Filtro habitaculo Bosch, Aceite Mobil 5W30 4L', prox: 'En 10.000 km o 6 meses', obs: '' },
    { desc: 'Diagnostico electronico computarizado', rep: '', prox: '', obs: 'Limpieza de cuerpo de acelerador y adaptacion' },
    { desc: 'Reparacion de aire acondicionado y carga', rep: 'Recarga gas R134a, filtro secador, valvula expansion', prox: 'Carga cada 2 anos', obs: 'A/C funcionando a 5 grados' }
  ];

  const kmBasePorVehiculo = [
    78000, 95000, 145000, 25000, 18000,
    110000, 135000, 14000, 52000, 28000,
    120000, 162000, 18000, 105000, 88000,
    35000, 138000, 175000, 195000, 92000,
    22000, 48000, 8000, 92000, 76000,
    142000, 72000, 180000, 35000, 58000
  ];

  const serviciosData = [];
  for (let i = 0; i < 30; i++) {
    const v = vehiculosData[i];
    const kmActual = kmBasePorVehiculo[i];

    const diasAntiguo = 220 + (i * 13) % 320;
    const kmAntiguo = Math.max(4000, kmActual - 6500 - (diasAntiguo * 28));

    const t1 = servAntiguo[i % servAntiguo.length];
    serviciosData.push({
      patente: v.patente,
      fecha: hace(diasAntiguo),
      km: Math.floor(kmAntiguo),
      descripcion: t1.desc + ' - ' + v.marca + ' ' + v.modelo,
      repuestos: t1.rep,
      proximoMantenimiento: t1.prox,
      observaciones: t1.obs
    });

    const diasReciente = 3 + (i * 7) % 75;
    const kmReciente = kmActual - 200 - (i * 40);

    const t2 = servReciente[(i + 7) % servReciente.length];
    serviciosData.push({
      patente: v.patente,
      fecha: hace(diasReciente),
      km: Math.floor(kmReciente),
      descripcion: t2.desc + ' - ' + v.marca + ' ' + v.modelo,
      repuestos: t2.rep,
      proximoMantenimiento: t2.prox,
      observaciones: t2.obs
    });
  }

  let serviciosCreados = 0;
  const serviciosSaltados = [];
  serviciosData.forEach(function (s) {
    const ref = mapaVehiculos[s.patente];
    if (!ref) {
      serviciosSaltados.push({ patente: s.patente, motivo: 'vehiculo no encontrado' });
      return;
    }
    const idVehiculo = ref.id;
    const clave = idVehiculo + '|' + s.fecha + '|' + s.descripcion;
    if (setServiciosExistentes.has(clave)) {
      serviciosSaltados.push({ patente: s.patente, motivo: 'duplicado' });
      return;
    }
    const id = nextId('Servicios', 'S');
    sheetServicios.appendRow([id, s.fecha, idVehiculo, s.km, s.descripcion, s.repuestos, s.proximoMantenimiento, s.observaciones, '', '']);
    setServiciosExistentes.add(clave);
    serviciosCreados++;
  });

  return {
    ok: true,
    mensaje: 'Datos demo masivos cargados correctamente',
    clientesCreados: clientesCreados,
    vehiculosCreados: vehiculosCreados,
    serviciosCreados: serviciosCreados,
    clientesTotalMapa: Object.keys(mapaClientes).length,
    vehiculosTotalMapa: Object.keys(mapaVehiculos).length,
    serviciosSaltados: serviciosSaltados.length,
    clientes: mapaClientes,
    vehiculos: mapaVehiculos,
    clientesDemoEsperados: clientesData.length,
    vehiculosDemoEsperados: vehiculosData.length,
    serviciosDemoEsperados: serviciosData.length
  };
}

function limpiarDatosDemo() {
  const ss = SpreadsheetApp.openById(obtenerSheetId());
  const sheetClientes = ss.getSheetByName('Clientes');
  const sheetVehiculos = ss.getSheetByName('Vehiculos');
  const sheetServicios = ss.getSheetByName('Servicios');

  const clientesDemo = [
    'Juan Perez', 'Maria Gonzalez', 'Carlos Rodriguez', 'Ana Martinez', 'Pedro Lopez',
    'Lucia Fernandez', 'Diego Sanchez', 'Carolina Romero', 'Sebastian Acosta', 'Valeria Diaz',
    'Miguel Angel Torres', 'Gabriela Suarez', 'Roberto Castro', 'Patricia Morales', 'Fernando Ortiz',
    'Silvina Vega', 'Hector Ramos', 'Norma Mendez', 'Ariel Juarez', 'Monica Pereira',
    'Cliente Demo'
  ];

  const patentesDemo = [
    'AB123CD', 'AC456EF', 'AD789GH', 'AE012IJ', 'AF345KL', 'AG678MN',
    'AA111AA', 'AB222BB', 'AC333CC', 'AD444DD', 'AE555EE', 'AF666FF', 'AG777GG', 'AH888HH', 'AI999II',
    'AJ000JJ', 'AK111KK', 'AL222LL', 'AM333MM', 'AN444NN', 'AO555OO', 'AP666PP', 'AQ777QQ', 'AR888RR',
    'AS999SS', 'AT000TT', 'AU111UU', 'AV222VV', 'AW333WW', 'AX444XX'
  ];

  const reporte = {
    clientesEliminados: 0,
    vehiculosEliminados: 0,
    serviciosEliminados: 0,
    clientesRestantes: 0,
    vehiculosRestantes: 0,
    serviciosRestantes: 0
  };

  const vehiculosEliminadosIds = [];
  if (sheetVehiculos) {
    const data = sheetVehiculos.getDataRange().getValues();
    const idCol = data[0].indexOf('ID_Vehiculo');
    const patenteCol = data[0].indexOf('Patente');

    for (let i = data.length - 1; i >= 1; i--) {
      const patente = String(data[i][patenteCol] || '').trim().toUpperCase();
      if (patentesDemo.indexOf(patente) !== -1) {
        vehiculosEliminadosIds.push(String(data[i][idCol]));
        sheetVehiculos.deleteRow(i + 1);
        reporte.vehiculosEliminados++;
      }
    }
  }

  if (sheetClientes) {
    const data = sheetClientes.getDataRange().getValues();
    const nombreCol = data[0].indexOf('Nombre');

    for (let i = data.length - 1; i >= 1; i--) {
      const nombre = String(data[i][nombreCol] || '').trim();
      if (clientesDemo.indexOf(nombre) !== -1) {
        sheetClientes.deleteRow(i + 1);
        reporte.clientesEliminados++;
      }
    }
  }

  if (sheetServicios) {
    const data = sheetServicios.getDataRange().getValues();
    const idVehiculoCol = data[0].indexOf('ID_Vehiculo');

    for (let i = data.length - 1; i >= 1; i--) {
      const idVehiculo = String(data[i][idVehiculoCol] || '').trim();
      if (vehiculosEliminadosIds.indexOf(idVehiculo) !== -1) {
        sheetServicios.deleteRow(i + 1);
        reporte.serviciosEliminados++;
      }
    }
  }

  reporte.clientesRestantes = Math.max(0, sheetClientes.getLastRow() - 1);
  reporte.vehiculosRestantes = Math.max(0, sheetVehiculos.getLastRow() - 1);
  reporte.serviciosRestantes = Math.max(0, sheetServicios.getLastRow() - 1);

  return {
    ok: true,
    mensaje: 'Limpieza completa. El Sheet queda listo para datos reales.',
    reporte: reporte
  };
}
