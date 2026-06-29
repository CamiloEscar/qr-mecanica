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

/**
 * Trigger de apertura del spreadsheet. Construye el menú "Catálogo" con las
 * acciones de sembrado, migración y rollback del catálogo de marcas/modelos.
 * Se ejecuta sólo dentro del editor de Sheets (los Web Apps no tienen UI),
 * por eso se ignora silenciosamente cualquier excepción.
 */
function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Catálogo')
      .addItem('Cargar catálogo…', 'cargarCatalogoMarcasModelos')
      .addItem('Migrar marcas…', 'migrarMarcasExistentes')
      .addItem('Eliminar catálogo', 'eliminarHojasCatalogo')
      .addToUi();
  } catch (err) {
    // Web App contexts no exponen UI; se ignora el error.
  }
}

const SHEETS = {
  CLIENTES: 'Clientes',
  VEHICULOS: 'Vehiculos',
  SERVICIOS: 'Servicios',
  ACCESOS_LOG: 'AccesosLog',
  TURNOS: 'Turnos',
  PRESUPUESTOS: 'Presupuestos',
  PRESUPUESTOS_ITEMS: 'Presupuestos_Items',
  MARCAS_MODELOS: 'Marcas_Modelos',
  MODELOS_VERSIONES: 'Modelos_Versiones'
};

const COLUMNAS = {
  CLIENTES: ['ID_Cliente', 'Nombre', 'Telefono', 'Email', 'Fecha_Alta', 'Consentimiento'],
  VEHICULOS: ['ID_Vehiculo', 'Token', 'Fecha_Alta', 'Activo', 'Patente', 'Marca', 'Modelo', 'Anio', 'Combustible', 'ID_Cliente', 'URL_QR_Impresa'],
  SERVICIOS: ['ID_Servicio', 'Fecha', 'ID_Vehiculo', 'Kilometraje', 'Descripcion', 'Repuestos', 'Proximo_Mantenimiento', 'Observaciones', 'Fotos_IDs', 'ID_Mecanico'],
  ACCESOS_LOG: ['Timestamp', 'Token_Suffix', 'Email_O_Anonimo'],
  TURNOS: ['ID_Turno', 'Fecha_Hora', 'Duracion_Minutos', 'ID_Vehiculo', 'Tipo_Servicio', 'Descripcion', 'Estado', 'ID_Mecanico', 'Fecha_Creacion', 'Notas'],
  PRESUPUESTOS: ['ID_Presupuesto', 'Fecha', 'Fecha_Vencimiento', 'ID_Vehiculo', 'Estado', 'Subtotal', 'Total', 'Validez_Dias', 'Token_Publico', 'Notas', 'ID_Servicio', 'Fecha_Aprobacion'],
  PRESUPUESTOS_ITEMS: ['ID_Item', 'ID_Presupuesto', 'Descripcion', 'Cantidad', 'Precio_Unitario', 'Subtotal_Item', 'Tipo'],
  MARCAS_MODELOS: ['ID', 'Marca', 'Modelo', 'Origen', 'Fecha_Alta'],
  MODELOS_VERSIONES: ['Marca', 'Modelo', 'Versión']
};

const ESTADOS_PRESUPUESTO = {
  BORRADOR: 'borrador',
  ENVIADO: 'enviado',
  APROBADO: 'aprobado',
  RECHAZADO: 'rechazado',
  VENCIDO: 'vencido',
  COMPLETADO: 'completado'
};

const TOKEN_REGEX = /^[a-zA-Z0-9]{16}$/;
const CHARSET = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';

function doGet(e) {
  const accion = (e.parameter && e.parameter.accion) || '';
  if (accion === 'diag') return serveDiagnostico();
  if (accion === 'admin') return serveAdminPage((e.parameter && e.parameter.session) || '');
  if (accion === 'imprimir' && e.parameter.v) return serveQRPrintPage(e.parameter.v, (e.parameter && e.parameter.session) || '');
  if (accion === 'api') return apiMecano(e);
  if (accion === 'presupuesto' && e.parameter.token) return servePresupuestoPublico(e.parameter.token);
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
  _autoGrowCatalogo_(payload.marca, payload.modelo, payload.version || '');
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
  const hojasACrear = ['Clientes', 'Vehiculos', 'Servicios', 'AccesosLog', 'Turnos', 'Presupuestos', 'Presupuestos_Items', 'Marcas_Modelos', 'Modelos_Versiones'];
  const headersPorHoja = {
    'Clientes': COLUMNAS.CLIENTES,
    'Vehiculos': COLUMNAS.VEHICULOS,
    'Servicios': COLUMNAS.SERVICIOS,
    'AccesosLog': COLUMNAS.ACCESOS_LOG,
    'Turnos': COLUMNAS.TURNOS,
    'Presupuestos': COLUMNAS.PRESUPUESTOS,
    'Presupuestos_Items': COLUMNAS.PRESUPUESTOS_ITEMS,
    'Marcas_Modelos': COLUMNAS.MARCAS_MODELOS,
    'Modelos_Versiones': COLUMNAS.MODELOS_VERSIONES,
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
        case 'altaTurno':
          resultado = altaTurnoInterno(payload);
          break;
        case 'listarTurnos':
          resultado = { ok: true, turnos: listarTurnos(payload || {}) };
          break;
        case 'actualizarTurno':
          resultado = actualizarTurnoInterno(payload);
          break;
        case 'cancelarTurno':
          resultado = cancelarTurnoInterno(payload.id);
          break;
        case 'completarTurno':
          resultado = completarTurnoInterno(payload.id);
          break;
        case 'altaPresupuesto':
          resultado = altaPresupuestoInterno(payload);
          break;
        case 'listarPresupuestos':
          resultado = { ok: true, presupuestos: listarPresupuestos(payload || {}) };
          break;
        case 'obtenerPresupuesto':
          resultado = obtenerPresupuesto(payload.id);
          break;
        case 'actualizarPresupuesto':
          resultado = actualizarPresupuestoInterno(payload);
          break;
        case 'marcarPresupuestoEnviado':
          resultado = marcarPresupuestoEnviado(payload.id);
          break;
        case 'aprobarPresupuesto':
          resultado = aprobarPresupuesto(payload.token);
          break;
        case 'rechazarPresupuesto':
          resultado = rechazarPresupuesto(payload.token, payload.motivo);
          break;
        case 'crearHojaPresupuestos':
          resultado = crearHojaPresupuestos();
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
  _autoGrowCatalogo_(p.marca, p.modelo, p.version || '');
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

function parsearFechaHora(s, fallback) {
  if (!s) return fallback || new Date();
  if (s instanceof Date) return s;
  const normalizado = String(s).replace('T', ' ').replace(/-/g, '-');
  const d = new Date(normalizado);
  if (!isNaN(d.getTime())) return d;
  return fallback || new Date();
}

function altaTurnoInterno(payload) {
  try {
    if (!obtenerSheet(SHEETS.TURNOS)) crearHojaTurnos();
    if (!payload || !payload.idVehiculo) return { ok: false, error: 'Falta el vehículo' };
    if (!payload.fechaHora) return { ok: false, error: 'Falta la fecha y hora' };
    if (!payload.tipoServicio) return { ok: false, error: 'Falta el tipo de servicio' };
    const sheet = obtenerSheet(SHEETS.TURNOS);
    const id = nextId(SHEETS.TURNOS, 'T');
    const fechaHora = parsearFechaHora(payload.fechaHora);
    const duracion = parseInt(payload.duracionMinutos, 10) || 60;
    const estado = payload.estado || 'pendiente';
    const fila = [
      id,
      fechaHora,
      duracion,
      payload.idVehiculo,
      payload.tipoServicio,
      payload.descripcion || '',
      estado,
      payload.idMecanico || '',
      ahoraComoString(),
      payload.notas || ''
    ];
    sheet.appendRow(fila);
    return { ok: true, id: id };
  } catch (err) {
    return { ok: false, error: 'No se pudo crear el turno: ' + (err && err.message ? err.message : err) };
  }
}

function buscarFilaPorId(nombreSheet, nombreColumnaId, idBuscado) {
  const sheet = obtenerSheet(nombreSheet);
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return -1;
  const headers = data[0];
  const idCol = headers.indexOf(nombreColumnaId);
  if (idCol === -1) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(idBuscado)) return i + 1;
  }
  return -1;
}

function listarTurnos(filtros) {
  try {
    if (!obtenerSheet(SHEETS.TURNOS)) crearHojaTurnos();
    const turnos = sheetAObjetos(SHEETS.TURNOS);
    const vehiculos = sheetAObjetos(SHEETS.VEHICULOS);
    const clientes = sheetAObjetos(SHEETS.CLIENTES);
    const vehiculoMap = {};
    vehiculos.forEach(function (v) { vehiculoMap[v.ID_Vehiculo] = v; });
    const clienteMap = {};
    clientes.forEach(function (c) { clienteMap[c.ID_Cliente] = c; });

    const f = filtros || {};
    const desdeMs = f.desde ? new Date(f.desde).getTime() : null;
    const hastaMs = f.hasta ? new Date(f.hasta).getTime() : null;
    const estadoFiltro = f.estado ? String(f.estado) : null;
    const idVehFiltro = f.idVehiculo ? String(f.idVehiculo) : null;

    let filtrados = turnos.filter(function (t) {
      if (estadoFiltro && String(t.Estado) !== estadoFiltro) return false;
      if (idVehFiltro && String(t.ID_Vehiculo) !== idVehFiltro) return false;
      if (desdeMs || hastaMs) {
        const fh = t.Fecha_Hora ? new Date(t.Fecha_Hora).getTime() : null;
        if (!fh || isNaN(fh)) return false;
        if (desdeMs && fh < desdeMs) return false;
        if (hastaMs && fh > hastaMs) return false;
      }
      return true;
    });

    filtrados.forEach(function (t) {
      const veh = vehiculoMap[t.ID_Vehiculo];
      t.Patente = veh ? veh.Patente : '';
      t.Marca = veh ? veh.Marca : '';
      t.Modelo = veh ? veh.Modelo : '';
      if (veh && veh.ID_Cliente) {
        const cli = clienteMap[veh.ID_Cliente];
        t.clienteNombre = cli ? cli.Nombre : '';
        t.Telefono_Cliente = cli ? cli.Telefono : '';
      } else {
        t.clienteNombre = '';
        t.Telefono_Cliente = '';
      }
    });

    filtrados.sort(function (a, b) {
      const fa = a.Fecha_Hora ? new Date(a.Fecha_Hora).getTime() : 0;
      const fb = b.Fecha_Hora ? new Date(b.Fecha_Hora).getTime() : 0;
      return fa - fb;
    });

    return filtrados;
  } catch (err) {
    Logger.log('Error en listarTurnos: ' + (err && err.message ? err.message : err));
    return [];
  }
}

function actualizarTurnoInterno(payload) {
  try {
    if (!obtenerSheet(SHEETS.TURNOS)) crearHojaTurnos();
    if (!payload || !payload.id) return { ok: false, error: 'Falta el ID del turno' };
    const fila = buscarFilaPorId(SHEETS.TURNOS, 'ID_Turno', payload.id);
    if (fila === -1) return { ok: false, error: 'Turno no encontrado: ' + payload.id };

    const sheet = obtenerSheet(SHEETS.TURNOS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const camposActualizables = {
      fechaHora: 'Fecha_Hora',
      duracionMinutos: 'Duracion_Minutos',
      idVehiculo: 'ID_Vehiculo',
      tipoServicio: 'Tipo_Servicio',
      descripcion: 'Descripcion',
      estado: 'Estado',
      idMecanico: 'ID_Mecanico',
      notas: 'Notas'
    };

    Object.keys(camposActualizables).forEach(function (key) {
      if (!(key in payload)) return;
      const colName = camposActualizables[key];
      const colIdx = headers.indexOf(colName);
      if (colIdx === -1) return;
      let valor = payload[key];
      if (colName === 'Fecha_Hora') valor = parsearFechaHora(valor);
      else if (colName === 'Duracion_Minutos') valor = parseInt(valor, 10) || 60;
      else if (valor === null || valor === undefined) valor = '';
      sheet.getRange(fila, colIdx + 1).setValue(valor);
    });

    return { ok: true };
  } catch (err) {
    return { ok: false, error: 'No se pudo actualizar el turno: ' + (err && err.message ? err.message : err) };
  }
}

function cancelarTurnoInterno(idTurno) {
  try {
    if (!obtenerSheet(SHEETS.TURNOS)) crearHojaTurnos();
    if (!idTurno) return { ok: false, error: 'Falta el ID del turno' };
    const fila = buscarFilaPorId(SHEETS.TURNOS, 'ID_Turno', idTurno);
    if (fila === -1) return { ok: false, error: 'Turno no encontrado: ' + idTurno };
    const sheet = obtenerSheet(SHEETS.TURNOS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const estadoCol = headers.indexOf('Estado') + 1;
    sheet.getRange(fila, estadoCol).setValue('cancelado');
    return { ok: true };
  } catch (err) {
    return { ok: false, error: 'No se pudo cancelar el turno: ' + (err && err.message ? err.message : err) };
  }
}

function completarTurnoInterno(idTurno) {
  try {
    if (!obtenerSheet(SHEETS.TURNOS)) crearHojaTurnos();
    if (!idTurno) return { ok: false, error: 'Falta el ID del turno' };
    const fila = buscarFilaPorId(SHEETS.TURNOS, 'ID_Turno', idTurno);
    if (fila === -1) return { ok: false, error: 'Turno no encontrado: ' + idTurno };

    const turnos = sheetAObjetos(SHEETS.TURNOS);
    const turno = turnos.find(function (t) { return String(t.ID_Turno) === String(idTurno); });
    if (!turno) return { ok: false, error: 'Turno no encontrado en datos' };

    const sheet = obtenerSheet(SHEETS.TURNOS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const estadoCol = headers.indexOf('Estado') + 1;
    sheet.getRange(fila, estadoCol).setValue('completado');

    const descripcionFinal = turno.Descripcion && String(turno.Descripcion).trim()
      ? String(turno.Descripcion)
      : String(turno.Tipo_Servicio || 'Servicio');

    const fechaServicio = turno.Fecha_Hora ? parsearFechaHora(turno.Fecha_Hora) : new Date();

    const resultadoServicio = altaServicioInterno({
      idVehiculo: turno.ID_Vehiculo,
      fecha: fechaServicio,
      kilometraje: 0,
      descripcion: descripcionFinal,
      repuestos: '',
      proximoMantenimiento: '',
      observaciones: 'Generado automáticamente desde turno ' + idTurno + ' · ' + String(turno.Tipo_Servicio || ''),
      fotosIds: '',
      idMecanico: turno.ID_Mecanico || ''
    });

    if (!resultadoServicio.ok) {
      return { ok: false, error: 'Turno marcado pero no se pudo crear el servicio: ' + resultadoServicio.error };
    }

    return { ok: true, servicioId: resultadoServicio.id, message: 'Turno completado y servicio creado' };
  } catch (err) {
    return { ok: false, error: 'No se pudo completar el turno: ' + (err && err.message ? err.message : err) };
  }
}

function generarTokenPresupuestoUnico() {
  const presupuestos = sheetAObjetos(SHEETS.PRESUPUESTOS);
  const existentes = new Set();
  for (let i = 0; i < presupuestos.length; i++) {
    if (presupuestos[i].Token_Publico) existentes.add(presupuestos[i].Token_Publico);
  }
  while (true) {
    let token = '';
    for (let i = 0; i < 16; i++) {
      token += CHARSET.charAt(Math.floor(Math.random() * CHARSET.length));
    }
    if (!existentes.has(token)) return token;
  }
}

function altaPresupuestoInterno(payload) {
  try {
    if (!payload || !payload.idVehiculo) return { ok: false, error: 'Falta el vehículo' };
    if (!payload.items || !Array.isArray(payload.items) || payload.items.length === 0) {
      return { ok: false, error: 'Agregá al menos un item al presupuesto' };
    }
    const vehiculo = buscarVehiculoPorId(payload.idVehiculo);
    if (!vehiculo) return { ok: false, error: 'Vehículo no encontrado: ' + payload.idVehiculo };

    const id = nextId(SHEETS.PRESUPUESTOS, 'P');
    const token = generarTokenPresupuestoUnico();
    const validezDias = parseInt(payload.validezDias, 10) || 15;
    const fecha = ahoraComoString();
    const fechaVencimiento = formatearFechaVencimiento(new Date(), validezDias);

    let subtotal = 0;
    const itemsNormalizados = [];
    for (let i = 0; i < payload.items.length; i++) {
      const it = payload.items[i];
      const descripcion = (it.descripcion || '').toString().trim();
      if (!descripcion) return { ok: false, error: 'Item ' + (i + 1) + ': la descripción es obligatoria' };
      const cantidad = parseFloat(it.cantidad) || 0;
      const precioUnitario = parseFloat(it.precioUnitario) || 0;
      if (cantidad <= 0) return { ok: false, error: 'Item ' + (i + 1) + ': la cantidad debe ser mayor a 0' };
      if (precioUnitario < 0) return { ok: false, error: 'Item ' + (i + 1) + ': el precio unitario no puede ser negativo' };
      const subItem = cantidad * precioUnitario;
      subtotal += subItem;
      const tipo = (it.tipo === 'mano_obra') ? 'mano_obra' : 'repuesto';
      itemsNormalizados.push({
        descripcion: descripcion,
        cantidad: cantidad,
        precioUnitario: precioUnitario,
        subtotal: subItem,
        tipo: tipo
      });
    }
    const total = Math.round(subtotal * 100) / 100;
    const subtotalRed = Math.round(subtotal * 100) / 100;

    const sheetP = obtenerSheet(SHEETS.PRESUPUESTOS);
    sheetP.appendRow([
      id,
      fecha,
      fechaVencimiento,
      payload.idVehiculo,
      ESTADOS_PRESUPUESTO.BORRADOR,
      subtotalRed,
      total,
      validezDias,
      token,
      payload.notas || '',
      '',
      ''
    ]);

    const sheetItems = obtenerSheet(SHEETS.PRESUPUESTOS_ITEMS);
    const itemsParaSheet = [];
    for (let j = 0; j < itemsNormalizados.length; j++) {
      const it = itemsNormalizados[j];
      const idItem = nextId(SHEETS.PRESUPUESTOS_ITEMS, 'PI');
      itemsParaSheet.push([
        idItem,
        id,
        it.descripcion,
        it.cantidad,
        it.precioUnitario,
        it.subtotal,
        it.tipo
      ]);
    }
    if (itemsParaSheet.length > 0) {
      const startRow = sheetItems.getLastRow() + 1;
      sheetItems.getRange(startRow, 1, itemsParaSheet.length, COLUMNAS.PRESUPUESTOS_ITEMS.length).setValues(itemsParaSheet);
    }

    return {
      ok: true,
      id: id,
      token: token,
      total: total,
      subtotal: subtotalRed,
      items: itemsNormalizados.length
    };
  } catch (err) {
    return { ok: false, error: 'No se pudo crear el presupuesto: ' + (err && err.message ? err.message : err) };
  }
}

function listarPresupuestos(filtros) {
  try {
    const presupuestos = sheetAObjetos(SHEETS.PRESUPUESTOS);
    const vehiculos = sheetAObjetos(SHEETS.VEHICULOS);
    const clientes = sheetAObjetos(SHEETS.CLIENTES);
    const items = sheetAObjetos(SHEETS.PRESUPUESTOS_ITEMS);

    const vehiculoMap = {};
    vehiculos.forEach(function (v) { vehiculoMap[v.ID_Vehiculo] = v; });
    const clienteMap = {};
    clientes.forEach(function (c) { clienteMap[c.ID_Cliente] = c; });

    const itemsPorPresupuesto = {};
    let totalItems = 0;
    items.forEach(function (it) {
      if (!itemsPorPresupuesto[it.ID_Presupuesto]) itemsPorPresupuesto[it.ID_Presupuesto] = 0;
      itemsPorPresupuesto[it.ID_Presupuesto]++;
      totalItems++;
    });

    const f = filtros || {};
    const estadoFiltro = f.estado ? String(f.estado) : null;
    const idVehFiltro = f.idVehiculo ? String(f.idVehiculo) : null;
    const desdeMs = f.desde ? new Date(f.desde).getTime() : null;
    const hastaMs = f.hasta ? new Date(f.hasta).getTime() : null;
    const q = (f.q || '').toString().toLowerCase().trim();

    const ahora = new Date();
    const filtrados = presupuestos.filter(function (p) {
      let estadoCalc = String(p.Estado || '');
      if (estadoCalc === ESTADOS_PRESUPUESTO.ENVIADO && p.Fecha_Vencimiento) {
        const fv = p.Fecha_Vencimiento instanceof Date ? p.Fecha_Vencimiento : new Date(p.Fecha_Vencimiento);
        if (!isNaN(fv.getTime()) && fv.getTime() < ahora.getTime()) {
          estadoCalc = ESTADOS_PRESUPUESTO.VENCIDO;
        }
      }
      p._EstadoCalculado = estadoCalc;

      if (estadoFiltro && estadoCalc !== estadoFiltro) return false;
      if (idVehFiltro && String(p.ID_Vehiculo) !== idVehFiltro) return false;
      if (desdeMs || hastaMs) {
        const fp = p.Fecha ? new Date(p.Fecha).getTime() : null;
        if (!fp || isNaN(fp)) return false;
        if (desdeMs && fp < desdeMs) return false;
        if (hastaMs && fp > hastaMs) return false;
      }
      if (q) {
        const veh = vehiculoMap[p.ID_Vehiculo];
        const cli = veh ? clienteMap[veh.ID_Cliente] : null;
        const haystack = [
          p.ID_Presupuesto,
          veh ? veh.Patente : '',
          veh ? veh.Marca : '',
          veh ? veh.Modelo : '',
          cli ? cli.Nombre : '',
          p.Notas || ''
        ].join(' ').toLowerCase();
        if (haystack.indexOf(q) === -1) return false;
      }
      return true;
    });

    filtrados.forEach(function (p) {
      const veh = vehiculoMap[p.ID_Vehiculo];
      p.Patente = veh ? veh.Patente : '';
      p.Marca = veh ? veh.Marca : '';
      p.Modelo = veh ? veh.Modelo : '';
      if (veh) {
        const cli = clienteMap[veh.ID_Cliente];
        p.clienteNombre = cli ? cli.Nombre : '';
        p.clienteTelefono = cli ? cli.Telefono : '';
      }
      p.itemsCount = itemsPorPresupuesto[p.ID_Presupuesto] || 0;
      p.urlPublica = p.Token_Publico ? ScriptApp.getService().getUrl() + '?accion=presupuesto&token=' + p.Token_Publico : '';
    });

    filtrados.sort(function (a, b) {
      const fa = a.Fecha ? new Date(a.Fecha).getTime() : 0;
      const fb = b.Fecha ? new Date(b.Fecha).getTime() : 0;
      return fb - fa;
    });

    return filtrados;
  } catch (err) {
    Logger.log('Error en listarPresupuestos: ' + (err && err.message ? err.message : err));
    return [];
  }
}

function obtenerPresupuesto(id) {
  try {
    if (!id) return { ok: false, error: 'Falta el ID del presupuesto' };
    const presupuestos = sheetAObjetos(SHEETS.PRESUPUESTOS);
    const presupuesto = presupuestos.find(function (p) { return p.ID_Presupuesto === id; });
    if (!presupuesto) return { ok: false, error: 'Presupuesto no encontrado: ' + id };

    const vehiculo = presupuesto.ID_Vehiculo ? buscarVehiculoPorId(presupuesto.ID_Vehiculo) : null;
    let cliente = null;
    if (vehiculo && vehiculo.ID_Cliente) {
      const clientes = sheetAObjetos(SHEETS.CLIENTES);
      cliente = clientes.find(function (c) { return c.ID_Cliente === vehiculo.ID_Cliente; }) || null;
    }
    const items = listarItemsPorPresupuesto(id);

    if (presupuesto.Token_Publico) {
      presupuesto.urlPublica = ScriptApp.getService().getUrl() + '?accion=presupuesto&token=' + presupuesto.Token_Publico;
    }

    return {
      ok: true,
      presupuesto: presupuesto,
      items: items,
      vehiculo: vehiculo || null,
      cliente: cliente
    };
  } catch (err) {
    return { ok: false, error: 'No se pudo obtener el presupuesto: ' + (err && err.message ? err.message : err) };
  }
}

function actualizarPresupuestoInterno(payload) {
  try {
    if (!payload || !payload.id) return { ok: false, error: 'Falta el ID del presupuesto' };
    const fila = buscarFilaPorId(SHEETS.PRESUPUESTOS, 'ID_Presupuesto', payload.id);
    if (fila === -1) return { ok: false, error: 'Presupuesto no encontrado: ' + payload.id };

    const sheet = obtenerSheet(SHEETS.PRESUPUESTOS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const camposActualizables = {
      estado: 'Estado',
      notas: 'Notas',
      idServicio: 'ID_Servicio',
      fechaAprobacion: 'Fecha_Aprobacion'
    };

    Object.keys(camposActualizables).forEach(function (key) {
      if (!(key in payload)) return;
      const colName = camposActualizables[key];
      const colIdx = headers.indexOf(colName);
      if (colIdx === -1) return;
      let valor = payload[key];
      if (valor === null || valor === undefined) valor = '';
      sheet.getRange(fila, colIdx + 1).setValue(valor);
    });

    return { ok: true };
  } catch (err) {
    return { ok: false, error: 'No se pudo actualizar el presupuesto: ' + (err && err.message ? err.message : err) };
  }
}

function marcarPresupuestoEnviado(id) {
  try {
    if (!id) return { ok: false, error: 'Falta el ID del presupuesto' };
    const presupuestos = sheetAObjetos(SHEETS.PRESUPUESTOS);
    const presupuesto = presupuestos.find(function (p) { return p.ID_Presupuesto === id; });
    if (!presupuesto) return { ok: false, error: 'Presupuesto no encontrado: ' + id };

    if (presupuesto.Estado === ESTADOS_PRESUPUESTO.APROBADO ||
        presupuesto.Estado === ESTADOS_PRESUPUESTO.RECHAZADO ||
        presupuesto.Estado === ESTADOS_PRESUPUESTO.COMPLETADO) {
      return { ok: false, error: 'No se puede enviar un presupuesto en estado ' + presupuesto.Estado };
    }

    const r = actualizarPresupuestoInterno({ id: id, estado: ESTADOS_PRESUPUESTO.ENVIADO });
    if (!r.ok) return r;

    return {
      ok: true,
      token: presupuesto.Token_Publico,
      urlPublica: ScriptApp.getService().getUrl() + '?accion=presupuesto&token=' + presupuesto.Token_Publico
    };
  } catch (err) {
    return { ok: false, error: 'No se pudo marcar como enviado: ' + (err && err.message ? err.message : err) };
  }
}

function aprobarPresupuesto(token) {
  try {
    if (!token || !/^[a-zA-Z0-9]{16}$/.test(token)) {
      return { ok: false, error: 'Token inválido' };
    }
    const presupuestos = sheetAObjetos(SHEETS.PRESUPUESTOS);
    const presupuesto = presupuestos.find(function (p) { return p.Token_Publico === token; });
    if (!presupuesto) return { ok: false, error: 'Presupuesto no encontrado' };

    if (presupuesto.Estado !== ESTADOS_PRESUPUESTO.ENVIADO &&
        presupuesto.Estado !== ESTADOS_PRESUPUESTO.BORRADOR) {
      return { ok: false, error: 'Este presupuesto ya fue ' + presupuesto.Estado + ' y no se puede aprobar de nuevo' };
    }

    const items = listarItemsPorPresupuesto(presupuesto.ID_Presupuesto);
    const repuestosTexto = items.map(function (it) {
      return (it.Cantidad || 1) + 'x ' + (it.Descripcion || '');
    }).join(', ');

    const descripcionItems = items.map(function (it) {
      return (it.Tipo === 'mano_obra' ? 'MO: ' : '') + (it.Descripcion || '');
    }).join(' · ');

    const fechaAprob = new Date();
    const resultadoServicio = altaServicioInterno({
      idVehiculo: presupuesto.ID_Vehiculo,
      fecha: fechaAprob,
      kilometraje: 0,
      descripcion: 'Presupuesto aprobado #' + presupuesto.ID_Presupuesto + ' · ' + descripcionItems,
      repuestos: repuestosTexto,
      proximoMantenimiento: '',
      observaciones: 'Generado automáticamente desde presupuesto ' + presupuesto.ID_Presupuesto,
      fotosIds: '',
      idMecanico: ''
    });

    if (!resultadoServicio.ok) {
      return { ok: false, error: 'No se pudo crear el servicio: ' + resultadoServicio.error };
    }

    const r = actualizarPresupuestoInterno({
      id: presupuesto.ID_Presupuesto,
      estado: ESTADOS_PRESUPUESTO.APROBADO,
      idServicio: resultadoServicio.id,
      fechaAprobacion: Utilities.formatDate(fechaAprob, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    });
    if (!r.ok) return r;

    return {
      ok: true,
      message: 'Presupuesto aprobado y servicio creado',
      servicioId: resultadoServicio.id,
      idPresupuesto: presupuesto.ID_Presupuesto
    };
  } catch (err) {
    return { ok: false, error: 'No se pudo aprobar el presupuesto: ' + (err && err.message ? err.message : err) };
  }
}

function rechazarPresupuesto(token, motivo) {
  try {
    if (!token || !/^[a-zA-Z0-9]{16}$/.test(token)) {
      return { ok: false, error: 'Token inválido' };
    }
    const presupuestos = sheetAObjetos(SHEETS.PRESUPUESTOS);
    const presupuesto = presupuestos.find(function (p) { return p.Token_Publico === token; });
    if (!presupuesto) return { ok: false, error: 'Presupuesto no encontrado' };

    if (presupuesto.Estado === ESTADOS_PRESUPUESTO.APROBADO ||
        presupuesto.Estado === ESTADOS_PRESUPUESTO.COMPLETADO ||
        presupuesto.Estado === ESTADOS_PRESUPUESTO.RECHAZADO) {
      return { ok: false, error: 'Este presupuesto ya está en estado ' + presupuesto.Estado };
    }

    const motivoLimpio = (motivo || '').toString().trim();
    const notasActuales = (presupuesto.Notas || '').toString();
    const sep = notasActuales ? ' | ' : '';
    const motivoTexto = motivoLimpio ? ('[Rechazado: ' + motivoLimpio + ']') : '[Rechazado]';
    const notasNuevas = notasActuales + sep + motivoTexto;

    const r = actualizarPresupuestoInterno({
      id: presupuesto.ID_Presupuesto,
      estado: ESTADOS_PRESUPUESTO.RECHAZADO,
      notas: notasNuevas
    });
    if (!r.ok) return r;

    return { ok: true, message: 'Presupuesto rechazado' };
  } catch (err) {
    return { ok: false, error: 'No se pudo rechazar el presupuesto: ' + (err && err.message ? err.message : err) };
  }
}

function buscarPresupuestoPorToken(token) {
  const presupuestos = sheetAObjetos(SHEETS.PRESUPUESTOS);
  return presupuestos.find(function (p) { return p.Token_Publico === token; }) || null;
}

function listarItemsPorPresupuesto(idPresupuesto) {
  const items = sheetAObjetos(SHEETS.PRESUPUESTOS_ITEMS);
  return items.filter(function (i) { return i.ID_Presupuesto === idPresupuesto; });
}

function formatearFechaVencimiento(fechaBase, dias) {
  const d = new Date(fechaBase.getTime());
  d.setDate(d.getDate() + (parseInt(dias, 10) || 15));
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function servePresupuestoPublico(token) {
  if (!token || !/^[a-zA-Z0-9]{16}$/.test(token)) {
    return serveError('Enlace inválido.');
  }
  const presupuesto = buscarPresupuestoPorToken(token);
  if (!presupuesto) return serveError('Presupuesto no encontrado.');

  const estadoCalc = calcularEstadoPresupuesto(presupuesto);
  presupuesto.Estado = estadoCalc;

  const items = listarItemsPorPresupuesto(presupuesto.ID_Presupuesto);
  items.sort(function (a, b) {
    return String(a.ID_Item).localeCompare(String(b.ID_Item));
  });

  const vehiculo = presupuesto.ID_Vehiculo ? buscarVehiculoPorId(presupuesto.ID_Vehiculo) : null;
  let cliente = null;
  if (vehiculo && vehiculo.ID_Cliente) {
    const clientes = sheetAObjetos(SHEETS.CLIENTES);
    cliente = clientes.find(function (c) { return c.ID_Cliente === vehiculo.ID_Cliente; }) || null;
  }

  const t = HtmlService.createTemplateFromFile('PresupuestoPage');
  t.presupuesto = presupuesto;
  t.items = items;
  t.vehiculo = vehiculo || {};
  t.cliente = cliente || {};
  t.webappUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle('Presupuesto #' + presupuesto.ID_Presupuesto + ' - Mecanica Martinez')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function calcularEstadoPresupuesto(p) {
  let estado = String(p.Estado || '');
  if (estado === ESTADOS_PRESUPUESTO.ENVIADO && p.Fecha_Vencimiento) {
    const fv = p.Fecha_Vencimiento instanceof Date ? p.Fecha_Vencimiento : new Date(p.Fecha_Vencimiento);
    if (!isNaN(fv.getTime()) && fv.getTime() < new Date().getTime()) {
      return ESTADOS_PRESUPUESTO.VENCIDO;
    }
  }
  return estado;
}

function crearHojaPresupuestos() {
  try {
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    const resultado = { presupuesto: 'existente', items: 'existente' };

    let sheet = ss.getSheetByName(SHEETS.PRESUPUESTOS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.PRESUPUESTOS);
      sheet.getRange(1, 1, 1, COLUMNAS.PRESUPUESTOS.length).setValues([COLUMNAS.PRESUPUESTOS]);
      sheet.getRange(1, 1, 1, COLUMNAS.PRESUPUESTOS.length).setFontWeight('bold').setBackground('#e2e8f0');
      sheet.setFrozenRows(1);
      sheet.getRange('B2:B').setNumberFormat('yyyy-MM-dd HH:mm:ss');
      sheet.getRange('C2:C').setNumberFormat('yyyy-MM-dd');
      sheet.getRange('M2:M').setNumberFormat('yyyy-MM-dd HH:mm:ss');
      sheet.getRange('F2:H').setNumberFormat('#,##0.00');
      resultado.presupuesto = 'creado';
    }

    sheet = ss.getSheetByName(SHEETS.PRESUPUESTOS_ITEMS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.PRESUPUESTOS_ITEMS);
      sheet.getRange(1, 1, 1, COLUMNAS.PRESUPUESTOS_ITEMS.length).setValues([COLUMNAS.PRESUPUESTOS_ITEMS]);
      sheet.getRange(1, 1, 1, COLUMNAS.PRESUPUESTOS_ITEMS.length).setFontWeight('bold').setBackground('#e2e8f0');
      sheet.setFrozenRows(1);
      sheet.getRange('E2:F').setNumberFormat('#,##0.00');
      resultado.items = 'creado';
    }

    return { ok: true, mensaje: 'Hojas de presupuestos listas', resultado: resultado };
  } catch (err) {
    return { ok: false, error: 'No se pudieron crear las hojas de presupuestos: ' + (err && err.message ? err.message : err) };
  }
}

function cargarPresupuestosDemo() {
  try {
    const sheetRes = crearHojaPresupuestos();
    if (!sheetRes.ok) return sheetRes;

    const vehiculos = sheetAObjetos(SHEETS.VEHICULOS);
    const vehiculosActivos = vehiculos.filter(function (v) { return esActivo(v.Activo); });
    if (vehiculosActivos.length === 0) {
      return { ok: false, error: 'No hay vehículos cargados. Corré primero cargarDatosDemoMasivos().' };
    }

    const existentes = sheetAObjetos(SHEETS.PRESUPUESTOS);
    if (existentes.length > 0) {
      return { ok: false, mensaje: 'Ya hay presupuestos cargados. Limpiá primero con limpiarPresupuestosDemo().', yaCargado: true };
    }

    const clientes = sheetAObjetos(SHEETS.CLIENTES);
    const clienteMap = {};
    clientes.forEach(function (c) { clienteMap[c.ID_Cliente] = c; });

    const plan = [
      {
        vehiculo: vehiculosActivos[0],
        estado: ESTADOS_PRESUPUESTO.BORRADOR,
        validez: 15,
        notas: 'Service 60.000 km - cliente frecuente',
        items: [
          { tipo: 'repuesto', desc: 'Aceite Shell Helix Ultra 5W40 5L', cant: 1, precio: 45000 },
          { tipo: 'repuesto', desc: 'Filtro de aceite Mann W712/75', cant: 1, precio: 5500 },
          { tipo: 'repuesto', desc: 'Filtro de aire Mann C25114', cant: 1, precio: 7800 },
          { tipo: 'mano_obra', desc: 'Mano de obra service completo', cant: 1, precio: 35000 }
        ]
      },
      {
        vehiculo: vehiculosActivos[1 % vehiculosActivos.length],
        estado: ESTADOS_PRESUPUESTO.BORRADOR,
        validez: 15,
        notas: 'Diagnostico + presupuesto reparacion tren delantero',
        items: [
          { tipo: 'repuesto', desc: 'Kit tren delantero (precaps, rotulas, bujes)', cant: 1, precio: 58000 },
          { tipo: 'repuesto', desc: 'Extremos de direccion x2', cant: 2, precio: 8500 },
          { tipo: 'mano_obra', desc: 'Mano de obra reparacion tren delantero', cant: 4, precio: 12000 }
        ]
      },
      {
        vehiculo: vehiculosActivos[2 % vehiculosActivos.length],
        estado: ESTADOS_PRESUPUESTO.BORRADOR,
        validez: 10,
        notas: 'Aprobado en mostrador, falta enviar',
        items: [
          { tipo: 'repuesto', desc: 'Pastillas freno Brembo P85020', cant: 1, precio: 22000 },
          { tipo: 'repuesto', desc: 'Disco de freno Fremax delantero x2', cant: 2, precio: 32000 },
          { tipo: 'mano_obra', desc: 'Cambio pastillas y discos delanteros', cant: 1, precio: 25000 }
        ]
      },
      {
        vehiculo: vehiculosActivos[3 % vehiculosActivos.length],
        estado: ESTADOS_PRESUPUESTO.ENVIADO,
        validez: 15,
        notas: 'Enviado por WhatsApp, esperando respuesta',
        items: [
          { tipo: 'repuesto', desc: 'Bujias NGK Iridium ILZKR7B x4', cant: 4, precio: 4500 },
          { tipo: 'repuesto', desc: 'Cables de bujia', cant: 1, precio: 12000 },
          { tipo: 'mano_obra', desc: 'Cambio bujias y cables', cant: 1, precio: 18000 }
        ]
      },
      {
        vehiculo: vehiculosActivos[4 % vehiculosActivos.length],
        estado: ESTADOS_PRESUPUESTO.ENVIADO,
        validez: 20,
        notas: 'Enviado por mail, vence pronto',
        items: [
          { tipo: 'repuesto', desc: 'Kit distribucion INA K015457XS', cant: 1, precio: 78000 },
          { tipo: 'repuesto', desc: 'Bomba de agua', cant: 1, precio: 35000 },
          { tipo: 'repuesto', desc: 'Correa poly-v', cant: 1, precio: 18000 },
          { tipo: 'mano_obra', desc: 'Cambio kit distribucion completo', cant: 5, precio: 14000 }
        ]
      },
      {
        vehiculo: vehiculosActivos[5 % vehiculosActivos.length],
        estado: ESTADOS_PRESUPUESTO.APROBADO,
        validez: 15,
        notas: 'Cliente aprobo por WhatsApp, listo para agendar',
        items: [
          { tipo: 'repuesto', desc: 'Aceite Mobil 1 ESP 5W30 4L', cant: 1, precio: 38000 },
          { tipo: 'repuesto', desc: 'Filtro de aceite Wix 51348', cant: 1, precio: 4800 },
          { tipo: 'mano_obra', desc: 'Cambio aceite y filtro', cant: 1, precio: 15000 }
        ]
      },
      {
        vehiculo: vehiculosActivos[6 % vehiculosActivos.length],
        estado: ESTADOS_PRESUPUESTO.APROBADO,
        validez: 15,
        notas: 'Aprobado - servicio agendado para la semana proxima',
        items: [
          { tipo: 'repuesto', desc: 'Amortiguadores Monroe G7308 x2', cant: 2, precio: 45000 },
          { tipo: 'repuesto', desc: 'Cazoletas y bujes', cant: 1, precio: 18000 },
          { tipo: 'mano_obra', desc: 'Cambio amortiguadores delanteros', cant: 3, precio: 15000 }
        ]
      },
      {
        vehiculo: vehiculosActivos[7 % vehiculosActivos.length],
        estado: ESTADOS_PRESUPUESTO.VENCIDO,
        validez: 7,
        notas: 'Vencio sin respuesta del cliente',
        items: [
          { tipo: 'repuesto', desc: 'Bateria Moura M70TD', cant: 1, precio: 95000 },
          { tipo: 'mano_obra', desc: 'Cambio bateria y chequeo sistema carga', cant: 1, precio: 8000 }
        ]
      }
    ];

    const sheetP = obtenerSheet(SHEETS.PRESUPUESTOS);
    const sheetI = obtenerSheet(SHEETS.PRESUPUESTOS_ITEMS);
    const idsCreados = [];

    const ahoraMs = Date.now();

    plan.forEach(function (p, idx) {
      const id = nextId(SHEETS.PRESUPUESTOS, 'P');
      const token = generarTokenPresupuestoUnico();
      const fecha = new Date(ahoraMs - (plan.length - idx) * 24 * 60 * 60 * 1000);
      const fechaVenc = new Date(fecha.getTime() + p.validez * 24 * 60 * 60 * 1000);

      let subtotal = 0;
      p.items.forEach(function (it) { subtotal += it.cant * it.precio; });
      const total = Math.round(subtotal * 100) / 100;
      const subRed = Math.round(subtotal * 100) / 100;

      const fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      const fechaVencStr = Utilities.formatDate(fechaVenc, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      let idServicio = '';
      let fechaAprobacion = '';
      if (p.estado === ESTADOS_PRESUPUESTO.APROBADO || p.estado === ESTADOS_PRESUPUESTO.COMPLETADO) {
        const repuestosTxt = p.items.map(function (it) { return it.cant + 'x ' + it.desc; }).join(', ');
        const descTxt = p.items.map(function (it) { return (it.tipo === 'mano_obra' ? 'MO: ' : '') + it.desc; }).join(' · ');
        const servicio = altaServicioInterno({
          idVehiculo: p.vehiculo.ID_Vehiculo,
          fecha: fecha,
          kilometraje: 0,
          descripcion: 'Presupuesto aprobado #' + id + ' · ' + descTxt,
          repuestos: repuestosTxt,
          proximoMantenimiento: '',
          observaciones: 'Generado automáticamente desde presupuesto demo ' + id,
          fotosIds: '',
          idMecanico: ''
        });
        if (servicio.ok) {
          idServicio = servicio.id;
          fechaAprobacion = fechaStr;
        }
      }

      sheetP.appendRow([
        id,
        fechaStr,
        fechaVencStr,
        p.vehiculo.ID_Vehiculo,
        p.estado,
        subRed,
        total,
        p.validez,
        token,
        p.notas,
        idServicio,
        fechaAprobacion
      ]);

      const itemsParaSheet = [];
      p.items.forEach(function (it) {
        const idItem = nextId(SHEETS.PRESUPUESTOS_ITEMS, 'PI');
        const subItem = it.cant * it.precio;
        itemsParaSheet.push([idItem, id, it.desc, it.cant, it.precio, subItem, it.tipo]);
      });
      if (itemsParaSheet.length > 0) {
        const startRow = sheetI.getLastRow() + 1;
        sheetI.getRange(startRow, 1, itemsParaSheet.length, COLUMNAS.PRESUPUESTOS_ITEMS.length).setValues(itemsParaSheet);
      }

      idsCreados.push(id);
    });

    return {
      ok: true,
      mensaje: 'Presupuestos demo cargados correctamente',
      presupuestosCreados: idsCreados.length,
      ids: idsCreados,
      distribucion: {
        borrador: plan.filter(function (p) { return p.estado === ESTADOS_PRESUPUESTO.BORRADOR; }).length,
        enviado: plan.filter(function (p) { return p.estado === ESTADOS_PRESUPUESTO.ENVIADO; }).length,
        aprobado: plan.filter(function (p) { return p.estado === ESTADOS_PRESUPUESTO.APROBADO; }).length,
        vencido: plan.filter(function (p) { return p.estado === ESTADOS_PRESUPUESTO.VENCIDO; }).length
      }
    };
  } catch (err) {
    return { ok: false, error: 'No se pudieron cargar los presupuestos demo: ' + (err && err.message ? err.message : err) };
  }
}

function limpiarPresupuestosDemo() {
  try {
    const sheetP = obtenerSheet(SHEETS.PRESUPUESTOS);
    const sheetI = obtenerSheet(SHEETS.PRESUPUESTOS_ITEMS);

    let presupuestosEliminados = 0;
    if (sheetP) {
      const data = sheetP.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        sheetP.deleteRow(i + 1);
        presupuestosEliminados++;
      }
    }

    let itemsEliminados = 0;
    if (sheetI) {
      const data = sheetI.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        sheetI.deleteRow(i + 1);
        itemsEliminados++;
      }
    }

    return {
      ok: true,
      mensaje: 'Presupuestos demo eliminados',
      presupuestosEliminados: presupuestosEliminados,
      itemsEliminados: itemsEliminados
    };
  } catch (err) {
    return { ok: false, error: 'No se pudieron limpiar los presupuestos: ' + (err && err.message ? err.message : err) };
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
        case 'altaTurno': resultado = altaTurnoInterno(payload); break;
        case 'listarTurnos': resultado = { ok: true, turnos: listarTurnos(payload || {}) }; break;
        case 'actualizarTurno': resultado = actualizarTurnoInterno(payload); break;
        case 'cancelarTurno': resultado = cancelarTurnoInterno(payload.id); break;
        case 'completarTurno': resultado = completarTurnoInterno(payload.id); break;
        case 'altaPresupuesto': resultado = altaPresupuestoInterno(payload); break;
        case 'listarPresupuestos': resultado = { ok: true, presupuestos: listarPresupuestos(payload || {}) }; break;
        case 'obtenerPresupuesto': resultado = obtenerPresupuesto(payload.id); break;
        case 'actualizarPresupuesto': resultado = actualizarPresupuestoInterno(payload); break;
        case 'marcarPresupuestoEnviado': resultado = marcarPresupuestoEnviado(payload.id); break;
        case 'aprobarPresupuesto': resultado = aprobarPresupuesto(payload.token); break;
        case 'rechazarPresupuesto': resultado = rechazarPresupuesto(payload.token, payload.motivo); break;
        case 'crearHojaPresupuestos': resultado = crearHojaPresupuestos(); break;
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

function migrarQuitarIvaPresupuestos() {
  try {
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    const sheet = ss.getSheetByName('Presupuestos');
    if (!sheet) return { ok: false, error: 'No existe la hoja Presupuestos' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const ivaColIndex = headers.indexOf('IVA');
    const subtotalColIndex = headers.indexOf('Subtotal');
    const totalColIndex = headers.indexOf('Total');

    if (ivaColIndex === -1) {
      return { ok: true, mensaje: 'No hay columna IVA, nada que migrar', accion: 'ninguna' };
    }

    const filasActualizadas = [];
    if (subtotalColIndex !== -1 && totalColIndex !== -1) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const subtotal = data[i][subtotalColIndex];
        if (typeof subtotal === 'number' && subtotal > 0) {
          sheet.getRange(i + 1, totalColIndex + 1).setValue(subtotal);
          filasActualizadas.push({ fila: i + 1, idPresupuesto: String(data[i][0] || ''), totalAnterior: Number(data[i][totalColIndex]) || 0, totalNuevo: subtotal });
        }
      }
    }

    sheet.deleteColumn(ivaColIndex + 1);

    return {
      ok: true,
      mensaje: 'Columna IVA eliminada. Total ajustado a Subtotal en ' + filasActualizadas.length + ' fila(s).',
      accion: 'migracion completa',
      filasActualizadas: filasActualizadas
    };
  } catch (err) {
    return { ok: false, error: 'No se pudo completar la migracion: ' + (err && err.message ? err.message : err) };
  }
}

function crearHojaTurnos() {
  try {
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    let hoja = ss.getSheetByName(SHEETS.TURNOS);
    if (hoja) {
      const headersActuales = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
      const headersEsperados = COLUMNAS.TURNOS;
      const headersCoinciden = headersEsperados.every(function (h, i) { return headersActuales[i] === h; }) &&
        headersActuales.length === headersEsperados.length;
      if (headersCoinciden) {
        return { ok: true, mensaje: 'La hoja Turnos ya existe con headers correctos', accion: 'ninguna' };
      }
      hoja.getRange(1, 1, 1, headersEsperados.length)
        .setValues([headersEsperados])
        .setFontWeight('bold')
        .setBackground('#e2e8f0');
      return { ok: true, mensaje: 'La hoja Turnos ya existía; headers normalizados', accion: 'headers actualizados' };
    }
    hoja = ss.insertSheet(SHEETS.TURNOS);
    hoja.getRange(1, 1, 1, COLUMNAS.TURNOS.length)
      .setValues([COLUMNAS.TURNOS]);
    hoja.getRange(1, 1, 1, COLUMNAS.TURNOS.length)
      .setFontWeight('bold')
      .setBackground('#e2e8f0');
    hoja.setFrozenRows(1);
    hoja.getRange('B2:B').setNumberFormat('yyyy-MM-dd HH:mm');
    hoja.getRange('I2:I').setNumberFormat('yyyy-MM-dd HH:mm');
    return { ok: true, mensaje: 'Hoja Turnos creada correctamente', accion: 'creada' };
  } catch (err) {
    return { ok: false, error: 'No se pudo crear la hoja Turnos: ' + (err && err.message ? err.message : err) };
  }
}

/**
 * Catálogo curado del mercado argentino: autos, clásicos, motos, camiones y
 * maquinaria. Cada entrada tiene al menos 2-3 modelos top. El seed se inserta
 * con Origen='seed' y las filas son inmutables (Q1 del design): un re-run no
 * duplica porque la clave única es (lower(Marca), lower(Modelo)).
 */
const MARCAS_MODELOS_SEED = [
  // AUTOS
  { marca: 'Toyota', modelos: ['Hilux', 'Etios', 'Corolla', 'Yaris', 'RAV4', 'Land Cruiser'] },
  { marca: 'Ford', modelos: ['Ranger', 'Focus', 'Fiesta', 'Ka', 'Ecosport', 'Mondeo', 'Territory', 'Bronco'] },
  { marca: 'Volkswagen', modelos: ['Gol', 'Voyage', 'Fox', 'Polo', 'Vento', 'Amarok', 'Tiguan', 'Nivus', 'T-Cross'] },
  { marca: 'Chevrolet', modelos: ['Onix', 'Prisma', 'Cruze', 'Tracker', 'S10', 'Spin', 'Equinox', 'Montana'] },
  { marca: 'Renault', modelos: ['Kwid', 'Sandero', 'Logan', 'Duster', 'Captur', 'Stepway', 'Kangoo'] },
  { marca: 'Fiat', modelos: ['Cronos', 'Argo', 'Mobi', 'Pulse', 'Fastback', 'Toro', '500', 'Uno'] },
  { marca: 'Peugeot', modelos: ['208', '2008', '308', '3008', '408', 'Partner', '5008'] },
  { marca: 'Citroën', modelos: ['C3', 'C4 Cactus', 'Berlingo', 'C5 Aircross'] },
  { marca: 'Honda', modelos: ['Civic', 'Accord', 'CR-V', 'HR-V', 'Fit', 'City', 'WR-V'] },
  { marca: 'Nissan', modelos: ['Frontier', 'Kicks', 'Versa', 'Sentra', 'X-Trail', 'Murano'] },
  { marca: 'Jeep', modelos: ['Renegade', 'Compass', 'Wrangler', 'Cherokee', 'Gladiator'] },
  { marca: 'Hyundai', modelos: ['HB20', 'Creta', 'Tucson', 'Santa Fe', 'Kona', 'i10', 'i20'] },
  { marca: 'Kia', modelos: ['Picanto', 'Rio', 'Cerato', 'Sportage', 'Sorento', 'Seltos'] },
  { marca: 'Mitsubishi', modelos: ['L200', 'Outlander', 'ASX', 'Eclipse Cross', 'Pajero'] },
  { marca: 'Subaru', modelos: ['Forester', 'Outback', 'XV', 'Impreza', 'Legacy'] },
  { marca: 'Suzuki', modelos: ['Swift', 'Vitara', 'Jimny', 'S-Cross', 'Baleno'] },
  { marca: 'Mazda', modelos: ['CX-3', 'CX-5', 'CX-30', 'Mazda2', 'Mazda3', 'BT-50'] },
  { marca: 'Daihatsu', modelos: ['Terios', 'Sirion', 'Move'] },
  { marca: 'Lexus', modelos: ['IS', 'ES', 'RX', 'NX', 'UX'] },
  { marca: 'Mercedes-Benz', modelos: ['Clase A', 'Clase C', 'Clase E', 'GLA', 'GLC', 'Sprinter', 'Vito'] },
  { marca: 'BMW', modelos: ['Serie 1', 'Serie 3', 'Serie 5', 'X1', 'X3', 'X5', 'X6'] },
  { marca: 'Audi', modelos: ['A1', 'A3', 'A4', 'Q3', 'Q5', 'Q7'] },
  { marca: 'Porsche', modelos: ['Cayenne', 'Macan', '911', 'Panamera', 'Taycan', 'Boxster'] },
  { marca: 'Lamborghini', modelos: ['Huracán', 'Urus', 'Aventador'] },
  { marca: 'Ferrari', modelos: ['488', 'Roma', 'F8', 'Portofino', 'SF90'] },
  { marca: 'Maserati', modelos: ['Ghibli', 'Quattroporte', 'Levante', 'MC20'] },
  { marca: 'Aston Martin', modelos: ['Vantage', 'DB11', 'DBX'] },
  { marca: 'McLaren', modelos: ['720S', 'GT', 'Artura'] },
  { marca: 'Bentley', modelos: ['Continental', 'Bentayga', 'Flying Spur'] },
  { marca: 'Rolls-Royce', modelos: ['Ghost', 'Phantom', 'Cullinan'] },
  { marca: 'Jaguar', modelos: ['F-Pace', 'E-Pace', 'XE', 'XF'] },
  { marca: 'Land Rover', modelos: ['Defender', 'Discovery', 'Range Rover', 'Evoque', 'Velar'] },
  { marca: 'Mini', modelos: ['Cooper', 'Countryman', 'One'] },
  { marca: 'Volvo', modelos: ['XC40', 'XC60', 'XC90', 'S60'] },
  { marca: 'Saab', modelos: ['9-3', '9-5'] },
  { marca: 'Smart', modelos: ['Fortwo', 'Forfour'] },
  { marca: 'Alfa Romeo', modelos: ['Giulietta', 'Stelvio', 'Giulia', 'MiTo'] },
  { marca: 'Lancia', modelos: ['Ypsilon', 'Delta'] },
  { marca: 'MG', modelos: ['ZR', 'ZS', 'HS', 'MG5'] },
  { marca: 'Geely', modelos: ['Emgrand', 'Coolray', 'GX3'] },
  { marca: 'Chery', modelos: ['Tiggo', 'QQ', 'Fulwin', 'Arrizo'] },
  { marca: 'JAC', modelos: ['S2', 'S3', 'S5', 'T6', 'JS2'] },
  { marca: 'DFSK', modelos: ['Glory', 'Glory 580', 'Seres 3'] },
  { marca: 'Foton', modelos: ['Tunland', 'Gratour'] },
  { marca: 'Baic', modelos: ['X35', 'X55', 'D20'] },
  { marca: 'Changan', modelos: ['CS15', 'CS35', 'CS55', 'MD201'] },
  { marca: 'Great Wall', modelos: ['Wingle', 'Poer'] },
  { marca: 'Haval', modelos: ['H1', 'H2', 'H6', 'Jolion'] },
  { marca: 'Jetour', modelos: ['X70', 'X90', 'X70 Plus'] },
  { marca: 'ZX Auto', modelos: ['Grand Tiger', 'Admiral'] },

  // CLASICOS
  { marca: 'Falcon', modelos: ['Falcon', 'Falcon Sprint'] },
  { marca: 'Torino', modelos: ['Torino 380', 'Torino TS'] },
  { marca: 'Sierra', modelos: ['Sierra GL', 'Sierra Ghia'] },
  { marca: '1500', modelos: ['1500 Standard', '1500 SS'] },
  { marca: 'Gordini', modelos: ['Gordini 850', 'Gordini 1100'] },
  { marca: 'Fiat 600', modelos: ['600 D', '600 R'] },
  { marca: 'Fiat 128', modelos: ['128 IAVA', '128 SE'] },
  { marca: 'Rastrojero', modelos: ['Rastrojero 42', 'Rastrojero Diesel'] },
  { marca: 'De Carlo', modelos: ['De Carlo 700'] },
  { marca: 'Savoia', modelos: ['Savoia Jankov'] },
  { marca: 'Di Tella', modelos: ['Di Tella 700'] },
  { marca: 'Dodge', modelos: ['Dart', 'Valiant', 'Polara'] },
  { marca: 'Pontiac', modelos: ['GTO', 'Firebird'] },
  { marca: 'Plymouth', modelos: ['Valiant', 'Barracuda'] },
  { marca: 'Oldsmobile', modelos: ['Cutlass', 'Delta 88'] },

  // MOTOS
  { marca: 'Yamaha', modelos: ['FZ', 'MT', 'YBR', 'XTZ', 'R6', 'R1', 'NMAX'] },
  { marca: 'Kawasaki', modelos: ['Ninja', 'Z', 'KLR', 'Versys', 'Vulcan'] },
  { marca: 'KTM', modelos: ['Duke', 'RC', 'Adventure', 'EXC'] },
  { marca: 'Ducati', modelos: ['Monster', 'Multistrada', 'Panigale', 'Scrambler'] },
  { marca: 'Harley-Davidson', modelos: ['Sportster', 'Softail', 'Touring', 'Street'] },
  { marca: 'BMW Motos', modelos: ['R 1200', 'F 800', 'G 650', 'S 1000'] },
  { marca: 'Triumph', modelos: ['Bonneville', 'Tiger', 'Street Triple', 'Speed Triple'] },
  { marca: 'Indian', modelos: ['Scout', 'Chief', 'Chieftain', 'FTR'] },
  { marca: 'Royal Enfield', modelos: ['Classic', 'Bullet', 'Himalayan', 'Meteor'] },
  { marca: 'Benelli', modelos: ['TNT', 'TRK', 'Imperiale'] },
  { marca: 'Kymco', modelos: ['Agility', 'Like', 'Downtown'] },
  { marca: 'Gilera', modelos: ['Smash', 'Sahel', 'VC'] },
  { marca: 'Zanella', modelos: ['ZB', 'RX', 'Custom'] },
  { marca: 'Mondial', modelos: ['RD', 'AC', 'LD'] },
  { marca: 'Corven', modelos: ['Energy', 'Hunter', 'Tria'] },
  { marca: 'Cerro', modelos: ['CE 110', 'CE 150'] },
  { marca: 'Bajaj', modelos: ['Rouser', 'Dominar', 'Pulsar'] },
  { marca: 'TVS', modelos: ['Apache', 'Ronin'] },
  { marca: 'Hero', modelos: ['Hunk', 'Glamour', 'Passion'] },

  // CAMIONES
  { marca: 'Iveco', modelos: ['Tector', 'Stralis', 'Cursor', 'Daily'] },
  { marca: 'Scania', modelos: ['R 450', 'G 410', 'P 280'] },
  { marca: 'Volvo Camiones', modelos: ['FH', 'FM', 'VM'] },
  { marca: 'Mercedes-Benz Camiones', modelos: ['Actros', 'Atego', 'Vario'] },
  { marca: 'Ford Cargo', modelos: ['1722', '1517', '1119'] },
  { marca: 'VW Constellation', modelos: ['17280', '24280', '31280'] },

  // MAQUINARIA
  { marca: 'John Deere', modelos: ['6110', '7210', '8370', '5090'] },
  { marca: 'Massey Ferguson', modelos: ['MF 4707', 'MF 6713', 'MF 7722'] },
  { marca: 'New Holland', modelos: ['T7', 'T8', 'TT4'] },
  { marca: 'Case IH', modelos: ['Magnum', 'Maxxum', 'Farmall'] },
  { marca: 'Deutz', modelos: ['Fahr 5', 'Fahr 6'] },
  { marca: 'Valtra', modelos: ['A134', 'BH180', 'T195'] }
];

/**
 * Versiones curadas (~30 pares marca|modelo). Sólo los modelos con variantes
 * conocidas reciben lista cerrada; el resto autogrowea desde el form.
 * Clave de duplicado: (lower(marca), lower(modelo), lower(version)).
 */
const MODELOS_VERSIONES_SEED = [
  { marca: 'Toyota', modelo: 'Hilux', versiones: ['Cabina Simple', 'Doble Cabina', 'GR Sport'] },
  { marca: 'Toyota', modelo: 'Etios', versiones: ['Sedán 4p', 'Hatch 5p', 'XLS'] },
  { marca: 'Toyota', modelo: 'Corolla', versiones: ['XEi', 'XLS', 'SEG'] },
  { marca: 'Toyota', modelo: 'Yaris', versiones: ['Sedán', 'Hatch'] },
  { marca: 'Toyota', modelo: 'RAV4', versiones: ['4x2', '4x4'] },
  { marca: 'Toyota', modelo: 'Land Cruiser', versiones: ['Prado', 'Full'] },
  { marca: 'Ford', modelo: 'Ranger', versiones: ['Cabina Simple', 'Doble Cabina', 'Raptor'] },
  { marca: 'Ford', modelo: 'Focus', versiones: ['4p', '5p'] },
  { marca: 'Ford', modelo: 'Bronco', versiones: ['Sport', 'Badlands'] },
  { marca: 'Volkswagen', modelo: 'Gol', versiones: ['Trendline', 'Comfortline', 'Highline'] },
  { marca: 'Volkswagen', modelo: 'Polo', versiones: ['MSI', 'Highline', 'GTS'] },
  { marca: 'Volkswagen', modelo: 'Amarok', versiones: ['Comfortline', 'Highline', 'V6'] },
  { marca: 'Volkswagen', modelo: 'Nivus', versiones: ['MSI', 'Highline'] },
  { marca: 'Chevrolet', modelo: 'Onix', versiones: ['LT', 'LTZ', 'Premier'] },
  { marca: 'Chevrolet', modelo: 'Cruze', versiones: ['LT', 'LTZ'] },
  { marca: 'Chevrolet', modelo: 'S10', versiones: ['LT', 'LTZ', 'High Country'] },
  { marca: 'Chevrolet', modelo: 'Tracker', versiones: ['LT', 'LTZ', 'Premier'] },
  { marca: 'Renault', modelo: 'Kwid', versiones: ['Life', 'Zen', 'Intens'] },
  { marca: 'Renault', modelo: 'Sandero', versiones: ['Life', 'Zen', 'Intens'] },
  { marca: 'Renault', modelo: 'Duster', versiones: ['Zen', 'Intens', 'Icon'] },
  { marca: 'Fiat', modelo: 'Cronos', versiones: ['1.3', '1.8', 'Precision'] },
  { marca: 'Fiat', modelo: 'Argo', versiones: ['Drive', 'Trekking', 'HGT'] },
  { marca: 'Fiat', modelo: 'Pulse', versiones: ['Drive', 'Audace', 'Impetus'] },
  { marca: 'Peugeot', modelo: '208', versiones: ['Allure', 'Feline', 'GT'] },
  { marca: 'Peugeot', modelo: '2008', versiones: ['Allure', 'Feline', 'GT'] },
  { marca: 'Jeep', modelo: 'Renegade', versiones: ['Sport', 'Longitude', 'Trailhawk'] },
  { marca: 'Jeep', modelo: 'Compass', versiones: ['Sport', 'Longitude', 'Limited'] },
  { marca: 'Nissan', modelo: 'Frontier', versiones: ['S', 'XE', 'LE'] },
  { marca: 'Hyundai', modelo: 'Creta', versiones: ['GLS', 'GL', 'Limited'] },
  { marca: 'Kia', modelo: 'Sportage', versiones: ['LX', 'EX', 'GT'] }
];

/**
 * Mapa de aliases para normalización de marca/modelo. Case-insensitive.
 * Se aplica DESPUES del Title Case sobre la primera palabra del string.
 * Ej: 'vw' → 'Volkswagen', 'chevr' → 'Chevrolet', 'citroen' → 'Citroën'.
 */
const ALIASES_MARCA = {
  'vw': 'Volkswagen',
  'chevro': 'Chevrolet',
  'chevr': 'Chevrolet',
  'merced': 'Mercedes-Benz',
  'bmw': 'BMW',
  'citroen': 'Citroën',
  'mb': 'Mercedes-Benz'
};

/**
 * Normaliza un string de marca o modelo:
 * 1) trim y colapso de espacios en blanco
 * 2) Title Case (primera letra mayúscula, resto minúscula por palabra)
 * 3) Alias map sobre la primera palabra (case-insensitive)
 * Pure fn — sin efectos colaterales, sin llamadas a Sheets.
 *
 * Q2: esta función está DUPLICADA en `AdminPage.html` (inline, dentro del
 * <script>) para evitar round-trip al backend en el pre-submit. Si modificás
 * la lógica acá (incluido el mapa `ALIASES_MARCA`), actualizá también el
 * duplicado en `AdminPage.html` cerca de la línea 4228.
 */
function normalizarMarcaModelo(s) {
  if (s === null || s === undefined) return '';
  const trimmed = String(s).trim().replace(/\s+/g, ' ');
  if (!trimmed) return '';
  const titled = trimmed.split(' ').map(function (w) {
    if (!w) return w;
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  }).join(' ');
  const lowerKey = titled.toLowerCase().split(' ')[0];
  if (Object.prototype.hasOwnProperty.call(ALIASES_MARCA, lowerKey)) {
    const canonical = ALIASES_MARCA[lowerKey];
    return titled.replace(/^\S+/, canonical);
  }
  return titled;
}

/**
 * Crea la hoja Marcas_Modelos idempotentemente. Espejo de crearHojaTurnos.
 * Devuelve {ok, mensaje, accion} donde accion ∈ {creada, headers actualizados, ninguna}.
 */
function crearHojaMarcasModelos() {
  try {
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    let hoja = ss.getSheetByName(SHEETS.MARCAS_MODELOS);
    if (hoja) {
      const headersActuales = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
      const headersEsperados = COLUMNAS.MARCAS_MODELOS;
      const headersCoinciden = headersEsperados.every(function (h, i) { return headersActuales[i] === h; }) &&
        headersActuales.length === headersEsperados.length;
      if (headersCoinciden) {
        return { ok: true, mensaje: 'La hoja Marcas_Modelos ya existe con headers correctos', accion: 'ninguna' };
      }
      hoja.getRange(1, 1, 1, headersEsperados.length)
        .setValues([headersEsperados])
        .setFontWeight('bold')
        .setBackground('#e2e8f0');
      return { ok: true, mensaje: 'La hoja Marcas_Modelos ya existía; headers normalizados', accion: 'headers actualizados' };
    }
    hoja = ss.insertSheet(SHEETS.MARCAS_MODELOS);
    hoja.getRange(1, 1, 1, COLUMNAS.MARCAS_MODELOS.length)
      .setValues([COLUMNAS.MARCAS_MODELOS]);
    hoja.getRange(1, 1, 1, COLUMNAS.MARCAS_MODELOS.length)
      .setFontWeight('bold')
      .setBackground('#e2e8f0');
    hoja.setFrozenRows(1);
    return { ok: true, mensaje: 'Hoja Marcas_Modelos creada correctamente', accion: 'creada' };
  } catch (err) {
    return { ok: false, error: 'No se pudo crear la hoja Marcas_Modelos: ' + (err && err.message ? err.message : err) };
  }
}

/**
 * Crea la hoja Modelos_Versiones idempotentemente. Al crear (no en re-runs),
 * siembra las versiones curadas de MODELOS_VERSIONES_SEED.
 */
function crearHojaModelosVersiones() {
  try {
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    let hoja = ss.getSheetByName(SHEETS.MODELOS_VERSIONES);
    if (hoja) {
      const headersActuales = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
      const headersEsperados = COLUMNAS.MODELOS_VERSIONES;
      const headersCoinciden = headersEsperados.every(function (h, i) { return headersActuales[i] === h; }) &&
        headersActuales.length === headersEsperados.length;
      if (headersCoinciden) {
        return { ok: true, mensaje: 'La hoja Modelos_Versiones ya existe con headers correctos', accion: 'ninguna' };
      }
      hoja.getRange(1, 1, 1, headersEsperados.length)
        .setValues([headersEsperados])
        .setFontWeight('bold')
        .setBackground('#e2e8f0');
      return { ok: true, mensaje: 'La hoja Modelos_Versiones ya existía; headers normalizados', accion: 'headers actualizados' };
    }
    hoja = ss.insertSheet(SHEETS.MODELOS_VERSIONES);
    hoja.getRange(1, 1, 1, COLUMNAS.MODELOS_VERSIONES.length)
      .setValues([COLUMNAS.MODELOS_VERSIONES]);
    hoja.getRange(1, 1, 1, COLUMNAS.MODELOS_VERSIONES.length)
      .setFontWeight('bold')
      .setBackground('#e2e8f0');
    hoja.setFrozenRows(1);
    // Seed de versiones curadas — sólo en creación
    const filasVersiones = [];
    MODELOS_VERSIONES_SEED.forEach(function (par) {
      const marcaN = normalizarMarcaModelo(par.marca);
      const modeloN = normalizarMarcaModelo(par.modelo);
      par.versiones.forEach(function (ver) {
        filasVersiones.push([marcaN, modeloN, ver]);
      });
    });
    if (filasVersiones.length > 0) {
      hoja.getRange(2, 1, filasVersiones.length, 3).setValues(filasVersiones);
    }
    return { ok: true, mensaje: 'Hoja Modelos_Versiones creada correctamente', accion: 'creada', versionesSembradas: filasVersiones.length };
  } catch (err) {
    return { ok: false, error: 'No se pudo crear la hoja Modelos_Versiones: ' + (err && err.message ? err.message : err) };
  }
}

/**
 * Rollback completo del catálogo: elimina ambas hojas si existen y borra
 * la clave de caché. Idempotente — llamar dos veces no produce error.
 */
function eliminarHojasCatalogo() {
  try {
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    [SHEETS.MARCAS_MODELOS, SHEETS.MODELOS_VERSIONES].forEach(function (nombre) {
      const hoja = ss.getSheetByName(nombre);
      if (hoja) ss.deleteSheet(hoja);
    });
    CacheService.getScriptCache().remove('catalogoVehiculos_v1');
    return { ok: true, mensaje: 'Catálogo eliminado' };
  } catch (err) {
    return { ok: false, error: 'No se pudo eliminar el catálogo: ' + (err && err.message ? err.message : err) };
  }
}

/**
 * Reconstruye el árbol del catálogo desde Sheets y lo guarda en cache.
 * Estructura: { ts, marcas, modelosPorMarca, versionesPorModelo }.
 * TTL 6h (21600 s). Read-only respecto de Sheets.
 */
function _rebuildCache_() {
  const tree = { ts: Date.now(), marcas: [], modelosPorMarca: {}, versionesPorModelo: {} };
  try {
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    const modelosPorMarca = {};
    const versionesPorModelo = {};

    const sheetMM = ss.getSheetByName(SHEETS.MARCAS_MODELOS);
    if (sheetMM) {
      const data = sheetMM.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const marcaN = normalizarMarcaModelo(data[i][1]);
        const modeloN = normalizarMarcaModelo(data[i][2]);
        if (!marcaN || !modeloN) continue;
        if (!modelosPorMarca[marcaN]) modelosPorMarca[marcaN] = [];
        if (modelosPorMarca[marcaN].indexOf(modeloN) === -1) modelosPorMarca[marcaN].push(modeloN);
      }
    }

    const sheetMV = ss.getSheetByName(SHEETS.MODELOS_VERSIONES);
    if (sheetMV) {
      const data = sheetMV.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const marcaN = normalizarMarcaModelo(data[i][0]);
        const modeloN = normalizarMarcaModelo(data[i][1]);
        const ver = String(data[i][2] || '').trim();
        if (!marcaN || !modeloN || !ver) continue;
        const key = marcaN + '|' + modeloN;
        if (!versionesPorModelo[key]) versionesPorModelo[key] = [];
        if (versionesPorModelo[key].indexOf(ver) === -1) versionesPorModelo[key].push(ver);
      }
    }

    const marcas = Object.keys(modelosPorMarca).sort();
    tree.marcas = marcas;
    for (let i = 0; i < marcas.length; i++) {
      const m = marcas[i];
      tree.modelosPorMarca[m] = modelosPorMarca[m].slice().sort();
      for (let j = 0; j < tree.modelosPorMarca[m].length; j++) {
        const modelo = tree.modelosPorMarca[m][j];
        const key = m + '|' + modelo;
        tree.versionesPorModelo[key] = (versionesPorModelo[key] || []).slice().sort();
      }
    }

    try {
      CacheService.getScriptCache().put('catalogoVehiculos_v1', JSON.stringify(tree), 21600);
    } catch (cacheErr) {
      Logger.log('Warning: no se pudo escribir cache catalogo: ' + (cacheErr && cacheErr.message ? cacheErr.message : cacheErr));
    }
  } catch (err) {
    Logger.log('Error en _rebuildCache_: ' + (err && err.message ? err.message : err));
  }
  return tree;
}

/**
 * Devuelve el árbol completo desde cache. Si la key falta o está corrupta,
 * reconstruye desde Sheets. Read-only.
 */
function _obtenerOCrearArbol_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('catalogoVehiculos_v1');
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed && parsed.marcas && parsed.modelosPorMarca && parsed.versionesPorModelo) {
        return parsed;
      }
    } catch (err) {
      // cache corrupto, rebuild
    }
  }
  return _rebuildCache_();
}

/**
 * Lista canónica de marcas. Cache-hit evita leer Sheets.
 */
function listarMarcas() {
  try {
    const tree = _obtenerOCrearArbol_();
    return (tree.marcas || []).slice();
  } catch (err) {
    Logger.log('Error en listarMarcas: ' + (err && err.message ? err.message : err));
    return [];
  }
}

/**
 * Lista de modelos de una marca. Array vacío si la marca no existe.
 * Cache-hit evita leer Sheets.
 */
function listarModelos(marca) {
  try {
    const marcaN = normalizarMarcaModelo(marca);
    if (!marcaN) return [];
    const tree = _obtenerOCrearArbol_();
    return (tree.modelosPorMarca[marcaN] || []).slice();
  } catch (err) {
    Logger.log('Error en listarModelos: ' + (err && err.message ? err.message : err));
    return [];
  }
}

/**
 * Lista de versiones curadas de (marca, modelo). Array vacío si el par
 * no tiene variantes curadas (R6) — el frontend debe revelar input libre.
 * Cache-hit evita leer Sheets.
 */
function listarVersiones(marca, modelo) {
  try {
    const marcaN = normalizarMarcaModelo(marca);
    const modeloN = normalizarMarcaModelo(modelo);
    if (!marcaN || !modeloN) return [];
    const tree = _obtenerOCrearArbol_();
    const key = marcaN + '|' + modeloN;
    return (tree.versionesPorModelo[key] || []).slice();
  } catch (err) {
    Logger.log('Error en listarVersiones: ' + (err && err.message ? err.message : err));
    return [];
  }
}

/**
 * Upsert idempotente de un par (marca, modelo) y opcionalmente versión.
 * Q1: si el par lowercase ya existe en Marcas_Modelos, NO upserts (skip).
 * Si llega `version` no-vacía y el trío es nuevo, también appends
 * Modelos_Versiones. Invalida el cache tras upsert exitoso.
 */
function registrarMarcaModelo(p) {
  try {
    const marca = normalizarMarcaModelo(p && p.marca || '');
    const modelo = normalizarMarcaModelo(p && p.modelo || '');
    const version = (p && p.version || '').toString().trim();
    const origen = (p && p.origen) || 'user';
    if (!marca || !modelo) {
      return { ok: false, error: 'Falta marca o modelo' };
    }

    const ss = SpreadsheetApp.openById(obtenerSheetId());
    const sheetMM = ss.getSheetByName(SHEETS.MARCAS_MODELOS);
    if (!sheetMM) return { ok: false, error: 'No existe la hoja Marcas_Modelos' };

    const pairKeyLower = marca.toLowerCase() + '|' + modelo.toLowerCase();
    const data = sheetMM.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const m = String(data[i][1] || '').trim().toLowerCase();
      const mo = String(data[i][2] || '').trim().toLowerCase();
      if (m + '|' + mo === pairKeyLower) {
        // Q1: par ya existe, no upsert
        return { ok: true, accion: 'ya_existente', marca: marca, modelo: modelo };
      }
    }

    const id = nextId(SHEETS.MARCAS_MODELOS, 'MM');
    sheetMM.appendRow([id, marca, modelo, origen, ahoraComoString()]);

    if (version) {
      const sheetMV = ss.getSheetByName(SHEETS.MODELOS_VERSIONES);
      if (sheetMV) {
        const dataMV = sheetMV.getDataRange().getValues();
        const trioLower = marca.toLowerCase() + '|' + modelo.toLowerCase() + '|' + version.toLowerCase();
        let yaExiste = false;
        for (let i = 1; i < dataMV.length; i++) {
          const m = String(dataMV[i][0] || '').trim().toLowerCase();
          const mo = String(dataMV[i][1] || '').trim().toLowerCase();
          const v = String(dataMV[i][2] || '').trim().toLowerCase();
          if (m + '|' + mo + '|' + v === trioLower) { yaExiste = true; break; }
        }
        if (!yaExiste) {
          sheetMV.appendRow([marca, modelo, version]);
        }
      }
    }

    try { CacheService.getScriptCache().remove('catalogoVehiculos_v1'); } catch (e) {}

    return { ok: true, accion: 'insertado', marca: marca, modelo: modelo };
  } catch (err) {
    return { ok: false, error: 'No se pudo registrar la marca/modelo: ' + (err && err.message ? err.message : err) };
  }
}

/**
 * Wrapper defensivo sobre registrarMarcaModelo. R4: una falla de catálogo
 * NO debe romper el alta de vehículo. Loguea y sigue.
 */
function _autoGrowCatalogo_(marca, modelo, version) {
  try {
    if (!marca || !modelo) return;
    const res = registrarMarcaModelo({ marca: marca, modelo: modelo, version: version || '', origen: 'user' });
    if (res && !res.ok) {
      Logger.log('catalogo auto-grow failed: ' + (res.error || 'unknown'));
    }
  } catch (e) {
    Logger.log('catalogo auto-grow failed: ' + (e && e.message ? e.message : e));
  }
}

/**
 * Carga el seed curado (MARCAS_MODELOS_SEED) en Marcas_Modelos. Idempotente
 * vía Q1. Si las hojas no existen, las crea primero. También siembra las
 * versiones curadas (MODELOS_VERSIONES_SEED) en Modelos_Versiones,
 * deduplicando por trío (marca, modelo, version) lowercase.
 */
function cargarCatalogoMarcasModelos() {
  try {
    crearHojaMarcasModelos();
    crearHojaModelosVersiones();

    let insertados = 0;
    let yaExistentes = 0;

    MARCAS_MODELOS_SEED.forEach(function (entry) {
      entry.modelos.forEach(function (modelo) {
        const res = registrarMarcaModelo({ marca: entry.marca, modelo: modelo, origen: 'seed' });
        if (res && res.ok) {
          if (res.accion === 'insertado') insertados++;
          else yaExistentes++;
        }
      });
    });

    // Sembrar versiones curadas directamente, dedup por trío lowercase.
    // Bypass Q1: registrarMarcaModelo no agrega versiones cuando el par
    // ya existe (caso típico del re-seed), por eso este path separado.
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    const sheetMV = ss.getSheetByName(SHEETS.MODELOS_VERSIONES);
    if (sheetMV) {
      const dataMV = sheetMV.getDataRange().getValues();
      const existentes = new Set();
      for (let i = 1; i < dataMV.length; i++) {
        const a = String(dataMV[i][0] || '').trim().toLowerCase();
        const b = String(dataMV[i][1] || '').trim().toLowerCase();
        const c = String(dataMV[i][2] || '').trim().toLowerCase();
        if (a && b && c) existentes.add(a + '|' + b + '|' + c);
      }
      const filasVersiones = [];
      MODELOS_VERSIONES_SEED.forEach(function (par) {
        const marcaN = normalizarMarcaModelo(par.marca);
        const modeloN = normalizarMarcaModelo(par.modelo);
        par.versiones.forEach(function (ver) {
          const verLimpia = String(ver || '').trim();
          if (!verLimpia) return;
          const trioLower = marcaN.toLowerCase() + '|' + modeloN.toLowerCase() + '|' + verLimpia.toLowerCase();
          if (!existentes.has(trioLower)) {
            filasVersiones.push([marcaN, modeloN, verLimpia]);
            existentes.add(trioLower);
          }
        });
      });
      if (filasVersiones.length > 0) {
        const startRow = sheetMV.getLastRow() + 1;
        sheetMV.getRange(startRow, 1, filasVersiones.length, 3).setValues(filasVersiones);
      }
    }

    try { CacheService.getScriptCache().remove('catalogoVehiculos_v1'); } catch (e) {}

    return { ok: true, insertados: insertados, ya_existentes: yaExistentes };
  } catch (err) {
    return { ok: false, error: 'No se pudo cargar el catálogo: ' + (err && err.message ? err.message : err) };
  }
}

/**
 * Migra Marcas/Modelos existentes en Vehiculos al catálogo. Aplica
 * normalización completa (incluye aliases como vw→Volkswagen). Dedup
 * vía Q1 (par lowercase). Devuelve métrica de aliases aplicados.
 */
function migrarMarcasExistentes() {
  try {
    crearHojaMarcasModelos();

    const ss = SpreadsheetApp.openById(obtenerSheetId());
    const sheetV = ss.getSheetByName(SHEETS.VEHICULOS);
    if (!sheetV) return { ok: false, error: 'No existe la hoja Vehiculos' };

    const data = sheetV.getDataRange().getValues();
    let insertados = 0;
    let yaExistentes = 0;
    let aliasesAplicados = 0;

    for (let i = 1; i < data.length; i++) {
      const marcaOrig = String(data[i][5] || '').trim();
      const modeloOrig = String(data[i][6] || '').trim();
      if (!marcaOrig || !modeloOrig) continue;
      const marcaN = normalizarMarcaModelo(marcaOrig);
      const modeloN = normalizarMarcaModelo(modeloOrig);
      if (marcaN !== marcaOrig || modeloN !== modeloOrig) aliasesAplicados++;
      const res = registrarMarcaModelo({ marca: marcaN, modelo: modeloN, origen: 'migration' });
      if (res && res.ok) {
        if (res.accion === 'insertado') insertados++;
        else yaExistentes++;
      }
    }

    return {
      ok: true,
      insertados: insertados,
      ya_existentes: yaExistentes,
      aliases_aplicados: aliasesAplicados
    };
  } catch (err) {
    return { ok: false, error: 'No se pudieron migrar las marcas: ' + (err && err.message ? err.message : err) };
  }
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

function cargarTurnosDemo() {
  try {
    const sheetRes = crearHojaTurnos();
    if (!sheetRes.ok) return { ok: false, error: 'No se pudo preparar la hoja Turnos: ' + sheetRes.error };

    const existentes = sheetAObjetos(SHEETS.TURNOS);
    if (existentes.length > 0) {
      return { ok: false, mensaje: 'Ya hay turnos cargados. Limpiá primero con limpiarTurnosDemo().', yaCargado: true };
    }

    const vehiculos = sheetAObjetos(SHEETS.VEHICULOS).filter(function (v) { return esActivo(v.Activo); });
    if (vehiculos.length === 0) {
      return { ok: false, error: 'No hay vehículos cargados. Corré primero cargarDatosDemoMasivos().' };
    }

    const tiposServicio = COLUMNAS_TIPOS_SERVICIO || [
      'Cambio de aceite', 'Service completo', 'Diagnostico', 'Freno', 'Embrague',
      'Distribucion', 'Suspension', 'Tren delantero', 'Aire acondicionado', 'Bateria', 'Otro'
    ];

    const descripcionesPorTipo = {
      'Cambio de aceite': 'Cliente pide también revisar filtro de aire',
      'Service completo': 'Service completo según manual',
      'Diagnostico': 'Vehículo con ruido extraño en el motor',
      'Freno': 'Pastillas gastadas, pedido de revisión',
      'Embrague': 'Embrague patinando',
      'Distribucion': 'Cambio de kit de distribución',
      'Suspension': 'Amortiguadores vencidos',
      'Tren delantero': 'Juego de tren delantero nuevo',
      'Aire acondicionado': 'A/C no enfría correctamente',
      'Bateria': 'Cambio de batería preventiva',
      'Otro': 'Revisión general'
    };

    const ahora = new Date();
    const semillaFecha = function (diasOffset, hora, minuto) {
      const d = new Date(ahora);
      d.setDate(d.getDate() + diasOffset);
      d.setHours(hora, minuto, 0, 0);
      return d;
    };

    const plan = [
      { dias: -28, hora: 9, min: 0, tipo: 'Service completo', estado: 'completado', duracion: 120, notas: 'Turno demo completado' },
      { dias: -21, hora: 10, min: 30, tipo: 'Cambio de aceite', estado: 'completado', duracion: 60, notas: '' },
      { dias: -14, hora: 14, min: 0, tipo: 'Freno', estado: 'completado', duracion: 90, notas: '' },
      { dias: -7, hora: 11, min: 0, tipo: 'Diagnostico', estado: 'completado', duracion: 45, notas: '' },
      { dias: -3, hora: 16, min: 0, tipo: 'Service completo', estado: 'cancelado', duracion: 120, notas: 'Cliente canceló a último momento' },
      { dias: -2, hora: 9, min: 30, tipo: 'Embrague', estado: 'no_show', duracion: 180, notas: 'Cliente no se presentó' },
      { dias: 0, hora: 9, min: 0, tipo: 'Cambio de aceite', estado: 'confirmado', duracion: 60, notas: '' },
      { dias: 1, hora: 11, min: 0, tipo: 'Diagnostico', estado: 'confirmado', duracion: 45, notas: '' },
      { dias: 1, hora: 15, min: 0, tipo: 'Freno', estado: 'confirmado', duracion: 90, notas: 'Traer llave de repuesto' },
      { dias: 2, hora: 10, min: 0, tipo: 'Aire acondicionado', estado: 'confirmado', duracion: 90, notas: '' },
      { dias: 3, hora: 14, min: 30, tipo: 'Distribucion', estado: 'confirmado', duracion: 240, notas: '' },
      { dias: 5, hora: 9, min: 0, tipo: 'Service completo', estado: 'pendiente', duracion: 120, notas: '' },
      { dias: 7, hora: 16, min: 0, tipo: 'Cambio de aceite', estado: 'pendiente', duracion: 60, notas: '' },
      { dias: 9, hora: 10, min: 30, tipo: 'Suspension', estado: 'pendiente', duracion: 150, notas: '' },
      { dias: 12, hora: 11, min: 0, tipo: 'Tren delantero', estado: 'pendiente', duracion: 120, notas: '' },
      { dias: 15, hora: 9, min: 0, tipo: 'Service completo', estado: 'pendiente', duracion: 120, notas: '' },
      { dias: 18, hora: 14, min: 0, tipo: 'Bateria', estado: 'pendiente', duracion: 30, notas: '' },
      { dias: 21, hora: 10, min: 0, tipo: 'Diagnostico', estado: 'pendiente', duracion: 45, notas: '' },
      { dias: 24, hora: 15, min: 30, tipo: 'Cambio de aceite', estado: 'pendiente', duracion: 60, notas: '' },
      { dias: 28, hora: 9, min: 30, tipo: 'Otro', estado: 'pendiente', duracion: 60, notas: 'Revisión pre-viaje' }
    ];

    const sheet = obtenerSheet(SHEETS.TURNOS);
    const idsCreados = [];

    for (let i = 0; i < plan.length; i++) {
      const item = plan[i];
      const vehiculo = vehiculos[i % vehiculos.length];
      const id = nextId(SHEETS.TURNOS, 'T');
      const fechaHora = semillaFecha(item.dias, item.hora, item.min);
      const descripcion = descripcionesPorTipo[item.tipo] || '';
      sheet.appendRow([
        id,
        fechaHora,
        item.duracion,
        vehiculo.ID_Vehiculo,
        item.tipo,
        descripcion,
        item.estado,
        '',
        ahoraComoString(),
        item.notas || ''
      ]);
      idsCreados.push(id);
    }

    return {
      ok: true,
      mensaje: 'Turnos demo cargados correctamente',
      turnosCreados: idsCreados.length,
      ids: idsCreados,
      distribucion: {
        completados: plan.filter(function (p) { return p.estado === 'completado'; }).length,
        confirmados: plan.filter(function (p) { return p.estado === 'confirmado'; }).length,
        pendientes: plan.filter(function (p) { return p.estado === 'pendiente'; }).length,
        cancelados: plan.filter(function (p) { return p.estado === 'cancelado'; }).length,
        noShow: plan.filter(function (p) { return p.estado === 'no_show'; }).length
      }
    };
  } catch (err) {
    return { ok: false, error: 'No se pudieron cargar los turnos demo: ' + (err && err.message ? err.message : err) };
  }
}

function limpiarTurnosDemo() {
  try {
    const sheet = obtenerSheet(SHEETS.TURNOS);
    if (!sheet) return { ok: false, error: 'No existe la hoja Turnos' };
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { ok: true, mensaje: 'La hoja Turnos ya estaba vacía', turnosEliminados: 0 };
    }
    let eliminados = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      sheet.deleteRow(i + 1);
      eliminados++;
    }
    return { ok: true, mensaje: 'Se eliminaron todos los turnos', turnosEliminados: eliminados };
  } catch (err) {
    return { ok: false, error: 'No se pudieron limpiar los turnos: ' + (err && err.message ? err.message : err) };
  }
}

const COLUMNAS_TIPOS_SERVICIO = [
  'Cambio de aceite', 'Service completo', 'Diagnostico', 'Freno', 'Embrague',
  'Distribucion', 'Suspension', 'Tren delantero', 'Aire acondicionado', 'Bateria', 'Otro'
];

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

  const sheetTurnos = ss.getSheetByName('Turnos');
  if (sheetTurnos) {
    const data = sheetTurnos.getDataRange().getValues();
    const idVehiculoCol = data[0].indexOf('ID_Vehiculo');

    for (let i = data.length - 1; i >= 1; i--) {
      const idVehiculo = String(data[i][idVehiculoCol] || '').trim();
      if (vehiculosEliminadosIds.indexOf(idVehiculo) !== -1) {
        sheetTurnos.deleteRow(i + 1);
        reporte.turnosEliminados = (reporte.turnosEliminados || 0) + 1;
      }
    }
  }

  const presupuestosEliminadosIds = [];
  const sheetPresupuestos = ss.getSheetByName('Presupuestos');
  if (sheetPresupuestos) {
    const data = sheetPresupuestos.getDataRange().getValues();
    const idCol = data[0].indexOf('ID_Presupuesto');
    const idVehiculoCol = data[0].indexOf('ID_Vehiculo');

    for (let i = data.length - 1; i >= 1; i--) {
      const idVehiculo = String(data[i][idVehiculoCol] || '').trim();
      if (vehiculosEliminadosIds.indexOf(idVehiculo) !== -1) {
        presupuestosEliminadosIds.push(String(data[i][idCol]));
        sheetPresupuestos.deleteRow(i + 1);
        reporte.presupuestosEliminados = (reporte.presupuestosEliminados || 0) + 1;
      }
    }
  }

  const sheetPresupuestosItems = ss.getSheetByName('Presupuestos_Items');
  if (sheetPresupuestosItems) {
    const data = sheetPresupuestosItems.getDataRange().getValues();
    const idPresupuestoCol = data[0].indexOf('ID_Presupuesto');

    for (let i = data.length - 1; i >= 1; i--) {
      const idPresupuesto = String(data[i][idPresupuestoCol] || '').trim();
      if (presupuestosEliminadosIds.indexOf(idPresupuesto) !== -1) {
        sheetPresupuestosItems.deleteRow(i + 1);
        reporte.presupuestosItemsEliminados = (reporte.presupuestosItemsEliminados || 0) + 1;
      }
    }
  }

  reporte.clientesRestantes = Math.max(0, sheetClientes.getLastRow() - 1);
  reporte.vehiculosRestantes = Math.max(0, sheetVehiculos.getLastRow() - 1);
  reporte.serviciosRestantes = Math.max(0, sheetServicios.getLastRow() - 1);
  const sheetTurnosRef = ss.getSheetByName('Turnos');
  reporte.turnosRestantes = sheetTurnosRef ? Math.max(0, sheetTurnosRef.getLastRow() - 1) : 0;
  const sheetPresupuestosRef = ss.getSheetByName('Presupuestos');
  reporte.presupuestosRestantes = sheetPresupuestosRef ? Math.max(0, sheetPresupuestosRef.getLastRow() - 1) : 0;
  const sheetPresupuestosItemsRef = ss.getSheetByName('Presupuestos_Items');
  reporte.presupuestosItemsRestantes = sheetPresupuestosItemsRef ? Math.max(0, sheetPresupuestosItemsRef.getLastRow() - 1) : 0;

  return {
    ok: true,
    mensaje: 'Limpieza completa. El Sheet queda listo para datos reales.',
    reporte: reporte
  };
}

/* ============================================================
 * Batería de pruebas del catálogo de Marcas/Modelos/Versión.
 * Lote C — sdd-apply task 7.1. Detrás de `DEBUG = true` para no
 * ejecutarse en producción. Sólo modificable a mano en el editor.
 * ============================================================ */

const DEBUG = false;

/**
 * Suite de integración del catálogo. Cubre:
 *  1. Idempotencia del seed (cargarCatalogoMarcasModelos ×2).
 *  2. Auto-grow independiente (altaVehiculoInterno → catálogo + Vehiculos).
 *  3. Migration aliases (vw → Volkswagen).
 *  4. Cache hit/miss (2da lectura de listarMarcas).
 *
 * Devuelve {ok, tests_run, tests_passed, failures[]}.
 */
function probarCatalogoVehiculos() {
  if (!DEBUG) {
    Logger.log('SKIP: probarCatalogoVehiculos() requires DEBUG=true');
    return { ok: false, skip: true, mensaje: 'Setear DEBUG = true en Codigo.gs' };
  }

  const testsRun = [];
  const failures = [];

  /* --- Test 1: Seed idempotencia ---------------------------- */
  Logger.log('===== TEST 1: Seed idempotencia =====');
  try {
    const r1 = cargarCatalogoMarcasModelos();
    Logger.log('Primera corrida: insertados=' + r1.insertados + ' ya_existentes=' + r1.ya_existentes);
    const r2 = cargarCatalogoMarcasModelos();
    Logger.log('Segunda corrida: insertados=' + r2.insertados + ' ya_existentes=' + r2.ya_existentes);
    testsRun.push('seed_idempotency');
    if (r2.insertados !== 0) {
      failures.push('seed_idempotency: 2da corrida insertó ' + r2.insertados + ' (esperado 0)');
    }
  } catch (e) {
    failures.push('seed_idempotency: excepción ' + e.message);
  }

  /* --- Test 2: Auto-grow independence ------------------------ */
  Logger.log('===== TEST 2: Auto-grow independence =====');
  try {
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    const sheetV = ss.getSheetByName(SHEETS.VEHICULOS);
    const sheetMM = ss.getSheetByName(SHEETS.MARCAS_MODELOS);
    if (!sheetV || !sheetMM) {
      Logger.log('WARN: Faltan hojas Vehiculos o Marcas_Modelos — skip test auto-grow');
      failures.push('auto_grow: hojas faltantes');
    } else {
      const idCliente = 'C-TEST-' + Date.now();
      const rAlta = altaVehiculoInterno({
        patente: 'TEST001',
        marca: 'Lamborghini',
        modelo: 'Urus',
        anio: '2024',
        combustible: 'Nafta',
        idCliente: idCliente,
        version: 'Suv'
      });
      Logger.log('altaVehiculoInterno result: ' + JSON.stringify(rAlta));
      testsRun.push('auto_grow');
      if (!rAlta.ok) {
        Logger.log('WARN: altaVehiculoInterno falló (idCliente puede no existir) — verificamos crecimiento del catálogo al menos');
        const dataMM = sheetMM.getDataRange().getValues();
        const found = dataMM.some(function (row) {
          return String(row[1]).toLowerCase() === 'lamborghini' && String(row[2]).toLowerCase() === 'urus';
        });
        if (!found) {
          failures.push('auto_grow: par Lamborghini/Urus no quedó en Marcas_Modelos');
        }
      }
    }
  } catch (e) {
    failures.push('auto_grow: excepción ' + e.message);
  }

  /* --- Test 3: Migration aliases ---------------------------- */
  Logger.log('===== TEST 3: Migration aliases (vw → Volkswagen) =====');
  try {
    const ss = SpreadsheetApp.openById(obtenerSheetId());
    const sheetV = ss.getSheetByName(SHEETS.VEHICULOS);
    const sheetMM = ss.getSheetByName(SHEETS.MARCAS_MODELOS);
    if (sheetV && sheetMM) {
      const dataV = sheetV.getDataRange().getValues();
      let foundVw = false;
      for (let i = 1; i < dataV.length; i++) {
        const marcaOrig = String(dataV[i][5] || '').trim();
        if (marcaOrig === 'vw' || marcaOrig === 'VW') { foundVw = true; break; }
      }
      if (!foundVw) {
        Logger.log('No hay fila con Marca=vw/VW en Vehiculos — no se puede probar alias in-situ');
        Logger.log('WARN: Test 3 no concluyente (sin fila de prueba vw)');
        failures.push('migration_aliases: sin fila de prueba');
      } else {
        const rMig = migrarMarcasExistentes();
        Logger.log('migrarMarcasExistentes: ' + JSON.stringify(rMig));
        const dataMM = sheetMM.getDataRange().getValues();
        const hasVwNormalized = dataMM.some(function (row) {
          return String(row[1]).toLowerCase() === 'volkswagen';
        });
        testsRun.push('migration_aliases');
        if (!hasVwNormalized) {
          failures.push('migration_aliases: Volkswagen no quedó en Marcas_Modelos tras migrar vw');
        }
      }
    }
  } catch (e) {
    failures.push('migration_aliases: excepción ' + e.message);
  }

  /* --- Test 4: Cache hit ------------------------------------- */
  Logger.log('===== TEST 4: Cache hit =====');
  try {
    CacheService.getScriptCache().remove('catalogoVehiculos_v1');
    Logger.log('Cache invalidada manualmente antes del test');
    const t0 = Date.now();
    const m1 = listarMarcas();
    const dt1 = Date.now() - t0;
    Logger.log('1ra lectura (cache MISS esperado): ' + m1.length + ' marcas, ' + dt1 + ' ms');
    const t1 = Date.now();
    const m2 = listarMarcas();
    const dt2 = Date.now() - t1;
    Logger.log('2da lectura (cache HIT esperado per spec R5): ' + m2.length + ' marcas, ' + dt2 + ' ms');
    testsRun.push('cache_hit');
    if (dt2 > dt1 - 50) {
      failures.push('cache_hit: cache hit no mas rapido que miss: dt1=' + dt1 + 'ms dt2=' + dt2 + 'ms');
    }
  } catch (e) {
    failures.push('cache_hit: excepción ' + e.message);
  }

  const testsPassed = testsRun.length - failures.length;
  Logger.log('===== RESUMEN =====');
  Logger.log('tests_run=' + testsRun.length + ' tests_passed=' + testsPassed + ' failures=' + failures.length);
  if (failures.length) {
    Logger.log('FAILURES: ' + JSON.stringify(failures));
  }

  return {
    ok: failures.length === 0,
    tests_run: testsRun.length,
    tests_passed: testsPassed,
    failures: failures
  };
}
