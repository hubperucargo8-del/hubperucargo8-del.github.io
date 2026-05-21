/**
 * ============================================================
 * HUB PERU CARGO — Backend Google Apps Script
 * ============================================================
 *
 * INSTRUCCIONES DE CONFIGURACIÓN:
 *
 * PASO 1: Crea una Google Sheet
 *   - Ve a https://sheets.google.com
 *   - Crea una hoja nueva llamada "HPC - Base de Datos"
 *   - Copia el ID de la URL (la parte entre /d/ y /edit)
 *   - Pégalo en la variable SHEET_ID más abajo
 *
 * PASO 2: Crea el Apps Script
 *   - Dentro de la hoja, ve a: Extensiones → Apps Script
 *   - Borra todo el código existente
 *   - Pega TODO este archivo
 *   - Cambia SHEET_ID y EMAIL_NOTIFICACION
 *
 * PASO 3: Despliega como Web App
 *   - Clic en "Desplegar" → "Nueva implementación"
 *   - Tipo: Aplicación web
 *   - Ejecutar como: Yo (tu email)
 *   - Quién puede acceder: Cualquier usuario
 *   - Haz clic en "Desplegar"
 *   - COPIA la URL que aparece
 *
 * PASO 4: Configura la URL en tus páginas HTML
 *   - En index.html: busca ENDPOINT y pega la URL
 *   - En reclamaciones.html: busca ENDPOINT y pega la URL
 *
 * ============================================================
 */

// ============================================================
//  ⚙️ CONFIGURACIÓN — CAMBIA ESTOS VALORES
// ============================================================

// ID de tu Google Sheet (lo encuentras en la URL de la hoja)
// Ejemplo: https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit
const SHEET_ID = '1tzK_NlurDknbtRCKWX5CPG6_xdX72AixUzFlvNQWwlc';

// Tu email para recibir notificaciones de nuevas reclamaciones
const EMAIL_NOTIFICACION = 'ventas@hubperucargo.com';

// Nombre de la empresa (para el asunto de emails)
const EMPRESA = 'Hub Peru Cargo S.A.C.';

// ============================================================
//  HANDLER PRINCIPAL — recibe peticiones POST del sitio web
// ============================================================

function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const body = JSON.parse(raw);

    if (body.type === 'tracking') {
      guardarTracking(body);
    } else if (body.type === 'reclamacion') {
      guardarReclamacion(body.datos || body);
    } else if (body.type === 'contacto') {
      guardarContacto(body.datos || body);
    }

    return respuesta({ ok: true });

  } catch (err) {
    console.error('Error en doPost:', err.toString());
    return respuesta({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  // Endpoint de prueba — abre la URL en el navegador para verificar que funciona
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', empresa: EMPRESA, timestamp: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  TRACKING — guarda eventos de comportamiento del usuario
// ============================================================

function guardarTracking(payload) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, '📊 Tracking', [
    'Fecha/Hora', 'Evento', 'Sesión', 'Tab Actual', 'Datos', 'URL', 'Timestamp ISO',
  ]);

  sheet.appendRow([
    formatFecha(payload.timestamp),
    payload.evento || payload.event || '',
    payload.sesion || payload.session || '',
    payload.tab_actual || '',
    JSON.stringify(payload.datos || payload.data || {}),
    payload.url || '',
    payload.timestamp || '',
  ]);
}

// ============================================================
//  RECLAMACIONES — guarda y notifica
// ============================================================

function guardarReclamacion(datos) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, '📋 Reclamaciones', [
    'N° Reclamación', 'Fecha Registro Lima', 'Estado', 'Tipo',
    'Nombre', 'Documento', 'Email', 'Teléfono', 'Domicilio',
    'Empresa Reclamante', 'RUC Empresa',
    'Monto Reclamado (USD)', 'N° Pedido / BL', 'Fecha Servicio',
    'Tipo de Servicio', 'Servicio Contratado',
    'Descripción de Hechos', 'Solución Esperada',
    'Desea Notificación', 'Timestamp ISO',
  ]);

  sheet.appendRow([
    datos.numero_reclamacion || generarNumero(),
    datos.fecha_registro     || formatFecha(datos.timestamp_iso),
    datos.estado             || 'PENDIENTE',
    datos.tipo               || '',
    datos.nombre             || '',
    datos.documento          || '',
    datos.email              || '',
    datos.telefono           || '',
    datos.domicilio          || '',
    datos.empresa            || '',
    datos.ruc_empresa        || '',
    datos.monto_reclamado_usd || '',
    datos.pedido_referencia  || '',
    datos.fecha_servicio     || '',
    datos.tipo_servicio      || '',
    datos.servicio_contratado || '',
    datos.descripcion        || '',
    datos.solucion_esperada  || '',
    datos.desea_notificacion || 'NO',
    datos.timestamp_iso      || new Date().toISOString(),
  ]);

  // Enviar email de notificación
  try {
    enviarEmailReclamacion(datos);
  } catch(err) {
    console.warn('No se pudo enviar email:', err.toString());
  }
}

// ============================================================
//  CONTACTO — guarda mensajes del formulario de contacto
// ============================================================

function guardarContacto(datos) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, '📩 Contacto', [
    'Fecha/Hora', 'Nombre', 'Empresa', 'Email', 'Teléfono', 'Servicio', 'Mensaje', 'Sesión',
  ]);

  sheet.appendRow([
    formatFecha(datos.timestamp || new Date().toISOString()),
    datos.nombre   || '',
    datos.empresa  || '',
    datos.email    || '',
    datos.telefono || '',
    datos.servicio || '',
    datos.mensaje  || '',
    datos.sesion   || '',
  ]);
}

// ============================================================
//  EMAIL DE NOTIFICACIÓN PARA RECLAMACIONES
// ============================================================

function enviarEmailReclamacion(datos) {
  const num    = datos.numero_reclamacion || 'N/A';
  const tipo   = datos.tipo || 'RECLAMO';
  const nombre = datos.nombre || 'No especificado';
  const fecha  = datos.fecha_registro || new Date().toLocaleString('es-PE');

  const asunto = `🔴 Nueva ${tipo} registrada — ${num} — ${EMPRESA}`;

  const cuerpo = `
Nueva reclamación registrada en el Libro de Reclamaciones Virtual.

═══════════════════════════════════════
NÚMERO DE CASO: ${num}
TIPO:           ${tipo}
ESTADO:         PENDIENTE
FECHA:          ${fecha}
═══════════════════════════════════════

DATOS DEL CONSUMIDOR:
  Nombre:     ${datos.nombre || '-'}
  Documento:  ${datos.documento || '-'}
  Email:      ${datos.email || '-'}
  Teléfono:   ${datos.telefono || '-'}
  Domicilio:  ${datos.domicilio || '-'}
  Empresa:    ${datos.empresa || '-'}

SERVICIO RECLAMADO:
  Tipo:       ${datos.tipo_servicio || '-'}
  Detalle:    ${datos.servicio_contratado || '-'}
  Pedido/BL:  ${datos.pedido_referencia || '-'}
  Monto:      USD ${datos.monto_reclamado_usd || '0'}

DESCRIPCIÓN:
${datos.descripcion || 'No especificada'}

SOLUCIÓN ESPERADA:
${datos.solucion_esperada || 'No especificada'}

═══════════════════════════════════════
IMPORTANTE: Debes dar respuesta en máximo 30 días calendario
conforme a la Ley N° 29571 y D.S. N° 011-2011-PCM.
═══════════════════════════════════════

Este email fue generado automáticamente por el sistema de
Libro de Reclamaciones Virtual de ${EMPRESA}.
`;

  MailApp.sendEmail({
    to:      EMAIL_NOTIFICACION,
    subject: asunto,
    body:    cuerpo,
    replyTo: datos.email || EMAIL_NOTIFICACION,
  });
}

// ============================================================
//  FUNCIONES DE UTILIDAD
// ============================================================

function getOrCreateSheet(ss, nombre, cabeceras) {
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    // Agregar cabeceras con estilo
    const headerRange = sheet.getRange(1, 1, 1, cabeceras.length);
    headerRange.setValues([cabeceras]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#003272');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    // Auto-ajustar columnas
    for (let i = 1; i <= cabeceras.length; i++) {
      sheet.setColumnWidth(i, 180);
    }
  }
  return sheet;
}

function formatFecha(isoString) {
  if (!isoString) return new Date().toLocaleString('es-PE', { timeZone: 'America/Lima' });
  try {
    return new Date(isoString).toLocaleString('es-PE', { timeZone: 'America/Lima' });
  } catch(e) {
    return isoString;
  }
}

function generarNumero() {
  const d = new Date();
  const p = n => String(n).padStart(2,'0');
  return `HPC-REC-${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}-${Math.floor(Math.random()*90000+10000)}`;
}

function respuesta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  FUNCIÓN DE PRUEBA — ejecuta esto manualmente para verificar
// ============================================================

function testConexion() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log('✅ Conectado a: ' + ss.getName());

  // Prueba de tracking
  guardarTracking({
    evento: 'test_conexion',
    sesion: 'TEST-001',
    tab_actual: 'inicio',
    datos: { prueba: true },
    url: 'http://localhost',
    timestamp: new Date().toISOString(),
  });

  // Prueba de reclamación
  guardarReclamacion({
    numero_reclamacion: 'HPC-REC-TEST-00001',
    tipo: 'RECLAMO',
    nombre: 'Usuario de Prueba',
    documento: '12345678',
    email: EMAIL_NOTIFICACION,
    telefono: '+51 999 999 999',
    domicilio: 'Lima, Perú',
    empresa: '',
    servicio_contratado: 'TEST — Prueba de conexión',
    descripcion: 'Esta es una reclamación de prueba generada por la función testConexion().',
    solucion_esperada: 'Verificar que el sistema funciona correctamente.',
    estado: 'PRUEBA',
    fecha_registro: new Date().toLocaleString('es-PE'),
    timestamp_iso: new Date().toISOString(),
  });

  Logger.log('✅ Pruebas escritas en la hoja. Revisa Google Sheets.');
}
