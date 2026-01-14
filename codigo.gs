/**
 * SISTEMA DE MONITOREO DE RED v6.0
 * Google Apps Script para monitoreo de hosts on-premise
 * Soporte para HTTP, HTTPS y puertos personalizados
 * Cálculo automático de disponibilidad y SLA con checkpoints diarios
 * Autor: Sistema de Monitoreo Institucional
 * Última actualización: 2025-01-11
 */

// ============================================
// CONFIGURACIÓN DE NOMBRES DE HOJAS
// ============================================
const SHEET_NAMES = {
  ESTADO: 'current state',
  HISTORIAL: 'story',
  CONFIG: 'configuration',
  REPORTES: 'Reports'
};

// Colores para diferenciar registros
const COLORES = {
  ACTIVO: '#d9ead3',           // Verde claro - Servicio activo
  INACTIVO: '#f4cccc',         // Rojo claro - Servicio caído
  CHECKPOINT_ACTIVO: '#b6d7a8', // Verde más oscuro - Checkpoint activo
  CHECKPOINT_INACTIVO: '#e06666' // Rojo más oscuro - Checkpoint inactivo
};

// ============================================
// FUNCIÓN PRINCIPAL - SE EJECUTA CADA 10 MIN
// ============================================
function monitorearRed() {
  try {
    Logger.log('=== Iniciando ciclo de monitoreo ===');
    const startTime = new Date();
    
    // Obtener configuración
    const config = obtenerConfiguracion();
    const hosts = obtenerHostsAMonitorear();
    
    if (hosts.length === 0) {
      Logger.log('No hay hosts configurados para monitorear');
      return;
    }
    
    // Monitorear cada host
    hosts.forEach(host => {
      monitorearHost(host, config);
    });
    
    // Actualizar reportes
    actualizarReportes();
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    Logger.log(`=== Monitoreo completado en ${duration} segundos ===`);
    
  } catch (error) {
    Logger.log('ERROR en monitoreo: ' + error.toString());
    enviarAlertaError(error);
  }
}

// ============================================
// CONSOLIDAR TIEMPOS DIARIOS (1 VEZ AL DÍA)
// ============================================
function consolidarTiemposDiarios() {
  try {
    Logger.log('=== Iniciando consolidación de tiempos diarios ===');
    const startTime = new Date();
    
    registrarCheckpointsDeTiempo();
    actualizarReportes();
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    Logger.log(`=== Consolidación completada en ${duration} segundos ===`);
    
  } catch (error) {
    Logger.log('ERROR en consolidación: ' + error.toString());
    enviarAlertaError(error);
  }
}

// ============================================
// REGISTRAR CHECKPOINTS DE TIEMPO
// ============================================
function registrarCheckpointsDeTiempo() {
  Logger.log('--- Registrando checkpoints de tiempo ---');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const estadoSheet = ss.getSheetByName(SHEET_NAMES.ESTADO);
  
  if (!estadoSheet) {
    Logger.log('No existe hoja "Estado Actual"');
    return;
  }
  
  const data = estadoSheet.getDataRange().getValues();
  
  // Iterar desde la fila 2 (saltar encabezados)
  for (let i = 1; i < data.length; i++) {
    const fechaHoraActual = data[i][0];
    const nombreHost = data[i][1];
    const protocolo = data[i][2];
    const ip = data[i][3];
    const puerto = data[i][4];
    const estado = data[i][5];
    const desde = data[i][6];
    
    if (!nombreHost || !estado || !desde) {
      continue; // Saltar filas vacías o incompletas
    }
    
    Logger.log(`Checkpoint para ${nombreHost}: Estado ${estado}`);
    
    // Calcular tiempo desde el último registro
    const ahora = new Date();
    const fechaDesde = new Date(desde);
    const diferenciaMs = ahora - fechaDesde;
    const minutos = Math.round(diferenciaMs / 60000);
    
    if (minutos < 1) {
      Logger.log(`  → Tiempo muy corto (${minutos} min), saltando`);
      continue;
    }
    
    const tiempoLegible = formatearTiempo(minutos);
    
    // Crear observación según el estado
    let observacion;
    if (estado === 'ACTIVO') {
      observacion = `Checkpoint de tiempo - Servicio activo (acumulado: ${tiempoLegible})`;
    } else {
      observacion = `Checkpoint de tiempo - Servicio caído (acumulado: ${tiempoLegible})`;
    }
    
    // Registrar checkpoint en historial
    registrarCheckpointEnHistorial({
      nombre: nombreHost,
      protocolo: protocolo,
      ip: ip,
      puerto: puerto
    }, estado, desde, minutos, tiempoLegible, observacion);
    
    // Actualizar "Estado Actual" con nuevos timestamps
    actualizarTimestampsEstadoActual(i + 1, ahora);
    
    Logger.log(`  ✓ Checkpoint registrado: ${minutos} minutos (${tiempoLegible})`);
  }
  
  Logger.log('--- Checkpoints completados ---');
}

// ============================================
// REGISTRAR CHECKPOINT EN HISTORIAL
// ============================================
function registrarCheckpointEnHistorial(host, estado, fechaDesde, minutos, tiempoLegible, observacion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historialSheet = ss.getSheetByName(SHEET_NAMES.HISTORIAL);
  
  if (!historialSheet) {
    historialSheet = ss.insertSheet(SHEET_NAMES.HISTORIAL);
    historialSheet.appendRow([
      'Fecha Hora Actual',
      'Nombre Host',
      'Protocolo',
      'IP',
      'Puerto',
      'Estado Anterior',
      'Fecha Hora Estado Anterior',
      'Estado Nuevo',
      'Tiempo en Minutos',
      'Tiempo Legible',
      'Observación'
    ]);
    historialSheet.getRange('A1:K1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  }
  
  const ahora = new Date();
  const fechaHoraActual = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const puertoDisplay = host.puerto || '-';
  
  historialSheet.appendRow([
    fechaHoraActual,
    host.nombre,
    host.protocolo.toUpperCase(),
    host.ip,
    puertoDisplay,
    estado, // Estado Anterior = Estado Nuevo (es un checkpoint)
    fechaDesde,
    estado, // Estado Nuevo = mismo estado
    minutos,
    tiempoLegible,
    observacion
  ]);
  
  // Aplicar color diferente para checkpoints
  const ultimaFila = historialSheet.getLastRow();
  const color = estado === 'ACTIVO' ? COLORES.CHECKPOINT_ACTIVO : COLORES.CHECKPOINT_INACTIVO;
  historialSheet.getRange(ultimaFila, 1, 1, 11).setBackground(color);
}

// ============================================
// ACTUALIZAR TIMESTAMPS EN ESTADO ACTUAL
// ============================================
function actualizarTimestampsEstadoActual(fila, ahora) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const estadoSheet = ss.getSheetByName(SHEET_NAMES.ESTADO);
  
  const fechaHoraActual = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  // Actualizar columna "Fecha Hora" (columna A)
  estadoSheet.getRange(fila, 1).setValue(fechaHoraActual);
  
  // Actualizar columna "Desde" (columna G) con la fecha actual
  estadoSheet.getRange(fila, 7).setValue(fechaHoraActual);
}

// ============================================
// OBTENER CONFIGURACIÓN DESDE HOJA
// ============================================
function obtenerConfiguracion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  
  if (!configSheet) {
    throw new Error('Hoja "Configuración" no encontrada');
  }
  
  const data = configSheet.getDataRange().getValues();
  const config = {};
  
  for (let i = 1; i < data.length; i++) {
    const parametro = data[i][0];
    const valor = data[i][1];
    
    if (parametro && valor) {
      config[parametro.toString().trim()] = valor.toString().trim();
    }
    
    if (parametro === 'Nombre Host') break;
  }
  
  Logger.log('Configuración cargada');
  return config;
}

// ============================================
// OBTENER HOSTS A MONITOREAR
// ============================================
function obtenerHostsAMonitorear() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  
  if (!configSheet) {
    Logger.log('ERROR: Hoja "Configuración" no existe');
    return [];
  }
  
  const data = configSheet.getDataRange().getValues();
  const hosts = [];
  
  let startRow = -1;
  for (let i = 0; i < data.length; i++) {
    const celda = data[i][0] ? data[i][0].toString().trim() : '';
    
    if (celda === 'Nombre Host') {
      startRow = i + 1;
      break;
    }
  }
  
  if (startRow === -1) {
    Logger.log('ERROR: No se encontró la tabla de hosts');
    return hosts;
  }
  
  for (let i = startRow; i < data.length; i++) {
    const nombre = data[i][0] ? data[i][0].toString().trim() : '';
    const protocolo = data[i][1] ? data[i][1].toString().trim().toLowerCase() : 'http';
    const ip = data[i][2] ? data[i][2].toString().trim() : '';
    const puerto = data[i][3] ? data[i][3].toString().trim() : '';
    const activo = data[i][4] ? data[i][4].toString().trim() : '';
    
    if (nombre && ip) {
      if (!activo) {
        hosts.push({
          nombre: nombre,
          protocolo: protocolo || 'http',
          ip: ip,
          puerto: puerto || ''
        });
      } else {
        const activoUpper = activo.toUpperCase();
        if (activoUpper === 'SI' || activoUpper === 'SÍ' || activoUpper === 'YES' || activoUpper === 'S') {
          hosts.push({
            nombre: nombre,
            protocolo: protocolo || 'http',
            ip: ip,
            puerto: puerto || ''
          });
        }
      }
    }
  }
  
  Logger.log(`Total hosts a monitorear: ${hosts.length}`);
  return hosts;
}

// ============================================
// MONITOREAR UN HOST ESPECÍFICO
// ============================================
function monitorearHost(host, config) {
  try {
    Logger.log(`Monitoreando: ${host.nombre} (${host.protocolo}://${host.ip}${host.puerto ? ':' + host.puerto : ''})`);
    
    const estadoActual = verificarConectividad(host);
    const estadoAnteriorData = obtenerEstadoAnterior(host.nombre);
    
    actualizarEstadoActual(host, estadoActual);
    
    // Solo notificar y registrar si HAY CAMBIO REAL de estado
    if (estadoAnteriorData && estadoAnteriorData.estado !== estadoActual) {
      Logger.log(`CAMBIO DE ESTADO detectado: ${estadoAnteriorData.estado} -> ${estadoActual}`);
      
      registrarCambioEnHistorial(host, estadoAnteriorData, estadoActual);
      enviarNotificaciones(host, estadoActual, estadoAnteriorData.estado, config);
    } else {
      Logger.log(`Sin cambios: ${estadoActual}`);
    }
    
  } catch (error) {
    Logger.log(`Error monitoreando ${host.nombre}: ${error}`);
  }
}

// ============================================
// VERIFICAR CONECTIVIDAD
// ============================================
function verificarConectividad(host) {
  try {
    let url = `${host.protocolo}://${host.ip}`;
    if (host.puerto) {
      url += `:${host.puerto}`;
    }
    
    const options = {
      'muteHttpExceptions': true,
      'validateHttpsCertificates': false,
      'followRedirects': true,
      'timeout': 10
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode >= 200 && responseCode < 600) {
      return 'ACTIVO';
    } else {
      return 'INACTIVO';
    }
    
  } catch (error) {
    return 'INACTIVO';
  }
}

// ============================================
// OBTENER ESTADO ANTERIOR CON TIMESTAMP
// ============================================
function obtenerEstadoAnterior(nombreHost) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let estadoSheet = ss.getSheetByName(SHEET_NAMES.ESTADO);
  
  if (!estadoSheet) {
    estadoSheet = ss.insertSheet(SHEET_NAMES.ESTADO);
    estadoSheet.appendRow(['Fecha Hora', 'Nombre Host', 'Protocolo', 'IP', 'Puerto', 'Estado', 'Desde']);
    estadoSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    return null;
  }
  
  const data = estadoSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === nombreHost) {
      return {
        estado: data[i][5],
        timestamp: data[i][6]
      };
    }
  }
  
  return null;
}

// ============================================
// ACTUALIZAR ESTADO ACTUAL
// ============================================
function actualizarEstadoActual(host, estado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let estadoSheet = ss.getSheetByName(SHEET_NAMES.ESTADO);
  
  if (!estadoSheet) {
    estadoSheet = ss.insertSheet(SHEET_NAMES.ESTADO);
    estadoSheet.appendRow(['Fecha Hora', 'Nombre Host', 'Protocolo', 'IP', 'Puerto', 'Estado', 'Desde']);
    estadoSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  }
  
  const ahora = new Date();
  const fechaHora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  const data = estadoSheet.getDataRange().getValues();
  let filaEncontrada = -1;
  let fechaDesde = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === host.nombre) {
      filaEncontrada = i + 1;
      if (data[i][5] === estado) {
        fechaDesde = data[i][6];
      }
      break;
    }
  }
  
  if (!fechaDesde) {
    fechaDesde = fechaHora;
  }
  
  const color = estado === 'ACTIVO' ? COLORES.ACTIVO : COLORES.INACTIVO;
  const puertoDisplay = host.puerto || '-';
  
  if (filaEncontrada > 0) {
    const rango = estadoSheet.getRange(filaEncontrada, 1, 1, 7);
    rango.setValues([[fechaHora, host.nombre, host.protocolo.toUpperCase(), host.ip, puertoDisplay, estado, fechaDesde]]);
    rango.setBackground(color);
  } else {
    estadoSheet.appendRow([fechaHora, host.nombre, host.protocolo.toUpperCase(), host.ip, puertoDisplay, estado, fechaDesde]);
    const ultimaFila = estadoSheet.getLastRow();
    estadoSheet.getRange(ultimaFila, 1, 1, 7).setBackground(color);
  }
  
  if (filaEncontrada > 0) {
    estadoSheet.getRange(filaEncontrada, 6).setFontWeight('bold');
  } else {
    estadoSheet.getRange(estadoSheet.getLastRow(), 6).setFontWeight('bold');
  }
}

// ============================================
// REGISTRAR CAMBIO REAL EN HISTORIAL
// ============================================
function registrarCambioEnHistorial(host, estadoAnteriorData, estadoNuevo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historialSheet = ss.getSheetByName(SHEET_NAMES.HISTORIAL);
  
  if (!historialSheet) {
    historialSheet = ss.insertSheet(SHEET_NAMES.HISTORIAL);
    historialSheet.appendRow([
      'Fecha Hora Actual',
      'Nombre Host',
      'Protocolo',
      'IP',
      'Puerto',
      'Estado Anterior',
      'Fecha Hora Estado Anterior',
      'Estado Nuevo',
      'Tiempo en Minutos',
      'Tiempo Legible',
      'Observación'
    ]);
    historialSheet.getRange('A1:K1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  }
  
  const ahora = new Date();
  const fechaHoraActual = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  let minutos = 0;
  let tiempoLegible = 'N/A';
  
  if (estadoAnteriorData.timestamp) {
    const fechaAnterior = new Date(estadoAnteriorData.timestamp);
    const diferenciaMs = ahora - fechaAnterior;
    minutos = Math.round(diferenciaMs / 60000);
    tiempoLegible = formatearTiempo(minutos);
  }
  
  const observacion = estadoNuevo === 'ACTIVO' ? 
    `Servicio restablecido (estuvo inactivo ${tiempoLegible})` : 
    `Servicio caído (estuvo activo ${tiempoLegible})`;
  
  const puertoDisplay = host.puerto || '-';
  
  historialSheet.appendRow([
    fechaHoraActual,
    host.nombre,
    host.protocolo.toUpperCase(),
    host.ip,
    puertoDisplay,
    estadoAnteriorData.estado,
    estadoAnteriorData.timestamp || 'N/A',
    estadoNuevo,
    minutos,
    tiempoLegible,
    observacion
  ]);
  
  // Color para cambios reales (más claro)
  const ultimaFila = historialSheet.getLastRow();
  const color = estadoNuevo === 'ACTIVO' ? COLORES.ACTIVO : COLORES.INACTIVO;
  historialSheet.getRange(ultimaFila, 1, 1, 11).setBackground(color);
}

// ============================================
// FORMATEAR TIEMPO
// ============================================
function formatearTiempo(minutos) {
  if (minutos < 1) return '< 1m';
  
  const dias = Math.floor(minutos / 1440);
  const horas = Math.floor((minutos % 1440) / 60);
  const mins = minutos % 60;
  
  let resultado = [];
  if (dias > 0) resultado.push(`${dias}d`);
  if (horas > 0) resultado.push(`${horas}h`);
  if (mins > 0) resultado.push(`${mins}m`);
  
  return resultado.join(' ') || '0m';
}

// ============================================
// ACTUALIZAR REPORTES
// ============================================
function actualizarReportes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let reportesSheet = ss.getSheetByName(SHEET_NAMES.REPORTES);
  
  if (!reportesSheet) {
    reportesSheet = ss.insertSheet(SHEET_NAMES.REPORTES);
  }
  
  reportesSheet.clear();
  
  reportesSheet.appendRow([
    'REPORTE DE DISPONIBILIDAD',
    '',
    '',
    '',
    '',
    'Última actualización: ' + new Date()
  ]);
  reportesSheet.getRange('A1:F1').merge().setFontSize(14).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  reportesSheet.getRange('A2:F2').merge();
  
  reportesSheet.appendRow([
    'Nombre Host',
    'Total Minutos ACTIVO',
    'Total Minutos INACTIVO',
    'Total Minutos Monitoreados',
    '% Disponibilidad',
    'Estado Actual',
    'Uptime'
  ]);
  reportesSheet.getRange('A3:G3').setFontWeight('bold').setBackground('#e8eaf6');
  
  const historialSheet = ss.getSheetByName(SHEET_NAMES.HISTORIAL);
  const estadoSheet = ss.getSheetByName(SHEET_NAMES.ESTADO);
  
  if (!historialSheet || !estadoSheet) {
    reportesSheet.appendRow(['No hay datos de historial disponibles', '', '', '', '', '', '']);
    return;
  }
  
  const historialData = historialSheet.getDataRange().getValues();
  const estadoData = estadoSheet.getDataRange().getValues();
  
  const hosts = {};
  
  for (let i = 1; i < historialData.length; i++) {
    const nombreHost = historialData[i][1];
    const estadoAnterior = historialData[i][5];
    const estadoNuevo = historialData[i][7];
    const minutos = historialData[i][8];
    
    if (!hosts[nombreHost]) {
      hosts[nombreHost] = {
        minutosActivo: 0,
        minutosInactivo: 0
      };
    }
    
    // Si estado anterior = estado nuevo = checkpoint, sumar según el estado
    if (estadoAnterior === estadoNuevo) {
      if (estadoNuevo === 'ACTIVO') {
        hosts[nombreHost].minutosActivo += minutos;
      } else {
        hosts[nombreHost].minutosInactivo += minutos;
      }
    } else {
      // Cambio real: si nuevo es ACTIVO, estuvo INACTIVO
      if (estadoNuevo === 'ACTIVO') {
        hosts[nombreHost].minutosInactivo += minutos;
      } else {
        hosts[nombreHost].minutosActivo += minutos;
      }
    }
  }
  
  for (const nombreHost in hosts) {
    const data = hosts[nombreHost];
    const totalMinutos = data.minutosActivo + data.minutosInactivo;
    const disponibilidad = totalMinutos > 0 ? (data.minutosActivo / totalMinutos * 100).toFixed(2) : 0;
    
    let estadoActual = 'DESCONOCIDO';
    for (let i = 1; i < estadoData.length; i++) {
      if (estadoData[i][1] === nombreHost) {
        estadoActual = estadoData[i][5];
        break;
      }
    }
    
    const uptime = formatearTiempo(data.minutosActivo);
    
    reportesSheet.appendRow([
      nombreHost,
      data.minutosActivo,
      data.minutosInactivo,
      totalMinutos,
      disponibilidad + '%',
      estadoActual,
      uptime
    ]);
    
    const ultimaFila = reportesSheet.getLastRow();
    let color = '#ffffff';
    if (disponibilidad >= 99.9) color = '#d9ead3';
    else if (disponibilidad >= 99) color = '#fff2cc';
    else color = '#f4cccc';
    
    reportesSheet.getRange(ultimaFila, 1, 1, 7).setBackground(color);
  }
  
  reportesSheet.autoResizeColumns(1, 7);
  
  const filaVacia = reportesSheet.getLastRow() + 1;
  reportesSheet.getRange(filaVacia, 1).setValue('');
  reportesSheet.appendRow(['LEYENDA:', '', '', '', '', '', '']);
  reportesSheet.appendRow(['Verde: ≥ 99.9% disponibilidad (SLA Tier 1)', '', '', '', '', '', '']);
  reportesSheet.appendRow(['Amarillo: ≥ 99% disponibilidad (SLA Tier 2)', '', '', '', '', '', '']);
  reportesSheet.appendRow(['Rojo: < 99% disponibilidad (Requiere atención)', '', '', '', '', '', '']);
  
  Logger.log('Reportes actualizados correctamente');
}

// ============================================
// ENVIAR NOTIFICACIONES
// ============================================
function enviarNotificaciones(host, estadoNuevo, estadoAnterior, config) {
  const ahora = new Date();
  const timestamp = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  let emoji, titulo, color;
  if (estadoNuevo === 'ACTIVO') {
    emoji = '✅';
    titulo = 'SERVICIO RESTABLECIDO';
    color = '#28a745';
  } else {
    emoji = '🚨';
    titulo = 'ALERTA - SERVICIO CAÍDO';
    color = '#dc3545';
  }
  
  const urlCompleta = `${host.protocolo}://${host.ip}${host.puerto ? ':' + host.puerto : ''}`;
  
  const mensajeTelegram = `${emoji} <b>${titulo}</b>\n\n` +
                         `🖥 Host: ${host.nombre}\n` +
                         `🌐 URL: ${urlCompleta}\n` +
                         `📡 Protocolo: ${host.protocolo.toUpperCase()}\n` +
                         `🔌 Puerto: ${host.puerto || 'predeterminado'}\n` +
                         `⏰ Hora: ${timestamp}\n` +
                         `📊 Estado Anterior: ${estadoAnterior}\n` +
                         `📊 Estado Actual: ${estadoNuevo}`;
  
  const asunto = `${emoji} ${titulo}: ${host.nombre}`;
  const cuerpoEmail = `
    <html>
    <body style="font-family: Arial, sans-serif;">
      <h2 style="color: ${color};">${emoji} ${titulo}</h2>
      <table style="border-collapse: collapse; width: 100%; max-width: 600px;">
        <tr style="background-color: #f2f2f2;">
          <td style="padding: 10px; border: 1px solid #ddd;"><strong>Host:</strong></td>
          <td style="padding: 10px; border: 1px solid #ddd;">${host.nombre}</td>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #ddd;"><strong>URL:</strong></td>
          <td style="padding: 10px; border: 1px solid #ddd;"><code>${urlCompleta}</code></td>
        </tr>
        <tr style="background-color: #f2f2f2;">
          <td style="padding: 10px; border: 1px solid #ddd;"><strong>Protocolo:</strong></td>
          <td style="padding: 10px; border: 1px solid #ddd;">${host.protocolo.toUpperCase()}</td>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #ddd;"><strong>Puerto:</strong></td>
          <td style="padding: 10px; border: 1px solid #ddd;">${host.puerto || 'predeterminado'}</td>
        </tr>
        <tr style="background-color: #f2f2f2;">
          <td style="padding: 10px; border: 1px solid #ddd;"><strong>Hora:</strong></td>
          <td style="padding: 10px; border: 1px solid #ddd;">${timestamp}</td>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #ddd;"><strong>Estado Anterior:</strong></td>
          <td style="padding: 10px; border: 1px solid #ddd;">${estadoAnterior}</td>
        </tr>
        <tr style="background-color: #f2f2f2;">
          <td style="padding: 10px; border: 1px solid #ddd;"><strong>Estado Actual:</strong></td>
          <td style="padding: 10px; border: 1px solid #ddd;"><strong style="color: ${color};">${estadoNuevo}</strong></td>
        </tr>
      </table>
      <hr>
      <p style="font-size: 12px; color: #666;">
        Monitoreo automático - Google Apps Script v5.0<br>
        Sistema de Monitoreo Institucional
      </p>
    </body>
    </html>
  `;
  
  enviarTelegram(mensajeTelegram, config);
  enviarEmail(asunto, cuerpoEmail, config);
}

// ============================================
// ENVIAR MENSAJE POR TELEGRAM
// ============================================
function enviarTelegram(mensaje, config) {
  try {
    const botToken = config['Bot Token'];
    const chatIds = config['Telegrams'];
    
    if (!botToken || !chatIds) {
      Logger.log('ERROR: Configuración de Telegram incompleta');
      return;
    }
    
    const idsArray = chatIds.split(';').map(id => id.trim()).filter(id => id);
    
    idsArray.forEach((chatId, index) => {
      const url = `https://api.telegram.org/bot${botToken}/sendMessage`;
      const payload = {
        'chat_id': chatId,
        'text': mensaje,
        'parse_mode': 'HTML'
      };
      
      const options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };
      
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode === 200) {
          Logger.log(`  ✓ Telegram enviado a ${chatId}`);
        } else {
          Logger.log(`  ✗ Error al enviar a ${chatId}: ${response.getContentText()}`);
        }
        
      } catch (fetchError) {
        Logger.log(`  ✗ Error de conexión: ${fetchError}`);
      }
    });
    
  } catch (error) {
    Logger.log('ERROR en enviarTelegram: ' + error);
  }
}

// ============================================
// ENVIAR EMAIL
// ============================================
function enviarEmail(asunto, cuerpo, config) {
  try {
    const emails = config['Emails'];
    
    if (!emails) {
      Logger.log('No hay emails configurados');
      return;
    }
    
    const emailsArray = emails.split(';').map(email => email.trim()).filter(email => email);
    
    emailsArray.forEach(email => {
      MailApp.sendEmail({
        to: email,
        subject: asunto,
        htmlBody: cuerpo
      });
      Logger.log(`Email enviado a: ${email}`);
    });
    
  } catch (error) {
    Logger.log('Error enviando email: ' + error);
  }
}

// ============================================
// ENVIAR ALERTA DE ERROR EN EL SCRIPT
// ============================================
function enviarAlertaError(error) {
  try {
    const config = obtenerConfiguracion();
    const emails = config['Emails'];
    
    if (emails) {
      const emailsArray = emails.split(';').map(e => e.trim()).filter(e => e);
      const asunto = '⚠️ Error en Script de Monitoreo';
      const cuerpo = `
        <html>
        <body>
          <h2>Error en el Script de Monitoreo</h2>
          <p><strong>Error:</strong> ${error.toString()}</p>
          <p><strong>Hora:</strong> ${new Date()}</p>
          <p>Por favor, revise el script en Google Apps Script.</p>
        </body>
        </html>
      `;
      
      emailsArray.forEach(email => {
        MailApp.sendEmail({
          to: email,
          subject: asunto,
          htmlBody: cuerpo
        });
      });
    }
  } catch (e) {
    Logger.log('Error enviando alerta de error: ' + e);
  }
}

// ============================================
// CREAR TRIGGER AUTOMÁTICO
// ============================================
function crearTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'monitorearRed') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('monitorearRed')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  Logger.log('Trigger creado: ejecución cada 10 minutos');
}

// ============================================
// ELIMINAR TRIGGER
// ============================================
function eliminarTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'monitorearRed') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Trigger eliminado');
    }
  });
}

// ============================================
// FUNCIÓN DE PRUEBA
// ============================================
function probarTelegram() {
  Logger.log('=== PRUEBA DE TELEGRAM ===\n');
  const config = obtenerConfiguracion();
  const mensajePrueba = '🧪 <b>PRUEBA DE TELEGRAM</b>\n\n' +
                       '✅ Si recibes este mensaje, Telegram está configurado correctamente.\n\n' +
                       '⏰ Hora: ' + new Date().toLocaleString('es-CO', { timeZone: 'America/Bogota' });
  
  enviarTelegram(mensajePrueba, config);
  Logger.log('\n=== FIN DE PRUEBA ===');
}
