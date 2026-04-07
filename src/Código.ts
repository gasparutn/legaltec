/**
 * Code.gs - Lógica de backend para la gestión de notificaciones
 * Este script maneja la lectura de datos de Google Sheet y el envío de correos.
 */

// Constantes de configuración global se inicializan de forma perezosa (Lazy Init)
// para evitar errores de orden de carga de archivos en Google Apps Script.
let __ss: GoogleAppsScript.Spreadsheet.Spreadsheet;
function getSS() {
  if (!__ss) __ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return __ss;
}

function getMainSheet() { return getSS().getSheetByName("Registros"); }
function getBajaResSheet() { return getSS().getSheetByName("BAJA RES"); }
function getControlSheet() { return getSS().getSheetByName("Control"); }

let CONFIG_CACHE: any = null; // Cache para la configuración de autoridades

// normalizeHeader se encuentra ahora en utils.ts

// parseSafeDate se encuentra ahora en utils.ts

/**
 * Lee la hoja "aux" y devuelve mapeo de autoridad -> {senadora, secretaria, correo}
 * Lee el rango A1:D44 (encabezados en fila 1, datos en filas 2-44)
 * Columnas esperadas: A=Senador/a (apellido), B=Senador/a (nombre+apellido), C=Secretario/a, D=Correo
 */
function loadAutoridadesConfig() {
  try {
    const auxSheet = getSS().getSheetByName("aux");
    if (!auxSheet) {
      Logger.log("Aviso: Hoja 'aux' no encontrada. Usando configuración vacía.");
      return {};
    }

    const range = auxSheet.getRange("A1:D44");
    const values = range.getValues();
    const config = {};

    // Saltamos la fila 1 (header - fila índice 0)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const apellido = String(row[0] || '').trim().toUpperCase();
      const nombreCompleto = String(row[1] || '').trim();
      const secretario = String(row[2] || '').trim();
      const correo = String(row[3] || '').trim();

      if (apellido && secretario && correo) {
        config[apellido] = {
          apellido: apellido,
          nombreCompleto: nombreCompleto,
          secretario: secretario,
          correo: correo
        };
      }
    }

    Logger.log("Config cargada: " + Object.keys(config).length + " autoridades encontradas");
    return config;
  } catch (e) {
    Logger.log("Error cargando config de autoridades: " + e.message);
    return {};
  }
}

/**
 * Obtiene el secretario y correo para una autoridad dada
 */
function getSecretarioYCorreo(autoridad) {
  if (!CONFIG_CACHE) {
    CONFIG_CACHE = loadAutoridadesConfig();
  }
  return CONFIG_CACHE[String(autoridad || '').trim().toUpperCase()] || null;
}

/**
 * Obtiene la lista completa de autoridades desde la hoja config
 */
function getAutoridadesList(ts?: any) {
  return loadAutoridadesConfig();
}

/**
 * Agrega una nueva autoridad a la hoja aux
 */
function addAutoridadConfig(apellido: string, nombreCompleto: string, secretario: string, correo: string) {
  try {
    const apellidoLower = String(apellido || '').trim().toUpperCase();
    if (!apellidoLower || !secretario || !correo) {
      return { success: false, message: "Faltan datos requeridos: Apellido, Secretario y Correo son obligatorios." };
    }

    const auxSheet = getSS().getSheetByName("aux");
    if (!auxSheet) {
      return { success: false, message: "Hoja 'aux' no encontrada." };
    }

    // Verificar que no exista ya
    const config = loadAutoridadesConfig();
    if (config[apellidoLower]) {
      return { success: false, message: "La autoridad '" + apellidoLower + "' ya existe." };
    }

    // Encontrar la primera fila vacía en el rango A2:D44
    const range = auxSheet.getRange("A2:D44");
    const values = range.getValues();
    let insertRow = -1;

    for (let i = 0; i < values.length; i++) {
      if (!values[i][0] || String(values[i][0]).trim() === '') {
        insertRow = i + 2; // +2 porque range empieza en fila 2 (i=0 es fila 2)
        break;
      }
    }

    if (insertRow === -1) {
      return { success: false, message: "No hay espacio disponible en la hoja aux (máximo 43 autoridades)." };
    }

    auxSheet.getRange(insertRow, 1).setValue(apellidoLower); // Columna A
    auxSheet.getRange(insertRow, 2).setValue(nombreCompleto); // Columna B
    auxSheet.getRange(insertRow, 3).setValue(secretario); // Columna C
    auxSheet.getRange(insertRow, 4).setValue(correo); // Columna D

    SpreadsheetApp.flush(); // Forzar la escritura

    // Limpiar cache para recargar
    CONFIG_CACHE = null;
    Logger.log("Autoridad agregada: " + apellidoLower);
    return { success: true, message: "Autoridad '" + apellidoLower + "' agregada exitosamente." };
  } catch (e) {
    Logger.log("Error agregando autoridad: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Edita una autoridad existente en la hoja aux
 */
function updateAutoridadConfig(oldApellido: string, newApellido: string, secretario: string, correo: string) {
  try {
    const auxSheet = getSS().getSheetByName("aux");
    if (!auxSheet) {
      return { success: false, message: "Hoja 'aux' no encontrada." };
    }

    const oldApellidoUpper = String(oldApellido || '').trim().toUpperCase();
    const newApellidoUpper = String(newApellido || '').trim().toUpperCase();
    
    if (!oldApellidoUpper || !newApellidoUpper || !secretario || !correo) {
      return { success: false, message: "Faltan datos requeridos." };
    }

    // Buscar la fila con esta autoridad (usando el apellido antiguo)
    const range = auxSheet.getRange("A1:D44");
    const values = range.getValues();
    let foundRow = -1;

    for (let i = 1; i < values.length; i++) { // i=1 para saltar header
      if (String(values[i][0] || '').trim().toUpperCase() === oldApellidoUpper) {
        foundRow = i + 1; // +1 porque range empieza en fila 1
        break;
      }
    }

    // Si la autoridad no está en el sheet (es predefinida), agregarla como custom
    if (foundRow === -1) {
      // Buscar primera fila vacía en A2:A44
      for (let i = 1; i < values.length; i++) {
        if (!values[i][0] || String(values[i][0]).trim() === '') {
          const newRow = i + 1; // fila real del sheet
          auxSheet.getRange(newRow, 1).setValue(newApellidoUpper); // A
          auxSheet.getRange(newRow, 2).setValue(""); // B (nombreCompleto vacío)
          auxSheet.getRange(newRow, 3).setValue(secretario); // C
          auxSheet.getRange(newRow, 4).setValue(correo); // D
          
          SpreadsheetApp.flush(); // Forzar la escritura antes de que la UI recargue los valores
          
          CONFIG_CACHE = null;
          Logger.log("Autoridad predefinida '"+oldApellidoUpper+"' agregada como custom: " + newApellidoUpper);
          return { success: true, message: "Autoridad agregada exitosamente." };
        }
      }
      return { success: false, message: "Límite de autoridades alcanzado (máx. 44)." };
    }

    // Si existe en el sheet, actualizarla
    auxSheet.getRange(foundRow, 1).setValue(newApellidoUpper); // Columna A - Apellido
    auxSheet.getRange(foundRow, 3).setValue(secretario); // Columna C
    auxSheet.getRange(foundRow, 4).setValue(correo); // Columna D

    SpreadsheetApp.flush(); // Forzar la escritura antes de que la UI recargue los valores

    // Limpiar cache
    CONFIG_CACHE = null;
    Logger.log("Autoridad actualizada: " + newApellidoUpper);
    return { success: true, message: "Autoridad '" + newApellidoUpper + "' actualizada exitosamente." };
  } catch (e) {
    Logger.log("Error actualizando autoridad: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Elimina una autoridad de la hoja aux
 */
function deleteAutoridadConfig(apellido: string) {
  try {

    const auxSheet = getSS().getSheetByName("aux");
    if (!auxSheet) {
      return { success: false, message: "Hoja 'aux' no encontrada." };
    }

    const apellidoUpper = String(apellido || '').trim().toUpperCase();
    if (!apellidoUpper) {
      return { success: false, message: "Apellido requerido." };
    }

    // Buscar y limpiar la fila
    const range = auxSheet.getRange("A1:D44");
    const values = range.getValues();
    let foundRow = -1;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0] || '').trim().toUpperCase() === apellidoUpper) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      // Si no existe en sheet, es una autoridad predefinida. Retornar suceso (desaparecerá del SELECT al recargar)
      Logger.log("Autoridad predefinida '" + apellidoUpper + "' no está en sheet (no se elimina).");
      return { success: true, message: "Autoridad eliminada (lista refresca automáticamente)." };
    }

    // Limpiar las 4 celdas de la fila
    auxSheet.getRange(foundRow, 1, 1, 4).clearContent();

    SpreadsheetApp.flush(); // Forzar la escritura

    // Limpiar cache
    CONFIG_CACHE = null;
    Logger.log("Autoridad eliminada: " + apellidoUpper);
    return { success: true, message: "Autoridad '" + apellidoUpper + "' eliminada exitosamente." };
  } catch (e) {
    Logger.log("Error eliminando autoridad: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Lee la hoja "config" y devuelve array de tareas customizadas
 * Lee el rango L1:L100 (tareas adicionales configuradas)
 */
/**
 * Obtiene lista de tareas desde la hoja "aux" rango F2:F100
 */
function getTareasList(ts?: any) {
  try {
    const auxSheet = getSS().getSheetByName("aux");
    if (!auxSheet) {
      Logger.log("Aviso: Hoja 'aux' no encontrada. Devolviendo lista vacía.");
      return [];
    }

    const range = auxSheet.getRange("F2:F100");
    const values = range.getValues();
    const tareas: string[] = [];

    for (let i = 0; i < values.length; i++) {
      const tarea = String(values[i][0] || '').trim();
      if (tarea) {
        tareas.push(tarea);
      }
    }

    Logger.log("Tareas cargadas desde aux F2:F100: " + tareas.length);
    return tareas;
  } catch (e) {
    Logger.log("Error obteniendo lista de tareas: " + e.message);
    return [];
  }
}

/**
 * Agrega nueva tarea en la hoja "aux" F2:F100
 */
function addTareaConfig(tarea: string) {
  try {
    const tareaLimpia = String(tarea || '').trim();
    if (!tareaLimpia) {
      return { success: false, message: "Tarea vacía." };
    }

    const auxSheet = getSS().getSheetByName("aux");
    if (!auxSheet) {
      return { success: false, message: "Hoja 'aux' no encontrada." };
    }

    // Encontrar primera celda vacía en F2:F100
    const range = auxSheet.getRange("F2:F100");
    const values = range.getValues();
    let firstEmpty = -1;

    for (let i = 0; i < values.length; i++) {
      if (!values[i][0] || String(values[i][0]).trim() === '') {
        firstEmpty = i + 2; // +2 porque range empieza en F2
        break;
      }
    }

    if (firstEmpty === -1) {
      return { success: false, message: "Límite de tareas alcanzado (máx. 100)." };
    }

    auxSheet.getRange('F' + firstEmpty).setValue(tareaLimpia);
    SpreadsheetApp.flush();
    
    Logger.log("Tarea agregada: " + tareaLimpia);
    return { success: true, message: "Tarea '" + tareaLimpia + "' agregada exitosamente." };
  } catch (e) {
    Logger.log("Error agregando tarea: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Actualiza una tarea por nombre en la hoja "aux" F2:F100
 */
function updateTareaConfig(oldName: string, newName: string) {
  try {
    const oldNameLimpia = String(oldName || '').trim().toUpperCase();
    const newNameLimpia = String(newName || '').trim();
    
    if (!newNameLimpia) {
      return { success: false, message: "Tarea vacía." };
    }

    const auxSheet = getSS().getSheetByName("aux");
    if (!auxSheet) {
      return { success: false, message: "Hoja 'aux' no encontrada." };
    }

    const range = auxSheet.getRange("F2:F100");
    const values = range.getValues();

    Logger.log("updateTareaConfig: Buscando '" + oldNameLimpia + "' en " + values.length + " filas");

    // Buscar la tarea por nombre (case-insensitive)
    let foundIdx = -1;
    const tareasTrovadas: string[] = [];
    for (let i = 0; i < values.length; i++) {
      const cellValue = String(values[i][0] || '').trim().toUpperCase();
      if (cellValue) {
        tareasTrovadas.push(cellValue);
      }
      if (cellValue === oldNameLimpia) {
        foundIdx = i + 2; // +2 porque range empieza en F2
        Logger.log("updateTareaConfig: ¡Encontrada en fila " + foundIdx + "!");
        break;
      }
    }
    Logger.log("updateTareaConfig: Tareas encontradas en la hoja: " + tareasTrovadas.join(" | "));

    // Si no existe, agregarla como nueva
    if (foundIdx === -1) {
      Logger.log("updateTareaConfig: NO ENCONTRADA. Buscando primera fila vacía para agregar como nueva...");
      for (let i = 0; i < values.length; i++) {
        if (!values[i][0] || String(values[i][0]).trim() === '') {
          const newRow = i + 2;
          auxSheet.getRange('F' + newRow).setValue(newNameLimpia);
          SpreadsheetApp.flush();
          Logger.log("Tarea '" + oldNameLimpia + "' agregada como nueva en fila " + newRow + ": " + newNameLimpia);
          return { success: true, message: "Tarea agregada exitosamente." };
        }
      }
      return { success: false, message: "Límite de tareas alcanzado (máx. 100)." };
    }

    // Si existe, actualizarla
    auxSheet.getRange('F' + foundIdx).setValue(newNameLimpia);
    SpreadsheetApp.flush();
    
    Logger.log("Tarea actualizada en fila " + foundIdx + ": " + oldNameLimpia + " -> " + newNameLimpia);
    return { success: true, message: "Tarea actualizada exitosamente." };
  } catch (e) {
    Logger.log("Error actualizando tarea: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Elimina una tarea por nombre en la hoja "aux" F2:F100
 */
function deleteTareaConfig(tareaName: string) {
  try {
    const tareaNameLimpia = String(tareaName || '').trim().toUpperCase();
    
    const auxSheet = getSS().getSheetByName("aux");
    if (!auxSheet) {
      return { success: false, message: "Hoja 'aux' no encontrada." };
    }

    const range = auxSheet.getRange("F2:F100");
    const values = range.getValues();

    // Buscar la tarea (case-insensitive)
    let foundIdx = -1;
    for (let i = 0; i < values.length; i++) {
      const cellValue = String(values[i][0] || '').trim().toUpperCase();
      if (cellValue === tareaNameLimpia) {
        foundIdx = i + 2; // +2 porque range empieza en F2
        break;
      }
    }

    if (foundIdx === -1) {
      Logger.log("Tarea '" + tareaNameLimpia + "' no encontrada en aux.");
      return { success: true, message: "Tarea eliminada (lista refresca automáticamente)." };
    }

    auxSheet.getRange('F' + foundIdx).clearContent();
    SpreadsheetApp.flush();
    
    Logger.log("Tarea eliminada: " + tareaNameLimpia);
    return { success: true, message: "Tarea eliminada exitosamente." };
  } catch (e) {
    Logger.log("Error eliminando tarea: " + e.message);
    return { success: false, message: "Error: " + e.message };
  }
}

/**
 * Obtiene el rol del usuario actual desde la hoja "config" (A1:E43)
 * Columnas: A=ID, B=NOMBRE, C=USUARIO, D=PASS, E=ROL
 */
function getCurrentUserRole() {
  try {
    const props = PropertiesService.getUserProperties();
    const userName = props.getProperty('CD_NOMBRE') || '';
    if (!userName) {
      Logger.log("Aviso: No se encontró usuario en sesión");
      return null;
    }

    const configSheet = getSS().getSheetByName("config");
    if (!configSheet) {
      Logger.log("Error: Hoja 'config' no encontrada");
      return null;
    }

    const range = configSheet.getRange("A1:E43");
    const values = range.getValues();

    // Buscar el usuario en la lista (columna B = NOMBRE)
    for (let i = 1; i < values.length; i++) { // i=1 para saltar header
      const nombre = String(values[i][1] || '').trim();
      if (nombre.toUpperCase() === userName.toUpperCase()) {
        return String(values[i][4] || '').trim().toUpperCase(); // Columna E = ROL
      }
    }

    Logger.log("Usuario no encontrado en config: " + userName);
    return null;
  } catch (e) {
    Logger.log("Error obteniendo rol del usuario: " + e.message);
    return null;
  }
}

/**
 * Comprueba si el usuario actual tiene permiso para gestionar autoridades
 * Solo "admin" y "supervisor" pueden gestionarlas
 */
function canManageAutoridadesByCurrentUser() {
  const rol = getCurrentUserRole();
  return rol === 'ADMIN' || rol === 'SUPERVISOR';
}

/**
 * Obtiene el rol del usuario actual (para enviar al frontend)
 */
function getUserRoleForFrontend() {
  return { role: getCurrentUserRole() };
}

/**
 * Registra el acceso al panel de Control Docus.
 */
function logControlDocusAccess() {
  logActivity("Ingresa a Control Docus");
  return { success: true };
}

function doGet(e) {
  const page = e.parameter.page || '';
  const appUrl = ScriptApp.getService().getUrl();

  // Si no hay página, mandamos al Login por defecto
  if (!page) {
    const t = HtmlService.createTemplateFromFile('LoginCD');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('Control Docus — Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── Módulo Control Docus ──────────────────────────────────────────────────
  if (page === 'ControlDocus') {
    const t = HtmlService.createTemplateFromFile('LoginCD');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('Control Docus — Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'PortalUsuario') {
    const t = HtmlService.createTemplateFromFile('PortalUsuario');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('Control Docus — Mi Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'PanelSupervisor') {
    const t = HtmlService.createTemplateFromFile('PanelSupervisor');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('Control Docus — Supervisor')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── Páginas existentes ────────────────────────────────────────────────────
  if (page === 'Notifications') {
    return HtmlService.createTemplateFromFile('Notifications')
      .evaluate()
      .setTitle('Gestión de Notificaciones')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'Index') {
    const t = HtmlService.createTemplateFromFile('Index');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('PLANILLA GENERAL')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Fallback al login para cualquier otra ruta no reconocida o protegida
  const t = HtmlService.createTemplateFromFile('LoginCD');
  t.appUrl = appUrl;
  return t.evaluate().setTitle('Control Docus — Login');
}

function include(filename) {
  // Función auxiliar para incluir archivos CSS o JS en la plantilla HTML.
  // Asegura que el nombre de archivo pasado sea el que se utiliza.
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Devuelve la URL para la página de notificaciones.
 */
function getNotificationsPageUrl() {
  const url = ScriptApp.getService().getUrl();
  return `${url}?page=Notifications`;
}

/**
 * Devuelve la URL para la página principal.
 */
function getMainPageUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Obtiene todos los datos de las notificaciones de la hoja de cálculo
 * para la vista de Notificaciones.
 * @return {Object} Un objeto con 'headers' y 'data'.
 */
function getNotificationsData() {
  const sheet = getMainSheet(); // Usar la hoja designada para notificaciones

  if (!sheet) {
    Logger.log("Hoja de notificaciones no encontrada: Hoja 1");
    return { headers: [], data: [] };
  }

  const range = sheet.getDataRange();
  const allValues = range.getValues();

  if (allValues.length === 0) {
    return { headers: [], data: [] };
  }

  const allHeaders = allValues[0].map(header => String(header).trim());
  const rawData = allValues.slice(1);

  // Definir los encabezados que queremos mostrar en la tabla de notificaciones
  // e en qué orden. Asegúrate de que estos nombres coincidan con los de tu hoja.
  const visibleNotificationHeaders = [
    'Id', 'Título', 'Mensaje', 'Fecha', 'Estado', 'Correo Secretario', 'Acción Notificar'
  ];

  const notificationsData = rawData.map((row, rowIndex) => {
    const rowObject = {};
    allHeaders.forEach((header, colIndex) => {
      rowObject[header] = row[colIndex];
    });

    // Mapear los datos de la fila a las columnas deseadas para la vista de notificaciones
    const notificationRow: any[] = [];
    notificationRow.push(rowObject['Id'] || ''); // Asumiendo 'Id' es una columna
    notificationRow.push(rowObject['REQ'] || ''); // Usamos REQ como 'Título' de la notificación
    notificationRow.push(rowObject['TAREAS'] || ''); // Usamos TAREAS como 'Mensaje' de la notificación
    notificationRow.push(rowObject['FECHA ALTA'] || ''); // Usamos FECHA ALTA como 'Fecha'
    notificationRow.push(rowObject['ESTADO'] || ''); // Usamos ESTADO como 'Estado'
    notificationRow.push(rowObject['CORREO SECRETARIO'] || ''); // Correo del secretario
    notificationRow.push(''); // Espacio para el botón de acción

    return notificationRow;
  });

  return { headers: visibleNotificationHeaders, data: notificationsData };
}

/**
 * Lee la hoja 'aux' y devuelve un mapa de AUTORIDAD → { secretario, correoSecretario }.
 * Columnas: B=AUTORIDAD, C=SECRETARIO, D=CORREO SECRETARIO (filas 2 a 43).
 */
function getAuxData() {
  const cacheKey = "AUX_DATA_MAP_" + SPREADSHEET_ID;
  const cached = getCache(cacheKey);
  if (cached) return { success: true, data: cached };

  try {
    const auxSheet = getSS().getSheetByName('aux');
    if (!auxSheet) {
      Logger.log("Hoja 'aux' no encontrada.");
      return { success: false, message: "Hoja 'aux' no encontrada.", data: {} };
    }

    const lastRow = auxSheet.getLastRow();
    if (lastRow < 2) return { success: true, data: {} };

    // Leer A2:D (columnas A, B, C, D — índices 0, 1, 2, 3 base-0)
    const range = auxSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const auxMap = {};

    range.forEach(row => {
      const autoridad = String(row[0] || '').trim();
      const secretario = String(row[2] || '').trim();
      const correo = String(row[3] || '').trim();
      if (autoridad) {
        auxMap[autoridad] = { secretario, correoSecretario: correo };
      }
    });

    setCache(cacheKey, auxMap, 1800); // Cache por 30 minutos
    return { success: true, data: auxMap };
  } catch (e) {
    Logger.log("Error en getAuxData: " + e.message);
    return { success: false, message: e.message, data: {} };
  }
}

interface FilterOptions {
  type: 'recent' | 'period' | 'all';
  month?: number | string;
  year?: number | string;
}

function getData(filterOptions: FilterOptions = { type: 'recent' }) {
  try {
    checkAndAssignIds(); // Asegurar IDs correlativos para cargas manuales
    const sheet = getMainSheet();

    if (!sheet) {
      console.error("Hoja principal no encontrada: Hoja 1");
      return { success: false, message: "Hoja de datos 'Hoja 1' no encontrada en el documento." };
    }

    const range = sheet.getDataRange();
    const allValues = range.getValues();

    if (!allValues || allValues.length === 0) {
      return { success: true, headers: [], data: [] };
    }

    const allHeaders = allValues[0].map(normalizeHeader);
    const rawData = allValues.slice(1);

    // Identificar índices para filtrado (Flexible: ALTA o DE ALTA)
    let altaIdx = allHeaders.indexOf('FECHA ALTA');
    if (altaIdx === -1) altaIdx = allHeaders.indexOf('FECHA DE ALTA');

    // Configuración de fechas para 'recent'
    const hoy = new Date();
    const mesActual = hoy.getMonth();
    const anioActual = hoy.getFullYear();
    const mesAnterior = mesActual === 0 ? 11 : mesActual - 1;
    const anioAnterior = mesActual === 0 ? anioActual - 1 : anioActual;

    // Filtrado y procesamiento en una sola pasada (optimización de CPU)
    const autoridadesConfig = loadAutoridadesConfig();
    const filteredAndProcessed: any[] = [];
    for (let i = 0; i < rawData.length; i++) {
       const row = rawData[i];
       const cellVal = row[altaIdx];
       let dateObj: any = null;

       // Lógica de filtrado
       if (filterOptions.type !== 'all' && altaIdx !== -1) {
         dateObj = parseSafeDate(cellVal);
         const isDateValid = dateObj && !isNaN(dateObj.getTime());

         if (filterOptions.type === 'recent') {
            if (isDateValid) {
              const m = dateObj.getMonth();
              const y = dateObj.getFullYear();
              if (!((y === anioActual && m === mesActual) || (y === anioAnterior && m === mesAnterior))) continue;
            }
         } else if (filterOptions.type === 'period') {
            if (!isDateValid) continue;
            const targetY = parseInt(String(filterOptions.year || '0'), 10);
            if (dateObj.getFullYear() !== targetY) continue;
            if (filterOptions.month !== 'all_months') {
              const targetM = parseInt(String(filterOptions.month || '0'), 10);
              if (dateObj.getMonth() !== targetM) continue;
            }
         }
       }

       // Si pasó el filtro, procesamos la fila para el objeto de respuesta
       const rowObject = {} as any;
       allHeaders.forEach((header, index) => {
         if (header) {
           rowObject[header] = row[index];
           if (header === 'FECHA DE ALTA') rowObject['FECHA ALTA'] = row[index];
           if (header === 'FECHA DE BAJA') rowObject['FECHA BAJA'] = row[index];
         }
       });

       try {
         const aux = calculateAuxColumns(rowObject);
         Object.assign(rowObject, aux);
       } catch (e) { /* ignore aux errors */ }

       filteredAndProcessed.push(rowObject);
    }

    // Encabezados para el cliente
    const visibleClientHeadersNames = ['ID', 'AUTORIDAD', 'APELLIDOS', 'NOMBRES', 'DNI', 'TAREAS', 'FECHA ALTA', 'TOTAL', 'REQ', 'CHECK', 'ESTADO'];
    const hiddenClientHeadersNames = ['SECRETARIO', 'CORREO SECRETARIO', 'CORREO LOCADOR', 'SEXO', 'FECHA DE BAJA', 'CUOTA', 'DOMICILIO', 'LOCALIDAD', 'GENERA CONTRATO', 'GENERA RESOLUCION', 'FECHA NOTIFICA', 'BAJA CONTRATO', 'AUTOR'];

    // Transformación final, sanitización y ordenamiento
    const finalProcessed = filteredAndProcessed.map(rowObject => {
      if (rowObject['BAJA CONTRATO'] && String(rowObject['BAJA CONTRATO']).trim() !== '') {
        rowObject['ESTADO'] = 'Baja';
      }

      const sanitize = (val) => {
        if (val === null || val === undefined) return '';
        if (val instanceof Date) {
          try {
            return `${val.getFullYear()}-${String(val.getMonth() + 1).padStart(2, '0')}-${String(val.getDate()).padStart(2, '0')}`;
          } catch (e) { return String(val); }
        }
        return val;
      };

      const visible = visibleClientHeadersNames.map(h => sanitize(rowObject[h]));
      const hidden = hiddenClientHeadersNames.map(h => sanitize(rowObject[h]));
      const recordId = rowObject['ID'];
      const sortDate = parseSafeDate(rowObject['FECHA ALTA']).getTime() || 0;

      return { visible, hidden, recordId, sortDate };
    });

    // Ordenar por Fecha Alta (Desc) y luego ID (Desc)
    finalProcessed.sort((a, b) => (b.sortDate - a.sortDate) || (Number(b.recordId) - Number(a.recordId)));

    const finalDataArray = finalProcessed.map(item => [...item.visible, '', ...item.hidden, item.recordId]);

    return {
      success: true,
      headers: [...visibleClientHeadersNames, 'OPC'],
      hiddenHeaders: hiddenClientHeadersNames,
      data: finalDataArray,
      autoridadesConfig: getAuxData().data
    };
  } catch (err) {
    console.error("Error crítico en getData:", err.toString());
    return {
      success: false,
      message: "Fallo crítico al recuperar datos: " + err.toString()
    };
  }
}

/**
 * Función evaluadora y comparadora: obtiene los datos más recientes para una lista de IDs.
 */
function getLatestDataByIDs(ids) {
  try {
    const sheet = getMainSheet();
    if (!sheet) return { success: false, message: "Hoja no encontrada" };
    
    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 2) return { success: true, data: [] };
    
    const allHeaders = allValues[0].map(normalizeHeader);
    const idIdx = allHeaders.indexOf('ID');
    if (idIdx === -1) return { success: false, message: "Columna ID no encontrada en el Sheet" };
    
    const visibleNames = ['ID', 'AUTORIDAD', 'APELLIDOS', 'NOMBRES', 'DNI', 'TAREAS', 'FECHA ALTA', 'TOTAL', 'REQ', 'CHECK', 'ESTADO'];
    const hiddenNames = ['SECRETARIO', 'CORREO SECRETARIO', 'CORREO LOCADOR', 'SEXO', 'FECHA BAJA', 'CUOTA', 'DOMICILIO', 'LOCALIDAD', 'GENERA CONTRATO', 'GENERA RESOLUCION', 'FECHA NOTIFICA', 'BAJA CONTRATO', 'AUTOR'];
    
    // Convertir IDs buscados a strings para comparación segura
    const idsToFind = (Array.isArray(ids) ? ids : [ids]).map(id => String(id).trim());
    
    const matchedRows = allValues.slice(1).filter(row => {
      const rowId = String(row[idIdx]).trim();
      return idsToFind.includes(rowId);
    });
    
    const processed = matchedRows.map(row => {
      const rowObj = {};
      allHeaders.forEach((h, i) => { 
        if (h) {
          rowObj[h] = row[i];
          // Alias dinámicos
          if (h === 'FECHA DE ALTA') rowObj['FECHA ALTA'] = row[i];
          if (h === 'FECHA DE BAJA') rowObj['FECHA BAJA'] = row[i];
        }
      });
      
      // Sanitización igual a getData
      const sanitize = (val) => {
        if (val === null || val === undefined) return '';
        if (val instanceof Date) {
          const y = val.getFullYear();
          const m = String(val.getMonth() + 1).padStart(2, '0');
          const d = String(val.getDate()).padStart(2, '0');
          return `${y}-${m}-${d}`;
        }
        return val;
      };
      
      const vPart = visibleNames.map(h => sanitize(rowObj[h] !== undefined ? rowObj[h] : ''));
      const hPart = hiddenNames.map(h => sanitize(rowObj[h] !== undefined ? rowObj[h] : ''));
      const recordId = String(rowObj['ID'] || '');
      
      // Estructura: [visibles..., '', ocultos..., ID]
      return [...vPart, '', ...hPart, recordId];
    });
    
    return { success: true, data: processed };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function addRow(rowData) {
  const sheet = getMainSheet();

  if (!sheet) {
    return { success: false, message: "Hoja principal no encontrada." };
  }

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idColumnIndex = headers.indexOf('Id'); // Obtener el índice de la columna 'Id'

    if (idColumnIndex === -1) {
      return { success: false, message: "Columna 'Id' no encontrada en la hoja." };
    }

    let nextId = 1;
    if (sheet.getLastRow() > 1) { // Si hay más de una fila (es decir, ya hay datos además de los encabezados)
      const lastId = sheet.getRange(sheet.getLastRow(), idColumnIndex + 1).getValue();
      if (typeof lastId === 'number') {
        nextId = lastId + 1;
      }
    }

    // Limpiar espacios vacíos y formatear DNI y Fechas
    let cleanData = trimData(rowData);
    if (cleanData['DNI']) {
      cleanData['DNI'] = formatearDNI(cleanData['DNI']);
    }

    // Formatear fechas si existen (FECHA ALTA, FECHA BAJA)
    if (cleanData['FECHA ALTA']) {
      cleanData['FECHA ALTA'] = fechaEnEspaniol(cleanData['FECHA ALTA']);
    }
    if (cleanData['FECHA BAJA']) {
      cleanData['FECHA BAJA'] = fechaEnEspaniol(cleanData['FECHA BAJA']);
    }

    // Asignar el nuevo ID y forzar CHECK=true en toda fila nueva
    cleanData['Id'] = nextId; // Usar la capitalización exacta 'Id'
    cleanData['CHECK'] = true; // Al agregar siempre se registra con CHECK tildado

    const newRowValues: any[] = [];
    headers.forEach(header => {
      // Usar cleanData[header] para obtener el valor, incluyendo el nuevo Id
      newRowValues.push(cleanData[header] !== undefined ? cleanData[header] : '');
    });

    sheet.appendRow(newRowValues);

    // Registrar actividad de ingreso
    logActivity(`Ingreso ${cleanData['APELLIDOS']}_${cleanData['AUTORIDAD']}`);

    return { success: true, message: "Registro agregado exitosamente con Id: " + nextId, newId: nextId };
  }
  catch (e) {
    Logger.log("Error al agregar fila: " + e.message);
    return { success: false, message: "Error al agregar registro: " + e.message };
  }
}

function uploadNotaPdf(base64Data: string, apellidos: string, autoridad: string, recordId: number) {
  try {
    // Usar el ID del registro proporcionado
    const fileNumber = String(recordId).padStart(3, '0');
    const fileName = `${fileNumber}) AUTORIZA ${apellidos}-${autoridad}.pdf`;

    // Convertir base64 a Blob
    const byteCharacters = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(byteCharacters, 'application/pdf', fileName);

    // Obtener la carpeta de notas y subir el archivo
    const notasFolder = DriveApp.getFolderById(CARPETA_NOTAS);
    const uploadedFile = notasFolder.createFile(blob);

    Logger.log("Archivo PDF subido: " + fileName);
    logActivity(`Sube Nota PDF ${apellidos}_${autoridad}`);

    return { 
      success: true, 
      message: "Nota PDF subida exitosamente: " + fileName,
      fileId: uploadedFile.getId()
    };
  } catch (e) {
    Logger.log("Error al subir nota PDF: " + e.message);
    return { success: false, message: "Error al subir nota PDF: " + e.message };
  }
}

function updateRow(rowIndex, rowData) {
  const sheet = getMainSheet();

  if (!sheet) {
    return { success: false, message: "Hoja principal no encontrada." };
  }

  const sheetRow = findRowById(sheet, rowIndex);
  if (sheetRow === -1) {
    return { success: false, message: "Registro con ID " + rowIndex + " no encontrado para actualizar." };
  }

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const updatedValues: any[] = [];

    let cleanData = trimData(rowData);
    if (cleanData['DNI']) {
      cleanData['DNI'] = formatearDNI(cleanData['DNI']);
    }

    // Formatear fechas si vienen en el update
    if (cleanData['FECHA ALTA']) {
      cleanData['FECHA ALTA'] = fechaEnEspaniol(cleanData['FECHA ALTA']);
    }
    if (cleanData['FECHA BAJA']) {
      cleanData['FECHA BAJA'] = fechaEnEspaniol(cleanData['FECHA BAJA']);
    }

    headers.forEach((header, index) => {
      const hNorm = normalizeHeader(header);
      
      // Intentar obtener el valor del objeto enviado usando la clave normalizada
      // cleanData ya viene con claves en MAYUSCULAS del formulario, pero las normalizaremos para seguridad
      const cleanDataNormalized = {};
      Object.keys(cleanData).forEach(k => cleanDataNormalized[normalizeHeader(k)] = cleanData[k]);

      let valToSet = cleanDataNormalized[hNorm];
      let fieldWasSent = Object.prototype.hasOwnProperty.call(cleanDataNormalized, hNorm);

      let finalVal;
      if (fieldWasSent) {
        finalVal = valToSet;
      } else {
        finalVal = sheet.getRange(sheetRow, index + 1).getValue();
      }

      if (hNorm === 'DNI' && finalVal) {
        finalVal = formatearDNI(String(finalVal));
      }

      if (typeof finalVal === 'string') {
        finalVal = finalVal.trim();
      }

      updatedValues.push(finalVal);
    });

    sheet.getRange(sheetRow, 1, 1, updatedValues.length).setValues([updatedValues]);

    // Verificar si existe BAJA CONTRATO y cambiar el estado automáticamente a "Baja"
    const normalizedHeaders = headers.map(normalizeHeader);
    const bajaContratIndex = normalizedHeaders.indexOf('BAJA CONTRATO');
    const estadoIndex = normalizedHeaders.indexOf('ESTADO');
    
    if (bajaContratIndex !== -1 && estadoIndex !== -1) {
      const bajaContratValue = updatedValues[bajaContratIndex];
      if (bajaContratValue && String(bajaContratValue).trim() !== '') {
        sheet.getRange(sheetRow, estadoIndex + 1).setValue('Baja');
      }
    }

    return { success: true, message: `Registro actualizado exitosamente.` };
  } catch (e) {
    Logger.log("Error al actualizar fila: " + e.message);
    return { success: false, message: "Error al actualizar registro: " + e.message };
  }
}

function deleteRow(rowIndex) {
  const sheet = getMainSheet();

  if (!sheet) {
    return { success: false, message: "Hoja principal no encontrada." };
  }

  const sheetRowToDelete = findRowById(sheet, rowIndex);
  if (sheetRowToDelete === -1) {
    return { success: false, message: "Registro con ID " + rowIndex + " no encontrado para eliminar." };
  }

  try {
    // Obtener datos antes de borrar para el log
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const rowValues = sheet.getRange(sheetRowToDelete, 1, 1, sheet.getLastColumn()).getValues()[0];
    const apellidosIdx = headers.indexOf('APELLIDOS');
    const autoridadIdx = headers.indexOf('AUTORIDAD');
    const apellidos = apellidosIdx !== -1 ? rowValues[apellidosIdx] : 'N/A';
    const autoridad = autoridadIdx !== -1 ? rowValues[autoridadIdx] : 'N/A';

    sheet.deleteRow(sheetRowToDelete);

    // Registrar actividad
    logActivity(`Elimina Registro ${apellidos}_${autoridad}`);

    const allData = sheet.getDataRange().getValues();
    if (allData.length > 1) {
      const idColumnIndex = headers.indexOf('Id');

      if (idColumnIndex === -1) {
        Logger.log("Columna 'Id' no encontrada para reajustar después de la eliminación.");
        return { success: true, message: `Registro en fila ${sheetRowToDelete} eliminado, pero no se pudo reajustar IDs.` };
      }

      const dataToUpdate = allData.slice(1);
      for (let i = 0; i < dataToUpdate.length; i++) {
        dataToUpdate[i][idColumnIndex] = i + 1;
      }
      sheet.getRange(1, 1, allData.length, headers.length).setValues([allData[0], ...dataToUpdate]);
    }

    return { success: true, message: `Registro eliminado y IDs reajustados.` };
  } catch (e) {
    Logger.log("Error al eliminar fila: " + e.message);
    return { success: false, message: "Error al eliminar registro: " + e.message };
  }
}

function updateCheckValue(rowIndex, value) {
  try {
    const sheet = getMainSheet();
    if (!sheet) return { success: false, message: 'Hoja no encontrada.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(normalizeHeader);
    const checkColIndex = headers.indexOf('CHECK');
    if (checkColIndex === -1) return { success: false, message: "Columna 'CHECK' no encontrada." };

    const sheetRow = findRowById(sheet, rowIndex);
    if (sheetRow === -1) return { success: false, message: "Registro con ID " + rowIndex + " no encontrado para actualizar CHECK." };
    sheet.getRange(sheetRow, checkColIndex + 1).setValue(value);
    SpreadsheetApp.flush();
    Logger.log('CHECK actualizado: fila=' + sheetRow + ', valor=' + value);
    return { success: true };
  } catch (e) {
    Logger.log('Error en updateCheckValue: ' + e.message);
    return { success: false, message: e.message };
  }
}

function executeAction(recordId, actionType) {
  const sheet = getMainSheet();

  if (!sheet) {
    Logger.log('Hoja principal no encontrada.');
    return { success: false, message: 'Hoja no encontrada.' };
  }

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(normalizeHeader);
    const idIndex = headers.indexOf('ID');
    
    if (idIndex === -1) {
      return { success: false, message: "Columna 'ID' no encontrada en el sheet." };
    }

    const sheetRow = findRowById(sheet, recordId);

    if (sheetRow === -1) {
      return { success: false, message: 'Registro con ID ' + recordId + ' no encontrado en el sheet.' };
    }

    let statusToSet = '';
    let actionMessage = '';

    const checkboxColumnIndex = headers.indexOf('CHECK') + 1;
    const estadoColumnIndex = headers.indexOf('ESTADO') + 1;

    const rowDataRaw = sheet.getRange(sheetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowObjectRaw = {};
    headers.forEach((upperHeader, index) => { rowObjectRaw[upperHeader] = rowDataRaw[index]; });
    const rowObject = trimData(rowObjectRaw);

    // DEBUG: Log de claves para troubleshooting
    Logger.log("=== Datos recuperados del Sheet (fila " + sheetRow + ") ===");
    Logger.log("Claves en rowObject: " + Object.keys(rowObject).join(" | "));

    // Robustez: Asegurar que APELLIDOS y NOMBRES existan si el Sheet tiene singular u otras variaciones
    // Búsqueda flexible: primero exacta, luego por patrón
    if (!rowObject['APELLIDOS'] || String(rowObject['APELLIDOS']).trim() === '') {
      // Buscar por patrón: cualquier clave que contenga 'APELLIDO' o 'APELLIDOS'
      const apellidoKey = Object.keys(rowObject).find(k => 
        k.includes('APELLIDO') && 
        (String(rowObject[k] || '').trim() !== '')
      );
      if (apellidoKey) {
        Logger.log("APELLIDOS recuperado de: " + apellidoKey + " = '" + String(rowObject[apellidoKey]).trim() + "'");
        rowObject['APELLIDOS'] = String(rowObject[apellidoKey]).trim();
      } else {
        const camposApellido = Object.keys(rowObject).filter(k => k.includes('APELLIDO'));
        Logger.log("ERROR: No se encontró APELLIDOS en ningún campo. Campos APELLIDO disponibles: " + camposApellido.join(" | "));
        rowObject['APELLIDOS'] = '';
      }
    } else {
      Logger.log("APELLIDOS encontrado: '" + String(rowObject['APELLIDOS']).trim() + "'");
      rowObject['APELLIDOS'] = String(rowObject['APELLIDOS']).trim();
    }

    if (!rowObject['NOMBRES'] || String(rowObject['NOMBRES']).trim() === '') {
      // Buscar por patrón: cualquier clave que contenga 'NOMBRE' o 'NOMBRES'
      const nombreKey = Object.keys(rowObject).find(k => 
        k.includes('NOMBRE') && 
        (String(rowObject[k] || '').trim() !== '')
      );
      if (nombreKey) {
        Logger.log("NOMBRES recuperado de: " + nombreKey + " = '" + String(rowObject[nombreKey]).trim() + "'");
        rowObject['NOMBRES'] = String(rowObject[nombreKey]).trim();
      } else {
        const camposNombre = Object.keys(rowObject).filter(k => k.includes('NOMBRE'));
        Logger.log("ERROR: No se encontró NOMBRES en ningún campo. Campos NOMBRE disponibles: " + camposNombre.join(" | "));
        rowObject['NOMBRES'] = '';
      }
    } else {
      Logger.log("NOMBRES encontrado: '" + String(rowObject['NOMBRES']).trim() + "'");
      rowObject['NOMBRES'] = String(rowObject['NOMBRES']).trim();
    }

    Logger.log("Después de robustez - APELLIDOS: '" + rowObject['APELLIDOS'] + "' | NOMBRES: '" + rowObject['NOMBRES'] + "'");

    if (rowObject['CHECK'] === true) {
      return { success: false, message: 'La fila ya está marcada como finalizada. Destíldala para volver a actuar.' };
    }

    const auxColumns = calculateAuxColumns(rowObject);
    const fullRowObject = { ...rowObject, ...auxColumns };

    switch (actionType) {
      case 'Nota Alta':
        autoNota_ALTA(fullRowObject);
        statusToSet = 'Pendiente';
        actionMessage = 'Nota Alta ejecutada.';
        break;
      case 'Contrato': {
        const resContrato = generaContratoWord(fullRowObject as unknown as ConRowData);
        statusToSet = 'Contrato';
        actionMessage = 'Contrato generado exitosamente (' + resContrato.nombre + ').';
        // Registrar la fecha de generación del contrato
        const gcIndex = headers.indexOf('GENERA CONTRATO') + 1;
        if (gcIndex > 0) {
          const now = new Date();
          const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy") + '_' + (fullRowObject['AUTOR'] || '');
          sheet.getRange(sheetRow, gcIndex).setValue(dateStr);
        }
        break;
      }
      case 'Resolución': {
        if (!fullRowObject['REQ'] || String(fullRowObject['REQ']).trim() === '') {
          return { success: false, message: 'Falta el N° de Requerimiento (REQ). Este campo es obligatorio para generar la Resolución.' };
        }
        const resResolucion = generaResolucionTXT(fullRowObject);
        statusToSet = 'Resolución';
        actionMessage = 'Resolución generada exitosamente (' + resResolucion.nombre + ').';
        // Registrar la fecha de generación de la resolución
        const grIndex = headers.indexOf('GENERA RESOLUCION') + 1;
        if (grIndex > 0) {
          const now = new Date();
          const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy") + '_' + (fullRowObject['AUTOR'] || '');
          sheet.getRange(sheetRow, grIndex).setValue(dateStr);
        }
        // Registrar la fecha de notificación
        const fnIndex = headers.indexOf('FECHA NOTIFICA') !== -1 ? headers.indexOf('FECHA NOTIFICA') + 1 :
          (headers.indexOf('FECHA NOTIFICACIÓN') !== -1 ? headers.indexOf('FECHA NOTIFICACIÓN') + 1 :
            (headers.indexOf('FECHA NOTIFICACION') !== -1 ? headers.indexOf('FECHA NOTIFICACION') + 1 : 0));
        if (fnIndex > 0) {
          const now = new Date();
          const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy HH:mm") + '_' + (fullRowObject['AUTOR'] || '');
          sheet.getRange(sheetRow, fnIndex).setValue(dateStr);
        }
        break;
      }
      case 'Baja': {
        if (!fullRowObject['REQ'] || String(fullRowObject['REQ']).trim() === '') {
          return { success: false, message: 'Falta el N° de Requerimiento (REQ). Este campo es obligatorio para procesar la Baja.' };
        }
        const resBaja = bajaResolucionTXT(fullRowObject);
        if (resBaja.success) {
          statusToSet = 'Baja';
          actionMessage = 'Baja de Resolución generada exitosamente (' + resBaja.nombre + ').';
          // Registrar la fecha de baja de contrato en la columna correspondiente
          const bcIndex = headers.indexOf('BAJA CONTRATO') + 1;
          if (bcIndex > 0) {
            const now = new Date();
            const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy") + '_' + (fullRowObject['AUTOR'] || '');
            sheet.getRange(sheetRow, bcIndex).setValue(dateStr);
          }
          // Registrar la fecha de notificación (FECHA NOTIFICA)
          const fnIndex = headers.indexOf('FECHA NOTIFICA') !== -1 ? headers.indexOf('FECHA NOTIFICA') + 1 :
            (headers.indexOf('FECHA NOTIFICACIÓN') !== -1 ? headers.indexOf('FECHA NOTIFICACIÓN') + 1 :
              (headers.indexOf('FECHA NOTIFICACION') !== -1 ? headers.indexOf('FECHA NOTIFICACION') + 1 : 0));
          if (fnIndex > 0) {
            const now = new Date();
            const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy HH:mm") + '_' + (fullRowObject['AUTOR'] || '');
            sheet.getRange(sheetRow, fnIndex).setValue(dateStr);
          }
        } else {
          return { success: false, message: 'Error en Baja: ' + resBaja.message };
        }
        break;
      }
      case 'Notificación': {
        const notificationResult = notificaLocaciones1(fullRowObject);
        if (notificationResult.success) {
          statusToSet = 'Notificado';
          actionMessage = notificationResult.message;
          const fnIndex = headers.indexOf('FECHA NOTIFICA') !== -1 ? headers.indexOf('FECHA NOTIFICA') + 1 :
            (headers.indexOf('FECHA NOTIFICACIÓN') !== -1 ? headers.indexOf('FECHA NOTIFICACIÓN') + 1 :
              (headers.indexOf('FECHA NOTIFICACION') !== -1 ? headers.indexOf('FECHA NOTIFICACION') + 1 : 0));
          if (fnIndex > 0) {
            const now = new Date();
            const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy HH:mm") + '_' + (fullRowObject['AUTOR'] || '');
            sheet.getRange(sheetRow, fnIndex).setValue(dateStr);
          }
        } else {
          return { success: false, message: 'Error en notificación: ' + notificationResult.message };
        }
        break;
      }
      default:
        return { success: false, message: 'Acción no reconocida: ' + actionType };
    }

    if (estadoColumnIndex > 0) sheet.getRange(sheetRow, estadoColumnIndex).setValue(statusToSet);
    if (checkboxColumnIndex > 0) sheet.getRange(sheetRow, checkboxColumnIndex).setValue(true);

    return { success: true, message: actionMessage + ' Estado: ' + statusToSet };
  } catch (e) {
    Logger.log('Error en executeAction [' + actionType + '] registro ID ' + recordId + ': ' + e.message + '\n' + e.stack);
    return { success: false, message: 'Error al ejecutar "' + actionType + '": ' + e.message };
  }
}

function markAsNotified(rowIndex, isNotified) {
  if (isNotified) {
    return executeAction(rowIndex, 'Notificación');
  } else {
    return { success: false, message: "La acción de desmarcar no está implementada directamente aquí." };
  }
}

function executeBulkAction(actionType) {
  const sheet = getMainSheet();
  if (!sheet) return { success: false, message: 'Hoja no encontrada.' };

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(normalizeHeader);
  const checkIndex = headers.indexOf('CHECK');
  const idIndex = headers.indexOf('ID');

  if (checkIndex === -1 || idIndex === -1) {
    return { success: false, message: "Columna 'CHECK' o 'ID' no encontrada." };
  }

  let count = 0;
  let errors: string[] = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const isChecked = row[checkIndex];
    const hasContent = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
    if (hasContent && (isChecked === false || isChecked === '' || isChecked === null || isChecked === undefined || isChecked === 'FALSE')) {
      try {
        const recordId = row[idIndex];
        const result = executeAction(recordId, actionType);
        if (result.success) count++;
        else errors.push('Fila ' + (i + 1) + ': ' + result.message);
      } catch (e) {
        errors.push('Fila ' + (i + 1) + ': ' + e.toString());
      }
    }
  }

  if (count === 0 && errors.length === 0) {
    return { success: false, message: "No hay registros sin procesar (CHECK destildado). Destildá las filas que querés procesar." };
  }
  if (count === 0 && errors.length > 0) {
    return { success: false, message: 'No se procesó ningún registro. Errores: ' + errors.join('; ') };
  }

  return {
    success: true,
    count: count,
    message: 'Se ejecutó "' + actionType + '" en ' + count + ' registros.' + (errors.length > 0 ? ' Errores en ' + errors.length + ' filas.' : '')
  };
}

function checkAndAssignIds() {
  const sheet = getMainSheet();
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).toUpperCase());
  const idIndex = headers.indexOf('ID');
  const estadoIndex = headers.indexOf('ESTADO');
  const checkIndex = headers.indexOf('CHECK');

  if (idIndex === -1) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = dataRange.getValues();

  let maxId = 0;
  for (let i = 0; i < values.length; i++) {
    const idValue = values[i][idIndex];
    if (typeof idValue === 'number' && idValue > maxId) {
      maxId = idValue;
    }
  }

  let nextId = maxId + 1;
  let hasChanges = false;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const idValue = row[idIndex];

    const hasContent = row.some((cell, idx) => idx !== idIndex && cell !== '' && cell !== null && cell !== undefined);
    if (!hasContent) continue;

    const isNewRow = (idValue === '' || idValue === null || idValue === undefined);

    if (isNewRow) {
      values[i][idIndex] = nextId;
      nextId++;
      hasChanges = true;

      if (estadoIndex !== -1) {
        const estadoValue = row[estadoIndex];
        if (estadoValue === '' || estadoValue === null || estadoValue === undefined) {
          values[i][estadoIndex] = 'Pendiente';
          hasChanges = true;
        }
      }

      if (checkIndex !== -1) {
        const checkValue = row[checkIndex];
        if (checkValue === '' || checkValue === null || checkValue === undefined) {
          values[i][checkIndex] = true;
          hasChanges = true;
        }
      }
    } else {
      if (estadoIndex !== -1) {
        const estadoValue = row[estadoIndex];
        if (estadoValue === '' || estadoValue === null || estadoValue === undefined) {
          values[i][estadoIndex] = 'Pendiente';
          hasChanges = true;
        }
      }
    }
  }

  if (hasChanges) {
    dataRange.setValues(values);
    SpreadsheetApp.flush();
  }
}

function getFilePreviewUrl(rowData, type) {
  try {
    let folderId = "";

    const req = String(rowData['REQ'] || 'S/N').trim();
    const apellidos = String(rowData['APELLIDOS'] || 'S/A').trim();
    const autoridad = String(rowData['AUTORIDAD'] || 'S/Aut').trim();

    if (type === 'CONTRATO') {
      folderId = CARPETA_FUSION_CONTRATO;
    } else if (type === 'RESOLUCION') {
      folderId = CARPETA_FUSION_RESOLUCIONES;
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      if (type === 'CONTRATO') {
        if (name.toLowerCase().includes("contrato") &&
          apellidos !== "" && name.toLowerCase().includes(apellidos.toLowerCase()) &&
          autoridad !== "" && name.toLowerCase().includes(autoridad.toLowerCase())) {
          return {
            success: true,
            url: file.getUrl().replace('/view', '/preview'),
            name: name
          };
        }
      } else {
        if (req !== "" && name.includes(req) && apellidos !== "" && name.includes(apellidos) && name.toLowerCase().includes(".txt")) {
          return {
            success: true,
            url: file.getUrl().replace('/view', '/preview'),
            name: name
          };
        }
      }
    }

    return { success: false, message: "Archivo no encontrado en Drive." };
  } catch (e) {
    return { success: false, message: "Error crítico al cargar datos: " + e.toString() };
  }
}

/**
 * Función ligera para verificar si hay cambios en el Sheet sin descargar todo.
 */
function checkDataVersion(filterOptions) {
  try {
    const sheet = getMainSheet();
    if (!sheet) return { success: false, message: "Hoja no encontrada." };
    
    // Obtenemos el número de filas y el valor de la última fila (o un hash simple)
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, version: "empty" };
    
    // Tomamos una muestra de la última fila para detectar cambios rápidos
    const lastRowValues = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const version = lastRow + "_" + lastRowValues.join("|").substring(0, 50).replace(/\s+/g, "");
    
    return { success: true, version: version };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Busca el índice de la fila (1-based) de un registro dado su ID único.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja donde buscar.
 * @param {string|number} id ID del registro a buscar.
 * @returns {number} El número de fila (1-indexed) o -1 si no se encuentra.
 */
function findRowById(sheet: any, id: string | number): number {
  if (!sheet || id === null || id === undefined) return -1;
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return -1;
  
  const headers = values[0].map(normalizeHeader);
  const idCol = headers.indexOf('ID');
  if (idCol === -1) return -1;
  
  const targetId = String(id).trim();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]).trim() === targetId) {
      return i + 1; // 1-indexed
    }
  }
  return -1;
}
