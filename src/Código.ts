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

function getMainSheet() { return getSS().getSheetByName("REGISTROS"); }
function getBajaResSheet() { return getSS().getSheetByName("BAJA RES"); }
function getControlSheet() { return getSS().getSheetByName("Control"); }

let CONFIG_CACHE: any = null; // Cache para la configuración de autoridades

// normalizeHeader se encuentra ahora en utils.ts

// parseSafeDate se encuentra ahora en utils.ts

/**
 * Lee la hoja "AUTORIDADES" y devuelve mapeo de autoridad -> {senadora, secretaria, correo}
 * Lee el rango A1:D44 (encabezados en fila 1, datos en filas 2-44)
 * Columnas esperadas: A=Senador/a (apellido), B=Senador/a (nombre+apellido), C=Secretario/a, D=Correo
 */
function loadAutoridadesConfig() {
  if (CONFIG_CACHE) return CONFIG_CACHE;

  try {
    const data = firestoreGetAllDocs("Autoridades");
    if (!data || data.length === 0) {
      Logger.log("Aviso: Colección 'Autoridades' no encontrada en Firebase.");
      return {};
    }

    const config = {};
    for (const row of data) {
      const apellido = String(row['APELLIDO'] || row['AUTORIDAD'] || '').trim().toUpperCase();
      const nombreCompleto = String(row['NOMBRE COMPLETO'] || '').trim();
      const secretario = String(row['SECRETARIO'] || '').trim();
      const correo = String(row['CORREO'] || '').trim();

      if (apellido && secretario && correo) {
        config[apellido] = {
          apellido: apellido,
          nombreCompleto: nombreCompleto,
          secretario: secretario,
          correo: correo
        };
      }
    }

    Logger.log("Config cargada desde Firebase: " + Object.keys(config).length + " autoridades encontradas");
    CONFIG_CACHE = config;
    return config;
  } catch (e) {
    Logger.log("Error cargando config de autoridades desde Firebase: " + e.message);
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

function addAutoridadConfig(apellido: string, nombreCompleto: string, secretario: string, correo: string) {
  try {
    const apellidoUpper = String(apellido || '').trim().toUpperCase();
    if (!apellidoUpper || !secretario || !correo) {
      return { success: false, message: "Faltan datos requeridos: Apellido, Secretario y Correo son obligatorios." };
    }

    // Verificar que no exista ya
    const config = loadAutoridadesConfig();
    if (config[apellidoUpper]) {
      return { success: false, message: "La autoridad '" + apellidoUpper + "' ya existe." };
    }

    const docId = apellidoUpper; 
    const data = {
      'APELLIDO': apellidoUpper,
      'NOMBRE COMPLETO': nombreCompleto.trim(),
      'SECRETARIO': secretario.trim(),
      'CORREO': correo.trim(),
      'ID_AUT': docId
    };

    firestoreUpdateDocument("Autoridades", docId, data);
    CONFIG_CACHE = null; 
    
    return { success: true, message: "Autoridad '" + apellidoUpper + "' agregada en Firebase." };
  } catch (e) {
    return { success: false, message: "Error (Firebase): " + e.message };
  }
}

/**
 * Edita una autoridad existente en la hoja AUTORIDADES
 */
function updateAutoridadConfig(oldApellido: string, newApellido: string, secretario: string, correo: string) {
  try {
    const oldKey = String(oldApellido || '').trim().toUpperCase();
    const newKey = String(newApellido || '').trim().toUpperCase();

    if (!oldKey || !newKey || !secretario || !correo) {
      return { success: false, message: "Faltan datos requeridos." };
    }

    // Si cambió el apellido, borramos el viejo y creamos el nuevo
    if (oldKey !== newKey) {
      firestoreDeleteDocument("Autoridades", oldKey);
    }

    const data = {
      'APELLIDO': newKey,
      'NOMBRE COMPLETO': String(newApellido || '').trim(),
      'SECRETARIO': secretario.trim(),
      'CORREO': correo.trim(),
      'ID_AUT': newKey
    };

    firestoreUpdateDocument("Autoridades", newKey, data);
    CONFIG_CACHE = null;

    return { success: true, message: "Autoridad actualizada en Firebase." };
  } catch (e) {
    return { success: false, message: "Error al actualizar (Firebase): " + e.message };
  }
}

/**
 * Elimina una autoridad de la hoja AUTORIDADES
 */
function deleteAutoridadConfig(apellido: string) {
  try {
    const key = String(apellido || '').trim().toUpperCase();
    if (!key) return { success: false, message: "Apellido requerido." };

    firestoreDeleteDocument("Autoridades", key);
    CONFIG_CACHE = null;

    return { success: true, message: "Autoridad '" + key + "' eliminada de Firebase." };
  } catch (e) {
    return { success: false, message: "Error al eliminar: " + e.message };
  }
}

/**
 * Obtiene lista de tareas desde la hoja "AUXILIARES" rango A2:A200
 */
function getTareasList(ts?: any) {
  try {
    const auxiliaries = firestoreGetAllDocs("Auxiliares");
    if (!auxiliaries || auxiliaries.length === 0) {
      Logger.log("Aviso: Colección 'Auxiliares' vacía o no encontrada en Firebase.");
      return [];
    }

    // Se asume que las tareas están en una propiedad llamada 'TAREA' o similar
    // o simplemente tomamos los valores no vacíos de los documentos
    const tareas: string[] = auxiliaries
      .map(doc => String(doc['TAREA'] || doc['VALOR'] || '').trim())
      .filter(t => t !== '');

    Logger.log("Tareas cargadas desde Firebase (Auxiliares): " + tareas.length);
    return tareas;
  } catch (e) {
    Logger.log("Error obteniendo lista de tareas (Firebase): " + e.message);
    return [];
  }
}

/**
 * Agrega nueva tarea en la hoja AUXILIARES
 */
function addTareaConfig(tarea: string) {
  try {
    const tareaLimpia = String(tarea || '').trim();
    if (!tareaLimpia) {
      return { success: false, message: "Tarea vacía." };
    }

    const sheet = getSS().getSheetByName("AUXILIARES");
    if (!sheet) {
      return { success: false, message: "Hoja 'AUXILIARES' no encontrada." };
    }

    const range = sheet.getRange("A2:A200");
    const values = range.getValues();
    let insertRow = -1;

    for (let i = 0; i < values.length; i++) {
      if (!values[i][0] || String(values[i][0]).trim() === '') {
        insertRow = i + 2;
        break;
      }
    }

    if (insertRow === -1) {
      return { success: false, message: "No hay espacio disponible en la hoja AUXILIARES." };
    }

    sheet.getRange(insertRow, 1).setValue(tareaLimpia);
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

    const range = auxSheet.getRange("E2:E100");
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
        foundIdx = i + 2; // +2 porque range empieza en E2
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
          auxSheet.getRange('E' + newRow).setValue(newNameLimpia);
          SpreadsheetApp.flush();
          Logger.log("Tarea '" + oldNameLimpia + "' agregada como nueva en fila " + newRow + ": " + newNameLimpia);
          return { success: true, message: "Tarea agregada exitosamente." };
        }
      }
      return { success: false, message: "Límite de tareas alcanzado (máx. 100)." };
    }

    // Si existe, actualizarla
    auxSheet.getRange('E' + foundIdx).setValue(newNameLimpia);
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

    const range = auxSheet.getRange("E2:E100");
    const values = range.getValues();

    // Buscar la tarea (case-insensitive)
    let foundIdx = -1;
    for (let i = 0; i < values.length; i++) {
      const cellValue = String(values[i][0] || '').trim().toUpperCase();
      if (cellValue === tareaNameLimpia) {
        foundIdx = i + 2; // +2 porque range empieza en E2
        break;
      }
    }

    if (foundIdx === -1) {
      Logger.log("Tarea '" + tareaNameLimpia + "' no encontrada en aux.");
      return { success: true, message: "Tarea eliminada (lista refresca automáticamente)." };
    }

    auxSheet.getRange('E' + foundIdx).clearContent();
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
  const props = PropertiesService.getUserProperties();
  const userName = props.getProperty('CD_NOMBRE') || '';
  const userStats: { role: string | null, autor: string | null } = { role: null, autor: null };

  if (userName) {
    try {
      const todosUsuarios = firestoreGetAllDocs("Usuarios");
      const user = todosUsuarios.find(u => String(u['NOMBRE'] || '').trim().toUpperCase() === userName.toUpperCase());
      if (user) {
        userStats.role = String(user['ROL'] || '').trim().toUpperCase();
        userStats.autor = String(user['AUTOR_COD'] || '').trim().toUpperCase();
      }
    } catch (e) {
      userStats.role = getCurrentUserRole();
    }
  }

  return userStats;
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
      .setTitle('LEGALTEC - CONTROL')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── Módulo Control Docus ──────────────────────────────────────────────────
  if (page === 'ControlDocus') {
    const t = HtmlService.createTemplateFromFile('LoginCD');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('LEGALTEC - CONTROL')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'PortalUsuario') {
    const t = HtmlService.createTemplateFromFile('PortalUsuario');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('LEGALTEC - CONTROL')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'PanelSupervisor') {
    const t = HtmlService.createTemplateFromFile('PanelSupervisor');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('LEGALTEC - CONTROL')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'PanelAdmin') {
    const t = HtmlService.createTemplateFromFile('PanelAdmin');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('LEGALTEC - ADMIN HUB')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'Index') {
    const t = HtmlService.createTemplateFromFile('Index');
    t.appUrl = appUrl;
    return t.evaluate()
      .setTitle('LEGALTEC - CONTROL')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Fallback al login para cualquier otra ruta no reconocida o protegida
  const t = HtmlService.createTemplateFromFile('LoginCD');
  t.appUrl = appUrl;
  return t.evaluate().setTitle('LEGALTEC - CONTROL');
}

function include(filename) {
  // Función auxiliar para incluir archivos CSS o JS en la plantilla HTML.
  // Asegura que el nombre de archivo pasado sea el que se utiliza.
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}



/**
 * Lee la hoja 'aux' y devuelve un mapa de AUTORIDAD → { secretario, correoSecretario }.
 * Columnas: B=AUTORIDAD, C=SECRETARIO, D=CORREO SECRETARIO (filas 2 a 43).
 */
function getAuxData() {
  const cacheKey = "AUX_DATA_MAP_V3_" + SPREADSHEET_ID;
  const cached = getCache(cacheKey);
  if (cached) return { success: true, data: cached.auxMap, autores: cached.autores };

  try {
    // Firebase-First: Consumimos de la colección Autoridades en lugar del Sheet
    const autoresFirebase = firestoreGetAllDocs("Autoridades") || [];
    const auxMap = {};
    
    autoresFirebase.forEach(doc => {
       // Buscamos las claves de forma robusta mediante normalización
       let autoridad = "", secretario = "", correo = "";
       
       for (let key in doc) {
         const normKey = normalizeHeader(key);
         // Buscamos coincidencias flexibles para evitar fallos por barras o minúsculas
         if (normKey.indexOf("APELLIDO SENADOR") !== -1) autoridad = String(doc[key] || '').trim().toUpperCase();
         if (normKey.indexOf("SECRETARIO") !== -1 && normKey.indexOf("CORREO") === -1) secretario = String(doc[key] || '').trim();
         if (normKey.indexOf("CORREO") !== -1) correo = String(doc[key] || '').trim();
       }

       if (autoridad) {
         auxMap[autoridad] = { secretario, correo: correo };
       }
    });

    // Obtener lista de autores desde Firebase
    const todosUsuarios = firestoreGetAllDocs("Usuarios");
    const listaAutores = Array.from(new Set(todosUsuarios.map(u => String(u['AUTOR_COD'] || '').trim().toUpperCase()).filter(a => a !== ''))).sort();

    const finalResult = { auxMap, autores: listaAutores };
    setCache(cacheKey, finalResult, 600); // 10 min cache
    return { success: true, data: auxMap, autores: listaAutores };
  } catch (e) {
    Logger.log("Error en getAuxData: " + e.message);
    return { success: false, message: e.message, data: {}, autores: [] };
  }
}

interface FilterOptions {
  type: 'recent' | 'period' | 'all';
  month?: number | string;
  year?: number | string;
}

function getData(filterOptions: FilterOptions = { type: 'recent' }) {
  try {
    let allHeaders: string[] = [];
    let rawDataObjects: any[] = [];
    let usingFirebase = false;

    // INTENTO DE LECTURA DESDE FIREBASE
    try {
      const records = firestoreGetAllDocs("Registros");
      if (records && records.length > 0) {
        rawDataObjects = records;
        allHeaders = Object.keys(records[0]).map(normalizeHeader);
        usingFirebase = true;
        Logger.log("✅ Datos obtenidos desde Firebase (" + records.length + " registros)");
      }
    } catch (firebaseErr) {
      Logger.log("⚠️ Falló lectura de Firebase, reintentando con Sheets: " + firebaseErr.message);
    }

    // SI FIREBASE FALLA O ESTÁ VACÍO, USAR SHEETS (Fallback)
    if (!usingFirebase) {
      checkAndAssignIds(); 
      const sheet = getMainSheet();
      if (!sheet) return { success: false, message: "Hoja de datos 'Registros' no encontrada." };
      
      const range = sheet.getDataRange();
      const allValues = range.getValues();
      if (!allValues || allValues.length === 0) return { success: true, headers: [], data: [] };

      allHeaders = allValues[0].map(normalizeHeader);
      rawDataObjects = allValues.slice(1); // Aquí rawDataObjects son FILAS (arrays)
    } else {
      // Si viene de Firebase, tenemos objetos. Para mantener compatibilidad con el loop inferior,
      // convertimos los objetos a FILAS (arrays) siguiendo el orden de allHeaders.
      rawDataObjects = rawDataObjects.map(obj => {
        return allHeaders.map(h => obj[h] !== undefined ? obj[h] : "");
      });
    }

    // A partir de aquí, rawData siempre es un array de FILAS (arrays)
    const rawData = rawDataObjects;
    
    // Identificar índices para filtrado (Flexible: ALTA o DE ALTA)
    // ... (El resto de la lógica de filtrado se mantiene igual)


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
    // Intentar detectar si el ID se llama ID_REG o ID
    let finalIdHeader = allHeaders.includes('ID_REG') ? 'ID_REG' : 'ID';

    const visibleClientHeadersNames = [finalIdHeader, 'AUTORIDAD', 'APELLIDOS', 'NOMBRES', 'DNI', 'TAREAS', 'FECHA ALTA', 'TOTAL', 'REQ', 'CHECK', 'ESTADO'];
    const hiddenClientHeadersNames = ['SECRETARIO', 'CORREO SECRETARIO', 'CORREO LOCADOR', 'SEXO', 'FECHA DE BAJA', 'CUOTA', 'DOMICILIO', 'LOCALIDAD', 'GENERA CONTRATO', 'GENERA RESOLUCION', 'FECHA NOTIFICA', 'BAJA CONTRATO', 'AUTOR', 'CREA REGISTRO'];

    const config = loadAutoridadesConfig();

    // Transformación final, sanitización y enriquecimiento
    const finalProcessed = filteredAndProcessed.map(rowObject => {
      if (rowObject['BAJA CONTRATO'] && String(rowObject['BAJA CONTRATO']).trim() !== '') {
        rowObject['ESTADO'] = 'Baja';
      }

      // Enriquecer con metadatos de autoridad (frescura ante todo)
      const authKey = String(rowObject['AUTORIDAD'] || '').trim().toUpperCase();
      if (config[authKey]) {
        rowObject['SECRETARIO'] = config[authKey].secretario;
        rowObject['CORREO SECRETARIO'] = config[authKey].correo;
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
      const recordId = rowObject['__id'] || rowObject[finalIdHeader] || rowObject['ID_REG'] || rowObject['ID'];
      const sortDate = parseSafeDate(rowObject['FECHA ALTA']).getTime() || 0;

      return { visible, hidden, recordId, sortDate };
    });

    // Ordenar por Fecha Alta (Desc) y luego ID (Desc).
    // Los registros sin FECHA ALTA (sortDate=0, como los renovados recén creados)
    // se tratan como los más recientes para que aparezcan al principio de la tabla.
    finalProcessed.sort((a, b) => {
      const sa = a.sortDate === 0 ? Infinity : a.sortDate;
      const sb = b.sortDate === 0 ? Infinity : b.sortDate;
      return (sb - sa) || (Number(b.recordId) - Number(a.recordId));
    });

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
    const idsToFind = (Array.isArray(ids) ? ids : [ids]).map(id => String(id).trim());
    if (idsToFind.length === 0) return { success: true, data: [] };

    const records = firestoreGetAllDocs("Registros");

    const config = loadAutoridadesConfig();
    const matched = records.filter(r => idsToFind.includes(String(r['ID'] || r['ID_REG'] || r['__id'])));

    const visibleNames = ['ID', 'AUTORIDAD', 'APELLIDOS', 'NOMBRES', 'DNI', 'TAREAS', 'FECHA ALTA', 'TOTAL', 'REQ', 'CHECK', 'ESTADO'];
    const hiddenNames = ['SECRETARIO', 'CORREO SECRETARIO', 'CORREO LOCADOR', 'SEXO', 'FECHA BAJA', 'CUOTA', 'DOMICILIO', 'LOCALIDAD', 'GENERA CONTRATO', 'GENERA RESOLUCION', 'FECHA NOTIFICA', 'BAJA CONTRATO', 'AUTOR', 'CREA REGISTRO'];

    const processed = matched.map(rowObject => {
      const authKey = String(rowObject['AUTORIDAD'] || '').trim().toUpperCase();
      if (config[authKey]) {
        rowObject['SECRETARIO'] = config[authKey].secretario;
        rowObject['CORREO SECRETARIO'] = config[authKey].correo;
      }

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

      const vPart = visibleNames.map(h => sanitize(rowObject[h] !== undefined ? rowObject[h] : ''));
      const hPart = hiddenNames.map(h => sanitize(rowObject[h] !== undefined ? rowObject[h] : ''));
      const recordId = String(rowObject['__id'] || rowObject['ID'] || rowObject['ID_REG'] || '');

      return [...vPart, '', ...hPart, recordId];
    });

    return { success: true, data: processed };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteRow(recordId: string) {
  try {
    const record = firestoreGetDoc("Registros", String(recordId));
    if (!record) return { success: false, message: "Registro no encontrado en Firebase." };

    firestoreDeleteDocument("Registros", String(recordId));
    
    const label = (record['APELLIDOS'] || 'N/A') + '_' + (record['AUTORIDAD'] || 'N/A');
    logActivity(`Elimina Registro ${label} (Firebase-First)`);

    return { success: true, message: "Registro eliminado de Firebase." };
  } catch (e) {
    Logger.log("Error en deleteRow (Firebase): " + e.message);
    return { success: false, message: "Error al eliminar: " + e.message };
  }
}

function addRow(rowData: any, activeAutor?: string) {
  try {
    const records = firestoreGetAllDocs("Registros") || [];
    
    // Calcular el siguiente ID
    let maxId = 0;
    records.forEach(r => {
      const idVal = parseInt(r['ID'] || r['ID_REG'] || 0, 10);
      if (!isNaN(idVal) && idVal > maxId) maxId = idVal;
    });
    const nextId = maxId + 1;

    // Limpiar espacios vacíos y formatear DNI y Fechas
    let cleanData = trimData(rowData);
    if (cleanData['DNI']) {
      cleanData['DNI'] = formatearDNI(cleanData['DNI']);
    }

    // Formatear fechas si existen (FECHA ALTA, FECHA BAJA)
    const dateFields = ['FECHA ALTA', 'FECHA BAJA'];
    dateFields.forEach(f => {
      if (cleanData[f]) {
        cleanData[f] = fechaEnEspaniol(cleanData[f]);
      }
    });

    // Asignar el nuevo ID en ambos campos para compatibilidad total con la web y Drive
    cleanData['ID'] = nextId; 
    cleanData['ID_REG'] = nextId; 
    cleanData['CHECK'] = true; 
    
    // Audit Trail: Creación
    if (activeAutor) {
      cleanData['CREA REGISTRO'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "d/M/yyyy HH:mm") + '_' + activeAutor;
    }

    // Escribimos en Firebase
    firestoreUpdateDocument("Registros", String(nextId), cleanData);

    // Registrar actividad de ingreso (usando el autor activo para el log optimizado)
    logActivity(`Ingreso ${cleanData['APELLIDOS'] || 'N/A'}_${cleanData['AUTORIDAD'] || 'N/A'}`, activeAutor);

    return { success: true, message: "Registro agregado exitosamente en Firebase con Id: " + nextId, newId: nextId };
  }
  catch (e) {
    Logger.log("Error al agregar registro (Firebase): " + e.message);
    return { success: false, message: "Error al agregar registro: " + e.message };
  }
}

function uploadNotaPdf(base64Data: string, apellidos: string, autoridad: string, recordId: number, autor: string = '') {
  try {
    // Usar el ID del registro proporcionado con formato de 3 dígitos
    const fileNumber = String(recordId).padStart(3, '0');
    // Mismo patrón de nombre que Contratos y Resoluciones: NNN) AUTORIZA APELLIDOS-AUTORIDAD_AUTOR
    const autorSuffix = autor ? `_${autor}` : '';
    const fileName = `${fileNumber}) AUTORIZA ${apellidos}-${autoridad}${autorSuffix}.pdf`;

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

function updateRow(recordId: string, rowData: any, activeAutor?: string) {
  try {
    const record = firestoreGetDoc("Registros", String(recordId));
    if (!record) return { success: false, message: "Registro no encontrado en Firebase (ID: " + recordId + ")." };

    let cleanData = trimData(rowData);
    
    // Formateo de DNI
    if (cleanData['DNI']) {
      cleanData['DNI'] = formatearDNI(cleanData['DNI']);
    }

    // Formateo de fechas
    const dateFields = ['FECHA ALTA', 'FECHA BAJA'];
    dateFields.forEach(f => {
      if (cleanData[f]) {
        cleanData[f] = fechaEnEspaniol(cleanData[f]);
      }
    });

    // Actualizamos Firebase
    firestoreUpdateDocument("Registros", String(recordId), cleanData);

    // Verificación automática de Estado de Baja
    if (cleanData['BAJA CONTRATO'] && String(cleanData['BAJA CONTRATO']).trim() !== '') {
      firestoreUpdateDocument("Registros", String(recordId), { "ESTADO": "Baja" });
    }

    return { success: true, message: "Registro actualizado exitosamente en Firebase." };
  } catch (e) {
    Logger.log("Error en updateRow (Firebase): " + e.message);
    return { success: false, message: "Error al actualizar registro: " + e.message };
  }
}

function updateCheckValue(recordId, value) {
  try {
    firestoreUpdateDocument("Registros", String(recordId), { "CHECK": value });
    Logger.log('CHECK actualizado en Firebase: ID=' + recordId + ', valor=' + value);
    return { success: true };
  } catch (e) {
    Logger.log('Error en updateCheckValue (Firebase): ' + e.message);
    return { success: false, message: e.message };
  }
}

function executeAction(recordId: string, actionType: string, preFetchedRecord?: any, activeAutor?: string): { success: boolean; message: string } {
  try {
    const record = preFetchedRecord || firestoreGetDoc("Registros", String(recordId));
    if (!record) return { success: false, message: "Registro no encontrado en Firebase (ID: " + recordId + ")." };

    // Robustez de nombres/apellidos
    if (!record['APELLIDOS'] || String(record['APELLIDOS']).trim() === '') {
      const apellidoKey = Object.keys(record).find(k => k.includes('APELLIDO') && String(record[k] || '').trim() !== '');
      record['APELLIDOS'] = apellidoKey ? String(record[apellidoKey]).trim() : '';
    }
    if (!record['NOMBRES'] || String(record['NOMBRES']).trim() === '') {
      const nombreKey = Object.keys(record).find(k => k.includes('NOMBRE') && String(record[k] || '').trim() !== '');
      record['NOMBRES'] = nombreKey ? String(record[nombreKey]).trim() : '';
    }

    // Enriquecer con metadatos de autoridad (frescura)
    const config = loadAutoridadesConfig();
    const authKey = String(record['AUTORIDAD'] || '').trim().toUpperCase();
    if (config[authKey]) {
      record['SECRETARIO'] = config[authKey].secretario;
      record['CORREO SECRETARIO'] = config[authKey].correo;
    }

    // Bypass del check de finalizado para Notificaciones (el usuario manda)
    // O si viene pre-fetchado (ya validado por el llamador)
    const isNotifica = (actionType === 'Notificación');
    
    if (!preFetchedRecord && !isNotifica) {
      const checkRaw = record['CHECK'];
      const isFinalized = (checkRaw === true || String(checkRaw || '').trim().toUpperCase() === 'TRUE');
      
      if (isFinalized) {
        return { success: false, message: 'La fila ya está marcada como finalizada (CHECK). Destíldala para volver a actuar.' };
      }
    }

    const auxColumns = calculateAuxColumns(record);
    const fullRowObject = { ...record, ...auxColumns };
    let statusToSet = '';
    let actionMessage = '';
    const updateData: any = {};

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
        const now = new Date();
        const firma = activeAutor || (fullRowObject['AUTOR'] || '');
        updateData['GENERA CONTRATO'] = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy") + '_' + firma;
        break;
      }

      case 'Resolución': {
        if (!fullRowObject['REQ'] || String(fullRowObject['REQ']).trim() === '') {
          return { success: false, message: 'Falta el N° de Requerimiento (REQ). Este campo es obligatorio para generar la Resolución.' };
        }
        const resResolucion = generaResolucionTXT(fullRowObject);
        statusToSet = 'Resolución';
        actionMessage = 'Resolución generada exitosamente (' + resResolucion.nombre + ').';
        const now = new Date();
        const firma = activeAutor || (fullRowObject['AUTOR'] || '');
        updateData['GENERA RESOLUCION'] = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy") + '_' + firma;
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
          const now = new Date();
          const firma = activeAutor || (fullRowObject['AUTOR'] || '');
          updateData['BAJA CONTRATO'] = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy") + '_' + firma;
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
          const now = new Date();
          const firma = activeAutor || (fullRowObject['AUTOR'] || '');
          updateData['FECHA NOTIFICA'] = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy HH:mm") + '_' + firma;
        } else {
          return { success: false, message: 'Error en notificación: ' + notificationResult.message };
        }
        break;
      }

      default:
        return { success: false, message: 'Acción no reconocida: ' + actionType };
    }

    if (statusToSet) updateData['ESTADO'] = statusToSet;
    updateData['CHECK'] = true;

    // Actualizamos Firebase. El synchronization mirror se encargará del Sheet.
    firestoreUpdateDocument("Registros", String(recordId), updateData);

    return { success: true, message: actionMessage + ' Estado: ' + (statusToSet || record['ESTADO']) };
  } catch (e) {
    Logger.log('Error en executeAction (Firestore) [' + actionType + '] registro ID ' + recordId + ': ' + e.message);
    return { success: false, message: 'Error al ejecutar "' + actionType + '": ' + e.message };
  }
}

/**
 * Versión determinista del proceso por lotes.
 * Recibe los IDs exactos desde el frontend, sin depender del estado CHECK en Firestore.
 */
function executeBulkActionByIDs(ids: any[], actionType: string, activeAutor?: string) {
  try {
    if (!ids || ids.length === 0) {
      return { success: false, count: 0, errors: [], message: "No se recibieron IDs para procesar." };
    }

    const config = loadAutoridadesConfig();
    let count = 0;
    const errors: string[] = [];

    for (const rawId of ids) {
      const docId = String(rawId).trim();
      try {
        const record = firestoreGetDoc("Registros", docId);
        if (!record) {
          errors.push(docId + ': Registro no encontrado en Firebase.');
          continue;
        }

        const apellidosReal = record['APELLIDOS'] || docId;

        // Enriquecer con datos del secretario (siempre frescos)
        const authKey = String(record['AUTORIDAD'] || '').trim().toUpperCase();
        if (config[authKey]) {
          record['SECRETARIO'] = config[authKey].secretario;
          record['CORREO SECRETARIO'] = config[authKey].correo;
        }

        // Pasar el record pre-cargado: bypasea chequeo de finalizado y evita re-lecturas
        const result = executeAction(docId, actionType, record);
        if (result.success) {
          count++;
        } else {
          errors.push(apellidosReal + ': ' + result.message);
        }
      } catch (e) {
        errors.push(docId + ': ' + e.toString());
      }
    }

    return {
      success: errors.length === 0,
      count: count,
      errors: errors,
      message: 'Se procesaron ' + count + ' de ' + ids.length + ' registros.' + (errors.length > 0 ? ' Errores: ' + errors.length : '')
    };
  } catch (e) {
    Logger.log("Error en executeBulkActionByIDs: " + e.message);
    return { success: false, count: 0, errors: [e.message], message: "Error crítico: " + e.message };
  }
}

function markAsNotified(rowIndex, isNotified) {
  if (isNotified) {
    return executeAction(rowIndex, 'Notificación');
  } else {
    return { success: false, message: "La acción de desmarcar no está implementada directamente aquí." };
  }
}

function executeBulkAction(actionType: string) {
  try {
    const records = firestoreGetAllDocs("Registros");
    if (!records || records.length === 0) {
      return { success: false, message: "No se encontraron registros en Firebase." };
    }

    let count = 0;
    let errors: string[] = [];

    for (const record of records) {
      const isChecked = record['CHECK'];
      const recordId = record['__id'] || record['ID'] || record['ID_REG'];
      const apellidos = record['APELLIDOS'] || 'Sin Apellido';

      // Procesamos solo si CHECK es falso, vacío o string 'FALSE'
      const checkStr = String(isChecked || '').trim().toUpperCase();
      if (isChecked === false || isChecked === '' || isChecked === null || isChecked === undefined || checkStr === 'FALSE') {
        try {
          // Pasamos el record directamente para evitar re-lectura y bugs de latencia/referencia
          const result = executeAction(String(recordId), actionType, record);
          if (result.success) {
            count++;
          } else {
            errors.push(apellidos + ': ' + result.message);
          }
        } catch (e) {
          errors.push(apellidos + ': ' + e.toString());
        }
      }
    }

    return {
      success: errors.length === 0,
      count: count,
      errors: errors,
      message: 'Se procesaron ' + count + ' registros.' + (errors.length > 0 ? ' Errores: ' + errors.length : '')
    };
  } catch (e) {
    Logger.log("Error en executeBulkAction (Firebase): " + e.message);
    return { success: false, message: "Error crítico en proceso por lotes: " + e.message };
  }
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

    const recordId = rowData['ID_REG'] || rowData['ID'];
    
    if (type === 'CONTRATO') {
      folderId = CARPETA_FUSION_CONTRATO;
    } else if (type === 'RESOLUCION') {
      folderId = CARPETA_FUSION_RESOLUCIONES;
    } else if (type === 'NOTA') {
      folderId = CARPETA_NOTAS;
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      
      if (type === 'NOTA') {
        const fileNumber = String(recordId).padStart(3, '0');
        // El patrón es: NNN) AUTORIZA ...
        if (name.startsWith(fileNumber + ")") && name.toLowerCase().includes(".pdf")) {
          return {
            success: true,
            url: file.getUrl().replace('/view', '/preview'),
            name: name
          };
        }
      } else if (type === 'CONTRATO') {
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
  const idCol = headers.indexOf('ID_REG') !== -1 ? headers.indexOf('ID_REG') : headers.indexOf('ID');
  if (idCol === -1) return -1;

  const targetId = String(id).trim();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]).trim() === targetId) {
      return i + 1; // 1-indexed
    }
  }
  return -1;
}
/**
 * ── FUNCIONES DE ADMINISTRACIÓN (SOLO ROL ADMIN) ──
 */

/**
 * Obtiene datos de cualquier colección para las tablas del Panel Admin
 */
function getAdminTableData(collection: string) {
  if (!canManageAutoridadesByCurrentUser()) return { success: false, message: "No autorizado" };
  try {
    const data = firestoreGetAllDocs(collection);
    return { success: true, data: data };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Trae todos los documentos de una colección para el Admin Hub
 */
function adminActionGet(collection: string) {
  try {
    const data = firestoreGetAllDocs(collection);
    return { ok: true, data: data };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

function syncAdminToSheet(collection: string, docId: string, data: any, isDelete: boolean) {
  const map = [
    { sheet: "REGISTROS", col: "Registros", id: "ID_REG" },
    { sheet: "REQUERIMIENTOS_CD", col: "Requerimientos", id: "ID_REQ" },
    { sheet: "DOCUMENTOS_CD", col: "Documentos", id: "DOC_ID" },
    { sheet: "CONTROL", col: "Control", id: "CONTROL" },
    { sheet: "AUTORIDADES", col: "Autoridades", id: "ID_AUT" },
    { sheet: "AUXILIARES", col: "Auxiliares", id: "ID_AUX" },
    { sheet: "USUARIOS", col: "Usuarios", id: "ID_USER" }
  ];

  const config = map.find(m => m.col === collection);
  if (!config) return;

  const ss = getSS();
  if (!ss) return;
  const sheet = ss.getSheetByName(config.sheet);
  if (!sheet) return;

  const d = sheet.getDataRange().getValues();
  if (d.length < 2) return;
  
  const headers = d[0].map(h => String(h).trim().toUpperCase());
  const idIndex = headers.indexOf(config.id);
  if (idIndex === -1) return;

  let rowIndex = -1;
  const targetId = String(docId).trim();

  for (let i = 1; i < d.length; i++) {
    let rowId = String(d[i][idIndex]).trim();
    if (rowId.endsWith(".0")) rowId = rowId.substring(0, rowId.length - 2);
    
    // Comparación robusta entre strings
    if (String(rowId) === String(targetId)) {
      rowIndex = i + 1; // 1-based index para Sheet UI
      break;
    }
  }

  if (isDelete) {
    if (rowIndex > -1) {
      sheet.deleteRow(rowIndex);
    }
  } else {
    // Si no es borrar, es Update o Create
    if (rowIndex > -1) {
      // Actualizar existente celda por celda
      headers.forEach((h, colIdx) => {
        if (data[h] !== undefined) {
           sheet.getRange(rowIndex, colIdx + 1).setValue(data[h]);
        }
      });
    } else {
      // Crear nueva fila
      const newRow = new Array(headers.length).fill("");
      headers.forEach((h, colIdx) => {
        if (data[h] !== undefined) {
          newRow[colIdx] = data[h];
        }
      });
      newRow[idIndex] = targetId;
      sheet.appendRow(newRow);
    }
  }
}

/**
 * Guarda o actualiza un documento en cualquier colección
 */
function adminActionSave(collection: string, docId: string, data: any) {
  try {
    if (!collection || !docId) throw new Error("Parámetros de guardado incompletos.");
    
    // Solo escribimos en Firestore. El synchronization mirror se encargará del Sheet.
    firestoreUpdateDocument(collection, docId, data);
    
    logActivity(`ADMIN: Actualizó/Creó registro en ${collection} (Firebase-First)`);
    return { ok: true };
  } catch (e) {
    Logger.log(`Error en adminActionSave [${collection}]: ` + e.message);
    return { ok: false, message: e.message };
  }
}

/**
 * Elimina un documento de cualquier colección en Firebase
 */
function adminActionDelete(collection: string, docId: string) {
  try {
    firestoreDeleteDocument(collection, docId);
    logActivity(`ADMIN: Eliminó registro ${docId} de ${collection} (Firebase-First)`);
    return { ok: true };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

/**
 * Elimina múltiples documentos de cualquier colección en Firebase
 */
function adminActionBulkDelete(collection: string, docIds: string[]) {
  try {
    let count = 0;
    docIds.forEach(id => {
      firestoreDeleteDocument(collection, id);
      count++;
    });
    logActivity(`ADMIN: Eliminó en lote ${count} registros de ${collection} (Firebase-First)`);
    return { ok: true, count: count };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}


/**
 * Fuerza una migración completa de los Sheets a Firebase para corregir desincronizaciones
 */
function adminActionSyncAll() {
  try {
    // RECTIFICACIÓN: En sistema Firebase-First, la sincronización masiva debe ser Firebase -> Sheet
    syncAllFirebaseToSheets();
    logActivity(`ADMIN: Sincronización Espejo (Firebase -> Sheets) completada`);
    return { ok: true, message: "Sincronización espejo finalizada. Los Sheets ahora reflejan los datos de Firebase." };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}
