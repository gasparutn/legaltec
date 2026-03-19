/**
 * Code.gs - Lógica de backend para la gestión de notificaciones
 * Este script maneja la lectura de datos de Google Sheet y el envío de correos.
 */

// Constantes de configuración global
const SPREADSHEET_ID = '1ypKf21sJ2LFx4eXL-u10tVnZ3_WXralrbik3cFqV8zc'; // ¡VERIFICA QUE ESTE SEA TU ID REAL!
const SS = SpreadsheetApp.openById(SPREADSHEET_ID); // Referencia al objeto Spreadsheet

const MAIN_SHEET = SS.getSheetByName("Hoja 1"); // Hoja principal de datos
const HOJA_BAJA_RESOLUCION = SS.getSheetByName("BAJA RES");
const HOJA_NOTIFICADOS = MAIN_SHEET; // Referencia a la misma hoja principal

const CARPETA_HOJA_FUSION = "13MVrXpbIdqtITqs78w0OIJkJvizEl0WK";
const CARPETA_NOTAS = "16gVooDiti5xyVorRv9EgNOCDaoaF_qT9"; // Actualizado por el usuario
const CARPETA_NOTAS_PDF = "1vN_YjSEsItkISD_Wshs8rLtF2zcFdu-D";
const CARPETA_FUSION_CONTRATO = "1OPpYdAc4ueXNLZAnfeH5ne_dAYYOoXId";
const CARPETA_FUSION_RESOLUCIONES = "1Es4FSVTBkFaXzZb33Y7KhqRN6nqT3d-y";
const CARPETA_FUSION_FORMPLANO = "1dzu8Kh84uX5TpPFW6xIp3-bkHaUpAiZA";
const CARPETA_FUSION_BAJA_DOCS = "1HpilrIg_KeVyhfl_B8Lsb2s0ef0t_pDn";

// IDs de las plantillas de documentos
const DOC_TEMPLATE_ALTA_ID = '13jqdi_9iILaD2xjwd2dyqGsZv2Nh1sgAZSsj8CxPC1I';
const DOC_TEMPLATE_ALTA_VICE_ID = '1oab9GdC2fRNATlXoiZd5z0NCMQ7tz4sfAug8H9KT5g0';
const DOC_TEMPLATE_BAJA_ID = '1758J9GrrtrsOIUhohI61MuiJU_dyuPLU3PI4GAOTuwI';
const DOC_TEMPLATE_BAJA_VICE_ID = '1xa2azNXF_PAWvVIKPGHLgfDPJTUhlUEXRTMNEyd_tno';
const DOC_TEMPLATE_DDJJ_ID = '15uh_ZjIqSeqhOBAtHb3X6G39T3V173fl9ebEYvIyDjM';
const DOC_TEMPLATE_PRORR_SENADO_ID = '1WIsIA82HFOiGjeDR8qXbRoHeMawQlu3tYvhq42vXg6M';
const DOC_TEMPLATE_PRORR_VICE_ID = '1ia234xrcWfCS-rR62nlQQswqkMFj0zTr8iHQ3NajZSE';
const DOC_TEMPLATE_RESOLUCION_ID = '18VJg98EStS8566yz42vAXyvPnV5vXm6RA-hpt31pTyg';
const DOC_TEMPLATE_BAJA_RESOLUCION_ID = '14a_8C52VoM-tE_0XqTgNoidxI5121vpOIwB8n9hZtHU';

// Asunto predeterminado para las notificaciones por correo
const ASUNTO_NOTIFICACION = "Notificación de Locación";

function doGet(e) {
  // Esta función es el punto de entrada principal para la aplicación web.
  // Renderiza la plantilla HTML 'Index' o 'Notifications' según el parámetro 'page'.
  if (e.parameter.page === 'Notifications') {
    return HtmlService.createTemplateFromFile('Notifications')
      .evaluate()
      .setTitle('Gestión de Notificaciones')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('PLANILLA') // Título de la página web
      .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // Meta tag para responsividad
  }
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
  const sheet = HOJA_NOTIFICADOS; // Usar la hoja designada para notificaciones

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
    const notificationRow = [];
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
  try {
    const auxSheet = SS.getSheetByName('aux');
    if (!auxSheet) {
      Logger.log("Hoja 'aux' no encontrada.");
      return { success: false, message: "Hoja 'aux' no encontrada.", data: {} };
    }

    const lastRow = auxSheet.getLastRow();
    if (lastRow < 2) return { success: true, data: {} };

    // Leer A2:D (columnas A, B, C, D — índices 0, 1, 2, 3 base-0)
    // El usuario especifica A (Autoridad), C (Secretario), D (Correo)
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

    return { success: true, data: auxMap };
  } catch (e) {
    Logger.log("Error en getAuxData: " + e.message);
    return { success: false, message: e.message, data: {} };
  }
}

function getData() {
  try {
    checkAndAssignIds(); // Asegurar IDs correlativos para cargas manuales
    const sheet = MAIN_SHEET;

    if (!sheet) {
      console.error("Hoja principal no encontrada: Hoja 1");
      return { success: false, message: "Hoja de datos 'Hoja 1' no encontrada en el documento." };
    }

    const range = sheet.getDataRange();
    const allValues = range.getValues();

    if (!allValues || allValues.length === 0) {
      return { success: true, headers: [], data: [] };
    }

    const allHeaders = allValues[0].map(header => String(header || '').trim().toUpperCase());
    const rawData = allValues.slice(1);

    const processedData = rawData.map(row => {
      const rowObject = {};
      allHeaders.forEach((header, index) => {
        if (header) rowObject[header] = row[index];
      });
      // Calcular las columnas auxiliares
      let auxColumns = {};
      try {
        auxColumns = calculateAuxColumns(rowObject);
      } catch (auxErr) {
        console.warn("Error calculando columnas auxiliares para una fila:", auxErr);
      }
      return { ...rowObject, ...auxColumns };
    });

    // Encabezados visibles y ocultos (nombres en mayúsculas para coincidir con el mapeo)
    // El usuario quiere CHECK al final de las visibles
    const visibleClientHeadersNames = [
      'ID', 'AUTORIDAD', 'APELLIDOS', 'NOMBRES', 'DNI', 'TAREAS', 'FECHA ALTA', 'CUOTA', 'TOTAL', 'REQ', 'ESTADO', 'CHECK'
    ];
    const hiddenClientHeadersNames = ['AUTOR', 'DOMICILIO', 'LOCALIDAD', 'FECHA BAJA', 'SECRETARIO', 'CORREO SECRETARIO', 'CORREO LOCADOR', 'FECHA NOTIFICA'];

    let finalDataForClient = processedData.map((rowObject, originalIndex) => {
      // 1. Columnas visibles
      const visiblePart = visibleClientHeadersNames.map(headerName =>
        rowObject[headerName] !== undefined ? rowObject[headerName] : ''
      );
      // 2. Columnas ocultas (al final para la expansión)
      const hiddenPart = hiddenClientHeadersNames.map(headerName =>
        rowObject[headerName] !== undefined ? rowObject[headerName] : ''
      );

      return {
        visiblePart: visiblePart,
        hiddenPart: hiddenPart,
        originalIndex: originalIndex
      };
    });

    // Ordenamiento por FECHA ALTA (DESC) y ID (DESC) para mostrar lo más nuevo arriba
    const altaIndex = visibleClientHeadersNames.indexOf('FECHA ALTA');
    const idIndex = visibleClientHeadersNames.indexOf('ID');

    finalDataForClient.sort((a, b) => {
      try {
        // 1. Prioridad: Fecha Alta (Desc)
        if (altaIndex !== -1) {
          const valA = a.visiblePart[altaIndex];
          const valB = b.visiblePart[altaIndex];
          const dateA = (valA instanceof Date) ? valA : new Date(valA);
          const dateB = (valB instanceof Date) ? valB : new Date(valB);
          
          if (!isNaN(dateA.getTime()) && !isNaN(dateB.getTime())) {
            if (dateA.getTime() !== dateB.getTime()) {
              return dateB - dateA;
            }
          }
        }
        
        // 2. Tie-breaker: ID (Desc)
        if (idIndex !== -1) {
          const idA = parseInt(a.visiblePart[idIndex]) || 0;
          const idB = parseInt(b.visiblePart[idIndex]) || 0;
          return idB - idA;
        }
        
        return 0;
      } catch (e) { return 0; }
    });

    // Preparación final para el cliente con sanitización estricta
    const sanitizedData = finalDataForClient.map(item => {
      // Sanitizar visiblePart y agregar placeholders vacíos para OPC y ACCION
      const sanitizedVisible = item.visiblePart.map(cell => {
        if (cell === null || cell === undefined) return '';
        if (cell instanceof Date) {
          try {
            const y = cell.getFullYear();
            const m = String(cell.getMonth() + 1).padStart(2, '0');
            const d = String(cell.getDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
          } catch (e) { return String(cell); }
        }
        return cell;
      });

      // Sanitizar hiddenPart
      const sanitizedHidden = item.hiddenPart.map(cell => {
        if (cell === null || cell === undefined) return '';
        if (cell instanceof Date) {
          try {
            const y = cell.getFullYear();
            const m = String(cell.getMonth() + 1).padStart(2, '0');
            const d = String(cell.getDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
          } catch (e) { return String(cell); }
        }
        return cell;
      });

      // Estructura final: [visibles..., '', ocultos..., originalIndex] (el '' es para OPC)
      return [...sanitizedVisible, '', ...sanitizedHidden, item.originalIndex];
    });

    return {
      success: true,
      headers: [...visibleClientHeadersNames, 'OPC'],
      hiddenHeaders: hiddenClientHeadersNames,
      data: sanitizedData
    };
  } catch (err) {
    console.error("Error crítico en getData:", err.toString());
    return {
      success: false,
      message: "Fallo crítico al recuperar datos: " + err.toString()
    };
  }
}

function addRow(rowData) {
  const sheet = MAIN_SHEET;

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

    // Limpiar espacios vacíos y formatear DNI
    let cleanData = trimData(rowData);
    if (cleanData['DNI']) {
      cleanData['DNI'] = formatearDNI(cleanData['DNI']);
    }

    // Asignar el nuevo ID y forzar CHECK=true en toda fila nueva
    cleanData['Id'] = nextId; // Usar la capitalización exacta 'Id'
    cleanData['CHECK'] = true; // Al agregar siempre se registra con CHECK tildado

    const newRowValues = [];
    headers.forEach(header => {
      // Usar cleanData[header] para obtener el valor, incluyendo el nuevo Id
      newRowValues.push(cleanData[header] !== undefined ? cleanData[header] : '');
    });

    sheet.appendRow(newRowValues);
    return { success: true, message: "Registro agregado exitosamente con Id: " + nextId };
  }
  catch (e) {
    Logger.log("Error al agregar fila: " + e.message);
    return { success: false, message: "Error al agregar registro: " + e.message };
  }
}

function updateRow(rowIndex, rowData) {
  const sheet = MAIN_SHEET;

  if (!sheet) {
    return { success: false, message: "Hoja principal no encontrada." };
  }

  const sheetRow = rowIndex + 2;

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const updatedValues = [];

    let cleanData = trimData(rowData);
    if (cleanData['DNI']) {
      cleanData['DNI'] = formatearDNI(cleanData['DNI']);
    }

    headers.forEach((header, index) => {
      let originalHeader = String(header || '').trim();
      let upperHeader = originalHeader.toUpperCase();

      let valToSet = cleanData[originalHeader];

      if (valToSet === undefined && cleanData[upperHeader] !== undefined) {
        valToSet = cleanData[upperHeader];
      }

      let fieldWasSent = Object.prototype.hasOwnProperty.call(cleanData, originalHeader) || Object.prototype.hasOwnProperty.call(cleanData, upperHeader);

      let finalVal;
      if (fieldWasSent) {
        finalVal = valToSet;
      } else {
        finalVal = sheet.getRange(sheetRow, index + 1).getValue();
      }

      if (upperHeader === 'DNI' && finalVal) {
        finalVal = formatearDNI(String(finalVal));
      }

      if (typeof finalVal === 'string') {
        finalVal = finalVal.trim();
      }

      updatedValues.push(finalVal);
    });

    sheet.getRange(sheetRow, 1, 1, updatedValues.length).setValues([updatedValues]);
    return { success: true, message: `Registro actualizado exitosamente.` };
  } catch (e) {
    Logger.log("Error al actualizar fila: " + e.message);
    return { success: false, message: "Error al actualizar registro: " + e.message };
  }
}

function deleteRow(rowIndex) {
  const sheet = MAIN_SHEET;

  if (!sheet) {
    return { success: false, message: "Hoja principal no encontrada." };
  }

  const sheetRowToDelete = rowIndex + 2;

  try {
    sheet.deleteRow(sheetRowToDelete);

    const allData = sheet.getDataRange().getValues();
    if (allData.length > 1) { 
      const headers = allData[0].map(header => String(header).trim());
      const idColumnIndex = headers.indexOf('Id');

      if (idColumnIndex === -1) {
        Logger.log("Columna 'Id' no encontrada para reajustar después de la eliminación.");
        return { success: true, message: `Registro en fila ${sheetRowToDelete} eliminado, pero no se pudo reajustar IDs.` };
      }

      const dataToUpdate = allData.slice(1);
      for (let i = 0; i < dataToUpdate.length; i++) {
        dataToUpdate[i][idColumnIndex] = i + 1;
      }
      sheet.getRange(1, 1, allData.length, headers.length).setValues([headers, ...dataToUpdate]);
    }

    return { success: true, message: `Registro en fila ${sheetRowToDelete} eliminado y IDs reajustados.` };
  } catch (e) {
    Logger.log("Error al eliminar fila: " + e.message);
    return { success: false, message: "Error al eliminar registro: " + e.message };
  }
}

function updateCheckValue(rowIndex, value) {
  try {
    const sheet = MAIN_SHEET;
    if (!sheet) return { success: false, message: 'Hoja no encontrada.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());
    const checkColIndex = headers.indexOf('CHECK');
    if (checkColIndex === -1) return { success: false, message: "Columna 'CHECK' no encontrada." };

    const sheetRow = rowIndex + 2;
    sheet.getRange(sheetRow, checkColIndex + 1).setValue(value);
    SpreadsheetApp.flush();
    Logger.log('CHECK actualizado: fila=' + sheetRow + ', valor=' + value);
    return { success: true };
  } catch (e) {
    Logger.log('Error en updateCheckValue: ' + e.message);
    return { success: false, message: e.message };
  }
}

function executeAction(rowIndex, actionType) {
  const sheet = MAIN_SHEET;

  if (!sheet) {
    Logger.log('Hoja principal no encontrada.');
    return { success: false, message: 'Hoja no encontrada.' };
  }

  const sheetRow = rowIndex + 2;
  let statusToSet = '';
  let actionMessage = '';

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());
    const checkboxColumnIndex = headers.indexOf('CHECK') + 1;
    const estadoColumnIndex = headers.indexOf('ESTADO') + 1;

    const rowDataRaw = sheet.getRange(sheetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowObjectRaw = {};
    headers.forEach((upperHeader, index) => { rowObjectRaw[upperHeader] = rowDataRaw[index]; });
    const rowObject = trimData(rowObjectRaw);

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
        const resContrato = generaContratoWord(fullRowObject);
        statusToSet = 'Contrato';
        actionMessage = 'Contrato generado exitosamente (' + resContrato.nombre + ').';
        break;
      }
      case 'Resolución': {
        if (!fullRowObject['REQ'] || String(fullRowObject['REQ']).trim() === '') {
          return { success: false, message: 'Falta el N° de Requerimiento (REQ). Este campo es obligatorio para generar la Resolución.' };
        }
        const resResolucion = generaResolucionTXT(fullRowObject);
        statusToSet = 'Resolución';
        actionMessage = 'Resolución generada exitosamente (' + resResolucion.nombre + ').';
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
            const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
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
    Logger.log('Error en executeAction [' + actionType + '] fila ' + sheetRow + ': ' + e.message + '\n' + e.stack);
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
  const sheet = MAIN_SHEET;
  if (!sheet) return { success: false, message: 'Hoja no encontrada.' };

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const checkIndex = headers.indexOf('CHECK');

  if (checkIndex === -1) {
    return { success: false, message: "Columna 'CHECK' no encontrada." };
  }

  let count = 0;
  let errors = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const isChecked = row[checkIndex];
    const hasContent = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
    if (hasContent && (isChecked === false || isChecked === '' || isChecked === null || isChecked === undefined || isChecked === 'FALSE')) {
      try {
        const result = executeAction(i - 1, actionType);
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
  const sheet = MAIN_SHEET;
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
    return { success: false, message: "Error al buscar archivo: " + e.toString() };
  }
}
