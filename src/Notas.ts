// --- Funciones de acción específicas (Notas, Resolución, Contrato, Notificación) ---

function autoNota_ALTA(rowData) {
  // Wrapper legacy para compatibilidad con llamadas directas (ej. desde flujos globales)
  const docActualId = (rowData['AUTORIDAD'] && String(rowData['AUTORIDAD']).toUpperCase() === 'VICE')
    ? DOC_TEMPLATE_ALTA_VICE_ID : DOC_TEMPLATE_ALTA_ID;
  return autoNota_ALTA_TEMPLATE(rowData, docActualId);
}

function autoNota_ALTA_TEMPLATE(rowData, templateId) {
  Logger.log("Función autoNota_ALTA_TEMPLATE() ejecutada. Template: " + templateId);
  try {
    const docTemplate = DriveApp.getFileById(templateId);
    const newDocName = "Nota_Alta_" + (rowData['APELLIDOS'] || 'N/A') + "_" + (rowData['AUTORIDAD'] || 'N/A') + "_" + (rowData['AUTOR'] || '');
    const newDoc = docTemplate.makeCopy(newDocName);
    const documento = DocumentApp.openById(newDoc.getId());
    const body = documento.getBody();

    // Fecha actual para marcador <<FECHA>>
    rowData['FECHA'] = ObtenerTextofecha();

    replaceFullRow(body, rowData);

    documento.saveAndClose();

    const folderNotas = DriveApp.getFolderById(CARPETA_NOTAS);
    newDoc.moveTo(folderNotas);

    // Solo se conserva el Doc por pedido del usuario

    return { success: true, message: `Nota Alta para ${rowData['NOMBRES']} generada.` };

  } catch (e) {
    Logger.log("Error en autoNota_ALTA_TEMPLATE(): " + e.message);
    return { success: false, message: "Error al generar Nota Alta: " + e.message };
  }
}

function autoNota_BAJA(rowData) {
  // Wrapper legacy para compatibilidad con llamadas directas
  const docActualId = (rowData['AUTORIDAD'] && String(rowData['AUTORIDAD']).toUpperCase() === 'VICE')
    ? DOC_TEMPLATE_BAJA_VICE_ID : DOC_TEMPLATE_BAJA_ID;
  return autoNota_BAJA_TEMPLATE(rowData, docActualId);
}

function autoNota_BAJA_TEMPLATE(rowData, templateId) {
  Logger.log("Función autoNota_BAJA_TEMPLATE() ejecutada. Template: " + templateId);
  try {
    const docTemplate = DriveApp.getFileById(templateId);
    const newDocName = "Nota_Baja_" + (rowData['APELLIDOS'] || 'N/A') + "_" + (rowData['AUTORIDAD'] || 'N/A') + "_" + (rowData['AUTOR'] || '');
    const newDoc = docTemplate.makeCopy(newDocName);
    const documento = DocumentApp.openById(newDoc.getId());
    const body = documento.getBody();

    // Fecha actual para marcador <<FECHA>>
    rowData['FECHA'] = ObtenerTextofecha();

    replaceFullRow(body, rowData);

    documento.saveAndClose();
    logActivity(`Genera Nota Baja ${rowData['APELLIDOS']}_${rowData['AUTORIDAD']}`);

    const folderNotas = DriveApp.getFolderById(CARPETA_NOTAS);
    newDoc.moveTo(folderNotas);

    // Solo se conserva el Doc por pedido del usuario

    Logger.log(`Documento "${newDoc.getName()}" y PDF creados exitosamente.`);
    return { success: true, message: `Nota Baja para ${rowData['NOMBRES']} generada.` };

  } catch (e) {
    Logger.log("Error en autoNota_BAJA_TEMPLATE(): " + e.message);
    return { success: false, message: "Error al generar Nota Baja: " + e.message };
  }
}

function autoNota_PRORR_TEMPLATE(rowData, templateId) {
  return autoNota_PRORR(rowData, templateId);
}

function ObtenerTextofecha() {
  //obetner fecha actual
  var fecha = new Date();
  var mes = fecha.getMonth();
  var dia = fecha.getDate();
  var anyo = fecha.getFullYear();

  let mesNombre = "";
  //obtener fecha en formato texto
  switch (mes) {

    case 0: mesNombre = "enero"; break;
    case 1: mesNombre = "febrero"; break;
    case 2: mesNombre = "marzo"; break;
    case 3: mesNombre = "abril"; break;
    case 4: mesNombre = "mayo"; break;
    case 5: mesNombre = "junio"; break;
    case 6: mesNombre = "julio"; break;
    case 7: mesNombre = "agosto"; break;
    case 8: mesNombre = "septiembre"; break;
    case 9: mesNombre = "octubre"; break;
    case 10: mesNombre = "noviembre"; break;
    case 11: mesNombre = "diciembre"; break;
  }
  return "Mendoza, " + mesNombre + " del " + anyo;
}

function formatearFechaTexto(fechaDateOrString) {
  if (!fechaDateOrString) return '';
  try {
    let dateObj;
    if (typeof fechaDateOrString === 'string') {
      const parts = fechaDateOrString.split('-');
      // if from <input type="date"> it will be "YYYY-MM-DD"
      if (parts.length >= 3) {
        dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2].split('T')[0]));
      } else {
        dateObj = new Date(fechaDateOrString);
      }
    } else {
      dateObj = fechaDateOrString;
    }

    if (isNaN(dateObj.getTime())) return String(fechaDateOrString);

    var mesNames = [
      "enero", "febrero", "marzo", "abril", "mayo", "junio",
      "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ];
    var mes = mesNames[dateObj.getMonth()];
    var dia = dateObj.getDate();
    var anyo = dateObj.getFullYear();

    return dia + " de " + mes + " de " + anyo;
  } catch (e) {
    return String(fechaDateOrString);
  }
}

function formatearDinero(valor) {
  if (valor === undefined || valor === null || valor === '') return '';
  let num = typeof valor === 'string' ? parseFloat(valor.replace(/[^0-9,-]+/g, "").replace(",", ".")) : parseFloat(valor);
  if (isNaN(num)) return String(valor);

  let parts = num.toFixed(2).split(".");
  let enteros = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  let decimales = parts[1];

  if (decimales === "00") {
    return "$" + enteros;
  }
  return "$" + enteros + "," + decimales;
}

function generarNotaDesdeModal(formDataRaw) {
  try {
    // Normalizar llaves a MAYÚSCULAS para consistencia con calculateAuxColumns
    const formData: {[key: string]: any} = {};
    for (let key in formDataRaw) {
      formData[key.trim().toUpperCase()] = formDataRaw[key];
    }

    const tipoNota = formData['TIPO_NOTA_SELECT'];
    const sexoNota = formData['SEXO_NOTA_SELECT'];

    if (!tipoNota) return { success: false, message: 'Tipo de nota no especificado.' };

    let sexoMapped = '';
    const sxUpper = String(sexoNota).toUpperCase().trim();

    if (sexoNota === 'Masculino' || sxUpper === 'EL SR.') {
      sexoMapped = 'el Sr.';
    } else if (sexoNota === 'Femenino' || sxUpper === 'LA SRA.') {
      sexoMapped = 'la Sra.';
    } else {
      sexoMapped = 'el/la Sr./Sra.';
    }

    formData['SEXO'] = sexoMapped;
    formData['delSr_delaSra'] = sexoMapped;

    let tipoNotaFinal = tipoNota;

    // La validación de autoridad permitida ya ocurrió en el cliente.
    // Determinar AUTORIDAD para naming
    if (tipoNotaFinal.includes('Vice')) {
      formData['AUTORIDAD'] = 'VICE';
    } else if (tipoNotaFinal.includes('Senado')) {
      formData['AUTORIDAD'] = 'SENADO';
    } else {
      formData['AUTORIDAD'] = ''; // DDJJ no usa autoridad
    }

    // Calcular auxiliares
    let fullData = { ...formData };
    try {
      const aux = calculateAuxColumns(formData);
      fullData = { ...formData, ...aux };
    } catch (e) {
      Logger.log("Advertencia generando nota desde modal (aux columns): " + e.message);
    }

    if (tipoNotaFinal === 'Nota Alta Senado') {
      return autoNota_ALTA_TEMPLATE(fullData, DOC_TEMPLATE_ALTA_ID);
    } else if (tipoNotaFinal === 'Nota Alta Vice') {
      return autoNota_ALTA_TEMPLATE(fullData, DOC_TEMPLATE_ALTA_VICE_ID);
    } else if (tipoNotaFinal === 'Nota Baja Senado') {
      return autoNota_BAJA_TEMPLATE(fullData, DOC_TEMPLATE_BAJA_ID);
    } else if (tipoNotaFinal === 'Nota Baja Vice') {
      return autoNota_BAJA_TEMPLATE(fullData, DOC_TEMPLATE_BAJA_VICE_ID);
    } else if (tipoNotaFinal === 'Prórroga Senado') {
      return autoNota_PRORR_TEMPLATE(fullData, DOC_TEMPLATE_PRORR_SENADO_ID);
    } else if (tipoNotaFinal === 'Prórroga Vice') {
      return autoNota_PRORR_TEMPLATE(fullData, DOC_TEMPLATE_PRORR_VICE_ID);
    } else if (tipoNotaFinal === 'Nota DDJJ') {
      return autoNota_DDJJ(fullData);
    } else {
      return { success: false, message: 'El tipo de nota seleccionado aún no tiene implementación de plantilla (' + tipoNotaFinal + ').' };
    }
  } catch (e) {
    Logger.log("Error en generarNotaDesdeModal: " + e.message);
    return { success: false, message: 'Error interno: ' + e.message };
  }
}

function autoNota_DDJJ(rowData) {
  try {
    const docTemplate = DriveApp.getFileById(DOC_TEMPLATE_DDJJ_ID);
    const newDocName = "Nota_DDJJ_" + (rowData['APELLIDOS'] || 'N/A') + "_" + (rowData['AUTOR'] || '');
    const newDoc = docTemplate.makeCopy(newDocName);
    const documento = DocumentApp.openById(newDoc.getId());
    const body = documento.getBody();

    // Fecha actual para marcador <<FECHA>>
    rowData['FECHA'] = ObtenerTextofecha();

    // ✅ Usar replaceFullRow para reemplazar todos los campos automáticamente
    replaceFullRow(body, rowData);

    documento.saveAndClose();

    const folderNotas = DriveApp.getFolderById(CARPETA_NOTAS);
    newDoc.moveTo(folderNotas);

    // Solo se conserva el Doc por pedido del usuario

    return { success: true, message: `Nota DDJJ para ${rowData['NOMBRES']} generada.` };

  } catch (e) {
    Logger.log("Error en autoNota_DDJJ(): " + e.message);
    return { success: false, message: "Error al generar Nota DDJJ: " + e.message };
  }
}

function autoNota_PRORR(rowData, templateIdOverride) {
  try {
    // Si viene un templateId del dispatcher, usarlo directamente; si no, inferir de AUTORIDAD (compatibilidad)
    const esVice = templateIdOverride ? templateIdOverride === DOC_TEMPLATE_PRORR_VICE_ID : (rowData['AUTORIDAD'] === 'VICE');
    const templateId = templateIdOverride || (esVice ? DOC_TEMPLATE_PRORR_VICE_ID : DOC_TEMPLATE_PRORR_SENADO_ID);
    const tipoLabel = esVice ? 'Vice' : 'Senado';

    const docTemplate = DriveApp.getFileById(templateId);
    const newDocName = "Nota_Prorr_" + (rowData['APELLIDOS'] || 'N/A') + "_" + (rowData['AUTORIDAD'] || 'N/A') + "_" + (rowData['AUTOR'] || '');
    const newDoc = docTemplate.makeCopy(newDocName);
    const documento = DocumentApp.openById(newDoc.getId());
    const body = documento.getBody();

    // Fecha actual para marcador <<FECHA>>
    rowData['FECHA'] = ObtenerTextofecha();

    // LECUOTAS2: "una" si es 1 cuota, sino el número en letras (específico para esta plantilla)
    let cantCuotas = parseInt(rowData['CUOTA'], 10);
    let leCuotas2 = (cantCuotas === 1) ? "una" : (rowData['NUM_LETRA'] || '');
    rowData['LECUOTAS2'] = leCuotas2;

    // ✅ Usar replaceFullRow para reemplazar todos los campos automáticamente
    replaceFullRow(body, rowData);

    documento.saveAndClose();

    const folderNotas = DriveApp.getFolderById(CARPETA_NOTAS);
    newDoc.moveTo(folderNotas);

    // Solo se conserva el Doc por pedido del usuario

    return { success: true, message: `Prórroga ${tipoLabel} para ${rowData['NOMBRES']} generada.` };

  } catch (e) {
    Logger.log("Error en autoNota_PRORR(): " + e.message);
    return { success: false, message: "Error al generar Prórroga: " + e.message };
  }
}
