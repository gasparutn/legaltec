// --- Funciones de acción específicas (Notas, Resolución, Contrato, Notificación) ---

function autoNota_ALTA(rowData) {
  // Wrapper legacy para compatibilidad con llamadas directas (ej. desde flujos globales)
  const docActualId = (rowData['AUTORIDAD'] && String(rowData['AUTORIDAD']).toUpperCase() === 'VICE')
    ? DOC_TEMPLATE_ALTA_VICE_ID : DOC_TEMPLATE_ALTA_ID;
  return autoNota_ALTA_TEMPLATE(rowData, docActualId);
}

function autoNota_ALTA_TEMPLATE(rowData, templateId) {
  Logger.log("Función autoNota_ALTA_TEMPLATE() ejecutada. Template: " + templateId);
  const docActualId = templateId;

  try {
    const docTemplate = DriveApp.getFileById(docActualId);

    // Crear una copia del documento de plantilla
    const newDocName = "Nota_Alta_" + (rowData['APELLIDOS'] || 'N/A') + "_" + (rowData['AUTORIDAD'] || 'N/A') + "_" + (rowData['AUTOR'] || '');
    const newDoc = docTemplate.makeCopy(newDocName);
    const documento = DocumentApp.openById(newDoc.getId());
    const body = documento.getBody();

    // Reemplazar marcadores de posición con datos de la fila
    // Asegúrate de que los nombres de las columnas en rowData (con su capitalización exacta)
    // y que los marcadores de posición en tu plantilla (ej. <<GENERO>>) coincidan.

    body.replaceText("<<SEXO>>", (rowData['SEXO'] || '').trim());
    body.replaceText("<<GENERO>>", (rowData['SEXO'] || '').trim());
    body.replaceText("<<FECHA>>", ObtenerTextofecha() || ''); // FECHA ACTUAL
    body.replaceText("<<NOMBRES>>", (rowData['NOMBRES'] || '').trim());
    body.replaceText("<<APELLIDOS>>", (rowData['APELLIDOS'] || '').trim());
    body.replaceText("<<DNI>>", formatearDNI(String(rowData['DNI'] || '').trim()));
    body.replaceText("<<CUOTA>>", rowData['CUOTA'] || ''); // Usar 'CUOTA'
    body.replaceText("<<MESLETRA>>", rowData['NUM_LETRA'] || ''); // cuotas en letra (ej: cinco)
    body.replaceText("<<MENSUAL_LETRA>>", rowData['NUM_A_LET_CUOTA'] || '');

    // Formatear CUOTA_MENSUAL a número AR (por ej: $150.000,00)
    let cuotaMensualStr = '';
    if (rowData['TOTAL_DIV_CUO'] !== undefined && rowData['TOTAL_DIV_CUO'] !== '') {
      cuotaMensualStr = formatearDinero(rowData['TOTAL_DIV_CUO']);
    }
    body.replaceText("<<CUOTA_MENSUAL>>", cuotaMensualStr);

    let cantCuotas = parseInt(rowData['CUOTA'], 10);
    let leMeses = (cantCuotas === 1) ? "mes" : "meses";
    let leCuotas = (cantCuotas === 1) ? "cuota" : "cuotas";
    let leCuotas2 = (cantCuotas === 1) ? "una" : (rowData['NUM_LETRA'] || "");
    body.replaceText("<<LEMESES>>", leMeses);
    body.replaceText("<<LECUOTAS>>", leCuotas);
    body.replaceText("<<LECUOTAS2>>", leCuotas2);

    body.replaceText("<<TAREAS>>", rowData['TAREAS'] || '');
    body.replaceText("<<FECHA ALTA>>", formatearFechaTexto(rowData['FECHA ALTA']));
    body.replaceText("<<DOMICILIO>>", rowData['DOMICILIO'] || ''); // Usar 'DOMICILIO'
    body.replaceText("<<LOCALIDAD>>", rowData['LOCALIDAD'] || '');
    body.replaceText("<<TOTAL>>", formatearDinero(rowData['TOTAL'])); // Usar 'TOTAL'
    body.replaceText("<<TOTALETRA>>", rowData['NUM_A_LETRAS'] || '');
    body.replaceText("<<MENLETRA>>", rowData['NUM_A_LET_CUOTA'] || '');
    body.replaceText("<<elSR_delSR>>", rowData['delSr_delaSra'] || ''); // Corregido: Usar 'delSr_delaSra'
    body.replaceText("<<el_denominado_a>>", rowData['el_denominado_a'] || '');
    body.replaceText("<<LOCADOR_AR>>", rowData['LOCADOR_AR'] || '');
    body.replaceText("<<MONTOMENSUAL>>", rowData['TOTAL_DIV_CUO'] || '');
    body.replaceText("<<NUM_LETRA>>", rowData['NUM_LETRA'] || '');
    body.replaceText("<<MES_ES>>", rowData['MES_ES'] || '');
    body.replaceText("<<NUMLETRA>>", rowData['NUMLETRA'] || '');
    body.replaceText("<<CUOTA_LE>>", rowData['CUOTA_LE'] || '');
    body.replaceText("<<FECHAHOY>>", rowData['FECHAHOY'] ? Utilities.formatDate(new Date(rowData['FECHAHOY']), Session.getScriptTimeZone(), "dd/MM/yyyy") : '');
    body.replaceText("<<AUTORIDAD>>", rowData['AUTORIDAD_COMPLETA'] || rowData['AUTORIDAD'] || ''); // Usando AUTORIDAD para la autoridad

    documento.saveAndClose();

    // Mover el nuevo documento a la carpeta de notas
    const folderNotas = DriveApp.getFolderById(CARPETA_NOTAS);
    newDoc.moveTo(folderNotas);

    // Crear PDF y moverlo a la carpeta de PDFs
    const folderNotasPdf = DriveApp.getFolderById(CARPETA_NOTAS_PDF);
    const docPdf = newDoc.getAs('application/pdf');
    docPdf.setName(newDoc.getName() + ".pdf");
    folderNotasPdf.createFile(docPdf);

    // Opcional: Eliminar el documento de Google Docs original si solo necesitas el PDF
    // newDoc.setTrashed(true);

    Logger.log(`Documento "${newDoc.getName()}" y PDF creados exitosamente.`);
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

    // Crear una copia del documento de plantilla
    const newDocName = "Nota_Baja_" + (rowData['APELLIDOS'] || 'N/A') + "_" + (rowData['AUTORIDAD'] || 'N/A') + "_" + (rowData['AUTOR'] || '');
    const newDoc = docTemplate.makeCopy(newDocName);
    const documento = DocumentApp.openById(newDoc.getId());
    const body = documento.getBody();

    // Reemplazar marcadores de posición con datos de la fila
    // Asegúrate de que los nombres de las columnas en rowData (con su capitalización exacta)
    // y que los marcadores de posición en tu plantilla (ej. <<GENERO>>) coincidan.

    body.replaceText("<<SEXO>>", (rowData['delSr_delaSra'] || '').trim()); // Corregido: Usar 'delSr_delaSra'
    body.replaceText("<<FECHA>>", ObtenerTextofecha() || ''); // FECHA ACTUAL
    body.replaceText("<<NOMBRES>>", (rowData['NOMBRES'] || '').trim());
    body.replaceText("<<APELLIDOS>>", (rowData['APELLIDOS'] || '').trim());
    body.replaceText("<<DNI>>", formatearDNI(String(rowData['DNI'] || '').trim()));
    body.replaceText("<<FECHA_BAJA>>", rowData['FECHA BAJA'] || ''); // Asumiendo columna 'FECHA BAJA'
    body.replaceText("<<MESES>>", rowData['CUOTA'] || ''); // Usar 'CUOTA'
    body.replaceText("<<DOMICILIO>>", rowData['DOMICILIO'] || ''); // Usar 'DOMICILIO'
    body.replaceText("<<LOCALIDAD>>", rowData['LOCALIDAD'] || '');
    // body.replaceText("<<AUTORIDAD>>", rowData['AUTORIDAD_COMPLETA'] || rowData['AUTORIDAD'] || ''); // Usando AUTORIDAD para la autoridad

    documento.saveAndClose();

    // Mover el nuevo documento a la carpeta de notas
    const folderNotas = DriveApp.getFolderById(CARPETA_NOTAS);
    newDoc.moveTo(folderNotas);

    // Crear PDF y moverlo a la carpeta de PDFs
    const folderNotasPdf = DriveApp.getFolderById(CARPETA_NOTAS_PDF);
    const docPdf = newDoc.getAs('application/pdf');
    docPdf.setName(newDoc.getName() + ".pdf");
    folderNotasPdf.createFile(docPdf);

    // Opcional: Eliminar el documento de Google Docs original si solo necesitas el PDF
    // newDoc.setTrashed(true);

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

  //obtener fecha en formato texto
  switch (mes) {

    case 0: mes = "enero"; break;
    case 1: mes = "febrero"; break;
    case 2: mes = "marzo"; break;
    case 3: mes = "abril"; break;
    case 4: mes = "mayo"; break;
    case 5: mes = "junio"; break;
    case 6: mes = "julio"; break;
    case 7: mes = "agosto"; break;
    case 8: mes = "septiembre"; break;
    case 9: mes = "octubre"; break;
    case 10: mes = "noviembre"; break;
    case 11: mes = "diciembre"; break;
  }
  return "Mendoza, " + mes + " del " + anyo;
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

function generarNotaDesdeModal(formData) {
  try {
    const tipoNota = formData['TIPO_NOTA_SELECT'];
    const sexoNota = formData['SEXO_NOTA_SELECT'];

    if (!tipoNota) return { success: false, message: 'Tipo de nota no especificado.' };

    let sexoMapped = '';
    if (sexoNota === 'Masculino') {
      sexoMapped = 'el Sr.';
    } else if (sexoNota === 'Femenino') {
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

    body.replaceText("<<FECHA>>", ObtenerTextofecha() || '');
    body.replaceText("<<NOMBRES>>", (rowData['NOMBRES'] || '').trim());
    body.replaceText("<<APELLIDOS>>", (rowData['APELLIDOS'] || '').trim());
    if (rowData['DNI']) body.replaceText("<<DNI>>", formatearDNI(String(rowData['DNI'] || '').trim()));
    body.replaceText("<<DOMICILIO>>", rowData['DOMICILIO'] || '');
    body.replaceText("<<LOCALIDAD>>", rowData['LOCALIDAD'] || '');
    body.replaceText("<<CORREO LOCADOR>>", (rowData['CORREO LOCADOR'] || '').trim());

    documento.saveAndClose();

    const folderNotas = DriveApp.getFolderById(CARPETA_NOTAS);
    newDoc.moveTo(folderNotas);

    const folderNotasPdf = DriveApp.getFolderById(CARPETA_NOTAS_PDF);
    const docPdf = newDoc.getAs('application/pdf');
    docPdf.setName(newDoc.getName() + ".pdf");
    folderNotasPdf.createFile(docPdf);

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

    body.replaceText("<<FECHA>>", ObtenerTextofecha() || '');
    body.replaceText("<<SEXO>>", (rowData['SEXO'] || '').trim());
    body.replaceText("<<GENERO>>", (rowData['SEXO'] || '').trim());
    body.replaceText("<<NOMBRES>>", (rowData['NOMBRES'] || '').trim());
    body.replaceText("<<APELLIDOS>>", (rowData['APELLIDOS'] || '').trim());
    body.replaceText("<<DNI>>", formatearDNI(String(rowData['DNI'] || '').trim()));
    body.replaceText("<<DOMICILIO>>", rowData['DOMICILIO'] || '');
    body.replaceText("<<LOCALIDAD>>", rowData['LOCALIDAD'] || '');
    body.replaceText("<<CUOTA>>", rowData['CUOTA'] || '');
    body.replaceText("<<MESLETRA>>", rowData['NUM_LETRA'] || '');

    // LECUOTAS2: "una" si es 1 cuota, sino el número en letras
    let cantCuotas = parseInt(rowData['CUOTA'], 10);
    let leCuotas2 = (cantCuotas === 1) ? "una" : (rowData['NUM_LETRA'] || '');
    body.replaceText("<<LECUOTAS2>>", leCuotas2);

    body.replaceText("<<AUTORIDAD>>", rowData['AUTORIDAD_COMPLETA'] || rowData['AUTORIDAD'] || '');

    documento.saveAndClose();

    const folderNotas = DriveApp.getFolderById(CARPETA_NOTAS);
    newDoc.moveTo(folderNotas);

    const folderNotasPdf = DriveApp.getFolderById(CARPETA_NOTAS_PDF);
    const docPdf = newDoc.getAs('application/pdf');
    docPdf.setName(newDoc.getName() + ".pdf");
    folderNotasPdf.createFile(docPdf);

    return { success: true, message: `Prórroga ${tipoLabel} para ${rowData['NOMBRES']} generada.` };

  } catch (e) {
    Logger.log("Error en autoNota_PRORR(): " + e.message);
    return { success: false, message: "Error al generar Prórroga: " + e.message };
  }
}
