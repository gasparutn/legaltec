
interface ResRowData {
  ID?: string | number;
  'APELLIDOS'?: string;
  'NOMBRES'?: string;
  'DNI'?: string;
  'REQ'?: string;
  'AUTORIDAD'?: string;
  'TOTAL'?: string | number;
  'CUOTA'?: string | number;
  'AUX9'?: string;
  'FECHA ALTA'?: string | Date;
  'FECHA BAJA'?: string | Date;
  'SEXO'?: string;
  'AUTOR'?: string;
  [key: string]: any;
}


function generaResolucionTXT(rowData: ResRowData): { url: string; nombre: string } {
  Logger.log("Iniciando generaResolucionTXT() para: " + rowData['APELLIDOS']);

  try {
    // 1. Validar datos mínimos requeridos
    if (!rowData.APELLIDOS || String(rowData.APELLIDOS).trim() === '') {
      throw new Error("Falta APELLIDOS para la generación de la resolución.");
    }
    if (!rowData.NOMBRES || String(rowData.NOMBRES).trim() === '') {
      throw new Error("Falta NOMBRES para la generación de la resolución.");
    }
    if (!rowData.AUTORIDAD || String(rowData.AUTORIDAD).trim() === '') {
      throw new Error("Falta AUTORIDAD para la generación de la resolución.");
    }
    if (!rowData.REQ || String(rowData.REQ).trim() === '') {
      throw new Error("Falta REQ (Nº de Requerimiento) para la generación de la resolución.");
    }

    const idPlantillaResolucion = DOC_TEMPLATE_RESOLUCION_ID;
    const idCarpetaResolucion = CARPETA_FUSION_RESOLUCIONES;

    // 2. Validar acceso a plantilla y carpeta
    if (!idPlantillaResolucion) throw new Error("ID de plantilla no configurado (DOC_TEMPLATE_RESOLUCION_ID).");
    if (!idCarpetaResolucion) throw new Error("ID de carpeta no configurado (CARPETA_FUSION_RESOLUCIONES).");

    const docBase = DriveApp.getFileById(idPlantillaResolucion);
    const recordId = String(rowData.ID || '0').padStart(3, '0');
    const apellidos = String(rowData['APELLIDOS']).trim();
    const autoridad = String(rowData['AUTORIDAD']).trim();
    const req = String(rowData['REQ']).trim();
    const nombreArchivo = recordId + ") " + req + "_" + apellidos + "_" + autoridad;
    const carpetaPDF = DriveApp.getFolderById(idCarpetaResolucion);

    const nuevoDocFile = docBase.makeCopy(nombreArchivo, carpetaPDF);
    const nuevoDocId = nuevoDocFile.getId();
    const documento = DocumentApp.openById(nuevoDocId);
    const body = documento.getBody();

    // 3. Calcular campos auxiliares (AUX1-AUX24)
    Logger.log("Calculando campos auxiliares para resolución...");
    const aux = calculateAuxColumns(rowData);
    const fullData = { ...rowData, ...aux };

    // Fecha actual para marcador <<FECHA>>
    fullData['FECHA'] = fechaEnEspaniol(new Date());

    replaceFullRow(body, fullData);
    
    const textoPlano = body.getText();
    
    // 4. Validar que el contenido no esté vacío
    if (!textoPlano || textoPlano.trim() === '') {
      throw new Error("El documento resultante está vacío. Verifique que la plantilla contiene marcadores válidos.");
    }
    
    documento.saveAndClose();

    const folderPlano = DriveApp.getFolderById(CARPETA_FUSION_RESOLUCIONES);
    const archivoFinal = folderPlano.createFile(nombreArchivo + ".txt", textoPlano, 'text/plain');
    const archivoFinalUrl = archivoFinal.getUrl();

    DriveApp.getFileById(nuevoDocId).setTrashed(true);

    logActivity(`Genera Resolucion ${apellidos}_${autoridad}`);
    Logger.log("=== Resolución generada exitosamente ===");
    return { url: archivoFinalUrl, nombre: nombreArchivo + ".txt" };
  } catch (err) {
    const errorPrefix = "ERROR en generaResolucionTXT: ";
    const msg = err.message || err.toString();
    Logger.log(errorPrefix + msg);
    if (err.stack) Logger.log("Stack: " + err.stack);
    throw new Error("Fallo al generar la resolución: " + msg);
  }
}

function bajaResolucionTXT(rowData: ResRowData): { success: boolean; url: string; nombre: string; message?: string } {
  Logger.log("Iniciando bajaResolucionTXT() para: " + rowData['APELLIDOS']);

  try {
    // 1. Validar datos mínimos requeridos
    if (!rowData.APELLIDOS || String(rowData.APELLIDOS).trim() === '') {
      throw new Error("Falta APELLIDOS para la baja de resolución.");
    }
    if (!rowData.NOMBRES || String(rowData.NOMBRES).trim() === '') {
      throw new Error("Falta NOMBRES para la baja de resolución.");
    }
    if (!rowData.AUTORIDAD || String(rowData.AUTORIDAD).trim() === '') {
      throw new Error("Falta AUTORIDAD para la baja de resolución.");
    }
    if (!rowData.REQ || String(rowData.REQ).trim() === '') {
      throw new Error("Falta REQ (Nº de Requerimiento) para la baja de resolución.");
    }
    if (!rowData['FECHA BAJA'] || String(rowData['FECHA BAJA']).trim() === '') {
      throw new Error("Falta FECHA BAJA para procesar la baja de resolución.");
    }

    const docTemplateBajaResolucionId = DOC_TEMPLATE_BAJA_RESOLUCION_ID;
    
    // 2. Validar acceso a plantilla y carpeta
    if (!docTemplateBajaResolucionId) throw new Error("ID de plantilla de baja no configurado (DOC_TEMPLATE_BAJA_RESOLUCION_ID).");
    if (!CARPETA_FUSION_RESOLUCIONES) throw new Error("ID de carpeta no configurado (CARPETA_FUSION_RESOLUCIONES).");
    
    const docBase = DriveApp.getFileById(docTemplateBajaResolucionId);
    const recordId = String(rowData.ID || '0').padStart(3, '0');
    const apellidos = String(rowData['APELLIDOS']).trim();
    const autoridad = String(rowData['AUTORIDAD']).trim();
    const req = String(rowData['REQ']).trim();
    const nombreArchivo = recordId + ") BAJA_" + req + "_" + apellidos + "_" + autoridad;
    const carpetaPDF = DriveApp.getFolderById(CARPETA_FUSION_RESOLUCIONES);

    const nuevoDocFile = docBase.makeCopy(nombreArchivo, carpetaPDF);
    const nuevoDocId = nuevoDocFile.getId();
    const documento = DocumentApp.openById(nuevoDocId);
    const body = documento.getBody();

    // 3. Calcular campos auxiliares (AUX1-AUX24)
    Logger.log("Calculando campos auxiliares para baja de resolución...");
    const aux = calculateAuxColumns(rowData);
    const fullData = { ...rowData, ...aux };

    // Agregar fecha de hoy si no está
    if (!fullData['FECHA']) {
      fullData['FECHA'] = fechaEnEspaniol(new Date());
    }

    replaceFullRow(body, fullData);

    const textoPlano = body.getText();
    
    // 4. Validar que el contenido no esté vacío
    if (!textoPlano || textoPlano.trim() === '') {
      throw new Error("El documento de baja resultante está vacío. Verifique que la plantilla contiene marcadores válidos.");
    }
    
    documento.saveAndClose();

    const folderPlano = DriveApp.getFolderById(CARPETA_FUSION_RESOLUCIONES);
    const archivoFinal = folderPlano.createFile(nombreArchivo + ".txt", textoPlano, 'text/plain');

    DriveApp.getFileById(nuevoDocId).setTrashed(true);
    
    logActivity(`Genera Baja Resolucion ${apellidos}_${autoridad}`);
    Logger.log("=== Baja de Resolución generada exitosamente ===");

    return { success: true, url: archivoFinal.getUrl(), nombre: nombreArchivo + ".txt" };
  } catch (err) {
    const errorPrefix = "ERROR en bajaResolucionTXT: ";
    const msg = err.message || err.toString();
    Logger.log(errorPrefix + msg);
    if (err.stack) Logger.log("Stack: " + err.stack);
    return { success: false, url: '', nombre: '', message: "Fallo al generar baja de resolución: " + msg };
  }
}