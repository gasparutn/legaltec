"use strict";
/**
 * Contratos.ts - Lógica para la generación de contratos de locación en formato Word.
 *
 * Este módulo utiliza una plantilla de Google Docs e inserta los datos del locador
 * y la autoridad correspondiente procesando todos los marcadores <<...>> automáticamente.
 */
/**
 * Genera un contrato en formato Word (.docx) a partir de una plantilla.
 * Reemplaza todos los marcadores estándar y auxiliares.
 *
 * @param {ConRowData} rowData Datos de la fila de la planilla.
 * @returns {{ url: string; nombre: string }} URL y nombre del documento generado.
 */
function generaContratoWord(rowData) {
    Logger.log("=== Iniciando generaContratoWord ===");
    Logger.log("Datos de entrada: " + JSON.stringify(rowData));
    // 1. Validar datos mínimos requeridos
    Logger.log("=== Validando datos en generaContratoWord ===");
    Logger.log("APELLIDOS: '" + (rowData.APELLIDOS || '') + "' (tipo: " + typeof rowData.APELLIDOS + ")");
    Logger.log("NOMBRES: '" + (rowData.NOMBRES || '') + "' (tipo: " + typeof rowData.NOMBRES + ")");
    Logger.log("Todas las claves disponibles: " + Object.keys(rowData).join(" | "));
    // Intentar recuperar APELLIDOS de campos alternativos
    let apellidos = String(rowData.APELLIDOS || '').trim();
    if (!apellidos) {
        const apellidoKeys = Object.keys(rowData).filter(k => k.toUpperCase().includes('APELLIDO') && String(rowData[k]).trim() !== '');
        if (apellidoKeys.length > 0) {
            apellidos = String(rowData[apellidoKeys[0]]).trim();
            Logger.log("APELLIDOS recuperado de campo alternativo: " + apellidoKeys[0]);
        }
    }
    // Intentar recuperar NOMBRES de campos alternativos
    let nombres = String(rowData.NOMBRES || '').trim();
    if (!nombres) {
        const nombreKeys = Object.keys(rowData).filter(k => k.toUpperCase().includes('NOMBRE') && String(rowData[k]).trim() !== '');
        if (nombreKeys.length > 0) {
            nombres = String(rowData[nombreKeys[0]]).trim();
            Logger.log("NOMBRES recuperado de campo alternativo: " + nombreKeys[0]);
        }
    }
    if (!apellidos) {
        const disponiblesApellido = Object.keys(rowData)
            .filter(k => k.toUpperCase().includes('APELLIDO'))
            .map(k => k + "='" + String(rowData[k] || '') + "'")
            .join("; ");
        throw new Error("Faltan APELLIDOS para la generación del contrato.\n" +
            "Campos APELLIDO* disponibles: [" + (disponiblesApellido || "ninguno") + "]\n" +
            "Verifique que la columna APELLIDOS no esté vacía en el registro.");
    }
    if (!nombres) {
        const disponiblesNombre = Object.keys(rowData)
            .filter(k => k.toUpperCase().includes('NOMBRE'))
            .map(k => k + "='" + String(rowData[k] || '') + "'")
            .join("; ");
        throw new Error("Faltan NOMBRES para la generación del contrato.\n" +
            "Campos NOMBRE* disponibles: [" + (disponiblesNombre || "ninguno") + "]\n" +
            "Verifique que la columna NOMBRES no esté vacía en el registro.");
    }
    // Actualizar rowData con los valores recuperados
    rowData.APELLIDOS = apellidos;
    rowData.NOMBRES = nombres;
    const cuotaVal = String(rowData.CUOTA || "").trim();
    const numericCuota = parseInt(cuotaVal || "0", 10);
    if (isNaN(numericCuota) || numericCuota < 1) {
        throw new Error("CUOTA no válida ('" + cuotaVal + "'). Debe ser un número mayor a 0 para generar el contrato.");
    }
    if (!rowData.AUTORIDAD) {
        throw new Error("Falta el campo AUTORIDAD. Es necesario para determinar el formato del contrato.");
    }
    const totalVal = String(rowData.TOTAL || "").trim();
    const numericTotal = parseFloat(totalVal.replace(/[^0-9,-]+/g, "").replace(",", ".")) || 0;
    if (isNaN(numericTotal) || numericTotal <= 0) {
        throw new Error("TOTAL no válido ('" + totalVal + "'). Debe ser un número mayor a 0 para generar el contrato.");
    }
    // 2. Calcular campos auxiliares (AUX1-AUX24) para la redacción
    // Estos campos contienen textos condicionados por género, autoridad y montos.
    Logger.log("Calculando campos auxiliares...");
    const aux = calculateAuxColumns(rowData);
    // 3. Combinar datos originales con auxiliares y metadatos adicionales
    const fullData = { ...rowData, ...aux };
    // Añadir fecha de hoy en formato largo para el documento
    fullData['FECHA'] = fechaEnEspaniol(new Date());
    // Intentar obtener metadatos de autoridad para el log
    try {
        const meta = getAutoridadMetadata(rowData.AUTORIDAD);
        if (meta.nombre) {
            Logger.log("Autoridad identificada: " + meta.nombre + " (" + (meta.tipo || "N/A") + ")");
        }
    }
    catch (e) {
        Logger.log("Aviso: No se pudo obtener metadatos de autoridad: " + e.message);
    }
    // 4. Generar el documento a partir de la copia de la plantilla
    // Formato del nombre: NNN) CONTRATO APELLIDOS-AUTORIDAD_AUTOR donde NNN es el ID con 3 dígitos
    const recordId = String(rowData.ID || '0').padStart(3, '0');
    const filename = `${recordId}) CONTRATO ${rowData.APELLIDOS}-${rowData.AUTORIDAD}_${rowData.AUTOR || ''}`.trim();
    Logger.log("Nombre del archivo a generar: " + filename);
    try {
        // DOC_TEMPLATE_CONTRATO_ID y CARPETA_FUSION_CONTRATO están en constants.ts
        const template = DriveApp.getFileById(DOC_TEMPLATE_CONTRATO_ID);
        const folder = DriveApp.getFolderById(CARPETA_FUSION_CONTRATO);
        Logger.log("Creando copia de la plantilla...");
        const copy = template.makeCopy(filename, folder);
        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();
        Logger.log("Reemplazando marcadores automáticos (replaceFullRow)...");
        // ✅ Reemplaza TODOS los campos (incluyendo AUX) transformando claves a <<KEY>>
        replaceFullRow(body, fullData);
        Logger.log("Guardando y cerrando documento...");
        doc.saveAndClose();
        Logger.log("Convirtiendo a formato Word .docx...");
        // Convertir a Word (.docx) usando la API de Drive con autenticación
        const fileId = copy.getId();
        Logger.log("ID de documento a exportar: " + fileId);
        const exportUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document`;
        Logger.log("URL de exportación: " + exportUrl);
        const params = {
            method: "get",
            headers: {
                "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
            },
            muteHttpExceptions: true,
        };
        const response = UrlFetchApp.fetch(exportUrl, params);
        const responseCode = response.getResponseCode();
        Logger.log("Código de respuesta de exportación: " + responseCode);
        if (responseCode !== 200) {
            const errorContent = response.getContentText();
            throw new Error("Fallo en exportación a .docx. Código HTTP: " + responseCode + ". Respuesta: " + errorContent);
        }
        const docxBlob = response.getBlob();
        if (!docxBlob || docxBlob.getBytes().length === 0) {
            throw new Error("El blob exportado está vacío. La exportación a .docx falló.");
        }
        Logger.log("Blob exportado exitosamente. Tamaño: " + docxBlob.getBytes().length + " bytes");
        Logger.log("Creando archivo en carpeta: " + CARPETA_FUSION_CONTRATO);
        const wordFilename = filename + '.docx';
        Logger.log("Nombre del archivo .docx: " + wordFilename);
        const wordxFile = folder.createFile(docxBlob).setName(wordFilename);
        Logger.log("Archivo creado exitosamente. ID: " + wordxFile.getId());
        // Eliminar la copia de Google Docs original
        copy.setTrashed(true);
        Logger.log("=== Contrato generado exitosamente en formato Word ===");
        logActivity(`Genera Contrato ${rowData.APELLIDOS}_${rowData.AUTORIDAD}`);
        return {
            url: wordxFile.getUrl(),
            nombre: wordxFile.getName()
        };
    }
    catch (err) {
        const errorPrefix = "ERROR en generaContratoWord: ";
        const msg = err.message || err.toString();
        Logger.log(errorPrefix + msg);
        if (err.stack)
            Logger.log("Stack: " + err.stack);
        // Si el error es de UrlFetchApp (probablemente cuotas o permisos de API)
        if (msg.includes("UrlFetchApp") || msg.includes("Access denied")) {
            throw new Error("Error de conexión con la API de Google Drive al exportar a Word. Por favor, verifique los permisos de la aplicación.");
        }
        throw new Error("Fallo al generar el contrato: " + msg);
    }
}
/**
 * Función de test para verificar la preparación de datos del contrato.
 * Puede ejecutarse desde el editor de Apps Script.
 */
function testGeneraContratoLocal() {
    const testData = {
        'SEXO': 'la Sra.',
        'NOMBRES': 'MARIA TEST',
        'APELLIDOS': 'GARCIA',
        'DNI': '12345678',
        'TAREAS': 'servicios de prueba legislativa',
        'FECHA ALTA': '2026-03-01',
        'CUOTA': '3',
        'TOTAL': '150000',
        'AUTORIDAD': 'SERRA'
    };
    try {
        const res = generaContratoWord(testData);
        Logger.log("Resultado Test: " + JSON.stringify(res));
    }
    catch (e) {
        Logger.log("Error en Test: " + e.message);
    }
}
