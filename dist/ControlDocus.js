"use strict";
// ─── Constantes del módulo ────────────────────────────────────────────────────
const CD_HOJA_CONFIG = 'config';
const CD_HOJA_REQUERIMIENTOS = 'REQUERIMIENTOS_CD';
const CD_HOJA_DOCUMENTOS = 'DOCUMENTOS_CD';
const CD_PROP_ROL = 'CD_ROL';
const CD_PROP_USUARIO_ID = 'CD_USUARIO_ID';
const CD_PROP_NOMBRE = 'CD_NOMBRE';
// Los 15 requisitos fijos (nombre corto para la tabla)
const CD_REQUISITOS = [
    'Nota solicitud Contrato de Locación',
    'Copia DNI',
    'Constancia últimos 2 votos emitidos / justificación',
    'Constancia de CUIL / Antecedentes institucionales',
    'Declaración Jurada sin cargo en dependencia estatal',
    'Declaración Jurada Horas en Docencia (si aplica)',
    'Constancia Deudores Alimentarios Morosos (online)',
    'Certificativa Negativa ANSES',
    'Constancia AFIP impuestos activos (monotrib/IVA/etc.)',
    'Constancia AFIP Ingresos Brutos',
    'Fotocopia DNI / Título de estudios (si corresponde)',
    'Partida de nacimiento actualizada',
    'Certificado NO Concursado / Quebrado',
    'Certificado NO Concursado firmado online',
    'Correo electrónico + celular del Secretario referente'
];
// ─── Autenticación ─────────────────────────────────────────────────────────────
/**
 * Valida usuario y contraseña contra la hoja 'config'.
 * Guarda la sesión en PropertiesService del script.
 */
function loginCD(usuario, pass) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hoja = ss.getSheetByName(CD_HOJA_CONFIG);
        if (!hoja)
            return { ok: false, message: 'Hoja de configuración no encontrada.' };
        const datos = hoja.getDataRange().getValues();
        // Encabezados: ID | NOMBRE | USUARIO | PASS | ROL
        for (let i = 1; i < datos.length; i++) {
            const [id, nombre, usr, pw, rol] = datos[i];
            if (String(usr).trim().toLowerCase() === usuario.trim().toLowerCase()
                && String(pw).trim() === pass.trim()) {
                const props = PropertiesService.getUserProperties();
                props.setProperty(CD_PROP_ROL, String(rol).trim().toLowerCase());
                props.setProperty(CD_PROP_USUARIO_ID, String(id).trim());
                props.setProperty(CD_PROP_NOMBRE, String(nombre).trim());
                return { ok: true, rol: String(rol).trim().toLowerCase(), nombre: String(nombre).trim(), uid: String(id).trim() };
            }
        }
        return { ok: false, message: 'Usuario o contraseña incorrectos.' };
    }
    catch (e) {
        return { ok: false, message: 'Error en login: ' + e.message };
    }
}
/** Retorna el rol y la sesión actual del usuario logueado. */
function getSesionCD() {
    const props = PropertiesService.getUserProperties();
    return {
        rol: props.getProperty(CD_PROP_ROL) || '',
        usuarioId: props.getProperty(CD_PROP_USUARIO_ID) || '',
        nombre: props.getProperty(CD_PROP_NOMBRE) || ''
    };
}
/** Cierra la sesión del módulo Control Docus. */
function logoutCD() {
    const props = PropertiesService.getUserProperties();
    props.deleteProperty(CD_PROP_ROL);
    props.deleteProperty(CD_PROP_USUARIO_ID);
    props.deleteProperty(CD_PROP_NOMBRE);
    return { ok: true };
}
// ─── Inicialización de hojas ───────────────────────────────────────────────────
/** Crea las hojas REQUERIMIENTOS_CD y DOCUMENTOS_CD si no existen. */
function initHojasCD() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        // REQUERIMIENTOS_CD
        if (!ss.getSheetByName(CD_HOJA_REQUERIMIENTOS)) {
            const h = ss.insertSheet(CD_HOJA_REQUERIMIENTOS);
            h.getRange(1, 1, 1, 8).setValues([[
                    'ID_REQ', 'ID_USUARIO', 'APELLIDO_NOMBRE', 'AUTORIDAD',
                    'FECHA_APERTURA', 'ESTADO', 'ID_PDF_DRIVE', 'ID_CARPETA_DRIVE'
                ]]);
        }
        // DOCUMENTOS_CD
        if (!ss.getSheetByName(CD_HOJA_DOCUMENTOS)) {
            const h = ss.insertSheet(CD_HOJA_DOCUMENTOS);
            h.getRange(1, 1, 1, 9).setValues([[
                    'ID_DOC', 'ID_REQ', 'NUM_REQ', 'ESTADO', 'URL_DRIVE',
                    'FECHA_CARGA', 'FECHA_APROBACION', 'OBS_RECHAZO', 'APROBADO_POR'
                ]]);
        }
        return { ok: true, message: 'Hojas inicializadas correctamente.' };
    }
    catch (e) {
        return { ok: false, message: 'Error al inicializar hojas: ' + e.message };
    }
}
// ─── Días hábiles ─────────────────────────────────────────────────────────────
/** Calcula cuántos días hábiles faltan desde hoy hasta la fecha de vencimiento (apertura + 7 hábiles). */
function calcularVencimientoCD(fechaAperturaStr) {
    const apertura = new Date(fechaAperturaStr);
    let diasHabiles = 0;
    let fecha = new Date(apertura);
    while (diasHabiles < 7) {
        fecha.setDate(fecha.getDate() + 1);
        const dow = fecha.getDay();
        if (dow !== 0 && dow !== 6)
            diasHabiles++; // excluir sáb y dom
    }
    const vencimiento = new Date(fecha);
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    vencimiento.setHours(0, 0, 0, 0);
    // Calcular días hábiles restantes entre hoy y vencimiento
    let restantes = 0;
    const cursor = new Date(hoy);
    if (hoy <= vencimiento) {
        while (cursor <= vencimiento) {
            const d = cursor.getDay();
            if (d !== 0 && d !== 6)
                restantes++;
            cursor.setDate(cursor.getDate() + 1);
        }
    }
    const fv = Utilities.formatDate(vencimiento, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    return {
        diasRestantes: restantes,
        vencido: hoy > vencimiento,
        fechaVencimiento: fv
    };
}
// ─── Requerimientos ───────────────────────────────────────────────────────────
/** Crea un nuevo requerimiento con los 15 documentos en PENDIENTE. */
function crearRequerimiento(apellidoNombre, autoridad, usuarioId) {
    try {
        initHojasCD();
        const sesion = getSesionCD();
        const uid = usuarioId || sesion.usuarioId;
        if (!uid)
            return { ok: false, message: 'No hay sesión activa.' };
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        // Generar ID_REQ
        const hojaReq = ss.getSheetByName(CD_HOJA_REQUERIMIENTOS);
        const lastRowReq = hojaReq.getLastRow();
        const idReq = lastRowReq < 1 ? 1 : lastRowReq; // lastRow incluye encabezado
        const fechaApertura = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        // Crear subcarpeta en DOCUMENTACION
        const folderName = `${apellidoNombre.trim()}_${autoridad.trim()}`;
        const mainFolder = DriveApp.getFolderById(CARPETA_DOCUMENTACION);
        const subFolder = mainFolder.createFolder(folderName);
        subFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // <--- HERENCIA PARA TODO LO QUE SE SUBA
        const idCarpetaDrive = subFolder.getId();
        hojaReq.appendRow([idReq, uid, apellidoNombre.trim(), autoridad.trim(), fechaApertura, 'EN CURSO', '', idCarpetaDrive]);
        // Crear los 15 documentos en PENDIENTE en una sola operación (Batch Write)
        const hojaDoc = ss.getSheetByName(CD_HOJA_DOCUMENTOS);
        const lastRowDoc = hojaDoc.getLastRow();
        let nextIdDoc = lastRowDoc < 1 ? 1 : lastRowDoc;
        const rowsToInsert = [];
        for (let i = 1; i <= 15; i++) {
            // ID_DOC, ID_REQ, NUM_REQ, ESTADO, URL_DRIVE, FECHA_CARGA, FECHA_APROBACION, OBS_RECHAZO, APROBADO_POR
            rowsToInsert.push([nextIdDoc, idReq, i, 'PENDIENTE', '', '', '', '', '']);
            nextIdDoc++;
        }
        hojaDoc.getRange(lastRowDoc + 1, 1, 15, 9).setValues(rowsToInsert);
        return { ok: true, idReq: idReq };
    }
    catch (e) {
        return { ok: false, message: 'Error al crear requerimiento: ' + e.message };
    }
}
/** Lista requerimientos. El usuario ve los suyos; el supervisor/admin ve todos. */
function listarRequerimientos(usuarioId, rol) {
    try {
        initHojasCD();
        const sesion = getSesionCD();
        const uid = usuarioId || sesion.usuarioId;
        const rол = (rol || sesion.rol).toLowerCase();
        if (!uid)
            return { ok: false, message: 'No hay sesión activa. Por favor iniciá sesión nuevamente.' };
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hojaReq = ss.getSheetByName(CD_HOJA_REQUERIMIENTOS);
        if (!hojaReq || hojaReq.getLastRow() < 2)
            return { ok: true, data: [] };
        // --- OPTIMIZACIÓN: Pre-cargar datos en Mapas para evitar O(N*M) ---
        // Map de aprobados: idReq -> cantidad
        const hojaDoc = ss.getSheetByName(CD_HOJA_DOCUMENTOS);
        const docsData = (hojaDoc && hojaDoc.getLastRow() > 1) ? hojaDoc.getDataRange().getValues().slice(1) : [];
        const aprobadosMap = {};
        for (const d of docsData) {
            const idR = String(d[1]);
            if (d[3] === 'APROBADO') {
                aprobadosMap[idR] = (aprobadosMap[idR] || 0) + 1;
            }
        }
        // Map de usuarios: uid -> nombre
        const hojaConfig = ss.getSheetByName(CD_HOJA_CONFIG);
        const configData = hojaConfig ? hojaConfig.getDataRange().getValues().slice(1) : [];
        const usuariosMap = {};
        for (const c of configData) {
            usuariosMap[String(c[0])] = String(c[1]);
        }
        const rows = hojaReq.getDataRange().getValues().slice(1);
        const result = [];
        for (const row of rows) {
            const [idReq, idUsuario, apellidoNombre, autoridad, fechaApertura, estado, idPdfDrive, idCarpetaDrive] = row;
            // Filtrado por rol
            if ((rол === 'usuario') && String(idUsuario) !== uid)
                continue;
            const idReqStr = String(idReq);
            const aprobados = aprobadosMap[idReqStr] || 0;
            const nombreUsuario = usuariosMap[String(idUsuario)] || String(idUsuario);
            // Vencimiento (esta función es ligera pero se podría optimizar más si fuera necesario)
            const venc = fechaApertura ? calcularVencimientoCD(String(fechaApertura)) : { diasRestantes: 0, vencido: false, fechaVencimiento: '' };
            result.push({
                idReq: idReqStr,
                idUsuario: String(idUsuario),
                nombreUsuario,
                apellidoNombre: String(apellidoNombre),
                autoridad: String(autoridad),
                fechaApertura: String(fechaApertura),
                estado: String(estado),
                idPdfDrive: String(idPdfDrive),
                idCarpetaDrive: String(idCarpetaDrive || ''),
                aprobados,
                total: 15,
                diasRestantes: venc.diasRestantes,
                vencido: venc.vencido,
                fechaVencimiento: venc.fechaVencimiento
            });
        }
        return { ok: true, data: result };
    }
    catch (e) {
        return { ok: false, message: 'Error al listar requerimientos: ' + e.message };
    }
}
/** Retorna los 15 documentos de un requerimiento específico. */
function getDocumentosReq(idReq) {
    try {
        initHojasCD();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hojaDoc = ss.getSheetByName(CD_HOJA_DOCUMENTOS);
        if (!hojaDoc || hojaDoc.getLastRow() < 2)
            return { ok: true, data: [] };
        const rows = hojaDoc.getDataRange().getValues().slice(1);
        const docs = rows
            .filter(r => String(r[1]) === String(idReq))
            .map(r => ({
            idDoc: String(r[0]),
            idReq: String(r[1]),
            numReq: Number(r[2]),
            nombreReq: CD_REQUISITOS[Number(r[2]) - 1] || 'Requisito ' + r[2],
            estado: String(r[3]), // PENDIENTE | CARGADO | APROBADO | RECHAZADO
            urlDrive: String(r[4]),
            fechaCarga: String(r[5]),
            fechaAprobacion: String(r[6]),
            obsRechazo: String(r[7]),
            aprobadoPor: String(r[8])
        }))
            .sort((a, b) => a.numReq - b.numReq);
        return { ok: true, data: docs };
    }
    catch (e) {
        return { ok: false, message: 'Error al obtener documentos: ' + e.message };
    }
}
// ─── Documentos ────────────────────────────────────────────────────────────────
/** Actualiza una fila de DOCUMENTOS_CD. Helper interno. */
function _updateDocRow(hojaDoc, idDoc, campos) {
    const rows = hojaDoc.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(idDoc)) {
            for (const col in campos) {
                hojaDoc.getRange(i + 1, Number(col) + 1).setValue(campos[col]);
            }
            return true;
        }
    }
    return false;
}
/** El usuario sube un archivo local (codificado en base64) a Google Drive y registra la URL en el sistema. */
function subirDocumentoLocalCD(idDoc, mimeType, base64Data, fileName) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hojaDoc = ss.getSheetByName(CD_HOJA_DOCUMENTOS);
        if (!hojaDoc)
            return { ok: false, message: 'Hoja de documentos no encontrada.' };
        // Decodificar el base64 a un Blob
        const data = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(data, mimeType, fileName);
        // Obtener ID_CARPETA_DRIVE del requerimiento
        let idCarpeta = CARPETA_DOCUMENTACION; // Fallback
        const hojaReq = ss.getSheetByName(CD_HOJA_REQUERIMIENTOS);
        if (hojaReq) {
            const reqRows = hojaReq.getDataRange().getValues();
            const docRows = hojaDoc.getDataRange().getValues();
            let idReq = '';
            for (const dr of docRows) {
                if (String(dr[0]) === String(idDoc)) {
                    idReq = String(dr[1]);
                    break;
                }
            }
            if (idReq) {
                for (const rr of reqRows) {
                    if (String(rr[0]) === idReq) {
                        idCarpeta = String(rr[7] || CARPETA_DOCUMENTACION);
                        break;
                    }
                }
            }
        }
        const carpeta = DriveApp.getFolderById(idCarpeta);
        const file = carpeta.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const urlDrive = file.getUrl();
        const fechaCarga = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        // Columnas (0-based): 3=ESTADO, 4=URL_DRIVE, 5=FECHA_CARGA, 7=OBS_RECHAZO
        const ok = _updateDocRow(hojaDoc, idDoc, { 3: 'CARGADO', 4: urlDrive, 5: fechaCarga, 7: '' });
        return ok ? { ok: true } : { ok: false, message: 'Documento no encontrado en la hoja.' };
    }
    catch (e) {
        return { ok: false, message: 'Error al subir archivo a Drive: ' + e.message };
    }
}
/** El supervisor aprueba un documento. */
function aprobarDocumento(idDoc) {
    try {
        const sesion = getSesionCD();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hojaDoc = ss.getSheetByName(CD_HOJA_DOCUMENTOS);
        if (!hojaDoc)
            return { ok: false, message: 'Hoja de documentos no encontrada.' };
        const fechaAprobacion = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        // 3=ESTADO, 6=FECHA_APROBACION, 7=OBS_RECHAZO, 8=APROBADO_POR
        const ok = _updateDocRow(hojaDoc, idDoc, { 3: 'APROBADO', 6: fechaAprobacion, 7: '', 8: sesion.nombre });
        if (!ok)
            return { ok: false, message: 'Documento no encontrado.' };
        _checkCompletitudReq(hojaDoc, ss, idDoc);
        return { ok: true };
    }
    catch (e) {
        return { ok: false, message: 'Error al aprobar: ' + e.message };
    }
}
/** Aprueba múltiples documentos en una sola llamada al servidor. */
function aprobarDocumentosLote(idDocs) {
    try {
        const sesion = getSesionCD();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hojaDoc = ss.getSheetByName(CD_HOJA_DOCUMENTOS);
        if (!hojaDoc)
            return { ok: false, message: 'Hoja de documentos no encontrada.' };
        const fechaAprobacion = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        const rows = hojaDoc.getDataRange().getValues();
        const idSet = new Set(idDocs.map(String));
        let reqsParaChequear = new Set();
        for (let i = 1; i < rows.length; i++) {
            const currentId = String(rows[i][0]);
            if (idSet.has(currentId)) {
                // Actualizar valores en la hoja (Col 3, 6, 7, 8)
                hojaDoc.getRange(i + 1, 4).setValue('APROBADO'); // Col 3 (0-based)
                hojaDoc.getRange(i + 1, 7).setValue(fechaAprobacion); // Col 6
                hojaDoc.getRange(i + 1, 8).setValue(''); // Col 7 (limpiar obs)
                hojaDoc.getRange(i + 1, 9).setValue(sesion.nombre); // Col 8
                reqsParaChequear.add(String(rows[i][1])); // Guardar idReq
            }
        }
        // Al final, chequear completitud de los requerimientos afectados
        for (const idReq of reqsParaChequear) {
            _finalCheckCompletitudReqManual(hojaDoc, ss, idReq);
        }
        return { ok: true };
    }
    catch (e) {
        return { ok: false, message: 'Error en aprobación masiva: ' + e.message };
    }
}
/** Versión de chequeo que recibe idReq directamente por eficiencia. */
function _finalCheckCompletitudReqManual(hojaDoc, ss, idReq) {
    const rows = hojaDoc.getDataRange().getValues();
    const docsDeeReq = rows.slice(1).filter(r => String(r[1]) === idReq);
    const todosAprobados = docsDeeReq.length === 15 && docsDeeReq.every(r => r[3] === 'APROBADO');
    if (todosAprobados) {
        const hojaReq = ss.getSheetByName(CD_HOJA_REQUERIMIENTOS);
        if (!hojaReq)
            return;
        const reqRows = hojaReq.getDataRange().getValues();
        for (let i = 1; i < reqRows.length; i++) {
            if (String(reqRows[i][0]) === idReq) {
                hojaReq.getRange(i + 1, 6).setValue('COMPLETO');
                break;
            }
        }
    }
}
/** Verifica si todos los docs de un req están aprobados y actualiza el estado. */
function _checkCompletitudReq(hojaDoc, ss, idDoc) {
    const rows = hojaDoc.getDataRange().getValues();
    // Buscar idReq del idDoc
    let idReq = '';
    for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(idDoc)) {
            idReq = String(rows[i][1]);
            break;
        }
    }
    if (!idReq)
        return;
    const docsDeeReq = rows.slice(1).filter(r => String(r[1]) === idReq);
    const todosAprobados = docsDeeReq.length === 15 && docsDeeReq.every(r => r[3] === 'APROBADO');
    if (todosAprobados) {
        const hojaReq = ss.getSheetByName(CD_HOJA_REQUERIMIENTOS);
        if (!hojaReq)
            return;
        const reqRows = hojaReq.getDataRange().getValues();
        for (let i = 1; i < reqRows.length; i++) {
            if (String(reqRows[i][0]) === idReq) {
                hojaReq.getRange(i + 1, 6).setValue('COMPLETO');
                break;
            }
        }
    }
}
/** El supervisor rechaza un documento con un motivo y elimina el archivo físico de Drive. */
function rechazarDocumento(idDoc, motivo) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hojaDoc = ss.getSheetByName(CD_HOJA_DOCUMENTOS);
        if (!hojaDoc)
            return { ok: false, message: 'Hoja de documentos no encontrada.' };
        // Buscar la URL del archivo antes de rechazar
        const rows = hojaDoc.getDataRange().getValues();
        let urlParaBorrar = '';
        for (let i = 1; i < rows.length; i++) {
            if (String(rows[i][0]) === String(idDoc)) {
                urlParaBorrar = String(rows[i][4] || ''); // Columna 4 (URL_DRIVE)
                break;
            }
        }
        // Borrado físico en Drive
        if (urlParaBorrar) {
            try {
                const fileIdMatch = urlParaBorrar.match(/[-\w]{25,}/);
                if (fileIdMatch) {
                    DriveApp.getFileById(fileIdMatch[0]).setTrashed(true);
                }
            }
            catch (e) {
                Logger.log("No se pudo borrar el archivo en Drive al rechazar: " + e.message);
            }
        }
        // 3=ESTADO, 4=URL_DRIVE, 6=FECHA_APROBACION, 7=OBS_RECHAZO, 8=APROBADO_POR
        // Limpiamos URL_DRIVE para que el sistema sepa que no hay archivo
        const ok = _updateDocRow(hojaDoc, idDoc, { 3: 'RECHAZADO', 4: '', 6: '', 7: motivo, 8: '' });
        return ok ? { ok: true } : { ok: false, message: 'Documento no encontrado.' };
    }
    catch (e) {
        return { ok: false, message: 'Error al rechazar: ' + e.message };
    }
}
/** El supervisor elimina un requerimiento completo y todos sus documentos. */
function eliminarRequerimientoComp(idReq, uid, rol) {
    try {
        if (rol !== 'supervisor' && rol !== 'admin') {
            return { ok: false, message: 'No tienes permisos para eliminar.' };
        }
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        // Eliminar de REQUERIMIENTOS_CD
        const hojaReq = ss.getSheetByName(CD_HOJA_REQUERIMIENTOS);
        if (!hojaReq)
            return { ok: false, message: 'Hoja de requerimientos no encontrada.' };
        const reqRows = hojaReq.getDataRange().getValues();
        let rowReq = -1;
        for (let i = 1; i < reqRows.length; i++) {
            if (String(reqRows[i][0]) === String(idReq)) {
                rowReq = i + 1;
                break;
            }
        }
        let idCarpetaTrash = '';
        if (rowReq !== -1) {
            // Obtener la carpeta antes de borrar la fila (ID_CARPETA_DRIVE está en el índice 7)
            idCarpetaTrash = String(reqRows[rowReq - 1][7] || '');
            hojaReq.deleteRow(rowReq);
        }
        else {
            return { ok: false, message: 'Requerimiento no encontrado.' };
        }
        const hojaDoc = ss.getSheetByName(CD_HOJA_DOCUMENTOS);
        if (hojaDoc) {
            const docRows = hojaDoc.getDataRange().getValues();
            for (let i = docRows.length - 1; i >= 1; i--) {
                if (String(docRows[i][1]) === String(idReq)) {
                    hojaDoc.deleteRow(i + 1);
                }
            }
        }
        // Borrar carpeta en Drive
        if (idCarpetaTrash) {
            try {
                DriveApp.getFolderById(idCarpetaTrash).setTrashed(true);
            }
            catch (e) {
                Logger.log("Error al borrar carpeta en Drive: " + e.message);
            }
        }
        return { ok: true, message: 'Requerimiento y carpeta eliminados con éxito.' };
    }
    catch (e) {
        return { ok: false, message: 'Error al eliminar requerimiento: ' + e.message };
    }
}
/** El usuario elimina un documento rechazado para poder subir uno nuevo. */
function eliminarDocumento(idDoc) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hojaDoc = ss.getSheetByName(CD_HOJA_DOCUMENTOS);
        if (!hojaDoc)
            return { ok: false, message: 'Hoja de documentos no encontrada.' };
        // 3=ESTADO, 4=URL_DRIVE, 5=FECHA_CARGA, 7=OBS_RECHAZO
        const ok = _updateDocRow(hojaDoc, idDoc, { 3: 'PENDIENTE', 4: '', 5: '', 7: '' });
        return ok ? { ok: true } : { ok: false, message: 'Documento no encontrado.' };
    }
    catch (e) {
        return { ok: false, message: 'Error al eliminar: ' + e.message };
    }
}
// ─── Generación de PDF ────────────────────────────────────────────────────────
/**
 * Genera el PDF de checklist aprobado para un requerimiento.
 * Retorna la URL del PDF en Drive.
 */
async function generarPdfCD(idReq) {
    try {
        const sesion = getSesionCD();
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        // Obtener datos del requerimiento
        const hojaReq = ss.getSheetByName(CD_HOJA_REQUERIMIENTOS);
        if (!hojaReq)
            return { ok: false, message: 'Hoja de requerimientos no encontrada.' };
        const reqRows = hojaReq.getDataRange().getValues();
        let reqData = null;
        for (let i = 1; i < reqRows.length; i++) {
            if (String(reqRows[i][0]) === String(idReq)) {
                reqData = reqRows[i];
                break;
            }
        }
        if (!reqData)
            return { ok: false, message: 'Requerimiento no encontrado.' };
        // Obtener docs
        const docsResult = getDocumentosReq(idReq);
        if (!docsResult.ok || !docsResult.data)
            return { ok: false, message: 'Error al obtener documentos.' };
        const docs = docsResult.data.filter(d => d.estado === 'APROBADO');
        if (docs.length === 0) {
            return { ok: false, message: 'Debe haber al menos un documento aprobado para generar el PDF.' };
        }
        // Obtener ID de subcarpeta para el PDF
        let idCarpetaFinal = String(reqData[7] || CARPETA_DOCUMENTACION);
        // Crear documento Google Docs
        const titulo = `CHECKLIST_${String(reqData[2]).replace(/,/g, '').replace(/ /g, '_')}_${idReq}`;
        const doc = DocumentApp.create(titulo);
        const body = doc.getBody();
        // --- Carátula ---
        body.appendParagraph('CHECKLIST - CONTRATO DE LOCACIÓN')
            .setHeading(DocumentApp.ParagraphHeading.HEADING1)
            .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('');
        const parContratado = body.appendParagraph(`Contratado: ${reqData[2]}`);
        parContratado.editAsText().setBold(true);
        body.appendParagraph(`Autoridad Solicitante: ${reqData[3]}`);
        body.appendParagraph(`Fecha de apertura: ${reqData[4]}`);
        body.appendParagraph(`Supervisor que aprobó: ${sesion.nombre}`);
        body.appendParagraph(`Fecha de generación: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}`);
        body.appendParagraph('').appendHorizontalRule();
        body.appendParagraph('');
        // --- Unificación de PDFs Reales ---
        const pdfBlobs = [];
        // 1. Agregar la carátula/reporte que ya creamos en el Google Doc (convertido a PDF)
        doc.saveAndClose();
        const caratulaBlob = DriveApp.getFileById(doc.getId()).getBlob().setName('00_Reporte.pdf');
        pdfBlobs.push(caratulaBlob);
        // 2. Agregar cada uno de los archivos originales aprobados
        for (const d of docs) {
            if (d.urlDrive) {
                try {
                    // Extraer ID de Drive desde la URL
                    const match = d.urlDrive.match(/[-\w]{25,}/);
                    if (match) {
                        const docBlob = DriveApp.getFileById(match[0]).getBlob();
                        pdfBlobs.push(docBlob);
                    }
                }
                catch (e) {
                    Logger.log("Error al obtener blob para unificar: " + e.message);
                }
            }
        }
        // 3. Unificar todos los blobs en uno solo usando la librería PDFApp
        let finalBlob;
        try {
            // PDFApp.mergePDFs acepta un array de blobs y devuelve el blob unificado
            finalBlob = await PDFApp.mergePDFs(pdfBlobs);
            finalBlob.setName(titulo + '.pdf');
        }
        catch (e) {
            Logger.log("Error en la unificación real, cayendo al reporte simple: " + e.message);
            // Fallback: usar solo la carátula si falla la unión
            finalBlob = caratulaBlob.setName(titulo + '.pdf');
        }
        // 4. Guardar archivo final y compartirlos
        const pdfFile = DriveApp.getFolderById(idCarpetaFinal).createFile(finalBlob);
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        // Borrar el .docs temporal
        DriveApp.getFileById(doc.getId()).setTrashed(true);
        const urlPdf = pdfFile.getUrl();
        // Guardar URL en la hoja de requerimientos (columna 7 = ID_PDF_DRIVE)
        for (let i = 1; i < reqRows.length; i++) {
            if (String(reqRows[i][0]) === String(idReq)) {
                hojaReq.getRange(i + 1, 7).setValue(pdfFile.getId());
                hojaReq.getRange(i + 1, 6).setValue('PDF_GENERADO');
                break;
            }
        }
        return { ok: true, urlPdf };
    }
    catch (e) {
        return { ok: false, message: 'Error al generar PDF: ' + e.message };
    }
}
/** Retorna la URL de previsualización del PDF ya generado. */
function getPdfUrlCD(idReq) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hojaReq = ss.getSheetByName(CD_HOJA_REQUERIMIENTOS);
        if (!hojaReq)
            return { ok: false, message: 'Hoja no encontrada.' };
        const rows = hojaReq.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
            if (String(rows[i][0]) === String(idReq) && rows[i][6]) {
                const fileId = String(rows[i][6]);
                const url = `https://drive.google.com/file/d/${fileId}/view`;
                return { ok: true, urlPdf: url };
            }
        }
        return { ok: false, message: 'PDF no generado aún.' };
    }
    catch (e) {
        return { ok: false, message: 'Error: ' + e.message };
    }
}
/** Asegura que un archivo sea público para visualización rápida antes de abrir el visor. */
function asegurarAccesoPublico(url) {
    try {
        const match = url.match(/[-\w]{25,}/);
        if (match) {
            const file = DriveApp.getFileById(match[0]);
            // Solo aplicar si no tiene ya el permiso correcto (para ahorrar tiempo)
            if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
                file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            }
        }
        return { ok: true };
    }
    catch (e) {
        return { ok: false };
    }
}
/** URL de la página de Control Docus (para el botón del Index principal). */
function getControlDocusUrl() {
    return ScriptApp.getService().getUrl() + '?page=PanelSupervisor';
}
