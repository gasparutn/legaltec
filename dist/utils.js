"use strict";
/**
 * Registra un movimiento en la hoja de Control.
 * Columnas: Usuario | Movimientos | Fecha y Hora
 */
function logActivity(movimiento, activeAutor) {
    try {
        const props = PropertiesService.getUserProperties();
        const userName = props.getProperty('CD_NOMBRE') || 'Usuario Desconocido';
        const autor = activeAutor || props.getProperty('AUTOR_COD') || '';
        // Abrir la hoja de control de forma eficiente
        const sheet = getSS().getSheetByName("Control");
        if (!sheet) {
            Logger.log("Error: La hoja 'Control' no existe.");
            return;
        }
        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        // Columnas en Sheet 'Control': Usuario | Movimiento | Fecha | Iniciales Autor
        sheet.appendRow([userName, movimiento, now, autor]);
    }
    catch (e) {
        Logger.log("Error al registrar actividad: " + e.message);
    }
}
/**
 * Sanitiza una cadena de texto para evitar inyecciones XSS quitando caracteres peligrosos.
 */
function sanitizeInput(str) {
    if (!str)
        return '';
    return String(str)
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}
/**
 * Normaliza un encabezado para comparación (Mayúsculas, Sin Espacios, Sin Acentos)
 */
function normalizeHeader(h) {
    return String(h || '').trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
/**
 * Crea un mapa de nombres de columna (normalizados) a índices (0-based) para una hoja.
 * Permite que el sistema sea resistente a cambios en el orden de las columnas.
 */
function getColumnMap(sheet) {
    if (!sheet)
        return {};
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0)
        return {};
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const map = {};
    headers.forEach((h, idx) => {
        const norm = normalizeHeader(h);
        if (norm)
            map[norm] = idx;
    });
    return map;
}
/**
 * Parsea una fecha de forma segura, priorizando el formato DD/MM/YYYY y el formato largo de Argentina (D de Mes de YYYY)
 */
function parseSafeDate(val) {
    if (val instanceof Date)
        return val;
    if (!val || String(val).trim() === '')
        return new Date(NaN);
    const s = String(val).trim().toLowerCase();
    // 1. Soporta DD/MM/YYYY o D/M/YYYY con separators / o -
    const partsDotsSlsh = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (partsDotsSlsh) {
        return new Date(parseInt(partsDotsSlsh[3], 10), parseInt(partsDotsSlsh[2], 10) - 1, parseInt(partsDotsSlsh[1], 10));
    }
    // 2. Soporta: "1 de febrero de 2026" o "01 de mayo de 2024"
    const partsLong = s.match(/^(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{4})/);
    if (partsLong) {
        const months = {
            'enero': 0, 'febrero': 1, 'marzo': 2, 'abril': 3, 'mayo': 4, 'junio': 5,
            'julio': 6, 'agosto': 7, 'septiembre': 8, 'octubre': 9, 'noviembre': 10, 'diciembre': 11
        };
        const day = parseInt(partsLong[1], 10);
        const monthName = partsLong[2];
        const year = parseInt(partsLong[3], 10);
        if (months[monthName] !== undefined) {
            return new Date(year, months[monthName], day);
        }
    }
    // 3. Formato YYYY-MM-DD
    const partsIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (partsIso) {
        return new Date(parseInt(partsIso[1], 10), parseInt(partsIso[2], 10) - 1, parseInt(partsIso[3], 10));
    }
    const d = new Date(val);
    return isNaN(d.getTime()) ? new Date(val.toString()) : d;
}
function fechaEnEspaniol(valor) {
    if (!valor)
        return '';
    const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    const d = parseSafeDate(valor);
    if (isNaN(d.getTime()))
        return String(valor);
    return d.getDate() + ' de ' + MESES[d.getMonth()] + ' de ' + d.getFullYear();
}
function replaceFields(target, fields) {
    for (const [key, value] of Object.entries(fields)) {
        const finalValue = value === undefined || value === null ? '' : String(value);
        // Obtener el nombre del marcador sin los brackets << >>
        let name = key.replace(/^<<|>>$/g, '').trim();
        // Escapar caracteres especiales en el nombre
        let escapedName = name.replace(/[$^*.+?()\]|{}]/g, '\\$&');
        // Crear un patrón que permita espacios opcionales y sea INSENSIBLE a mayúsculas: (?i)<<\s*NOMBRE\s*>>
        const pattern = "(?i)<<\\s*" + escapedName + "\\s*>>";
        try {
            target.replaceText(pattern, finalValue);
        }
        catch (e) {
            Logger.log(`Error al reemplazar marcador ${key}: ${e.message}`);
        }
    }
}
/**
 * Reemplaza todos los campos de la planilla y auxiliares en el documento.
 * Genera marcadores automáticos <<KEY>> y <<KEY_CON_ESPACIO>>.
 */
function replaceFullRow(body, data) {
    const fields = {};
    for (const [key, value] of Object.entries(data)) {
        const upperKey = key.trim().toUpperCase();
        // 1. Limpiar espacios invisibles (Trimming) de los valores
        let finalValue = value;
        if (typeof value === 'string') {
            finalValue = value.trim();
        }
        // 2. Formatear CUOTA con cero a la izquierda si es dígito único
        if (upperKey === 'CUOTA' && finalValue !== undefined && finalValue !== null && finalValue !== '') {
            const num = parseInt(String(finalValue), 10);
            if (!isNaN(num) && num > 0 && num < 10) {
                finalValue = '0' + num;
            }
        }
        fields[`<<${upperKey}>>`] = finalValue;
        // Si la clave tiene espacios, también generar versión con guión bajo para compatibilidad
        if (upperKey.includes(' ')) {
            fields[`<<${upperKey.replace(/\s+/g, '_')}>>`] = finalValue;
        }
    }
    // Reemplazos específicos que requieren formato (DNI, Dinero, Fechas)
    if (data['DNI']) {
        const dni = String(data['DNI']).replace(/\D/g, '');
        if (dni) {
            fields['<<DNI>>'] = dni.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
        }
    }
    // Formatear automáticamente campos de fecha conocidos si tienen valor
    const camposFecha = ['FECHA ALTA', 'FECHA BAJA', 'FECHA RESOLUCION', 'VENCIMIENTO'];
    camposFecha.forEach(key => {
        if (data[key]) {
            const f = fechaEnEspaniol(data[key]);
            fields[`<<${key.toUpperCase().replace(/\s+/g, '_')}>>`] = f;
            fields[`<<${key.toUpperCase()}>>`] = f;
        }
    });
    // Lista de campos que deben formatearse como moneda ($1.000)
    const camposMoneda = ['TOTAL', 'TOTAL_DIV_CUO', 'AUX18'];
    camposMoneda.forEach(key => {
        const val = data[key];
        if (val !== undefined && val !== null && val !== '') {
            const n = typeof val === 'string' ? parseFloat(val.replace(/[^0-9,-]+/g, "").replace(",", ".")) : Number(val);
            if (!isNaN(n)) {
                // Formato: $1.000 (Puntos para miles, sin decimales para enteros)
                fields[`<<${key.toUpperCase()}>>`] = "$" + n.toLocaleString('es-AR', {
                    minimumFractionDigits: 0,
                    maximumFractionDigits: 2
                });
            }
        }
    });
    replaceFields(body, fields);
}
function generateDocumentFromTemplate(templateId, fields, folderId, filename, format = 'docx') {
    const template = DriveApp.getFileById(templateId);
    const folder = DriveApp.getFolderById(folderId);
    const copy = template.makeCopy(filename, folder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    replaceFields(body, fields);
    doc.saveAndClose();
    if (format === 'pdf') {
        // Usar 'application/pdf' o MimeType.PDF si está disponible.
        // En los tipos, a veces está como MimeType.PDF
        const pdfBlob = copy.getAs('application/pdf');
        const pdfFile = folder.createFile(pdfBlob);
        copy.setTrashed(true);
        return { url: pdfFile.getUrl(), nombre: pdfFile.getName() };
    }
    return { url: copy.getUrl(), nombre: copy.getName() };
}
/**
 * Guarda un objeto en el CacheService de forma serializada.
 */
function setCache(key, value, expirationInSeconds = 600) {
    try {
        const cache = CacheService.getScriptCache();
        const serialized = JSON.stringify(value);
        // CacheService tiene un límite de 100KB por entrada.
        if (serialized.length < 100000) {
            cache.put(key, serialized, expirationInSeconds);
        }
    }
    catch (e) {
        Logger.log("Error al escribir en caché: " + e.message);
    }
}
/**
 * Recupera un objeto del CacheService.
 */
function getCache(key) {
    try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get(key);
        return cached ? JSON.parse(cached) : null;
    }
    catch (e) {
        Logger.log("Error al leer de caché: " + e.message);
        return null;
    }
}
/**
 * Devuelve la URL de ejecución de la web app
 */
function getAppUrl() {
    return ScriptApp.getService().getUrl();
}
