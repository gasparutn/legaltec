"use strict";
/**
 * Firebase.ts - Conector para Cloud Firestore
 * Permite la gestión de datos en tiempo real para Legaltec.
 */
// --- CONFIGURACIÓN DE FIREBASE ---
// Reemplaza esto con los datos de tu archivo JSON de Cuenta de Servicio
const FIREBASE_CONFIG = {
    project_id: "legaltec-bd",
    client_email: "firebase-adminsdk-fbsvc@legaltec-bd.iam.gserviceaccount.com",
    private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCTxxOTM5H7PdYD\nCb6uUQmLRYiC7ilfkgk5lvPiqUy8J1LqkWhCCnPD5ygWDu1MW59X9PM5JEFsaF3N\nbh/WHT3kejhewcjAJeraqyP8+1gZfWWZm7uEZpck/8/h5mxAxdaScmTDy4jAYtrB\neLYQ/KR3RSDOdE3/EhpyLF7LVvGtDGaRSz7JAH91Mfj5IM1INvNrv6Qzvrh0KBu4\nRuPjD+IbiPQjd+W25pWesKdxiXDMZ0N9tkVk/LAVqTgi8QvP55sa2rlIy+fLjC8y\nFCep9LE4gVNCqTI2i15sjvkCLTPfkDJLxTiIeQmBPj4mw+lWEZHhn3DgwxA4G6Vm\nqomw7TQ7AgMBAAECggEAHlxZtLX1KYmQRQqE/vpOPCtaew/kNyrhM1BEpg5DYiqw\nGdoq8dqE4zxEG5gQf84ZJDCCofxFZrjggLq91rcmRqEnoKwuiu+fWzKrD6gx9vaw\n34jD01VieysFcjNtjTc4ONlAw4s2+tO0fuhGe6H0Rj7RGSmC9sMo5Mkh5qPjvjoM\nkpNaKmvRQdlL/D97RaRlHm+sfNu6NteWlLEwputF1oJXpZBBIz08vNNVy7qUj7VC\nzhqEjSUDnPhps529b5555eISa82qzrbhbwOQ5JVJAuCcDTRDo8QIlvaSIL9v51hG\n/r0ukPEEtaiFgxNPgj12+5sDG3wrs+BVwBFgkegHEQKBgQDC3fNcKiSnCh8QhOdS\nfGq2Hzy/RvYdizaMLWtIWMDfUoJJWjgPw1tEqMy5RTgCSMAOCwZr2Z+sA5LwNLlt\nY3fiC6LnGHwS/mp152vc6vkoIzmy5p3HGbftgRR71QYPnyxpAEQlHwWWpVhu/gIu\n4yE1whMEIqBw7QgfvXTf6MJ20QKBgQDCI0/ivSpum6Q1TbR0gowVsrW8gLW5T3I7\nfYPdoJMBfxwXfpjLJvBI9uCtRWrk37x+WKR8kWOjE4qBAcOBdcOR2Q9pqGfQYg6C\nJYP9DmV0c3Lc/GTW3jQWMxZ1xjcqqby2hPLQp7Kip3l9OJQzPzIVWwy4XefGCAHg\nLIx5SLFVSwKBgGuzvJPN/yALqgu5SRkABwLN1QdrMxA1J1rmp4r+8ur7cWMVaDLe\nKI2UNkKYDVLF3tBkK5JkX6n097uniGz7MwFOqSTNFZZx42lzFNyvSjJy9ar5Z27p\nugyc8TNYE9eChEssiH1z3eTbUjtWTOKZSnup3lTExqRjfb/9OKGgxgphAoGAJN2J\nxPfBAAIEO68Gzn0f1tr50dLrL0Zb3dom2UYtxEH79qYuo51AUcq3vY+KDe5CoAd7\nLMB7GDOUwIyyMHpJaf+7AohdNl/4e1RebQZ305Y8wKQZDghiAbkv0auTPEHBOTOj\n3GWL8YE1+8ZrthOREicZEkWaxqssmAI9pqZcLmMCgYAmZBJzKEkj7EzBn4aSOJrE\n4X6k/iqqTK6IekfNHR82HaEX6PyaSEZFojXgsueKY3JoKJJkoiU34dKEIGqXhSCu\nguJ+xguDjSYY7ed99NoDSRAF7DRuSc4Kgz5BcYUV1qlD770Moi8lv9y39Jb1evF2\nPZ+2Eu6xwVcrwvPKJgX24A==\n-----END PRIVATE KEY-----\n"
};
/**
 * Prueba la conexión enviando un registro ficticio a la colección 'test'.
 */
function testFirebaseConnection() {
    const dummyData = {
        id: "TEST-001",
        usuario: "Admin Test",
        fecha: new Date().toLocaleString(),
        mensaje: "Hola desde Google Apps Script"
    };
    try {
        const result = firestoreCreateDocumentWithId("ConexionesPrueba", dummyData.id, dummyData);
        Logger.log("✅ Éxito al conectar con Firebase!");
        return { success: true, response: result };
    }
    catch (e) {
        Logger.log("❌ Error de conexión: " + e.message);
        return { success: false, error: e.message };
    }
}
/**
 * Genera un token de acceso OAuth2 usando la llave privada de la cuenta de servicio.
 */
function getServiceAccountToken() {
    const header = JSON.stringify({ alg: "RS256", typ: "JWT" });
    const now = Math.floor(Date.now() / 1000);
    const claimSet = JSON.stringify({
        iss: FIREBASE_CONFIG.client_email,
        scope: "https://www.googleapis.com/auth/datastore",
        aud: "https://oauth2.googleapis.com/token",
        exp: now + 3600,
        iat: now
    });
    const toSign = Utilities.base64EncodeWebSafe(header) + "." + Utilities.base64EncodeWebSafe(claimSet);
    const signature = Utilities.computeRsaSha256Signature(toSign, FIREBASE_CONFIG.private_key);
    const jwt = toSign + "." + Utilities.base64EncodeWebSafe(signature);
    const options = {
        method: "post",
        payload: { grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: jwt },
        muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", options);
    const res = JSON.parse(response.getContentText());
    if (res.error)
        throw new Error("Error obteniendo token: " + res.error_description);
    return res.access_token;
}
/**
 * Crea o sobreescribe completamente un documento en Firestore con un ID específico.
 */
function firestoreCreateDocumentWithId(collection, docId, data) {
    const projectId = FIREBASE_CONFIG.project_id;
    const encodedId = encodeURIComponent(String(docId).trim());
    // Sin updateMask ni allowMissing — sobreescritura total del documento
    const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/${collection}/${encodedId}`;
    const payload = { fields: {} };
    for (const key in data) {
        const val = data[key];
        if (val === true || val === false)
            payload.fields[key] = { booleanValue: val };
        else if (val instanceof Date)
            payload.fields[key] = { timestampValue: val.toISOString() };
        else if (typeof val === 'number')
            payload.fields[key] = { doubleValue: val };
        else
            payload.fields[key] = { stringValue: String(val || "") };
    }
    const options = {
        method: "patch", contentType: "application/json", payload: JSON.stringify(payload),
        headers: { "Authorization": "Bearer " + getServiceAccountToken() },
        muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    const respJson = JSON.parse(response.getContentText());
    if (respJson.error)
        throw new Error(`Error Firestore [${collection}/${docId}]: ${respJson.error.message}`);
    return respJson;
}
/**
 * Actualiza SOLO los campos especificados en un documento existente (PATCH parcial).
 * Usa updateMask para no sobreescribir el resto del documento.
 */
function firestoreUpdateDocument(collection, docId, data) {
    const projectId = FIREBASE_CONFIG.project_id;
    const encodedId = encodeURIComponent(String(docId).trim());
    // Los nombres de campo con espacios o caracteres especiales deben escaparse con backticks
    const fieldPaths = Object.keys(data).map(key => {
        const escapedKey = /[^a-zA-Z_0-9]/.test(key) ? '`' + key + '`' : key;
        return `updateMask.fieldPaths=${encodeURIComponent(escapedKey)}`;
    }).join('&');
    // PATCH parcial con updateMask — NO se usa allowMissing
    const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/${collection}/${encodedId}?${fieldPaths}`;
    const payload = { fields: {} };
    for (const key in data) {
        const val = data[key];
        if (val === true || val === false)
            payload.fields[key] = { booleanValue: val };
        else if (val instanceof Date)
            payload.fields[key] = { timestampValue: val.toISOString() };
        else if (typeof val === 'number')
            payload.fields[key] = { doubleValue: val };
        else
            payload.fields[key] = { stringValue: String(val || "") };
    }
    const options = {
        method: "patch", contentType: "application/json", payload: JSON.stringify(payload),
        headers: { "Authorization": "Bearer " + getServiceAccountToken() },
        muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    const respJson = JSON.parse(response.getContentText());
    if (respJson.error)
        throw new Error(`Error Firestore [${collection}/${docId}]: ${respJson.error.message}`);
    return respJson;
}
function firestoreDeleteDocument(collection, docId) {
    const projectId = FIREBASE_CONFIG.project_id;
    const encodedId = encodeURIComponent(String(docId).trim());
    const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/${collection}/${encodedId}`;
    const options = {
        method: "delete", headers: { "Authorization": "Bearer " + getServiceAccountToken() },
        muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200 && response.getResponseCode() !== 204) {
        throw new Error(`Fallo al eliminar documento ${docId}`);
    }
}
function firestoreGetDoc(collection, docId) {
    const projectId = FIREBASE_CONFIG.project_id;
    const encodedId = encodeURIComponent(String(docId).trim());
    const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/${collection}/${encodedId}`;
    const options = {
        method: "get", headers: { "Authorization": "Bearer " + getServiceAccountToken() },
        muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200)
        return null;
    const doc = JSON.parse(response.getContentText());
    const obj = {};
    const fields = doc.fields || {};
    for (const key in fields) {
        const valObj = fields[key];
        const valType = Object.keys(valObj)[0];
        obj[key] = valObj[valType];
    }
    const nameParts = doc.name.split('/');
    obj['__id'] = nameParts[nameParts.length - 1];
    return obj;
}
function firestoreGetAllDocs(collection) {
    try {
        const projectId = FIREBASE_CONFIG.project_id;
        const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/${collection}?pageSize=1000`;
        const options = {
            method: "get", headers: { "Authorization": "Bearer " + getServiceAccountToken() },
            muteHttpExceptions: true
        };
        const response = UrlFetchApp.fetch(url, options);
        const data = JSON.parse(response.getContentText());
        if (data.error)
            throw new Error(`Error Firestore: ${data.error.message}`);
        if (!data.documents)
            return [];
        return data.documents.map(doc => {
            const fields = doc.fields || {};
            const obj = {};
            const nameParts = doc.name.split('/');
            obj['__id'] = nameParts[nameParts.length - 1];
            for (const key in fields) {
                const valObj = fields[key];
                const valType = Object.keys(valObj)[0];
                obj[key] = valObj[valType];
            }
            return obj;
        });
    }
    catch (e) {
        Logger.log("Error en firestoreGetAllDocs: " + e.message);
        return [];
    }
}
/**
 * Sincroniza Firebase -> Sheet (Mirror).
 */
function firestoreToSheetSync(collectionName, sheetName) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet)
        return;
    const data = firestoreGetAllDocs(collectionName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const resultRows = [];
    data.forEach(item => {
        const row = [];
        headers.forEach(h => {
            const normalizedH = String(h).trim().toUpperCase();
            row.push(item[normalizedH] || '');
        });
        resultRows.push(row);
    });
    if (resultRows.length > 0) {
        if (sheet.getLastRow() > 1)
            sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
        sheet.getRange(2, 1, resultRows.length, headers.length).setValues(resultRows);
    }
}
function syncAllFirebaseToSheets() {
    const mapArr = [
        { sheet: "REGISTROS", col: "Registros" },
        { sheet: "REQUERIMIENTOS_CD", col: "Requerimientos" },
        { sheet: "DOCUMENTOS_CD", col: "Documentos" },
        { sheet: "USUARIOS", col: "Usuarios" },
        { sheet: "AUTORIDADES", col: "Autoridades" },
        { sheet: "AUXILIARES", col: "Auxiliares" }
    ];
    mapArr.forEach(m => {
        try {
            firestoreToSheetSync(m.col, m.sheet);
            Logger.log(`✅ Sincronizado: ${m.col} -> ${m.sheet}`);
        }
        catch (e) {
            Logger.log(`❌ Fallo sync de ${m.col}: ` + e.message);
        }
    });
}
/**
 * Función pensada para ser ejecutada por un activador de tiempo (Trigger).
 * Mantiene el Sheet actualizado como un espejo de Firebase.
 */
function scheduledSync() {
    syncAllFirebaseToSheets();
}
function migrateAllSheetsToFirebase() {
    const hojasAMigrar = [
        { sheet: "REGISTROS", col: "Registros", id: "ID_REG" },
        { sheet: "REQUERIMIENTOS_CD", col: "Requerimientos", id: "ID_REQ" },
        { sheet: "DOCUMENTOS_CD", col: "Documentos", id: "DOC_ID" },
        { sheet: "CONTROL", col: "Control", id: "CONTROL" },
        { sheet: "AUTORIDADES", col: "Autoridades", id: "ID_AUT" },
        { sheet: "AUXILIARES", col: "Auxiliares", id: "ID_AUX" },
        { sheet: "USUARIOS", col: "Usuarios", id: "ID_USER" }
    ];
    for (const item of hojasAMigrar) {
        migrateAnySheet(item.sheet, item.col, item.id);
    }
}
function migrateAnySheet(sheetName, collectionName, idColumnName) {
    const ss = getSS();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet)
        return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2)
        return;
    const headers = data[0].map(normalizeHeader);
    const targetIdHeader = normalizeHeader(idColumnName);
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowObject = {};
        headers.forEach((h, idx) => { if (h)
            rowObject[h] = row[idx]; });
        let docId = String(rowObject[targetIdHeader] || "").trim();
        if (docId.endsWith(".0"))
            docId = docId.substring(0, docId.length - 2);
        if (!docId)
            docId = "row_" + i;
        firestoreCreateDocumentWithId(collectionName, docId, rowObject);
    }
}
function onEdit(e) {
    if (!e)
        return;
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    if (row === 1)
        return;
    const mapping = {
        "REGISTROS": { col: "Registros", idCol: "ID_REG" },
        "AUTORIDADES": { col: "Autoridades", idCol: "ID_AUT" },
        "USUARIOS": { col: "Usuarios", idCol: "ID_USER" },
        "REQUERIMIENTOS_CD": { col: "Requerimientos", idCol: "ID_REQ" },
        "DOCUMENTOS_CD": { col: "Documentos", idCol: "DOC_ID" },
        "AUXILIARES": { col: "Auxiliares", idCol: "ID_AUX" },
        "CONTROL": { col: "Control", idCol: "CONTROL" }
    };
    const config = mapping[sheetName];
    if (!config)
        return;
    try {
        const fullData = sheet.getDataRange().getValues();
        const headers = fullData[0].map(normalizeHeader);
        const rowValues = fullData[row - 1];
        const rowObject = {};
        headers.forEach((h, idx) => { if (h)
            rowObject[h] = rowValues[idx]; });
        const targetIdHeader = normalizeHeader(config.idCol);
        let docId = String(rowObject[targetIdHeader] || "").trim();
        if (docId.endsWith(".0"))
            docId = docId.substring(0, docId.length - 2);
        if (docId)
            firestoreCreateDocumentWithId(config.col, docId, rowObject);
    }
    catch (err) { }
}
