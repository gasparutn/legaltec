"use strict";
/**
 * ============================================================
 *  ManualUtils.ts  —  Funciones de Ejecución Manual
 * ============================================================
 *  ATENCIÓN: Estas funciones se ejecutan MANUALMENTE desde el
 *  editor de Google Apps Script (▶ Ejecutar).
 *  NO se invocan desde el frontend ni de forma automática.
 *
 *  Cómo usarlas:
 *  1. Abrí el proyecto en script.google.com
 *  2. Seleccioná la función en el menú desplegable superior
 *  3. Hacé clic en ▶ Ejecutar
 *  4. Revisá los logs en Ver > Registros (Ctrl+Enter)
 * ============================================================
 */
// ─────────────────────────────────────────────────────────────
//  COLOR DE ESTADO EN HOJA DE CÁLCULO
//  (Sincroniza los colores de fondo de la columna ESTADO
//   en la hoja "REGISTROS" de Google Sheets, de forma que
//   coincidan con la paleta visual del frontend.)
// ─────────────────────────────────────────────────────────────
/**
 * Aplica el color de fondo y texto de la columna ESTADO en
 * la hoja "REGISTROS" para TODOS los registros existentes.
 *
 * Paleta sincronizada con Stylesheet.html (.status-*):
 *   Pendiente   → fondo #fc0404ff  texto #ffffffff
 *   Contrato    → fondo rgba(108, 233, 255, 1)  texto #ffffffff
 *   Resolución  → fondo #fee925  texto #ffffffff
 *   Notificado  → fondo #67fe74ff  texto #ffffffff
 *   Baja        → fondo #8672fbff  texto #ffffff
 *   (otro)      → sin color (reset)
 *
 * Ejecutar manualmente desde el editor de Apps Script.
 */
function aplicarColorEstadoTodosLosRegistros() {
    const sheet = getMainSheet();
    if (!sheet) {
        Logger.log("❌ Hoja 'REGISTROS' no encontrada.");
        return;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const estadoColIdx = headers.findIndex(h => String(h).trim().toUpperCase() === 'ESTADO');
    if (estadoColIdx === -1) {
        Logger.log("❌ Columna 'ESTADO' no encontrada en la hoja.");
        return;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        Logger.log("ℹ️ No hay registros para procesar.");
        return;
    }
    // Colum number (1-indexed)
    const estadoCol = estadoColIdx + 1;
    const range = sheet.getRange(2, estadoCol, lastRow - 1, 1);
    const valores = range.getValues();
    // Paleta de colores — sincronizada con Stylesheet.html
    const PALETA = {
        'PENDIENTE': { bg: '#fc0404', fg: '#ffffff' },
        'CONTRATO': { bg: '#6ce9ff', fg: '#000000' },
        'RESOLUCION': { bg: '#fee925', fg: '#000000' },
        'RESOLUCIÖN': { bg: '#fee925', fg: '#000000' },
        'RESOLUCIÓN': { bg: '#fee925', fg: '#000000' },
        'NOTIFICADO': { bg: '#67fe74', fg: '#000000' },
        'BAJA': { bg: '#8672fb', fg: '#ffffff' },
    };
    let actualizados = 0;
    let reseteados = 0;
    for (let i = 0; i < valores.length; i++) {
        const celda = sheet.getRange(i + 2, estadoCol);
        const valor = String(valores[i][0] || '').trim().toUpperCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // quitar tildes
        const palette = PALETA[valor];
        if (palette) {
            celda.setBackground(palette.bg);
            celda.setFontColor(palette.fg);
            celda.setFontWeight('bold');
            actualizados++;
        }
        else {
            // Resetear a sin color si el estado no coincide con ninguno conocido
            celda.setBackground(null);
            celda.setFontColor(null);
            celda.setFontWeight('normal');
            reseteados++;
        }
    }
    SpreadsheetApp.flush();
    Logger.log(`✅ Colores de ESTADO aplicados: ${actualizados} actualizados, ${reseteados} reseteados.`);
}
// ─────────────────────────────────────────────────────────────
//  REASIGNAR IDs FALTANTES
//  Recorre todos los registros en Firebase y asigna un ID
//  secuencial a los que no tienen uno.
// ─────────────────────────────────────────────────────────────
/**
 * Revisa todos los documentos en la colección "Registros" de
 * Firebase y asigna un campo 'ID' a los que no lo tienen,
 * continuando desde el máximo ID existente.
 * Ejecutar manualmente.
 */
function reasignarIDsFaltantes() {
    try {
        const records = firestoreGetAllDocs("Registros");
        if (!records || records.length === 0) {
            Logger.log("ℹ️ No hay registros en Firebase.");
            return;
        }
        let maxId = 0;
        const sinId = [];
        records.forEach(r => {
            const id = parseInt(r['ID'] || r['ID_REG'] || 0, 10);
            if (!isNaN(id) && id > maxId)
                maxId = id;
            if (!r['ID'] && !r['ID_REG'])
                sinId.push(r);
        });
        Logger.log(`ℹ️ Máximo ID actual: ${maxId}. Registros sin ID: ${sinId.length}`);
        sinId.forEach(r => {
            maxId++;
            r['ID'] = maxId;
            firestoreUpdateDocument("Registros", String(maxId), r);
            Logger.log(`✅ Asignado ID ${maxId} a: ${r['APELLIDOS'] || '?'} - ${r['AUTORIDAD'] || '?'}`);
        });
        Logger.log(`✅ Proceso finalizado. Total IDs asignados: ${sinId.length}`);
    }
    catch (e) {
        Logger.log("❌ Error en reasignarIDsFaltantes: " + e.message);
    }
}
// ─────────────────────────────────────────────────────────────
//  RECALCULAR CAMPO ESTADO EN TODOS LOS REGISTROS (Firebase)
//  Revisa el campo BAJA CONTRATO y asigna ESTADO='Baja'
//  si corresponde. Útil para migración de datos.
// ─────────────────────────────────────────────────────────────
/**
 * Recorre todos los registros en Firebase y actualiza el campo
 * ESTADO a 'Baja' cuando existe un valor en 'BAJA CONTRATO'.
 * No modifica registros que ya tienen ESTADO='Baja'.
 * Ejecutar manualmente.
 */
function recalcularEstadoBaja() {
    try {
        const records = firestoreGetAllDocs("Registros");
        if (!records || records.length === 0) {
            Logger.log("ℹ️ No hay registros para procesar.");
            return;
        }
        let actualizados = 0;
        records.forEach(r => {
            const bajaCont = String(r['BAJA CONTRATO'] || '').trim();
            const estadoActual = String(r['ESTADO'] || '').trim();
            if (bajaCont !== '' && estadoActual !== 'Baja') {
                const id = String(r['__id'] || r['ID'] || r['ID_REG'] || '');
                if (!id)
                    return;
                firestoreUpdateDocument("Registros", id, { ...r, ESTADO: 'Baja' });
                Logger.log(`✅ Registro ${id} actualizado a ESTADO='Baja' (tenía BAJA CONTRATO: ${bajaCont})`);
                actualizados++;
            }
        });
        Logger.log(`✅ recalcularEstadoBaja finalizado. Actualizados: ${actualizados}`);
    }
    catch (e) {
        Logger.log("❌ Error en recalcularEstadoBaja: " + e.message);
    }
}
// ─────────────────────────────────────────────────────────────
//  LIMPIAR CACHÉ GLOBAL
//  Borra todas las entradas de CacheService para forzar
//  recarga fresca de datos desde Firebase/Sheets.
// ─────────────────────────────────────────────────────────────
/**
 * Limpia el caché global de Google Apps Script (CacheService).
 * Útil cuando los datos se ven desactualizados en el frontend
 * a pesar de haber hecho cambios en Firebase.
 * Ejecutar manualmente.
 */
function limpiarCacheGlobal() {
    try {
        CacheService.getScriptCache().removeAll([]);
        Logger.log("✅ Caché global limpiado exitosamente.");
    }
    catch (e) {
        Logger.log("❌ Error al limpiar caché: " + e.message);
    }
}
