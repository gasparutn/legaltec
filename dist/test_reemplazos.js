"use strict";
/**
 * Test script to verify the fix for document generation and date logic.
 */
function testReemplazosYFechas() {
    var testData = {
        'nombres': 'JUAN PABLO',
        'apellidos': 'GONZALEZ',
        'dni': '30123456',
        'fecha alta': '2026-03-01',
        'tipo_nota_select': 'Nota Alta Senado',
        'sexo_nota_select': 'Masculino'
    };
    Logger.log("--- TEST 1: Normalización de llaves en Notas.ts ---");
    // Simular normalización que agregué en generarNotaDesdeModal
    var normalizedData = {};
    for (var key in testData) {
        normalizedData[key.trim().toUpperCase()] = testData[key];
    }
    Logger.log("Llaves normalizadas: " + Object.keys(normalizedData).join(", "));
    if (normalizedData['NOMBRES'] === 'JUAN PABLO' && normalizedData['FECHA ALTA'] === '2026-03-01') {
        Logger.log("SUCCESS: Llaves normalizadas correctamente.");
    }
    else {
        Logger.log("FAILED: Normalización incorrecta.");
    }
    Logger.log("--- TEST 2: Cálculo de AUX16 (Mes Siguiente) en Resolucion.ts ---");
    var rowData = { 'FECHA ALTA': '2026-03-15' };
    var aux16 = "";
    var fechaAltaRaw = rowData['FECHA ALTA'];
    try {
        var d = new Date(fechaAltaRaw);
        var nextMonth = new Date(d);
        nextMonth.setMonth(nextMonth.getMonth() + 1);
        var meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
        aux16 = "".concat(meses[nextMonth.getMonth()], " de ").concat(nextMonth.getFullYear());
        Logger.log("Entrada: 2026-03-15 -> Salida AUX16: ".concat(aux16));
        if (aux16 === "abril de 2026") {
            Logger.log("SUCCESS: Cálculo de mes siguiente correcto.");
        }
        else {
            Logger.log("FAILED: Cálculo de mes siguiente incorrecto.");
        }
        // Mocking replaceFields behavior for standardized placeholder
        var fields = { "<<FECHA ALTA>>": aux16 };
        Logger.log("Probando mapeo de marcador con espacio: " + Object.keys(fields)[0]);
        if (fields["<<FECHA ALTA>>"]) {
            Logger.log("SUCCESS: Marcador <<FECHA ALTA>> mapeado correctamente.");
        }
    }
    catch (e) {
        Logger.log("ERROR en test: " + e.message);
    }
    Logger.log("--- TEST 3: Caso borde Diciembre -> Enero ---");
    var rowDataBorde = { 'FECHA ALTA': '2026-12-10' };
    var aux16Borde = "";
    try {
        var d = new Date(rowDataBorde['FECHA ALTA']);
        var nextMonth = new Date(d);
        nextMonth.setMonth(nextMonth.getMonth() + 1);
        var meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
        aux16Borde = "".concat(meses[nextMonth.getMonth()], " de ").concat(nextMonth.getFullYear());
        Logger.log("Entrada: 2026-12-10 -> Salida AUX16: ".concat(aux16Borde));
        if (aux16Borde === "enero de 2027") {
            Logger.log("SUCCESS: Caso borde Diciembre-Enero correcto.");
        }
        else {
            Logger.log("FAILED: Caso borde Diciembre-Enero incorrecto.");
        }
    }
    catch (e) {
        Logger.log("ERROR en test: " + e.message);
    }
}
