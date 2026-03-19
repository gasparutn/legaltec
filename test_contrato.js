/**
 * FUNCIÓN DE PRUEBA - Ejecutar manualmente desde el editor de Apps Script
 * para diagnosticar la generación de contratos sin depender del frontend.
 * 
 * INSTRUCCIONES:
 * 1. Abrí el editor de Apps Script (script.google.com)
 * 2. Seleccioná la función "testGeneraContrato" en el menú desplegable
 * 3. Hacé click en ▶ Ejecutar
 * 4. Revisá el registro de ejecución (Ver > Registros)
 */
function testGeneraContrato() {
    Logger.log('=== TEST generaContratoWord ===');

    // Datos de prueba — podés cambiarlos por datos reales de tu Sheet
    const rowDataTest = {
        'SEXO': 'la Sra.',
        'NOMBRES': 'MARIA',
        'APELLIDOS': 'GARCIA',
        'DNI': '12345678',
        'TAREAS': 'de asesoramiento legislativo',
        'FECHA ALTA': '1 de marzo de 2026',
        'CUOTA': '3',
        'TOTAL': 150000,
        'DOMICILIO': 'Av. San Martín 123',
        'LOCALIDAD': 'Capital, Mendoza',
        'AUTORIDAD': 'SERRA',
        'AUTOR': 'JU',
    };

    Logger.log('Llamando generaContratoWord con datos de prueba...');
    try {
        const resultado = generaContratoWord(rowDataTest);
        Logger.log('ÉXITO. Nombre: ' + resultado.nombre);
        Logger.log('URL: ' + resultado.url);
    } catch (e) {
        Logger.log('ERROR: ' + e.message);
        Logger.log('Stack: ' + e.stack);
    }

    Logger.log('=== FIN TEST ===');
}

/**
 * Verifica que la carpeta de contratos existe y está accesible.
 */
function testCarpetaContrato() {
    Logger.log('=== TEST carpeta contratos ===');
    try {
        const carpeta = DriveApp.getFolderById(CARPETA_FUSION_CONTRATO);
        Logger.log('Carpeta encontrada: ' + carpeta.getName());
        Logger.log('ID: ' + carpeta.getId());

        const archivos = carpeta.getFiles();
        let count = 0;
        while (archivos.hasNext()) {
            const f = archivos.next();
            count++;
            Logger.log('  Archivo ' + count + ': ' + f.getName() + ' (' + f.getMimeType() + ')');
            if (count >= 5) { Logger.log('  (mostrando solo primeros 5)'); break; }
        }
        if (count === 0) Logger.log('  La carpeta está vacía.');
    } catch (e) {
        Logger.log('ERROR accediendo a la carpeta: ' + e.message);
    }
    Logger.log('=== FIN TEST ===');
}
