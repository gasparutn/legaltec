// Notifica.js — Funciones de notificación por correo

function notificaLocaciones1(rowData) {
  Logger.log("Función notificaLocaciones1() ejecutada para fila: " + JSON.stringify(rowData));

  // Define el asunto predeterminado para el correo de notificación
  const ASUNTO_NOTIFICACION = "Notifica Contrato";

  try {
    // Verificar si los datos necesarios existen en rowData
    // Usar los nombres de cabecera con la capitalización exacta
    // 'CHECK' es la columna T (índice 19 en un array 0-indexado)
    // 'CORREO SECRETARIO' es la columna D (índice 3)
    // 'NOMBRES' es la columna F (índice 5)
    // 'SECRETARIO' es la columna C (índice 2)
    const isChecked = rowData['CHECK'];
    const correoSecretario = rowData['CORREO SECRETARIO'];
    const nombreLocador = (rowData['APELLIDOS'] ? rowData['APELLIDOS'] + ' ' : '') + (rowData['NOMBRES'] || '');
    const secretario = rowData['SECRETARIO'];
    const estado = rowData['ESTADO'];

    // 1) Unicamente se envia notificación si el estado de la columna U del sheet esta en estado Resolución o Notificado
    if (!estado || (!estado.toLowerCase().includes('resoluci') && !estado.toLowerCase().includes('notificad'))) {
      let errorMessage = "El estado actual no es 'Resolución' ni 'Notificado' (" + estado + "). No se envía notificación.";
      Logger.log(errorMessage);
      return { success: false, message: errorMessage };
    }

    // Solo procede si el checkbox NO está marcado y los campos de correo, locador y secretario no están vacíos
    if (isChecked !== true && correoSecretario && nombreLocador && secretario) {
      // Crea la plantilla HTML para el cuerpo del correo
      let templateName = "htmlNotificaAlta";
      if (String(rowData['REQ'] || '').toLowerCase().includes('baja')) {
        templateName = "htmlNotificaBaja";
      }
      const templateHtml = HtmlService.createTemplateFromFile(templateName);
      templateHtml.secretarios = secretario;
      templateHtml.nombreLocador = nombreLocador;

      const mensajeHtml = templateHtml.evaluate().getContent();

      // Configuración del correo
      let options = {
        bcc: 'jmaizdeaurelli@legislaturamendoza.gov.ar',
        cc: 'legaltec.locaciones@legislaturamendoza.gov.ar',
        htmlBody: mensajeHtml
      };

      // Manejo del remitente (from)
      const aliases = GmailApp.getAliases();
      const desiredAlias = 'legaltec.locaciones@legislaturamendoza.gov.ar';

      if (aliases.includes(desiredAlias)) {
        options.from = desiredAlias;
      } else if (aliases.length > 0) {
        options.from = aliases[0];
      }

      // Envía el correo
      GmailApp.sendEmail(
        correoSecretario,
        ASUNTO_NOTIFICACION,
        "Este es un mensaje de texto plano de respaldo...", // Mensaje texto
        options
      );

      Logger.log("Correo enviado exitosamente a: " + correoSecretario);
      return { success: true, message: `Notificación enviada a ${secretario} (${correoSecretario}).` };
    } else {
      let errorMessage = "Faltan datos o el checkbox ya está marcado para enviar la notificación: ";
      if (isChecked === true) errorMessage += "El checkbox ya está marcado. ";
      if (!correoSecretario) errorMessage += "Correo del Secretario vacío. ";
      if (!nombreLocador) errorMessage += "Nombre del Locador vacío. ";
      if (!secretario) errorMessage += "Nombre del Secretario vacío. ";
      Logger.log(errorMessage);
      return { success: false, message: errorMessage };
    }
  } catch (e) {
    Logger.log("Error en notificaLocaciones1(): " + e.message);
    return { success: false, message: "Error al enviar notificación: " + e.message };
  }
}