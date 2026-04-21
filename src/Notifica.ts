function notificaLocaciones1(rowData: {[key: string]: any}): { success: boolean; message: string } {
  Logger.log("Función notificaLocaciones1() ejecutada para fila: " + JSON.stringify(rowData));

  const ASUNTO_NOTIFICACION = "Notifica Contrato";

  try {
    const isChecked = rowData['CHECK'];
    const correoSecretario = rowData['CORREO SECRETARIO'];
    const nombreLocador = (rowData['APELLIDOS'] ? rowData['APELLIDOS'] + ' ' : '') + (rowData['NOMBRES'] || '');
    const secretario = rowData['SECRETARIO'];
    const estado = rowData['ESTADO'];

    if (!estado || (!estado.toLowerCase().includes('resoluci') && !estado.toLowerCase().includes('notificad'))) {
      const errorMessage = "El estado actual no es 'Resolución' ni 'Notificado' (" + estado + "). No se envía notificación.";
      Logger.log(errorMessage);
      return { success: false, message: errorMessage };
    }

    // El check ya se hizo o se bypassó en executeAction. 
    // Solo validamos que tengamos los datos mínimos para el envío.
    if (correoSecretario && nombreLocador && secretario) {
      let templateName = "htmlNotificaAlta";
      if (String(rowData['REQ'] || '').toLowerCase().includes('baja')) {
        templateName = "htmlNotificaBaja";
      }
      const templateHtml = HtmlService.createTemplateFromFile(templateName) as any;
      templateHtml.secretarios = secretario;
      templateHtml.nombreLocador = nombreLocador;

      const mensajeHtml = templateHtml.evaluate().getContent();

      const options: GoogleAppsScript.Gmail.GmailAdvancedOptions = {
        bcc: 'jmaizdeaurelli@legislaturamendoza.gov.ar',
        cc: 'legaltec.locaciones@legislaturamendoza.gov.ar',
        htmlBody: mensajeHtml
      };

      const aliases = GmailApp.getAliases();
      const desiredAlias = 'legaltec.locaciones@legislaturamendoza.gov.ar';

      if (aliases.includes(desiredAlias)) {
        options.from = desiredAlias;
      } else if (aliases.length > 0) {
        options.from = aliases[0];
      }

      GmailApp.sendEmail(
        correoSecretario,
        ASUNTO_NOTIFICACION,
        "Este es un mensaje de texto plano de respaldo...",
        options
      );

      Logger.log("Correo enviado exitosamente a: " + correoSecretario);
      logActivity(`Genera Notifica ${rowData['APELLIDOS']}_${rowData['AUTORIDAD']}`);
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
  } catch (e: any) {
    Logger.log("Error en notificaLocaciones1(): " + e.message);
    return { success: false, message: "Error al enviar notificación: " + e.message };
  }
}