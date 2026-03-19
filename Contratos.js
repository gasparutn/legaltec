// --- Funciones de acción específicas ---


function generaContratoWordAnsioso() {

  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.alert('GENERANDO DOCUMENTOS, pulse  >> SI <<  para continuar', ui.ButtonSet.YES_NO);

  if (respuesta == ui.Button.YES) {
    //obtener documento origen
    var DocActual2 = DriveApp.getFileById('1-pyl7Y0Z1vGdSmj2FDAueQdqLAbmTHO3hIJbPF5bQ3E');
    var fila = 3;
    var nombreCelda2 = 'B' + fila;
    var celdaActual2 = HOJA_FUSION.getRange(nombreCelda2);
    var documentosGenerados = 0;

    while (!celdaActual2.isBlank()) {

      if (HOJA_FUSION.getRange('S' + fila).getValue() != true) { // Era R
        documentosGenerados++;
        //cuando sea el primer documento generados
        if (documentosGenerados >= 1) {
          //creamos nuevo documento
          var docNuevo1 = DocActual2.makeCopy("CONTRATO " + HOJA_FUSION.getRange('D' + fila).getValue() + "-" + HOJA_FUSION.getRange('B' + fila).getValue() + "_" + HOJA_FUSION.getRange('O' + fila).getValue()); // N -> O
          var documento = DocumentApp.openById(docNuevo1.getId());
          var fechaT = ObtenerTextofecha();
        }
        //reemplazamos los datos
        documento.getBody().replaceText("<<SEXO>>", HOJA_FUSION.getRange('F' + fila).getDisplayValue().trim());
        documento.getBody().replaceText("<<APELLIDOS>>", HOJA_FUSION.getRange('G' + fila).getDisplayValue().trim());
        documento.getBody().replaceText("<<NOMBRES>>", HOJA_FUSION.getRange('H' + fila).getDisplayValue().trim());
        documento.getBody().replaceText("<<DNI>>", formatearDNI(HOJA_FUSION.getRange('I' + fila).getDisplayValue().trim()));
        documento.getBody().replaceText("<<TAREAS>>", HOJA_FUSION.getRange('J' + fila).getDisplayValue().trim());
        documento.getBody().replaceText("<<FECHA_ALTA>>", HOJA_FUSION.getRange('K' + fila).getDisplayValue().trim());
        documento.getBody().replaceText("<<CUOTA>>", HOJA_FUSION.getRange('L' + fila).getDisplayValue().trim());
        documento.getBody().replaceText("<<CUOTAS>>", HOJA_FUSION.getRange('L' + fila).getDisplayValue().trim());
        documento.getBody().replaceText("<<MONTO_TOTAL>>", HOJA_FUSION.getRange('M' + fila).getDisplayValue().trim());
        documento.getBody().replaceText("<<DOMICILIO>>", HOJA_FUSION.getRange('N' + fila).getDisplayValue().trim());
        documento.getBody().replaceText("<<LOCALIDAD>>", HOJA_FUSION.getRange('O' + fila).getDisplayValue().trim());

        // añadimos el check
        //HOJA_FUSION.getRange('P' + fila).setValue('SIN NOTA');
        HOJA_FUSION.getRange('O' + fila).insertCheckboxes().setValue('true');
        HOJA_FUSION.getRange('R' + fila).insertCheckboxes();
        //marcamos el check
        HOJA_FUSION.getRange('S' + fila).setValue('true'); // S (era R)
        //añadimos la fecha
        HOJA_FUSION.getRange('T' + fila).setValue(new Date()); // T (era S)
        HOJA_FUSION.getRange('X' + fila).setValue("INGRESE FECHA"); // X (era W)
        HOJA_FUSION.getRange('Y' + fila).insertCheckboxes().setValue('true'); // Y (era X)

        //convertimos a Word (.docx) usando la API de Drive
        documento.saveAndClose();

        var fileId = docNuevo1.getId();
        var url = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document`;
        var params = {
          method: "GET",
          headers: {
            "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
          },
          muteHttpExceptions: true,
        };
        var response = UrlFetchApp.fetch(url, params);
        var docxBlob = response.getBlob();

        const folder3 = DriveApp.getFolderById(CARPETA_FUSION_CONTRATO);
        folder3.createFile(docxBlob).setName("CONTRATO " + HOJA_FUSION.getRange('D' + fila).getValue() + "-" + HOJA_FUSION.getRange('B' + fila).getValue() + "_" + HOJA_FUSION.getRange('O' + fila).getValue() + ".docx"); // N -> O

        // Eliminar el documento de Google Docs original (opcional)
        DriveApp.getFileById(docNuevo1.getId()).setTrashed(true);
      }

      fila++;
      nombreCelda2 = 'B' + fila;
      celdaActual2 = HOJA_FUSION.getRange(nombreCelda2);
    }

    //mensaje final
    if (documentosGenerados > 0 && documentosGenerados < 2) {
      ui.alert('DOUMENTO GENERADO ' + documentosGenerados + ' "FINALIZADO"');
    } else if (documentosGenerados > 1) {
      ui.alert('DOUMENTOS GENERADOS ' + documentosGenerados + ' "FINALIZADO"');
    } else {
      ui.alert('No se han encontrado datos para generar documentos');
    }
  }

  function ObtenerTextofecha() {
    var fecha = new Date();
    var mes = fecha.getMonth();
    var dia = fecha.getDate();
    var anyo = fecha.getFullYear();

    switch (mes) {
      case 0: mes = "enero"; break;
      case 1: mes = "febrero"; break;
      case 2: mes = "marzo"; break;
      case 3: mes = "abril"; break;
      case 4: mes = "mayo"; break;
      case 5: mes = "junio"; break;
      case 6: mes = "julio"; break;
      case 7: mes = "agosto"; break;
      case 8: mes = "septiembre"; break;
      case 9: mes = "octubre"; break;
      case 10: mes = "noviembre"; break;
      case 11: mes = "diciembre"; break;
    }
    return dia + " de " + mes + " del " + anyo;
  }

  function ObtenerTextofecha2() {
    var fecha = new Date();
    var mes2 = fecha.getMonth();
    var dia2 = fecha.getDate();
    var anyo2 = fecha.getFullYear();

    mes2 = fecha.getMonth() + 1;
    mes2 = mes2 < 10 ? '0' + mes2 : mes2;
    dia2 = dia2 < 10 ? '0' + dia2 : dia2;

    return anyo2 + "-" + mes2 + "-" + dia2;
  }
}

function generaContratoWord(rowData) {
  Logger.log("Iniciando generaContratoWord() para: " + rowData['APELLIDOS']);

  // Helper: convierte cualquier valor de fecha a "D de mes de YYYY" en español
  function fechaEnEspaniol(valor) {
    if (!valor) return '';
    const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    let d;
    if (valor instanceof Date) {
      d = valor;
    } else {
      const s = String(valor).trim();
      // Ya está en formato español: "1 de marzo de 2026"
      if (/\d{1,2}\s+de\s+\w+\s+de\s+\d{4}/i.test(s)) return s;
      // Formato YYYY-MM-DD
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
        const p = s.split('-');
        d = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
        // Formato DD/MM/YYYY
      } else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) {
        const p = s.split('/');
        d = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
      } else {
        // Intentar parseo genérico
        d = new Date(s);
      }
    }
    if (!d || isNaN(d.getTime())) return String(valor);
    return d.getDate() + ' de ' + MESES[d.getMonth()] + ' de ' + d.getFullYear();
  }

  // Extract variables (Trimmed and Formatted)
  const sexo = String(rowData['SEXO'] || '').trim();
  const nombres = String(rowData['NOMBRES'] || '').trim();
  const apellidos = String(rowData['APELLIDOS'] || '').trim();
  const dni = formatearDNI(String(rowData['DNI'] || '').trim());
  const tareas = String(rowData['TAREAS'] || '').trim();
  const fechaAlta = fechaEnEspaniol(rowData['FECHA ALTA']);  // ← Conversión correcta
  const cuota = String(rowData['CUOTA'] || '').trim();
  const totalRaw = rowData['TOTAL'];
  const domicilio = String(rowData['DOMICILIO'] || '').trim();
  const localidad = String(rowData['LOCALIDAD'] || '').trim();
  const autoridad = String(rowData['AUTORIDAD'] || '').trim();
  const autor = String(rowData['AUTOR'] || '').trim();

  Logger.log('fechaAlta convertida: "' + fechaAlta + '"');

  // AUX3 = denominado / denominada  (basado en sexo)
  const aux3 = (sexo === "el Sr.") ? "denominado" : (sexo === "la Sra.") ? "denominada" : "";

  // AUX4 = EL LOCADOR / LA LOCADORA
  const aux4 = (sexo === "el Sr.") ? "EL LOCADOR" : (sexo === "la Sra.") ? "LA LOCADORA" : "";

  // AUX5 map check:
  const aux5Map = {
    "ADM": "por", "LEGISLATIVA": "por", "VICE": "por",
    "ANA": "por el Senador", "CHAPPEL": "por el Senador", "DIUMENJO": "por el Senador", "FREIDEMBERG": "por el Senador", "GONZALEZ F.": "por el Senador", "GONZALEZ V.": "por el Senador", "IGLESIAS": "por el Senador", "KERCHNER": "por el Senador", "MAGISTRETTI": "por el Senador", "MARCOLINI": "por el Senador", "MARQUEZ": "por el Senador", "PERVIU": "por el Senador", "PEZZUTTI": "por el Senador", "PRADINES": "por el Senador", "PRINGLES": "por el Senador", "QUATTRINI": "por el Senador", "ROSTAND": "por el Senador", "SAEZ": "por el Senador", "SAT": "por el Senador", "SERRA": "por el Senador", "SEVILLA": "por el Senador", "SOTO": "por el Senador", "VAQUER": "por el Senador",
    "ASES": "por la Senadora", "BARRO": "por la Senadora", "CANO": "por la Senadora", "DERRACHE": "por la Senadora", "EISENCHLAS": "por la Senadora", "FLORIDIA": "por la Senadora", "GALIÑARES": "por la Senadora", "GOMEZ": "por la Senadora", "LAFERTE": "por la Senadora", "MANONI": "por la Senadora", "NAJUL": "por la Senadora", "SABADIN": "por la Senadora", "SAINZ": "por la Senadora", "VICCHI": "por la Senadora", "ZLOBEC": "por la Senadora"
  };
  const aux5 = aux5Map[autoridad] || "";

  // AUX7
  let aux7 = "";
  if (autoridad !== "ADM" && autoridad !== "LEGISLATIVA" && autoridad !== "VICE") {
    if (aux5 === "por el Senador") aux7 = "el Senador";
    else if (aux5 === "por la Senadora") aux7 = "la Senadora";
  }

  // AUX8
  let aux8 = "";
  if (["VICE", "LEGISLATIVA", "ADM"].includes(autoridad)) {
    aux8 = "de";
  } else if (aux5 === "por el Senador") {
    aux8 = "del Senador";
  } else if (aux5 === "por la Senadora") {
    aux8 = "de la Senadora";
  }

  // AUX9
  const aux9Map = {
    "ADM": "SECRETARIA ADMINISTRATIVA", "ANA": "ANA MARIO ESTEBAN", "ASES": "ASES YAMEL", "BARRO": "BARRO JOHANA", "CANO": "CANO ADRIANA", "CHAPPEL": "CHAPPEL DUGAR", "DERRACHE": "DERRACHE MARIA", "DIUMENJO": "DIUMENJO ALEJANDRO", "EISENCHLAS": "EISENCHLAS NATALIA", "FLORIDIA": "FLORIDIA ANGELA", "FREIDEMBERG": "FREIDEMBRERG ABEL", "GALIÑARES": "GALIÑARES MARÍA", "GOMEZ": "GÓMEZ OLGA CRISTINA", "GONZALEZ F.": "GONZÁLEZ FELIX", "GONZALEZ V.": "GONZÁLEZ VALENTIN", "IGLESIAS": "IGLESIAS MARCELINO", "KERCHNER": "KERCHNER MARTIN", "LAFERTE": "LAFERTE JESICA", "LEGISLATIVA": "SECRETARIA LEGISLATIVA", "MAGISTRETTI": "MAGISTRETTI ARMANDO", "MANONI": "MANONI FLAVIA", "MARCOLINI": "MARCOLINI WALTHER", "MARQUEZ": "MARQUEZ SERGIO", "NAJUL": "NAJUL CLAUDIA", "PERVIU": "PERVIU HELIO", "PEZZUTTI": "PEZZUTTI DULIO", "PRADINES": "PRADINES GABRIEL", "PRINGLES": "PRINGLES ARIEL", "QUATTRINI": "QUATTRINI MARCOS", "ROSTAND": "ROSTAND MARTÍN", "SABADIN": "SABADIN MARIA FERNANDA", "SAEZ": "SÁEZ DAVID", "SAINZ": "SAINZ MARÍA LAURA", "SAT": "SAT MAURICIO", "SERRA": "SERRA PEDRO", "SEVILLA": "SEVILLA OSCAR", "SOTO": "SOTO OSCAR", "VAQUER": "VAQUER GERARDO", "VICCHI": "VICCHI GERMAN", "VICE": "VICEGOBERNACION", "ZLOBEC": "ZLOBEC LEIVA MARIANA G."
  };
  const aux9 = aux9Map[autoridad] || "";

  // AUX10
  const aux10Map = {
    "1": "un", "01": "un", "2": "dos", "02": "dos", "3": "tres", "03": "tres", "4": "cuatro", "04": "cuatro", "5": "cinco", "05": "cinco", "6": "seis", "06": "seis", "7": "siete", "07": "siete", "8": "ocho", "08": "ocho", "9": "nueve", "09": "nueve", "10": "diez", "11": "once", "12": "doce"
  };
  const aux10 = (aux10Map[cuota] || "").trim();

  // AUX11
  const aux11 = (aux10 === "") ? "" : (aux10 === "un" ? "mes" : "meses");

  // AUX12
  const aux12Map = {
    "1": "una", "01": "una", "2": "dos", "02": "dos", "3": "tres", "03": "tres", "4": "cuatro", "04": "cuatro", "5": "cinco", "05": "cinco", "6": "seis", "06": "seis", "7": "siete", "07": "siete", "8": "ocho", "08": "ocho", "9": "nueve", "09": "nueve", "10": "diez", "11": "once", "12": "doce"
  };
  const aux12 = aux12Map[cuota] || "";

  // AUX13
  const aux13 = (aux10 === "") ? "" : (aux10 === "un" ? "cuota" : "cuotas");

  // AUX14
  const aux14 = (aux13 === "") ? "" : (aux13 === "cuotas" ? "cada una efectivizadas a mes vencido abonándose la primera en el mes de" : "efectivizada a mes vencido, abonándose en");

  // AUX16 (FECHA ALTA -> Next month text)
  let aux16 = "";
  const fechaAltaLower = fechaAlta.toLowerCase();
  if (fechaAltaLower.includes("enero")) aux16 = "febrero de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("febrero")) aux16 = "marzo de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("marzo")) aux16 = "abril de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("abril")) aux16 = "mayo de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("mayo")) aux16 = "junio de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("junio")) aux16 = "julio de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("julio")) aux16 = "agosto de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("agosto")) aux16 = "septiembre de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("septiembre")) aux16 = "octubre de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("octubre")) aux16 = "noviembre de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("noviembre")) aux16 = "diciembre de " + (fechaAltaLower.match(/\d{4}/)?.[0] || "");
  else if (fechaAltaLower.includes("diciembre")) {
    let year = parseInt(fechaAltaLower.match(/\d{4}/)?.[0] || "0");
    aux16 = "enero de " + (year ? (year + 1) : "");
  }

  // numericTotal
  let numericTotal = 0;
  if (totalRaw) {
    if (typeof totalRaw === 'string') {
      numericTotal = parseFloat(totalRaw.replace(/[^0-9,-]+/g, "").replace(",", "."));
    } else {
      numericTotal = parseFloat(totalRaw);
    }
  }

  // AUX17
  let aux17 = "";
  if (!isNaN(numericTotal) && numericTotal > 0) {
    if (typeof NumeroALetras === 'function') {
      aux17 = NumeroALetras(numericTotal).trim();
    }
  }

  // AUX18
  const cuotaNum = parseInt(cuota, 10);
  let divResult = "";
  if (!isNaN(numericTotal) && !isNaN(cuotaNum) && cuotaNum > 0) {
    divResult = (numericTotal / cuotaNum).toFixed(2);
  }

  // Format AUX18 string
  let aux18Str = "";
  if (divResult !== "") {
    let cVal = Number(divResult);
    aux18Str = "$" + cVal.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // AUX19
  let aux19 = "";
  if (divResult !== "") {
    if (typeof NumeroALetras === 'function') {
      aux19 = NumeroALetras(Number(divResult)).trim();
    }
  }

  // AUX22
  const aux22 = (sexo === "el Sr.") ? "obligado" : (sexo === "la Sra.") ? "obligada" : "";

  // AUX24
  const aux24 = (sexo === "el Sr.") ? "éste" : (sexo === "la Sra.") ? "ésta" : "";

  // Custom format total
  let totalFormateado = "";
  if (!isNaN(numericTotal) && numericTotal > 0) {
    totalFormateado = "$" + numericTotal.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // ID Documento base
  const idPlantilla = '1-pyl7Y0Z1vGdSmj2FDAueQdqLAbmTHO3hIJbPF5bQ3E';

  // Create file
  const docBase = DriveApp.getFileById(idPlantilla);
  const nombreArchivo = "CONTRATO " + apellidos + "-" + autoridad + "_" + autor;
  const carpetaPDF = DriveApp.getFolderById(CARPETA_FUSION_CONTRATO);

  const nuevoDocFile = docBase.makeCopy(nombreArchivo, carpetaPDF);
  const nuevoDocId = nuevoDocFile.getId();
  const documento = DocumentApp.openById(nuevoDocId);
  const body = documento.getBody();

  // replaceText
  // AUX10 y CUOTA: se usa el patrón '[ ]+' para consumir espacios extra
  // que la plantilla pueda tener después del marcador (el motor regex de Apps Script
  // trata [ ] como un espacio literal seguro, a diferencia de \s o \s+).
  // Si la plantilla tiene "<<AUX10>>  <<CUOTA>>" (2 espacios), el primer reemplazo
  // consume ambos espacios y reinserta exactamente uno. Si hay solo 1 espacio,
  // el patrón también funciona y el resultado es correcto.
  body.replaceText("<<AUX10>>[ ]+", aux10 + " ");
  body.replaceText("<<AUX10>>", aux10); // fallback si no hay espacio después

  body.replaceText("<<AUX19>>[ ]+", aux19 + " ");
  body.replaceText("<<AUX19>>", aux19);

  body.replaceText("<<CUOTA>>[ ]+", cuota + " ");
  body.replaceText("<<CUOTA>>", cuota);
  body.replaceText("<<CUOTAS>>[ ]+", cuota + " ");
  body.replaceText("<<CUOTAS>>", cuota);

  body.replaceText("<<SEXO>>", sexo);
  body.replaceText("<<NOMBRES>>", nombres);
  body.replaceText("<<APELLIDOS>>", apellidos);
  body.replaceText("<<DNI>>", dni);
  body.replaceText("<<TAREAS>>", tareas);
  body.replaceText("<<FECHA_ALTA>>", fechaAlta);
  body.replaceText("<<AUX3>>", aux3);
  body.replaceText("<<AUX4>>", aux4);
  body.replaceText("<<AUX5>>", aux5);
  body.replaceText("<<AUX7>>", aux7);
  body.replaceText("<<AUX8>>", aux8);
  body.replaceText("<<AUX9>>", aux9);
  body.replaceText("<<AUX11>>", aux11);
  body.replaceText("<<AUX12>>", aux12);
  body.replaceText("<<AUX13>>", aux13);
  body.replaceText("<<AUX14>>", aux14);
  body.replaceText("<<AUX16>>", aux16);
  body.replaceText("<<AUX17>>", aux17);
  body.replaceText("<<AUX18>>", aux18Str);
  body.replaceText("<<AUX22>>", aux22);
  body.replaceText("<<AUX24>>", aux24);
  body.replaceText("<<TOTAL>>", totalFormateado);
  body.replaceText("<<MONTO_TOTAL>>", totalFormateado);
  body.replaceText("<<DOMICILIO>>", domicilio);
  body.replaceText("<<LOCALIDAD>>", localidad);

  documento.saveAndClose();

  // Pausa breve antes de exportar
  Utilities.sleep(1000);

  // Exportar a Word (.docx) via Drive API v3 — único método soportado para Google Docs → Word
  var token = ScriptApp.getOAuthToken();
  var exportUrl = 'https://www.googleapis.com/drive/v3/files/' + nuevoDocId +
    '/export?mimeType=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document';

  var fetchResponse = UrlFetchApp.fetch(exportUrl, {
    method: 'GET',
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  var respCode = fetchResponse.getResponseCode();
  Logger.log('Drive export HTTP code: ' + respCode);

  if (respCode !== 200) {
    try { DriveApp.getFileById(nuevoDocId).setTrashed(true); } catch (e2) { }
    Logger.log('Respuesta error: ' + fetchResponse.getContentText().substring(0, 300));
    throw new Error('Error exportando a Word (HTTP ' + respCode + ').');
  }

  // Crear el .docx en la carpeta destino
  var docxBlob = fetchResponse.getBlob();
  docxBlob.setName(nombreArchivo + '.docx');

  Logger.log('Tamaño del blob exportado: ' + docxBlob.getBytes().length + ' bytes');

  var archivoFinal = carpetaPDF.createFile(docxBlob);
  var archivoFinalUrl = archivoFinal.getUrl();

  Logger.log('Contrato .docx creado exitosamente.');
  Logger.log('  Nombre : ' + archivoFinal.getName());
  Logger.log('  ID     : ' + archivoFinal.getId());
  Logger.log('  Carpeta: ' + CARPETA_FUSION_CONTRATO);
  Logger.log('  URL    : ' + archivoFinalUrl);

  // Eliminar el Google Doc temporal
  DriveApp.getFileById(nuevoDocId).setTrashed(true);

  return { url: archivoFinalUrl, nombre: nombreArchivo };
}
