function generaResolucionTXT(rowData) {
  Logger.log("Iniciando generaResolucionTXT() para: " + rowData['APELLIDOS']);

  // Helper: convierte a "D de mes de YYYY" en español
  function fechaEnEspaniol(valor) {
    if (!valor) return '';
    const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    let d;
    if (valor instanceof Date) {
      d = valor;
    } else {
      const s = String(valor).trim();
      if (/\d{1,2}\s+de\s+\w+\s+de\s+\d{4}/i.test(s)) return s;
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
        const p = s.split('-');
        d = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
      } else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) {
        const p = s.split('/');
        d = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
      } else {
        d = new Date(s);
      }
    }
    if (!d || isNaN(d.getTime())) return String(valor);
    let dStr = String(d.getDate());
    if (dStr.length === 1) dStr = '0' + dStr;
    return dStr + ' de ' + MESES[d.getMonth()] + ' de ' + d.getFullYear();
  }

  // Extract variables
  const sexo = String(rowData['SEXO'] || '').trim();
  const nombres = String(rowData['NOMBRES'] || '').trim();
  const apellidos = String(rowData['APELLIDOS'] || '').trim();
  const dni = formatearDNI(String(rowData['DNI'] || '').trim());
  const tareas = String(rowData['TAREAS'] || '').trim();
  const fechaAlta = fechaEnEspaniol(rowData['FECHA ALTA']);
  const cuota = String(rowData['CUOTA'] || '').trim();
  const totalRaw = rowData['TOTAL'];
  const req = String(rowData['REQ'] || '').trim();
  const autoridad = String(rowData['AUTORIDAD'] || '').trim();
  const nroRes = String(rowData['NRO_RESOLUCION'] || 'S-NroRes').trim();

  // --- Mapeo de AUXiliares ---

  // AUX1: "con el Sr." o "con la Sra."
  const aux1 = (sexo === "el Sr.") ? "con el Sr." : (sexo === "la Sra.") ? "con la Sra." : "";

  // AUX2: "del Sr." o "de la Sra."
  const aux2 = (sexo === "el Sr.") ? "del Sr." : (sexo === "la Sra.") ? "de la Sra." : "";

  // AUX6: "por el contratado" / "por la contratada"
  const aux6 = (sexo === "el Sr.") ? "por el contratado" : (sexo === "la Sra.") ? "por la contratada" : "";

  // Necesitamos aux5Map para AUX7
  const aux5Map = {
    "ADM": "por", "LEGISLATIVA": "por", "VICE": "por",
    "ANA": "por el Senador", "CHAPPEL": "por el Senador", "DIUMENJO": "por el Senador", "FREIDEMBERG": "por el Senador", "GONZALEZ F.": "por el Senador", "GONZALEZ V.": "por el Senador", "IGLESIAS": "por el Senador", "KERCHNER": "por el Senador", "MAGISTRETTI": "por el Senador", "MARCOLINI": "por el Senador", "MARQUEZ": "por el Senador", "PERVIU": "por el Senador", "PEZZUTTI": "por el Senador", "PRADINES": "por el Senador", "PRINGLES": "por el Senador", "QUATTRINI": "por el Senador", "ROSTAND": "por el Senador", "SAEZ": "por el Senador", "SAT": "por el Senador", "SERRA": "por el Senador", "SEVILLA": "por el Senador", "SOTO": "por el Senador", "VAQUER": "por el Senador",
    "ASES": "por la Senadora", "BARRO": "por la Senadora", "CANO": "por la Senadora", "DERRACHE": "por la Senadora", "EISENCHLAS": "por la Senadora", "FLORIDIA": "por la Senadora", "GALIÑARES": "por la Senadora", "GOMEZ": "por la Senadora", "LAFERTE": "por la Senadora", "MANONI": "por la Senadora", "NAJUL": "por la Senadora", "SABADIN": "por la Senadora", "SAINZ": "por la Senadora", "VICCHI": "por la Senadora", "ZLOBEC": "por la Senadora"
  };
  const aux5 = aux5Map[autoridad] || "";

  // AUX7: el Senador / la Senadora
  let aux7 = "";
  if (autoridad !== "ADM" && autoridad !== "LEGISLATIVA" && autoridad !== "VICE") {
    if (aux5 === "por el Senador") aux7 = "el Senador";
    else if (aux5 === "por la Senadora") aux7 = "la Senadora";
  }

  // AUX9: Autoridad Completa
  const aux9Map = {
    "ADM": "SECRETARIA ADMINISTRATIVA", "ANA": "ANA MARIO ESTEBAN", "ASES": "ASES YAMEL", "BARRO": "BARRO JOHANA", "CANO": "CANO ADRIANA", "CHAPPEL": "CHAPPEL DUGAR", "DERRACHE": "DERRACHE MARIA", "DIUMENJO": "DIUMENJO ALEJANDRO", "EISENCHLAS": "EISENCHLAS NATALIA", "FLORIDIA": "FLORIDIA ANGELA", "FREIDEMBERG": "FREIDEMBRERG ABEL", "GALIÑARES": "GALIÑARES MARÍA", "GOMEZ": "GÓMEZ OLGA CRISTINA", "GONZALEZ F.": "GONZÁLEZ FELIX", "GONZALEZ V.": "GONZÁLEZ VALENTIN", "IGLESIAS": "IGLESIAS MARCELINO", "KERCHNER": "KERCHNER MARTIN", "LAFERTE": "LAFERTE JESICA", "LEGISLATIVA": "SECRETARIA LEGISLATIVA", "MAGISTRETTI": "MAGISTRETTI ARMANDO", "MANONI": "MANONI FLAVIA", "MARCOLINI": "MARCOLINI WALTHER", "MARQUEZ": "MARQUEZ SERGIO", "NAJUL": "NAJUL CLAUDIA", "PERVIU": "PERVIU HELIO", "PEZZUTTI": "PEZZUTTI DULIO", "PRADINES": "PRADINES GABRIEL", "PRINGLES": "PRINGLES ARIEL", "QUATTRINI": "QUATTRINI MARCOS", "ROSTAND": "ROSTAND MARTÍN", "SABADIN": "SABADIN MARIA FERNANDA", "SAEZ": "SÁEZ DAVID", "SAINZ": "SAINZ MARÍA LAURA", "SAT": "SAT MAURICIO", "SERRA": "SERRA PEDRO", "SEVILLA": "SEVILLA OSCAR", "SOTO": "SOTO OSCAR", "VAQUER": "VAQUER GERARDO", "VICCHI": "VICCHI GERMAN", "VICE": "VICEGOBERNACION", "ZLOBEC": "ZLOBEC LEIVA MARIANA G."
  };
  const aux9 = aux9Map[autoridad] || "";

  // AUX10: "un", "dos"
  const aux10Map = {
    "1": "un", "01": "un", "2": "dos", "02": "dos", "3": "tres", "03": "tres", "4": "cuatro", "04": "cuatro", "5": "cinco", "05": "cinco", "6": "seis", "06": "seis", "7": "siete", "07": "siete", "8": "ocho", "08": "ocho", "9": "nueve", "09": "nueve", "10": "diez", "11": "once", "12": "doce"
  };
  const aux10 = (aux10Map[cuota] || "").trim();

  // AUX11: "mes" / "meses"
  const aux11 = (aux10 === "") ? "" : (aux10 === "un" ? "mes" : "meses");

  // AUX12: "una", "dos"
  const aux12Map = {
    "1": "una", "01": "una", "2": "dos", "02": "dos", "3": "tres", "03": "tres", "4": "cuatro", "04": "cuatro", "5": "cinco", "05": "cinco", "6": "seis", "06": "seis", "7": "siete", "07": "siete", "8": "ocho", "08": "ocho", "9": "nueve", "09": "nueve", "10": "diez", "11": "once", "12": "doce"
  };
  const aux12 = aux12Map[cuota] || "";

  // AUX13
  const aux13 = (aux10 === "") ? "" : (aux10 === "un" ? "cuota" : "cuotas");

  // AUX14
  const aux14 = (aux13 === "") ? "" : (aux13 === "cuotas" ? "cada una efectivizadas a mes vencido abonándose la primera en el mes de" : "efectivizada a mes vencido, abonándose en");

  // AUX15
  const aux15 = (aux13 === "") ? "" : (aux13 === "cuotas" ? "cuotas mensuales, iguales y consecutivas" : "cuota mensual");

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

  // divResult
  const cuotaNum = parseInt(cuota, 10);
  let divResult = "";
  if (!isNaN(numericTotal) && !isNaN(cuotaNum) && cuotaNum > 0) {
    divResult = (numericTotal / cuotaNum).toFixed(2);
  }

  // AUX18
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

  // AUX21 logic based on formula: =ESPACIOS(SI(B6="";"";SI(B6="VICE";"L96003 H. LEGISLATURA";SI(B6="LEGISLATIVA";"L96003 H. LEGISLATURA";SI(B6="ADM";"L96003 H. LEGISLATURA";"L96000 H. SENADO")))))
  const authUpper = autoridad.toUpperCase();
  const aux21 = (autoridad === "") ? "" :
    (authUpper === "VICE" || authUpper === "LEGISLATIVA" || authUpper === "ADM") ?
      "L96003 H. LEGISLATURA" :
      "L96000 H. SENADO";

  // FECHA ACTUAL (especialmente para etiqueta <<FECHA>>)
  const fechaActual = fechaEnEspaniol(new Date());

  // Custom format total
  let totalFormateado = "";
  if (!isNaN(numericTotal) && numericTotal > 0) {
    totalFormateado = "$" + numericTotal.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // ID Plantilla Resolución
  const idPlantillaResolucion = '1PTRDI0JcACTe34nP4cBLG4w5rwM2CrBfxxRT1nXE33o';
  // Carpeta destino Resoluciones
  const idCarpetaResolucion = '1Es4FSVTBkFaXzZb33Y7KhqRN6nqT3d-y';

  // Create file
  const docBase = DriveApp.getFileById(idPlantillaResolucion);
  const autor = String(rowData['AUTOR'] || '').trim();
  const nombreArchivo = (req ? req : 'S_REQ') + "_" + apellidos + "_" + autoridad + (autor ? "_" + autor : '');
  const carpetaPDF = DriveApp.getFolderById(idCarpetaResolucion);

  const nuevoDocFile = docBase.makeCopy(nombreArchivo, carpetaPDF);
  const nuevoDocId = nuevoDocFile.getId();
  const documento = DocumentApp.openById(nuevoDocId);
  const body = documento.getBody();

  // Reemplazar textos
  body.replaceText("<<SEXO>>", sexo);
  body.replaceText("<<NOMBRES>>", nombres);
  body.replaceText("<<APELLIDOS>>", apellidos);
  body.replaceText("<<DNI>>", dni);
  body.replaceText("<<TAREAS>>", tareas);
  body.replaceText("<<FECHA_ALTA>>", fechaAlta);
  body.replaceText("<<FECHA>>", fechaActual);

  body.replaceText("<<CUOTA>>[ ]+", cuota + " ");
  body.replaceText("<<CUOTA>>", cuota);

  body.replaceText("<<TOTAL>>", totalFormateado);
  body.replaceText("<<REQ>>", req);

  // Tags of AUX
  body.replaceText("<<AUX1>>", aux1);
  body.replaceText("<<AUX2>>", aux2);
  body.replaceText("<<AUX5>>", aux5);
  body.replaceText("<<AUX6>>", aux6);
  body.replaceText("<<AUX7>>", aux7);
  body.replaceText("<<AUX9>>", aux9);

  body.replaceText("<<AUX10>>[ ]+", aux10 + " ");
  body.replaceText("<<AUX10>>", aux10);

  body.replaceText("<<AUX11>>", aux11);
  body.replaceText("<<AUX12>>", aux12);
  body.replaceText("<<AUX14>>", aux14);
  body.replaceText("<<AUX15>>", aux15);
  body.replaceText("<<AUX16>>", aux16);
  body.replaceText("<<AUX17>>", aux17);
  body.replaceText("<<AUX18>>", aux18Str);

  body.replaceText("<<AUX19>>[ ]+", aux19 + " ");
  body.replaceText("<<AUX19>>", aux19);

  body.replaceText("<<AUX21>>", aux21);
  const textoPlano = body.getText();
  documento.saveAndClose();

  // Create TXT file
  const folderPlano = DriveApp.getFolderById(CARPETA_FUSION_RESOLUCIONES);
  const archivoFinal = folderPlano.createFile(nombreArchivo + ".txt", textoPlano, MimeType.PLAIN_TEXT);
  const archivoFinalUrl = archivoFinal.getUrl();

  Logger.log('Resolución .txt creada exitosamente.');

  // Eliminar temp docs original (no lo guardamos en docs, solo el txt)
  DriveApp.getFileById(nuevoDocId).setTrashed(true);

  return { url: archivoFinalUrl, nombre: nombreArchivo + ".txt" };
}

function bajaResolucionTXT(rowData) {
  Logger.log("Iniciando bajaResolucionTXT() para: " + rowData['APELLIDOS']);

  // Helper: convierte a "D de mes de YYYY" en español
  function fechaEnEspaniol(valor) {
    if (!valor) return '';
    const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    let d;
    if (valor instanceof Date) {
      d = valor;
    } else {
      const s = String(valor).trim();
      if (/\d{1,2}\s+de\s+\w+\s+de\s+\d{4}/i.test(s)) return s;
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
        const p = s.split('-');
        d = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
      } else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) {
        const p = s.split('/');
        d = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
      } else {
        d = new Date(s);
      }
    }
    if (!d || isNaN(d.getTime())) return String(valor);
    let dStr = String(d.getDate());
    if (dStr.length === 1) dStr = '0' + dStr;
    return dStr + ' de ' + MESES[d.getMonth()] + ' de ' + d.getFullYear();
  }

  // Extract variables requeridas: REQ, APELLIDOS, NOMBRES, AUX5, AUX9, FECHA DE BAJA, AUX2, AUX1, DNI
  const sexo = String(rowData['SEXO'] || '').trim();
  const nombres = String(rowData['NOMBRES'] || '').trim();
  const apellidos = String(rowData['APELLIDOS'] || '').trim();
  const dni = formatearDNI(String(rowData['DNI'] || '').trim());
  const fechaBaja = fechaEnEspaniol(rowData['FECHA BAJA']);
  const req = String(rowData['REQ'] || '').trim();
  const autoridad = String(rowData['AUTORIDAD'] || '').trim();

  // --- Mapeo de AUXiliares ---
  const aux1 = (sexo === "el Sr.") ? "con el Sr." : (sexo === "la Sra.") ? "con la Sra." : "";

  const aux2 = (sexo === "el Sr.") ? "del Sr." : (sexo === "la Sra.") ? "de la Sra." : "";

  const aux5Map = {
    "ADM": "por", "LEGISLATIVA": "por", "VICE": "por",
    "ANA": "por el Senador", "CHAPPEL": "por el Senador", "DIUMENJO": "por el Senador", "FREIDEMBERG": "por el Senador", "GONZALEZ F.": "por el Senador", "GONZALEZ V.": "por el Senador", "IGLESIAS": "por el Senador", "KERCHNER": "por el Senador", "MAGISTRETTI": "por el Senador", "MARCOLINI": "por el Senador", "MARQUEZ": "por el Senador", "PERVIU": "por el Senador", "PEZZUTTI": "por el Senador", "PRADINES": "por el Senador", "PRINGLES": "por el Senador", "QUATTRINI": "por el Senador", "ROSTAND": "por el Senador", "SAEZ": "por el Senador", "SAT": "por el Senador", "SERRA": "por el Senador", "SEVILLA": "por el Senador", "SOTO": "por el Senador", "VAQUER": "por el Senador",
    "ASES": "por la Senadora", "BARRO": "por la Senadora", "CANO": "por la Senadora", "DERRACHE": "por la Senadora", "EISENCHLAS": "por la Senadora", "FLORIDIA": "por la Senadora", "GALIÑARES": "por la Senadora", "GOMEZ": "por la Senadora", "LAFERTE": "por la Senadora", "MANONI": "por la Senadora", "NAJUL": "por la Senadora", "SABADIN": "por la Senadora", "SAINZ": "por la Senadora", "VICCHI": "por la Senadora", "ZLOBEC": "por la Senadora"
  };
  const aux5 = aux5Map[autoridad] || "";

  const aux9Map = {
    "ADM": "SECRETARIA ADMINISTRATIVA", "ANA": "ANA MARIO ESTEBAN", "ASES": "ASES YAMEL", "BARRO": "BARRO JOHANA", "CANO": "CANO ADRIANA", "CHAPPEL": "CHAPPEL DUGAR", "DERRACHE": "DERRACHE MARIA", "DIUMENJO": "DIUMENJO ALEJANDRO", "EISENCHLAS": "EISENCHLAS NATALIA", "FLORIDIA": "FLORIDIA ANGELA", "FREIDEMBERG": "FREIDEMBRERG ABEL", "GALIÑARES": "GALIÑARES MARÍA", "GOMEZ": "GÓMEZ OLGA CRISTINA", "GONZALEZ F.": "GONZÁLEZ FELIX", "GONZALEZ V.": "GONZÁLEZ VALENTIN", "IGLESIAS": "IGLESIAS MARCELINO", "KERCHNER": "KERCHNER MARTIN", "LAFERTE": "LAFERTE JESICA", "LEGISLATIVA": "SECRETARIA LEGISLATIVA", "MAGISTRETTI": "MAGISTRETTI ARMANDO", "MANONI": "MANONI FLAVIA", "MARCOLINI": "MARCOLINI WALTHER", "MARQUEZ": "MARQUEZ SERGIO", "NAJUL": "NAJUL CLAUDIA", "PERVIU": "PERVIU HELIO", "PEZZUTTI": "PEZZUTTI DULIO", "PRADINES": "PRADINES GABRIEL", "PRINGLES": "PRINGLES ARIEL", "QUATTRINI": "QUATTRINI MARCOS", "ROSTAND": "ROSTAND MARTÍN", "SABADIN": "SABADIN MARIA FERNANDA", "SAEZ": "SÁEZ DAVID", "SAINZ": "SAINZ MARÍA LAURA", "SAT": "SAT MAURICIO", "SERRA": "SERRA PEDRO", "SEVILLA": "SEVILLA OSCAR", "SOTO": "SOTO OSCAR", "VAQUER": "VAQUER GERARDO", "VICCHI": "VICCHI GERMAN", "VICE": "VICEGOBERNACION", "ZLOBEC": "ZLOBEC LEIVA MARIANA G."
  };
  const aux9 = aux9Map[autoridad] || "";

  // ID Plantilla Baja Resolución
  const docTemplateBajaResolucionId = DOC_TEMPLATE_BAJA_RESOLUCION_ID;

  // Creation
  const docBase = DriveApp.getFileById(docTemplateBajaResolucionId);
  const autor = String(rowData['AUTOR'] || '').trim();
  const nombreArchivo = "BAJA_" + (req ? req : 'S_REQ') + "_" + apellidos + "_" + autoridad + (autor ? "_" + autor : '');
  const carpetaPDF = DriveApp.getFolderById(CARPETA_FUSION_RESOLUCIONES);

  const nuevoDocFile = docBase.makeCopy(nombreArchivo, carpetaPDF);
  const nuevoDocId = nuevoDocFile.getId();
  const documento = DocumentApp.openById(nuevoDocId);
  const body = documento.getBody();

  // Reemplazar textos indicados en el prompt
  body.replaceText("<<REQ>>", req);
  body.replaceText("<<APELLIDOS>>", apellidos);
  body.replaceText("<<NOMBRES>>", nombres);
  body.replaceText("<<AUX5>>", aux5);
  body.replaceText("<<AUX9>>", aux9);
  body.replaceText("<<FECHA_BAJA>>", fechaBaja);
  body.replaceText("<<AUX2>>", aux2);
  body.replaceText("<<AUX1>>", aux1);
  body.replaceText("<<DNI>>", dni);

  // Guardamos texto
  const textoPlano = body.getText();
  documento.saveAndClose();

  // Guardar como TXT
  const folderPlano = DriveApp.getFolderById(CARPETA_FUSION_RESOLUCIONES);
  const archivoFinal = folderPlano.createFile(nombreArchivo + ".txt", textoPlano, MimeType.PLAIN_TEXT);

  Logger.log('Baja Resolución .txt creada exitosamente.');

  // Eliminar temp docs original
  DriveApp.getFileById(nuevoDocId).setTrashed(true);

  return { success: true, url: archivoFinal.getUrl(), nombre: nombreArchivo + ".txt" };
}