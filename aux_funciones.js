// --- Funciones auxiliares de cálculo ---
// (Mantener las funciones calculateAuxColumns, Unidades, Decenas, etc. aquí)
// ... (El resto de tu código para AUX_FUNCIONES.GS, NOTAS.GS, RESOLUCION.GS, NOTIFICA.GS, NUMALETRAS.GS, CONTRATOS.GS)
// Asegúrate de que todas las funciones de tu archivo 'para la iA' estén presentes aquí.

/**
 * Calcula los valores para las columnas auxiliares (elSR_conElSr a FECHAHOY)
 * basándose en los datos de una fila.
 * Se asume que rowData contiene las claves con la capitalización exacta de los encabezados de la hoja.
 * @param {Object} rowData Objeto que contiene los datos de la fila, mapeados por encabezado.
 * @returns {Object} Un objeto con las propiedades AUX calculadas.
 */
function calculateAuxColumns(rowData) {
  const aux = {};

  // Columna Q, cabecera “elSR_conElSr”.
  const sexoC = rowData['SEXO'] || '';
  if (sexoC === 'el Sr.') {
    aux['elSR_conElSr'] = 'con el Sr.';
  } else if (sexoC === 'la Sra.') {
    aux['elSR_conElSr'] = 'con la Sra.';
  } else {
    aux['elSR_conElSr'] = '';
  }
  aux['elSR_conElSr'] = aux['elSR_conElSr'].trim();

  // Columna R, cabecera “elSR_delSR”. FORMULA =ESPACIOS(IFS(C2="el Sr.","del Sr.",C2="la Sra.","de la Sra.",C2="",""))
  if (sexoC === 'el Sr.') {
    aux['delSr_delaSra'] = 'del Sr.';
  } else if (sexoC === 'la Sra.') {
    aux['delSr_delaSra'] = 'de la Sra.';
  } else {
    aux['delSr_delaSra'] = '';
  }
  aux['delSr_delaSra'] = aux['delSr_delaSra'].trim();

  // Columna S, cabecera “el_denominado_a”. FORMULA =ESPACIOS(IFS(C2="el Sr.","denominado",C2="la Sra.","denominada",C2="",""))
  if (sexoC === 'el Sr.') {
    aux['el_denominado_a'] = 'denominado';
  } else if (sexoC === 'la Sra.') {
    aux['el_denominado_a'] = 'denominada';
  } else {
    aux['el_denominado_a'] = '';
  }
  aux['el_denominado_a'] = aux['el_denominado_a'].trim();

  // Columna T, cabecera “LOCADOR_AR”. FORMULA =ESPACIOS(IFS(C2="el Sr.","EL LOCADOR",C2="la Sra.","LA LOCADORA",C2="",""))
  if (sexoC === 'el Sr.') {
    aux['LOCADOR_AR'] = 'EL LOCADOR';
  } else if (sexoC === 'la Sra.') {
    aux['LOCADOR_AR'] = 'LA LOCADORA';
  } else {
    aux['LOCADOR_AR'] = '';
  }
  aux['LOCADOR_AR'] = aux['LOCADOR_AR'].trim();

  // Columna U, cabecera “VICE_LEGIS_Xel_SENADOR”. FORMULA =ESPACIOS(IFS(B2="VICE","por",B2="LEGISLATIVA","por",B2="ADM","por",C2="el Sr.","por el Senador",C2="la Sra.","por la Senadora",C2="",""))
  const autoridadB = rowData['AUTORIDAD'] || ''; // Asumiendo 'AUTORIDAD' es el encabezado de la columna B
  if (['VICE', 'LEGISLATIVA', 'ADM'].includes(String(autoridadB).toUpperCase())) {
    aux['VICE_LEGIS_Xel_SENADOR'] = 'por';
  } else if (sexoC === 'el Sr.') {
    aux['VICE_LEGIS_Xel_SENADOR'] = 'por el Senador';
  } else if (sexoC === 'la Sra.') {
    aux['VICE_LEGIS_Xel_SENADOR'] = 'por la Senadora';
  } else {
    aux['VICE_LEGIS_Xel_SENADOR'] = '';
  }
  aux['VICE_LEGIS_Xel_SENADOR'] = aux['VICE_LEGIS_Xel_SENADOR'].trim();

  // Columna V, cabecera “porElContratado”. FORMULA =ESPACIOS(IFS(C2="el Sr.","por el contratado",C2="la Sra.","por la contratada",C2="",""))
  if (sexoC === 'el Sr.') {
    aux['porElContratado'] = 'por el contratado';
  } else if (sexoC === 'la Sra.') {
    aux['porElContratado'] = 'por la contratada';
  } else {
    aux['porElContratado'] = '';
  }
  aux['porElContratado'] = aux['porElContratado'].trim();

  // Columna W, cabecera “elSenador_a”. FORMULA =ESPACIOS(IFS(C2="el Sr.","el Senador",C2="la Sra.","la Senadora",C2="",""))
  if (sexoC === 'el Sr.') {
    aux['elSenador_a'] = 'el Senador';
  } else if (sexoC === 'la Sra.') {
    aux['elSenador_a'] = 'la Senadora';
  } else {
    aux['elSenador_a'] = '';
  }
  aux['elSenador_a'] = aux['elSenador_a'].trim();

  // Columna X, cabecera “VICE_LEGIS_DelSenador_a”. FORMULA =ESPACIOS(IFS(B2="VICE","de",B2="LEGISLATIVA","de",B2="ADM","de",C2="el Sr.","del Senador",C2="la Sra.","de la Senadora",C2="",""))
  if (['VICE', 'LEGISLATIVA', 'ADM'].includes(String(autoridadB).toUpperCase())) {
    aux['VICE_LEGIS_DelSenador_a'] = 'de';
  } else if (sexoC === 'el Sr.') {
    aux['VICE_LEGIS_DelSenador_a'] = 'del Senador';
  } else if (sexoC === 'la Sra.') {
    aux['VICE_LEGIS_DelSenador_a'] = 'de la Senadora';
  } else {
    aux['VICE_LEGIS_DelSenador_a'] = '';
  }
  aux['VICE_LEGIS_DelSenador_a'] = aux['VICE_LEGIS_DelSenador_a'].trim();

  // Columna Y, cabecera “AUTORIDAD”. FORMULA =SWITCH(B2,...)
  switch (String(autoridadB).toUpperCase()) {
    case 'ADM': aux['AUTORIDAD_COMPLETA'] = 'Secretaría Administrativa'; break;
    case 'ANA': aux['AUTORIDAD_COMPLETA'] = 'ANA MARIO ESTEBAN'; break;
    case 'ASES': aux['AUTORIDAD_COMPLETA'] = 'ASES YAMEL'; break;
    case 'BARRO': aux['AUTORIDAD_COMPLETA'] = 'BARRO JOHANA A.'; break;
    case 'CANO': aux['AUTORIDAD_COMPLETA'] = 'CANO ADRIANA E.'; break;
    case 'CHAPPEL': aux['AUTORIDAD_COMPLETA'] = 'CHAPPEL DUGAR'; break;
    case 'DERRACHE': aux['AUTORIDAD_COMPLETA'] = 'DERRACHE MIRIAM.'; break;
    case 'DIUMENJO': aux['AUTORIDAD_COMPLETA'] = 'DIUMENJO ALEJANDRO D.'; break;
    case 'EISENCHLAS': aux['AUTORIDAD_COMPLETA'] = 'EISENCHLAS NATALIA F.'; break;
    case 'FLORIDIA': aux['AUTORIDAD_COMPLETA'] = 'FLORIDIA ANGELA'; break;
    case 'FREIDEMBERG': aux['AUTORIDAD_COMPLETA'] = 'FREIDEMBRERG ABEL L.'; break;
    case 'GALIÑARES': aux['AUTORIDAD_COMPLETA'] = 'GALIÑARES MARÍA'; break;
    case 'GÓMEZ': aux['AUTORIDAD_COMPLETA'] = 'GÓMEZ OLGA CRISTINA'; break;
    case 'GONZÁLEZ F.': aux['AUTORIDAD_COMPLETA'] = 'GONZÁLEZ FELIX'; break;
    case 'GONZÁLEZ V.': aux['AUTORIDAD_COMPLETA'] = 'GONZÁLEZ VELANTIN'; break;
    case 'IGLESIAS': aux['AUTORIDAD_COMPLETA'] = 'IGLESIAS MARCELINO'; break;
    case 'KERCHNER': aux['AUTORIDAD_COMPLETA'] = 'KERCHNER MARTIN'; break;
    case 'LAFERTE': aux['AUTORIDAD_COMPLETA'] = 'LAFERTE JESICA C.'; break;
    case 'LEGISLATIVA': aux['AUTORIDAD_COMPLETA'] = 'Secretaría Legislativa'; break;
    case 'MAGISTRETTI': aux['AUTORIDAD_COMPLETA'] = 'MAGISTRETTI ARMANDO'; break;
    case 'MANONI': aux['AUTORIDAD_COMPLETA'] = 'MANONI FLAVIA'; break;
    case 'MARCOLINI': aux['AUTORIDAD_COMPLETA'] = 'MARCOLINI WALTHER'; break;
    case 'MARQUEZ': aux['AUTORIDAD_COMPLETA'] = 'MARQUEZ SERGIO'; break;
    case 'NAJUL': aux['AUTORIDAD_COMPLETA'] = 'NAJUL CLAUDIA I.'; break;
    case 'PERVIU': aux['AUTORIDAD_COMPLETA'] = 'PERVIU HELIO M.'; break;
    case 'PEZZUTTI': aux['AUTORIDAD_COMPLETA'] = 'PEZZUTTI DULIO'; break;
    case 'PRADINES': aux['AUTORIDAD_COMPLETA'] = 'PRADINES GABRIEL A.'; break;
    case 'PRINGLES': aux['AUTORIDAD_COMPLETA'] = 'PRINGLES ARIEL'; break;
    case 'QUATTRINI': aux['AUTORIDAD_COMPLETA'] = 'QUATTRINI MARCOS'; break;
    case 'ROSTAND': aux['AUTORIDAD_COMPLETA'] = 'ROSTAND MARTÍN G.'; break;
    case 'RUS': aux['AUTORIDAD_COMPLETA'] = 'RUS MARÍA M.'; break;
    case 'SABADIN': aux['AUTORIDAD_COMPLETA'] = 'SABADIN MARIA FERNANDA'; break;
    case 'SÁEZ': aux['AUTORIDAD_COMPLETA'] = 'SÁEZ DAVID'; break;
    case 'SAINZ': aux['AUTORIDAD_COMPLETA'] = 'SAINZ MARÍA LAURA'; break;
    case 'SAT': aux['AUTORIDAD_COMPLETA'] = 'SAT MAURICIO'; break;
    case 'SERRA': aux['AUTORIDAD_COMPLETA'] = 'SERRA PEDRO'; break;
    case 'SEVILLA': aux['AUTORIDAD_COMPLETA'] = 'SEVILLA OSCAR'; break;
    case 'SOTO': aux['AUTORIDAD_COMPLETA'] = 'SOTO OSCAR'; break;
    case 'VAQUER': aux['AUTORIDAD_COMPLETA'] = 'VAQUER GERARDO R.'; break;
    case 'VICCHI': aux['AUTORIDAD_COMPLETA'] = 'VICCHI GERMAN A.'; break;
    case 'VICE': aux['AUTORIDAD_COMPLETA'] = 'VICEGOBERNACION'; break;
    case 'ZLOBEC': aux['AUTORIDAD_COMPLETA'] = 'ZLOBEC LEIVA MARIANA G.'; break;
    default: aux['AUTORIDAD_COMPLETA'] = ''; break;
  }

  // Columna Z, cabecera “NUM_LETRA”. FORMULA =SWITCH(I2,...)
  const cuotasI = String(rowData['CUOTA'] || '').trim(); // Asumiendo 'CUOTA' es el encabezado de la columna I
  switch (cuotasI) {
    case '1': aux['NUM_LETRA'] = 'un'; break;
    case '2': aux['NUM_LETRA'] = 'dos'; break;
    case '3': aux['NUM_LETRA'] = 'tres'; break;
    case '4': aux['NUM_LETRA'] = 'cuatro'; break;
    case '5': aux['NUM_LETRA'] = 'cinco'; break;
    case '6': aux['NUM_LETRA'] = 'seis'; break;
    case '7': aux['NUM_LETRA'] = 'siete'; break;
    case '8': aux['NUM_LETRA'] = 'ocho'; break;
    case '9': aux['NUM_LETRA'] = 'nueve'; break;
    case '10': aux['NUM_LETRA'] = 'diez'; break;
    case '11': aux['NUM_LETRA'] = 'once'; break;
    case '12': aux['NUM_LETRA'] = 'doce'; break;
    default: aux['NUM_LETRA'] = ''; break;
  }

  // Columna AA, cabecera “MES_ES”. FORMULA =SI(Z2="","",SI(Z2="un","mes","meses"))
  if (aux['NUM_LETRA'] === '') {
    aux['MES_ES'] = '';
  } else if (aux['NUM_LETRA'] === 'un') {
    aux['MES_ES'] = 'mes';
  } else {
    aux['MES_ES'] = 'meses';
  }
  const mesesCUOTA = String(rowData['CUOTA'] || ''); // Asumiendo 'CUOTA' es el encabezado de la columna I
  // Columna AB, cabecera “NUMLETRA”. FORMULA =SWITCH(I2,...)
  switch (mesesCUOTA) {
    case '1': aux['NUMLETRA'] = 'UN'; break;
    case '2': aux['NUMLETRA'] = 'DOS'; break;
    case '3': aux['NUMLETRA'] = 'TRES'; break;
    case '4': aux['NUMLETRA'] = 'CUATRO'; break;
    case '5': aux['NUMLETRA'] = 'CINCO'; break;
    case '6': aux['NUMLETRA'] = 'SEIS'; break;
    case '7': aux['NUMLETRA'] = 'SIETE'; break;
    case '8': aux['NUMLETRA'] = 'OCHO'; break;
    case '9': aux['NUMLETRA'] = 'NUEVE'; break;
    case '10': aux['NUMLETRA'] = 'DIEZ'; break;
    case '11': aux['NUMLETRA'] = 'ONCE'; break;
    case '12': aux['NUMLETRA'] = 'DOCE'; break;
    default: aux['NUMLETRA'] = ''; break;
  }

  // Columna AC, cabecera “elSR_conElSr3”. FORMULA =SI(Z2="","",SI(Z2="un","cuota","cuotas"))
  if (aux['NUM_LETRA'] === '') {
    aux['CUOTA_LE'] = '';
  } else if (aux['NUM_LETRA'] === 'un') {
    aux['CUOTA_LE'] = 'cuota';
  } else {
    aux['CUOTA_LE'] = 'cuotas';
  }

  // Columna AD, cabecera “elSR_conElSr4”. FORMULA =ESPACIOS(SI(AC2="","",SI(AC2="cuotas","cada una, efectivizadas a mes vencido abonándose la primera en el mes de","efectivizada a mes vencido, abonándose en")))
  if (aux['CUOTA_LE'] === '') {
    aux['elSR_conElSr4'] = '';
  } else if (aux['CUOTA_LE'] === 'cuotas') {
    aux['elSR_conElSr4'] = 'cada una, efectivizadas a mes vencido abonándose la primera en el mes de';
  } else {
    aux['elSR_conElSr4'] = 'efectivizada a mes vencido, abonándose en';
  }
  aux['elSR_conElSr4'] = aux['elSR_conElSr4'].trim();

  // Columna AE, cabecera “elSR_conElSr5”. FORMULA =ESPACIOS(SI(AC2="","",SI(AC2="cuotas","cuotas mensuales, iguales y consecutivas","cuota mensual")))
  if (aux['CUOTA_LE'] === '') {
    aux['elSR_conElSr5'] = '';
  } else if (aux['CUOTA_LE'] === 'cuotas') {
    aux['elSR_conElSr5'] = 'cuotas mensuales, iguales y consecutivas';
  } else {
    aux['elSR_conElSr5'] = 'cuota mensual';
  }
  aux['elSR_conElSr5'] = aux['elSR_conElSr5'].trim();

  // Columna AF, cabecera “FECHACORTA”. FORMULA =SWITCH(H2,...)
  const fechaH = rowData['FECHA ALTA'] || ''; // Asumiendo 'FECHA ALTA' es el encabezado de la columna H
  switch (String(fechaH).trim()) {
    case '1 de enero de 2024': aux['FECHACORTA'] = 'febrero de 2024'; break;
    case '1 de febrero de 2024': aux['FECHACORTA'] = 'marzo de 2024'; break;
    case '1 de marzo de 2024': aux['FECHACORTA'] = 'abril de 2024'; break;
    case '1 de abril de 2024': aux['FECHACORTA'] = 'mayo de 2024'; break;
    case '1 de mayo de 2024': aux['FECHACORTA'] = 'junio de 2024'; break;
    case '1 de junio de 2024': aux['FECHACORTA'] = 'julio de 2024'; break;
    case '1 de julio de 2024': aux['FECHACORTA'] = 'agosto de 2024'; break;
    case '1 de agosto de 2024': aux['FECHACORTA'] = 'septiembre de 2024'; break;
    case '1 de septiembre de 2024': aux['FECHACORTA'] = 'octubre de 2024'; break;
    case '1 de octubre de 2024': aux['FECHACORTA'] = 'noviembre de 2024'; break;
    case '1 de noviembre de 2024': aux['FECHACORTA'] = 'diciembre de 2024'; break;
    default: aux['FECHACORTA'] = ''; break;
  }

  // Columna AG, cabecera “NUM_A_LETRAS”. FORMULA =ESPACIOS(SI(J2="","",NumeroALetras(J2)))
  const importeTotalJ = rowData['TOTAL']; // Asumiendo 'TOTAL' es el encabezado de la columna J
  if (importeTotalJ === '' || importeTotalJ === null || importeTotalJ === undefined) {
    aux['NUM_A_LETRAS'] = '';
  } else {
    // Asegurarse de que el valor sea un número antes de pasarlo a NumeroALetras
    const numericTotal = typeof importeTotalJ === 'string' ? parseFloat(importeTotalJ.replace(/[^0-9,-]+/g, "").replace(",", ".")) : importeTotalJ;
    aux['NUM_A_LETRAS'] = NumeroALetras(numericTotal);
  }
  aux['NUM_A_LETRAS'] = aux['NUM_A_LETRAS'].trim();

  // Columna AH, cabecera “TOTAL_DIV_CUO”. FORMULA =SI.ERROR(J2/I2,"")
  const cuotasNum = parseFloat(String(rowData['CUOTA'] || '').replace(',', '.')); // Convertir a número, manejar coma
  const totalNum = typeof importeTotalJ === 'string' ? parseFloat(importeTotalJ.replace(/[^0-9,-]+/g, "").replace(",", ".")) : parseFloat(importeTotalJ); // Convertir a número, manejar coma
  if (!isNaN(totalNum) && !isNaN(cuotasNum) && cuotasNum > 0) {
    aux['TOTAL_DIV_CUO'] = totalNum / cuotasNum;
  } else {
    aux['TOTAL_DIV_CUO'] = '';
  }

  // Columna AI, cabecera “NUM_A_LET_CUOTA”. FORMULA =ESPACIOS(SI(AH2="","",NumeroALetras(AH2)))
  if (aux['TOTAL_DIV_CUO'] === '' || aux['TOTAL_DIV_CUO'] === null || aux['TOTAL_DIV_CUO'] === undefined) {
    aux['NUM_A_LET_CUOTA'] = '';
  } else {
    // Asegurarse de que el valor sea un número antes de pasarlo a NumeroALetras
    const numericCuota = parseFloat(aux['TOTAL_DIV_CUO']);
    aux['NUM_A_LET_CUOTA'] = NumeroALetras(numericCuota);
  }
  aux['NUM_A_LET_CUOTA'] = aux['NUM_A_LET_CUOTA'].trim();

  // Columna AJ, cabecera “FECHAHOY”. FORMULA =SI(AF2="","",HOY())
  if (aux['FECHACORTA'] === '') {
    aux['FECHAHOY'] = '';
  } else {
    aux['FECHAHOY'] = new Date(); // Retorna un objeto Date, puedes formatearlo si es necesario en la plantilla
  }

  // AUX21 - Basado en fórmula: =ESPACIOS(SI(B6="";"";SI(B6="VICE";"L96003 H. LEGISLATURA";SI(B6="LEGISLATIVA";"L96003 H. LEGISLATURA";SI(B6="ADM";"L96003 H. LEGISLATURA";"L96000 H. SENADO")))))
  if (autoridadB === "") {
    aux['AUX21'] = "";
  } else if (['VICE', 'LEGISLATIVA', 'ADM'].includes(String(autoridadB).toUpperCase())) {
    aux['AUX21'] = "L96003 H. LEGISLATURA";
  } else {
    aux['AUX21'] = "L96000 H. SENADO";
  }

  return aux;
}

/**
 * Formatea un DNI al estilo xx.xxx.xxx (con puntos)
 * @param {string|number} dni El DNI a formatear.
 * @returns {string} El DNI formateado con puntos.
 */
function formatearDNI(dni) {
  if (!dni) return '';
  // Si ya tiene puntos o comas, limpiamos primero para estandarizar
  // Nota: El usuario pidió puntos en vez de comas.
  let soloNumeros = String(dni).replace(/\D/g, '');
  if (!soloNumeros) return String(dni);

  // Formatear con puntos como separadores de miles
  // Esto cubrirá tanto xx.xxx.xxx como cualquier otra longitud
  return soloNumeros.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

/**
 * Recorre un objeto y hace trim a todos sus valores de tipo string.
 * @param {Object} data El objeto a limpiar.
 * @returns {Object} El objeto con los valores limpios.
 */
function trimData(data) {
  const result = {};
  for (let key in data) {
    if (typeof data[key] === 'string') {
      result[key] = data[key].trim();
    } else {
      result[key] = data[key];
    }
  }
  return result;
}

