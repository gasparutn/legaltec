/**
 * aCalcula los valores para las columnas auxiliares (AUX1 a AUX20)
 */
function calculateAuxColumns(rowDataRaw: { [key: string]: any }): { [key: string]: any } {
  // Normalizar datos de entrada: trimming de valores
  const rowData: { [key: string]: any } = {};
  for (let key in rowDataRaw) {
    const val = rowDataRaw[key];
    rowData[key.trim().toUpperCase()] = (typeof val === 'string') ? val.trim() : val;
  }

  // Función interna para buscar claves de forma flexible (espacio vs guion bajo)
  const getVal = (key: string) => {
    const k = key.toUpperCase();
    return rowData[k] || rowData[k.replace(/\s+/g, '_')] || rowData[k.replace(/_/g, ' ')] || '';
  };

  const aux: { [key: string]: any } = {};
  const sexoC = getVal('SEXO');
  const autoridadB = getVal('AUTORIDAD');
  const cuotaTotal = String(getVal('CUOTA')).trim();
  const cuotaInt = parseInt(cuotaTotal, 10) || 0;

  // Metadatos de autoridad (género)
  const authInfo = getAutoridadMetadata(autoridadB);
  const isSenadora = authInfo.tipo === 'senadora';
  const isSenador = authInfo.tipo === 'senador';

  // aux1 y aux2: Tratamientos por género (el Sr./la Sra. y del Sr./de la Sra.)
  const sx = String(sexoC).toUpperCase().trim();
  if (sx === 'EL SR.' || sx === 'MASCULINO' || sx === 'M') {
    aux['AUX1'] = 'con el Sr.';
    aux['AUX2'] = 'del Sr.';
  } else if (sx === 'LA SRA.' || sx === 'FEMENINO' || sx === 'F') {
    aux['AUX1'] = 'con la Sra.';
    aux['AUX2'] = 'de la Sra.';
  } else {
    aux['AUX1'] = '';
    aux['AUX2'] = '';
  }

  // aux3: "denominado" o "denominada"
  if (sexoC === 'el Sr.') aux['AUX3'] = 'denominado';
  else if (sexoC === 'la Sra.') aux['AUX3'] = 'denominada';
  else aux['AUX3'] = '';

  // aux4: "LA LOCADORA" o "EL LOCADOR"
  if (sexoC === 'el Sr.') aux['AUX4'] = 'EL LOCADOR';
  else if (sexoC === 'la Sra.') aux['AUX4'] = 'LA LOCADORA';
  else aux['AUX4'] = '';

  // aux5: "por el Senador" , "por la Senadora", ""
  if (isSenadora) aux['AUX5'] = 'por la Senadora';
  else if (isSenador) aux['AUX5'] = 'por el Senador';
  else aux['AUX5'] = 'por'; // Secretaría / Vice

  // aux6: "por el contratado" o "por la contratada"
  if (sexoC === 'el Sr.') aux['AUX6'] = 'por el contratado';
  else if (sexoC === 'la Sra.') aux['AUX6'] = 'por la contratada';
  else aux['AUX6'] = '';

  // aux7: "el Senador" , "la Senadora", ""
  if (isSenadora) aux['AUX7'] = 'la Senadora';
  else if (isSenador) aux['AUX7'] = 'el Senador';
  else aux['AUX7'] = '';

  // aux8: "del Senador" ,"de la Senadora", "de"
  if (isSenadora) aux['AUX8'] = 'de la Senadora';
  else if (isSenador) aux['AUX8'] = 'del Senador';
  else aux['AUX8'] = 'de'; // Secretaría / Vice

  // aux10 y aux12 (Números a letras de la cuota)
  const letrasMin = ["", "un", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve", "diez", "once", "doce"];
  const letrasMay = ["", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE"];

  const padNum = (n: number) => (n > 0 && n < 10) ? '0' + n : String(n);

  aux['AUX10'] = (cuotaInt <= 12 && cuotaInt > 0) ? letrasMin[cuotaInt] : padNum(cuotaInt);
  aux['AUX12'] = (cuotaInt <= 12 && cuotaInt > 0) ? letrasMay[cuotaInt] : padNum(cuotaInt);

  // aux11: "meses" o "mes"
  aux['AUX11'] = (cuotaInt === 1) ? 'mes' : 'meses';

  // aux13: "cuotas" o "cuota"
  aux['AUX13'] = (cuotaInt === 1) ? 'cuota' : 'cuotas';

  // aux14
  if (cuotaInt > 1) {
    aux['AUX14'] = 'cada una efectivizadas a mes vencido abonándose la primera en el mes de';
  } else {
    aux['AUX14'] = 'efectivizada a mes vencido, abonándose en';
  }

  // aux15
  if (cuotaInt > 1) {
    aux['AUX15'] = 'cuotas mensuales, iguales y consecutivas';
  } else {
    aux['AUX15'] = 'cuota mensual';
  }

  // aux16 (Mes siguiente a la fecha de alta)
  const fechaH = getVal('FECHA ALTA');
  if (fechaH && String(fechaH).trim() !== '') {
    try {
      const d = parseSafeDate(fechaH);
      if (d && !isNaN(d.getTime()) && d.getFullYear() > 1900) {
        // Normalizar al día 1 para evitar problemas con meses de distinta duración (ej. 31 de enero)
        const nextMonth = new Date(d.getFullYear(), d.getMonth() + 1, 1);
        const mesesNombres = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
        aux['AUX16'] = `${mesesNombres[nextMonth.getMonth()]} de ${nextMonth.getFullYear()}`;
      } else {
        Logger.log("ERROR: FECHA ALTA inválida para AUX16: " + fechaH + ". Asegúrese que está en formato DD/MM/YYYY o D de mes de YYYY");
        aux['AUX16'] = '';
      }
    } catch (e) {
      Logger.log("ERROR calculando AUX16: " + e.message + " para fecha: " + fechaH);
      aux['AUX16'] = '';
    }
  } else {
    Logger.log("ERROR: FECHA ALTA no encontrada o vacía para AUX16");
    aux['AUX16'] = '';
  }

  const total = getVal('TOTAL');
  const numericTotal = typeof total === 'string' ? parseFloat(total.replace(/[^0-9,-]+/g, "").replace(",", ".")) : Number(total);
  if (!isNaN(numericTotal) && numericTotal > 0) {
    aux['AUX17'] = NumeroALetras(numericTotal);
  } else { aux['AUX17'] = ''; }

  if (!isNaN(numericTotal) && numericTotal > 0 && cuotaInt > 0) {
    const divCuo = numericTotal / cuotaInt;
    if (!isNaN(divCuo) && isFinite(divCuo) && divCuo > 0) {
      aux['AUX18'] = divCuo; // Valor numérico para total/cuota
      aux['AUX19'] = NumeroALetras(divCuo); // Letras del valor mensual
    } else {
      Logger.log("Aviso: División de TOTAL/CUOTA resultó en valor inválido: " + divCuo);
      aux['AUX18'] = '';
      aux['AUX19'] = '';
    }
  } else {
    aux['AUX18'] = '';
    aux['AUX19'] = '';
  }

  const today = new Date();
  const meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  aux['AUX20'] = `${today.getDate()} de ${meses[today.getMonth()]} de ${today.getFullYear()}`;

  // Otros auxiliares usados anteriormente para compatibilidad
  aux['AUTORIDAD_COMPLETA'] = authInfo.nombre;
  aux['AUX9'] = authInfo.nombre;

  if (autoridadB === "") {
    aux['AUX21'] = "";
  } else if (['VICE', 'LEGISLATIVA', 'ADM'].includes(String(autoridadB).toUpperCase())) {
    aux['AUX21'] = "L96003 H. LEGISLATURA";
  } else {
    aux['AUX21'] = "L96000 H. SENADO";
  }

  aux['AUX22'] = (sexoC === 'la Sra.') ? 'obligada' : 'obligado';
  aux['AUX24'] = (sexoC === 'la Sra.') ? 'ésta última' : 'éste último';

  // Sincronizar descriptivos para retrocompatibilidad
  aux['elSR_conElSr'] = aux['AUX1'];
  aux['delSr_delaSra'] = aux['AUX2'];
  aux['el_denominado_a'] = aux['AUX3'];
  aux['LOCADOR_AR'] = aux['AUX4'];
  aux['VICE_LEGIS_Xel_SENADOR'] = aux['AUX5'];
  aux['porElContratado'] = aux['AUX6'];
  aux['elSenador_a'] = aux['AUX7'];
  aux['VICE_LEGIS_DelSenador_a'] = aux['AUX8'];
  aux['NUM_LETRA'] = aux['AUX10'];
  aux['MES_ES'] = aux['AUX11'];
  aux['NUMLETRA'] = aux['AUX12'];
  aux['CUOTA_LE'] = aux['AUX13'];
  aux['elSR_conElSr4'] = aux['AUX14'];
  aux['elSR_conElSr5'] = aux['AUX15'];
  aux['FECHACORTA'] = aux['AUX16'];
  aux['NUM_A_LETRAS'] = aux['AUX17'];
  aux['TOTAL_DIV_CUO'] = aux['AUX18'];
  aux['NUM_A_LET_CUOTA'] = aux['AUX19'];
  aux['FECHAHOY'] = aux['AUX20'];

  return aux;
}

function formatearDNI(dni: string | number): string {
  if (!dni) return '';
  let soloNumeros = String(dni).replace(/\D/g, '');
  if (!soloNumeros) return String(dni);
  return soloNumeros.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

// parseFechaGeneral ha sido eliminada. Se utiliza parseSafeDate de utils.ts en su lugar.

function trimData(data: { [key: string]: any }): { [key: string]: any } {
  const result: { [key: string]: any } = {};
  for (let key in data) {
    if (typeof data[key] === 'string') {
      result[key] = sanitizeInput(data[key].trim());
    } else {
      result[key] = data[key];
    }
  }
  return result;
}
