
const letrasMin = ["", "un", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve", "diez", "once", "doce"];
const letrasMay = ["", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE"];

function padNum(n) {
  return (n > 0 && n < 10) ? '0' + n : String(n);
}

function calculateAuxColumnsMock(rowDataRaw) {
  const rowData = {};
  for (let key in rowDataRaw) {
    const val = rowDataRaw[key];
    rowData[key.trim().toUpperCase()] = (typeof val === 'string') ? val.trim() : val;
  }

  const aux = {};
  const cuotaTotal = String(rowData['CUOTA'] || '').trim();
  const cuotaInt = parseInt(cuotaTotal, 10) || 0;

  aux['AUX10'] = (cuotaInt <= 12 && cuotaInt > 0) ? letrasMin[cuotaInt] : padNum(cuotaInt);
  aux['AUX12'] = (cuotaInt <= 12 && cuotaInt > 0) ? letrasMay[cuotaInt] : padNum(cuotaInt);

  const fechaH = rowData['FECHA ALTA'] || '';
  if (fechaH) {
    let d;
    if (fechaH instanceof Date) { d = fechaH; }
    else {
      const match = String(fechaH).match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (match) { d = new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3])); }
      else { d = new Date(fechaH); }
    }
    if (!isNaN(d.getTime())) {
      const nextMonth = new Date(d);
      nextMonth.setMonth(nextMonth.getMonth() + 1);
      const mesesNombres = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
      aux['AUX16'] = `${mesesNombres[nextMonth.getMonth()]} de ${nextMonth.getFullYear()}`;
    } else { aux['AUX16'] = 'INVALID'; }
  } else { aux['AUX16'] = 'EMPTY'; }

  return { ...rowData, ...aux };
}

function replaceFullRowMock(data) {
  const fields = {};
  for (const [key, value] of Object.entries(data)) {
    const upperKey = key.trim().toUpperCase();
    let finalValue = value;
    if (typeof value === 'string') { finalValue = value.trim(); }
    if (upperKey === 'CUOTA' && finalValue !== undefined && finalValue !== null && finalValue !== '') {
      const num = parseInt(String(finalValue), 10);
      if (!isNaN(num) && num > 0 && num < 10) { finalValue = '0' + num; }
    }
    fields[`<<${upperKey}>>`] = finalValue;
  }
  return fields;
}

// Test cases
const testData = {
  'cuota ': ' 3 ',
  'fecha alta': '2026-03-01',
  'apellidos': '  GARCIA  '
};

console.log("--- Test 1: Trimming and Padding ---");
const processed = calculateAuxColumnsMock(testData);
const fields = replaceFullRowMock(processed);

console.log("CUOTA original:", testData['cuota ']);
console.log("CUOTA final:", fields['<<CUOTA>>']);
console.log("APELLIDOS final:", fields['<<APELLIDOS>>']);
console.log("AUX16 final:", fields['<<AUX16>>']);

if (fields['<<CUOTA>>'] === '03' && fields['<<APELLIDOS>>'] === 'GARCIA' && fields['<<AUX16>>'] === 'abril de 2026') {
  console.log("\n✅ PRUEBA EXITOSA");
} else {
  console.log("\n❌ PRUEBA FALLIDA");
  console.log(JSON.stringify(fields, null, 2));
}
