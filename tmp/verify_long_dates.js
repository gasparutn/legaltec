
function parseFechaGeneral(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const s = String(val).trim();
  
  let match = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  
  match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) return new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]));
  
  match = s.match(/^(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})/i);
  if (match) {
    const dia = parseInt(match[1]);
    const mesNombre = match[2].toLowerCase();
    const anyo = parseInt(match[3]);
    const mesesNombres = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
    const mesIndex = mesesNombres.indexOf(mesNombre);
    if (mesIndex !== -1) return new Date(anyo, mesIndex, dia);
  }
  
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function calculateNextMonth(fechaH) {
  const d = parseFechaGeneral(fechaH);
  if (d && !isNaN(d.getTime())) {
    const nextMonth = new Date(d.getFullYear(), d.getMonth() + 1, 1);
    const mesesNombres = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
    return `${mesesNombres[nextMonth.getMonth()]} de ${nextMonth.getFullYear()}`;
  }
  return 'INVALID';
}

const cases = [
  { input: '1 de febrero de 2026', expected: 'marzo de 2026' },
  { input: '1 de marzo de 2026', expected: 'abril de 2026' },
  { input: '31 de enero de 2026', expected: 'febrero de 2026' }
];

console.log("--- Test 3: Long Text Date Parsing for AUX16 ---");
cases.forEach(c => {
  const result = calculateNextMonth(c.input);
  const status = result === c.expected ? "✅" : "❌";
  console.log(`Input: ${c.input} -> Result: ${result} (Expected: ${c.expected}) ${status}`);
});
