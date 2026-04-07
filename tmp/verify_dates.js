
function parseFechaGeneral(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const s = String(val).trim();
  
  // Formato YYYY-MM-DD (ISO)
  let match = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  
  // Formato DD/MM/YYYY (Común en planillas)
  match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) return new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]));
  
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

// Test cases
const cases = [
  { input: '2026-02-01', expected: 'marzo de 2026' },
  { input: '01/02/2026', expected: 'marzo de 2026' },
  { input: '31/01/2026', expected: 'febrero de 2026' }, // Fin de mes
  { input: new Date(2026, 2, 1), expected: 'abril de 2026' } // Objeto Date (Marzo -> Abril)
];

console.log("--- Test 2: Robust Date Parsing for AUX16 ---");
let allPassed = true;
cases.forEach(c => {
  const result = calculateNextMonth(c.input);
  const status = result === c.expected ? "✅" : "❌";
  console.log(`Input: ${c.input} -> Result: ${result} (Expected: ${c.expected}) ${status}`);
  if (result !== c.expected) allPassed = false;
});

if (allPassed) {
  console.log("\n✅ TODAS LAS PRUEBAS DE FECHA EXITOSAS");
} else {
  console.log("\n❌ ALGUNAS PRUEBAS FALLARON");
  process.exit(1);
}
