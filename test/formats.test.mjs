// Pruebas de formato de números, lectura de lo que se escribe y traducción de fórmulas.
import assert from 'node:assert/strict';
import { formatValue, parseInput, changeDecimals } from '../apps/calcula/src/format.js';
import { toEnglish, toSpanish } from '../apps/calcula/src/i18n.js';

const f = (v, fmt) => formatValue(v, fmt).text;
assert.equal(f(42000, '"$"#,##0'), '$42,000');
assert.equal(f(-1234.5, '"$"#,##0.00;[Red]-"$"#,##0.00'), '-$1,234.50');
assert.equal(f(0.4234, '0.0%'), '42.3%');
assert.equal(f(46335, 'dd/mm/yyyy'), '09/11/2026');
assert.equal(f(46335, 'd "de" mmmm "de" yyyy'), '9 de noviembre de 2026');
assert.equal(f(0.5625, 'hh:mm'), '13:30');
assert.equal(f(1234567.891, '#,##0.00'), '1,234,567.89');
assert.equal(f(0.1 + 0.2, 'General'), '0.3');
assert.equal(f(true), 'VERDADERO');

assert.deepEqual(parseInput('$1,250.50'), { value: 1250.5, fmt: '"$"#,##0.00' });
assert.deepEqual(parseInput('15%'), { value: 0.15, fmt: '0%' });
assert.equal(parseInput('15/03/2026').value, 46096);
assert.equal(parseInput('verdadero').value, true);
assert.equal(parseInput("'007").value, '007');
assert.equal(changeDecimals('0.0%', 1), '0.00%');

assert.equal(toEnglish('=SUMA(A1:A3)'), '=SUM(A1:A3)');
assert.equal(toEnglish('=SI(B2>10,"Mucho","Poco")'), '=IF(B2>10,"Mucho","Poco")');
assert.equal(toEnglish('=BUSCARV(A2,D:E,2,FALSO)'), '=VLOOKUP(A2,D:E,2,FALSE)');
assert.equal(toEnglish('=CONTAR.SI(A:A,"SUMA(")'), '=COUNTIF(A:A,"SUMA(")');
assert.equal(toSpanish('=AVERAGE(B2:B9)+MAX(C1,C2)'), '=PROMEDIO(B2:B9)+MAX(C1,C2)');
assert.equal(toSpanish('=IFERROR(XLOOKUP(A1,B:B,C:C),0)'), '=SI.ERROR(BUSCARX(A1,B:B,C:C),0)');
console.log('formatos: todas las pruebas pasaron');
