// Traduce nombres de funciones y errores entre español (lo que ve el usuario) e inglés (lo que se guarda). GPL-3.0
import { esES } from 'hyperformula/i18n/languages';
import { getLang, t } from '../../shared/i18n.js';
// Con la interfaz en inglés, las fórmulas se muestran con los nombres en inglés (como Excel en inglés);
// en español se pueden escribir en cualquiera de los dos idiomas.

// Ajustes para que coincidan con los nombres en español más usados
const EXTRA = {
  XLOOKUP: 'BUSCARX', VLOOKUP: 'BUSCARV', HLOOKUP: 'BUSCARH', FILTER: 'FILTRAR', UNIQUE: 'UNICOS', SORT: 'ORDENAR', SORTBY: 'ORDENARPOR',
  DATEDIF: 'SIFECHA', SEQUENCE: 'SECUENCIA', LET: 'LET', IFS: 'SI.CONJUNTO', SWITCH: 'CAMBIAR', TEXTJOIN: 'UNIRCADENAS', CONCAT: 'CONCAT', IFERROR: 'SI.ERROR', IFNA: 'SI.ND',
  XMATCH: 'COINCIDIRX', SUMIFS: 'SUMAR.SI.CONJUNTO', AVERAGEIFS: 'PROMEDIO.SI.CONJUNTO', MAXIFS: 'MAX.SI.CONJUNTO', MINIFS: 'MIN.SI.CONJUNTO', TODAY: 'HOY', NOW: 'AHORA',
};
export const EN_TO_ES = { ...esES.functions, ...EXTRA };
export const ES_TO_EN = Object.fromEntries(Object.entries(EN_TO_ES).map(([en, es]) => [es.toUpperCase(), en]));
export const ERR_ES = { ...esES.errors, CYCLE: '#¡CIRCULAR!' };
const ERR_EN = { CYCLE: '#CYCLE!', DIV_BY_ZERO: '#DIV/0!', ERROR: '#ERROR!', NA: '#N/A', NAME: '#NAME?', NUM: '#NUM!', REF: '#REF!', SPILL: '#SPILL!', VALUE: '#VALUE!' };

const ID = /([A-Za-zÁÉÍÓÚÜÑáéíóúüñ_][A-Za-z0-9ÁÉÍÓÚÜÑáéíóúüñ._]*)(\s*\()/g;

function mapOutsideStrings(f, fn) {
  // No toca lo que está entre comillas
  return f.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : fn(part))).join('');
}

export function toEnglish(formula) {
  if (typeof formula !== 'string' || !formula.startsWith('=')) return formula;
  return mapOutsideStrings(formula, (p) => p.replace(ID, (all, name, paren) => {
    const en = ES_TO_EN[name.toUpperCase()];
    return (en || name.toUpperCase()) + paren;
  }).replace(/\bVERDADERO\b/gi, 'TRUE').replace(/\bFALSO\b/gi, 'FALSE'));
}
export function toSpanish(formula) {
  if (typeof formula !== 'string' || !formula.startsWith('=')) return formula;
  if (getLang() === 'en') return formula;
  return mapOutsideStrings(formula, (p) => p.replace(ID, (all, name, paren) => (EN_TO_ES[name.toUpperCase()] || name) + paren).replace(/\bTRUE\b/g, 'VERDADERO').replace(/\bFALSE\b/g, 'FALSO'));
}
export function errorText(err) {
  if (getLang() === 'en') return ERR_EN[err?.type] || err?.value || '#ERROR!';
  if (!err) return '#¡ERROR!';
  return ERR_ES[err.type] || err.value || '#¡ERROR!';
}
export const errorEnglish = (err) => ERR_EN[err?.type] || '#VALUE!';

// Catálogo para "Insertar función"
export const FUNCTION_HELP = [
  ['Matemáticas', 'SUM', 'Suma todos los números de un rango.', 'SUMA(número1; [número2]; …)'],
  ['Matemáticas', 'SUMIF', 'Suma las celdas que cumplen una condición.', 'SUMAR.SI(rango; criterio; [rango_suma])'],
  ['Matemáticas', 'SUMIFS', 'Suma con varias condiciones.', 'SUMAR.SI.CONJUNTO(rango_suma; rango1; criterio1; …)'],
  ['Matemáticas', 'ROUND', 'Redondea a un número de decimales.', 'REDONDEAR(número; decimales)'],
  ['Matemáticas', 'ROUNDUP', 'Redondea hacia arriba.', 'REDONDEAR.MAS(número; decimales)'],
  ['Matemáticas', 'ROUNDDOWN', 'Redondea hacia abajo.', 'REDONDEAR.MENOS(número; decimales)'],
  ['Matemáticas', 'INT', 'Parte entera de un número.', 'ENTERO(número)'],
  ['Matemáticas', 'ABS', 'Valor absoluto.', 'ABS(número)'],
  ['Matemáticas', 'MOD', 'Residuo de una división.', 'RESIDUO(número; divisor)'],
  ['Matemáticas', 'POWER', 'Eleva a una potencia.', 'POTENCIA(número; potencia)'],
  ['Matemáticas', 'SQRT', 'Raíz cuadrada.', 'RAIZ(número)'],
  ['Matemáticas', 'PRODUCT', 'Multiplica los números.', 'PRODUCTO(número1; …)'],
  ['Matemáticas', 'SUMPRODUCT', 'Suma de productos de rangos.', 'SUMAPRODUCTO(matriz1; matriz2; …)'],
  ['Matemáticas', 'RAND', 'Número aleatorio entre 0 y 1.', 'ALEATORIO()'],
  ['Matemáticas', 'RANDBETWEEN', 'Entero aleatorio entre dos números.', 'ALEATORIO.ENTRE(inferior; superior)'],
  ['Estadística', 'AVERAGE', 'Promedio de los números.', 'PROMEDIO(número1; …)'],
  ['Estadística', 'AVERAGEIF', 'Promedio de las celdas que cumplen una condición.', 'PROMEDIO.SI(rango; criterio; [rango_promedio])'],
  ['Estadística', 'COUNT', 'Cuenta las celdas con números.', 'CONTAR(valor1; …)'],
  ['Estadística', 'COUNTA', 'Cuenta las celdas no vacías.', 'CONTARA(valor1; …)'],
  ['Estadística', 'COUNTBLANK', 'Cuenta las celdas vacías.', 'CONTAR.BLANCO(rango)'],
  ['Estadística', 'COUNTIF', 'Cuenta las celdas que cumplen una condición.', 'CONTAR.SI(rango; criterio)'],
  ['Estadística', 'COUNTIFS', 'Cuenta con varias condiciones.', 'CONTAR.SI.CONJUNTO(rango1; criterio1; …)'],
  ['Estadística', 'MAX', 'El valor más grande.', 'MAX(número1; …)'],
  ['Estadística', 'MIN', 'El valor más pequeño.', 'MIN(número1; …)'],
  ['Estadística', 'MEDIAN', 'La mediana.', 'MEDIANA(número1; …)'],
  ['Estadística', 'MODE', 'El valor más frecuente.', 'MODA(número1; …)'],
  ['Estadística', 'STDEV', 'Desviación estándar de una muestra.', 'DESVEST(número1; …)'],
  ['Estadística', 'LARGE', 'El k-ésimo valor más grande.', 'K.ESIMO.MAYOR(matriz; k)'],
  ['Estadística', 'SMALL', 'El k-ésimo valor más pequeño.', 'K.ESIMO.MENOR(matriz; k)'],
  ['Estadística', 'RANK', 'Posición de un número en una lista.', 'JERARQUIA(número; ref; [orden])'],
  ['Lógicas', 'IF', 'Devuelve un valor si la condición es verdadera y otro si es falsa.', 'SI(prueba_lógica; valor_si_verdadero; [valor_si_falso])'],
  ['Lógicas', 'IFS', 'Revisa varias condiciones en orden.', 'SI.CONJUNTO(prueba1; valor1; …)'],
  ['Lógicas', 'IFERROR', 'Devuelve otro valor si hay error.', 'SI.ERROR(valor; valor_si_error)'],
  ['Lógicas', 'AND', 'VERDADERO si todas las condiciones lo son.', 'Y(lógico1; …)'],
  ['Lógicas', 'OR', 'VERDADERO si alguna condición lo es.', 'O(lógico1; …)'],
  ['Lógicas', 'NOT', 'Invierte el valor lógico.', 'NO(lógico)'],
  ['Búsqueda', 'VLOOKUP', 'Busca en la primera columna y devuelve un valor de la misma fila.', 'BUSCARV(valor; tabla; columna; [FALSO])'],
  ['Búsqueda', 'HLOOKUP', 'Busca en la primera fila y devuelve un valor de la misma columna.', 'BUSCARH(valor; tabla; fila; [FALSO])'],
  ['Búsqueda', 'XLOOKUP', 'Búsqueda moderna en cualquier dirección.', 'BUSCARX(valor; matriz_búsqueda; matriz_devuelta; [si_no_se_encuentra])'],
  ['Búsqueda', 'INDEX', 'Valor en una fila y columna de un rango.', 'INDICE(matriz; fila; [columna])'],
  ['Búsqueda', 'MATCH', 'Posición de un valor en un rango.', 'COINCIDIR(valor; matriz; [tipo])'],
  ['Búsqueda', 'CHOOSE', 'Elige un valor de una lista.', 'ELEGIR(índice; valor1; …)'],
  ['Búsqueda', 'FILTER', 'Filtra un rango con una condición (matriz dinámica).', 'FILTRAR(matriz; incluir)'],
  ['Texto', 'CONCATENATE', 'Une textos.', 'CONCATENAR(texto1; …)'],
  ['Texto', 'TEXTJOIN', 'Une textos con un separador.', 'UNIRCADENAS(separador; ignorar_vacías; texto1; …)'],
  ['Texto', 'LEFT', 'Caracteres del inicio de un texto.', 'IZQUIERDA(texto; [n])'],
  ['Texto', 'RIGHT', 'Caracteres del final de un texto.', 'DERECHA(texto; [n])'],
  ['Texto', 'MID', 'Caracteres del medio de un texto.', 'EXTRAE(texto; inicio; n)'],
  ['Texto', 'LEN', 'Número de caracteres.', 'LARGO(texto)'],
  ['Texto', 'UPPER', 'Convierte a mayúsculas.', 'MAYUSC(texto)'],
  ['Texto', 'LOWER', 'Convierte a minúsculas.', 'MINUSC(texto)'],
  ['Texto', 'PROPER', 'Primera letra de cada palabra en mayúscula.', 'NOMPROPIO(texto)'],
  ['Texto', 'TRIM', 'Quita espacios sobrantes.', 'ESPACIOS(texto)'],
  ['Texto', 'SUBSTITUTE', 'Reemplaza texto.', 'SUSTITUIR(texto; texto_original; texto_nuevo)'],
  ['Texto', 'TEXT', 'Da formato a un número como texto.', 'TEXTO(valor; formato)'],
  ['Texto', 'FIND', 'Posición de un texto dentro de otro.', 'ENCONTRAR(texto_buscado; dentro_de)'],
  ['Fecha y hora', 'TODAY', 'Fecha de hoy.', 'HOY()'],
  ['Fecha y hora', 'NOW', 'Fecha y hora actuales.', 'AHORA()'],
  ['Fecha y hora', 'DATE', 'Crea una fecha.', 'FECHA(año; mes; día)'],
  ['Fecha y hora', 'YEAR', 'Año de una fecha.', 'AÑO(fecha)'],
  ['Fecha y hora', 'MONTH', 'Mes de una fecha.', 'MES(fecha)'],
  ['Fecha y hora', 'DAY', 'Día de una fecha.', 'DIA(fecha)'],
  ['Fecha y hora', 'WEEKDAY', 'Día de la semana.', 'DIASEM(fecha; [tipo])'],
  ['Fecha y hora', 'DATEDIF', 'Diferencia entre dos fechas.', 'SIFECHA(inicio; fin; unidad)'],
  ['Fecha y hora', 'NETWORKDAYS', 'Días laborables entre dos fechas.', 'DIAS.LAB(inicio; fin; [festivos])'],
  ['Fecha y hora', 'EOMONTH', 'Último día del mes.', 'FIN.MES(fecha; meses)'],
  ['Financieras', 'PMT', 'Pago de un préstamo.', 'PAGO(tasa; nper; va; [vf]; [tipo])'],
  ['Financieras', 'FV', 'Valor futuro de una inversión.', 'VF(tasa; nper; pago; [va]; [tipo])'],
  ['Financieras', 'PV', 'Valor actual.', 'VA(tasa; nper; pago; [vf]; [tipo])'],
  ['Financieras', 'NPV', 'Valor neto actual.', 'VNA(tasa; valor1; …)'],
  ['Financieras', 'IRR', 'Tasa interna de retorno.', 'TIR(valores; [estimar])'],
  ['Financieras', 'RATE', 'Tasa de interés por periodo.', 'TASA(nper; pago; va; …)'],
].map(([cat, en, desc, syntax]) => ({
  en, esName: EN_TO_ES[en] || en, catEs: cat, descEs: desc, syntaxEs: syntax.replace(/;/g, ','),
  get es() { return getLang() === 'en' ? this.en : this.esName; },
  get cat() { return t(this.catEs); },
  get desc() { return t(this.descEs); },
  get syntax() { return getLang() === 'en' ? syntaxEn(this.syntaxEs) : this.syntaxEs; },
}));
const PARAM_EN = { número: 'number', número1: 'number1', número2: 'number2', rango: 'range', rango1: 'range1', criterio: 'criteria', criterio1: 'criteria1', rango_suma: 'sum_range', rango_promedio: 'average_range', decimales: 'num_digits', divisor: 'divisor', potencia: 'power', matriz: 'array', matriz1: 'array1', matriz2: 'array2', inferior: 'bottom', superior: 'top', valor: 'value', valor1: 'value1', ref: 'ref', orden: 'order', prueba_lógica: 'logical_test', valor_si_verdadero: 'value_if_true', valor_si_falso: 'value_if_false', prueba1: 'test1', valor_si_error: 'value_if_error', lógico: 'logical', lógico1: 'logical1', tabla: 'table', columna: 'col_index', fila: 'row_index', FALSO: 'FALSE', matriz_búsqueda: 'lookup_array', matriz_devuelta: 'return_array', si_no_se_encuentra: 'if_not_found', tipo: 'type', índice: 'index', incluir: 'include', texto: 'text', texto1: 'text1', separador: 'delimiter', ignorar_vacías: 'ignore_empty', inicio: 'start', texto_original: 'old_text', texto_nuevo: 'new_text', formato: 'format', texto_buscado: 'find_text', dentro_de: 'within_text', año: 'year', mes: 'month', día: 'day', fecha: 'date', fin: 'end', unidad: 'unit', festivos: 'holidays', meses: 'months', tasa: 'rate', va: 'pv', vf: 'fv', pago: 'pmt', valores: 'values', estimar: 'guess' };
function syntaxEn(s) {
  const m = /^([^(]+)\((.*)\)$/.exec(s); if (!m) return s;
  const name = ES_TO_EN[m[1].toUpperCase()] || m[1];
  return `${name}(${m[2].replace(/[\p{L}_][\p{L}\p{N}_]*/gu, (w) => PARAM_EN[w] || w)})`;
}
