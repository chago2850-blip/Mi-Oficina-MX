// Ecuaciones de Word (OMML) -> LaTeX, para editarlas en Escribiente. GPL-3.0
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const kids = (el, name) => [...el.childNodes].filter((n) => n.nodeType === 1 && n.namespaceURI === M && (!name || n.localName === name));
const kid = (el, name) => kids(el, name)[0];
const attr = (el, name) => el?.getAttributeNS(M, 'val') || el?.getAttribute('m:val') || (name ? el?.getAttribute(name) : '') || '';
const prop = (el, pr, name) => attr(kid(kid(el, pr) || el, name) || null);

const SYM = { '∑': '\\sum', '∏': '\\prod', '∐': '\\coprod', '∫': '\\int', '∬': '\\iint', '∭': '\\iiint', '∮': '\\oint', '⋃': '\\bigcup', '⋂': '\\bigcap' };
const ACC = { '̂': '\\hat', '^': '\\hat', '̃': '\\tilde', '~': '\\tilde', '̄': '\\bar', '¯': '\\bar', '⃗': '\\vec', '→': '\\vec', '̇': '\\dot', '̈': '\\ddot', '̌': '\\check', '́': '\\acute', '̀': '\\grave', '̆': '\\breve' };
const FUNCS = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'log', 'ln', 'exp', 'lim', 'max', 'min', 'sinh', 'cosh', 'tanh', 'arcsin', 'arccos', 'arctan', 'det', 'gcd']);
const escTxt = (s) => s.replace(/\\/g, '\\backslash ').replace(/([{}#%&_$^])/g, '\\$1');
const grp = (s) => (s.length === 1 ? s : `{${s}}`);

export function ommlToLatex(el) { return conv(el).replace(/\s+/g, ' ').trim(); }
function seq(el) { return [...el.childNodes].filter((n) => n.nodeType === 1).map(conv).join(''); }
function conv(el) {
  if (!el) return '';
  if (el.namespaceURI !== M) {
    // w:r dentro de una ecuación (texto normal) o propiedades
    if (el.localName === 'r') return [...el.getElementsByTagName('*')].filter((t) => t.localName === 't').map((t) => escTxt(t.textContent)).join('');
    return '';
  }
  switch (el.localName) {
    case 'oMathPara': return kids(el, 'oMath').map(conv).join(' \\\\ ');
    case 'oMath': case 'e': case 'num': case 'den': case 'sub': case 'sup': case 'deg': case 'lim': case 'fName': return seq(el);
    case 'r': {
      const t = [...el.childNodes].filter((n) => n.nodeType === 1 && n.localName === 't').map((n) => n.textContent).join('');
      const plain = !!kid(kid(el, 'rPr') || el, 'nor') || attr(kid(kid(el, 'rPr') || el, 'sty')) === 'p';
      if (FUNCS.has(t.trim())) return `\\${t.trim()} `;
      if (plain && /[A-Za-z]{2,}/.test(t)) return `\\text{${t}}`;
      return escTxt(t);
    }
    case 'f': {
      const type = prop(el, 'fPr', 'type');
      const n = seq(kid(el, 'num')), d = seq(kid(el, 'den'));
      return type === 'lin' ? `${grp(n)}/${grp(d)}` : type === 'noBar' ? `\\genfrac{}{}{0pt}{}{${n}}{${d}}` : `\\frac{${n}}{${d}}`;
    }
    case 'sSup': return `${grp(seq(kid(el, 'e')))}^{${seq(kid(el, 'sup'))}}`;
    case 'sSub': return `${grp(seq(kid(el, 'e')))}_{${seq(kid(el, 'sub'))}}`;
    case 'sSubSup': return `${grp(seq(kid(el, 'e')))}_{${seq(kid(el, 'sub'))}}^{${seq(kid(el, 'sup'))}}`;
    case 'sPre': return `{}_{${seq(kid(el, 'sub'))}}^{${seq(kid(el, 'sup'))}}${grp(seq(kid(el, 'e')))}`;
    case 'rad': {
      const hide = prop(el, 'radPr', 'degHide') === '1' || prop(el, 'radPr', 'degHide') === 'on';
      const d = seq(kid(el, 'deg'));
      return !hide && d ? `\\sqrt[${d}]{${seq(kid(el, 'e'))}}` : `\\sqrt{${seq(kid(el, 'e'))}}`;
    }
    case 'nary': {
      const chr = prop(el, 'naryPr', 'chr') || '∫';
      const sub = seq(kid(el, 'sub')), sup = seq(kid(el, 'sup'));
      return `${SYM[chr] || escTxt(chr)}${sub ? `_{${sub}}` : ''}${sup ? `^{${sup}}` : ''} ${seq(kid(el, 'e'))}`;
    }
    case 'd': {
      const pr = kid(el, 'dPr');
      const beg = pr && kid(pr, 'begChr') ? attr(kid(pr, 'begChr')) : '(';
      const end = pr && kid(pr, 'endChr') ? attr(kid(pr, 'endChr')) : ')';
      const sep = pr && kid(pr, 'sepChr') ? attr(kid(pr, 'sepChr')) : '|';
      const L = (c) => (c === '' ? '.' : c === '{' ? '\\{' : c === '}' ? '\\}' : c === '|' ? '|' : c === '‖' ? '\\|' : c === '⟨' || c === '〈' || c === '<' ? '\\langle' : c === '⟩' || c === '〉' || c === '>' ? '\\rangle' : c === '⟦' ? '\\llbracket' : c === '⟧' ? '\\rrbracket' : c === '⌊' ? '\\lfloor' : c === '⌋' ? '\\rfloor' : c === '⌈' ? '\\lceil' : c === '⌉' ? '\\rceil' : c);
      // delimitadores que LaTeX no admite: se escriben como texto
      const ok = (c) => /^(\\?.|\\[a-z]+)$/.test(c) && !/[^\x20-\x7e]/.test(c);
      if (!ok(L(beg)) || !ok(L(end))) return `\\mathrm{${escTxt(beg)}}${kids(el, 'e').map(seq).join(escTxt(sep))}\\mathrm{${escTxt(end)}}`;
      return `\\left${L(beg)}${kids(el, 'e').map(seq).join(sep === '|' ? '\\middle|' : escTxt(sep))}\\right${L(end)}`;
    }
    case 'func': return `${seq(kid(el, 'fName'))}${grp(seq(kid(el, 'e')))}`;
    case 'acc': { const c = prop(el, 'accPr', 'chr') || '̂'; return `${ACC[c] || '\\hat'}{${seq(kid(el, 'e'))}}`; }
    case 'bar': return `${prop(el, 'barPr', 'pos') === 'bot' ? '\\underline' : '\\overline'}{${seq(kid(el, 'e'))}}`;
    case 'groupChr': { const pos = prop(el, 'groupChrPr', 'pos'); return `${pos === 'top' ? '\\overbrace' : '\\underbrace'}{${seq(kid(el, 'e'))}}`; }
    case 'limLow': { const e = seq(kid(el, 'e')); const l = seq(kid(el, 'lim')); return /^\\(lim|max|min)\s*$/.test(e) ? `${e.trim()}_{${l}}` : `\\underset{${l}}{${e}}`; }
    case 'limUpp': return `\\overset{${seq(kid(el, 'lim'))}}{${seq(kid(el, 'e'))}}`;
    case 'm': return `\\begin{matrix}${kids(el, 'mr').map((r) => kids(r, 'e').map(seq).join(' & ')).join(' \\\\ ')}\\end{matrix}`;
    case 'eqArr': return `\\begin{aligned}${kids(el, 'e').map(seq).join(' \\\\ ')}\\end{aligned}`;
    case 'box': case 'borderBox': case 'phant': return seq(kid(el, 'e'));
    default: return /Pr$/.test(el.localName) ? '' : seq(el);
  }
}
export const OMML_NS = M;
