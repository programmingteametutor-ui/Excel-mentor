/**
 * Text-level formula utilities that don't need the calculation engine:
 *  - shifting relative references (fill handle, copy/paste, shared formulas in XLSX)
 *  - listing function names (compatibility checks)
 *  - Excel "_xlfn." prefix handling for newer functions in XLSX files
 */
import { colToLetters, lettersToCol } from '../model/address'
import { MAX_COLS, MAX_ROWS } from '../model/types'

type Segment = { text: string; code: boolean }

/** Splits a formula into code segments and literal segments ("strings" and 'sheet names'). */
function segments(formula: string): Segment[] {
  const out: Segment[] = []
  let i = 0
  let buf = ''
  while (i < formula.length) {
    const ch = formula[i]
    if (ch === '"' || ch === "'") {
      if (buf) out.push({ text: buf, code: true })
      buf = ''
      let j = i + 1
      while (j < formula.length) {
        if (formula[j] === ch) {
          if (formula[j + 1] === ch) j += 2
          else break
        } else j++
      }
      out.push({ text: formula.slice(i, j + 1), code: false })
      i = j + 1
    } else {
      buf += ch
      i++
    }
  }
  if (buf) out.push({ text: buf, code: true })
  return out
}

// cell/range reference tokens inside code segments
const REF_RE =
  /(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])|(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(])|(^|[^A-Za-z0-9_.$:])(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_(:])/g

export function shiftFormula(formula: string, dRow: number, dCol: number, maxRows = MAX_ROWS, maxCols = MAX_COLS): string {
  if (!formula.startsWith('=') || (dRow === 0 && dCol === 0)) return formula
  const shiftCol = (abs: string, letters: string): string | null => {
    const c = lettersToCol(letters)
    if (abs || dCol === 0) return abs + letters.toUpperCase()
    const n = c + dCol
    return n < 0 || n >= maxCols ? null : colToLetters(n)
  }
  const shiftRow = (abs: string, digits: string): string | null => {
    const r = parseInt(digits, 10) - 1
    if (abs || dRow === 0) return abs + digits
    const n = r + dRow
    return n < 0 || n >= maxRows ? null : String(n + 1)
  }
  return segments(formula)
    .map((seg) => {
      if (!seg.code) return seg.text
      return seg.text.replace(REF_RE, (m, p1, ca, cl, ra, rd, p2, ca1, cl1, ca2, cl2, p3, ra1, rd1, ra2, rd2) => {
        if (cl !== undefined) {
          if (lettersToCol(cl) >= maxCols) return m // not a cell ref (e.g. a name)
          const col = shiftCol(ca, cl)
          const row = shiftRow(ra, rd)
          if (col === null || row === null) return `${p1}#REF!`
          return `${p1}${col}${row}`
        }
        if (cl1 !== undefined) {
          const a = shiftCol(ca1, cl1)
          const b = shiftCol(ca2, cl2)
          if (a === null || b === null) return `${p2}#REF!`
          return `${p2}${a}:${b}`
        }
        const a = shiftRow(ra1, rd1)
        const b = shiftRow(ra2, rd2)
        if (a === null || b === null) return `${p3}#REF!`
        return `${p3}${a}:${b}`
      })
    })
    .join('')
}

/** Upper-case function names used in a formula, in order of first appearance. */
export function functionNames(formula: string): string[] {
  const names = new Set<string>()
  for (const seg of segments(formula)) {
    if (!seg.code) continue
    for (const m of seg.text.matchAll(/(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)) {
      names.add(m[1].toUpperCase().replace(/^_XLFN\.(_XLWS\.)?|^_XLWS\./, ''))
    }
  }
  return [...names]
}

/** Functions Excel stores with a "_xlfn." prefix in XLSX files (introduced after Excel 2007). */
export const XLFN_FUNCTIONS = new Set(
  (
    'ACOT ACOTH AGGREGATE ARABIC BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR ' +
    'CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T ' +
    'COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP ERF.PRECISE ERFC.PRECISE EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST ' +
    'FILTER FLOOR.MATH FLOOR.PRECISE FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HSTACK IFNA IFS IMCOSH IMCOT IMCSC IMCSCH ' +
    'IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NETWORKDAYS.INTL ' +
    'NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI ' +
    'POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ RRI SEC SECH SEQUENCE SHEET SHEETS SINGLE SKEW.P SORT SORTBY STDEV.P STDEV.S SWITCH ' +
    'T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR UNICODE UNIQUE VAR.P VAR.S VSTACK ' +
    'WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST'
  ).split(' '),
)

/** Removes "_xlfn." / "_xlws." prefixes from a formula read from XLSX. */
export function stripXlfn(formula: string): string {
  return segments(formula)
    .map((s) => (s.code ? s.text.replace(/_xlfn\.(_xlws\.)?|_xlws\./gi, '') : s.text))
    .join('')
}

/** Adds "_xlfn." to newer functions so Excel recognizes them when opening the file. */
export function addXlfn(formula: string): string {
  return segments(formula)
    .map((s) =>
      s.code
        ? s.text.replace(/(?<![A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)(\s*\()/g, (m, name, paren) => {
            const up = String(name).toUpperCase()
            if (!XLFN_FUNCTIONS.has(up)) return m
            const ws = up === 'FILTER' || up === 'SORT' || up === 'SORTBY' ? '_xlws.' : ''
            return `_xlfn.${ws}${up}${paren}`
          })
        : s.text,
    )
    .join('')
}

/**
 * HyperFormula implements TRUE/FALSE as zero-argument functions. Excel writes them as bare
 * literals. These two helpers convert between the engine's form and the Excel form.
 */
export function toEngineFormula(formula: string): string {
  if (!formula.startsWith('=')) return formula
  return segments(formula)
    .map((s) => (s.code ? s.text.replace(/(?<![A-Za-z0-9_.!])(TRUE|FALSE)(?![A-Za-z0-9_.(!])/gi, (m) => `${m.toUpperCase()}()`) : s.text))
    .join('')
}

export function toExcelFormula(formula: string): string {
  if (!formula.startsWith('=')) return formula
  return segments(formula)
    .map((s) => (s.code ? s.text.replace(/(?<![A-Za-z0-9_.!])(TRUE|FALSE)\(\s*\)/gi, (_m, w: string) => w.toUpperCase()) : s.text))
    .join('')
}

/** A cell/range reference found in a formula, with its position in the text. */
export type FormulaRef = { sheet: string | null; text: string; start: number; end: number; r1: number; c1: number; r2: number; c2: number }

/**
 * Lists the references in a formula (A1, B2:C9, C:D, 3:5, Sheet2!A1, 'My sheet'!B:B) so the editor can
 * outline each referenced range in its own colour, as Excel does while you type.
 */
export function formulaRefs(formula: string): FormulaRef[] {
  if (!formula.startsWith('=')) return []
  // blank out string literals so "A1" inside quotes isn't treated as a reference
  const text = formula.replace(/"(?:[^"]|"")*"?/g, (m) => ' '.repeat(m.length))
  const re =
    /(^|[^A-Za-z0-9_.$'!])(?:('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})(?![A-Za-z0-9_(!])/g
  const out: FormulaRef[] = []
  let m: RegExpExecArray | null
  while ((m = re.exec(text))) {
    const start = m.index + m[1].length
    const sheetRaw = m[2]
    const ref = m[3].replace(/\$/g, '')
    const [a, b = a] = ref.split(':')
    let r1: number, c1: number, r2: number, c2: number
    if (/^\d+$/.test(a)) {
      r1 = Number(a) - 1
      r2 = Number(b) - 1
      c1 = 0
      c2 = MAX_COLS - 1
    } else if (/^[A-Za-z]+$/.test(a)) {
      c1 = lettersToCol(a.toUpperCase())
      c2 = lettersToCol(b.toUpperCase())
      r1 = 0
      r2 = MAX_ROWS - 1
    } else {
      const pa = /^([A-Za-z]+)(\d+)$/.exec(a)!
      const pb = /^([A-Za-z]+)(\d+)$/.exec(b)
      if (!pb) continue
      c1 = lettersToCol(pa[1].toUpperCase())
      r1 = Number(pa[2]) - 1
      c2 = lettersToCol(pb[1].toUpperCase())
      r2 = Number(pb[2]) - 1
    }
    if (r1 < 0 || r2 < 0 || c1 >= MAX_COLS || c2 >= MAX_COLS || r1 >= MAX_ROWS || r2 >= MAX_ROWS) continue
    const sheet = sheetRaw ? (sheetRaw.startsWith("'") ? sheetRaw.slice(1, -1).replace(/''/g, "'") : sheetRaw) : null
    out.push({ sheet, text: formula.slice(start, m.index + m[0].length), start, end: m.index + m[0].length, r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) })
  }
  return out
}

/** Excel-like colours for formula references, in order of appearance. */
export const REF_COLORS = ['#2f6fd6', '#d93025', '#7c3fc4', '#188038', '#c26401', '#0b8a8a', '#c2185b', '#6b5b00']
