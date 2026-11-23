/* Writes worked examples and practice exercises into their own sheets, and checks practice answers. */
import { Workbook } from '../engine/workbook'
import { engineFunctions } from '../engine/compat'
import { colToLetters } from '../model/address'
import { allowedFunctions, CONCEPTS } from './concepts'
import { computeExpected, HEADER_STYLE, SHEETS, TARGET_STYLE } from './build'
import { functionsIn, translateFormula, type TableLayout } from './layout'
import { diagnoseCell } from './validate'
import type { CellTarget, ConceptId, ConceptLesson, PracticeTarget } from './types'

export interface Block {
  title?: string
  prompt?: string
  headers: string[]
  rows: (string | number)[][]
  resultLabel: string
  formula: string
  note?: string
}

const isPerRow = (formula: string) => /\[[^[\]:]+\](?!\s*:)/.test(formula.replace(/"(?:[^"]|"")*"/g, '').replace(/\[[^[\]]+:(all|first|last)\]/g, ''))

interface Placed {
  layout: TableLayout
  resultR: number
  resultC: number
  perRow: boolean
  formulaAt: (r: number) => string | null
  lastRow: number
}

function place(block: Block, sheet: string, start: number): Placed | { error: string } {
  const perRow = isPerRow(block.formula)
  const cols = new Map(block.headers.map((h, i) => [h.trim().toLowerCase(), i]))
  const layout: TableLayout = { sheet, headerRow: start + 1, firstRow: start + 2, lastRow: start + 1 + block.rows.length, cols }
  const ctx = { main: layout, fromSheet: sheet }
  const resultC = perRow ? block.headers.length : Math.max(1, block.headers.length - 1)
  const resultR = perRow ? layout.firstRow : layout.lastRow + 1
  const formulaAt = (r: number) => {
    const t = translateFormula(block.formula, ctx, perRow ? r : undefined)
    return 'error' in t ? null : t.formula
  }
  const test = translateFormula(block.formula, ctx, perRow ? resultR : undefined)
  if ('error' in test) return { error: test.error }
  return { layout, resultR, resultC, perRow, formulaAt, lastRow: perRow ? layout.lastRow : resultR }
}

/** Returns an error message if a lesson block can't be used. */
export function validateBlock(block: Block, concept: ConceptId): string | null {
  if (!block || !Array.isArray(block.headers) || block.headers.length < 1 || block.headers.length > 5) return 'Example needs 1–5 columns.'
  if (!Array.isArray(block.rows) || block.rows.length < 1 || block.rows.length > 8) return 'Example needs 1–8 rows.'
  if (block.rows.some((r) => !Array.isArray(r) || r.length !== block.headers.length)) return 'Every example row needs a value for each column.'
  const p = place(block, 'Scratch', 0)
  if ('error' in p) return p.error
  const f = p.formulaAt(p.resultR)!
  const allowed = allowedFunctions([concept])
  if (!CONCEPTS[concept].functions.length) for (const fn of ['SUM', 'AVERAGE', 'MAX', 'MIN', 'COUNT', 'COUNTA', 'COUNTIF', 'SUMIF']) allowed.add(fn)
  const engine = engineFunctions()
  for (const fn of functionsIn(f)) if (!engine.has(fn) || !allowed.has(fn)) return `Example uses ${fn}, which doesn't fit this concept.`
  // the practice check requires the concept's own function, so the formula must use it
  const own = CONCEPTS[concept].functions
  if (own.length && !functionsIn(f).some((fn) => own.includes(fn))) return `Example must use ${own.join(' or ')}.`
  const wb = Workbook.blank('scratch', ['Scratch'])
  const sid = wb.order[0]
  writeCells(wb, sid, block, 0, p, true)
  const v = wb.getValue(sid, p.resultR, p.resultC)
  if (v === null || v === '' || (typeof v === 'object' && v && 'error' in v)) return 'Example formula does not calculate.'
  return null
}

function writeCells(wb: Workbook, sid: string, block: Block, start: number, p: Placed, withFormulas: boolean, practice = false) {
  const entries: { r: number; c: number; input: string }[] = []
  const lit = (v: string | number) => (typeof v === 'number' ? String(v) : /^[=+\-@]/.test(v) || /^\d/.test(v) ? `'${v}` : v)
  if (block.title) entries.push({ r: start, c: 0, input: lit(block.title) })
  block.headers.forEach((h, c) => entries.push({ r: start + 1, c, input: lit(h) }))
  if (p.perRow) entries.push({ r: start + 1, c: p.resultC, input: lit(block.resultLabel) })
  block.rows.forEach((row, i) => row.forEach((v, c) => entries.push({ r: start + 2 + i, c, input: typeof v === 'number' ? String(v) : lit(v) })))
  if (!p.perRow) entries.push({ r: p.resultR, c: Math.max(0, p.resultC - 1), input: lit(block.resultLabel) })
  if (withFormulas) {
    if (p.perRow) for (let r = p.layout.firstRow; r <= p.layout.lastRow; r++) entries.push({ r, c: p.resultC, input: p.formulaAt(r)! })
    else entries.push({ r: p.resultR, c: p.resultC, input: p.formulaAt(p.resultR)! })
  }
  const noteCol = p.resultC + 2
  if (block.note) entries.push({ r: start + 1, c: noteCol, input: lit(block.note) })
  if (practice && block.prompt) entries.push({ r: start + 1, c: noteCol, input: lit(block.prompt) })
  // values and labels may collide with the result cell when the block is a single column: drop duplicates keeping the last
  const seen = new Map<string, { r: number; c: number; input: string }>()
  for (const e of entries) seen.set(`${e.r},${e.c}`, e)
  wb.setInputs(sid, [...seen.values()], practice ? 'Mentor practice' : 'Mentor example')
}

function ensureSheet(wb: Workbook, name: string): string {
  const s = wb.sheetByName(name)
  if (s) return s.id
  return wb.addSheet(name)
}

/** Adds a worked example to the Learning Example sheet. Returns the cell holding the example formula. */
export interface WrittenExample {
  cell: { r: number; c: number }
  nextRow: number
  /** formula of the (first) result cell, in A1 form */
  formula: string
  /** every result formula (several for per-row examples), to write when revealing */
  results: { r: number; c: number; input: string }[]
  headers: string[]
  firstDataRow: number
}

/** Adds a worked example to the Learning Example sheet. With withResult=false the result cells stay empty (for a prediction first). */
export function writeExample(wb: Workbook, lesson: ConceptLesson, startRow: number, withResult = true): WrittenExample | { error: string } {
  const sid = ensureSheet(wb, SHEETS.example)
  const block: Block = { ...lesson.example, title: `Example — ${lesson.example.title}` }
  const p = place(block, SHEETS.example, startRow)
  if ('error' in p) return p
  writeCells(wb, sid, block, startRow, p, withResult)
  const cols = Math.max(block.headers.length, p.resultC + 1)
  const ranges = [{ r1: startRow + 1, c1: 0, r2: startRow + 1, c2: cols - 1 }]
  wb.applyStyle(sid, [{ r1: startRow, c1: 0, r2: startRow, c2: 0 }], { bold: true, color: '#0f766e' }, 'Mentor example')
  wb.applyStyle(sid, ranges, HEADER_STYLE, 'Mentor example')
  wb.applyStyle(sid, [{ r1: p.resultR, c1: p.resultC, r2: p.perRow ? p.layout.lastRow : p.resultR, c2: p.resultC }], { bold: true, fill: '#e6f4f1' }, 'Mentor example')
  if (block.note) wb.applyStyle(sid, [{ r1: startRow + 1, c1: p.resultC + 2, r2: startRow + 1, c2: p.resultC + 2 }], { italic: true, color: '#5d6670', fill: undefined, bold: false }, 'Mentor example')
  const results = p.perRow ? Array.from({ length: p.layout.lastRow - p.layout.firstRow + 1 }, (_, i) => ({ r: p.layout.firstRow + i, c: p.resultC, input: p.formulaAt(p.layout.firstRow + i)! })) : [{ r: p.resultR, c: p.resultC, input: p.formulaAt(p.resultR)! }]
  return { cell: { r: p.resultR, c: p.resultC }, nextRow: p.lastRow + 3, formula: results[0].input, results, headers: block.headers, firstDataRow: p.layout.firstRow }
}

/** Writes the example's result formulas (after the student predicted the answer). */
export function revealExample(wb: Workbook, results: { r: number; c: number; input: string }[]) {
  const sid = ensureSheet(wb, SHEETS.example)
  wb.setInputs(sid, results, 'Mentor example')
}

/** Adds a practice exercise to the Practice sheet with one empty yellow answer cell. */
export function writePractice(wb: Workbook, lesson: ConceptLesson, startRow: number): { target: PracticeTarget; nextRow: number } | { error: string } {
  const sid = ensureSheet(wb, SHEETS.practice)
  const src = lesson.practice
  const block: Block = { ...src, title: `Practice — ${src.prompt}` }
  const p = place(block, SHEETS.practice, startRow)
  if ('error' in p) return p
  writeCells(wb, sid, block, startRow, p, false, true)
  const cols = Math.max(block.headers.length, p.resultC + 1)
  wb.applyStyle(sid, [{ r1: startRow, c1: 0, r2: startRow, c2: 0 }], { bold: true, color: '#0f766e' }, 'Mentor practice')
  wb.applyStyle(sid, [{ r1: startRow + 1, c1: 0, r2: startRow + 1, c2: cols - 1 }], HEADER_STYLE, 'Mentor practice')
  wb.applyStyle(sid, [{ r1: p.resultR, c1: p.resultC, r2: p.resultR, c2: p.resultC }], TARGET_STYLE, 'Mentor practice')
  wb.applyStyle(sid, [{ r1: startRow + 1, c1: p.resultC + 2, r2: startRow + 1, c2: p.resultC + 2 }], { italic: true, color: '#5d6670', fill: undefined, bold: false }, 'Mentor practice')
  const formula = p.formulaAt(p.resultR)!
  return {
    target: { concept: lesson.concept, cell: `${colToLetters(p.resultC)}${p.resultR + 1}`, r: p.resultR, c: p.resultC, formula, requireFunctions: [], attempts: 0, passed: false },
    nextRow: p.lastRow + 3,
  }
}

/** Checks the practice answer with the same engine as Check My Work. */
export function checkPractice(wb: Workbook, pt: PracticeTarget): { ok: boolean; state: string; message: string } {
  const t: CellTarget = { kind: 'cell', id: 'practice', sheet: SHEETS.practice, cell: pt.cell, r: pt.r, c: pt.c, requirement: 'the practice result', concept: pt.concept, formula: pt.formula, requireFunctions: pt.requireFunctions, requireFormula: true, tolerance: 1e-6 }
  const expected = computeExpected(wb.toDoc(), [t])
  const d = diagnoseCell(wb, t, expected.practice)
  return { ok: d.state === 'correct', state: d.state, message: d.message }
}

/** Renders a block into a small grid (for the concept review page). The result cell shows the formula and value. */
export function renderBlock(block: Block, withFormula: boolean): { rows: { text: string; formula?: string; target?: boolean }[][]; resultCell: string; firstRow: number } | null {
  const p = place(block, 'Scratch', 0)
  if ('error' in p) return null
  const wb = Workbook.blank('scratch', ['Scratch'])
  const sid = wb.order[0]
  writeCells(wb, sid, { ...block, title: undefined, note: undefined, prompt: undefined }, 0, p, withFormula)
  const out: { text: string; formula?: string; target?: boolean }[][] = []
  const width = Math.max(block.headers.length, p.resultC + 1)
  for (let r = 1; r <= p.lastRow; r++) {
    const row: { text: string; formula?: string; target?: boolean }[] = []
    for (let c = 0; c < width; c++) {
      const isF = wb.isFormula(sid, r, c)
      row.push({ text: wb.getDisplay(sid, r, c).text, formula: isF ? wb.getEditText(sid, r, c) : undefined, target: r === p.resultR && c === p.resultC })
    }
    out.push(row)
  }
  return { rows: out, resultCell: `${colToLetters(p.resultC)}${p.resultR + 1}`, firstRow: 1 }
}

/** Checks a formula typed for a practice block shown at rows starting at 1 (as rendered by renderBlock). */
export function checkBlockAnswer(block: Block, concept: ConceptId, formula: string): { ok: boolean; message: string } {
  const p = place(block, 'Scratch', 0)
  if ('error' in p) return { ok: false, message: p.error }
  const wb = Workbook.blank('scratch', ['Scratch'])
  const sid = wb.order[0]
  writeCells(wb, sid, { ...block, title: undefined, note: undefined, prompt: undefined }, 0, p, false)
  const f = formula.trim().startsWith('=') ? formula.trim() : `=${formula.trim()}`
  wb.setInputs(sid, [{ r: p.resultR, c: p.resultC, input: f }])
  const t: CellTarget = { kind: 'cell', id: 'x', sheet: 'Scratch', cell: `${colToLetters(p.resultC)}${p.resultR + 1}`, r: p.resultR, c: p.resultC, requirement: 'the result', concept, formula: p.formulaAt(p.resultR)!, requireFunctions: [], requireFormula: true, tolerance: 1e-6 }
  const expected = computeExpected(wb.toDoc(), [t])
  const d = diagnoseCell(wb, t, expected.x)
  return { ok: d.state === 'correct', message: d.state === 'correct' ? 'Correct!' : d.message }
}
