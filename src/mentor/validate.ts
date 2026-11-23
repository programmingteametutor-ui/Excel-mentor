/* Check My Work: validates every required cell / operation and explains what is wrong without giving the answer. */
import type { Workbook } from '../engine/workbook'
import { colToLetters, parseRangeA1 } from '../model/address'
import type { WorkbookDoc } from '../model/types'
import { computeExpected } from './build'
import { CONCEPTS } from './concepts'
import { functionsIn, lettersToCol, referencedCells } from './layout'
import type { CellState, CellTarget, ChartTarget, FormatTarget, SortTarget, Target, TargetResult } from './types'

type Val = number | string | boolean | null | { error: string }

const isErr = (v: unknown): v is { error: string } => !!v && typeof v === 'object' && 'error' in (v as object)

export function sameValue(a: unknown, b: unknown, tol = 1e-6): boolean {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= Math.max(tol, tol * Math.abs(b))
  if (typeof a === 'string' && typeof b === 'string') return a.trim().toLowerCase() === b.trim().toLowerCase()
  if (typeof a === 'boolean' || typeof b === 'boolean') return a === b
  if (typeof a === 'number' && typeof b === 'string') return b.trim() !== '' && Number(b) === a
  return false
}

function methodOk(t: CellTarget, studentFns: string[]): boolean {
  const fns = CONCEPTS[t.concept]?.functions ?? []
  if (!fns.length) return true
  if (t.concept === 'index-match') return studentFns.includes('INDEX') && studentFns.includes('MATCH')
  return fns.some((f) => studentFns.includes(f))
}

/** Describes cells for a student ("E8", "E8 and F8", "3 cells"). */
function describeCells(keys: string[], wb: Workbook, sheetId: string): string {
  const pretty = keys.map((k) => {
    const [sheet, rc] = k.split('|')
    const [r, c] = rc.split(',').map(Number)
    void wb
    void sheetId
    return `${sheet ? `${sheet}!` : ''}${colToLetters(c)}${r + 1}`
  })
  if (pretty.length <= 3) return pretty.join(pretty.length === 2 ? ' and ' : ', ')
  return `${pretty.length} cells (${pretty.slice(0, 2).join(', ')}…)`
}

function normRefs(formula: string, ownSheet: string) {
  const r = referencedCells(formula)
  const own = ownSheet.toLowerCase()
  const cells = new Set([...r.cells].map((k) => (k.startsWith(own + '|') ? '|' + k.slice(own.length + 1) : k)))
  return cells
}

export interface CellDiagnosis {
  state: CellState
  message: string
  expected: string
}

/** Diagnoses one formula target given the expected value (from computeExpected). */
export function diagnoseCell(wb: Workbook, t: CellTarget, expectedValue: unknown): CellDiagnosis & { student: string } {
  const sid = wb.sheetByName(t.sheet)?.id
  if (!sid) return { state: 'incorrect', message: `The ${t.sheet} sheet is missing.`, expected: 'Sheet restored', student: '' }
  const text = wb.getEditText(sid, t.r, t.c)
  const fnName = (CONCEPTS[t.concept]?.functions ?? [])[0]
  const expectedLabel = fnName ? `A formula using ${t.concept === 'index-match' ? 'INDEX and MATCH' : fnName}` : 'A formula'
  if (!text.trim()) return { state: 'not_attempted', message: 'This cell still needs a formula.', expected: 'Formula required', student: '' }
  const isFormula = wb.isFormula(sid, t.r, t.c)
  const val = wb.getValue(sid, t.r, t.c) as Val
  const same = !isErr(val) && sameValue(val, expectedValue, t.tolerance)
  const studentFns = isFormula ? functionsIn(text) : []
  if (!isFormula) {
    if (same) return { state: 'correct_value', message: 'This is the right number, but it was typed in. Use a formula so it updates when the data changes.', expected: expectedLabel, student: text }
    return { state: 'incorrect', message: 'This cell needs a formula. Start with = and refer to the cells instead of typing a number.', expected: expectedLabel, student: text }
  }
  if (isErr(val)) {
    const why = wb.explainError(sid, t.r, t.c)
    return { state: 'incorrect', message: `Your formula shows ${val.error}. ${why ?? 'Check the cell references and brackets.'}`, expected: expectedLabel, student: text }
  }
  const ok = methodOk(t, studentFns)
  if (same && ok) return { state: 'correct', message: 'Correct.', expected: 'Correct', student: text }
  if (same && !ok) {
    const fn = CONCEPTS[t.concept].functions[0]
    return { state: 'correct_value', message: `Your answer gives the correct number, but this task is designed to practise ${t.concept === 'index-match' ? 'INDEX and MATCH' : fn}.`, expected: expectedLabel, student: text }
  }
  // wrong value: explain the closest mistake
  const exp = normRefs(t.formula, t.sheet)
  const got = normRefs(text, t.sheet)
  const missing = [...exp].filter((k) => !got.has(k))
  const extra = [...got].filter((k) => !exp.has(k))
  if (ok && !missing.length && !extra.length) {
    return { state: 'correct_formula', message: 'Your formula is set up the right way, but the result is different because some cells it uses are not correct yet. Fix those cells and check again.', expected: 'Correct formula', student: text }
  }
  const fnUsed = studentFns[0]
  if (!ok) {
    const want = CONCEPTS[t.concept]
    const hint = want.functions.length ? `Think about which function this task needs: ${want.summary}` : `Re-read the task: ${t.requirement}.`
    return { state: 'incorrect', message: `${fnUsed ? `${fnUsed} isn't the right tool here. ` : ''}${hint}`, expected: expectedLabel, student: text }
  }
  if (missing.length && !extra.length) {
    const fn = fnUsed ? `Your ${fnUsed} formula is the right idea, ` : 'You are close, '
    return { state: 'incorrect', message: `${fn}but check whether you included every value it needs.`, expected: `Missing ${describeCells(missing, wb, sid)}`, student: text }
  }
  if (extra.length && !missing.length) {
    return { state: 'incorrect', message: `You're close. Your formula includes cells it shouldn't — check where your range starts and ends.`, expected: `Includes extra ${describeCells(extra, wb, sid)}`, student: text }
  }
  if (missing.length && extra.length) {
    return { state: 'incorrect', message: `The right idea, but the formula points at different cells than this task needs. Check which row or column it should use.`, expected: 'Different cells needed', student: text }
  }
  return { state: 'incorrect', message: `The result isn't what the task asks for. Re-read it: ${t.requirement}.`, expected: expectedLabel, student: text }
}

function checkSort(wb: Workbook, t: SortTarget, original: WorkbookDoc | null): TargetResult {
  const sid = wb.sheetByName(t.sheet)?.id
  const rg = parseRangeA1(t.range)
  const base = { id: t.id, expected: `Sorted by ${t.columnName}, ${t.order === 'desc' ? 'largest first' : 'smallest first'}` }
  if (!sid || !rg) return { ...base, state: 'incorrect', student: '', message: 'The table is missing.' }
  const c = lettersToCol(t.column)
  const vals: unknown[] = []
  for (let r = rg.r1 + 1; r <= rg.r2; r++) vals.push(wb.getValue(sid, r, c))
  if (vals.every((v) => v === null || v === '')) return { ...base, state: 'not_attempted', student: '', message: `Fill in ${t.columnName} first, then sort by it.` }
  const nums = vals.map((v) => (typeof v === 'number' ? v : Number.NaN))
  let sorted = nums.every((n) => !Number.isNaN(n))
  for (let i = 1; sorted && i < nums.length; i++) sorted = t.order === 'desc' ? nums[i - 1] >= nums[i] : nums[i - 1] <= nums[i]
  // rows must stay together: compare data rows with the original data
  let intact = true
  const origSheet = original?.sheets.find((s) => s.name === t.sheet)
  if (origSheet) {
    // compare only the generated data columns (result columns hold the student's formulas)
    const dataSheet = original?.sheets.find((s) => s.name === 'Data')
    const width = (dataSheet?.cells[0] ?? origSheet.cells[0] ?? []).length
    const sig = (row: unknown[]) => JSON.stringify(row.slice(0, width))
    const want = origSheet.cells.slice(rg.r1 + 1, rg.r2 + 1).map((row) => sig(row.map((v) => (typeof v === 'string' && v.startsWith('=') ? null : v))))
    const got: string[] = []
    for (let r = rg.r1 + 1; r <= rg.r2; r++) {
      const row: unknown[] = []
      for (let cc = 0; cc < width; cc++) row.push(wb.getRaw(sid, r, cc) ?? null)
      got.push(sig(row))
    }
    intact = JSON.stringify([...want].sort()) === JSON.stringify([...got].sort())
  }
  if (!intact) return { ...base, state: 'incorrect', student: 'Rows mixed up', message: 'Some rows got split apart. Undo, click one cell inside the table, then sort so whole rows move together.' }
  if (!sorted) return { ...base, state: 'incorrect', student: 'Not sorted', message: `The table isn't sorted by ${t.columnName} yet. Click a cell in that column and use Sort ${t.order === 'desc' ? 'Z→A (largest first)' : 'A→Z (smallest first)'}.` }
  return { ...base, state: 'correct', student: 'Sorted', message: 'Correct.' }
}

function checkChart(wb: Workbook, t: ChartTarget): TargetResult {
  const sid = wb.sheetByName(t.sheet)?.id
  const base = { id: t.id, expected: `A chart of ${t.valueName}` }
  const charts = sid ? wb.meta(sid).charts ?? [] : []
  if (!charts.length) return { ...base, state: 'not_attempted', student: 'No chart', message: 'There is no chart on the Task sheet yet. Select the labels and values, then Insert → Column.' }
  const vc = lettersToCol(t.valueCol)
  const lc = lettersToCol(t.labelCol)
  const good = charts.find((ch) => ch.source.sheetId === sid && ch.source.range.c1 <= Math.min(vc, lc) && ch.source.range.c2 >= Math.max(vc, lc))
  const anyValue = charts.find((ch) => ch.source.range.c1 <= vc && ch.source.range.c2 >= vc)
  if (good) return { ...base, state: 'correct', student: `${good.type} chart`, message: 'Correct.' }
  if (anyValue) return { ...base, state: 'partially_correct', student: `${anyValue.type} chart`, message: `Your chart shows ${t.valueName}, but select the names column too so each bar has a label.` }
  return { ...base, state: 'incorrect', student: 'Chart of other data', message: `The chart doesn't include the ${t.valueName} column. Select the label column and ${t.valueName}, then insert the chart.` }
}

function checkFormat(wb: Workbook, t: FormatTarget): TargetResult {
  const sid = wb.sheetByName(t.sheet)?.id
  const rg = parseRangeA1(t.range)
  const base = { id: t.id, expected: t.format === 'currency' ? 'Currency format' : 'Percent format' }
  if (!sid || !rg) return { ...base, state: 'incorrect', student: '', message: 'The column is missing.' }
  let ok = 0
  let total = 0
  for (let r = rg.r1; r <= rg.r2; r++) {
    total++
    const f = wb.getStyle(sid, r, rg.c1)?.numFmt ?? ''
    if (t.format === 'percent' ? f.includes('%') : /[$₹€£¥]|\[\$|Accounting|"\p{Sc}"/u.test(f)) ok++
  }
  if (ok === total) return { ...base, state: 'correct', student: base.expected, message: 'Correct.' }
  if (ok === 0) return { ...base, state: 'not_attempted', student: 'General', message: `Select the column and apply the ${t.format === 'currency' ? 'Currency' : 'Percent'} format from the Home tab.` }
  return { ...base, state: 'partially_correct', student: `${ok} of ${total} cells`, message: 'Some cells in the column are formatted, but not all of them.' }
}

/** Checks every target. `original` is the generated workbook (for "rows kept together"). */
export function checkWork(wb: Workbook, targets: Target[], original: WorkbookDoc | null): TargetResult[] {
  const cells = targets.filter((t): t is CellTarget => t.kind === 'cell')
  const expected = computeExpected(wb.toDoc(), cells)
  return targets.map((t) => {
    if (t.kind === 'cell') {
      const d = diagnoseCell(wb, t, expected[t.id])
      return { id: t.id, state: d.state, student: d.student, expected: d.expected, message: d.message }
    }
    if (t.kind === 'sort') return checkSort(wb, t, original)
    if (t.kind === 'chart') return checkChart(wb, t)
    return checkFormat(wb, t)
  })
}

export const isCorrect = (s: CellState) => s === 'correct'
/** Beginner-friendly three-state view. */
export function simpleState(s: CellState): 'Correct' | 'Wrong' | 'Not attempted' {
  return s === 'correct' ? 'Correct' : s === 'not_attempted' ? 'Not attempted' : 'Wrong'
}
