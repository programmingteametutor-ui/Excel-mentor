/* Turns a ProblemDraft into an Excel workbook plus a hidden answer model, and checks it (guardrails). */
import { emptyMeta, newId, Workbook } from '../engine/workbook'
import { engineFunctions } from '../engine/compat'
import { colToLetters } from '../model/address'
import type { CellStyle, SheetDoc, WorkbookDoc } from '../model/types'
import { allowedFunctions, CONCEPTS, GENERATABLE } from './concepts'
import { functionsIn, translateFormula, type LayoutContext, type TableLayout } from './layout'
import { hashString } from './random'
import type { CellTarget, ColumnType, ConceptId, Domain, ProblemDraft, ProblemLevel, ProblemPublic, SessionMode, Target } from './types'

export const SHEETS = { instructions: 'Instructions', task: 'Task', data: 'Data', example: 'Learning Example', practice: 'Practice' } as const

export const HEADER_STYLE: CellStyle = { bold: true, fill: '#0f766e', color: '#ffffff', vAlign: 'middle' }
export const WORK_HEADER_STYLE: CellStyle = { bold: true, fill: '#fde68a', color: '#3d2c00', vAlign: 'middle' }
export const TARGET_STYLE: CellStyle = { fill: '#fff7d1', borders: { top: { style: 'thin', color: '#e0b400' }, bottom: { style: 'thin', color: '#e0b400' }, left: { style: 'thin', color: '#e0b400' }, right: { style: 'thin', color: '#e0b400' } } }

const MIN_ROWS: Record<ProblemLevel, number> = { beginner: 5, intermediate: 8, advanced: 10 }

export function numFmt(type: ColumnType | undefined, currency = '$', decimals = false): string | undefined {
  switch (type) {
    case 'currency':
      return `${currency}#,##0${decimals ? '.00' : ''}`
    case 'percent':
      return '0.0%'
    case 'date':
      return 'dd-mmm-yyyy'
    case 'number':
      return '#,##0.00'
    case 'integer':
      return '#,##0'
    default:
      return undefined
  }
}

export function dateSerial(v: string): number | null {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(v.trim())
  if (!m) return null
  const t = Date.UTC(+m[1], +m[2] - 1, +m[3])
  if (Number.isNaN(t)) return null
  return Math.round((t - Date.UTC(1899, 11, 30)) / 86400000)
}

type Cell = string | number | boolean | null

class SheetBuilder {
  cells: Cell[][] = []
  styles: Record<string, CellStyle> = {}
  widths: Record<number, number> = {}
  heights: Record<number, number> = {}
  constructor(public name: string) {}
  set(r: number, c: number, v: Cell, style?: CellStyle) {
    while (this.cells.length <= r) this.cells.push([])
    const row = this.cells[r]
    while (row.length <= c) row.push(null)
    row[c] = v
    if (style) this.styles[`${r},${c}`] = { ...this.styles[`${r},${c}`], ...style }
  }
  style(r: number, c: number, style: CellStyle) {
    this.styles[`${r},${c}`] = { ...this.styles[`${r},${c}`], ...style }
  }
  autoWidth(min = 64, max = 240) {
    const w: number[] = []
    for (const row of this.cells) row.forEach((v, c) => (w[c] = Math.max(w[c] ?? 0, String(v ?? '').length)))
    w.forEach((n, c) => (this.widths[c] = Math.max(min, Math.min(max, Math.round(n * 7.2 + 18)))))
  }
  doc(extra: Partial<SheetDoc> = {}): SheetDoc {
    return { ...emptyMeta(newId(), this.name), cells: this.cells, styles: this.styles, colWidths: this.widths, rowHeights: this.heights, ...extra }
  }
}

export interface BuiltProblem {
  doc: WorkbookDoc
  pub: ProblemPublic
  targets: Target[]
  /** answers computed by the engine, for guardrail checks and tests */
  expected: Record<string, unknown>
}

export interface BuildInput {
  draft: ProblemDraft
  level: ProblemLevel
  domain: Domain
  newConcepts: ConceptId[]
  source: 'gemini' | 'assessment'
  mode?: SessionMode
  /** take-home: summary of the uploaded brief, shown on the Instructions sheet */
  brief?: string
  /** take-home: parts of the brief the app can't check (taught in conversation) */
  uncheckedRequirements?: string[]
}

/** Normalizes AI output in place (numbers given as strings, trimmed names…). */
function normalizeDraft(d: ProblemDraft, realData = false): ProblemDraft {
  const fixTable = (t: { columns: { name: string; type: ColumnType }[]; rows: (string | number)[][] }) => {
    t.columns = t.columns.map((c) => ({ name: String(c.name ?? '').trim(), type: (['text', 'integer', 'number', 'currency', 'percent', 'date'].includes(c.type) ? c.type : 'text') as ColumnType }))
    t.rows = (t.rows ?? []).map((row) =>
      t.columns.map((col, i) => {
        const v = row[i]
        if (col.type === 'text' || col.type === 'date') return v === null || v === undefined ? '' : String(v).trim()
        if (typeof v === 'number') return v
        // real (uploaded) data may have blank cells; generated data may not
        if (realData && (v === null || v === undefined || String(v).trim() === '')) return ''
        const n = Number(String(v ?? '').replace(/[,$₹€£%\s]/g, ''))
        return Number.isFinite(n) ? (col.type === 'percent' && /%/.test(String(v)) ? n / 100 : n) : String(v ?? '')
      }),
    )
  }
  fixTable(d.dataset)
  if (d.lookup) fixTable(d.lookup)
  d.rowTasks = d.rowTasks ?? []
  d.summaryTasks = d.summaryTasks ?? []
  d.operations = d.operations ?? []
  d.instructions = d.instructions ?? []
  d.concepts = [...new Set(d.concepts ?? [])]
  return d
}

/** Builds the workbook and hidden answers, returning the problems found (empty = OK).
 * Uploaded assessments are never rejected for their tasks: a task that can't be checked automatically is moved to
 * the "covered in conversation" list and the rest of the problem is built. Only unusable data is an error. */
export function buildProblem(input: BuildInput): { built?: BuiltProblem; errors: string[] } {
  if (input.source !== 'assessment') {
    const r = buildOnce(input)
    return { built: r.built, errors: r.errors }
  }
  let draft = structuredClone(input.draft)
  const unchecked = [...(input.uncheckedRequirements ?? [])]
  for (let round = 0; round < 10; round++) {
    const r = buildOnce({ ...input, draft, uncheckedRequirements: unchecked })
    if (r.errors.length || !r.drop.length) return { built: r.built, errors: r.errors }
    const gone = new Set(r.drop)
    const keep = <T extends { requirement: string }>(list: T[], p: string) =>
      list.filter((t, i) => {
        if (!gone.has(`${p}${i}`)) return true
        if (!unchecked.includes(t.requirement)) unchecked.push(t.requirement)
        return false
      })
    draft = { ...draft, rowTasks: keep(draft.rowTasks, 'r'), summaryTasks: keep(draft.summaryTasks, 's'), operations: keep(draft.operations, 'o') }
  }
  return { errors: ['The assessment tasks could not be laid out.'] }
}

function buildOnce(input: BuildInput): { built?: BuiltProblem; errors: string[]; drop: string[] } {
  const real = input.source === 'assessment'
  const d = normalizeDraft(structuredClone(input.draft), real)
  const errors: string[] = []
  /** real data: task keys (r0, s1, o2) whose task can't be checked automatically */
  const drop = new Set<string>()
  const fail = (msg: string, key?: string) => (real && key ? drop.add(key) : errors.push(msg))
  if (real) {
    // take the tasks as given: unknown concepts count as general formulas, result columns never clash with data columns
    const dataNames = d.dataset.columns.map((c) => c.name.trim().toLowerCase())
    for (const t of [...d.rowTasks, ...d.summaryTasks]) if (!GENERATABLE.includes(t.concept)) t.concept = 'formula'
    for (const t of d.rowTasks) while (dataNames.includes(t.header.trim().toLowerCase())) t.header = `${t.header} (result)`
    d.concepts = [...new Set([...d.rowTasks.map((t) => t.concept), ...d.summaryTasks.map((t) => t.concept), ...d.operations.map((o) => (o.type === 'format' ? (o.format === 'percent' ? 'percent' : 'formatting') : o.type))])] as ConceptId[]
    if (!d.concepts.length) d.concepts = ['formula']
  }
  const cur = d.currencySymbol || '$'
  // ---------------------------------------------------------------- structural guardrails
  const cols = d.dataset.columns
  if (cols.length < 2 || cols.length > (real ? 60 : 14)) errors.push(`The dataset needs between 2 and ${real ? 60 : 14} columns.`)
  const names = cols.map((c) => c.name.toLowerCase())
  if (names.some((n) => !n)) errors.push('Every column needs a name.')
  if (new Set(names).size !== names.length) errors.push('Column names must be unique.')
  const rows = d.dataset.rows
  // generated data must be clean and small; uploaded assessment data is used as it is (blanks and repeats allowed)
  const minRows = real ? 2 : MIN_ROWS[input.level]
  if (rows.length < minRows) errors.push(`The dataset needs at least ${minRows} rows${real ? '' : ` for a ${input.level} problem`} (it has ${rows.length}).`)
  if (rows.length > (real ? 5000 : 80)) errors.push(`The dataset is too large (max ${real ? 5000 : 80} rows).`)
  if (!real) rows.forEach((row, i) =>
    cols.forEach((c, j) => {
      const v = row[j]
      if (c.type !== 'text' && c.type !== 'date' && typeof v !== 'number') errors.push(`Row ${i + 1}, column “${c.name}” should be a number but is “${v}”.`)
      if (c.type === 'date' && dateSerial(String(v)) === null) errors.push(`Row ${i + 1}, column “${c.name}” should be a date as YYYY-MM-DD.`)
      if (c.type === 'text' && v === '') errors.push(`Row ${i + 1}, column “${c.name}” is empty.`)
    }),
  )
  if (!real) {
    const sig = rows.map((r) => JSON.stringify(r))
    if (new Set(sig).size !== sig.length) errors.push('The dataset contains duplicate rows.')
  }
  const concepts = d.concepts.filter((c) => GENERATABLE.includes(c))
  if (!concepts.length) errors.push('The problem lists no supported concepts.')
  const allowed = allowedFunctions(concepts)
  const engine = engineFunctions()
  const taskHeaders = d.rowTasks.map((t) => t.header.trim().toLowerCase())
  if (taskHeaders.some((h) => names.includes(h))) errors.push('A result column has the same name as a data column.')
  if (!real && !d.rowTasks.length && !d.summaryTasks.length && !d.operations.length) errors.push('The problem has no tasks.')
  if (!real) for (const t of [...d.rowTasks, ...d.summaryTasks]) if (!concepts.includes(t.concept)) errors.push(`Task “${'header' in t ? t.header : t.label}” uses concept ${t.concept}, which is not in the concept list.`)
  if (!real) for (const c of concepts) {
    const used = d.rowTasks.some((t) => t.concept === c) || d.summaryTasks.some((t) => t.concept === c) || d.operations.some((o) => (o.type === 'format' ? c === 'formatting' || c === 'percent' : o.type === c))
    if (!used) errors.push(`Concept “${c}” is listed but no task practises it.`)
  }
  if (errors.length) return { errors, drop: [] }

  // ---------------------------------------------------------------- layout
  const task = new SheetBuilder(SHEETS.task)
  const dataCols = cols.length
  const allCols = [...cols.map((c) => c.name), ...d.rowTasks.map((t) => t.header)]
  const layout: TableLayout = { sheet: SHEETS.task, headerRow: 0, firstRow: 1, lastRow: rows.length, cols: new Map(allCols.map((n, i) => [n.trim().toLowerCase(), i])) }
  const others = new Map<string, TableLayout>()
  let lookupSheet: SheetBuilder | null = null
  if (d.lookup) {
    const lname = (d.lookup.name || 'Lookup').replace(/[\\/?*[\]:]/g, '').slice(0, 28) || 'Lookup'
    lookupSheet = new SheetBuilder(lname)
    others.set(lname.toLowerCase(), { sheet: lname, headerRow: 0, firstRow: 1, lastRow: d.lookup.rows.length, cols: new Map(d.lookup.columns.map((c, i) => [c.name.toLowerCase(), i])) })
    others.set('lookup', others.get(lname.toLowerCase())!)
  }
  allCols.forEach((h, c) => task.set(0, c, h, c < dataCols ? HEADER_STYLE : WORK_HEADER_STYLE))
  const writeValue = (sb: SheetBuilder, r: number, c: number, v: string | number, type: ColumnType) => {
    if (v === '' || v === null || v === undefined) return
    if (type === 'date') sb.set(r, c, dateSerial(String(v)), { numFmt: numFmt('date') })
    else sb.set(r, c, v, numFmt(type, cur) ? { numFmt: numFmt(type, cur) } : undefined)
  }
  rows.forEach((row, i) => cols.forEach((col, j) => writeValue(task, i + 1, j, row[j], col.type)))

  // summary placement
  const summaryRow0 = rows.length + 2
  const summaries = new Map<string, string>()
  const summaryPos = d.summaryTasks.map((s, i) => {
    let c = s.underColumn ? layout.cols.get(s.underColumn.trim().toLowerCase()) : undefined
    if (c === undefined) c = Math.max(1, dataCols)
    if (c < 1) c = 1
    const r = summaryRow0 + i
    summaries.set(s.label.trim().toLowerCase(), `${colToLetters(c)}${r + 1}`)
    return { r, c }
  })
  const ctx: LayoutContext = { main: layout, others, summaries, fromSheet: SHEETS.task }

  // ---------------------------------------------------------------- targets
  const targets: Target[] = []
  // real assessments accept any correct formula: the student's own method counts
  const conceptFns = (c: ConceptId) => (real ? [] : CONCEPTS[c]?.functions ?? [])
  const checkFns = (f: string, where: string, key: string) => {
    for (const fn of functionsIn(f)) {
      if (!engine.has(fn)) fail(`${where} uses ${fn}, which the formula engine can't calculate.`, key)
      else if (!real && !allowed.has(fn)) errors.push(`${where} uses ${fn}, which isn't one of this problem's concepts.`)
    }
  }
  d.rowTasks.forEach((t, k) => {
    const c = dataCols + k
    for (let i = 0; i < rows.length; i++) {
      const r = i + 1
      const tr = translateFormula(t.formula, ctx, r)
      if ('error' in tr) {
        fail(`Column “${t.header}”: ${tr.error}`, `r${k}`)
        return
      }
      if (i === 0) checkFns(tr.formula, `Column “${t.header}”`, `r${k}`)
      task.style(r, c, { ...TARGET_STYLE, ...(numFmt(t.format, cur, t.format === 'currency' && t.concept === 'average') ? { numFmt: numFmt(t.format, cur, t.concept === 'average') } : {}) })
      targets.push({ kind: 'cell', id: `t${k}_${r}`, sheet: SHEETS.task, cell: `${colToLetters(c)}${r + 1}`, r, c, requirement: t.requirement, concept: t.concept, formula: tr.formula, requireFunctions: conceptFns(t.concept), requireFormula: true, tolerance: t.concept === 'round' ? 0.0005 : 1e-6, format: t.format })
    }
  })
  d.summaryTasks.forEach((s, i) => {
    const { r, c } = summaryPos[i]
    const tr = translateFormula(s.formula, ctx)
    if ('error' in tr) {
      fail(`Summary “${s.label}”: ${tr.error}`, `s${i}`)
      return
    }
    checkFns(tr.formula, `Summary “${s.label}”`, `s${i}`)
    task.set(r, Math.max(0, c - 1), s.label, { bold: true, hAlign: 'right' })
    task.style(r, c, { ...TARGET_STYLE, ...(numFmt(s.format, cur, s.concept === 'average' || s.concept === 'averageif') ? { numFmt: numFmt(s.format, cur, s.concept === 'average' || s.concept === 'averageif') } : {}) })
    targets.push({ kind: 'cell', id: `s${i}`, sheet: SHEETS.task, cell: `${colToLetters(c)}${r + 1}`, r, c, requirement: s.requirement, concept: s.concept, formula: tr.formula, requireFunctions: conceptFns(s.concept), requireFormula: true, tolerance: 1e-6, format: s.format })
  })
  const lastCol = allCols.length - 1
  d.operations.forEach((o, i) => {
    if (o.type === 'sort') {
      const c = layout.cols.get(o.column.trim().toLowerCase())
      if (c === undefined) return fail(`Sort column “${o.column}” doesn't exist.`, `o${i}`)
      targets.push({ kind: 'sort', id: `o${i}`, sheet: SHEETS.task, requirement: o.requirement, concept: 'sort', range: `A1:${colToLetters(lastCol)}${rows.length + 1}`, column: colToLetters(c), columnName: allCols[c], order: o.order === 'asc' ? 'asc' : 'desc' })
    } else if (o.type === 'chart') {
      const lc = layout.cols.get(o.labelColumn.trim().toLowerCase())
      const vc = layout.cols.get(o.valueColumn.trim().toLowerCase())
      if (lc === undefined || vc === undefined) return fail('The chart task refers to a column that doesn’t exist.', `o${i}`)
      targets.push({ kind: 'chart', id: `o${i}`, sheet: SHEETS.task, requirement: o.requirement, concept: 'chart', chartType: o.chartType, labelCol: colToLetters(lc), valueCol: colToLetters(vc), valueName: allCols[vc] })
    } else if (o.type === 'format') {
      const c = layout.cols.get(o.column.trim().toLowerCase())
      if (c === undefined) return fail(`Format column “${o.column}” doesn't exist.`, `o${i}`)
      // the column must not already carry that format
      for (let r = 1; r <= rows.length; r++) {
        const st = task.styles[`${r},${c}`]
        if (st?.numFmt) task.styles[`${r},${c}`] = { ...st, numFmt: undefined }
      }
      targets.push({ kind: 'format', id: `o${i}`, sheet: SHEETS.task, requirement: o.requirement, concept: o.format === 'percent' ? 'percent' : 'formatting', range: `${colToLetters(c)}2:${colToLetters(c)}${rows.length + 1}`, format: o.format })
    }
  })
  if (errors.length || drop.size) return { errors, drop: [...drop] }
  task.autoWidth(70, 200)

  // ---------------------------------------------------------------- other sheets
  const ins = new SheetBuilder(SHEETS.instructions)
  let r = 0
  ins.set(r++, 0, d.title, { bold: true, fontSize: 16, color: '#0f766e' })
  r++
  const block = (title: string, lines: string[]) => {
    ins.set(r++, 0, title, { bold: true, fill: '#e6f4f1' })
    for (const l of lines) ins.set(r++, 0, l, { wrap: true })
    r++
  }
  block(real ? 'Your take-home assessment' : 'Business situation', [d.context])
  if (input.brief) block('What the brief asks for (mentor summary)', input.brief.split(/\n+/).filter(Boolean))
  block('Objective', [d.objective])
  block('Your tasks (on the Task sheet — fill the yellow cells)', d.instructions.map((s, i) => `${i + 1}. ${s}`))
  if (input.uncheckedRequirements?.length) block('Also in the brief (your mentor covers these in conversation; they are not checked automatically)', input.uncheckedRequirements.map((s) => `• ${s}`))
  block('How the session works', ['1. Understand the problem: your mentor presents it and teaches you to ask clarifying questions.', '2. Concept roadmap: which Excel concepts you need and why.', '3–4. Learn each concept (worked example on the Learning Example sheet), then practise it on the Practice sheet.', '5. Plan your solution like an analyst.', '6. Solve: fill in the yellow cells on the Task sheet yourself. Ask for a hint on any cell.', '7. Check My Work, then explain how your formulas could break.', '8. Trade-offs and explaining your work to a manager.', '9. Review and report.'])
  ins.widths[0] = 760
  ins.heights[0] = 30

  const data = new SheetBuilder(SHEETS.data)
  cols.forEach((c, j) => data.set(0, j, c.name, HEADER_STYLE))
  rows.forEach((row, i) => cols.forEach((col, j) => writeValue(data, i + 1, j, row[j], col.type)))
  data.autoWidth(70, 200)

  if (lookupSheet && d.lookup) {
    d.lookup.columns.forEach((c, j) => lookupSheet!.set(0, j, c.name, HEADER_STYLE))
    d.lookup.rows.forEach((row, i) => d.lookup!.columns.forEach((col, j) => writeValue(lookupSheet!, i + 1, j, row[j], col.type)))
    lookupSheet.autoWidth(80, 200)
  }

  const ex = new SheetBuilder(SHEETS.example)
  ex.set(0, 0, 'Learning examples', { bold: true, fontSize: 14, color: '#0f766e' })
  ex.set(1, 0, 'Your mentor adds a small worked example here for each new concept. This sheet is separate from your Task.', { italic: true, color: '#5d6670' })
  ex.widths = { 0: 150, 1: 110, 2: 110, 3: 110, 4: 330 }
  const pr = new SheetBuilder(SHEETS.practice)
  pr.set(0, 0, 'Practice', { bold: true, fontSize: 14, color: '#0f766e' })
  pr.set(1, 0, 'Short exercises for each concept appear here. Type your answer in the yellow cell, then click “Check practice” in the mentor panel.', { italic: true, color: '#5d6670' })
  pr.widths = { 0: 150, 1: 110, 2: 110, 3: 110, 4: 330 }

  const sheetDocs: SheetDoc[] = [ins.doc(), task.doc({ freeze: { rows: 1, cols: 0 } }), data.doc({ freeze: { rows: 1, cols: 0 } })]
  if (lookupSheet) sheetDocs.push(lookupSheet.doc())
  sheetDocs.push(ex.doc(), pr.doc())
  const now = new Date().toISOString()
  const doc: WorkbookDoc = {
    schemaVersion: 1,
    id: newId('wb'),
    name: d.title.slice(0, 80),
    activeSheetId: sheetDocs[0].id,
    sheetOrder: sheetDocs.map((s) => s.id),
    sheets: sheetDocs,
    namedRanges: [],
    createdAt: now,
    updatedAt: now,
  }

  // ---------------------------------------------------------------- compute every answer with the engine
  const expected = computeExpected(doc, targets.filter((t): t is CellTarget => t.kind === 'cell'))
  const bad = (v: unknown) => v === null || v === undefined || v === '' || typeof v === 'object'
  if (real) {
    // real data can have blanks: rows of a result column that don't calculate aren't checked; a task with nothing checkable is dropped
    for (let i = targets.length - 1; i >= 0; i--) {
      const t = targets[i]
      if (t.kind !== 'cell' || !bad(expected[t.id])) continue
      if (t.id.startsWith('t')) {
        const k = t.id.slice(1).split('_')[0]
        if (targets.some((x) => x.id.startsWith(`t${k}_`) && !bad(expected[x.id]))) {
          targets.splice(i, 1)
          const st = task.styles[`${t.r},${t.c}`]
          if (st) task.styles[`${t.r},${t.c}`] = { ...st, fill: undefined, borders: undefined }
          continue
        }
        drop.add(`r${k}`)
      } else drop.add(t.id)
    }
    if (drop.size) return { errors, drop: [...drop] }
  }
  for (const t of targets) {
    if (t.kind !== 'cell' || real) continue
    const v = expected[t.id]
    if (v === null || v === undefined || v === '') errors.push(`${t.cell} (“${t.requirement}”) calculates to an empty result.`)
    else if (typeof v === 'object') errors.push(`${t.cell} (“${t.requirement}”) calculates to an error (${(v as { error: string }).error}).`)
  }
  // answers must not leak into the visible text (the employer's own brief is shown as it is)
  const visible = [d.context, d.objective, ...d.instructions, d.explanation.situation, d.explanation.goal, d.explanation.data, d.explanation.think].join(' ')
  for (const t of targets) {
    if (t.kind !== 'cell' || real) continue
    const v = expected[t.id]
    if (typeof v === 'number' && Math.abs(v) >= 100) {
      const forms = [String(v), v.toLocaleString('en-US'), v.toFixed(2)]
      if (forms.some((f) => visible.includes(f))) errors.push(`The instructions reveal the answer for ${t.cell}.`)
    }
  }
  if (errors.length) return { errors: [...new Set(errors)].slice(0, 12), drop: [] }

  const datasetVersion = (hashString(JSON.stringify(d.dataset) + JSON.stringify(d.lookup ?? null)) >>> 0).toString(36)
  const pub: ProblemPublic = {
    problemId: newId('pb'),
    datasetVersion,
    title: d.title,
    context: d.context,
    objective: d.objective,
    explanation: d.explanation,
    concepts,
    newConcepts: input.newConcepts.filter((c) => concepts.includes(c)),
    instructions: d.instructions,
    level: input.level,
    domain: input.domain,
    estimatedMinutes: Math.max(10, Math.min(90, Math.round(d.estimatedMinutes || 20))),
    source: input.source,
    mode: input.mode ?? (real ? 'takehome' : 'topic'),
    ...(input.brief ? { brief: input.brief } : {}),
    ...(input.uncheckedRequirements?.length ? { uncheckedRequirements: input.uncheckedRequirements } : {}),
    targets: targets.map((t) => ({ id: t.id, kind: t.kind, sheet: t.sheet, cell: t.kind === 'cell' ? t.cell : undefined, requirement: t.requirement, concept: t.concept })),
  }
  return { built: { doc, pub, targets, expected }, errors: [], drop: [] }
}

/** Evaluates the hidden formulas in a scratch copy of the workbook. */
export function computeExpected(doc: WorkbookDoc, targets: CellTarget[]): Record<string, unknown> {
  const wb = Workbook.fromDoc(structuredClone(doc))
  const bySheet = new Map<string, { r: number; c: number; input: string }[]>()
  for (const t of targets) {
    const sid = wb.sheetByName(t.sheet)?.id
    if (!sid) continue
    const list = bySheet.get(sid) ?? []
    list.push({ r: t.r, c: t.c, input: t.formula })
    bySheet.set(sid, list)
  }
  for (const [sid, entries] of bySheet) wb.setInputs(sid, entries, 'expected')
  const out: Record<string, unknown> = {}
  for (const t of targets) {
    const sid = wb.sheetByName(t.sheet)?.id
    out[t.id] = sid ? wb.getValue(sid, t.r, t.c) : null
  }
  return out
}
