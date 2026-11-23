/* Gemini calls for ExcelMentor: building the problem (topic or take-home), lesson tables, and each mentor turn.
 * Every session needs Gemini; nothing here falls back to canned problems or keyword judging. */
import { allowedFunctions, CONCEPTS, conceptName, GENERATABLE } from './concepts'
import { buildProblem, type BuiltProblem } from './build'
import type { ExtractedFile } from './extract'
import { detectHeaderRow, largestSheet, tableFromSheet } from './extract'
import { GeminiClient, GeminiError, parseJsonLoose } from './gemini'
import { validateBlock } from './lesson'
import { checkTurn, normalizeTurn, TURN_SCHEMA, type TurnCheckContext, type TurnProblem } from './protocol'
import type { Recommendation } from './progress'
import { conceptProgress, scores } from './progress'
import type { ChatMessage, ConceptId, ConceptLesson, LearnerState, MentorTurn, ProblemDraft, ProblemLevel, ProblemPublic, Profile, StageId } from './types'

export class MentorUnavailableError extends Error {}

function requireClient(client: GeminiClient | null): GeminiClient {
  if (!client?.available) throw new MentorUnavailableError('ExcelMentor needs Gemini. Open Mentor settings and add your Gemini API key (or a proxy URL).')
  return client
}

const DESIGNER_SYSTEM = (p: Profile) => `You design realistic business Excel problems for ExcelMentor, an Excel tutor. You write data and tasks as JSON; the app builds the workbook and computes every answer with its formula engine.
Student: role ${p.role}, domain ${p.domain}, Excel level ${p.level}${p.goal ? `, goal: ${p.goal}` : ''}.
Never write any answer or calculated result in the text.`

const COLUMN_TYPES = ['text', 'integer', 'number', 'currency', 'percent', 'date']
const COLUMNS_SCHEMA = { type: 'array', items: { type: 'object', properties: { name: { type: 'string' }, type: { type: 'string', enum: COLUMN_TYPES } }, required: ['name', 'type'] } }
const TABLE_SCHEMA = { type: 'object', properties: { name: { type: 'string' }, columns: COLUMNS_SCHEMA, rows: { type: 'array', items: { type: 'array', items: { type: 'string' } } } }, required: ['columns', 'rows'] }
const TASK_PROPS = {
  rowTasks: { type: 'array', items: { type: 'object', properties: { header: { type: 'string' }, requirement: { type: 'string' }, concept: { type: 'string', enum: GENERATABLE }, formula: { type: 'string' }, format: { type: 'string', enum: COLUMN_TYPES } }, required: ['header', 'requirement', 'concept', 'formula'] } },
  summaryTasks: { type: 'array', items: { type: 'object', properties: { label: { type: 'string' }, requirement: { type: 'string' }, concept: { type: 'string', enum: GENERATABLE }, formula: { type: 'string' }, format: { type: 'string', enum: COLUMN_TYPES }, underColumn: { type: 'string' } }, required: ['label', 'requirement', 'concept', 'formula'] } },
  operations: {
    type: 'array',
    items: {
      type: 'object',
      properties: {
        type: { type: 'string', enum: ['sort', 'chart', 'format'] },
        requirement: { type: 'string' },
        column: { type: 'string' },
        order: { type: 'string', enum: ['asc', 'desc'] },
        chartType: { type: 'string', enum: ['column', 'bar', 'line', 'pie'] },
        labelColumn: { type: 'string' },
        valueColumn: { type: 'string' },
        format: { type: 'string', enum: ['currency', 'percent'] },
      },
      required: ['type', 'requirement'],
    },
  },
  instructions: { type: 'array', items: { type: 'string' } },
}
const EXPLANATION_SCHEMA = { type: 'object', properties: { situation: { type: 'string' }, goal: { type: 'string' }, data: { type: 'string' }, think: { type: 'string' } }, required: ['situation', 'goal', 'data', 'think'] }

const PROBLEM_SCHEMA = {
  type: 'object',
  properties: {
    title: { type: 'string' },
    context: { type: 'string' },
    objective: { type: 'string' },
    explanation: EXPLANATION_SCHEMA,
    concepts: { type: 'array', items: { type: 'string', enum: GENERATABLE } },
    estimatedMinutes: { type: 'integer' },
    currencySymbol: { type: 'string' },
    dataset: TABLE_SCHEMA,
    lookup: TABLE_SCHEMA,
    ...TASK_PROPS,
  },
  required: ['title', 'context', 'objective', 'explanation', 'concepts', 'estimatedMinutes', 'dataset', 'rowTasks', 'summaryTasks', 'operations', 'instructions'],
}

const TOKEN_RULES = `- Write formulas with column tokens, not cell addresses:
  [Column]           value of that column in the same row  (row tasks only)
  [Column:all]       the whole column, e.g. =SUM([Total Sales:all])
  [First:all]:[Last:all]  a block of columns (for VLOOKUP tables)
  [Lookup.Column:all]     a column of the lookup table
  [[Summary label]]  another summary cell
  Example row task: {"header":"Total Sales","concept":"sum","formula":"=SUM([Jan]:[Apr])"}
  Example summary: {"label":"Top total","concept":"max","formula":"=MAX([Total Sales:all])","underColumn":"Total Sales"}
- Task sheet: header row, then the data rows. Each rowTask becomes a new column to the right (one result per data row). Each summaryTask is one cell below the table, in the column named by underColumn.
- Concept "sort" -> an operation {type:"sort", column, order}. "chart" -> {type:"chart", chartType, labelColumn, valueColumn}. "formatting" -> {type:"format", column, format:"currency"}. "percent" -> a row task dividing two columns with format "percent".
- Lookup concepts (xlookup, vlookup, index-match) need a lookup table whose first column matches values in the main data.`

function learnerSummary(s: LearnerState): string {
  const parts: string[] = []
  for (const id of GENERATABLE) {
    const p = conceptProgress(s, id)
    if (p.stage === 'locked') continue
    const sc = scores(p)
    parts.push(`${id}: ${p.stage}${p.appliedTotal ? ` (task accuracy ${sc.application}%)` : ''}`)
  }
  return parts.length ? parts.join('; ') : 'No concepts learned yet.'
}

function draftFromAI(raw: Record<string, unknown>): ProblemDraft {
  const d = raw as unknown as ProblemDraft
  const ops: ProblemDraft['operations'] = []
  for (const o of ((raw.operations ?? []) as Record<string, string>[])) {
    if (o.type === 'sort') ops.push({ type: 'sort', column: o.column, order: o.order === 'asc' ? 'asc' : 'desc', requirement: o.requirement })
    else if (o.type === 'chart') ops.push({ type: 'chart', chartType: (['column', 'bar', 'line', 'pie'].includes(o.chartType) ? o.chartType : 'column') as 'column', labelColumn: o.labelColumn, valueColumn: o.valueColumn ?? o.column, requirement: o.requirement })
    else if (o.type === 'format') ops.push({ type: 'format', column: o.column, format: o.format === 'percent' ? 'percent' : 'currency', requirement: o.requirement })
  }
  const lookup = d.lookup && Array.isArray(d.lookup.rows) && d.lookup.rows.length ? d.lookup : undefined
  return { ...d, lookup, operations: ops, rowTasks: d.rowTasks ?? [], summaryTasks: d.summaryTasks ?? [], instructions: d.instructions ?? [] }
}

/** Concepts named by the tasks are added to the concept list (Gemini sometimes forgets to list them). */
function withTaskConcepts(d: ProblemDraft, extra: ConceptId[] = []): ProblemDraft {
  const used = [...d.rowTasks.map((t) => t.concept), ...d.summaryTasks.map((t) => t.concept)]
  d.concepts = [...new Set([...(d.concepts ?? []), ...extra.filter((c) => used.includes(c)), ...used])].filter((c) => GENERATABLE.includes(c))
  return d
}

export interface CreatedProblem {
  built: BuiltProblem
  notice?: string
}

async function generateWithFeedback(
  client: GeminiClient,
  profile: Profile,
  prompt: string,
  schema: Record<string, unknown>,
  build: (raw: Record<string, unknown>) => { built?: BuiltProblem; errors: string[] },
  onStep?: (s: string) => void,
  signal?: AbortSignal,
): Promise<BuiltProblem> {
  const turns: { role: 'user' | 'model'; text: string }[] = [{ role: 'user', text: prompt }]
  let lastErrors: string[] = []
  for (let attempt = 0; attempt < 3; attempt++) {
    onStep?.(attempt === 0 ? 'Gemini is designing your problem…' : `Fixing issues the checks found (attempt ${attempt + 1} of 3)…`)
    let text: string
    try {
      text = await client.generate({ system: DESIGNER_SYSTEM(profile), turns, schema, signal, timeoutMs: 150_000 })
    } catch (e) {
      if (e instanceof GeminiError && ['auth', 'model', 'quota', 'network', 'busy'].includes(e.kind)) throw e
      lastErrors = [e instanceof Error ? e.message : String(e)]
      continue
    }
    let raw: Record<string, unknown>
    try {
      raw = parseJsonLoose<Record<string, unknown>>(text)
    } catch {
      lastErrors = ['The response was not valid JSON.']
      turns.push({ role: 'model', text }, { role: 'user', text: 'That was not valid JSON. Return only the corrected JSON object.' })
      continue
    }
    onStep?.('Checking every answer cell with the formula engine…')
    const res = build(raw)
    if (res.built) return res.built
    lastErrors = res.errors
    turns.push({ role: 'model', text }, { role: 'user', text: `The problem failed these checks:\n- ${res.errors.join('\n- ')}\nFix every issue and return the complete corrected JSON.` })
  }
  throw new Error(`Gemini's problem didn't pass the answer checks after 3 tries (${lastErrors.slice(0, 3).join(' ') || 'unknown issue'}). Try again.`)
}

// ------------------------------------------------------------------ mode 1: topic learning
const ROWS: Record<ProblemLevel, string> = { beginner: '8 to 10', intermediate: '12 to 16', advanced: '16 to 25' }

export async function createTopicProblem(client: GeminiClient | null, profile: Profile, learner: LearnerState, rec: Recommendation, onStep?: (s: string) => void, signal?: AbortSignal): Promise<CreatedProblem> {
  const c = requireClient(client)
  const concepts = [...rec.newConcepts, ...rec.revision]
  const fns = [...allowedFunctions(concepts)].join(', ')
  const attach = profile.attachments.map((a) => `--- ${a.name} ---\n${a.text.slice(0, 2500)}`).join('\n')
  const avoid = learner.history.slice(-5).map((h) => h.title)
  const prompt = `Create one realistic business Excel problem as JSON.

STUDENT
Role: ${profile.role}. Domain: ${profile.domain}. Excel level: ${profile.level}. Goal: ${profile.goal ?? 'not given'}. Preferred difficulty: ${profile.difficulty}. Session length: ${profile.duration} minutes.
What they want to learn: ${profile.description || 'not given'}
${attach ? `Their attached material (use it for context and realism, never copy personal data):\n${attach}\n` : ''}Progress so far: ${learnerSummary(learner)}
${avoid.length ? `Avoid repeating these previous problems: ${avoid.join('; ')}.` : ''}

REQUIREMENTS
- Make the problem something a ${profile.role} in ${profile.domain} would really be asked to do.
- Difficulty: ${rec.level}. Dataset: ${ROWS[rec.level]} rows of realistic ${profile.domain} data (realistic names, categories, numbers, dates as YYYY-MM-DD), logically consistent, no duplicate rows.
- Concepts (use exactly these; every one must be practised by at least one task): ${concepts.join(', ')}. New concepts to teach: ${rec.newConcepts.join(', ') || 'none (revision problem)'}.
- Formulas may use only these functions: ${fns}. Arithmetic (+ - * /) is allowed.
- Every task must be solvable with the given data. Instructions describe the BUSINESS need ("Calculate the total sales for each representative"), never the formula. Never write any answer or result value in the text.
${TOKEN_RULES}
- Row values are strings in JSON; numbers without thousands separators or currency symbols; percents as decimals (0.25).
- explanation: situation, goal, data (what the columns mean) and a "think" question for the student.`
  const built = await generateWithFeedback(
    c,
    profile,
    prompt,
    PROBLEM_SCHEMA,
    (raw) => buildProblem({ draft: withTaskConcepts(draftFromAI(raw), concepts), level: rec.level, domain: profile.domain, newConcepts: rec.newConcepts, source: 'gemini', mode: 'topic' }),
    onStep,
    signal,
  )
  return { built }
}

// ------------------------------------------------------------------ mode 2: take-home assessment
const TABLE_REF = { type: 'object', properties: { file: { type: 'string' }, sheet: { type: 'string' }, header_row: { type: 'integer' } }, required: ['file', 'sheet', 'header_row'] }
const ASSESSMENT_SCHEMA = {
  type: 'object',
  properties: {
    title: { type: 'string' },
    context: { type: 'string' },
    objective: { type: 'string' },
    explanation: EXPLANATION_SCHEMA,
    brief_summary: { type: 'array', items: { type: 'string' } },
    level: { type: 'string', enum: ['beginner', 'intermediate', 'advanced'] },
    concepts: { type: 'array', items: { type: 'string', enum: GENERATABLE } },
    estimatedMinutes: { type: 'integer' },
    currencySymbol: { type: 'string' },
    main_table: TABLE_REF,
    lookup_table: TABLE_REF,
    dataset: TABLE_SCHEMA,
    lookup: TABLE_SCHEMA,
    ...TASK_PROPS,
    unchecked_requirements: { type: 'array', items: { type: 'string' } },
  },
  required: ['title', 'context', 'objective', 'explanation', 'brief_summary', 'level', 'concepts', 'estimatedMinutes', 'rowTasks', 'summaryTasks', 'operations', 'instructions', 'unchecked_requirements'],
}

function filesForPrompt(files: ExtractedFile[]): string {
  return files
    .map((f) => {
      if (f.kind === 'document') return `=== DOCUMENT: ${f.name} ===\n${f.text}`
      const sheets = (f.sheets ?? []).map((s) => `--- sheet "${s.name}" (${s.rows} rows × ${s.cols} columns used); first rows, numbered as in Excel:\n${s.preview.map((r, i) => `${i + 1} | ${r.join(' | ')}`).join('\n')}`)
      return `=== SPREADSHEET: ${f.name} ===\n${sheets.join('\n')}`
    })
    .join('\n\n')
}

/** Builds the session problem from an uploaded take-home: the real data becomes the Task sheet; the brief's questions become its tasks. */
export async function createAssessmentProblem(client: GeminiClient | null, profile: Profile, learner: LearnerState, files: ExtractedFile[], note: string, onStep?: (s: string) => void, signal?: AbortSignal): Promise<CreatedProblem> {
  const c = requireClient(client)
  if (!files.length) throw new Error('Upload at least one file from your take-home assessment.')
  const sheetsAvailable = files.some((f) => f.kind === 'spreadsheet' && f.sheets?.length)
  const fns = [...allowedFunctions(GENERATABLE)].join(', ')
  const prompt = `A student uploaded a take-home Excel assessment from an employer. Turn it into a mentored practice session as JSON. The student will solve their REAL tasks on their REAL data; you only describe the structure — the app computes every answer with its formula engine and never shows them.

STUDENT
Role: ${profile.role}. Domain: ${profile.domain}. Excel level: ${profile.level}.
${note.trim() ? `Student's note about the assessment: ${note.trim()}\n` : ''}
FILES
${filesForPrompt(files)}

WHAT TO RETURN
- title (short), context (the business situation in the brief, 2–4 sentences), objective (what the employer wants overall), explanation {situation, goal, data (what the columns mean), think (a question to get the student thinking)}.
- brief_summary: every question or deliverable in the brief, one short line each, in order. Never answer them.
- level: how hard the assessment is. estimatedMinutes: realistic time to learn and solve it with a mentor.
${sheetsAvailable
    ? `- main_table: the table the tasks are computed on: the spreadsheet file name, the sheet name, and header_row = the Excel row number (as numbered above) of its header row. The app reads the table from there (until the first empty row) — do NOT copy its rows.
- lookup_table: only if a second table (e.g. rates, targets, a product list) is needed for lookups — same format. Omit otherwise.`
    : `- There is no spreadsheet. Build "dataset" (and "lookup" if needed) from the tables in the documents EXACTLY as given (do not invent rows). If the documents contain no data at all, create realistic data that matches the brief's description and say in context that it is sample data.`}
- Mirror every requirement in the brief that can be answered with a formula, a sort, a chart or a number format as a task: rowTasks (a new result column computed for every row), summaryTasks (single cells), operations (sort/chart/format). Requirement text describes the business need in the brief's own terms, never the formula.
- Column tokens must use the EXACT header names of the table.
- unchecked_requirements: every part of the brief that is NOT mirrored as a task (PivotTables, conditional formatting, dashboards, written commentary, presentations, anything needing a function not allowed below) — one short line each.
- concepts: the concept ids the mirrored tasks use (every listed concept must be used by a task).
- Formulas may use only these functions: ${fns}. Arithmetic (+ - * /) is allowed. If a requirement needs another function, put it in unchecked_requirements instead.
${TOKEN_RULES}
- Never write any answer or calculated value anywhere in the text.`
  const fileByName = (name: string) => files.find((f) => f.name === name) ?? files.find((f) => f.name.toLowerCase() === name.trim().toLowerCase()) ?? files.find((f) => f.kind === 'spreadsheet')
  let summary = ''
  let unchecked: string[] = []
  const built = await generateWithFeedback(
    c,
    profile,
    prompt,
    ASSESSMENT_SCHEMA,
    (raw) => {
      const r = raw as Record<string, unknown> & { main_table?: { file: string; sheet: string; header_row: number }; lookup_table?: { file: string; sheet: string; header_row: number }; level?: ProblemLevel; brief_summary?: string[]; unchecked_requirements?: string[] }
      const errors: string[] = []
      const draft = draftFromAI(r)
      /** Reads a table, finding the header row itself when Gemini's guess doesn't work. */
      const readTable = (ref: { file?: string; sheet?: string; header_row?: number } | undefined) => {
        const f = (ref?.file && fileByName(ref.file)) || files.find((x) => x.kind === 'spreadsheet' && x.doc)
        if (!f?.doc) return null
        const sheet = (ref?.sheet && f.sheets?.find((x) => x.name.toLowerCase() === ref.sheet!.trim().toLowerCase())?.name) || largestSheet(f)
        if (!sheet) return null
        const guess = Number.isFinite(ref?.header_row) ? Math.max(0, Math.round(ref!.header_row!) - 1) : -1
        const tries = [guess, detectHeaderRow(f.doc, sheet) ?? -1, 0].filter((x, i, a) => x >= 0 && a.indexOf(x) === i)
        for (const row of tries) {
          const t = tableFromSheet(f.doc, sheet, row)
          if (!('error' in t)) return { ...t, sheet }
        }
        return null
      }
      if (sheetsAvailable) {
        const t = readTable(r.main_table)
        if (t) draft.dataset = { columns: t.columns, rows: t.rows }
        else errors.push('No data table could be read from the uploaded spreadsheet: say which file, sheet and header row hold the data (main_table).')
        // a lookup table that can't be read is left out; tasks that needed it are covered in conversation instead
        const l = r.lookup_table ? readTable(r.lookup_table) : null
        draft.lookup = l && l.sheet !== t?.sheet ? { name: l.sheet, columns: l.columns, rows: l.rows } : undefined
      } else if (!draft.dataset?.rows?.length) errors.push('The dataset is missing.')
      if (errors.length) return { errors }
      summary = (r.brief_summary ?? []).map((x) => `• ${x}`).join('\n')
      unchecked = (r.unchecked_requirements ?? []).filter((x) => typeof x === 'string' && x.trim())
      const level = (['beginner', 'intermediate', 'advanced'] as const).includes(r.level as never) ? (r.level as ProblemLevel) : 'intermediate'
      const d = withTaskConcepts(draft)
      // teach what the student hasn't shown they can apply yet (at most 6)
      const known = (id: ConceptId) => ['applying', 'strong', 'mastered'].includes(conceptProgress(learner, id).stage)
      let teach = d.concepts.filter((x) => !known(x))
      if (!teach.length) teach = d.concepts.slice(0, 2)
      return buildProblem({ draft: d, level, domain: profile.domain, newConcepts: teach.slice(0, 6), source: 'assessment', mode: 'takehome', brief: summary, uncheckedRequirements: unchecked })
    },
    onStep,
    signal,
  )
  return { built }
}

// ------------------------------------------------------------------ lesson tables (worked example + practice)
const BLOCK = { type: 'object', properties: { prompt: { type: 'string' }, headers: { type: 'array', items: { type: 'string' } }, rows: { type: 'array', items: { type: 'array', items: { type: 'string' } } }, resultLabel: { type: 'string' }, formula: { type: 'string' } }, required: ['prompt', 'headers', 'rows', 'resultLabel', 'formula'] }
const LESSON_SCHEMA = {
  type: 'object',
  properties: {
    example: { type: 'object', properties: { title: { type: 'string' }, headers: { type: 'array', items: { type: 'string' } }, rows: { type: 'array', items: { type: 'array', items: { type: 'string' } } }, resultLabel: { type: 'string' }, formula: { type: 'string' }, note: { type: 'string' } }, required: ['title', 'headers', 'rows', 'resultLabel', 'formula', 'note'] },
    practice: BLOCK,
  },
  required: ['example', 'practice'],
}

const toCell = (v: string | number) => (typeof v === 'number' ? v : /^-?\d+(\.\d+)?$/.test(String(v).trim()) ? Number(v) : String(v))

export function builtinLesson(concept: ConceptId): ConceptLesson {
  return { ...structuredClone(CONCEPTS[concept].lesson), concept, source: 'built-in' }
}

/** A worked example and a practice exercise in the problem's world; any table that doesn't calculate is replaced by a verified one. */
export async function generateLesson(client: GeminiClient | null, concept: ConceptId, pub: ProblemPublic, profile: Profile): Promise<ConceptLesson> {
  const fb = builtinLesson(concept)
  if (!client?.available) return fb
  const c = CONCEPTS[concept]
  try {
    const raw = await client.generateJson<{ example: ConceptLesson['example']; practice: ConceptLesson['practice'] }>({
      system: DESIGNER_SYSTEM(profile),
      schema: LESSON_SCHEMA,
      timeoutMs: 45_000,
      turns: [
        {
          role: 'user',
          text: `Write two small tables for teaching "${c.name}" (${c.summary}) in the world of this problem: "${pub.title}" — ${pub.context}
- example: a worked example (2–3 columns, 3–4 rows) with DIFFERENT data from the problem, a resultLabel and a formula; note: one sentence explaining it.
- practice: an exercise (2–3 columns, 2–4 rows) with different data again; prompt: what to calculate, in business words, without the formula.
Formulas use column tokens: [Col:all] for a whole column (one result), or [Col] for the same row (one result per row).${c.functions.length ? ` The formula must use ${c.functions.join(' or ')}.` : ''}
Rows are arrays of strings; numbers without symbols.`,
        },
      ],
    })
    const norm = <T extends { rows: (string | number)[][] }>(b: T): T => ({ ...b, rows: b.rows.map((r) => r.map(toCell)) })
    const example = raw.example && !validateBlock(norm(raw.example), concept) ? norm(raw.example) : fb.example
    const practice = raw.practice && !validateBlock(norm(raw.practice), concept) ? norm(raw.practice) : fb.practice
    return { ...fb, example, practice, source: example === fb.example && practice === fb.practice ? 'built-in' : 'gemini' }
  } catch {
    return fb
  }
}

// ------------------------------------------------------------------ one mentor turn
export interface TurnRequest {
  client: GeminiClient | null
  system: string
  /** the conversation so far (without the current event) */
  history: ChatMessage[]
  state: string
  event: string
  stage: StageId
  concept: number
  check: Omit<TurnCheckContext, 'stage' | 'concept'>
  signal?: AbortSignal
}

export interface TurnResult {
  turn: MentorTurn
  /** rules still broken after the correction round */
  problems: TurnProblem[]
  corrected: boolean
}

function historyTurns(history: ChatMessage[]): { role: 'user' | 'model'; text: string }[] {
  const turns: { role: 'user' | 'model'; text: string }[] = []
  for (const m of history.slice(-24)) {
    if (m.kind === 'error') continue
    const role = m.from === 'mentor' && m.kind !== 'event' ? 'model' : 'user'
    const text = m.kind === 'event' ? `[App note shown to the student] ${m.text}` : m.text
    const last = turns[turns.length - 1]
    if (last && last.role === role) last.text += `\n\n${text}`
    else turns.push({ role, text })
  }
  while (turns.length && turns[0].role === 'model') turns.shift()
  return turns
}

export async function mentorTurn(req: TurnRequest): Promise<TurnResult> {
  const client = requireClient(req.client)
  const ctx: TurnCheckContext = { ...req.check, stage: req.stage, concept: req.concept }
  const turns = historyTurns(req.history)
  const current = `[STATE — from the app, authoritative]\n${req.state}\n\n[EVENT]\n${req.event}\n\nReply with the JSON object only.`
  if (turns.length && turns[turns.length - 1].role === 'user') turns[turns.length - 1].text += `\n\n${current}`
  else turns.push({ role: 'user', text: current })
  let best: TurnResult | null = null
  for (let round = 0; round < 2; round++) {
    const text = await client.generate({ system: req.system, turns, schema: TURN_SCHEMA, signal: req.signal, timeoutMs: 90_000 })
    let raw: Partial<MentorTurn>
    try {
      raw = parseJsonLoose<Partial<MentorTurn>>(text)
    } catch {
      turns.push({ role: 'model', text }, { role: 'user', text: 'That was not valid JSON. Reply again with only the JSON object.' })
      continue
    }
    const turn = normalizeTurn(raw, req.stage, req.concept)
    const problems = checkTurn(turn, ctx)
    const result = { turn, problems, corrected: round > 0 }
    if (!problems.length) return result
    if (!best || problems.filter((p) => p.hard).length <= best.problems.filter((p) => p.hard).length) best = result
    turns.push({ role: 'model', text }, { role: 'user', text: `Your reply broke these protocol rules:\n- ${problems.map((p) => p.rule).join('\n- ')}\nRewrite the complete JSON reply so it follows every rule. Keep the same teaching intent.` })
  }
  if (!best) throw new GeminiError('Gemini returned data in an unexpected format.', 200, 'format')
  return best
}

export { conceptName }
