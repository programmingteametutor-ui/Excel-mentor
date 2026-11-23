/* The 9-stage session as data: where it may move next, what evidence each gate needs, what the mentor is told each
 * turn, how a validated reply updates the session, and how answer leaks are detected. No React or Gemini here. */
import { colToLetters } from '../model/address'
import type { WorkbookDoc } from '../model/types'
import { SHEETS } from './build'
import { CONCEPTS, conceptName } from './concepts'
import { STAGE_INFO } from './protocol'
import type { CellTarget, ConceptId, MentorSession, MentorTurn, ProblemPublic, StagePosition, StageId, Target } from './types'

export const MAX_HINT = 4

export function newSession(o: { wbId: string; problem: ProblemPublic; notice?: string }): MentorSession {
  const t = new Date().toISOString()
  return {
    version: 2,
    wbId: o.wbId,
    kind: 'problem',
    mode: o.problem.mode,
    problem: o.problem,
    position: { stage: 'problem', concept: 0 },
    attempts: 0,
    messages: [],
    lessons: {},
    exampleRow: 3,
    practiceRow: 3,
    practice: {},
    ledger: { added: [], rejected: [] },
    pressureTests: {},
    stageLog: [],
    hints: {},
    stats: {},
    checks: [],
    startedAt: t,
    lastActivityAt: t,
    timeSpentMs: 0,
    originalDocKey: o.wbId,
    notice: o.notice,
  }
}

// ------------------------------------------------------------------ positions
export function nextPosition(p: StagePosition, nConcepts: number): StagePosition {
  switch (p.stage) {
    case 'problem':
      return { stage: 'roadmap', concept: 0 }
    case 'roadmap':
      return nConcepts ? { stage: 'learn', concept: 0 } : { stage: 'plan', concept: 0 }
    case 'learn':
      return { stage: 'practice', concept: p.concept }
    case 'practice':
      return p.concept + 1 < nConcepts ? { stage: 'learn', concept: p.concept + 1 } : { stage: 'plan', concept: 0 }
    case 'plan':
      return { stage: 'solve', concept: 0 }
    case 'solve':
      return { stage: 'check', concept: 0 }
    case 'check':
      return { stage: 'followup', concept: 0 }
    default:
      return { stage: 'report', concept: 0 }
  }
}

/** 0-based index in the 9-stage stepper. */
export const stageIndex = (s: StageId) => STAGE_INFO[s].n - 1

export function currentConcept(s: MentorSession): ConceptId | null {
  const st = s.position.stage
  return st === 'learn' || st === 'practice' ? (s.problem?.newConcepts[s.position.concept] ?? null) : null
}

/** Why the current gate can't pass yet, from the app's own evidence (null when nothing is missing). */
export function missingEvidence(s: MentorSession): string | null {
  const c = currentConcept(s)
  if (s.position.stage === 'practice' && c) {
    const pt = s.practice[c]
    if (!pt) return 'the practice exercise has not been written yet.'
    if (!pt.passed) return `the formula engine has not confirmed the practice answer in ${SHEETS.practice}!${pt.cell} as correct yet.`
  }
  if (s.position.stage === 'solve' && !s.checks.length) return 'the student has not run Check My Work yet.'
  if (s.position.stage === 'check') {
    const last = s.checks[s.checks.length - 1]
    if (!last) return 'the student has not run Check My Work yet.'
    if (last.correct < last.total) return `the last Check My Work had ${last.total - last.correct} cell(s) still wrong; every Task cell must be correct.`
  }
  if (s.teachBack) return `the student still owes a teach-back for "${s.teachBack}" after a Level-4 reveal.`
  return null
}

// ------------------------------------------------------------------ task groups (one per result column / summary / operation)
export interface TaskGroup {
  id: string
  requirement: string
  concept: ConceptId
  targets: Target[]
  where: string
}

export function taskGroups(targets: Target[]): TaskGroup[] {
  const groups: TaskGroup[] = []
  for (const t of targets) {
    const key = t.kind === 'cell' && t.id.startsWith('t') ? t.id.split('_')[0] : t.id
    let g = groups.find((x) => x.id === key)
    if (!g) groups.push((g = { id: key, requirement: t.requirement, concept: t.concept, targets: [], where: '' }))
    g.targets.push(t)
  }
  for (const g of groups) {
    const cells = g.targets.filter((t): t is CellTarget => t.kind === 'cell')
    g.where = cells.length ? (cells.length > 1 ? `${cells[0].cell}:${cells[cells.length - 1].cell}` : cells[0].cell) : g.targets[0].kind === 'sort' ? 'the whole table (Sort)' : g.targets[0].kind === 'chart' ? 'a new chart' : 'the column (Format)'
  }
  return groups
}

export const groupOf = (targetId: string) => (targetId.startsWith('t') ? targetId.split('_')[0] : targetId)

// ------------------------------------------------------------------ what the mentor knows about the problem (system prompt)
function sheetTable(doc: WorkbookDoc | null, name: string, maxRows: number): string {
  const s = doc?.sheets.find((x) => x.name === name)
  if (!s) return ''
  const rows = s.cells.slice(0, maxRows + 1)
  const width = Math.max(0, ...rows.map((r) => r?.length ?? 0))
  const lines = rows.map((r, i) => `${i + 1} | ${Array.from({ length: width }, (_, c) => (r?.[c] === null || r?.[c] === undefined ? '' : String(r[c]))).join(' | ')}`)
  const total = s.cells.length
  return `${lines.join('\n')}${total > maxRows + 1 ? `\n… (${total - 1} data rows in total)` : ''}`
}

export function problemBrief(pub: ProblemPublic, targets: Target[], original: WorkbookDoc | null): string {
  const task = original?.sheets.find((s) => s.name === SHEETS.task)
  const headers = (task?.cells[0] ?? []).map((h, c) => `${colToLetters(c)}: ${h ?? ''}`).join(', ')
  const lookup = original?.sheets.find((s) => ![SHEETS.instructions, SHEETS.task, SHEETS.data, SHEETS.example, SHEETS.practice].includes(s.name as never))
  const groups = taskGroups(targets)
  const lines = groups.map((g, i) => {
    const t = g.targets[0]
    const secret = t.kind === 'cell' ? `SECRET expected formula (first cell, never reveal): ${t.formula}` : t.kind === 'sort' ? `expected: sort the table by ${t.columnName} ${t.order === 'asc' ? 'ascending' : 'descending'}` : t.kind === 'chart' ? `expected: a ${t.chartType} chart of ${t.valueName} by column ${t.labelCol}` : `expected: ${t.format} number format on ${t.range}`
    return `Task ${i + 1} [group ${g.id}] — ${g.requirement} — concept: ${conceptName(g.concept)} — where: ${g.where} — ${secret}`
  })
  return [
    `Title: ${pub.title}`,
    `Business context: ${pub.context}`,
    `Objective: ${pub.objective}`,
    `About the data: ${pub.explanation.data}`,
    pub.brief ? `The uploaded assessment brief (mentor summary):\n${pub.brief}` : '',
    pub.uncheckedRequirements?.length ? `Parts of the brief the app can't check automatically (teach and discuss them in conversation, still without doing them for the student):\n${pub.uncheckedRequirements.map((x) => `- ${x}`).join('\n')}` : '',
    `Level: ${pub.level}. Concepts in this problem: ${pub.concepts.map(conceptName).join(', ')}. New concepts to teach: ${pub.newConcepts.map(conceptName).join(', ') || 'none'}.`,
    `Sheets: ${SHEETS.instructions} (the brief), ${SHEETS.task} (the data with yellow result cells the student fills), ${SHEETS.data} (a read-only copy of the data)${lookup ? `, ${lookup.name} (lookup table)` : ''}, ${SHEETS.example} (your worked examples), ${SHEETS.practice} (practice exercises).`,
    `Task sheet columns: ${headers}`,
    `Task sheet (first rows; row 1 is the header):\n${sheetTable(original, SHEETS.task, 8)}`,
    lookup ? `${lookup.name} sheet:\n${sheetTable(original, lookup.name, 10)}` : '',
    `Tasks (the yellow cells):\n${lines.join('\n')}`,
  ]
    .filter(Boolean)
    .join('\n\n')
}

// ------------------------------------------------------------------ the state the mentor gets every turn
export function stateText(s: MentorSession, targets: Target[]): string {
  const pub = s.problem!
  const pos = s.position
  const info = STAGE_INFO[pos.stage]
  const c = currentConcept(s)
  const next = nextPosition(pos, pub.newConcepts.length)
  const lines: string[] = []
  lines.push(`Stage: "${pos.stage}" (stage ${info.n} of 9 — ${info.area})${c ? ` — concept_index ${pos.concept}: ${conceptName(c)} (concept ${pos.concept + 1} of ${pub.newConcepts.length})` : ''}`)
  if (pos.stage === 'report') lines.push('This is the final stage. Set session_complete true only after the student has given their reflection.')
  else lines.push(`If the gate passes, the ONLY allowed next position: "${next.stage}"${next.stage === 'learn' || next.stage === 'practice' ? `, concept_index ${next.concept} (${conceptName(pub.newConcepts[next.concept])})` : ''}${pos.stage === 'solve' ? ' — the app makes this move when the student clicks Check My Work' : ''}.`)
  const miss = missingEvidence(s)
  if (miss && pos.stage !== 'report') lines.push(`Evidence still missing before this gate can pass: ${miss}`)
  lines.push(`Student replies at this gate so far: ${s.attempts}`)
  // the current concept, and the next one (prepared before the reply that moves to it)
  const nextC = (next.stage === 'learn' || next.stage === 'practice') && pos.stage !== 'report' ? pub.newConcepts[next.concept] : null
  for (const k of [...new Set([c, nextC].filter((x): x is ConceptId => !!x))]) {
    const tag = k === c ? '' : ' (prepared for the next concept)'
    const ex = s.exampleCells?.[k]
    if (ex) lines.push(`Worked example for ${conceptName(k)}${tag} on the ${SHEETS.example} sheet: data ${ex.range}, result cell ${ex.cell} = \`${ex.formula}\` → ${ex.value}. (You may show and explain this example formula freely.)`)
    const pt = s.practice[k]
    const lesson = s.lessons[k]
    if (pt) lines.push(`Practice for ${conceptName(k)}${tag}: yellow cell ${SHEETS.practice}!${pt.cell}${lesson ? ` — "${lesson.practice.prompt}"` : ''} — status: ${pt.passed ? 'CORRECT (engine-verified)' : pt.attempts ? `not correct yet (${pt.attempts} check${pt.attempts === 1 ? '' : 's'})` : 'not checked yet'}. SECRET practice formula (Level 4 only): ${pt.formula}`)
  }
  if (c) {
    const used = s.pressureTests[c] ?? []
    if (used.length) lines.push(`Pressure Test logics already used for ${conceptName(c)}: ${used.join(', ')} — use a different one.`)
  } else {
    const used = s.pressureTests[pos.stage] ?? []
    if (used.length) lines.push(`Pressure Test logics already used in this stage: ${used.join(', ')}.`)
  }
  if (s.ledger.added.length) lines.push(`Official solution plan (validated so far): ${s.ledger.added.join('; ')}`)
  if (s.ledger.rejected.length) lines.push(`Rejected assumptions: ${s.ledger.rejected.join('; ')}`)
  const last = s.checks[s.checks.length - 1]
  if (last) {
    const wrong = last.results.filter((r) => r.state !== 'correct')
    lines.push(`Check My Work: run ${s.checks.length} time(s); last result ${last.correct}/${last.total} correct${wrong.length ? `; not correct: ${wrong.slice(0, 12).map((r) => `${targets.find((t) => t.id === r.id)?.kind === 'cell' ? (targets.find((t) => t.id === r.id) as CellTarget).cell : r.id} (${r.state.replace(/_/g, ' ')})`).join(', ')}${wrong.length > 12 ? '…' : ''}` : ''}.`)
  }
  const hinted = Object.entries(s.hints).filter(([, v]) => v > 0)
  if (hinted.length) lines.push(`Hint levels already given: ${hinted.map(([k, v]) => `${k} → level ${v}`).join(', ')}.`)
  if (s.teachBack) lines.push(`Teach-back pending for ${s.teachBack} (after a Level-4 reveal): ask for it and set teach_back_passed true only when the explanation is right.`)
  if (s.stageLog.length) lines.push(`Gates passed so far:\n${s.stageLog.map((r) => `- ${r.stage}${r.concept ? ` (${conceptName(r.concept)})` : ''}: score ${r.score} — ${r.summary}`).join('\n')}`)
  return lines.join('\n')
}

// ------------------------------------------------------------------ applying a validated reply
export function applyTurn(s: MentorSession, t: MentorTurn, now: string): MentorSession {
  const pub = s.problem!
  const n = structuredClone(s)
  const key = currentConcept(s) ?? s.position.stage
  if (t.pressure_test !== 'none') n.pressureTests[key] = [...new Set([...(n.pressureTests[key] ?? []), t.pressure_test])]
  for (const a of t.context_added) if (!n.ledger.added.includes(a)) n.ledger.added.push(a)
  for (const a of t.assumptions_rejected) if (!n.ledger.rejected.includes(a)) n.ledger.rejected.push(a)
  if (t.teach_back_passed) n.teachBack = undefined
  const moved = t.current_stage !== s.position.stage || t.concept_index !== s.position.concept
  if (t.gate_passed && (moved || s.position.stage === 'report')) {
    const concept = currentConcept(s) ?? undefined
    n.stageLog.push({ stage: s.position.stage, concept, passedAt: now, summary: t.stage_summary || '(no summary)', score: t.section_score, attempts: s.attempts })
  }
  if (moved) {
    n.position = { stage: t.current_stage, concept: t.current_stage === 'learn' || t.current_stage === 'practice' ? t.concept_index : 0 }
    n.attempts = 0
  }
  void pub
  return n
}

// ------------------------------------------------------------------ answer leaks
const norm = (s: string) => s.replace(/[\s$`]/g, '').toUpperCase()

export interface LeakGuard {
  targets: CellTarget[]
  expected: Record<string, unknown>
  /** task groups the mentor may reveal right now (topic mode, Level 4 after Level 3) */
  allowGroups: Set<string>
  /** the current practice formula, revealable only at Level 4 */
  practice?: { cell: string; formula: string; level: number }
}

/** The cell whose formula (or, for single summary cells, exact result) appears in the text. */
export function findLeak(text: string, g: LeakGuard, hintLevel: number): string | null {
  const tx = norm(text)
  for (const t of g.targets) {
    if (g.allowGroups.has(groupOf(t.id)) && hintLevel >= MAX_HINT) continue
    const f = norm(t.formula)
    if (f.length > 5 && (tx.includes(f) || tx.includes(f.replace(/^=/, '')))) return `${t.sheet}!${t.cell}`
    const v = g.expected[t.id]
    // results are only matched for single summary cells: row results are too numerous and too often look like data
    if (!t.id.startsWith('t') && typeof v === 'number' && Math.abs(v) >= 100) {
      const forms = [v.toLocaleString('en-US'), String(v), v.toFixed(2), v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })]
      if (forms.some((x) => new RegExp(`(^|[^\\d.,])${x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}($|[^\\d])`).test(text))) return `${t.sheet}!${t.cell}`
    }
  }
  if (g.practice && hintLevel < MAX_HINT && g.practice.level < MAX_HINT - 1) {
    const f = norm(g.practice.formula)
    if (f.length > 5 && (tx.includes(f) || tx.includes(f.replace(/^=/, '')))) return `${SHEETS.practice}!${g.practice.cell}`
  }
  return null
}

// ------------------------------------------------------------------ the student's attempt at a task cell
export function cellContext(conceptId: ConceptId): string {
  const c = CONCEPTS[conceptId]
  return `${c.name}${c.functions.length ? ` (${c.functions.join(', ')})` : ''}`
}
