import { test } from 'node:test'
import assert from 'node:assert/strict'
import { Workbook } from '../src/engine/workbook'
import { builtinProblem } from './helpers/builtins'
import { buildProblem, SHEETS } from '../src/mentor/build'
import { checkWork } from '../src/mentor/validate'
import { translateFormula, functionsIn, referencedCells } from '../src/mentor/layout'
import { computeStage, emptyProgress, recommend, recordPractice, recordProblem, nextLevel, performanceScore } from '../src/mentor/progress'
import { createTopicProblem, builtinLesson, MentorUnavailableError } from '../src/mentor/ai'
import { GeminiClient, parseJsonLoose } from '../src/mentor/gemini'
import { writeExample, writePractice, checkPractice, validateBlock } from '../src/mentor/lesson'
import { MentorStore, memoryKV } from '../src/mentor/store'
import { CONCEPTS, GENERATABLE } from '../src/mentor/concepts'
import type { CellTarget, ConceptId, Domain, Profile, ProblemDraft } from '../src/mentor/types'
import { sortRange } from '../src/engine/features/data'

const DOMAINS: Domain[] = ['Sales', 'HR', 'Finance', 'Marketing', 'Operations', 'Healthcare', 'Education', 'Customer Support', 'General']
const SETS: ConceptId[][] = [['sum'], ['sum', 'average', 'max'], ['min', 'count', 'sort', 'chart'], ['percent', 'if', 'formatting'], ['countif', 'sumif', 'averageif'], ['round', 'concat', 'formula'], ['xlookup'], ['vlookup'], ['index-match']]

const profile: Profile = { level: 'beginner', role: 'Sales', domain: 'Sales', description: '', attachments: [], difficulty: 'normal', duration: 30, createdAt: '' }

test('built-in problems pass every guardrail for all domains and concept sets', () => {
  for (const domain of DOMAINS)
    for (const concepts of SETS) {
      const draft = builtinProblem({ domain, level: 'beginner', concepts, seed: `${domain}${concepts}` })
      const res = buildProblem({ draft, level: 'beginner', domain, newConcepts: concepts, source: 'gemini' })
      assert.deepEqual(res.errors, [], `${domain} ${concepts}`)
      assert.ok(res.built!.targets.length > 0)
    }
})

test('built-in problems are deterministic for a seed', () => {
  const a = builtinProblem({ domain: 'HR', level: 'intermediate', concepts: ['sum', 'if'], seed: 'same' })
  const b = builtinProblem({ domain: 'HR', level: 'intermediate', concepts: ['sum', 'if'], seed: 'same' })
  assert.deepEqual(a.dataset, b.dataset)
})

test('task sheet never contains answers; workbook has the required sheets', () => {
  const draft = builtinProblem({ domain: 'Sales', level: 'beginner', concepts: ['sum', 'max'], seed: 's1' })
  const { built } = buildProblem({ draft, level: 'beginner', domain: 'Sales', newConcepts: ['sum', 'max'], source: 'gemini' })
  const names = built!.doc.sheets.map((s) => s.name)
  for (const n of [SHEETS.instructions, SHEETS.task, SHEETS.data, SHEETS.example, SHEETS.practice]) assert.ok(names.includes(n))
  const task = built!.doc.sheets.find((s) => s.name === SHEETS.task)!
  for (const t of built!.targets) if (t.kind === 'cell') assert.equal(task.cells[t.r]?.[t.c] ?? null, null, `${t.cell} must be empty`)
  assert.ok(!JSON.stringify(task.cells).includes('=SUM'))
})

function loadSolved(seed: string, concepts: ConceptId[]) {
  const draft = builtinProblem({ domain: 'Sales', level: 'beginner', concepts, seed })
  const { built } = buildProblem({ draft, level: 'beginner', domain: 'Sales', newConcepts: concepts, source: 'gemini' })
  const wb = Workbook.fromDoc(structuredClone(built!.doc))
  const sid = wb.sheetByName(SHEETS.task)!.id
  return { wb, sid, built: built! }
}

test('Check My Work: correct, wrong range, wrong method, typed value, not attempted', () => {
  const { wb, sid, built } = loadSolved('cw', ['sum', 'max'])
  const cells = built.targets.filter((t): t is CellTarget => t.kind === 'cell')
  const rowT = cells.filter((t) => t.id.startsWith('t0_'))
  // solve everything correctly
  wb.setInputs(sid, cells.map((t) => ({ r: t.r, c: t.c, input: t.formula })))
  let res = checkWork(wb, built.targets, built.doc)
  assert.ok(res.every((r) => r.state === 'correct'), JSON.stringify(res.filter((r) => r.state !== 'correct')))
  // row 1: range missing the last month
  const t1 = rowT[0]
  const shortRange = t1.formula.replace(/([A-Z]+)(\d+)\)$/, (_m, col: string, row: string) => `${String.fromCharCode(col.charCodeAt(0) - 1)}${row})`)
  wb.setInputs(sid, [{ r: t1.r, c: t1.c, input: shortRange }])
  // row 2: right number with + instead of SUM
  const t2 = rowT[1]
  const v2 = wb.getValue(sid, t2.r, t2.c) as number
  const plus = '=' + referencedCells(t2.formula).ranges[0].split(':').length && `=${[...referencedCells(t2.formula).cells].map((k) => { const [r, c] = k.split('|')[1].split(',').map(Number); return String.fromCharCode(65 + c) + (r + 1) }).join('+')}`
  wb.setInputs(sid, [{ r: t2.r, c: t2.c, input: plus }])
  // row 3: typed number
  const t3 = rowT[2]
  wb.setInputs(sid, [{ r: t3.r, c: t3.c, input: String(wb.getValue(sid, t3.r, t3.c)) }])
  // row 4: blank
  const t4 = rowT[3]
  wb.setInputs(sid, [{ r: t4.r, c: t4.c, input: '' }])
  res = checkWork(wb, built.targets, built.doc)
  const by = (id: string) => res.find((r) => r.id === id)!
  assert.equal(by(t1.id).state, 'incorrect')
  assert.match(by(t1.id).expected, /Missing/)
  assert.match(by(t1.id).message, /included every value/)
  assert.equal(by(t2.id).state, 'correct_value', `${plus} -> ${by(t2.id).message} (${v2})`)
  assert.match(by(t2.id).message, /practise SUM/)
  assert.equal(by(t3.id).state, 'correct_value')
  assert.equal(by(t4.id).state, 'not_attempted')
  // MAX depends on the totals: its formula is right but the value differs -> correct_formula
  const maxT = cells.find((t) => t.concept === 'max')!
  assert.ok(['correct', 'correct_formula'].includes(by(maxT.id).state))
})

test('sort target: sorted with rows intact passes, unsorted fails', () => {
  const { wb, sid, built } = loadSolved('sort', ['sum', 'sort'])
  const cells = built.targets.filter((t): t is CellTarget => t.kind === 'cell')
  wb.setInputs(sid, cells.map((t) => ({ r: t.r, c: t.c, input: t.formula })))
  const sortT = built.targets.find((t) => t.kind === 'sort')!
  let res = checkWork(wb, built.targets, built.doc)
  assert.equal(res.find((r) => r.id === sortT.id)!.state, 'incorrect')
  const totalCol = cells[0].c
  const rows = built.doc.sheets.find((s) => s.name === SHEETS.task)!.cells.length
  void rows
  sortRange(wb, sid, { r1: 0, c1: 0, r2: cells.filter((t) => t.id.startsWith('t0_')).length, c2: totalCol }, [{ col: totalCol, desc: true }], true)
  res = checkWork(wb, built.targets, built.doc)
  assert.equal(res.find((r) => r.id === sortT.id)!.state, 'correct', JSON.stringify(res.find((r) => r.id === sortT.id)))
  assert.ok(res.filter((r) => r.id !== sortT.id && r.id.startsWith('t')).every((r) => r.state === 'correct'))
})

test('chart and format targets', () => {
  const { wb, sid, built } = loadSolved('chart', ['sum', 'chart', 'formatting'])
  const cells = built.targets.filter((t): t is CellTarget => t.kind === 'cell')
  wb.setInputs(sid, cells.map((t) => ({ r: t.r, c: t.c, input: t.formula })))
  let res = checkWork(wb, built.targets, built.doc)
  const chartT = built.targets.find((t) => t.kind === 'chart')!
  const fmtT = built.targets.find((t) => t.kind === 'format')!
  assert.equal(res.find((r) => r.id === chartT.id)!.state, 'not_attempted')
  assert.notEqual(res.find((r) => r.id === fmtT.id)!.state, 'correct')
  const totalCol = cells[0].c
  wb.run({ label: 'chart', dirty: [sid], metaScope: [sid], metaOps: () => (wb.meta(sid).charts = [{ id: 'c1', type: 'column', source: { sheetId: sid, range: { r1: 0, c1: 0, r2: 8, c2: totalCol } }, seriesIn: 'cols', firstRowHeader: true, firstColLabels: true, title: '', xTitle: '', yTitle: '', legend: 'none', dataLabels: false, gridlines: true, stacked: false, colors: [], x: 0, y: 0, w: 400, h: 300 }]) })
  wb.applyStyle(sid, [{ r1: 1, c1: totalCol, r2: 8, c2: totalCol }], { numFmt: '$#,##0' })
  res = checkWork(wb, built.targets, built.doc)
  assert.equal(res.find((r) => r.id === chartT.id)!.state, 'correct')
  assert.equal(res.find((r) => r.id === fmtT.id)!.state, 'correct')
})

test('formula templates translate to A1 references', () => {
  const ctx = { main: { sheet: 'Task', headerRow: 0, firstRow: 1, lastRow: 10, cols: new Map([['rep', 0], ['jan', 1], ['apr', 4], ['total', 5]]) }, fromSheet: 'Task', others: new Map([['lookup', { sheet: 'Lookup', headerRow: 0, firstRow: 1, lastRow: 4, cols: new Map([['region', 0], ['rate', 1]]) }]]), summaries: new Map([['grand total', 'F13']]) }
  assert.deepEqual(translateFormula('=SUM([Jan]:[Apr])', ctx, 3), { formula: '=SUM(B4:E4)' })
  assert.deepEqual(translateFormula('=MAX([Total:all])', ctx), { formula: '=MAX(F2:F11)' })
  assert.deepEqual(translateFormula('=XLOOKUP([Rep],[Lookup.Region:all],[Lookup.Rate:all])', ctx, 1), { formula: '=XLOOKUP(A2,Lookup!A2:A5,Lookup!B2:B5)' })
  assert.deepEqual(translateFormula('=VLOOKUP("x",[Lookup.Region:all]:[Lookup.Rate:all],2,FALSE)', ctx), { formula: '=VLOOKUP("x",Lookup!$A$2:$B$5,2,FALSE)' })
  assert.deepEqual(translateFormula('=[[Grand total]]/2', ctx), { formula: '=F13/2' })
  assert.deepEqual(translateFormula('=COUNTIF([Rep:all],"[Jan]")', ctx), { formula: '=COUNTIF(A2:A11,"[Jan]")' })
  assert.ok('error' in translateFormula('=SUM([Nope:all])', ctx))
  assert.ok('error' in translateFormula('=[Jan]*2', ctx))
  assert.deepEqual(functionsIn('=IF(SUM(A1:A2)>1,"MAX(","x")'), ['IF', 'SUM'])
})

test('mastery needs evidence across problems', () => {
  const p = emptyProgress()
  p.stage = 'introduced'
  assert.equal(computeStage(p), 'introduced')
  let s = recordPractice({ concepts: {}, history: [] }, 'sum', true)
  assert.equal(s.concepts.sum!.stage, 'practicing')
  const draft = builtinProblem({ domain: 'Sales', level: 'beginner', concepts: ['sum'], seed: 'm' })
  const { built } = buildProblem({ draft, level: 'beginner', domain: 'Sales', newConcepts: ['sum'], source: 'gemini' })
  const pub = built!.pub
  const allCorrect = { at: '', correct: pub.targets.length, total: pub.targets.length, results: pub.targets.map((t) => ({ id: t.id, state: 'correct' as const, student: '', expected: '', message: '' })) }
  s = recordProblem(s, pub, allCorrect, {})
  assert.equal(s.concepts.sum!.stage, 'strong', 'one clean problem = strong, not mastered')
  s = recordProblem(s, pub, allCorrect, {})
  assert.equal(s.concepts.sum!.stage, 'mastered')
  const perf = performanceScore(pub, [allCorrect], {})
  assert.equal(nextLevel('beginner', perf), 'intermediate')
  assert.equal(nextLevel('intermediate', { firstCheck: 0.3, final: 1, hintRate: 3, assisted: true }), 'beginner')
})

test('recommendations follow the dependency order', () => {
  const r = recommend({ ...profile, role: 'Other' }, { concepts: {}, history: [] })
  assert.deepEqual(r.newConcepts, ['sum', 'max'], 'AVERAGE waits until SUM is known')
  assert.deepEqual(recommend(profile, { concepts: {}, history: [] }).newConcepts, ['sum', 'sort'], 'a Sales role gets Sort early')
  assert.equal(r.level, 'beginner')
  const known = { stage: 'strong' as const, understandingScore: 1, understandingCount: 1, practiceCorrect: 1, practiceAttempts: 1, appliedCorrect: 10, appliedTotal: 10, hintLevels: 0, independentProblems: 1, problems: 1 }
  const r2 = recommend(profile, { concepts: { sum: known, average: known, max: known, min: known, count: known }, history: [] })
  assert.ok(!r2.newConcepts.includes('sum'))
  assert.ok(!r2.newConcepts.includes('if'), 'IF is intermediate and needs the beginner level raised')
})

test('Gemini problem: guardrail failures are fed back, then the fixed problem is used', async () => {
  const good: ProblemDraft = builtinProblem({ domain: 'Sales', level: 'beginner', concepts: ['sum', 'average'], seed: 'ai' })
  const bad = structuredClone(good)
  bad.rowTasks[0].formula = '=SUM([January]:[Apr])' // unknown column
  const asStrings = (d: ProblemDraft) => ({ ...d, dataset: { ...d.dataset, rows: d.dataset.rows.map((r) => r.map(String)) } })
  const replies = [asStrings(bad), asStrings(good)]
  const bodies: unknown[] = []
  const fetcher = (async (_url: string, init: RequestInit) => {
    bodies.push(JSON.parse(String(init.body)))
    const draft = replies.shift()
    return new Response(JSON.stringify({ candidates: [{ content: { parts: [{ text: '```json\n' + JSON.stringify(draft) + '\n```' }] } }] }), { status: 200 })
  }) as unknown as typeof fetch
  const client = new GeminiClient({ connection: 'key', apiKey: 'k', proxyUrl: '', model: 'gemini-test' }, fetcher)
  const steps: string[] = []
  const res = await createTopicProblem(client, profile, { concepts: {}, history: [] }, { level: 'beginner', newConcepts: ['sum', 'average'], revision: [], estimatedMinutes: 20, reason: '' }, (s) => steps.push(s))
  assert.equal(res.built.pub.source, 'gemini')
  assert.equal(res.built.pub.mode, 'topic')
  assert.equal(bodies.length, 2)
  const second = bodies[1] as { contents: { role: string; parts: { text: string }[] }[]; generationConfig: Record<string, unknown> }
  assert.match(second.contents[2].parts[0].text, /Unknown column “January”/)
  assert.equal(second.contents[1].role, 'model')
  assert.ok(second.generationConfig.responseMimeType === 'application/json')
  assert.ok(!('temperature' in second.generationConfig), 'sampling parameters are not sent')
})

test('every session needs Gemini: no key, or a rejected key, is an error (no canned problems)', async () => {
  await assert.rejects(createTopicProblem(null, profile, { concepts: {}, history: [] }, { level: 'beginner', newConcepts: ['sum'], revision: [], estimatedMinutes: 20, reason: '' }), MentorUnavailableError)
  const fetcher = (async () => new Response(JSON.stringify({ error: { message: 'API key not valid' } }), { status: 400 })) as unknown as typeof fetch
  const client = new GeminiClient({ connection: 'key', apiKey: 'bad', proxyUrl: '', model: 'm' }, fetcher)
  await assert.rejects(createTopicProblem(client, profile, { concepts: {}, history: [] }, { level: 'beginner', newConcepts: ['sum'], revision: [], estimatedMinutes: 20, reason: '' }), /API key/)
})

test('lessons: examples go to their own sheet, practice is checked', () => {
  const draft = builtinProblem({ domain: 'Sales', level: 'beginner', concepts: ['sum'], seed: 'l' })
  const { built } = buildProblem({ draft, level: 'beginner', domain: 'Sales', newConcepts: ['sum'], source: 'gemini' })
  const wb = Workbook.fromDoc(structuredClone(built!.doc))
  const taskBefore = JSON.stringify(wb.serializeSheet(wb.sheetByName(SHEETS.task)!.id).cells)
  const lesson = builtinLesson('sum')
  const ex = writeExample(wb, lesson, 3)
  assert.ok(!('error' in ex))
  const exId = wb.sheetByName(SHEETS.example)!.id
  if (!('error' in ex)) {
    assert.equal(wb.getValue(exId, ex.cell.r, ex.cell.c), 570)
    assert.match(wb.getEditText(exId, ex.cell.r, ex.cell.c), /^=SUM\(B6:B8\)$/)
  }
  const pr = writePractice(wb, lesson, 3)
  assert.ok(!('error' in pr))
  if ('error' in pr) return
  const pid = wb.sheetByName(SHEETS.practice)!.id
  assert.equal(checkPractice(wb, pr.target).state, 'not_attempted')
  wb.setInputs(pid, [{ r: pr.target.r, c: pr.target.c, input: '=10+20+30' }])
  assert.equal(checkPractice(wb, pr.target).state, 'correct_value')
  wb.setInputs(pid, [{ r: pr.target.r, c: pr.target.c, input: pr.target.formula }])
  assert.equal(checkPractice(wb, pr.target).ok, true)
  assert.equal(JSON.stringify(wb.serializeSheet(wb.sheetByName(SHEETS.task)!.id).cells), taskBefore, 'Task sheet untouched')
})

test('every built-in lesson example and practice is valid', () => {
  for (const id of Object.keys(CONCEPTS) as ConceptId[]) {
    const l = builtinLesson(id)
    assert.equal(validateBlock(l.example, id), null, `${id} example`)
    assert.equal(validateBlock(l.practice, id), null, `${id} practice`)
  }
  assert.ok(GENERATABLE.length > 10)
})

test('loose JSON parsing', () => {
  assert.deepEqual(parseJsonLoose('Here you go:\n```json\n{"a":1}\n```'), { a: 1 })
  assert.deepEqual(parseJsonLoose('prefix {"a":[1,2]} suffix'), { a: [1, 2] })
})

test('answers are stored apart from the workbook and not in plain text', () => {
  const kv = memoryKV()
  const st = new MentorStore(kv)
  const draft = builtinProblem({ domain: 'Sales', level: 'beginner', concepts: ['sum'], seed: 'st' })
  const { built } = buildProblem({ draft, level: 'beginner', domain: 'Sales', newConcepts: ['sum'], source: 'gemini' })
  st.saveAnswers(built!.pub.problemId, built!.targets)
  const raw = kv.get(`excelmentor.answers.${built!.pub.problemId}`)!
  assert.ok(!raw.includes('SUM('))
  assert.deepEqual(st.answers(built!.pub.problemId), built!.targets)
})

