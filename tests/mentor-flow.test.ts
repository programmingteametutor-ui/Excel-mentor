/* The Gemini-driven 9-stage flow: protocol checks, gates and evidence, state, answer leaks, file reading, take-home builds. */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import { Workbook } from '../src/engine/workbook'
import { buildProblem, computeExpected, SHEETS } from '../src/mentor/build'
import { createAssessmentProblem, mentorTurn } from '../src/mentor/ai'
import { docxText, spreadsheet, tableFromSheet } from '../src/mentor/extract'
import { applyTurn, findLeak, missingEvidence, newSession, nextPosition, problemBrief, stateText } from '../src/mentor/flow'
import { GeminiClient } from '../src/mentor/gemini'
import { checkTurn, chunkParagraphs, CLOSING_PHRASE, normalizeTurn, sentences, systemPrompt, type TurnCheckContext } from '../src/mentor/protocol'
import { buildReport, recommend, topicRecommendation } from '../src/mentor/progress'
import type { CellTarget, MentorTurn, Profile } from '../src/mentor/types'
import { builtinProblem } from './helpers/builtins'

const profile: Profile = { level: 'beginner', role: 'Finance', domain: 'Finance', description: '', attachments: [], difficulty: 'normal', duration: 30, createdAt: '' }

function built(concepts: ('sum' | 'max' | 'average')[] = ['sum', 'max']) {
  const draft = builtinProblem({ domain: 'Sales', level: 'beginner', concepts, seed: 'flow' })
  const res = buildProblem({ draft, level: 'beginner', domain: 'Sales', newConcepts: concepts, source: 'gemini' })
  assert.deepEqual(res.errors, [])
  return res.built!
}

const turn = (x: Partial<MentorTurn>): MentorTurn => normalizeTurn({ message_to_student: 'Para one.\n\nPara two.\n\nWhat would you ask first?', verdict: 'none', ...x }, 'problem', 0)

function ctxFor(stage: MentorTurn['current_stage'], concept = 0, over: Partial<TurnCheckContext> = {}): TurnCheckContext {
  return { stage, concept, next: nextPosition({ stage, concept }, 2), missingEvidence: null, leaks: () => null, closingPhrase: CLOSING_PHRASE.topic, ...over }
}

// ------------------------------------------------------------------ positions and gates
test('the nine stages run in order, with learn/practice repeated per concept', () => {
  const seq: string[] = []
  let p = { stage: 'problem' as MentorTurn['current_stage'], concept: 0 }
  for (let i = 0; i < 12 && p.stage !== 'report'; i++) {
    seq.push(p.stage === 'learn' || p.stage === 'practice' ? `${p.stage}${p.concept}` : p.stage)
    p = nextPosition(p, 2)
  }
  seq.push(p.stage)
  assert.deepEqual(seq, ['problem', 'roadmap', 'learn0', 'practice0', 'learn1', 'practice1', 'plan', 'solve', 'check', 'followup', 'report'])
  assert.deepEqual(nextPosition({ stage: 'roadmap', concept: 0 }, 0), { stage: 'plan', concept: 0 }, 'a revision problem skips learn/practice')
})

test('State Lockdown: skipping, going back, or moving without a passed gate is rejected', () => {
  assert.deepEqual(checkTurn(turn({ current_stage: 'problem' }), ctxFor('problem')), [])
  const skip = checkTurn(turn({ current_stage: 'learn', gate_passed: true, verdict: 'pass' }), ctxFor('problem'))
  assert.ok(skip.some((p) => p.hard && /State Lockdown/.test(p.rule)))
  const noGate = checkTurn(turn({ current_stage: 'roadmap', verdict: 'partial' }), ctxFor('problem'))
  assert.ok(noGate.some((p) => p.hard && /Wait for Signal/.test(p.rule)))
  const weak = checkTurn(turn({ current_stage: 'roadmap', gate_passed: true, verdict: 'partial' }), ctxFor('problem'))
  assert.ok(weak.some((p) => p.hard && /Three-Pillar/.test(p.rule)))
  assert.deepEqual(checkTurn(turn({ current_stage: 'roadmap', gate_passed: true, verdict: 'pass' }), ctxFor('problem')), [])
  const wrongConcept = checkTurn(turn({ current_stage: 'learn', concept_index: 1, gate_passed: true, verdict: 'pass' }), ctxFor('roadmap'))
  assert.ok(wrongConcept.some((p) => p.hard))
  const back = checkTurn(turn({ current_stage: 'problem', gate_passed: true, verdict: 'pass' }), ctxFor('plan'))
  assert.ok(back.some((p) => p.hard))
})

test('Gatekeeper: practice needs the engine, check needs every cell correct, solve → check is the app’s move', () => {
  const evid = checkTurn(turn({ current_stage: 'learn', concept_index: 1, gate_passed: true, verdict: 'pass' }), ctxFor('practice', 0, { missingEvidence: 'the formula engine has not confirmed the practice answer.' }))
  assert.ok(evid.some((p) => p.hard && /Gatekeeper/.test(p.rule)))
  const solve = checkTurn(turn({ current_stage: 'check', gate_passed: true, verdict: 'pass' }), ctxFor('solve', 0, { appDriven: true }))
  assert.ok(solve.some((p) => p.hard && /Check My Work/.test(p.rule)))

  const b = built()
  const s = newSession({ wbId: b.doc.id, problem: b.pub })
  s.position = { stage: 'practice', concept: 0 }
  s.practice.sum = { concept: 'sum', cell: 'D6', r: 5, c: 3, formula: '=SUM(B4:B5)', requireFunctions: [], attempts: 1, passed: false }
  assert.match(missingEvidence(s)!, /practice answer/)
  s.practice.sum.passed = true
  assert.equal(missingEvidence(s), null)
  s.position = { stage: 'check', concept: 0 }
  s.checks = [{ at: '', results: [], correct: 3, total: 5 }]
  assert.match(missingEvidence(s)!, /2 cell\(s\) still wrong/)
  s.checks.push({ at: '', results: [], correct: 5, total: 5 })
  assert.equal(missingEvidence(s), null)
  s.teachBack = 't0'
  assert.match(missingEvidence(s)!, /teach-back/)
})

test('soft rules: banned templates, surface questions and praise without substance', () => {
  const t = turn({ message_to_student: "Imagine you're tasked with the report. Does this make sense? Excellent start!", verdict: 'partial' })
  const rules = checkTurn(t, ctxFor('learn')).map((p) => p.rule)
  assert.ok(rules.some((r) => /Banned template/.test(r)))
  assert.ok(rules.some((r) => /Anti-Surface/.test(r)))
  assert.ok(rules.some((r) => /Praise-to-Substance/.test(r)))
  assert.ok(!checkTurn(turn({ message_to_student: 'Excellent — that covers all three pillars. Why would it break with blanks?', verdict: 'pass' }), ctxFor('learn')).some((p) => /Praise/.test(p.rule)))
})

test('the final hand-off needs the closing sentence and a 10–15 sentence mastery profile', () => {
  const profile12 = Array.from({ length: 12 }, (_, i) => `Sentence number ${i + 1} about the student.`).join(' ')
  assert.equal(sentences(profile12), 12)
  const ok = turn({ current_stage: 'report', gate_passed: true, verdict: 'pass', session_complete: true, mastery_profile: profile12, message_to_student: `Well done today.\n\n${CLOSING_PHRASE.topic}` })
  assert.deepEqual(checkTurn(ok, ctxFor('report')), [])
  const short = checkTurn({ ...ok, mastery_profile: 'Too short. Only two.' }, ctxFor('report'))
  assert.ok(short.some((p) => /10–15 sentences/.test(p.rule)))
  const noPhrase = checkTurn({ ...ok, message_to_student: 'Bye.' }, ctxFor('report'))
  assert.ok(noPhrase.some((p) => /exact sentence/.test(p.rule)))
  assert.ok(checkTurn(ok, ctxFor('followup')).some((p) => /session_complete/.test(p.rule)))
})

test('Chunking: prose paragraphs over 3 sentences are split; lists and code are kept', () => {
  const out = chunkParagraphs('One. Two. Three. Four. Five.\n\n--> a. b. c. d.\n\n```\nx. y. z. w.\n```')
  assert.equal(out.split('\n\n')[0], 'One. Two. Three.')
  assert.equal(out.split('\n\n')[1], 'Four. Five.')
  assert.ok(out.includes('--> a. b. c. d.'))
  assert.ok(out.includes('x. y. z. w.'))
})

// ------------------------------------------------------------------ answer leaks
test('No Spoiling: task formulas and summary results are detected; Level 4 may reveal the authorised task only', () => {
  const b = built()
  const cells = b.targets.filter((t): t is CellTarget => t.kind === 'cell')
  const expected = computeExpected(b.doc, cells)
  const row = cells.find((t) => t.id.startsWith('t0_'))!
  const summary = cells.find((t) => t.id.startsWith('s'))!
  const guard = { targets: cells, expected, allowGroups: new Set<string>() }
  assert.equal(findLeak(`Try typing ${row.formula} in the cell.`, guard, 1), `${SHEETS.task}!${row.cell}`)
  assert.equal(findLeak('Think about which months belong in the total.', guard, 1), null)
  const v = expected[summary.id] as number
  if (Math.abs(v) >= 100) assert.equal(findLeak(`The top value is ${v.toLocaleString('en-US')}.`, guard, 2), `${SHEETS.task}!${summary.cell}`)
  const allowed = { ...guard, allowGroups: new Set(['t0']) }
  assert.equal(findLeak(`Here it is: ${row.formula}`, allowed, 4), null, 'authorised at level 4')
  assert.notEqual(findLeak(`Here it is: ${row.formula}`, allowed, 3), null, 'not below level 4')
  const practice = { ...guard, practice: { cell: 'D6', formula: '=SUM(B4:B5)', level: 0 } }
  assert.match(findLeak('Use =SUM(B4:B5) here', practice, 1)!, /Practice!D6/)
  assert.equal(findLeak('Use =SUM(B4:B5) here', practice, 4), null)
})

// ------------------------------------------------------------------ session state
test('applying a reply: the gate log, the solution plan ledger and pressure tests are recorded', () => {
  const b = built()
  let s = newSession({ wbId: b.doc.id, problem: b.pub })
  s = applyTurn(s, turn({ current_stage: 'problem', context_added: ['Totals per rep'], assumptions_rejected: ['Typing totals by hand'], pressure_test: 'A' }), 't')
  assert.deepEqual(s.ledger, { added: ['Totals per rep'], rejected: ['Typing totals by hand'] })
  assert.deepEqual(s.pressureTests.problem, ['A'])
  s.attempts = 3
  s = applyTurn(s, turn({ current_stage: 'roadmap', gate_passed: true, verdict: 'pass', stage_summary: 'Asked about refresh and scope.', section_score: 80 }), 't2')
  assert.deepEqual(s.position, { stage: 'roadmap', concept: 0 })
  assert.equal(s.attempts, 0)
  assert.deepEqual(s.stageLog, [{ stage: 'problem', concept: undefined, passedAt: 't2', summary: 'Asked about refresh and scope.', score: 80, attempts: 3 }])
  const st = stateText(s, b.targets)
  assert.match(st, /Stage: "roadmap"/)
  assert.match(st, /ONLY allowed next position: "learn", concept_index 0/)
  assert.match(st, /Official solution plan.*Totals per rep/)
  const brief = problemBrief(b.pub, b.targets, b.doc)
  assert.match(brief, /SECRET expected formula/)
  const sys = systemPrompt({ profile, problem: b.pub, problemBrief: brief })
  for (const rule of ['ATOMIC TEACHING', 'CONTEXT VALIDATION', 'PRESSURE TEST', 'HUMAN-FIRST', 'BUILD-IN-THE-SHEET', 'RIGOROUS GATEKEEPER', 'ANTI-PARROT', 'TEACH-THEN-GUIDE', 'ADAPTIVE PIVOT', 'MASTERY PROFILE', 'READABILITY']) assert.ok(sys.includes(rule), rule)
  assert.ok(sys.includes(CLOSING_PHRASE.topic))
})

test('the report scores each of the nine stages from the gates passed', () => {
  const b = built()
  const next = recommend(profile, { concepts: {}, history: [] })
  const r = buildReport(b.pub, [], {}, 600000, next, [
    { stage: 'problem', passedAt: '', summary: 'Good questions.', score: 90, attempts: 1 },
    { stage: 'learn', concept: 'sum', passedAt: '', summary: 'Needed nudges.', score: 40, attempts: 4 },
  ], 'Profile.')
  assert.equal(r.sections.length, 9)
  assert.equal(r.sections[0].score, 90)
  assert.equal(r.sections[3].score, null)
  assert.ok(r.needsPractice.includes('SUM'))
  assert.equal(r.masteryProfile, 'Profile.')
})

test('topic choice: chosen topics in curriculum order (max 3); the mentor’s choice follows the role', () => {
  const r = topicRecommendation(profile, { concepts: {}, history: [] }, ['xlookup', 'sum', 'if', 'countif'])
  assert.deepEqual(r.newConcepts, ['sum', 'if', 'countif'])
  const hr = recommend({ ...profile, role: 'HR' }, { concepts: {}, history: [] })
  assert.equal(hr.newConcepts[0], 'count', 'HR starts with counting')
})

// ------------------------------------------------------------------ one mentor turn with Gemini
function mockClient(replies: unknown[], bodies: unknown[] = []) {
  const fetcher = (async (_u: string, init: RequestInit) => {
    bodies.push(JSON.parse(String(init.body)))
    return new Response(JSON.stringify({ candidates: [{ content: { parts: [{ text: JSON.stringify(replies.shift()) }] } }] }), { status: 200 })
  }) as unknown as typeof fetch
  return new GeminiClient({ connection: 'key', apiKey: 'k', proxyUrl: '', model: 'm' }, fetcher)
}

test('mentor turn: a reply that breaks a rule is sent back once with the rules it broke', async () => {
  const bodies: { contents: { role: string; parts: { text: string }[] }[] }[] = []
  const bad = { ...turn({ current_stage: 'plan', gate_passed: true, verdict: 'pass' }) }
  const good = { ...turn({ current_stage: 'problem', verdict: 'partial', hint_level: 1 }) }
  const client = mockClient([bad, good], bodies)
  const res = await mentorTurn({ client, system: 'sys', history: [{ id: '1', from: 'mentor', text: 'Hello', at: '' }, { id: '2', from: 'student', text: 'Hi', at: '' }], state: 'Stage: "problem"', event: 'Student says: "ok"', stage: 'problem', concept: 0, check: { next: { stage: 'roadmap', concept: 0 }, missingEvidence: null, leaks: () => null, closingPhrase: CLOSING_PHRASE.topic } })
  assert.equal(res.corrected, true)
  assert.deepEqual(res.problems, [])
  assert.equal(res.turn.current_stage, 'problem')
  assert.equal(bodies.length, 2)
  const last = bodies[1].contents[bodies[1].contents.length - 1].parts[0].text
  assert.match(last, /State Lockdown/)
  assert.equal(bodies[0].contents[0].role, 'user', 'history starts with the user')
})

// ------------------------------------------------------------------ files
test('Word documents: paragraphs, numbered items and tables become text', async () => {
  const zip = new JSZip()
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>
    <w:p><w:r><w:t>Sales assessment</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Calculate total revenue</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Target</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>5000</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
  </w:body></w:document>`)
  const text = await docxText(await zip.generateAsync({ type: 'uint8array' }))
  assert.match(text, /Sales assessment/)
  assert.match(text, /• Calculate total revenue/)
  assert.match(text, /Region \| Target\nNorth \| 5000/)
})

function assessmentWorkbook() {
  const wb = Workbook.blank('Assessment', ['Orders'])
  const sid = wb.order[0]
  const rows = [['Q3 order export'], [], ['Order', 'Rep', 'Date', 'Amount'], ['A1', 'Ann', '2024-07-01', '120'], ['A2', 'Bob', '2024-07-03', '80.5'], ['A3', 'Ann', '2024-07-09', ''], ['A4', 'Cy', '2024-07-12', '200']]
  wb.setInputs(sid, rows.flatMap((r, i) => r.map((v, c) => ({ r: i, c, input: v }))))
  return wb.toDoc()
}

test('uploaded sheets: the table under the header row is read with column types, blanks kept', () => {
  const t = tableFromSheet(assessmentWorkbook(), 'Orders', 2)
  assert.ok(!('error' in t))
  if ('error' in t) return
  assert.deepEqual(t.columns.map((c) => `${c.name}:${c.type}`), ['Order:text', 'Rep:text', 'Date:date', 'Amount:number'])
  assert.equal(t.rows.length, 4)
  assert.equal(t.rows[0][2], '2024-07-01')
  assert.equal(t.rows[2][3], '')
  assert.ok('error' in tableFromSheet(assessmentWorkbook(), 'Orders', 0), 'a title row is not a header')
  assert.ok('error' in tableFromSheet(assessmentWorkbook(), 'Nope', 2))
})

test('take-home: one Gemini call; a wrong header row is found automatically; unusable tasks move to conversation instead of blocking', async () => {
  const file = spreadsheet('orders.xlsx', assessmentWorkbook())
  assert.equal(file.sheets![0].name, 'Orders')
  const brief = { name: 'brief.docx', kind: 'document' as const, text: 'Q1. Total amount. Q2. Largest order. Q3. Build a PivotTable by rep.' }
  const base = {
    title: 'Q3 orders',
    context: 'An employer sent Q3 orders.',
    objective: 'Summarise the orders.',
    explanation: { situation: 's', goal: 'g', data: 'd', think: 't' },
    brief_summary: ['Total amount', 'Largest order', 'PivotTable by rep'],
    level: 'beginner',
    concepts: ['sum', 'max'],
    estimatedMinutes: 30,
    rowTasks: [],
    summaryTasks: [
      { label: 'Total amount', requirement: 'Q1: total of all order amounts', concept: 'sum', formula: '=SUM([Amount:all])', underColumn: 'Amount' },
      { label: 'Largest order', requirement: 'Q2: the largest order amount', concept: 'max', formula: '=MAX([Amount:all])', underColumn: 'Amount' },
      { label: 'Profit', requirement: 'Q4: total profit', concept: 'sum', formula: '=SUM([Profit:all])', underColumn: 'Amount' },
      { label: 'Odd', requirement: 'Q5: something exotic', concept: 'pivot', formula: '=FOOBAR([Amount:all])' },
    ],
    operations: [],
    instructions: ['Find the total amount', 'Find the largest order'],
    unchecked_requirements: ['Q3: PivotTable of amount by rep'],
  }
  const bodies: { contents: { parts: { text: string }[] }[] }[] = []
  const client = mockClient([{ ...base, main_table: { file: 'orders.xlsx', sheet: 'Orders', header_row: 1 } }], bodies)
  const res = await createAssessmentProblem(client, profile, { concepts: {}, history: [] }, [file, brief], 'Applying for a junior analyst job')
  const b = res.built
  assert.equal(b.pub.mode, 'takehome')
  assert.equal(b.pub.source, 'assessment')
  assert.deepEqual(b.pub.uncheckedRequirements, ['Q3: PivotTable of amount by rep', 'Q4: total profit', 'Q5: something exotic'])
  assert.equal(bodies.length, 1, 'no re-asking Gemini')
  assert.equal(b.targets.length, 2)
  assert.match(b.pub.brief!, /Largest order/)
  assert.deepEqual(b.pub.newConcepts, ['sum', 'max'])
  const values = Object.values(b.expected)
  assert.ok(values.includes(400.5) && values.includes(200), JSON.stringify(b.expected))
  const task = b.doc.sheets.find((s) => s.name === SHEETS.task)!
  assert.deepEqual(task.cells[0].slice(0, 4), ['Order', 'Rep', 'Date', 'Amount'])
  assert.match(bodies[0].contents[0].parts[0].text, /Applying for a junior analyst job/)
  const ins = JSON.stringify(b.doc.sheets.find((s) => s.name === SHEETS.instructions)!.cells)
  assert.match(ins, /PivotTable/)
  const sys = systemPrompt({ profile, problem: b.pub, problemBrief: problemBrief(b.pub, b.targets, b.doc) })
  assert.match(sys, /real take-home assessment/)
  assert.ok(sys.includes(CLOSING_PHRASE.takehome))
})

test('settings default to the API key mode; model picker prefers the newest general flash model', async () => {
  const { MentorStore, memoryKV } = await import('../src/mentor/store')
  const kv = memoryKV()
  assert.equal(new MentorStore(kv).settings().connection, 'key')
  kv.set('excelmentor.settings', JSON.stringify({ connection: 'none', apiKey: 'k', proxyUrl: '', model: '' }))
  assert.equal(new MentorStore(kv).settings().connection, 'key', 'the old built-in mode maps to the key mode')
  const { pickModel } = await import('../src/mentor/gemini')
  assert.equal(pickModel(['gemini-2.0-flash', 'gemini-2.5-flash', 'gemini-2.5-flash-lite', 'gemini-2.5-pro', 'gemini-2.5-flash-image', 'text-embedding-004']), 'gemini-2.5-flash')
})

test('Gemini errors include Google’s own reason', async () => {
  const fetcher = (async () => new Response(JSON.stringify({ error: { message: 'Requests from referer http://localhost:5173/ are blocked.' } }), { status: 403 })) as unknown as typeof fetch
  const c = new GeminiClient({ connection: 'key', apiKey: 'k', proxyUrl: '', model: 'm' }, fetcher)
  await assert.rejects(c.generate({ system: 's', turns: [{ role: 'user', text: 'hi' }] }), /restricted.*Google says: Requests from referer/s)
})

test('an overloaded model (503) is retried, then another model is tried', async () => {
  let calls = 0
  const urls: string[] = []
  const fetcher = (async (url: string) => {
    urls.push(url)
    if (url.includes('models?')) return new Response(JSON.stringify({ models: [{ name: 'models/gemini-busy', supportedGenerationMethods: ['generateContent'] }, { name: 'models/gemini-2.5-flash', supportedGenerationMethods: ['generateContent'] }] }), { status: 200 })
    calls++
    if (url.includes('gemini-busy')) return new Response(JSON.stringify({ error: { message: 'This model is currently experiencing high demand.' } }), { status: 503 })
    return new Response(JSON.stringify({ candidates: [{ content: { parts: [{ text: 'OK' }] } }] }), { status: 200 })
  }) as unknown as typeof fetch
  const c = new GeminiClient({ connection: 'key', apiKey: 'k', proxyUrl: '', model: 'gemini-busy' }, fetcher, [0, 0])
  assert.equal(await c.generate({ system: 's', turns: [{ role: 'user', text: 'hi' }] }), 'OK')
  assert.equal(c.model, 'gemini-2.5-flash')
  assert.equal(calls, 4, '3 tries on the busy model, then the stand-in')
  const p = new GeminiClient({ connection: 'proxy', apiKey: '', proxyUrl: 'https://p', model: 'x-busy' }, (async () => new Response(JSON.stringify({ error: { message: 'high demand' } }), { status: 503 })) as unknown as typeof fetch, [0])
  await assert.rejects(p.generate({ system: 's', turns: [{ role: 'user', text: 'hi' }] }), /overloaded right now/)
})

test('when the newest models are all overloaded, a stable alias is tried before them and the stand-in is remembered', async () => {
  const urls: string[] = []
  const fetcher = (async (url: string) => {
    urls.push(url)
    if (url.includes('models?')) return new Response(JSON.stringify({ models: ['gemini-9.9-flash', 'gemini-9.8-flash', 'gemini-flash-latest', 'gemini-2.5-flash'].map((n) => ({ name: `models/${n}`, supportedGenerationMethods: ['generateContent'] })) }), { status: 200 })
    if (/gemini-9\.\d-flash:/.test(url)) return new Response(JSON.stringify({ error: { message: 'This model is currently experiencing high demand.' } }), { status: 503 })
    return new Response(JSON.stringify({ candidates: [{ content: { parts: [{ text: 'OK' }] } }] }), { status: 200 })
  }) as unknown as typeof fetch
  const c = new GeminiClient({ connection: 'key', apiKey: 'k', proxyUrl: '', model: 'gemini-9.9-flash' }, fetcher, [0, 0])
  assert.equal(await c.generate({ system: 's', turns: [{ role: 'user', text: 'hi' }] }), 'OK')
  assert.equal(c.model, 'gemini-flash-latest', 'the stable alias, not the next-newest (also overloaded) model')
  assert.ok(!urls.some((u) => u.includes('gemini-9.8-flash')))
  const before = urls.length
  await c.generate({ system: 's', turns: [{ role: 'user', text: 'again' }] })
  assert.equal(urls.length, before + 1, 'later calls go straight to the stand-in')
})

test('if the first stand-in is also overloaded, the next one is tried', async () => {
  const fetcher = (async (url: string) => {
    if (url.includes('models?')) return new Response(JSON.stringify({ models: ['gemini-8.1-flash', 'gemini-flash-latest', 'gemini-2.5-flash'].map((n) => ({ name: `models/${n}`, supportedGenerationMethods: ['generateContent'] })) }), { status: 200 })
    if (/gemini-9\.0-flash-x:|gemini-8\.1-flash:|flash-latest:/.test(url)) return new Response(JSON.stringify({ error: { message: 'high demand' } }), { status: 503 })
    return new Response(JSON.stringify({ candidates: [{ content: { parts: [{ text: 'OK' }] } }] }), { status: 200 })
  }) as unknown as typeof fetch
  const c = new GeminiClient({ connection: 'key', apiKey: 'k', proxyUrl: '', model: 'gemini-9.0-flash-x' }, fetcher, [0])
  assert.equal(await c.generate({ system: 's', turns: [{ role: 'user', text: 'hi' }] }), 'OK')
  assert.equal(c.model, 'gemini-2.5-flash')
})

test('Groq connection: OpenAI-style request, JSON mode with the schema in the prompt, rate limits fall back to another model', async () => {
  const bodies: Record<string, unknown>[] = []
  const headers: Record<string, string>[] = []
  const fetcher = (async (url: string, init?: RequestInit) => {
    if (url.endsWith('/models')) return new Response(JSON.stringify({ data: [{ id: 'llama-3.3-70b-versatile' }, { id: 'openai/gpt-oss-120b' }, { id: 'whisper-large-v3' }] }), { status: 200 })
    assert.ok(url.startsWith('https://api.groq.com/openai/v1/chat/completions'))
    const body = JSON.parse(String(init?.body))
    bodies.push(body)
    headers.push(init?.headers as Record<string, string>)
    if (body.model === 'llama-3.3-70b-versatile') return new Response(JSON.stringify({ error: { message: 'Rate limit reached' } }), { status: 429 })
    return new Response(JSON.stringify({ choices: [{ message: { content: '{"a":1}' } }] }), { status: 200 })
  }) as unknown as typeof fetch
  const c = new GeminiClient({ connection: 'groq', apiKey: '', proxyUrl: '', model: '', groqKey: 'gsk_x', groqModel: '' }, fetcher, [0])
  assert.equal(c.available, true)
  const out = await c.generateJson<{ a: number }>({ system: 'sys', turns: [{ role: 'user', text: 'hi' }, { role: 'model', text: 'yo' }, { role: 'user', text: 'again' }], schema: { type: 'object', properties: { a: { type: 'number' } } } })
  assert.deepEqual(out, { a: 1 })
  assert.equal(c.model, 'openai/gpt-oss-120b')
  const last = bodies[bodies.length - 1] as { messages: { role: string; content: string }[]; response_format: { type: string } }
  assert.equal(last.response_format.type, 'json_object')
  assert.equal(last.messages[0].role, 'system')
  assert.match(last.messages[0].content, /JSON schema/)
  assert.deepEqual(last.messages.map((m) => m.role), ['system', 'user', 'assistant', 'user'])
  assert.equal(headers[0].Authorization, 'Bearer gsk_x')
  assert.deepEqual(await c.listModels(), ['llama-3.3-70b-versatile', 'openai/gpt-oss-120b'])
})
