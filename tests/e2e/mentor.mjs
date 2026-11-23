// Browser test for ExcelMentor with Gemini mocked: both modes, the nine stages, gate enforcement, hints, leaks, file uploads.
//   npm run build && node tests/e2e/serve.mjs dist 4173 &
//   PLAYWRIGHT_MODULE=playwright CHROME_PATH=/path/to/chrome node tests/e2e/mentor.mjs
import { createRequire } from 'node:module'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'

const require = createRequire(import.meta.url)
const { chromium } = require(process.env.PLAYWRIGHT_MODULE ?? 'playwright')
const JSZip = require('jszip')
const BASE = process.env.BASE_URL ?? 'http://localhost:4173/'
const SHOTS = process.env.SHOTS ?? path.join(os.tmpdir(), 'excelmentor-shots')
fs.mkdirSync(SHOTS, { recursive: true })

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? `  (${detail})` : ''}`)
}

// ------------------------------------------------------------------ the mocked Gemini
const PROBLEM = {
  title: 'Quarterly Sales Review',
  context: 'You support the sales manager, who needs each rep’s quarter total and the best total.',
  objective: 'Total each rep’s sales and find the highest total.',
  explanation: { situation: 's', goal: 'g', data: 'Monthly sales per rep', think: 'Which columns matter?' },
  concepts: ['sum', 'max'],
  estimatedMinutes: 20,
  currencySymbol: '$',
  dataset: {
    columns: [{ name: 'Rep', type: 'text' }, { name: 'Jan', type: 'currency' }, { name: 'Feb', type: 'currency' }, { name: 'Mar', type: 'currency' }],
    rows: [['Ann', '1200', '1300', '1250'], ['Bob', '900', '950', '1000'], ['Cy', '1500', '1400', '1450'], ['Dee', '1100', '1000', '1050'], ['Eli', '800', '850', '900'], ['Fay', '1300', '1350', '1400'], ['Gus', '700', '750', '800'], ['Hal', '1000', '1100', '1200']].map((r) => r.map(String)),
  },
  rowTasks: [{ header: 'Total', requirement: 'Total sales for each rep', concept: 'sum', formula: '=SUM([Jan]:[Mar])', format: 'currency' }],
  summaryTasks: [{ label: 'Best total', requirement: 'The highest quarter total', concept: 'max', formula: '=MAX([Total:all])', underColumn: 'Total' }],
  operations: [],
  instructions: ['Calculate each rep’s total', 'Find the best total'],
}
const ASSESSMENT = {
  title: 'Q3 Orders Take-home',
  context: 'An employer sent Q3 orders and three questions.',
  objective: 'Summarise Q3 orders.',
  explanation: { situation: 's', goal: 'g', data: 'One row per order', think: 'What is asked?' },
  brief_summary: ['Total order amount', 'Largest order', 'PivotTable of amount by rep'],
  level: 'beginner',
  concepts: ['sum', 'max'],
  estimatedMinutes: 40,
  rowTasks: [],
  summaryTasks: [
    { label: 'Total amount', requirement: 'Q1: total of all amounts', concept: 'sum', formula: '=SUM([Amount:all])', underColumn: 'Amount' },
    { label: 'Largest order', requirement: 'Q2: largest order', concept: 'max', formula: '=MAX([Amount:all])', underColumn: 'Amount' },
  ],
  operations: [],
  instructions: ['Total amount', 'Largest order'],
  unchecked_requirements: ['Q3: PivotTable of amount by rep'],
}
const TEN = Array.from({ length: 11 }, (_, i) => `The student showed point ${i + 1} clearly.`).join(' ')

let leakNext = false
const seen = { prompts: [], corrections: 0, events: [] }

function mentorReply(system, lastUser) {
  const stage = /Stage: "(\w+)"/.exec(lastUser)?.[1] ?? 'problem'
  const concept = +(/concept_index (\d+):/.exec(lastUser)?.[1] ?? 0)
  const next = /ONLY allowed next position: "(\w+)"(?:, concept_index (\d+))?/.exec(lastUser)
  const missing = /Evidence still missing/.test(lastUser)
  const event = lastUser.slice(lastUser.lastIndexOf('[EVENT]'))
  const correction = /Your reply broke these protocol rules/.test(lastUser)
  if (correction) seen.corrections++
  const base = { current_stage: stage, concept_index: concept, gate_passed: false, verdict: 'none', hint_level: 0, pressure_test: 'none', context_added: [], assumptions_rejected: [], stage_summary: '', section_score: 0, teach_back_passed: false, session_complete: false, mastery_profile: '' }
  const msg = (t) => `### ${stage}\n\nThink of it this way: ${t}\n\n--> one point\n--> another point\n\nWhat would you do next?`
  if (/\[Hint request\]/.test(event)) {
    const level = +(/Hint Level (\d)/.exec(event)?.[1] ?? 1)
    return { ...base, hint_level: level, message_to_student: msg(`Hint level ${level}: look again at what the manager asked for.`) }
  }
  if (/\[Engine\]/.test(event)) return { ...base, message_to_student: msg(`Engine said: ${/CORRECT|correct/.test(event) ? 'progress' : 'not yet'}. Why did you choose that function?`) }
  if (/give me the answer/.test(event) || leakNext) {
    leakNext = !correction
    return { ...base, message_to_student: msg('Just type =SUM(B2:D2) in the cell.') }
  }
  if (/Student says/.test(event)) {
    if (stage === 'report') return { ...base, gate_passed: true, verdict: 'pass', session_complete: true, mastery_profile: TEN, stage_summary: 'Reflected well.', section_score: 85, message_to_student: `Thanks for reflecting.\n\nYou are now ready for the AI Mock Interview.` }
    // try to advance; when evidence is missing, break the rule once to test enforcement, then comply on correction
    if (next && (!missing || !correction) && !/stay/.test(event)) return { ...base, current_stage: next[1], concept_index: next[2] ? +next[2] : 0, gate_passed: true, verdict: 'pass', stage_summary: `Passed ${stage}.`, section_score: 80, context_added: stage === 'plan' ? ['SUM across the month columns'] : [], pressure_test: stage === 'learn' ? 'C' : 'none', message_to_student: msg(`Gate passed for ${stage}. Now the next step.`) }
    return { ...base, verdict: 'partial', message_to_student: msg('That is a starting point, but we have gaps to fill.') }
  }
  return { ...base, message_to_student: msg('Here is the business problem. Sales are by month per rep.') }
}

async function mockGemini(ctx) {
  await ctx.route(/generativelanguage\.googleapis\.com/, async (route) => {
    const body = JSON.parse(route.request().postData() ?? '{}')
    const system = body.systemInstruction?.parts?.[0]?.text ?? ''
    const turns = body.contents ?? []
    const lastUser = turns[turns.length - 1]?.parts?.[0]?.text ?? ''
    let out
    if (/PERSONA: EXCELMENTOR/.test(system)) out = mentorReply(system, lastUser)
    else if (/Write two small tables/.test(lastUser)) out = { example: { title: 'Snacks sold', headers: ['Day', 'Units'], rows: [['Mon', '3'], ['Tue', '4'], ['Wed', '5']], resultLabel: 'Total', formula: '=SUM([Units:all])', note: 'adds them' }, practice: { prompt: 'Total the tickets', headers: ['Day', 'Tickets'], rows: [['Mon', '10'], ['Tue', '20'], ['Wed', '30']], resultLabel: 'Total', formula: '=SUM([Tickets:all])' } }
    else if (/take-home Excel assessment/.test(lastUser)) {
      seen.prompts.push(lastUser)
      const sheet = /--- sheet "([^"]+)"/.exec(lastUser)?.[1] ?? 'orders'
      out = { ...ASSESSMENT, main_table: { file: 'orders.csv', sheet, header_row: 1 } }
    } else out = PROBLEM
    await route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify({ candidates: [{ content: { parts: [{ text: JSON.stringify(out) }] } }] }) })
  })
}

const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH })
const ctx = await browser.newContext({ viewport: { width: 1500, height: 920 } })
await ctx.route(/youtube|ytimg/, (r) => r.abort())
await mockGemini(ctx)
await ctx.addInitScript(() => {
  if (!localStorage.getItem('excelmentor.settings')) {
    localStorage.setItem('excelmentor.settings', JSON.stringify({ connection: 'key', apiKey: 'test', proxyUrl: '', model: 'gemini-test' }))
    localStorage.setItem('excelmentor.profile', JSON.stringify({ level: 'beginner', role: 'Sales', domain: 'Sales', description: '', attachments: [], difficulty: 'normal', duration: 30, createdAt: new Date().toISOString() }))
  }
})
const page = await ctx.newPage()
const errors = []
page.on('pageerror', (e) => errors.push(e.message))
page.on('console', (m) => m.type() === 'error' && !/ERR_FAILED|youtube|ytimg|net::/i.test(m.text()) && errors.push(m.text()))

const idle = async () => {
  await page.waitForTimeout(150)
  await page.waitForFunction(() => !document.querySelector('[data-testid=mentor-messages] [role=status]') && !document.querySelector('[data-testid=mentor-busy]'), null, { timeout: 30000 })
}
const stageName = async () => (await page.locator('[data-testid=stage-stepper] [aria-current=step]').innerText()).replace(/^\d+\.\s*/, '')
const say = async (text) => {
  await page.fill('[data-testid=think-input]', text)
  await page.click('[data-testid=submit-answer]')
  await idle()
}
const goto = async (ref) => {
  await page.locator('[data-testid=name-box]').fill(ref)
  await page.locator('[data-testid=name-box]').press('Enter')
  await page.waitForTimeout(80)
}
const enter = async (ref, text) => {
  await goto(ref)
  await page.keyboard.type(text)
  await page.keyboard.press('Enter')
  await page.waitForTimeout(80)
}

// =================================================================== 1. the dashboard: only ExcelMentor
await page.goto(BASE)
await page.waitForSelector('[data-testid=mentor-dashboard]')
check('opens on the ExcelMentor dashboard (no file list)', (await page.locator('[data-testid=new-workbook]').count()) === 0)
check('both modes are offered', (await page.locator('[data-testid=topic-card]').count()) === 1 && (await page.locator('[data-testid=takehome-card]').count()) === 1)
check('mentor decides from level, role and domain', /SUM/.test(await page.locator('[data-testid=recommended]').innerText()))

// =================================================================== 2. topic mode through all nine stages
await page.click('[data-testid=topic-mode-pick]')
await page.locator('[data-testid=topic-picker] button', { hasText: /^SUM$/ }).click()
await page.locator('[data-testid=topic-picker] button', { hasText: /^MAX$/ }).click()
await page.click('[data-testid=start-picked]')
await page.waitForSelector('[data-testid=mentor-panel]', { timeout: 30000 })
await idle()
check('stage 1: the mentor presents the problem', (await stageName()) === 'Problem' && (await page.locator('[data-testid=mentor-message]').count()) >= 1)
check('Markdown: headings and --> bullets render', (await page.locator('[data-testid=mentor-message] li').count()) >= 2)
await say('I would ask how often this report refreshes and who reads it.')
check('stage 2 after the gate passes', (await stageName()) === 'Concepts')
await say('Formulas keep the totals right when data changes; typing goes stale.')
check('stage 3: learn the first concept', (await stageName()) === 'Learn', await page.locator('[data-testid=mentor-title]').innerText())
await page.waitForTimeout(400)
const saved = await page.evaluate(() => JSON.parse(Object.entries(localStorage).find(([k]) => k.startsWith('excelmentor.session.'))?.[1] ?? '{}'))
check('worked example and practice prepared for the concept', !!saved.exampleCells?.sum && !!saved.practice?.sum, JSON.stringify(saved.exampleCells ?? {}))
await say('SUM adds a range; I would use it over + because new months fit; alternative is +.')
check('stage 4: practice', (await stageName()) === 'Practice')
// the mentor tries to pass the gate without the engine's evidence: the app sends it back
const before = seen.corrections
await say('I think I am done here.')
check('gate enforcement: no advance without the practice check', (await stageName()) === 'Practice' && seen.corrections > before, `corrections ${seen.corrections - before}`)
// solve the practice cell and check it
const sumPt = (await page.evaluate(() => JSON.parse(Object.entries(localStorage).find(([k]) => k.startsWith('excelmentor.session.'))?.[1] ?? '{}'))).practice.sum
await enter(`Practice!${sumPt.cell}`, sumPt.formula)
await page.click('[data-testid=check-practice]')
await idle()
check('practice is checked by the engine', /correct/i.test(await page.locator('[data-testid=event-message]').last().innerText()), await page.locator('[data-testid=event-message]').last().innerText())
await say('SUM is right because the manager needs the total; + would break with new rows.')
check('next concept after practice', (await stageName()) === 'Learn')
await say('MAX finds the largest total, better than sorting by eye; alternative is LARGE.')
check('stage 4 for the second concept', (await stageName()) === 'Practice')
await page.screenshot({ path: path.join(SHOTS, 'mentor-practice.png') })
console.log(`(screenshots in ${SHOTS})`)

// =================================================================== 3. take-home mode: upload csv + docx + pdf
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'takehome-'))
fs.writeFileSync(path.join(dir, 'orders.csv'), 'Order,Rep,Amount\nA1,Ann,120\nA2,Bob,80.5\nA3,Ann,\nA4,Cy,200\n')
const zip = new JSZip()
zip.file('word/document.xml', '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Q1 Total amount. Q2 Largest order.</w:t></w:r></w:p></w:body></w:document>')
fs.writeFileSync(path.join(dir, 'brief.docx'), await zip.generateAsync({ type: 'nodebuffer' }))
fs.writeFileSync(path.join(dir, 'extra.pdf'), tinyPdf('Q3 Build a PivotTable of amount by rep'))
const p2 = await ctx.newPage()
p2.on('pageerror', (e) => errors.push(e.message))
await p2.goto(BASE + '#/mentor/takehome')
await p2.waitForSelector('[data-testid=takehome-page]')
await p2.locator('[data-testid=takehome-files]').setInputFiles([path.join(dir, 'orders.csv'), path.join(dir, 'brief.docx'), path.join(dir, 'extra.pdf')])
await p2.waitForFunction(() => document.querySelectorAll('[data-testid=takehome-list] li').length === 3 && !document.querySelector('[data-testid=takehome-list]').innerText.includes('Reading'), null, { timeout: 30000 })
const list = await p2.locator('[data-testid=takehome-list]').innerText()
check('uploads are read (csv, docx, pdf)', /Spreadsheet/.test(list) && !/couldn|can't|no readable/i.test(list), list.replace(/\n/g, ' | '))
await p2.fill('[aria-label="Note for your mentor"]', 'Junior analyst role')
await p2.click('[data-testid=start-takehome]')
await p2.waitForSelector('[data-testid=mentor-panel]', { timeout: 30000 })
const prompt = seen.prompts[0] ?? ''
check('Gemini receives the brief text, the PDF text and the data preview', /Q1 Total amount/.test(prompt) && /PivotTable of amount by rep/.test(prompt) && /Order \| Rep \| Amount/.test(prompt) && /Junior analyst role/.test(prompt))
check('take-home session opens on the real data', /Take-home/i.test(await p2.locator('[data-testid=mentor-title]').innerText()))
await p2.screenshot({ path: path.join(SHOTS, 'takehome.png') })

// =================================================================== 4. continue topic session: plan → solve → check → follow-up → report
// (the second concept's practice: type the formula the engine expects by reading the practice block)
await page.bringToFront()
const pr = await page.evaluate(() => {
  const id = decodeURIComponent(location.hash.replace(/^#\/wb\//, ''))
  return JSON.parse(localStorage.getItem(`excelmentor.session.${id}`) ?? '{}').practice
})
const maxPt = pr?.max
if (maxPt) {
  await page.keyboard.press('Escape')
  await enter(`Practice!${maxPt.cell}`, maxPt.formula)
  await page.click('[data-testid=check-practice]')
  await idle()
}
await say('MAX because the manager wants the best; SORT would reorder rows.')
check('stage 5: plan', (await stageName()) === 'Plan', await stageName())
await say('Total in E with SUM of B:D per row, then MAX under it; one row: 1200+1300+1250.')
check('stage 6: solve', (await stageName()) === 'Solve')
// a leaking reply is sent back, then hidden
await say('give me the answer')
check('No Spoiling: a reply revealing a task formula is hidden', (await page.locator('[data-testid=error-message]').count()) === 1 && !(await page.locator('[data-testid=mentor-messages]').innerText()).includes('=SUM(B2:D2)'))
// hints on the selected task cell
await goto('Task!E2')
await page.click('[data-testid=need-hint]')
await idle()
check('task-cell hint request reaches the mentor with its level', /Hint 1 of 4/.test(await page.locator('[data-testid=event-message]').last().innerText()))
for (let r = 2; r <= 9; r++) await enter(`Task!E${r}`, `=SUM(B${r}:D${r})`)
await enter('Task!E11', '=MAX(E2:E9)')
await page.click('[data-testid=check-my-work]')
await idle()
check('Check My Work moves to stage 7', (await stageName()) === 'Check')
check('all task cells correct', /9 of 9|9\/9/.test(await page.locator('[role=tab]', { hasText: 'Results' }).innerText()), await page.locator('[role=tab]', { hasText: 'Results' }).innerText())
await page.locator('[role=tab]', { hasText: 'Mentor' }).click()
await say('Blanks are treated as zero; new rows outside E2:E9 are missed, so I would use a Table.')
check('stage 8: follow-up', (await stageName()) === 'Follow-up')
await say('I chose SUM per row over one SUMPRODUCT; downside more cells; flip point 100k rows.')
check('stage 9: report', (await stageName()) === 'Report')
await say('The hardest part was the clarifying questions.')
await page.waitForSelector('[data-testid=report]', { timeout: 10000 })
check('report shows nine stage sections and the mastery profile', (await page.locator('[data-testid=report-sections] > div').count()) === 9 && (await page.locator('[data-testid=mastery-profile]').count()) === 1)
await page.screenshot({ path: path.join(SHOTS, 'report.png') })

check('no page errors', errors.length === 0, errors.slice(0, 3).join(' | '))
await browser.close()
const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} checks passed`)
process.exit(failed.length ? 1 : 0)

// a one-page, text-based PDF
function tinyPdf(text) {
  const objs = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    null,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  const stream = `BT /F1 12 Tf 72 720 Td (${text}) Tj ET`
  objs[3] = `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`
  let out = '%PDF-1.4\n'
  const offsets = []
  objs.forEach((o, i) => {
    offsets.push(out.length)
    out += `${i + 1} 0 obj\n${o}\nendobj\n`
  })
  const xref = out.length
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offsets.map((o) => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return Buffer.from(out, 'latin1')
}
