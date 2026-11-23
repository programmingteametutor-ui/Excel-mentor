// Browser test for formula range dragging and the Analysis ToolPak / Solver add-ins.
//   npm run build && node tests/e2e/serve.mjs dist 4173 &
//   PLAYWRIGHT_MODULE=playwright CHROME_PATH=/path/to/chrome node tests/e2e/addins.mjs
import { createRequire } from 'node:module'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'

const require = createRequire(import.meta.url)
const { chromium } = require(process.env.PLAYWRIGHT_MODULE ?? 'playwright')
const BASE = process.env.BASE_URL ?? 'http://localhost:4173/'
const SHOTS = process.env.SHOTS ?? path.join(os.tmpdir(), 'excelpro-shots')
fs.mkdirSync(SHOTS, { recursive: true })

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? `  (${detail})` : ''}`)
}

const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH })
const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
const page = await ctx.newPage()
const errors = []
page.on('pageerror', (e) => errors.push(e.message))

await page.goto(BASE)
await page.evaluate(() => localStorage.clear())
await page.goto(BASE)
await page.click('[data-testid=new-workbook]')
await page.waitForSelector('[data-testid=grid-scroller]')
let box = await page.locator('[data-testid=grid-scroller]').boundingBox()
const xy = (r, c) => [box.x + c * 88 + 44, box.y + r * 22 + 11]
const bar = () => page.locator('[data-testid=formula-bar]').inputValue()
const drag = async (a, b) => {
  await page.mouse.move(...xy(...a))
  await page.mouse.down()
  await page.mouse.move(...xy(...b), { steps: 6 })
  await page.mouse.up()
}
const typeAt = async (r, c, text) => {
  await page.mouse.click(...xy(r, c))
  await page.keyboard.type(text)
  await page.keyboard.press('Enter')
}
const cellValue = async (r, c) => {
  await page.mouse.click(...xy(r, c))
  return bar()
}
const valueAt = async (ref) => {
  const nb = page.locator('[data-testid=name-box]')
  await nb.click()
  await nb.fill(ref)
  await nb.press('Enter')
  await page.waitForTimeout(60)
  return bar()
}
const tab = (name) => page.getByRole('tab', { name, exact: true }).click()

// ---------------------------------------------------------------- formula drag
const scores = [72, 85, 90, 66, 78, 95, 88, 70, 81, 77]
const hours = [5, 8, 9, 3, 6, 10, 8, 4, 7, 6]
await typeAt(0, 0, 'Score')
await typeAt(0, 1, 'Hours')
for (let i = 0; i < scores.length; i++) {
  await typeAt(i + 1, 0, String(scores[i]))
  await typeAt(i + 1, 1, String(hours[i]))
}
// a formula in D1 whose range sits directly under the formula hint popup
await page.mouse.click(...xy(0, 3))
await page.keyboard.type('=SUM(')
await drag([1, 3], [4, 3])
check('drag selects a range under the formula hint popup', (await bar()) === '=SUM(D2:D5', await bar())
await page.keyboard.press('Escape')
await page.mouse.click(...xy(0, 3))
await page.keyboard.type('=AVERAGE(')
await drag([1, 0], [10, 0])
check('drag inserts a column range', (await bar()) === '=AVERAGE(A2:A11', await bar())
check('referenced range is outlined in colour', (await page.locator('[data-testid=ref-highlight]').count()) >= 1)
await page.keyboard.press('Enter')
check('formula commits with the dragged range', (await cellValue(0, 3)) === '=AVERAGE(A2:A11)')

// ---------------------------------------------------------------- add-ins on the Data tab
await tab('Data')
check('Data Analysis button is on the Data tab', await page.getByRole('button', { name: /^Data Analysis:/ }).isVisible())
check('Solver button is on the Data tab', await page.getByRole('button', { name: /^Solver: find/ }).isVisible())
await page.getByRole('button', { name: /^Add-ins:/ }).click()
await page.getByLabel('Solver Add-in').uncheck()
await page.getByRole('button', { name: 'OK', exact: true }).click()
check('turning Solver off hides its button', (await page.getByRole('button', { name: /^Solver: find/ }).count()) === 0)
await page.getByRole('button', { name: /^Add-ins:/ }).click()
await page.getByLabel('Solver Add-in').check()
await page.getByRole('button', { name: 'OK', exact: true }).click()
check('turning Solver on shows it again', await page.getByRole('button', { name: /^Solver: find/ }).isVisible())

// ---------------------------------------------------------------- Descriptive Statistics
await page.mouse.click(...xy(2, 0))
await page.getByRole('button', { name: /^Data Analysis:/ }).click()
await page.getByRole('option', { name: 'Descriptive Statistics' }).click()
await page.getByTestId('analysis-ok').click()
check('input range is prefilled from the data around the active cell', (await page.getByTestId('f-input').inputValue()) === '$A$1:$B$11', await page.getByTestId('f-input').inputValue())
check('labels are detected', await page.getByLabel('Labels in first row').isChecked())
await page.screenshot({ path: path.join(SHOTS, 'addins-descriptive.png') })
await page.getByTestId('tool-ok').click()
await page.waitForTimeout(150)
const tabs = await page.getByTestId('sheet-tabs').innerText()
check('results go to a new sheet', /Analysis1/.test(tabs), tabs.replace(/\n/g, ' '))
check('Mean of Score is correct', (await valueAt('B3')) === String(scores.reduce((a, b) => a + b) / 10), await valueAt('B3'))
check('Mean label is in A3', (await valueAt('A3')) === 'Mean')
await page.screenshot({ path: path.join(SHOTS, 'addins-descriptive-output.png') })

// ---------------------------------------------------------------- Regression with range picking
await page.getByTestId('sheet-tabs').getByText('Sheet1').click()
await page.waitForTimeout(100)
box = await page.locator('[data-testid=grid-scroller]').boundingBox()
await tab('Data')
await page.getByRole('button', { name: /^Data Analysis:/ }).click()
await page.getByRole('option', { name: 'Regression' }).click()
await page.getByTestId('analysis-ok').click()
await page.getByRole('button', { name: 'Select Input Y Range on the sheet' }).click()
check('dialog collapses to the pick bar', await page.getByTestId('pick-bar').isVisible())
await drag([0, 0], [10, 0])
await page.getByTestId('pick-done').click()
check('picked Y range fills the box', (await page.getByTestId('f-y').inputValue()) === '$A$1:$A$11', await page.getByTestId('f-y').inputValue())
await page.getByTestId('f-x').fill('B1:B11')
await page.getByLabel('Labels', { exact: true }).check()
await page.getByLabel('Residuals', { exact: true }).check()
await page.getByTestId('tool-ok').click()
await page.waitForTimeout(150)
const r2 = await valueAt('B5')
check('regression R Square is written', Number(r2) > 0.8 && Number(r2) < 1, r2)
await page.screenshot({ path: path.join(SHOTS, 'addins-regression.png') })

// ---------------------------------------------------------------- Solver
await page.getByTestId('sheet-tabs').getByText('Sheet1').click()
await page.waitForTimeout(100)
const model = [
  ['Doors', '0', '3'],
  ['Windows', '0', '5'],
]
await typeAt(14, 0, 'Product')
await typeAt(14, 1, 'Units')
await typeAt(14, 2, 'Profit')
for (let i = 0; i < model.length; i++) for (let j = 0; j < 3; j++) await typeAt(15 + i, j, model[i][j])
await typeAt(17, 0, 'Total')
await typeAt(17, 2, '=SUMPRODUCT(B16:B17,C16:C17)')
await typeAt(18, 0, 'Plant 1')
await typeAt(18, 1, '=B16')
await typeAt(18, 2, '4')
await typeAt(19, 0, 'Plant 2')
await typeAt(19, 1, '=2*B17')
await typeAt(19, 2, '12')
await typeAt(20, 0, 'Plant 3')
await typeAt(20, 1, '=3*B16+2*B17')
await typeAt(20, 2, '18')
await page.mouse.click(...xy(17, 2))
await tab('Data')
await page.getByRole('button', { name: /^Solver: find/ }).click()
check('objective defaults to the active formula cell', (await page.getByTestId('solver-objective').inputValue()) === '$C$18', await page.getByTestId('solver-objective').inputValue())
await page.getByTestId('solver-variables').fill('B16:B17')
await page.getByTestId('constraint-add').click()
await page.getByTestId('c-ref').fill('B19:B21')
await page.getByTestId('c-rhs').fill('C19:C21')
await page.getByTestId('constraint-ok').click()
check('constraint is listed', (await page.getByTestId('solver-constraints').innerText()).includes('B19:B21 <= C19:C21'))
await page.getByTestId('solver-method').selectOption('simplex')
await page.screenshot({ path: path.join(SHOTS, 'addins-solver.png') })
await page.getByTestId('solver-solve').click()
await page.getByTestId('solver-message').waitFor()
const msg = await page.getByTestId('solver-message').innerText()
check('Solver finds the optimum', /found a solution/.test(msg), msg)
await page.getByLabel('Answer').check()
await page.getByLabel('Sensitivity').check()
await page.screenshot({ path: path.join(SHOTS, 'addins-solver-results.png') })
await page.getByTestId('solver-results-ok').click()
await page.waitForTimeout(200)
const tabs2 = await page.getByTestId('sheet-tabs').innerText()
check('answer and sensitivity reports are created', /Answer Report 1/.test(tabs2) && /Sensitivity Report 1/.test(tabs2), tabs2.replace(/\n/g, ' '))
await page.screenshot({ path: path.join(SHOTS, 'addins-answer-report.png') })
await page.getByTestId('sheet-tabs').getByText('Sheet1').click()
await page.waitForTimeout(100)
check('variable cells hold the solution', (await cellValue(15, 1)) === '2' && (await cellValue(16, 1)) === '6')
check('objective is 36', (await page.locator('.xp-cell', { hasText: /^36$/ }).count()) >= 1)
// the model is remembered
await tab('Data')
await page.getByRole('button', { name: /^Solver: find/ }).click()
check('Solver remembers the model on the sheet', (await page.getByTestId('solver-variables').inputValue()) === 'B16:B17')
await page.getByRole('button', { name: 'Close', exact: true }).last().click()

check('no page errors', errors.length === 0, errors.join(' | '))
await browser.close()
const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} passed`)
process.exit(failed.length ? 1 : 0)
