import { test } from 'node:test'
import assert from 'node:assert/strict'
import { Workbook } from '../src/engine/workbook'
import * as st from '../src/engine/stats'
import {
  anovaSingleFactor,
  anovaTwoFactorNoRep,
  anovaTwoFactorRep,
  correlationTool,
  covarianceTool,
  descriptiveStatistics,
  exponentialSmoothing,
  formulaCtx,
  fourierTool,
  fTestTwoSample,
  histogramTool,
  movingAverage,
  randomNumbers,
  rankAndPercentile,
  readSeries,
  regressionTool,
  samplingTool,
  tTestPaired,
  tTestTwoSample,
  writeTable,
  zTestTwoSample,
  type OutTable,
} from '../src/engine/features/analysis'
import { answerReport, sensitivityReport, solve, solveLP } from '../src/engine/features/solver'
import type { SolverModel } from '../src/model/types'
import { formulaRefs } from '../src/engine/formulaText'

const close = (a: unknown, b: number, tol = 1e-6) => {
  assert.equal(typeof a, 'number', `expected a number, got ${String(a)}`)
  assert.ok(Math.abs((a as number) - b) <= tol * Math.max(1, Math.abs(b)), `${a} ≠ ${b}`)
}
const v = (c: unknown) => (c !== null && typeof c === 'object' && 'v' in (c as object) ? (c as { v: unknown }).v : c)
/** value next to a row label anywhere in the table */
const find = (t: OutTable | string, label: string, col = 1) => {
  assert.notEqual(typeof t, 'string', String(t))
  const row = (t as OutTable).rows.find((r) => v(r[0]) === label)
  assert.ok(row, `row “${label}” not found`)
  return v(row![col])
}
const S = (name: string, values: number[]) => ({ name, values })

function sheet(rows: (string | number | null)[][]) {
  const wb = Workbook.blank('T')
  const id = wb.order[0]
  wb.setInputs(id, rows.flatMap((r, i) => r.map((x, c) => ({ r: i, c, input: x === null ? '' : String(x) }))))
  return { wb, id }
}
const val = (wb: Workbook, id: string, r: number, c: number) => wb.getValue(id, r, c)

// ---------------------------------------------------------------- distributions & functions
test('statistical distributions match the engine and Excel', () => {
  close(st.tInv2T(0.05, 9), 2.262157163, 1e-8)
  close(st.tTwoTail(2.1, 7), 0.07387119621, 1e-8)
  close(st.fRight(3.2, 4, 12), 0.05257985916, 1e-8)
  close(st.fInv(0.95, 3, 20), 3.098391212, 1e-8)
  close(st.normInv(0.975), 1.959963985, 1e-9)
  close(st.chiCdf(5.2, 3), 0.8422755496, 1e-8)
  close(st.kurtosis([3, 4, 5, 2, 3, 4, 5, 6, 4, 7]), -0.151799637, 1e-8)
  close(st.determinant([[4, -1], [2, 1]]), 6)
})

test('new worksheet functions: LINEST, TREND, FORECAST, KURT, MINVERSE, …', () => {
  const { wb, id } = sheet([[1, 2], [2, 4], [3, 5], [4, 4], [5, 5]])
  wb.setInputs(id, [
    { r: 9, c: 0, input: '=LINEST(B1:B5,A1:A5,TRUE,TRUE)' },
    { r: 20, c: 0, input: '=TREND(B1:B5,A1:A5,6)' },
    { r: 20, c: 1, input: '=FORECAST.LINEAR(6,B1:B5,A1:A5)' },
    { r: 20, c: 2, input: '=INTERCEPT(B1:B5,A1:A5)' },
    { r: 20, c: 3, input: '=KURT(3,4,5,2,3,4,5,6,4,7)' },
    { r: 20, c: 4, input: '=RANK.AVG(4,B1:B5)' },
    { r: 20, c: 5, input: '=PERCENTRANK.INC(B1:B5,4)' },
    { r: 20, c: 6, input: '=TRIMMEAN(A1:B5,0.2)' },
    { r: 20, c: 7, input: '=PERMUT(100,3)' },
    { r: 22, c: 0, input: '=MINVERSE(A1:B2)' },
    { r: 22, c: 3, input: '=MDETERM(A1:B2)' },
    { r: 25, c: 0, input: '=GROWTH(B1:B5,A1:A5,6)' },
  ])
  // LINEST with statistics: slope, intercept / SEs / R², SEy / F, df / SSreg, SSres
  close(val(wb, id, 9, 0), 0.6)
  close(val(wb, id, 9, 1), 2.2)
  close(val(wb, id, 10, 0), 0.2828427125)
  close(val(wb, id, 10, 1), 0.938083152)
  close(val(wb, id, 11, 0), 0.6)
  close(val(wb, id, 11, 1), 0.894427191)
  close(val(wb, id, 12, 0), 4.5)
  close(val(wb, id, 12, 1), 3)
  close(val(wb, id, 13, 0), 3.6)
  close(val(wb, id, 13, 1), 2.4)
  close(val(wb, id, 20, 0), 5.8)
  close(val(wb, id, 20, 1), 5.8)
  close(val(wb, id, 20, 2), 2.2)
  close(val(wb, id, 20, 3), -0.151799637, 1e-8)
  close(val(wb, id, 20, 4), 3.5)
  close(val(wb, id, 20, 5), 0.25)
  close(val(wb, id, 20, 6), 3.625)
  close(val(wb, id, 20, 7), 970200)
  // inverse of [[1,2],[2,4]] is singular → #NUM!
  assert.equal((val(wb, id, 22, 0) as { type?: string }).type, 'NUM')
  close(val(wb, id, 22, 3), 0)
  assert.equal(typeof val(wb, id, 25, 0), 'number')
})

test('formulaRefs finds references for colour outlines', () => {
  const refs = formulaRefs('=SUM(B2:B7)+C2*Sheet2!A1+"A9"+LOG10(4)+$D$3+E:E')
  assert.deepEqual(refs.map((r) => r.text), ['B2:B7', 'C2', 'Sheet2!A1', '$D$3', 'E:E'])
  assert.equal(refs[2].sheet, 'Sheet2')
  assert.deepEqual([refs[0].r1, refs[0].r2, refs[0].c1], [1, 6, 1])
})

// ---------------------------------------------------------------- Analysis ToolPak
test('ANOVA single factor (textbook example: F = 9.2647)', () => {
  const t = anovaSingleFactor([S('A', [6, 8, 4, 5, 3, 4]), S('B', [8, 12, 9, 11, 6, 8]), S('C', [13, 9, 11, 8, 7, 12])], 0.05)
  close(find(t, 'Between Groups', 1), 84)
  close(find(t, 'Within Groups', 1), 68)
  close(find(t, 'Between Groups', 4), 9.264705882)
  close(find(t, 'Between Groups', 5), 0.002398777, 1e-5)
  close(find(t, 'Between Groups', 6), 3.682320344, 1e-7)
  close(find(t, 'Total', 2), 17)
})

test('ANOVA two-factor without and with replication', () => {
  const t = anovaTwoFactorNoRep([[4, 5, 6], [6, 7, 9], [8, 9, 12]], ['r1', 'r2', 'r3'], ['a', 'b', 'c'], 0.05)
  // grand mean 7.333; rows SS = 3·Σ(rowMean − G)² ; cols SS = 3·Σ(colMean − G)²
  close(find(t, 'Rows', 1), 3 * ((5 - 22 / 3) ** 2 + (22 / 3 - 22 / 3) ** 2 + (29 / 3 - 22 / 3) ** 2))
  close(find(t, 'Columns', 1), 3 * ((6 - 22 / 3) ** 2 + (7 - 22 / 3) ** 2 + (9 - 22 / 3) ** 2))
  const rep = anovaTwoFactorRep([[1, 2], [2, 3], [3, 6], [4, 7]], 2, ['g1', 'g2'], ['x', 'y'], 0.05)
  close(find(rep, 'Sample', 1), 4 * (2 - 3.5) ** 2 + 4 * (5 - 3.5) ** 2)
  close(find(rep, 'Within', 2), 4)
  const totalRow = (rep as OutTable).rows.filter((r) => v(r[0]) === 'Total').pop()!
  close(v(totalRow[1]), st.devsq([1, 2, 2, 3, 3, 6, 4, 7]))
})

test('t-tests agree with T.TEST', () => {
  const a = [14, 15, 15, 16, 13, 8, 14, 17, 16, 14, 19, 20, 21, 15, 15, 16, 16, 13, 14, 12]
  const b = [15, 17, 14, 17, 14, 8, 12, 19, 19, 14, 17, 22, 24, 16, 13, 16, 13, 18, 15, 13]
  const { wb, id } = sheet(a.map((x, i) => [x, b[i]]))
  wb.setInputs(id, [
    { r: 0, c: 3, input: '=T.TEST(A1:A20,B1:B20,2,1)' },
    { r: 1, c: 3, input: '=T.TEST(A1:A20,B1:B20,2,2)' },
  ])
  close(find(tTestPaired(S('a', a), S('b', b), 0, 0.05), 'P(T<=t) two-tail'), val(wb, id, 0, 3) as number, 1e-8)
  close(find(tTestTwoSample(S('a', a), S('b', b), 0, 0.05, true), 'P(T<=t) two-tail'), val(wb, id, 1, 3) as number, 1e-8)
  const uneq = tTestTwoSample(S('a', a), S('b', b), 0, 0.05, false)
  const q1 = st.variance(a) / 20
  const q2 = st.variance(b) / 20
  assert.equal(find(uneq, 'df'), Math.round((q1 + q2) ** 2 / (q1 ** 2 / 19 + q2 ** 2 / 19))) // Welch df rounded, as Excel's tool shows
  close(find(zTestTwoSample(S('a', a), S('b', b), 0, 9, 11, 0.05), 'z'), (st.mean(a) - st.mean(b)) / Math.sqrt(9 / 20 + 11 / 20))
  const f = fTestTwoSample(S('a', a), S('b', b), 0.05)
  close(find(f, 'F'), st.variance(a) / st.variance(b))
})

test('regression tool matches LINEST and reports residuals', () => {
  const y = [10, 12, 15, 18, 20, 25, 24, 30]
  const x1 = [1, 2, 3, 4, 5, 6, 7, 8]
  const x2 = [3, 1, 4, 1, 5, 9, 2, 6]
  const t = regressionTool(S('Sales', y), [S('Ads', x1), S('Price', x2)], { constantZero: false, confidence: 0.9, residuals: true, standardized: true, normalProbability: true, lineFitChart: false })
  close(find(t, 'R Square'), 0.9871202556)
  close(find(t, 'Intercept'), 6.315398075)
  close(find(t, 'Ads'), 2.576771654)
  close(find(t, 'Ads', 2), 0.1617904095)
  close(find(t, 'Regression', 4), 191.6032315, 1e-7)
  assert.ok((t as OutTable).rows.some((r) => v(r[0]) === 'Lower 90.0%' || r.some((c) => v(c) === 'Lower 90.0%')))
  assert.ok((t as OutTable).rows.some((r) => v(r[0]) === 'RESIDUAL OUTPUT'))
  assert.ok((t as OutTable).rows.some((r) => v(r[0]) === 'PROBABILITY OUTPUT'))
})

test('descriptive statistics, histogram, correlation, covariance, rank', () => {
  const xs = [3, 4, 5, 2, 3, 4, 5, 6, 4, 7]
  const d = descriptiveStatistics([S('Score', xs)], { summary: true, confidence: 0.95, kthLargest: 2, kthSmallest: 1 })
  close(find(d, 'Mean'), 4.3)
  close(find(d, 'Mode'), 4)
  close(find(d, 'Kurtosis'), -0.151799637, 1e-8)
  close(find(d, 'Largest(2)'), 6)
  close(find(d, 'Confidence Level(95.0%)'), st.tInv2T(0.05, 9) * (st.stdev(xs) / Math.sqrt(10)))
  const h = histogramTool(S('x', xs), { bins: [3, 5], pareto: true, cumulative: true, chart: true })
  assert.equal(typeof h, 'object')
  const rows = (h as OutTable).rows
  assert.deepEqual(rows.slice(1, 4).map((r) => v(r[1])), [3, 5, 2])
  assert.ok((h as OutTable).chart)
  const c = correlationTool([S('a', [1, 2, 3, 4]), S('b', [2, 4, 6, 9])])
  close(find(c, 'a', 1), 1)
  assert.equal(find(c, 'a', 2), null)
  close(find(covarianceTool([S('a', [1, 2, 3, 4]), S('b', [2, 4, 6, 8])]), 'b', 1), 2.5)
  const rk = rankAndPercentile([S('x', [10, 30, 20])]) as OutTable
  assert.deepEqual(rk.rows.slice(1).map((r) => [v(r[0]), v(r[2])]), [[2, 1], [3, 2], [1, 3]])
})

test('moving average and exponential smoothing write live formulas', () => {
  const { wb, id } = sheet([['Sales'], [10], [12], [14], [13], [17], [20]])
  const ctx = formulaCtx(wb, id, { sheetId: id, r: 0, c: 2 })
  const ma = movingAverage({ r1: 0, c1: 0, r2: 6, c2: 0 }, true, 3, true, false, ctx)
  assert.equal(typeof ma, 'object')
  writeTable(wb, ma as OutTable, { sheetId: id, r: 0, c: 2 }, 'Moving Average')
  assert.equal((val(wb, id, 0, 2) as { type?: string }).type, 'NA')
  close(val(wb, id, 2, 2), 12)
  close(val(wb, id, 5, 2), (14 + 13 + 17 + 20) / 4 === 0 ? 0 : (13 + 17 + 20) / 3)
  assert.equal(typeof val(wb, id, 4, 3), 'number')
  // edits flow through
  wb.setInputs(id, [{ r: 1, c: 0, input: '13' }])
  close(val(wb, id, 2, 2), 13)
  const ctx2 = formulaCtx(wb, id, { sheetId: id, r: 0, c: 5 })
  const es = exponentialSmoothing({ r1: 1, c1: 0, r2: 6, c2: 0 }, false, 0.3, false, false, ctx2)
  writeTable(wb, es as OutTable, { sheetId: id, r: 0, c: 5 }, 'Exponential Smoothing')
  close(val(wb, id, 1, 5), 13)
  close(val(wb, id, 2, 5), 0.7 * 12 + 0.3 * 13)
})

test('random numbers, sampling and Fourier analysis', () => {
  const r = randomNumbers(2, 500, { kind: 'normal', mean: 50, sd: 5 }, 42) as OutTable
  assert.equal(r.rows.length, 500)
  const col = r.rows.map((x) => x[0] as number)
  assert.ok(Math.abs(st.mean(col) - 50) < 1)
  assert.deepEqual(randomNumbers(2, 3, { kind: 'normal', mean: 0, sd: 1 }, 7), randomNumbers(2, 3, { kind: 'normal', mean: 0, sd: 1 }, 7))
  const pat = randomNumbers(1, 1, { kind: 'patterned', from: 1, to: 3, step: 1, repeatEach: 2, repeatSeq: 1 }, 1) as OutTable
  assert.deepEqual(pat.rows.map((x) => x[0]), [1, 1, 2, 2, 3, 3])
  const sp = samplingTool(S('x', [1, 2, 3, 4, 5, 6, 7]), 'periodic', 3) as OutTable
  assert.deepEqual(sp.rows.map((x) => x[0]), [3, 6])
  const ft = fourierTool([1, 2, 3, 4], false) as OutTable
  assert.deepEqual(ft.rows.map((x) => x[0]), ['10', '-2+2i', '-2', '-2-2i'])
  assert.equal(typeof fourierTool([1, 2, 3], false), 'string')
})

test('readSeries reports text cells and uses labels', () => {
  const { wb, id } = sheet([['Height', 'Weight'], [170, 65], [180, 'x']])
  assert.match(readSeries(wb, id, { r1: 0, c1: 0, r2: 2, c2: 1 }, 'cols', true) as string, /non-numeric/)
  const ok = readSeries(wb, id, { r1: 0, c1: 0, r2: 1, c2: 1 }, 'cols', true)
  assert.deepEqual(ok, [S('Height', [170]), S('Weight', [65])])
})

// ---------------------------------------------------------------- Solver
test('LP simplex solves the classic product-mix problem', () => {
  // max 3x + 5y  s.t. x ≤ 4, 2y ≤ 12, 3x + 2y ≤ 18
  const r = solveLP({ c: [3, 5], A: [[1, 0], [0, 2], [3, 2]], b: [4, 12, 18], ops: ['<=', '<=', '<='], lower: [0, 0], upper: [Infinity, Infinity], maximize: true })
  assert.equal(r.status, 'optimal')
  close(r.x[0], 2)
  close(r.x[1], 6)
  close(r.z, 36)
  assert.equal(solveLP({ c: [1], A: [[1]], b: [5], ops: ['>='], lower: [0], upper: [3], maximize: true }).status, 'infeasible')
  assert.equal(solveLP({ c: [1], A: [[1]], b: [5], ops: ['>='], lower: [0], upper: [Infinity], maximize: true }).status, 'unbounded')
})

function productMix() {
  // A: labels; B: quantities (variables); C: profit per unit; usage table
  return sheet([
    ['Product', 'Units', 'Profit'],
    ['Doors', 0, 3],
    ['Windows', 0, 5],
    ['Total profit', null, '=SUMPRODUCT(B2:B3,C2:C3)'],
    ['Plant', 'Used', 'Capacity'],
    ['Plant 1', '=B2', 4],
    ['Plant 2', '=2*B3', 12],
    ['Plant 3', '=3*B2+2*B3', 18],
  ])
}
const mixModel = (method: SolverModel['method']): SolverModel => ({
  objective: 'C4',
  goal: 'max',
  variables: 'B2:B3',
  constraints: [{ ref: 'B6:B8', op: '<=', rhs: 'C6:C8' }],
  nonNegative: true,
  method,
})

test('Solver: Simplex LP on a worksheet model, with reports', async () => {
  const { wb, id } = productMix()
  const res = await solve(wb, id, mixModel('simplex'))
  assert.ok(res.ok, res.message)
  assert.match(res.message, /found a solution/)
  close(res.x![0], 2)
  close(res.x![1], 6)
  close(res.objective, 36)
  // workbook untouched until applied
  assert.equal(val(wb, id, 1, 1), 0)
  wb.setInputs(id, res.details!.varCells.map((c, i) => ({ r: c.r, c: c.c, input: String(res.x![i]) })), 'Solver')
  close(val(wb, id, 3, 2), 36)
  const ans = answerReport(wb, id, mixModel('simplex'), res)
  assert.ok(ans.rows.some((r) => v(r[4]) === 'Binding'))
  const sens = sensitivityReport(wb, id, res)!
  const shadow = sens.rows.filter((r) => typeof v(r[3]) === 'number' && String(v(r[0])).startsWith('$B$') && [6, 7, 8].some((k) => v(r[0]) === `$B$${k}`)).map((r) => v(r[3]) as number)
  close(shadow[0], 0, 1e-6)
  close(shadow[1], 1.5, 1e-6)
  close(shadow[2], 1, 1e-6)
})

test('Solver: integer and binary constraints (knapsack)', async () => {
  const { wb, id } = sheet([
    ['Item', 'Take', 'Value', 'Weight'],
    ['A', 0, 10, 5],
    ['B', 0, 40, 4],
    ['C', 0, 30, 6],
    ['D', 0, 50, 3],
    ['Total', null, '=SUMPRODUCT(B2:B5,C2:C5)', '=SUMPRODUCT(B2:B5,D2:D5)'],
  ])
  const m: SolverModel = { objective: 'C6', goal: 'max', variables: 'B2:B5', constraints: [{ ref: 'D6', op: '<=', rhs: '10' }, { ref: 'B2:B5', op: 'bin' }], nonNegative: true, method: 'simplex' }
  const res = await solve(wb, id, m)
  assert.ok(res.ok, res.message)
  assert.deepEqual(res.x, [0, 1, 0, 1])
  close(res.objective, 90)
  // the same model with GRG + branch and bound
  const g = await solve(wb, id, { ...m, method: 'grg' })
  assert.ok(g.ok, g.message)
  close(g.objective, 90)
})

test('Solver: GRG nonlinear and Evolutionary', async () => {
  // minimise x² + y² subject to x + y ≥ 1  →  (0.5, 0.5), objective 0.5
  const { wb, id } = sheet([
    ['x', 3],
    ['y', -2],
    ['f', '=B1^2+B2^2'],
    ['g', '=B1+B2'],
  ])
  const m: SolverModel = { objective: 'B3', goal: 'min', variables: 'B1:B2', constraints: [{ ref: 'B4', op: '>=', rhs: '1' }], nonNegative: false, method: 'grg' }
  const r = await solve(wb, id, m)
  assert.ok(r.ok, r.message)
  close(r.x![0], 0.5, 1e-3)
  close(r.x![1], 0.5, 1e-3)
  close(r.objective, 0.5, 1e-4)
  // Simplex refuses a nonlinear model
  const lp = await solve(wb, id, { ...m, method: 'simplex' })
  assert.equal(lp.ok, false)
  assert.match(lp.message, /linearity/)
  // Evolutionary needs bounds…
  const noB = await solve(wb, id, { ...m, method: 'evolutionary' })
  assert.match(noB.message, /bounds/)
  // …and finds the optimum when given them
  const ev = await solve(wb, id, { ...m, method: 'evolutionary', constraints: [...m.constraints, { ref: 'B1:B2', op: '<=', rhs: '5' }, { ref: 'B1:B2', op: '>=', rhs: '-5' }], options: { seed: 3, maxTimeSec: 20 } })
  assert.ok(ev.ok, ev.message)
  close(ev.objective, 0.5, 1e-3)
  // Value Of: x² + (−2)² = 5  →  x = 1
  const vo = await solve(wb, id, { objective: 'B3', goal: 'value', targetValue: 5, variables: 'B1', constraints: [], nonNegative: true, method: 'grg' })
  assert.ok(vo.ok, vo.message)
  close(vo.x![0], 1, 1e-4)
  // …and an impossible target is reported as infeasible (x² + 4 can never be 2)
  const imp = await solve(wb, id, { objective: 'B3', goal: 'value', targetValue: 2, variables: 'B1', constraints: [], nonNegative: true, method: 'grg' })
  assert.equal(imp.ok, false)
})

test('Solver validates its inputs', async () => {
  const { wb, id } = productMix()
  assert.match((await solve(wb, id, { ...mixModel('simplex'), objective: 'A1' })).message, /formula/)
  assert.match((await solve(wb, id, { ...mixModel('simplex'), variables: 'C4' })).message, /formula/)
  assert.match((await solve(wb, id, { ...mixModel('simplex'), constraints: [{ ref: 'B6', op: '<=', rhs: 'C1:C3' }] })).message, /right-hand side/)
  const inf = await solve(wb, id, { ...mixModel('simplex'), constraints: [...mixModel('simplex').constraints, { ref: 'B2', op: '>=', rhs: '10' }] })
  assert.match(inf.message, /could not find a feasible/)
})
