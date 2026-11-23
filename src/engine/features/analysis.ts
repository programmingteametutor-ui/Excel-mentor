/**
 * Analysis ToolPak — the data-analysis tools of Excel's Analysis ToolPak add-in.
 *
 * Every tool is a pure function from numbers to an output table that mirrors Excel's layout
 * (same row labels, same order). `writeTable` then places the table in an output range or on a
 * new worksheet as one undoable command. Like Excel, the tools write static values, except
 * Moving Average and Exponential Smoothing, which write formulas linked to the input.
 */
import type { RawCellContent } from 'hyperformula'
import type { Workbook } from '../workbook'
import { toEngineFormula } from '../formulaText'
import { cellKey, colToLetters, quoteSheetName } from '../../model/address'
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, type CellStyle, type ChartType, type Range } from '../../model/types'
import * as st from '../stats'
import { uid } from './common'

// ---------------------------------------------------------------- output tables
export type Val = string | number | boolean | null
export type OutCell = Val | { v: Val; b?: boolean; i?: boolean; fmt?: string; top?: boolean; bottom?: boolean; center?: boolean }
export interface OutTable {
  rows: OutCell[][]
  /** optional chart over part of the output (range relative to the table's top-left cell) */
  chart?: { type: ChartType; range: Range; title: string; xTitle?: string; yTitle?: string }
}

const NA = '=NA()'

export class TableBuilder {
  rows: OutCell[][] = []
  title(text: string) {
    this.rows.push([{ v: text, b: true }])
    return this
  }
  blank() {
    this.rows.push([])
    return this
  }
  /** Table header: italic, rule above and below (Excel's ToolPak look). */
  header(cells: Val[]) {
    this.rows.push(cells.map((v, j) => ({ v, i: true, top: true, bottom: true, center: j > 0 })))
    return this
  }
  row(cells: Val[], fmt?: string) {
    this.rows.push(cells.map((v) => (fmt && typeof v === 'number' ? { v, fmt } : v)))
    return this
  }
  /** Puts a rule under the last row. */
  end() {
    const last = this.rows[this.rows.length - 1]
    if (last) this.rows[this.rows.length - 1] = last.map((c) => (c !== null && typeof c === 'object' ? { ...c, bottom: true } : { v: c, bottom: true }))
    return this
  }
  build(chart?: OutTable['chart']): OutTable {
    return { rows: this.rows, chart }
  }
}

const finite = (x: number): Val => (Number.isFinite(x) ? x : NA)

// ---------------------------------------------------------------- reading input
export interface Series {
  name: string
  values: number[]
}

/**
 * Splits an input range into series (by columns or rows). With `labels`, the first cell of each
 * series is its name. Blank cells are skipped; any other non-number is an error (as in Excel).
 */
export function readSeries(wb: Workbook, sheetId: string, range: Range, by: 'cols' | 'rows', labels: boolean, kind = 'Column'): Series[] | string {
  const out: Series[] = []
  const count = by === 'cols' ? range.c2 - range.c1 + 1 : range.r2 - range.r1 + 1
  const len = by === 'cols' ? range.r2 - range.r1 + 1 : range.c2 - range.c1 + 1
  for (let s = 0; s < count; s++) {
    let name = `${kind} ${s + 1}`
    const values: number[] = []
    for (let i = 0; i < len; i++) {
      const r = by === 'cols' ? range.r1 + i : range.r1 + s
      const c = by === 'cols' ? range.c1 + s : range.c1 + i
      const v = wb.getValue(sheetId, r, c)
      if (labels && i === 0) {
        const t = wb.getDisplay(sheetId, r, c).text
        if (t) name = t
        continue
      }
      if (v === null || v === '' || v === undefined) continue
      if (typeof v !== 'number') return `Input range contains non-numeric data (${colToLetters(c)}${r + 1}).${!labels && i === 0 ? ' If the first row holds labels, tick “Labels in first row”.' : ''}`
      values.push(v)
    }
    out.push({ name, values })
  }
  return out
}

/** A single column/row of numbers (for t-tests, regression Y, histogram input). */
export function readVector(wb: Workbook, sheetId: string, range: Range, labels: boolean, fallback: string): Series | string {
  const by = range.c1 === range.c2 ? 'cols' : range.r1 === range.r2 ? 'rows' : null
  if (!by) return `${fallback} must be a single column or row.`
  const s = readSeries(wb, sheetId, range, by, labels, fallback)
  if (typeof s === 'string') return s
  return { ...s[0], name: labels ? s[0].name : fallback }
}

/** Numbers of a rectangular block, keeping blanks as null (for ANOVA two-factor, regression X). */
export function readMatrix(wb: Workbook, sheetId: string, range: Range): (number | null)[][] | string {
  const out: (number | null)[][] = []
  for (let r = range.r1; r <= range.r2; r++) {
    const row: (number | null)[] = []
    for (let c = range.c1; c <= range.c2; c++) {
      const v = wb.getValue(sheetId, r, c)
      if (v === null || v === '' || v === undefined) row.push(null)
      else if (typeof v === 'number') row.push(v)
      else return `Input range contains non-numeric data (${colToLetters(c)}${r + 1}).`
    }
    out.push(row)
  }
  return out
}

const need = (s: Series, n: number, what = 'values') => (s.values.length < n ? `“${s.name}” needs at least ${n} ${what}.` : null)

// ---------------------------------------------------------------- descriptive statistics
export interface DescriptiveOpts {
  summary: boolean
  confidence?: number // e.g. 0.95
  kthLargest?: number
  kthSmallest?: number
}

export function descriptiveStatistics(series: Series[], o: DescriptiveOpts): OutTable | string {
  if (!series.length) return 'Choose an input range.'
  for (const s of series) if (!s.values.length) return `“${s.name}” has no numbers.`
  const blocks: Val[][][] = series.map((s) => {
    const xs = s.values
    const n = xs.length
    const sorted = st.sorted(xs)
    const rows: Val[][] = []
    if (o.summary) {
      const sd = n > 1 ? st.stdev(xs) : NaN
      const md = st.mode(xs)
      rows.push(
        ['Mean', st.mean(xs)],
        ['Standard Error', finite(sd / Math.sqrt(n))],
        ['Median', st.median(xs)],
        ['Mode', md ?? NA],
        ['Standard Deviation', finite(sd)],
        ['Sample Variance', finite(n > 1 ? st.variance(xs) : NaN)],
        ['Kurtosis', finite(n > 3 ? st.kurtosis(xs) : NaN)],
        ['Skewness', finite(n > 2 ? st.skewness(xs) : NaN)],
        ['Range', sorted[n - 1] - sorted[0]],
        ['Minimum', sorted[0]],
        ['Maximum', sorted[n - 1]],
        ['Sum', st.sum(xs)],
        ['Count', n],
      )
    }
    if (o.kthLargest) rows.push([`Largest(${o.kthLargest})`, sorted[n - o.kthLargest] ?? NA])
    if (o.kthSmallest) rows.push([`Smallest(${o.kthSmallest})`, sorted[o.kthSmallest - 1] ?? NA])
    if (o.confidence) {
      const cl = n > 1 ? st.tInv2T(1 - o.confidence, n - 1) * (st.stdev(xs) / Math.sqrt(n)) : NaN
      rows.push([`Confidence Level(${(o.confidence * 100).toFixed(1)}%)`, finite(cl)])
    }
    return rows
  })
  if (!blocks[0].length) return 'Choose at least one of Summary statistics, Confidence level, Kth largest or Kth smallest.'
  const t = new TableBuilder()
  t.rows.push(series.flatMap((s) => [{ v: s.name, i: true, top: true, bottom: true, center: true }, { v: null, top: true, bottom: true }]))
  t.rows.push(series.flatMap(() => [null, null]))
  for (let i = 0; i < blocks[0].length; i++) t.rows.push(blocks.flatMap((b) => b[i]))
  t.end()
  return t.build()
}

// ---------------------------------------------------------------- histogram
export interface HistogramOpts {
  bins?: number[]
  pareto: boolean
  cumulative: boolean
  chart: boolean
  binLabel?: string
}

export function histogramTool(input: Series, o: HistogramOpts): OutTable | string {
  const xs = input.values
  if (!xs.length) return 'The input range has no numbers.'
  let bins = o.bins?.length ? [...o.bins].sort((a, b) => a - b) : []
  if (!bins.length) {
    // Excel: equal-width bins from the minimum to the maximum
    const lo = Math.min(...xs)
    const hi = Math.max(...xs)
    const k = Math.max(1, Math.ceil(Math.sqrt(xs.length)))
    const w = (hi - lo) / k
    bins = w > 0 ? Array.from({ length: k + 1 }, (_, i) => +(lo + i * w).toPrecision(12)) : [lo]
  }
  const counts = new Array(bins.length + 1).fill(0)
  for (const x of xs) {
    const i = bins.findIndex((b) => x <= b)
    counts[i < 0 ? bins.length : i]++
  }
  const labels: Val[] = [...bins, 'More']
  const n = xs.length
  const t = new TableBuilder()
  const head: Val[] = [o.binLabel || 'Bin', 'Frequency']
  if (o.cumulative) head.push('Cumulative %')
  const order = labels.map((_, i) => i).sort((a, b) => counts[b] - counts[a] || a - b)
  if (o.pareto) head.push(o.binLabel || 'Bin', 'Frequency', ...(o.cumulative ? ['Cumulative %'] : []))
  t.header(head)
  let cum = 0
  let cumP = 0
  labels.forEach((lab, i) => {
    cum += counts[i]
    const row: OutCell[] = [lab, counts[i]]
    if (o.cumulative) row.push({ v: cum / n, fmt: '0.00%' })
    if (o.pareto) {
      const j = order[i]
      cumP += counts[j]
      row.push(labels[j], counts[j])
      if (o.cumulative) row.push({ v: cumP / n, fmt: '0.00%' })
    }
    t.rows.push(row)
  })
  t.end()
  return t.build(o.chart ? { type: 'column', range: { r1: 0, c1: 0, r2: labels.length, c2: 1 }, title: 'Histogram', xTitle: o.binLabel || 'Bin', yTitle: 'Frequency' } : undefined)
}

// ---------------------------------------------------------------- correlation / covariance
function matrixTable(series: Series[], f: (a: number[], b: number[]) => number, what: string): OutTable | string {
  if (series.length < 2) return `${what} needs at least two columns (or rows) of data.`
  const n = series[0].values.length
  if (series.some((s) => s.values.length !== n)) return `Every ${what === 'Correlation' ? 'variable' : 'series'} must have the same number of values (blank cells are skipped).`
  if (n < 2) return 'Each series needs at least two values.'
  const t = new TableBuilder()
  t.header([null, ...series.map((s) => s.name)])
  series.forEach((a, i) => t.row([a.name, ...series.map((b, j) => (j <= i ? finite(f(a.values, b.values)) : null))]))
  t.end()
  return t.build()
}

export const correlationTool = (series: Series[]) =>
  matrixTable(
    series,
    (a, b) => {
      const ma = st.mean(a)
      const mb = st.mean(b)
      let sab = 0
      let saa = 0
      let sbb = 0
      a.forEach((x, i) => {
        sab += (x - ma) * (b[i] - mb)
        saa += (x - ma) ** 2
        sbb += (b[i] - mb) ** 2
      })
      return sab / Math.sqrt(saa * sbb)
    },
    'Correlation',
  )

/** Excel's Covariance tool uses the population covariance (COVARIANCE.P). */
export const covarianceTool = (series: Series[]) =>
  matrixTable(
    series,
    (a, b) => {
      const ma = st.mean(a)
      const mb = st.mean(b)
      return a.reduce((s, x, i) => s + (x - ma) * (b[i] - mb), 0) / a.length
    },
    'Covariance',
  )

// ---------------------------------------------------------------- tests of two samples
const pOne = (t: number, df: number) => st.tTwoTail(Math.abs(t), df) / 2

export function tTestPaired(a: Series, b: Series, hyp: number, alpha: number): OutTable | string {
  if (a.values.length !== b.values.length) return 'Paired t-Test needs the same number of values in both variables.'
  const e = need(a, 2) ?? need(b, 2)
  if (e) return e
  const d = a.values.map((x, i) => x - b.values[i])
  const n = d.length
  const df = n - 1
  const t = (st.mean(d) - hyp) / (st.stdev(d) / Math.sqrt(n))
  const r = correlationTool([a, b])
  const pearson = typeof r === 'string' ? NA : (r.rows[2][1] as Val)
  return new TableBuilder()
    .title('t-Test: Paired Two Sample for Means')
    .blank()
    .header([null, a.name, b.name])
    .row(['Mean', st.mean(a.values), st.mean(b.values)])
    .row(['Variance', st.variance(a.values), st.variance(b.values)])
    .row(['Observations', n, n])
    .row(['Pearson Correlation', pearson])
    .row(['Hypothesized Mean Difference', hyp])
    .row(['df', df])
    .row(['t Stat', finite(t)])
    .row(['P(T<=t) one-tail', finite(pOne(t, df))])
    .row(['t Critical one-tail', st.tInv(1 - alpha, df)])
    .row(['P(T<=t) two-tail', finite(2 * pOne(t, df))])
    .row(['t Critical two-tail', st.tInv2T(alpha, df)])
    .end()
    .build()
}

export function tTestTwoSample(a: Series, b: Series, hyp: number, alpha: number, equalVariances: boolean): OutTable | string {
  const e = need(a, 2) ?? need(b, 2)
  if (e) return e
  const n1 = a.values.length
  const n2 = b.values.length
  const m1 = st.mean(a.values)
  const m2 = st.mean(b.values)
  const v1 = st.variance(a.values)
  const v2 = st.variance(b.values)
  const t = new TableBuilder().title(`t-Test: Two-Sample Assuming ${equalVariances ? 'Equal' : 'Unequal'} Variances`).blank().header([null, a.name, b.name])
  t.row(['Mean', m1, m2]).row(['Variance', v1, v2]).row(['Observations', n1, n2])
  let stat: number
  let df: number
  if (equalVariances) {
    df = n1 + n2 - 2
    const pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / df
    stat = (m1 - m2 - hyp) / Math.sqrt(pooled * (1 / n1 + 1 / n2))
    t.row(['Pooled Variance', pooled])
  } else {
    const q1 = v1 / n1
    const q2 = v2 / n2
    // Welch–Satterthwaite, rounded to a whole number as Excel's tool does
    df = Math.round((q1 + q2) ** 2 / (q1 ** 2 / (n1 - 1) + q2 ** 2 / (n2 - 1)))
    stat = (m1 - m2 - hyp) / Math.sqrt(q1 + q2)
  }
  return t
    .row(['Hypothesized Mean Difference', hyp])
    .row(['df', df])
    .row(['t Stat', finite(stat)])
    .row(['P(T<=t) one-tail', finite(pOne(stat, df))])
    .row(['t Critical one-tail', st.tInv(1 - alpha, df)])
    .row(['P(T<=t) two-tail', finite(2 * pOne(stat, df))])
    .row(['t Critical two-tail', st.tInv2T(alpha, df)])
    .end()
    .build()
}

export function zTestTwoSample(a: Series, b: Series, hyp: number, var1: number, var2: number, alpha: number): OutTable | string {
  const e = need(a, 1) ?? need(b, 1)
  if (e) return e
  if (!(var1 > 0) || !(var2 > 0)) return 'Enter the known variance of each variable (a number greater than 0).'
  const n1 = a.values.length
  const n2 = b.values.length
  const m1 = st.mean(a.values)
  const m2 = st.mean(b.values)
  const z = (m1 - m2 - hyp) / Math.sqrt(var1 / n1 + var2 / n2)
  const p1 = 1 - st.normCdf(Math.abs(z))
  return new TableBuilder()
    .title('z-Test: Two Sample for Means')
    .blank()
    .header([null, a.name, b.name])
    .row(['Mean', m1, m2])
    .row(['Known Variance', var1, var2])
    .row(['Observations', n1, n2])
    .row(['Hypothesized Mean Difference', hyp])
    .row(['z', z])
    .row(['P(Z<=z) one-tail', p1])
    .row(['z Critical one-tail', st.normInv(1 - alpha)])
    .row(['P(Z<=z) two-tail', 2 * p1])
    .row(['z Critical two-tail', st.normInv(1 - alpha / 2)])
    .end()
    .build()
}

export function fTestTwoSample(a: Series, b: Series, alpha: number): OutTable | string {
  const e = need(a, 2) ?? need(b, 2)
  if (e) return e
  const v1 = st.variance(a.values)
  const v2 = st.variance(b.values)
  const d1 = a.values.length - 1
  const d2 = b.values.length - 1
  const F = v1 / v2
  // Excel reports the tail on the side F falls (left tail when F < 1)
  const left = F < 1
  return new TableBuilder()
    .title('F-Test Two-Sample for Variances')
    .blank()
    .header([null, a.name, b.name])
    .row(['Mean', st.mean(a.values), st.mean(b.values)])
    .row(['Variance', v1, v2])
    .row(['Observations', d1 + 1, d2 + 1])
    .row(['df', d1, d2])
    .row(['F', finite(F)])
    .row(['P(F<=f) one-tail', finite(left ? st.fCdf(F, d1, d2) : st.fRight(F, d1, d2))])
    .row(['F Critical one-tail', left ? st.fInv(alpha, d1, d2) : st.fInv(1 - alpha, d1, d2)])
    .end()
    .build()
}

// ---------------------------------------------------------------- ANOVA
function anovaRows(t: TableBuilder, rows: { name: string; ss: number; df: number; test: boolean }[], errMs: number, errDf: number, alpha: number) {
  t.header(['Source of Variation', 'SS', 'df', 'MS', 'F', 'P-value', 'F crit'])
  for (const r of rows) {
    if (r.name === 'Total') {
      t.blank()
      t.row(['Total', r.ss, r.df])
      continue
    }
    const ms = r.ss / r.df
    if (r.test) {
      const F = ms / errMs
      t.row([r.name, r.ss, r.df, finite(ms), finite(F), finite(st.fRight(F, r.df, errDf)), st.fInv(1 - alpha, r.df, errDf)])
    } else t.row([r.name, r.ss, r.df, finite(ms)])
  }
  t.end()
}

const summaryRow = (name: string, xs: number[]): Val[] => [name, xs.length, st.sum(xs), xs.length ? st.mean(xs) : NA, xs.length > 1 ? st.variance(xs) : NA]

export function anovaSingleFactor(groups: Series[], alpha: number): OutTable | string {
  if (groups.length < 2) return 'ANOVA needs at least two groups.'
  for (const g of groups) if (!g.values.length) return `“${g.name}” has no numbers.`
  const all = groups.flatMap((g) => g.values)
  const N = all.length
  const k = groups.length
  if (N <= k) return 'There are not enough values: each group needs more than one value on average.'
  const G = st.mean(all)
  const ssb = groups.reduce((a, g) => a + g.values.length * (st.mean(g.values) - G) ** 2, 0)
  const ssw = groups.reduce((a, g) => a + st.devsq(g.values), 0)
  const t = new TableBuilder().title('Anova: Single Factor').blank().title('SUMMARY').header(['Groups', 'Count', 'Sum', 'Average', 'Variance'])
  for (const g of groups) t.row(summaryRow(g.name, g.values))
  t.end().blank().blank().title('ANOVA')
  anovaRows(t, [{ name: 'Between Groups', ss: ssb, df: k - 1, test: true }, { name: 'Within Groups', ss: ssw, df: N - k, test: false }, { name: 'Total', ss: ssb + ssw, df: N - 1, test: false }], ssw / (N - k), N - k, alpha)
  return t.build()
}

/** Two-factor ANOVA without replication. `data` is rows × columns (no blanks allowed). */
export function anovaTwoFactorNoRep(data: number[][], rowNames: string[], colNames: string[], alpha: number): OutTable | string {
  const r = data.length
  const c = data[0]?.length ?? 0
  if (r < 2 || c < 2) return 'Two-factor ANOVA needs at least two rows and two columns of numbers.'
  const all = data.flat()
  const G = st.mean(all)
  const rowMeans = data.map((row) => st.mean(row))
  const cols = data[0].map((_, j) => data.map((row) => row[j]))
  const colMeans = cols.map((col) => st.mean(col))
  const ssr = c * rowMeans.reduce((a, m) => a + (m - G) ** 2, 0)
  const ssc = r * colMeans.reduce((a, m) => a + (m - G) ** 2, 0)
  const sst = st.devsq(all)
  const sse = sst - ssr - ssc
  const dfe = (r - 1) * (c - 1)
  const t = new TableBuilder().title('Anova: Two-Factor Without Replication').blank().header(['SUMMARY', 'Count', 'Sum', 'Average', 'Variance'])
  data.forEach((row, i) => t.row(summaryRow(rowNames[i], row)))
  t.blank()
  cols.forEach((col, j) => t.row(summaryRow(colNames[j], col)))
  t.end().blank().blank().title('ANOVA')
  anovaRows(t, [{ name: 'Rows', ss: ssr, df: r - 1, test: true }, { name: 'Columns', ss: ssc, df: c - 1, test: true }, { name: 'Error', ss: sse, df: dfe, test: false }, { name: 'Total', ss: sst, df: r * c - 1, test: false }], sse / dfe, dfe, alpha)
  return t.build()
}

/** Two-factor ANOVA with replication: `data` holds a groups of `perSample` rows each. */
export function anovaTwoFactorRep(data: number[][], perSample: number, groupNames: string[], colNames: string[], alpha: number): OutTable | string {
  const b = data[0]?.length ?? 0
  if (perSample < 2) return 'Rows per sample must be at least 2.'
  if (data.length % perSample) return `The number of data rows (${data.length}) must be a multiple of the rows per sample (${perSample}).`
  const a = data.length / perSample
  if (a < 2 || b < 2) return 'Two-factor ANOVA needs at least two samples and two columns.'
  const cell = (i: number, j: number) => data.slice(i * perSample, (i + 1) * perSample).map((row) => row[j])
  const all = data.flat()
  const G = st.mean(all)
  const groupVals = Array.from({ length: a }, (_, i) => data.slice(i * perSample, (i + 1) * perSample).flat())
  const colVals = Array.from({ length: b }, (_, j) => data.map((row) => row[j]))
  const ssA = b * perSample * groupVals.reduce((s, g) => s + (st.mean(g) - G) ** 2, 0)
  const ssB = a * perSample * colVals.reduce((s, g) => s + (st.mean(g) - G) ** 2, 0)
  let ssW = 0
  for (let i = 0; i < a; i++) for (let j = 0; j < b; j++) ssW += st.devsq(cell(i, j))
  const ssT = st.devsq(all)
  const ssAB = ssT - ssA - ssB - ssW
  const dfW = a * b * (perSample - 1)
  const t = new TableBuilder().title('Anova: Two-Factor With Replication').blank()
  const block = (title: string, vals: (j: number) => number[], total: number[]) => {
    t.header(['SUMMARY', ...colNames, 'Total'])
    t.title(title)
    const cols = colNames.map((_, j) => vals(j))
    t.row(['Count', ...cols.map((v) => v.length), total.length])
    t.row(['Sum', ...cols.map((v) => st.sum(v)), st.sum(total)])
    t.row(['Average', ...cols.map((v) => st.mean(v)), st.mean(total)])
    t.row(['Variance', ...cols.map((v) => finite(st.variance(v))), finite(st.variance(total))])
    t.blank()
  }
  for (let i = 0; i < a; i++) block(groupNames[i], (j) => cell(i, j), groupVals[i])
  t.title('Total')
  const cols = colNames.map((_, j) => colVals[j])
  t.row(['Count', ...cols.map((v) => v.length)])
  t.row(['Sum', ...cols.map((v) => st.sum(v))])
  t.row(['Average', ...cols.map((v) => st.mean(v))])
  t.row(['Variance', ...cols.map((v) => finite(st.variance(v)))])
  t.end().blank().blank().title('ANOVA')
  anovaRows(
    t,
    [
      { name: 'Sample', ss: ssA, df: a - 1, test: true },
      { name: 'Columns', ss: ssB, df: b - 1, test: true },
      { name: 'Interaction', ss: ssAB, df: (a - 1) * (b - 1), test: true },
      { name: 'Within', ss: ssW, df: dfW, test: false },
      { name: 'Total', ss: ssT, df: a * b * perSample - 1, test: false },
    ],
    ssW / dfW,
    dfW,
    alpha,
  )
  return t.build()
}

// ---------------------------------------------------------------- regression
export interface RegressionOpts {
  constantZero: boolean
  confidence: number // e.g. 0.95
  residuals: boolean
  standardized: boolean
  normalProbability: boolean
  lineFitChart: boolean
}

export function regressionTool(y: Series, xs: Series[], o: RegressionOpts): OutTable | string {
  const n = y.values.length
  if (xs.some((x) => x.values.length !== n)) return 'Input Y and Input X ranges must have the same number of values (and no blanks).'
  if (xs.length > 16) return 'Input X can have at most 16 columns.'
  if (n <= xs.length + (o.constantZero ? 0 : 1)) return 'There are not enough observations for this many X variables.'
  const res = st.ols(
    y.values,
    xs.map((x) => x.values),
    !o.constantZero,
  )
  if (!res) return 'The regression could not be calculated (the X columns may be identical or constant).'
  const k = xs.length
  const t95 = st.tInv2T(0.05, res.dfRes)
  const tc = st.tInv2T(1 - o.confidence, res.dfRes)
  const tb = new TableBuilder().title('SUMMARY OUTPUT').blank()
  tb.header(['Regression Statistics', null])
  tb.row(['Multiple R', res.multipleR]).row(['R Square', res.r2]).row(['Adjusted R Square', res.adjR2]).row(['Standard Error', finite(res.seY)]).row(['Observations', n]).end()
  tb.blank().title('ANOVA')
  tb.header([null, 'df', 'SS', 'MS', 'F', 'Significance F'])
  const msReg = res.ssReg / res.dfReg
  const msRes = res.ssRes / res.dfRes
  tb.row(['Regression', res.dfReg, res.ssReg, finite(msReg), finite(res.F), finite(res.pF)])
  tb.row(['Residual', res.dfRes, res.ssRes, finite(msRes)])
  tb.row(['Total', res.dfReg + res.dfRes, res.ssReg + res.ssRes]).end()
  tb.blank()
  const pct = (o.confidence * 100).toFixed(1)
  const extra = Math.abs(o.confidence - 0.95) > 1e-9
  tb.header([null, 'Coefficients', 'Standard Error', 't Stat', 'P-value', 'Lower 95%', 'Upper 95%', ...(extra ? [`Lower ${pct}%`, `Upper ${pct}%`] : [])])
  const names = ['Intercept', ...xs.map((x) => x.name)]
  names.forEach((name, i) => {
    if (i === 0 && o.constantZero) {
      tb.row(['Intercept', 0, NA, NA, NA, NA, NA, ...(extra ? [NA, NA] : [])])
      return
    }
    const b = res.coef[i]
    const se = res.se[i]
    const ok = Number.isFinite(se) && se > 0
    tb.row([name, b, ok ? se : 0, ok ? res.t[i] : NA, ok ? res.p[i] : NA, b - t95 * se, b + t95 * se, ...(extra ? [b - tc * se, b + tc * se] : [])])
  })
  tb.end()
  let chart: OutTable['chart']
  if (o.residuals || o.standardized || o.lineFitChart) {
    tb.blank().blank().title('RESIDUAL OUTPUT').blank()
    const head: Val[] = ['Observation', `Predicted ${y.name}`, 'Residuals']
    if (o.standardized) head.push('Standard Residuals')
    if (o.lineFitChart) head.push(xs[0].name, y.name)
    const start = tb.rows.length
    tb.header(head)
    // Excel standardises residuals by their standard deviation (n − 1 denominator)
    const sdRes = Math.sqrt(res.residuals.reduce((a, r) => a + r * r, 0) / (n - 1))
    for (let i = 0; i < n; i++) {
      const row: Val[] = [i + 1, res.fitted[i], res.residuals[i]]
      if (o.standardized) row.push(res.residuals[i] / sdRes)
      if (o.lineFitChart) row.push(xs[0].values[i], y.values[i])
      tb.row(row)
    }
    tb.end()
    if (o.lineFitChart) {
      // a scatter of Y and predicted Y against the first X (Excel's “Line Fit Plot”)
      chart = { type: 'scatter', range: { r1: start, c1: o.standardized ? 4 : 3, r2: start + n, c2: o.standardized ? 5 : 4 }, title: `${xs[0].name} Line Fit Plot`, xTitle: xs[0].name, yTitle: y.name }
    }
  }
  if (o.normalProbability) {
    tb.blank().blank().title('PROBABILITY OUTPUT').blank()
    tb.header(['Percentile', y.name])
    st.sorted(y.values).forEach((v, i) => tb.row([(100 * (2 * i + 1)) / (2 * n), v]))
    tb.end()
  }
  void k
  return tb.build(chart)
}

// ---------------------------------------------------------------- time series (formulas)
export interface FormulaCtx {
  /** reference to an input cell (sheet-qualified when the output is on another sheet) */
  inRef: (r: number, c: number) => string
  /** reference to an output cell, relative to the output's top-left corner */
  outRef: (i: number, j: number) => string
}

/** Moving Average: one forecast per input value, linked to the input with =AVERAGE(...). */
export function movingAverage(input: Range, labels: boolean, interval: number, stdErrors: boolean, chart: boolean, ctx: FormulaCtx, name = 'Actual'): OutTable | string {
  const vertical = input.c1 === input.c2
  if (!vertical && input.r1 !== input.r2) return 'Input range must be a single column or row.'
  if (!(interval >= 2)) return 'Interval must be a whole number of at least 2.'
  const cells: { r: number; c: number }[] = []
  for (let r = input.r1; r <= input.r2; r++) for (let c = input.c1; c <= input.c2; c++) cells.push({ r, c })
  if (labels) cells.shift()
  const n = cells.length
  if (n < interval) return `Interval (${interval}) is larger than the number of values (${n}).`
  const rows: OutCell[][] = []
  const off = chart ? 1 : 0 // a header row when charting
  const fcCol = chart ? 1 : 0
  if (chart) rows.push([{ v: name, i: true, bottom: true }, { v: 'Forecast', i: true, bottom: true }, ...(stdErrors ? [{ v: 'Standard Error', i: true, bottom: true }] : [])])
  const span = (a: number, b: number) => `${ctx.inRef(cells[a].r, cells[a].c)}:${ctx.inRef(cells[b].r, cells[b].c)}`
  for (let i = 0; i < n; i++) {
    const row: OutCell[] = []
    if (chart) row.push(`=${ctx.inRef(cells[i].r, cells[i].c)}`)
    row.push(i < interval - 1 ? NA : `=AVERAGE(${span(i - interval + 1, i)})`)
    if (stdErrors) {
      const ok = i >= 2 * interval - 2
      row.push(ok ? `=SQRT(SUMXMY2(${span(i - interval + 1, i)},${ctx.outRef(off + i - interval + 1, fcCol)}:${ctx.outRef(off + i, fcCol)})/${interval})` : NA)
    }
    rows.push(row)
  }
  return { rows, chart: chart ? { type: 'line', range: { r1: 0, c1: 0, r2: n, c2: 1 }, title: 'Moving Average', xTitle: 'Data Point', yTitle: 'Value' } : undefined }
}

/** Exponential Smoothing: F(t) = α·A(t−1) + (1−α)·F(t−1), with α = 1 − damping factor. */
export function exponentialSmoothing(input: Range, labels: boolean, damping: number, stdErrors: boolean, chart: boolean, ctx: FormulaCtx, name = 'Actual'): OutTable | string {
  if (input.c1 !== input.c2 && input.r1 !== input.r2) return 'Input range must be a single column or row.'
  if (!(damping >= 0 && damping <= 1)) return 'Damping factor must be between 0 and 1.'
  const cells: { r: number; c: number }[] = []
  for (let r = input.r1; r <= input.r2; r++) for (let c = input.c1; c <= input.c2; c++) cells.push({ r, c })
  if (labels) cells.shift()
  const n = cells.length
  if (n < 2) return 'Exponential Smoothing needs at least two values.'
  const alpha = +(1 - damping).toPrecision(12)
  const off = chart ? 1 : 0
  const fc = chart ? 1 : 0
  const rows: OutCell[][] = []
  if (chart) rows.push([{ v: name, i: true, bottom: true }, { v: 'Forecast', i: true, bottom: true }, ...(stdErrors ? [{ v: 'Standard Error', i: true, bottom: true }] : [])])
  const inR = (i: number) => ctx.inRef(cells[i].r, cells[i].c)
  for (let i = 0; i < n; i++) {
    const row: OutCell[] = []
    if (chart) row.push(`=${inR(i)}`)
    row.push(i === 0 ? NA : i === 1 ? `=${inR(0)}` : `=${alpha}*${inR(i - 1)}+${damping}*${ctx.outRef(off + i - 1, fc)}`)
    if (stdErrors) row.push(i >= 3 ? `=SQRT(SUMXMY2(${inR(i - 2)}:${inR(i)},${ctx.outRef(off + i - 2, fc)}:${ctx.outRef(off + i, fc)})/3)` : NA)
    rows.push(row)
  }
  return { rows, chart: chart ? { type: 'line', range: { r1: 0, c1: 0, r2: n, c2: 1 }, title: 'Exponential Smoothing', xTitle: 'Data Point', yTitle: 'Value' } : undefined }
}

// ---------------------------------------------------------------- rank, sampling, random numbers, Fourier
export function rankAndPercentile(series: Series[]): OutTable | string {
  if (!series.length) return 'Choose an input range.'
  const blocks = series.map((s) => {
    const items = s.values.map((v, i) => ({ v, i }))
    items.sort((a, b) => b.v - a.v || a.i - b.i)
    return items.map((it) => [it.i + 1, it.v, s.values.filter((x) => x > it.v).length + 1, { v: st.percentRank(s.values, it.v, 3, false) ?? 0, fmt: '0.00%' }] as OutCell[])
  })
  const rows: OutCell[][] = [series.flatMap((s) => ['Point', s.name, 'Rank', 'Percent'].map((v, j) => ({ v, i: true, top: true, bottom: true, center: j > 0 })))]
  const len = Math.max(...blocks.map((b) => b.length))
  for (let i = 0; i < len; i++) rows.push(blocks.flatMap((b) => b[i] ?? [null, null, null, null]))
  return { rows }
}

export function samplingTool(input: Series, method: 'periodic' | 'random', amount: number, rand: () => number = Math.random): OutTable | string {
  const xs = input.values
  if (!xs.length) return 'The input range has no numbers.'
  if (!(amount >= 1)) return method === 'periodic' ? 'Period must be at least 1.' : 'Number of samples must be at least 1.'
  const out: number[] = []
  if (method === 'periodic') {
    if (amount > xs.length) return `Period (${amount}) is larger than the number of values (${xs.length}).`
    for (let i = amount - 1; i < xs.length; i += amount) out.push(xs[i])
  } else for (let i = 0; i < amount; i++) out.push(xs[Math.floor(rand() * xs.length)]) // with replacement, like Excel
  return { rows: out.map((v) => [v]) }
}

export type RandomDist =
  | { kind: 'uniform'; from: number; to: number }
  | { kind: 'normal'; mean: number; sd: number }
  | { kind: 'bernoulli'; p: number }
  | { kind: 'binomial'; p: number; trials: number }
  | { kind: 'poisson'; lambda: number }
  | { kind: 'patterned'; from: number; to: number; step: number; repeatEach: number; repeatSeq: number }
  | { kind: 'discrete'; values: number[]; probs: number[] }

export function randomNumbers(vars: number, count: number, dist: RandomDist, seed?: number): OutTable | string {
  const rand = st.seededRandom(seed)
  if (dist.kind === 'patterned') {
    if (!(dist.step > 0) || dist.to < dist.from) return 'Patterned: “to” must be at least “from”, and the step must be positive.'
    const seq: number[] = []
    for (let r = 0; r < Math.max(1, dist.repeatSeq); r++)
      for (let x = dist.from; x <= dist.to + 1e-9; x += dist.step) for (let k = 0; k < Math.max(1, dist.repeatEach); k++) seq.push(+x.toPrecision(12))
    const v = Math.max(1, vars || 1)
    return { rows: seq.map((x) => new Array(v).fill(x)) }
  }
  if (!(vars >= 1) || !(count >= 1)) return 'Number of variables and number of random numbers must be at least 1.'
  if (vars * count > 200_000) return 'That is too many numbers at once (limit 200,000).'
  let one: () => number
  switch (dist.kind) {
    case 'uniform':
      one = () => dist.from + rand() * (dist.to - dist.from)
      break
    case 'normal':
      if (!(dist.sd >= 0)) return 'Standard deviation must not be negative.'
      one = () => st.randNormal(rand, dist.mean, dist.sd)
      break
    case 'bernoulli':
      if (!(dist.p >= 0 && dist.p <= 1)) return 'p must be between 0 and 1.'
      one = () => (rand() < dist.p ? 1 : 0)
      break
    case 'binomial':
      if (!(dist.p >= 0 && dist.p <= 1) || !(dist.trials >= 1)) return 'Binomial needs 0 ≤ p ≤ 1 and at least one trial.'
      one = () => st.randBinomial(rand, Math.floor(dist.trials), dist.p)
      break
    case 'poisson':
      if (!(dist.lambda > 0)) return 'Lambda must be greater than 0.'
      one = () => st.randPoisson(rand, dist.lambda)
      break
    case 'discrete': {
      const total = st.sum(dist.probs)
      if (!dist.values.length || dist.values.length !== dist.probs.length || dist.probs.some((p) => p < 0) || Math.abs(total - 1) > 1e-6)
        return 'Discrete needs a two-column range of values and probabilities that add up to 1.'
      one = () => {
        let u = rand()
        for (let i = 0; i < dist.values.length; i++) {
          u -= dist.probs[i]
          if (u < 0) return dist.values[i]
        }
        return dist.values[dist.values.length - 1]
      }
      break
    }
  }
  return { rows: Array.from({ length: count }, () => Array.from({ length: vars }, one)) }
}

/** Complex number as Excel's COMPLEX text ("3+4i"). */
function complexText(re: number, im: number): string {
  const f = (x: number) => String(+x.toPrecision(15))
  const r = Math.abs(re) < 1e-12 ? 0 : re
  const i = Math.abs(im) < 1e-12 ? 0 : im
  if (i === 0) return f(r)
  const imag = i === 1 ? 'i' : i === -1 ? '-i' : `${f(i)}i`
  if (r === 0) return imag
  return `${f(r)}${i > 0 ? '+' : ''}${imag}`
}

/** Fourier Analysis (FFT); input length must be a power of 2, as in Excel. */
export function fourierTool(xs: number[], inverse: boolean): OutTable | string {
  const n = xs.length
  if (n < 2 || n > 4096 || n & (n - 1)) return `Fourier Analysis needs 2, 4, 8, … up to 4096 values (you have ${n}).`
  const re = [...xs]
  const im = new Array(n).fill(0)
  // iterative radix-2 Cooley–Tukey
  for (let i = 1, j = 0; i < n; i++) {
    let bit = n >> 1
    for (; j & bit; bit >>= 1) j ^= bit
    j ^= bit
    if (i < j) {
      ;[re[i], re[j]] = [re[j], re[i]]
      ;[im[i], im[j]] = [im[j], im[i]]
    }
  }
  for (let len = 2; len <= n; len <<= 1) {
    const ang = ((inverse ? 2 : -2) * Math.PI) / len
    for (let i = 0; i < n; i += len)
      for (let k = 0; k < len / 2; k++) {
        const wr = Math.cos(ang * k)
        const wi = Math.sin(ang * k)
        const ur = re[i + k]
        const ui = im[i + k]
        const vr = re[i + k + len / 2] * wr - im[i + k + len / 2] * wi
        const vi = re[i + k + len / 2] * wi + im[i + k + len / 2] * wr
        re[i + k] = ur + vr
        im[i + k] = ui + vi
        re[i + k + len / 2] = ur - vr
        im[i + k + len / 2] = ui - vi
      }
  }
  return { rows: re.map((r, i) => [complexText(inverse ? r / n : r, inverse ? im[i] / n : im[i])]) }
}

// ---------------------------------------------------------------- writing output
export type OutputTarget = { kind: 'range'; sheetId: string; r: number; c: number } | { kind: 'newSheet'; name?: string }

/** "Analysis1", "Analysis2", … — the first unused name for `base` (Excel numbers reports from 1). */
export function uniqueSheetName(wb: Workbook, base: string): string {
  for (let k = 1; ; k++) if (!wb.sheetByName(`${base}${k}`)) return `${base}${k}`
}

/** Resolves the output sheet/corner (creating a new worksheet when asked). */
export function resolveTarget(wb: Workbook, target: OutputTarget): { sheetId: string; r: number; c: number } | string {
  if (target.kind === 'range') return { sheetId: target.sheetId, r: target.r, c: target.c }
  const name = target.name?.trim() || uniqueSheetName(wb, 'Analysis')
  if (wb.sheetByName(name)) return `A sheet named “${name}” already exists. Choose another name.`
  try {
    return { sheetId: wb.addSheet(name), r: 0, c: 0 }
  } catch (e) {
    return (e as Error).message
  }
}

/** Context for formula-writing tools: input references qualified when needed. */
export function formulaCtx(wb: Workbook, inputSheetId: string, out: { sheetId: string; r: number; c: number }): FormulaCtx {
  const prefix = inputSheetId === out.sheetId ? '' : `${quoteSheetName(wb.meta(inputSheetId).name)}!`
  return {
    inRef: (r, c) => `${prefix}$${colToLetters(c)}$${r + 1}`,
    outRef: (i, j) => `${colToLetters(out.c + j)}${out.r + i + 1}`,
  }
}

export function tableSize(t: OutTable) {
  return { rows: t.rows.length, cols: Math.max(1, ...t.rows.map((r) => r.length)) }
}

/** True when the output area already holds data (Excel asks before overwriting). */
export function outputOverlapsData(wb: Workbook, out: { sheetId: string; r: number; c: number }, t: OutTable): boolean {
  const { rows, cols } = tableSize(t)
  for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) if (!wb.isCellEmpty(out.sheetId, out.r + i, out.c + j)) return true
  return false
}

/**
 * Writes a table (values, formulas and ToolPak styling) as one undoable command, widens
 * columns on a fresh sheet, and adds the optional chart beside it. Returns the written range.
 */
export function writeTable(wb: Workbook, t: OutTable, out: { sheetId: string; r: number; c: number }, label: string, freshSheet = false): Range {
  const { rows, cols } = tableSize(t)
  const hfId = wb.hfId(out.sheetId)
  const content: RawCellContent[][] = []
  const styles: [string, CellStyle | undefined][] = []
  const widths = new Array(cols).fill(0)
  t.rows.forEach((row, i) => {
    const line: RawCellContent[] = []
    for (let j = 0; j < cols; j++) {
      const cell = row[j] ?? null
      const obj = cell !== null && typeof cell === 'object' ? cell : null
      const v = obj ? obj.v : (cell as Val)
      line.push(typeof v === 'string' && v.startsWith('=') ? toEngineFormula(v) : (v as RawCellContent))
      const s: CellStyle = {}
      if (obj?.b) s.bold = true
      if (obj?.i) s.italic = true
      if (obj?.fmt) s.numFmt = obj.fmt
      if (obj?.center) s.hAlign = 'center'
      if (obj?.top || obj?.bottom) s.borders = { ...(obj.top ? { top: { style: 'thin', color: '#000000' } } : {}), ...(obj.bottom ? { bottom: { style: 'thin', color: '#000000' } } : {}) }
      styles.push([cellKey(out.r + i, out.c + j), Object.keys(s).length ? s : undefined])
      // single-cell rows (titles, notes) overflow into the next columns, so they don't set widths
      const len = row.length === 1 ? 0 : typeof v === 'number' ? Math.min(12, String(+v.toPrecision(10)).length) : typeof v === 'string' && !v.startsWith('=') ? v.length : 0
      widths[j] = Math.max(widths[j], len)
    }
    content.push(line)
  })
  wb.run({
    label,
    dirty: [out.sheetId],
    metaScope: [out.sheetId],
    hfOps: () => wb.hf.setCellContents({ sheet: hfId, row: out.r, col: out.c }, content),
    metaOps: () => {
      const m = wb.meta(out.sheetId)
      const next = { ...m.styles }
      for (const [k, s] of styles) {
        if (s) next[k] = { ...next[k], ...s }
        else if (freshSheet) delete next[k]
      }
      m.styles = next
      if (freshSheet) {
        const cw = { ...m.colWidths }
        widths.forEach((w, j) => {
          const px = Math.min(260, Math.max(DEFAULT_COL_WIDTH, Math.round(w * 7.2 + 14)))
          if (px > DEFAULT_COL_WIDTH) cw[out.c + j] = px
        })
        m.colWidths = cw
      }
      if (t.chart) {
        const m2 = wb.meta(out.sheetId)
        const colX = (c: number) => {
          let x = 0
          for (let k = 0; k < c; k++) x += m2.colWidths?.[k] ?? DEFAULT_COL_WIDTH
          return x
        }
        const ch = t.chart
        m2.charts = [
          ...(m2.charts ?? []),
          {
            id: uid('ch'),
            type: ch.type,
            source: { sheetId: out.sheetId, range: { r1: out.r + ch.range.r1, c1: out.c + ch.range.c1, r2: out.r + ch.range.r2, c2: out.c + ch.range.c2 } },
            seriesIn: 'cols',
            firstRowHeader: true,
            firstColLabels: true,
            title: ch.title,
            xTitle: ch.xTitle ?? '',
            yTitle: ch.yTitle ?? '',
            legend: 'bottom',
            dataLabels: false,
            gridlines: true,
            stacked: false,
            colors: ['#2f6fd6', '#e8710a', '#188038', '#d93025', '#7c3fc4', '#0b8a8a'],
            x: colX(out.c + cols) + 24,
            y: out.r * DEFAULT_ROW_HEIGHT + 8,
            w: 460,
            h: 280,
          },
        ]
      }
    },
  })
  return { r1: out.r, c1: out.c, r2: out.r + rows - 1, c2: out.c + cols - 1 }
}
