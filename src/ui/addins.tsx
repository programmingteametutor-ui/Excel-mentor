/*
 * Add-ins: the Add-ins manager, the Analysis ToolPak (Data Analysis) and Solver dialogs.
 *
 * Range boxes work like Excel's: type a reference, or press the select button, which collapses the
 * dialog so you can drag across cells on the sheet, then press Done (or Enter).
 */
import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import type { Workbook } from '../engine/workbook'
import type { Range, SolverConstraint, SolverModel, SolverOp } from '../model/types'
import { parseRangeA1, quoteSheetName } from '../model/address'
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
  outputOverlapsData,
  randomNumbers,
  rankAndPercentile,
  readMatrix,
  readSeries,
  readVector,
  regressionTool,
  resolveTarget,
  samplingTool,
  tTestPaired,
  tTestTwoSample,
  writeTable,
  zTestTwoSample,
  uniqueSheetName,
  type FormulaCtx,
  type OutTable,
  type RandomDist,
  type Series,
} from '../engine/features/analysis'
import { answerReport, applySolution, saveModel, sensitivityReport, solve, type SolverResult } from '../engine/features/solver'
import { seededRandom } from '../engine/stats'
import { Button, Check, Dialog, Field, inputClass } from './components/ui'
import { Icon } from './icons'
import { cn } from './lib/utils'

// ---------------------------------------------------------------- add-in switches
export interface AddinState {
  analysis: boolean
  solver: boolean
}
const KEY = 'excelpro.addins'
const EVT = 'excelpro:addins'
export function readAddins(): AddinState {
  try {
    const v = JSON.parse(localStorage.getItem(KEY) ?? 'null')
    if (v && typeof v === 'object') return { analysis: v.analysis !== false, solver: v.solver !== false }
  } catch {
    /* storage unavailable */
  }
  return { analysis: true, solver: true }
}
export function useAddins(): [AddinState, (s: AddinState) => void] {
  const [s, setS] = useState(readAddins)
  useEffect(() => {
    const h = () => setS(readAddins())
    window.addEventListener(EVT, h)
    return () => window.removeEventListener(EVT, h)
  }, [])
  return [
    s,
    (next) => {
      try {
        localStorage.setItem(KEY, JSON.stringify(next))
      } catch {
        /* storage unavailable */
      }
      setS(next)
      window.dispatchEvent(new Event(EVT))
    },
  ]
}

export function AddinsDialog({ onClose }: { onClose: () => void }) {
  const [state, save] = useAddins()
  const [draft, setDraft] = useState(state)
  return (
    <Dialog
      title="Add-ins"
      onClose={onClose}
      width={460}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" onClick={() => (save(draft), onClose())}>OK</Button>
        </>
      }
    >
      <p className="mb-2 text-[12.5px] text-[var(--muted)]">Add-ins available. Ticked add-ins appear on the <b>Data</b> tab, in the <b>Analyze</b> group.</p>
      <div className="rounded-md border border-[var(--border)] p-1" data-testid="addins-list">
        <Check label="Analysis ToolPak" checked={draft.analysis} onChange={(v) => setDraft({ ...draft, analysis: v })} />
        <p className="mb-1 ml-7 text-[11.5px] text-[var(--muted)]">Data analysis tools for statistics and engineering: descriptive statistics, histograms, regression, t-tests, ANOVA, correlation, moving averages, random numbers and more.</p>
        <Check label="Solver Add-in" checked={draft.solver} onChange={(v) => setDraft({ ...draft, solver: v })} />
        <p className="mb-1 ml-7 text-[11.5px] text-[var(--muted)]">Optimisation and what-if analysis: find the best values for decision cells subject to constraints (Simplex LP, GRG Nonlinear, Evolutionary).</p>
      </div>
    </Dialog>
  )
}

// ---------------------------------------------------------------- range picking

function RangeInput({ value, onChange, onPick, label, testid }: { value: string; onChange: (v: string) => void; onPick: () => void; label: string; testid?: string }) {
  return (
    <div className="flex items-center gap-1">
      <input className={cn(inputClass, 'min-w-0 flex-1 font-mono')} value={value} onChange={(e) => onChange(e.target.value)} aria-label={label} data-testid={testid} spellCheck={false} />
      <Button size="icon" variant="outline" tip="Collapse the dialog and select cells on the sheet" aria-label={`Select ${label} on the sheet`} onClick={onPick}>
        <Icon name="pickRange" size={14} />
      </Button>
    </div>
  )
}

/** The small bar shown while the dialog is collapsed for range selection. */
function PickBar({ title, value, onChange, onDone }: { title: string; value: string; onChange: (v: string) => void; onDone: () => void }) {
  const ref = useRef<HTMLInputElement>(null)
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      if (e.key === 'Enter' || e.key === 'Escape') {
        e.preventDefault()
        e.stopPropagation()
        onDone()
      }
    }
    document.addEventListener('keydown', k, true)
    return () => document.removeEventListener('keydown', k, true)
  }, [onDone])
  return createPortal(
    <div className="fixed left-1/2 top-24 z-[85] flex w-[440px] max-w-[94vw] -translate-x-1/2 items-center gap-2 rounded-lg border border-[var(--accent)] bg-[var(--bg)] px-3 py-2 text-[12.5px] shadow-2xl" role="dialog" aria-label={title} data-testid="pick-bar">
      <span className="shrink-0 font-medium">{title}</span>
      <input ref={ref} className={cn(inputClass, 'min-w-0 flex-1 font-mono')} value={value} onChange={(e) => onChange(e.target.value)} aria-label="Selected range" />
      <Button size="sm" variant="default" onClick={onDone} data-testid="pick-done">Done</Button>
    </div>,
    document.body,
  )
}

/** Resolves "Sheet1!$A$1:$B$5" or "A1:B5" to a sheet and range. */
export function resolveRef(wb: Workbook, baseSheetId: string, text: string): { sheetId: string; range: Range } | string {
  const t = text.trim().replace(/\$/g, '')
  if (!t) return 'Enter a range.'
  const bang = t.lastIndexOf('!')
  let sheetId = baseSheetId
  let ref = t
  if (bang >= 0) {
    const name = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'")
    const s = wb.sheetByName(name)
    if (!s) return `There is no sheet named “${name}”.`
    sheetId = s.id
    ref = t.slice(bang + 1)
  }
  const range = parseRangeA1(ref.toUpperCase())
  if (!range) return `“${text}” is not a valid range.`
  return { sheetId, range }
}

const absRef = (r: Range) => {
  const c = (n: number) => {
    let s = ''
    n++
    while (n > 0) {
      const m = (n - 1) % 26
      s = String.fromCharCode(65 + m) + s
      n = Math.floor((n - 1) / 26)
    }
    return s
  }
  const a = `$${c(r.c1)}$${r.r1 + 1}`
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:$${c(r.c2)}$${r.r2 + 1}`
}
export function liveRefText(wb: Workbook, baseSheetId: string, activeSheetId: string, prim: Range) {
  const ref = absRef(prim)
  return activeSheetId === baseSheetId ? ref : `${quoteSheetName(wb.meta(activeSheetId).name)}!${ref}`
}

// ---------------------------------------------------------------- Data Analysis
type ToolId =
  | 'anova1'
  | 'anova2rep'
  | 'anova2'
  | 'correl'
  | 'covar'
  | 'desc'
  | 'expsmooth'
  | 'ftest'
  | 'fourier'
  | 'hist'
  | 'movavg'
  | 'random'
  | 'rank'
  | 'regression'
  | 'sampling'
  | 'tpaired'
  | 'tequal'
  | 'tunequal'
  | 'ztest'

export const TOOLS: { id: ToolId; name: string; about: string }[] = [
  { id: 'anova1', name: 'Anova: Single Factor', about: 'Tests whether the means of two or more groups differ.' },
  { id: 'anova2rep', name: 'Anova: Two-Factor With Replication', about: 'Two factors with several observations per combination; also tests their interaction.' },
  { id: 'anova2', name: 'Anova: Two-Factor Without Replication', about: 'Two factors with one observation per combination.' },
  { id: 'correl', name: 'Correlation', about: 'Correlation matrix: how strongly each pair of variables moves together (−1 to 1).' },
  { id: 'covar', name: 'Covariance', about: 'Covariance matrix of several variables.' },
  { id: 'desc', name: 'Descriptive Statistics', about: 'Mean, median, mode, standard deviation, range, skewness, confidence level and more.' },
  { id: 'expsmooth', name: 'Exponential Smoothing', about: 'Forecast that weights recent values more heavily (writes linked formulas).' },
  { id: 'ftest', name: 'F-Test Two-Sample for Variances', about: 'Tests whether two samples have equal variances.' },
  { id: 'fourier', name: 'Fourier Analysis', about: 'Fast Fourier transform of 2ⁿ values (results as complex numbers).' },
  { id: 'hist', name: 'Histogram', about: 'Frequency table for bins, with optional Pareto order, cumulative % and chart.' },
  { id: 'movavg', name: 'Moving Average', about: 'Average of the last N values as a forecast (writes linked formulas).' },
  { id: 'random', name: 'Random Number Generation', about: 'Random numbers from uniform, normal, Bernoulli, binomial, Poisson, patterned or discrete distributions.' },
  { id: 'rank', name: 'Rank and Percentile', about: 'Ranks each value and gives its percentile.' },
  { id: 'regression', name: 'Regression', about: 'Linear regression: coefficients, R², ANOVA, p-values, residuals.' },
  { id: 'sampling', name: 'Sampling', about: 'Takes a periodic or random sample of the input.' },
  { id: 'tpaired', name: 't-Test: Paired Two Sample for Means', about: 'Before/after style test on paired observations.' },
  { id: 'tequal', name: 't-Test: Two-Sample Assuming Equal Variances', about: 'Compares two means assuming equal variances (pooled).' },
  { id: 'tunequal', name: 't-Test: Two-Sample Assuming Unequal Variances', about: 'Compares two means without assuming equal variances (Welch).' },
  { id: 'ztest', name: 'z-Test: Two Sample for Means', about: 'Compares two means when the population variances are known.' },
]

type FormVals = Record<string, string | boolean>

interface Spec {
  k: string
  label: string
  type: 'range' | 'number' | 'text' | 'check' | 'radio' | 'select'
  options?: [string, string][]
  def?: string | boolean
  hint?: string
  show?: (f: FormVals) => boolean
}

const inputR = (label = 'Input Range'): Spec => ({ k: 'input', label, type: 'range' })
const grouped: Spec = { k: 'grouped', label: 'Grouped By', type: 'radio', options: [['cols', 'Columns'], ['rows', 'Rows']], def: 'cols' }
const labels = (label = 'Labels in first row'): Spec => ({ k: 'labels', label, type: 'check', def: false })
const alpha: Spec = { k: 'alpha', label: 'Alpha', type: 'number', def: '0.05' }
const twoVar: Spec[] = [
  { k: 'var1', label: 'Variable 1 Range', type: 'range' },
  { k: 'var2', label: 'Variable 2 Range', type: 'range' },
]

const SPECS: Record<ToolId, Spec[]> = {
  anova1: [inputR(), grouped, labels(), alpha],
  anova2rep: [inputR(), { k: 'perSample', label: 'Rows per sample', type: 'number', def: '2' }, alpha],
  anova2: [inputR(), labels('Labels'), alpha],
  correl: [inputR(), grouped, labels()],
  covar: [inputR(), grouped, labels()],
  desc: [
    inputR(),
    grouped,
    labels(),
    { k: 'summary', label: 'Summary statistics', type: 'check', def: true },
    { k: 'conf', label: 'Confidence Level for Mean', type: 'check', def: true },
    { k: 'confLevel', label: 'Confidence level (%)', type: 'number', def: '95', show: (f) => !!f.conf },
    { k: 'kl', label: 'Kth Largest', type: 'check', def: false },
    { k: 'klv', label: 'k (largest)', type: 'number', def: '1', show: (f) => !!f.kl },
    { k: 'ks', label: 'Kth Smallest', type: 'check', def: false },
    { k: 'ksv', label: 'k (smallest)', type: 'number', def: '1', show: (f) => !!f.ks },
  ],
  expsmooth: [inputR(), { k: 'damping', label: 'Damping factor', type: 'number', def: '0.3', hint: 'α = 1 − damping factor' }, labels('Labels'), { k: 'chart', label: 'Chart Output', type: 'check', def: false }, { k: 'stderr', label: 'Standard Errors', type: 'check', def: false }],
  ftest: [...twoVar, labels('Labels'), alpha],
  fourier: [inputR(), labels('Labels in first row'), { k: 'inverse', label: 'Inverse', type: 'check', def: false }],
  hist: [
    inputR(),
    { k: 'bins', label: 'Bin Range (optional)', type: 'range', hint: 'Upper limits of each bin, in ascending order. Leave empty for equal-width bins.' },
    labels('Labels'),
    { k: 'pareto', label: 'Pareto (sorted histogram)', type: 'check', def: false },
    { k: 'cumulative', label: 'Cumulative Percentage', type: 'check', def: false },
    { k: 'chart', label: 'Chart Output', type: 'check', def: true },
  ],
  movavg: [inputR(), labels('Labels in First Row'), { k: 'interval', label: 'Interval', type: 'number', def: '3' }, { k: 'chart', label: 'Chart Output', type: 'check', def: false }, { k: 'stderr', label: 'Standard Errors', type: 'check', def: false }],
  random: [
    { k: 'vars', label: 'Number of Variables', type: 'number', def: '1' },
    { k: 'count', label: 'Number of Random Numbers', type: 'number', def: '10' },
    {
      k: 'dist',
      label: 'Distribution',
      type: 'select',
      def: 'uniform',
      options: [
        ['uniform', 'Uniform'],
        ['normal', 'Normal'],
        ['bernoulli', 'Bernoulli'],
        ['binomial', 'Binomial'],
        ['poisson', 'Poisson'],
        ['patterned', 'Patterned'],
        ['discrete', 'Discrete'],
      ],
    },
    { k: 'from', label: 'Between', type: 'number', def: '0', show: (f) => f.dist === 'uniform' || f.dist === 'patterned' },
    { k: 'to', label: 'and', type: 'number', def: '1', show: (f) => f.dist === 'uniform' || f.dist === 'patterned' },
    { k: 'step', label: 'in steps of', type: 'number', def: '1', show: (f) => f.dist === 'patterned' },
    { k: 'repEach', label: 'repeating each number (times)', type: 'number', def: '1', show: (f) => f.dist === 'patterned' },
    { k: 'repSeq', label: 'repeating the sequence (times)', type: 'number', def: '1', show: (f) => f.dist === 'patterned' },
    { k: 'mean', label: 'Mean', type: 'number', def: '0', show: (f) => f.dist === 'normal' },
    { k: 'sd', label: 'Standard deviation', type: 'number', def: '1', show: (f) => f.dist === 'normal' },
    { k: 'p', label: 'p Value', type: 'number', def: '0.5', show: (f) => f.dist === 'bernoulli' || f.dist === 'binomial' },
    { k: 'trials', label: 'Number of trials', type: 'number', def: '10', show: (f) => f.dist === 'binomial' },
    { k: 'lambda', label: 'Lambda', type: 'number', def: '1', show: (f) => f.dist === 'poisson' },
    { k: 'vp', label: 'Value and Probability Input Range', type: 'range', show: (f) => f.dist === 'discrete', hint: 'Two columns: values, then probabilities that add up to 1.' },
    { k: 'seed', label: 'Random Seed (optional)', type: 'number', def: '' },
  ],
  rank: [inputR(), grouped, labels()],
  regression: [
    { k: 'y', label: 'Input Y Range', type: 'range' },
    { k: 'x', label: 'Input X Range', type: 'range' },
    labels('Labels'),
    { k: 'zero', label: 'Constant is Zero', type: 'check', def: false },
    { k: 'conf', label: 'Confidence Level', type: 'check', def: false },
    { k: 'confLevel', label: 'Confidence level (%)', type: 'number', def: '95', show: (f) => !!f.conf },
    { k: 'resid', label: 'Residuals', type: 'check', def: false },
    { k: 'std', label: 'Standardized Residuals', type: 'check', def: false },
    { k: 'linefit', label: 'Line Fit Plots', type: 'check', def: false },
    { k: 'normal', label: 'Normal Probability Plots (table)', type: 'check', def: false },
  ],
  sampling: [
    inputR(),
    labels('Labels'),
    { k: 'method', label: 'Sampling Method', type: 'radio', options: [['periodic', 'Periodic'], ['random', 'Random']], def: 'random' },
    { k: 'period', label: 'Period', type: 'number', def: '2', show: (f) => f.method === 'periodic' },
    { k: 'samples', label: 'Number of Samples', type: 'number', def: '10', show: (f) => f.method === 'random' },
  ],
  tpaired: [...twoVar, { k: 'hyp', label: 'Hypothesized Mean Difference', type: 'number', def: '0' }, labels('Labels'), alpha],
  tequal: [...twoVar, { k: 'hyp', label: 'Hypothesized Mean Difference', type: 'number', def: '0' }, labels('Labels'), alpha],
  tunequal: [...twoVar, { k: 'hyp', label: 'Hypothesized Mean Difference', type: 'number', def: '0' }, labels('Labels'), alpha],
  ztest: [
    ...twoVar,
    { k: 'hyp', label: 'Hypothesized Mean Difference', type: 'number', def: '0' },
    { k: 'v1', label: 'Variable 1 Variance (known)', type: 'number', def: '' },
    { k: 'v2', label: 'Variable 2 Variance (known)', type: 'number', def: '' },
    labels('Labels'),
    alpha,
  ],
}

const OUTPUT_SPECS: Spec[] = [
  { k: 'out', label: 'Output options', type: 'radio', options: [['sheet', 'New Worksheet Ply'], ['range', 'Output Range']], def: 'sheet' },
  { k: 'outRange', label: 'Output Range', type: 'range', show: (f) => f.out === 'range' },
  { k: 'sheetName', label: 'New sheet name', type: 'text', def: '', show: (f) => f.out === 'sheet', hint: 'Leave empty for an automatic name.' },
]

function defaults(id: ToolId, input: string, guessLabels: boolean): FormVals {
  const f: FormVals = {}
  for (const s of [...SPECS[id], ...OUTPUT_SPECS]) if (s.def !== undefined) f[s.k] = s.def
  if ('labels' in f) f.labels = guessLabels
  for (const k of ['input', 'var1', 'y']) if (SPECS[id].some((s) => s.k === k)) f[k] = input
  return f
}

export interface AnalysisHostProps {
  wb: Workbook
  sheetId: string
  /** suggested input (current data region) */
  initialInput: Range | null
  /** live selection text while collapsed */
  liveRef: (baseSheetId: string) => string
  onClose: () => void
  onWritten: (sheetId: string, range: Range, label: string) => void
}

export function AnalysisHost(p: AnalysisHostProps) {
  const { wb } = p
  const [tool, setTool] = useState<ToolId | null>(null)
  const [listSel, setListSel] = useState<ToolId>('desc')
  const [f, setF] = useState<FormVals>({})
  const [picking, setPicking] = useState<string | null>(null)
  const [pickText, setPickText] = useState('')
  const [err, setErr] = useState<string | null>(null)
  const [confirm, setConfirm] = useState<null | (() => void)>(null)
  const baseSheet = useRef(p.sheetId).current

  // keep the pick bar showing the live selection
  const live = p.liveRef(baseSheet)
  const lastLive = useRef(live)
  useEffect(() => {
    if (picking && live !== lastLive.current) setPickText(live)
    lastLive.current = live
  }, [live, picking])

  const initialText = p.initialInput ? absRef(p.initialInput) : ''
  const guessLabels = useMemo(() => {
    const rg = p.initialInput
    if (!rg) return false
    for (let c = rg.c1; c <= rg.c2; c++) if (typeof wb.getValue(p.sheetId, rg.r1, c) === 'string') return true
    return false
  }, [p.initialInput, p.sheetId, wb])

  const open = (id: ToolId) => {
    setTool(id)
    setF(defaults(id, initialText, guessLabels))
    setErr(null)
  }

  const set = (k: string, v: string | boolean) => {
    setF((o) => ({ ...o, [k]: v }))
    setErr(null)
  }

  if (picking) {
    return <PickBar title={(tool ? [...SPECS[tool], ...OUTPUT_SPECS].find((s) => s.k === picking)?.label : '') ?? 'Select range'} value={pickText} onChange={setPickText} onDone={() => (set(picking, pickText), setPicking(null))} />
  }

  if (!tool) {
    return (
      <Dialog
        title="Data Analysis"
        onClose={p.onClose}
        width={460}
        footer={
          <>
            <Button variant="outline" onClick={p.onClose}>Cancel</Button>
            <Button variant="default" onClick={() => open(listSel)} data-testid="analysis-ok">OK</Button>
          </>
        }
      >
        <p className="mb-1 text-[12px] font-medium text-[var(--muted)]">Analysis Tools</p>
        <div className="max-h-[46vh] overflow-auto rounded-md border border-[var(--border)]" role="listbox" aria-label="Analysis Tools" data-testid="analysis-tools">
          {TOOLS.map((t) => (
            <div
              key={t.id}
              role="option"
              aria-selected={listSel === t.id}
              className={cn('cursor-default px-2 py-1 text-[13px]', listSel === t.id ? 'bg-[var(--accent)] text-[var(--accent-contrast)]' : 'hover:bg-[var(--hover)]')}
              onClick={() => setListSel(t.id)}
              onDoubleClick={() => open(t.id)}
            >
              {t.name}
            </div>
          ))}
        </div>
        <p className="mt-2 min-h-[2.5em] text-[12px] text-[var(--muted)]">{TOOLS.find((t) => t.id === listSel)?.about}</p>
      </Dialog>
    )
  }

  const toolName = TOOLS.find((t) => t.id === tool)!.name
  const specs = SPECS[tool]

  const run = (force = false) => {
    const res = runTool(wb, baseSheet, tool, f, force)
    if (typeof res === 'string') return setErr(res)
    if ('confirm' in res) return setConfirm(() => () => run(true))
    p.onWritten(res.sheetId, res.range, toolName)
  }

  const fieldEl = (s: Spec): ReactNode => {
    if (s.show && !s.show(f)) return null
    const v = f[s.k]
    switch (s.type) {
      case 'range':
        return (
          <Field key={s.k} label={s.label} hint={s.hint}>
            <RangeInput value={String(v ?? '')} onChange={(x) => set(s.k, x)} label={s.label} testid={`f-${s.k}`} onPick={() => (setPicking(s.k), setPickText(String(v ?? '') || live))} />
          </Field>
        )
      case 'number':
      case 'text':
        return (
          <Field key={s.k} label={s.label} hint={s.hint}>
            <input className={inputClass} value={String(v ?? '')} onChange={(e) => set(s.k, e.target.value)} inputMode={s.type === 'number' ? 'decimal' : undefined} aria-label={s.label} data-testid={`f-${s.k}`} />
          </Field>
        )
      case 'check':
        return <Check key={s.k} label={s.label} checked={!!v} onChange={(x) => set(s.k, x)} />
      case 'select':
        return (
          <Field key={s.k} label={s.label}>
            <select className={inputClass} value={String(v)} onChange={(e) => set(s.k, e.target.value)} aria-label={s.label} data-testid={`f-${s.k}`}>
              {s.options!.map(([val, lab]) => <option key={val} value={val}>{lab}</option>)}
            </select>
          </Field>
        )
      case 'radio':
        return (
          <div key={s.k} className="mb-2">
            <div className="mb-1 text-[12px] font-medium text-[var(--muted)]">{s.label}</div>
            <div className="flex flex-wrap gap-3" role="radiogroup" aria-label={s.label}>
              {s.options!.map(([val, lab]) => (
                <label key={val} className="flex items-center gap-1.5 text-[12.5px]">
                  <input type="radio" name={s.k} checked={v === val} onChange={() => set(s.k, val)} className="accent-[var(--accent)]" />
                  {lab}
                </label>
              ))}
            </div>
          </div>
        )
    }
  }

  return (
    <Dialog
      title={toolName}
      onClose={p.onClose}
      width={500}
      footer={
        confirm ? (
          <>
            <span className="mr-auto text-[12.5px]">Output range will overwrite existing data.</span>
            <Button variant="outline" onClick={() => setConfirm(null)}>Cancel</Button>
            <Button variant="danger" onClick={() => (confirm(), setConfirm(null))} data-testid="overwrite-ok">Overwrite</Button>
          </>
        ) : (
          <>
            <Button variant="ghost" className="mr-auto" onClick={() => setTool(null)}>← All tools</Button>
            <Button variant="outline" onClick={p.onClose}>Cancel</Button>
            <Button variant="default" onClick={() => run()} data-testid="tool-ok">OK</Button>
          </>
        )
      }
    >
      <p className="mb-3 text-[12px] text-[var(--muted)]">{TOOLS.find((t) => t.id === tool)?.about}</p>
      <div className="grid grid-cols-1 gap-x-4 sm:grid-cols-2">
        <div>{specs.filter((s) => s.type !== 'check').map(fieldEl)}</div>
        <div className="flex flex-col">{specs.filter((s) => s.type === 'check').map(fieldEl)}</div>
      </div>
      <div className="mt-2 border-t border-[var(--border)] pt-3">{OUTPUT_SPECS.map(fieldEl)}</div>
      {err && <p className="mt-1 rounded-md bg-[var(--panel)] p-2 text-[12.5px] text-[var(--danger)]" role="alert" data-testid="tool-error">{err}</p>}
    </Dialog>
  )
}

// ---------------------------------------------------------------- tool execution
const num = (v: unknown, name: string, opts: { int?: boolean; min?: number; max?: number; optional?: boolean } = {}): number | string => {
  const s = String(v ?? '').trim()
  if (!s && opts.optional) return NaN
  const n = Number(s)
  if (!s || !Number.isFinite(n)) return `${name} must be a number.`
  if (opts.int && !Number.isInteger(n)) return `${name} must be a whole number.`
  if (opts.min !== undefined && n < opts.min) return `${name} must be at least ${opts.min}.`
  if (opts.max !== undefined && n > opts.max) return `${name} must be at most ${opts.max}.`
  return n
}

type RunResult = { sheetId: string; range: Range } | { confirm: true }

export function runTool(wb: Workbook, baseSheet: string, tool: ToolId, f: FormVals, force: boolean): RunResult | string {
  const labelsOn = !!f.labels
  const ref = (k: string, what: string) => {
    const r = resolveRef(wb, baseSheet, String(f[k] ?? ''))
    return typeof r === 'string' ? `${what}: ${r}` : r
  }
  const alphaV = 'alpha' in f ? num(f.alpha, 'Alpha', { min: 0.000001, max: 0.999999 }) : 0.05
  if (typeof alphaV === 'string') return alphaV
  const two = (): [Series, Series] | string => {
    const a = ref('var1', 'Variable 1 Range')
    if (typeof a === 'string') return a
    const b = ref('var2', 'Variable 2 Range')
    if (typeof b === 'string') return b
    const s1 = readVector(wb, a.sheetId, a.range, labelsOn, 'Variable 1')
    if (typeof s1 === 'string') return s1
    const s2 = readVector(wb, b.sheetId, b.range, labelsOn, 'Variable 2')
    if (typeof s2 === 'string') return s2
    return [s1, s2]
  }
  const many = (kind = 'Column') => {
    const r = ref('input', 'Input Range')
    if (typeof r === 'string') return r
    const by = f.grouped === 'rows' ? 'rows' : 'cols'
    return readSeries(wb, r.sheetId, r.range, by, labelsOn, by === 'rows' ? 'Row' : kind)
  }

  // ---- build the tool (formula tools get the output context later)
  let build: (ctx: FormulaCtx) => OutTable | string
  let inSheet = baseSheet
  switch (tool) {
    case 'anova1': {
      const s = many()
      if (typeof s === 'string') return s
      build = () => anovaSingleFactor(s, alphaV)
      break
    }
    case 'anova2rep':
    case 'anova2': {
      const r = ref('input', 'Input Range')
      if (typeof r === 'string') return r
      const withLabels = tool === 'anova2rep' || labelsOn
      const rg = r.range
      const body = withLabels ? { r1: rg.r1 + 1, c1: rg.c1 + 1, r2: rg.r2, c2: rg.c2 } : rg
      if (body.r2 < body.r1 || body.c2 < body.c1) return 'The input range is too small.'
      const m = readMatrix(wb, r.sheetId, body)
      if (typeof m === 'string') return m
      if (m.some((row) => row.some((x) => x === null))) return 'Two-factor ANOVA needs a complete block of numbers (no blank cells).'
      const data = m as number[][]
      const txt = (rr: number, cc: number) => wb.getDisplay(r.sheetId, rr, cc).text
      const colNames = data[0].map((_, j) => (withLabels ? txt(rg.r1, rg.c1 + 1 + j) : '') || `Column ${j + 1}`)
      if (tool === 'anova2') {
        const rowNames = data.map((_, i) => (withLabels ? txt(rg.r1 + 1 + i, rg.c1) : '') || `Row ${i + 1}`)
        build = () => anovaTwoFactorNoRep(data, rowNames, colNames, alphaV)
      } else {
        const per = num(f.perSample, 'Rows per sample', { int: true, min: 2 })
        if (typeof per === 'string') return per
        const groups = Array.from({ length: Math.floor(data.length / per) }, (_, i) => txt(rg.r1 + 1 + i * per, rg.c1) || `Sample ${i + 1}`)
        build = () => anovaTwoFactorRep(data, per, groups, colNames, alphaV)
      }
      break
    }
    case 'correl':
    case 'covar': {
      const s = many()
      if (typeof s === 'string') return s
      build = () => (tool === 'correl' ? correlationTool(s) : covarianceTool(s))
      break
    }
    case 'desc': {
      const s = many()
      if (typeof s === 'string') return s
      const cl = f.conf ? num(f.confLevel, 'Confidence level', { min: 0.0001, max: 99.9999 }) : NaN
      if (typeof cl === 'string') return cl
      const kl = f.kl ? num(f.klv, 'k (largest)', { int: true, min: 1 }) : NaN
      if (typeof kl === 'string') return kl
      const ks = f.ks ? num(f.ksv, 'k (smallest)', { int: true, min: 1 }) : NaN
      if (typeof ks === 'string') return ks
      build = () => descriptiveStatistics(s, { summary: !!f.summary, confidence: f.conf ? cl / 100 : undefined, kthLargest: f.kl ? kl : undefined, kthSmallest: f.ks ? ks : undefined })
      break
    }
    case 'expsmooth':
    case 'movavg': {
      const r = ref('input', 'Input Range')
      if (typeof r === 'string') return r
      const name = labelsOn ? wb.getDisplay(r.sheetId, r.range.r1, r.range.c1).text || 'Actual' : 'Actual'
      if (tool === 'movavg') {
        const iv = num(f.interval, 'Interval', { int: true, min: 2 })
        if (typeof iv === 'string') return iv
        build = (ctx) => movingAverage(r.range, labelsOn, iv, !!f.stderr, !!f.chart, ctx, name)
      } else {
        const d = num(f.damping, 'Damping factor', { min: 0, max: 1 })
        if (typeof d === 'string') return d
        build = (ctx) => exponentialSmoothing(r.range, labelsOn, d, !!f.stderr, !!f.chart, ctx, name)
      }
      // the formulas refer to the input sheet
      inSheet = r.sheetId
      break
    }
    case 'ftest':
    case 'tpaired':
    case 'tequal':
    case 'tunequal':
    case 'ztest': {
      const s = two()
      if (typeof s === 'string') return s
      const hyp = tool === 'ftest' ? 0 : num(f.hyp, 'Hypothesized Mean Difference')
      if (typeof hyp === 'string') return hyp
      if (tool === 'ftest') build = () => fTestTwoSample(s[0], s[1], alphaV)
      else if (tool === 'tpaired') build = () => tTestPaired(s[0], s[1], hyp, alphaV)
      else if (tool === 'ztest') {
        const v1 = num(f.v1, 'Variable 1 Variance (known)', { min: 0 })
        if (typeof v1 === 'string') return v1
        const v2 = num(f.v2, 'Variable 2 Variance (known)', { min: 0 })
        if (typeof v2 === 'string') return v2
        build = () => zTestTwoSample(s[0], s[1], hyp, v1, v2, alphaV)
      } else build = () => tTestTwoSample(s[0], s[1], hyp, alphaV, tool === 'tequal')
      break
    }
    case 'fourier': {
      const r = ref('input', 'Input Range')
      if (typeof r === 'string') return r
      const s = readVector(wb, r.sheetId, r.range, labelsOn, 'Input')
      if (typeof s === 'string') return s
      build = () => fourierTool(s.values, !!f.inverse)
      break
    }
    case 'hist': {
      const r = ref('input', 'Input Range')
      if (typeof r === 'string') return r
      const s = readVector(wb, r.sheetId, r.range, labelsOn, 'Input')
      if (typeof s === 'string') return s
      let bins: number[] | undefined
      let binLabel: string | undefined
      if (String(f.bins ?? '').trim()) {
        const b = ref('bins', 'Bin Range')
        if (typeof b === 'string') return b
        const bs = readVector(wb, b.sheetId, b.range, labelsOn, 'Bin')
        if (typeof bs === 'string') return bs
        bins = bs.values
        binLabel = labelsOn ? bs.name : undefined
      }
      build = () => histogramTool(s, { bins, binLabel, pareto: !!f.pareto, cumulative: !!f.cumulative, chart: !!f.chart })
      break
    }
    case 'random': {
      const vars = num(f.vars, 'Number of Variables', { int: true, min: 1 })
      if (typeof vars === 'string') return vars
      const count = f.dist === 'patterned' ? 1 : num(f.count, 'Number of Random Numbers', { int: true, min: 1 })
      if (typeof count === 'string') return count
      const seed = num(f.seed, 'Random Seed', { optional: true })
      if (typeof seed === 'string') return seed
      const n = (k: string, label: string) => num(f[k], label)
      let dist: RandomDist
      const need = (...vals: (number | string)[]) => vals.find((x) => typeof x === 'string') as string | undefined
      switch (f.dist) {
        case 'normal': {
          const m = n('mean', 'Mean')
          const sd = n('sd', 'Standard deviation')
          const e = need(m, sd)
          if (e) return e
          dist = { kind: 'normal', mean: m as number, sd: sd as number }
          break
        }
        case 'bernoulli': {
          const pv = n('p', 'p Value')
          if (typeof pv === 'string') return pv
          dist = { kind: 'bernoulli', p: pv }
          break
        }
        case 'binomial': {
          const pv = n('p', 'p Value')
          const t = n('trials', 'Number of trials')
          const e = need(pv, t)
          if (e) return e
          dist = { kind: 'binomial', p: pv as number, trials: t as number }
          break
        }
        case 'poisson': {
          const l = n('lambda', 'Lambda')
          if (typeof l === 'string') return l
          dist = { kind: 'poisson', lambda: l }
          break
        }
        case 'patterned': {
          const a = n('from', 'Between')
          const b = n('to', 'and')
          const st = n('step', 'Step')
          const re = n('repEach', 'Repeat each')
          const rs = n('repSeq', 'Repeat sequence')
          const e = need(a, b, st, re, rs)
          if (e) return e
          dist = { kind: 'patterned', from: a as number, to: b as number, step: st as number, repeatEach: re as number, repeatSeq: rs as number }
          break
        }
        case 'discrete': {
          const r = ref('vp', 'Value and Probability Input Range')
          if (typeof r === 'string') return r
          if (r.range.c2 - r.range.c1 !== 1) return 'Value and Probability Input Range must be two columns wide.'
          const m = readMatrix(wb, r.sheetId, r.range)
          if (typeof m === 'string') return m
          const rows = m.filter((row) => row[0] !== null && row[1] !== null) as number[][]
          dist = { kind: 'discrete', values: rows.map((x) => x[0]), probs: rows.map((x) => x[1]) }
          break
        }
        default: {
          const a = n('from', 'Between')
          const b = n('to', 'and')
          const e = need(a, b)
          if (e) return e
          dist = { kind: 'uniform', from: a as number, to: b as number }
        }
      }
      build = () => randomNumbers(vars, count, dist, Number.isFinite(seed) ? seed : undefined)
      break
    }
    case 'rank': {
      const s = many()
      if (typeof s === 'string') return s
      build = () => rankAndPercentile(s)
      break
    }
    case 'regression': {
      const yr = ref('y', 'Input Y Range')
      if (typeof yr === 'string') return yr
      const xr = ref('x', 'Input X Range')
      if (typeof xr === 'string') return xr
      const y = readVector(wb, yr.sheetId, yr.range, labelsOn, 'Y')
      if (typeof y === 'string') return y
      const xs = readSeries(wb, xr.sheetId, xr.range, 'cols', labelsOn, 'X Variable')
      if (typeof xs === 'string') return xs
      if (!labelsOn) xs.forEach((x, i) => (x.name = xs.length === 1 ? 'X Variable 1' : `X Variable ${i + 1}`))
      if (!labelsOn) y.name = 'Y'
      const cl = f.conf ? num(f.confLevel, 'Confidence level', { min: 0.0001, max: 99.9999 }) : 95
      if (typeof cl === 'string') return cl
      build = () => regressionTool(y, xs, { constantZero: !!f.zero, confidence: cl / 100, residuals: !!f.resid, standardized: !!f.std, normalProbability: !!f.normal, lineFitChart: !!f.linefit })
      break
    }
    case 'sampling': {
      const r = ref('input', 'Input Range')
      if (typeof r === 'string') return r
      const s = readSeries(wb, r.sheetId, r.range, r.range.c1 === r.range.c2 ? 'cols' : 'rows', labelsOn)
      if (typeof s === 'string') return s
      const all: Series = { name: 'Sample', values: s.flatMap((x) => x.values) }
      const amount = num(f.method === 'periodic' ? f.period : f.samples, f.method === 'periodic' ? 'Period' : 'Number of Samples', { int: true, min: 1 })
      if (typeof amount === 'string') return amount
      build = () => samplingTool(all, f.method === 'periodic' ? 'periodic' : 'random', amount, seededRandom())
      break
    }
  }

  // ---- validate with a throw-away context, then place the output
  const dryCtx = formulaCtx(wb, inSheet, { sheetId: inSheet, r: 0, c: 0 })
  const dry = build(dryCtx)
  if (typeof dry === 'string') return dry
  let target: { sheetId: string; r: number; c: number }
  let fresh = false
  if (f.out === 'range') {
    const o = ref('outRange', 'Output Range')
    if (typeof o === 'string') return o
    target = { sheetId: o.sheetId, r: o.range.r1, c: o.range.c1 }
    if (!force && outputOverlapsData(wb, target, dry)) return { confirm: true }
  } else {
    const t = resolveTarget(wb, { kind: 'newSheet', name: String(f.sheetName ?? '') })
    if (typeof t === 'string') return t
    target = t
    fresh = true
  }
  const table = build(formulaCtx(wb, inSheet, target))
  if (typeof table === 'string') return table
  const range = writeTable(wb, table, target, TOOLS.find((t) => t.id === tool)!.name, fresh)
  return { sheetId: target.sheetId, range }
}

// ---------------------------------------------------------------- Solver
const OPS: { v: SolverOp; label: string }[] = [
  { v: '<=', label: '<=' },
  { v: '=', label: '=' },
  { v: '>=', label: '>=' },
  { v: 'int', label: 'int' },
  { v: 'bin', label: 'bin' },
]
const METHODS: { v: SolverModel['method']; label: string; about: string }[] = [
  { v: 'grg', label: 'GRG Nonlinear', about: 'Select the GRG Nonlinear engine for Solver problems that are smooth nonlinear.' },
  { v: 'simplex', label: 'Simplex LP', about: 'Select the LP Simplex engine for linear Solver problems.' },
  { v: 'evolutionary', label: 'Evolutionary', about: 'Select the Evolutionary engine for Solver problems that are non-smooth. Every variable needs upper and lower bounds.' },
]

export interface SolverHostProps {
  wb: Workbook
  sheetId: string
  activeCell: { r: number; c: number }
  liveRef: (baseSheetId: string) => string
  onClose: () => void
  /** after the solution is kept / reports are created */
  onDone: (msg: string, reportSheet?: string) => void
}

const constraintText = (c: SolverConstraint) => (c.op === 'int' ? `${c.ref} = integer` : c.op === 'bin' ? `${c.ref} = binary` : `${c.ref} ${c.op} ${c.rhs}`)

export function SolverHost(p: SolverHostProps) {
  const { wb, sheetId } = p
  const saved = wb.meta(sheetId).solver
  const activeA1 = absRef({ r1: p.activeCell.r, c1: p.activeCell.c, r2: p.activeCell.r, c2: p.activeCell.c })
  const [m, setM] = useState<SolverModel>(
    () =>
      saved ?? {
        objective: wb.isFormula(sheetId, p.activeCell.r, p.activeCell.c) ? activeA1 : '',
        goal: 'max',
        variables: '',
        constraints: [],
        nonNegative: true,
        method: 'grg',
        options: { maxTimeSec: 60, maxIterations: 500, precision: 0.000001, multistart: false },
      },
  )
  const [sel, setSel] = useState<number | null>(null)
  const [edit, setEdit] = useState<null | { index: number | null; c: SolverConstraint }>(null)
  const [showOptions, setShowOptions] = useState(false)
  const [picking, setPicking] = useState<null | { key: string; title: string }>(null)
  const [pickText, setPickText] = useState('')
  const [err, setErr] = useState<string | null>(null)
  const [running, setRunning] = useState<null | { iteration: number; objective: number; seconds: number }>(null)
  const stopRef = useRef(false)
  const [result, setResult] = useState<SolverResult | null>(null)
  const [keep, setKeep] = useState(true)
  const [reports, setReports] = useState({ answer: false, sensitivity: false })
  const live = p.liveRef(sheetId)
  const lastLive = useRef(live)
  useEffect(() => {
    if (picking && live !== lastLive.current) setPickText(live)
    lastLive.current = live
  }, [live, picking])

  const startPick = (key: string, title: string, current: string) => {
    setPicking({ key, title })
    setPickText(current || live)
  }
  const finishPick = () => {
    if (!picking) return
    const v = pickText.replace(new RegExp(`^${quoteSheetName(wb.meta(sheetId).name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}!`), '')
    if (picking.key === 'objective') setM((o) => ({ ...o, objective: v }))
    else if (picking.key === 'variables') setM((o) => ({ ...o, variables: v }))
    else if (picking.key === 'c.ref' && edit) setEdit({ ...edit, c: { ...edit.c, ref: v } })
    else if (picking.key === 'c.rhs' && edit) setEdit({ ...edit, c: { ...edit.c, rhs: v } })
    setPicking(null)
  }

  if (picking) return <PickBar title={picking.title} value={pickText} onChange={setPickText} onDone={finishPick} />

  const doSolve = async () => {
    setErr(null)
    saveModel(wb, sheetId, m)
    stopRef.current = false
    setRunning({ iteration: 0, objective: NaN, seconds: 0 })
    await new Promise((r) => setTimeout(r, 30)) // let the progress dialog paint
    const res = await solve(wb, sheetId, m, {
      onProgress: (info) => {
        setRunning(info)
        return !stopRef.current
      },
    })
    setRunning(null)
    if (!res.ok && !res.x) {
      setResult(res)
      return
    }
    if (res.ok) applySolution(wb, sheetId, res) // show the answer in the sheet while the results dialog is open
    setKeep(true)
    setReports({ answer: false, sensitivity: false })
    setResult(res)
  }

  // ---- progress
  if (running) {
    return (
      <Dialog title="Solver — working" onClose={() => (stopRef.current = true)} width={380} footer={<Button variant="outline" onClick={() => (stopRef.current = true)} data-testid="solver-stop">Stop</Button>}>
        <p className="text-[13px]" role="status">Solving… iteration {running.iteration}{Number.isFinite(running.objective) ? `, current objective ${+running.objective.toPrecision(8)}` : ''} ({running.seconds.toFixed(1)} s)</p>
      </Dialog>
    )
  }

  // ---- results
  if (result) {
    const canReport = !!result.ok && !!result.details
    const sens = canReport && !!result.details!.shadow
    const finish = () => {
      if (result.ok && !keep) {
        if (wb.undoLabel() === 'Solver') wb.undo()
        p.onDone('Original values restored.')
        return
      }
      let reportSheet: string | undefined
      if (result.ok && reports.answer) {
        const t = resolveTarget(wb, { kind: 'newSheet', name: uniqueSheetName(wb, 'Answer Report ') })
        if (typeof t !== 'string') {
          writeTable(wb, answerReport(wb, sheetId, m, result), t, 'Answer Report', true)
          reportSheet = t.sheetId
        }
      }
      if (result.ok && reports.sensitivity && sens) {
        const tbl = sensitivityReport(wb, sheetId, result)
        const t = resolveTarget(wb, { kind: 'newSheet', name: uniqueSheetName(wb, 'Sensitivity Report ') })
        if (tbl && typeof t !== 'string') {
          writeTable(wb, tbl, t, 'Sensitivity Report', true)
          reportSheet = reportSheet ?? t.sheetId
        }
      }
      p.onDone(result.ok ? 'Solver solution kept.' : result.message, reportSheet)
    }
    return (
      <Dialog
        title="Solver Results"
        onClose={() => (result.ok && wb.undoLabel() === 'Solver' && wb.undo(), p.onClose())}
        width={520}
        footer={
          <>
            {!result.ok && <Button variant="outline" onClick={() => setResult(null)}>Return to Solver Parameters</Button>}
            <Button variant="default" onClick={finish} data-testid="solver-results-ok">OK</Button>
          </>
        }
      >
        <p className={cn('mb-3 rounded-md p-2 text-[13px]', result.ok ? 'bg-[var(--panel)]' : 'bg-[var(--panel)] text-[var(--danger)]')} role="status" data-testid="solver-message">{result.message}</p>
        {result.ok && (
          <div className="grid grid-cols-2 gap-4">
            <div role="radiogroup" aria-label="Keep or restore">
              <label className="flex items-center gap-1.5 text-[12.5px]"><input type="radio" checked={keep} onChange={() => setKeep(true)} className="accent-[var(--accent)]" /> Keep Solver Solution</label>
              <label className="mt-1 flex items-center gap-1.5 text-[12.5px]"><input type="radio" checked={!keep} onChange={() => setKeep(false)} className="accent-[var(--accent)]" /> Restore Original Values</label>
            </div>
            <div>
              <div className="mb-1 text-[12px] font-medium text-[var(--muted)]">Reports</div>
              <Check label="Answer" checked={reports.answer} onChange={(v) => setReports({ ...reports, answer: v })} />
              {sens && <Check label="Sensitivity" checked={reports.sensitivity} onChange={(v) => setReports({ ...reports, sensitivity: v })} />}
            </div>
          </div>
        )}
        <p className="mt-3 text-[11.5px] text-[var(--muted)]">{result.engine} · {result.seconds.toFixed(2)} s · {result.iterations} iterations{result.objective !== undefined ? ` · objective ${+result.objective.toPrecision(10)}` : ''}</p>
      </Dialog>
    )
  }

  // ---- constraint editor
  if (edit) {
    const c = edit.c
    const save = (another: boolean) => {
      if (!c.ref.trim()) return setErr('Enter a cell reference.')
      if (c.op !== 'int' && c.op !== 'bin' && !String(c.rhs ?? '').trim()) return setErr('Enter a constraint value or cell.')
      const list = [...m.constraints]
      const clean: SolverConstraint = c.op === 'int' || c.op === 'bin' ? { ref: c.ref.trim(), op: c.op } : { ref: c.ref.trim(), op: c.op, rhs: String(c.rhs).trim() }
      if (edit.index === null) list.push(clean)
      else list[edit.index] = clean
      setM({ ...m, constraints: list })
      setErr(null)
      setEdit(another ? { index: null, c: { ref: '', op: '<=', rhs: '' } } : null)
    }
    return (
      <Dialog
        title={edit.index === null ? 'Add Constraint' : 'Change Constraint'}
        onClose={() => (setEdit(null), setErr(null))}
        width={520}
        footer={
          <>
            <Button variant="outline" onClick={() => (setEdit(null), setErr(null))}>Cancel</Button>
            {edit.index === null && <Button variant="outline" onClick={() => save(true)}>Add</Button>}
            <Button variant="default" onClick={() => save(false)} data-testid="constraint-ok">OK</Button>
          </>
        }
      >
        <div className="grid grid-cols-[1fr_auto_1fr] items-end gap-2">
          <Field label="Cell Reference">
            <RangeInput value={c.ref} onChange={(v) => setEdit({ ...edit, c: { ...c, ref: v } })} label="Cell Reference" testid="c-ref" onPick={() => startPick('c.ref', 'Cell Reference', c.ref)} />
          </Field>
          <Field label=" ">
            <select className={inputClass} value={c.op} onChange={(e) => setEdit({ ...edit, c: { ...c, op: e.target.value as SolverOp } })} aria-label="Relation" data-testid="c-op">
              {OPS.map((o) => <option key={o.v} value={o.v}>{o.label}</option>)}
            </select>
          </Field>
          <Field label="Constraint">
            {c.op === 'int' || c.op === 'bin' ? (
              <input className={inputClass} value={c.op === 'int' ? 'integer' : 'binary'} disabled aria-label="Constraint" />
            ) : (
              <RangeInput value={c.rhs ?? ''} onChange={(v) => setEdit({ ...edit, c: { ...c, rhs: v } })} label="Constraint" testid="c-rhs" onPick={() => startPick('c.rhs', 'Constraint', c.rhs ?? '')} />
            )}
          </Field>
        </div>
        <p className="text-[11.5px] text-[var(--muted)]">The constraint can be a number, a cell, or a range the same size as the cell reference (e.g. <code>B6:B8 &lt;= C6:C8</code>).</p>
        {err && <p className="mt-2 text-[12.5px] text-[var(--danger)]" role="alert">{err}</p>}
      </Dialog>
    )
  }

  // ---- options
  if (showOptions) {
    const o = m.options ?? {}
    const setO = (patch: Partial<NonNullable<SolverModel['options']>>) => setM({ ...m, options: { ...o, ...patch } })
    return (
      <Dialog title="Solver Options" onClose={() => setShowOptions(false)} width={420} footer={<Button variant="default" onClick={() => setShowOptions(false)}>OK</Button>}>
        <Field label="Constraint Precision"><input className={inputClass} value={String(o.precision ?? 0.000001)} onChange={(e) => setO({ precision: Number(e.target.value) || 0.000001 })} /></Field>
        <Field label="Max Time (seconds)"><input className={inputClass} value={String(o.maxTimeSec ?? 60)} onChange={(e) => setO({ maxTimeSec: Math.max(1, Number(e.target.value) || 60) })} /></Field>
        <Field label="Iterations"><input className={inputClass} value={String(o.maxIterations ?? 500)} onChange={(e) => setO({ maxIterations: Math.max(1, Number(e.target.value) || 500) })} /></Field>
        <Check label="Use Multistart (GRG Nonlinear)" checked={!!o.multistart} onChange={(v) => setO({ multistart: v })} />
        <Field label="Random Seed (Evolutionary / Multistart, optional)"><input className={inputClass} value={o.seed === undefined ? '' : String(o.seed)} onChange={(e) => setO({ seed: e.target.value.trim() === '' ? undefined : Number(e.target.value) })} /></Field>
      </Dialog>
    )
  }

  // ---- parameters
  return (
    <Dialog
      title="Solver Parameters"
      onClose={p.onClose}
      width={600}
      footer={
        <>
          <Button variant="ghost" className="mr-auto" onClick={() => setShowOptions(true)}>Options…</Button>
          <Button variant="outline" onClick={() => (saveModel(wb, sheetId, m), p.onClose())}>Close</Button>
          <Button variant="default" onClick={doSolve} data-testid="solver-solve">Solve</Button>
        </>
      }
    >
      <Field label="Set Objective">
        <RangeInput value={m.objective} onChange={(v) => setM({ ...m, objective: v })} label="Set Objective" testid="solver-objective" onPick={() => startPick('objective', 'Set Objective', m.objective)} />
      </Field>
      <div className="mb-3 flex flex-wrap items-center gap-4 text-[12.5px]" role="radiogroup" aria-label="To">
        <span className="text-[12px] font-medium text-[var(--muted)]">To:</span>
        {(['max', 'min', 'value'] as const).map((g) => (
          <label key={g} className="flex items-center gap-1.5">
            <input type="radio" name="goal" checked={m.goal === g} onChange={() => setM({ ...m, goal: g })} className="accent-[var(--accent)]" />
            {g === 'max' ? 'Max' : g === 'min' ? 'Min' : 'Value Of:'}
          </label>
        ))}
        <input className={cn(inputClass, 'w-24')} disabled={m.goal !== 'value'} value={m.targetValue === undefined ? '' : String(m.targetValue)} onChange={(e) => setM({ ...m, targetValue: e.target.value === '' ? undefined : Number(e.target.value) })} aria-label="Value Of" />
      </div>
      <Field label="By Changing Variable Cells">
        <RangeInput value={m.variables} onChange={(v) => setM({ ...m, variables: v })} label="By Changing Variable Cells" testid="solver-variables" onPick={() => startPick('variables', 'By Changing Variable Cells', m.variables)} />
      </Field>
      <div className="mb-1 text-[12px] font-medium text-[var(--muted)]">Subject to the Constraints</div>
      <div className="flex gap-2">
        <div className="h-32 flex-1 overflow-auto rounded-md border border-[var(--border)] font-mono text-[12.5px]" role="listbox" aria-label="Constraints" data-testid="solver-constraints">
          {m.constraints.map((c, i) => (
            <div key={i} role="option" aria-selected={sel === i} className={cn('cursor-default px-2 py-0.5', sel === i ? 'bg-[var(--accent)] text-[var(--accent-contrast)]' : 'hover:bg-[var(--hover)]')} onClick={() => setSel(i)} onDoubleClick={() => setEdit({ index: i, c: { ...c } })}>
              {constraintText(c)}
            </div>
          ))}
        </div>
        <div className="flex w-24 flex-col gap-1">
          <Button size="sm" variant="outline" onClick={() => setEdit({ index: null, c: { ref: '', op: '<=', rhs: '' } })} data-testid="constraint-add">Add</Button>
          <Button size="sm" variant="outline" disabled={sel === null} onClick={() => sel !== null && setEdit({ index: sel, c: { ...m.constraints[sel] } })}>Change</Button>
          <Button size="sm" variant="outline" disabled={sel === null} onClick={() => sel !== null && (setM({ ...m, constraints: m.constraints.filter((_, i) => i !== sel) }), setSel(null))}>Delete</Button>
          <Button size="sm" variant="outline" onClick={() => (setM({ ...m, objective: '', variables: '', constraints: [], targetValue: undefined }), setSel(null))}>Reset All</Button>
        </div>
      </div>
      <div className="mt-2">
        <Check label="Make Unconstrained Variables Non-Negative" checked={m.nonNegative} onChange={(v) => setM({ ...m, nonNegative: v })} />
      </div>
      <Field label="Select a Solving Method">
        <select className={inputClass} value={m.method} onChange={(e) => setM({ ...m, method: e.target.value as SolverModel['method'] })} aria-label="Solving method" data-testid="solver-method">
          {METHODS.map((x) => <option key={x.v} value={x.v}>{x.label}</option>)}
        </select>
      </Field>
      <p className="text-[11.5px] text-[var(--muted)]">{METHODS.find((x) => x.v === m.method)?.about}</p>
      {err && <p className="mt-2 text-[12.5px] text-[var(--danger)]" role="alert">{err}</p>}
    </Dialog>
  )
}
