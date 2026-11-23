/**
 * Solver add-in: finds values for variable cells that maximise, minimise or hit a target for an
 * objective cell, subject to constraints — like Excel's Solver.
 *
 * Methods
 *  - Simplex LP: the model is checked for linearity (by probing), turned into a linear program and
 *    solved with a two-phase simplex; integer/binary constraints use branch and bound.
 *  - GRG Nonlinear: augmented-Lagrangian method with a projected quasi-Newton (BFGS) inner solver
 *    and finite-difference gradients; optional multistart; integers by branch and bound.
 *  - Evolutionary: differential evolution with feasibility-first selection (needs bounds on every
 *    variable), followed by a local polish.
 *
 * The workbook is never touched while solving: trial values go into a private copy of the
 * calculation engine. Only the final answer is written, as one undoable "Solver" step.
 */
import { HyperFormula } from 'hyperformula'
import { HF_CONFIG, type Workbook } from '../workbook'
import { colToLetters, parseRangeA1 } from '../../model/address'
import type { Range, SolverConstraint, SolverModel } from '../../model/types'
import { TableBuilder, type OutTable } from './analysis'

// ---------------------------------------------------------------- parsing
type Cell = { r: number; c: number }
const a1 = (p: Cell) => `$${colToLetters(p.c)}$${p.r + 1}`

/** Parses "B2:B5, D7" (optionally prefixed with the sheet name) into ranges on `sheetName`. */
export function parseRefList(text: string, sheetName: string): Range[] | string {
  const parts = text.split(/[,;]/).map((s) => s.trim()).filter(Boolean)
  if (!parts.length) return 'Enter a cell or range.'
  const out: Range[] = []
  for (const p of parts) {
    let ref = p.replace(/\$/g, '')
    const bang = ref.lastIndexOf('!')
    if (bang >= 0) {
      const sh = ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'")
      if (sh.toLowerCase() !== sheetName.toLowerCase()) return `“${p}” is on another sheet. Solver cells must be on the active sheet (${sheetName}).`
      ref = ref.slice(bang + 1)
    }
    const rg = parseRangeA1(ref.toUpperCase())
    if (!rg) return `“${p}” is not a valid cell or range.`
    out.push(rg)
  }
  return out
}

const cellsOf = (rgs: Range[]): Cell[] => {
  const out: Cell[] = []
  for (const rg of rgs) for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) out.push({ r, c })
  return out
}

// ---------------------------------------------------------------- evaluation engine
interface Row {
  /** constraint index in the model, and element index inside a range constraint */
  ci: number
  label: string
  lhs: Cell
  op: '<=' | '>=' | '='
  rhs: { cell: Cell } | { value: number }
}

class Probe {
  hf: HyperFormula
  sheet: number
  evals = 0
  constructor(wb: Workbook, sheetId: string, private vars: Cell[]) {
    const sheets = wb.hf.getAllSheetsSerialized()
    let named: ReturnType<HyperFormula['getAllNamedExpressionsSerialized']> = []
    try {
      named = wb.hf.getAllNamedExpressionsSerialized()
    } catch {
      named = []
    }
    this.hf = HyperFormula.buildFromSheets(sheets, { ...HF_CONFIG, undoLimit: 0 }, named)
    this.sheet = this.hf.getSheetId(wb.meta(sheetId).name)!
  }
  set(x: number[]) {
    this.evals++
    this.hf.batch(() => this.vars.forEach((v, i) => this.hf.setCellContents({ sheet: this.sheet, row: v.r, col: v.c }, [[x[i]]])))
  }
  num(p: Cell): number {
    const v = this.hf.getCellValue({ sheet: this.sheet, row: p.r, col: p.c })
    return typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : v === null ? 0 : NaN
  }
  destroy() {
    this.hf.destroy()
  }
}

/** Values of objective and constraint rows at x (g ≤ 0 / h = 0 normalised). */
interface Eval {
  f: number
  /** lhs − rhs for every row (sign meaning given by op) */
  d: number[]
}

// ---------------------------------------------------------------- results
export interface SolverResult {
  ok: boolean
  /** Excel-style status line */
  message: string
  /** solution values for the variable cells (present when a usable solution was found) */
  x?: number[]
  objective?: number
  feasible?: boolean
  engine: string
  iterations: number
  seconds: number
  /** for reports */
  details?: {
    varCells: Cell[]
    original: number[]
    originalObjective: number
    rows: Row[]
    integerVars: Set<number>
    shadow?: number[]
    reduced?: number[]
    objCoef?: number[]
    linear: boolean
  }
}

export interface SolveHooks {
  /** called now and then so the UI can show progress; return false to stop */
  onProgress?: (info: { iteration: number; objective: number; seconds: number }) => boolean | void
}

const tick = () => new Promise<void>((res) => setTimeout(res, 0))

// ---------------------------------------------------------------- linear programming
export interface LP {
  c: number[]
  A: number[][]
  b: number[]
  ops: ('<=' | '>=' | '=')[]
  lower: number[]
  upper: number[]
  maximize: boolean
}
export interface LPResult {
  status: 'optimal' | 'infeasible' | 'unbounded' | 'limit'
  x: number[]
  z: number
  iterations: number
}

const EPS = 1e-9

/** Two-phase dense tableau simplex with Bland's rule. Handles general bounds by substitution. */
export function solveLP(lp: LP, maxIter = 20000): LPResult {
  const n = lp.c.length
  // variable substitution: x_j = shift_j + sign_j * y_k  (and free variables as y⁺ − y⁻)
  type Map = { cols: number[]; signs: number[]; shift: number }
  const maps: Map[] = []
  let ncol = 0
  const extraRows: { coef: Map; rhs: number }[] = []
  for (let j = 0; j < n; j++) {
    const l = lp.lower[j]
    const u = lp.upper[j]
    if (l > u + EPS) return { status: 'infeasible', x: [], z: NaN, iterations: 0 }
    if (Number.isFinite(l)) {
      const m = { cols: [ncol++], signs: [1], shift: l }
      maps.push(m)
      if (Number.isFinite(u)) extraRows.push({ coef: m, rhs: u - l })
    } else if (Number.isFinite(u)) {
      maps.push({ cols: [ncol++], signs: [-1], shift: u })
    } else {
      maps.push({ cols: [ncol++, ncol++], signs: [1, -1], shift: 0 })
    }
  }
  // rows in y-space
  const rows: { a: number[]; op: '<=' | '>=' | '='; b: number }[] = []
  lp.A.forEach((arow, i) => {
    const a = new Array(ncol).fill(0)
    let shift = 0
    arow.forEach((v, j) => {
      if (!v) return
      shift += v * maps[j].shift
      maps[j].cols.forEach((col, k) => (a[col] += v * maps[j].signs[k]))
    })
    rows.push({ a, op: lp.ops[i], b: lp.b[i] - shift })
  })
  for (const er of extraRows) {
    const a = new Array(ncol).fill(0)
    a[er.coef.cols[0]] = 1
    rows.push({ a, op: '<=', b: er.rhs })
  }
  const cy = new Array(ncol).fill(0)
  let c0 = 0
  lp.c.forEach((v, j) => {
    c0 += v * maps[j].shift
    maps[j].cols.forEach((col, k) => (cy[col] += v * maps[j].signs[k]))
  })
  // standard form: make b ≥ 0
  for (const r of rows)
    if (r.b < 0) {
      r.a = r.a.map((v) => -v)
      r.b = -r.b
      r.op = r.op === '<=' ? '>=' : r.op === '>=' ? '<=' : '='
    }
  const m = rows.length
  const nSlack = rows.filter((r) => r.op !== '=').length
  const nArt = rows.filter((r) => r.op !== '<=').length
  const total = ncol + nSlack + nArt
  const T: number[][] = []
  const basis: number[] = []
  const artificial = new Set<number>()
  let s = ncol
  let a = ncol + nSlack
  rows.forEach((r) => {
    const row = new Array(total + 1).fill(0)
    r.a.forEach((v, j) => (row[j] = v))
    row[total] = r.b
    if (r.op === '<=') {
      row[s] = 1
      basis.push(s++)
    } else if (r.op === '>=') {
      row[s++] = -1
      row[a] = 1
      artificial.add(a)
      basis.push(a++)
    } else {
      row[a] = 1
      artificial.add(a)
      basis.push(a++)
    }
    T.push(row)
  })
  let iterations = 0
  const pivot = (pr: number, pc: number) => {
    const p = T[pr][pc]
    const prow = T[pr]
    for (let j = 0; j <= total; j++) prow[j] /= p
    for (let i = 0; i < T.length; i++) {
      if (i === pr) continue
      const f = T[i][pc]
      if (Math.abs(f) > 1e-15) {
        const row = T[i]
        for (let j = 0; j <= total; j++) row[j] -= f * prow[j]
      }
    }
    basis[pr] = pc
  }
  /** minimises cost·y over the tableau; obj row is appended temporarily */
  const run = (cost: number[], allowed: (j: number) => boolean): 'optimal' | 'unbounded' | 'limit' => {
    const obj = [...cost, 0]
    for (let i = 0; i < m; i++) {
      const cb = cost[basis[i]]
      if (cb) for (let j = 0; j <= total; j++) obj[j] -= cb * T[i][j]
    }
    T.push(obj)
    try {
      for (;;) {
        if (++iterations > maxIter) return 'limit'
        let pc = -1
        for (let j = 0; j < total; j++)
          if (allowed(j) && obj[j] < -1e-10) {
            pc = j
            break
          }
        if (pc < 0) return 'optimal'
        let pr = -1
        let best = Infinity
        for (let i = 0; i < m; i++) {
          const v = T[i][pc]
          if (v > EPS) {
            const ratio = T[i][total] / v
            if (ratio < best - 1e-12 || (Math.abs(ratio - best) <= 1e-12 && basis[i] < basis[pr])) {
              best = ratio
              pr = i
            }
          }
        }
        if (pr < 0) return 'unbounded'
        pivot(pr, pc)
      }
    } finally {
      T.pop()
    }
  }
  if (nArt) {
    const cost = new Array(total).fill(0)
    for (const j of artificial) cost[j] = 1
    const st = run(cost, () => true)
    if (st === 'limit') return { status: 'limit', x: [], z: NaN, iterations }
    const infeas = basis.reduce((acc, b, i) => acc + (artificial.has(b) ? T[i][total] : 0), 0)
    if (infeas > 1e-7 * Math.max(1, ...rows.map((r) => Math.abs(r.b)))) return { status: 'infeasible', x: [], z: NaN, iterations }
    // drive remaining artificials out of the basis
    for (let i = 0; i < m; i++)
      if (artificial.has(basis[i])) {
        const pc = T[i].findIndex((v, j) => j < total && !artificial.has(j) && Math.abs(v) > 1e-9)
        if (pc >= 0) pivot(i, pc)
      }
  }
  const cost = new Array(total).fill(0)
  cy.forEach((v, j) => (cost[j] = lp.maximize ? -v : v))
  const st = run(cost, (j) => !artificial.has(j))
  if (st !== 'optimal') return { status: st, x: [], z: NaN, iterations }
  const y = new Array(total).fill(0)
  basis.forEach((b, i) => (y[b] = T[i][total]))
  const x = maps.map((mp) => mp.shift + mp.cols.reduce((acc, col, k) => acc + mp.signs[k] * y[col], 0))
  const z = c0 + cy.reduce((acc, v, j) => acc + v * y[j], 0)
  return { status: 'optimal', x, z, iterations }
}

/** Branch and bound over a relaxation solver (for integer / binary variables). */
async function branchAndBound(
  relax: (lower: number[], upper: number[]) => Promise<{ status: string; x: number[]; z: number }>,
  lower: number[],
  upper: number[],
  ints: Set<number>,
  maximize: boolean,
  limits: { nodes: number; deadline: number },
): Promise<{ status: 'optimal' | 'infeasible' | 'unbounded' | 'limit'; x: number[]; z: number; nodes: number }> {
  let best: { x: number[]; z: number } | null = null
  let nodes = 0
  let hitLimit = false
  let unbounded = false
  const better = (z: number, ref: number) => (maximize ? z > ref + 1e-9 * Math.max(1, Math.abs(ref)) : z < ref - 1e-9 * Math.max(1, Math.abs(ref)))
  const stack: { lo: number[]; up: number[] }[] = [{ lo: lower, up: upper }]
  while (stack.length) {
    if (++nodes > limits.nodes || Date.now() > limits.deadline) {
      hitLimit = true
      break
    }
    if (nodes % 25 === 0) await tick()
    const node = stack.pop()!
    const r = await relax(node.lo, node.up)
    if (r.status === 'unbounded') {
      unbounded = true
      continue
    }
    if (r.status !== 'optimal') continue
    if (best && !better(r.z, best.z)) continue
    // most fractional integer variable
    let bj = -1
    let frac = 0
    for (const j of ints) {
      const f = Math.abs(r.x[j] - Math.round(r.x[j]))
      if (f > 1e-6 && f > frac) {
        frac = f
        bj = j
      }
    }
    if (bj < 0) {
      best = { x: r.x.map((v, j) => (ints.has(j) ? Math.round(v) : v)), z: r.z }
      continue
    }
    const v = r.x[bj]
    const down = { lo: node.lo, up: node.up.map((u, j) => (j === bj ? Math.floor(v) : u)) }
    const upB = { lo: node.lo.map((l, j) => (j === bj ? Math.ceil(v) : l)), up: node.up }
    // explore the nearer branch first
    if (v - Math.floor(v) > 0.5) stack.push(down, upB)
    else stack.push(upB, down)
  }
  if (best) return { status: hitLimit ? 'limit' : 'optimal', ...best, nodes }
  return { status: hitLimit ? 'limit' : unbounded ? 'unbounded' : 'infeasible', x: [], z: NaN, nodes }
}

// ---------------------------------------------------------------- nonlinear (GRG-style)
interface NlpProblem {
  n: number
  lower: number[]
  upper: number[]
  /** objective to minimise (already signed/scaled) and constraint values: ineq g ≤ 0, eq h = 0 */
  evaluate: (x: number[]) => { f: number; g: number[]; h: number[] }
}

function clamp(x: number[], lo: number[], up: number[]) {
  return x.map((v, j) => Math.min(up[j], Math.max(lo[j], v)))
}

async function minimizeNLP(p: NlpProblem, x0: number[], opts: { precision: number; deadline: number; maxIter: number; hooks?: SolveHooks; started: number }) {
  let x = clamp(x0, p.lower, p.upper)
  const first = p.evaluate(x)
  let lamG = new Array(first.g.length).fill(0)
  let lamH = new Array(first.h.length).fill(0)
  let mu = 10
  let iterations = 0
  let stopped = false
  const viol = (e: { g: number[]; h: number[] }) => Math.max(0, ...e.g.map((v) => Math.max(0, v)), ...e.h.map(Math.abs))
  const L = (x: number[]) => {
    const e = p.evaluate(x)
    if (!Number.isFinite(e.f) || e.g.some((v) => !Number.isFinite(v)) || e.h.some((v) => !Number.isFinite(v))) return Infinity
    let v = e.f
    e.g.forEach((g, i) => {
      const t = Math.max(0, g + lamG[i] / mu)
      v += (mu / 2) * t * t - (lamG[i] * lamG[i]) / (2 * mu)
    })
    e.h.forEach((h, i) => (v += lamH[i] * h + (mu / 2) * h * h))
    return v
  }
  const grad = (x: number[], fx: number) => {
    const g = new Array(p.n).fill(0)
    for (let j = 0; j < p.n; j++) {
      const h = 1e-6 * Math.max(1, Math.abs(x[j]))
      const xp = [...x]
      const forward = x[j] + h <= p.upper[j]
      xp[j] = forward ? x[j] + h : x[j] - h
      const fp = L(xp)
      g[j] = Number.isFinite(fp) ? (forward ? fp - fx : fx - fp) / h : 0
    }
    return g
  }
  let lastViol = viol(first)
  for (let outer = 0; outer < 40 && !stopped; outer++) {
    // ---- inner: projected BFGS on L
    let fx = L(x)
    if (!Number.isFinite(fx)) break
    let g = grad(x, fx)
    // first step: about 10% of the variables' size, so a steep start doesn't jump straight to a bound
    const gmax = Math.max(1e-12, ...g.map(Math.abs))
    const h0 = Math.min(1, (0.1 * Math.max(1, ...x.map(Math.abs))) / gmax)
    const identity = () => Array.from({ length: p.n }, (_, i) => Array.from({ length: p.n }, (_, j) => (i === j ? h0 : 0)))
    let H: number[][] = identity()
    let stall = 0
    for (let it = 0; it < opts.maxIter; it++) {
      iterations++
      if (Date.now() > opts.deadline) {
        stopped = true
        break
      }
      if (iterations % 10 === 0) {
        await tick()
        if (opts.hooks?.onProgress?.({ iteration: iterations, objective: p.evaluate(x).f, seconds: (Date.now() - opts.started) / 1000 }) === false) {
          stopped = true
          break
        }
      }
      // free variables: not stuck at a bound with the gradient pushing outwards
      const free = x.map((v, j) => !((v <= p.lower[j] + 1e-12 && g[j] > 0) || (v >= p.upper[j] - 1e-12 && g[j] < 0)))
      const pg = g.map((v, j) => (free[j] ? v : 0))
      if (Math.max(...pg.map(Math.abs)) < 1e-9 * (1 + Math.abs(fx))) break
      let d = H.map((row, i) => (free[i] ? -row.reduce((acc, h, j) => acc + (free[j] ? h * pg[j] : 0), 0) : 0))
      let slope = d.reduce((acc, v, j) => acc + v * pg[j], 0)
      if (!(slope < 0)) {
        // not a descent direction: reset to steepest descent
        H = identity()
        d = pg.map((v) => -h0 * v)
        slope = d.reduce((acc, v, j) => acc + v * pg[j], 0)
      }
      let t = 1
      let xn = x
      let fn = fx
      let accepted = false
      for (let ls = 0; ls < 40; ls++) {
        xn = clamp(
          x.map((v, j) => v + t * d[j]),
          p.lower,
          p.upper,
        )
        fn = L(xn)
        const dec = g.reduce((acc, gv, j) => acc + gv * (xn[j] - x[j]), 0)
        if (Number.isFinite(fn) && fn <= fx + 1e-4 * Math.min(0, dec)) {
          accepted = true
          break
        }
        t /= 2
      }
      if (!accepted) break
      const gn = grad(xn, fn)
      const sv = xn.map((v, j) => v - x[j])
      const yv = gn.map((v, j) => v - g[j])
      const sy = sv.reduce((acc, v, j) => acc + v * yv[j], 0)
      if (sy > 1e-12) {
        // BFGS inverse-Hessian update
        const Hy = H.map((row) => row.reduce((acc, h, j) => acc + h * yv[j], 0))
        const yHy = yv.reduce((acc, v, j) => acc + v * Hy[j], 0)
        H = H.map((row, i) => row.map((h, j) => h + ((sy + yHy) * sv[i] * sv[j]) / (sy * sy) - (Hy[i] * sv[j] + sv[i] * Hy[j]) / sy))
      }
      const rel = Math.abs(fx - fn) / Math.max(1, Math.abs(fx))
      x = xn
      g = gn
      fx = fn
      if (rel < opts.precision * 1e-3) {
        if (++stall >= 5) break
      } else stall = 0
    }
    // ---- multiplier update
    const e = p.evaluate(x)
    const v = viol(e)
    lamG = lamG.map((l, i) => Math.max(0, l + mu * e.g[i]))
    lamH = lamH.map((l, i) => l + mu * e.h[i])
    if (v < opts.precision && outer > 0 && Math.abs(v - lastViol) < opts.precision) break
    if (v > 0.25 * lastViol) mu = Math.min(1e10, mu * 10)
    lastViol = v
    if (!e.g.length && !e.h.length) break
  }
  const e = p.evaluate(x)
  return { x, f: e.f, violation: viol(e), iterations, stopped }
}

// ---------------------------------------------------------------- evolutionary
async function evolve(p: NlpProblem, x0: number[], ints: Set<number>, opts: { precision: number; deadline: number; maxIter: number; seed?: number; hooks?: SolveHooks; started: number }) {
  let seed = (opts.seed ?? Date.now()) >>> 0
  const rand = () => {
    seed = (seed + 0x6d2b79f5) >>> 0
    let t = seed
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
  const n = p.n
  const NP = Math.max(20, Math.min(100, 10 * n))
  const fix = (x: number[]) => clamp(x.map((v, j) => (ints.has(j) ? Math.round(v) : v)), p.lower, p.upper)
  const score = (x: number[]) => {
    const e = p.evaluate(x)
    const v = e.g.reduce((a, g) => a + Math.max(0, g), 0) + e.h.reduce((a, h) => a + Math.max(0, Math.abs(h) - opts.precision), 0)
    return { f: Number.isFinite(e.f) ? e.f : Infinity, v: Number.isFinite(v) ? v : Infinity }
  }
  const beats = (a: { f: number; v: number }, b: { f: number; v: number }) => (a.v <= 1e-9 && b.v <= 1e-9 ? a.f < b.f : a.v < b.v)
  const pop: number[][] = [fix(x0)]
  while (pop.length < NP) pop.push(fix(Array.from({ length: n }, (_, j) => p.lower[j] + rand() * (p.upper[j] - p.lower[j]))))
  const fit = pop.map(score)
  let bestI = 0
  fit.forEach((f, i) => beats(f, fit[bestI]) && (bestI = i))
  let lastBest = fit[bestI]
  let stallGens = 0
  let gen = 0
  let stopped = false
  for (; gen < opts.maxIter; gen++) {
    if (Date.now() > opts.deadline) {
      stopped = true
      break
    }
    for (let i = 0; i < NP; i++) {
      let a: number, b: number, c: number
      do a = Math.floor(rand() * NP)
      while (a === i)
      do b = Math.floor(rand() * NP)
      while (b === i || b === a)
      do c = Math.floor(rand() * NP)
      while (c === i || c === a || c === b)
      const F = 0.5 + 0.3 * rand()
      const jr = Math.floor(rand() * n)
      const trial = fix(pop[i].map((v, j) => (rand() < 0.9 || j === jr ? pop[a][j] + F * (pop[b][j] - pop[c][j]) : v)))
      const s = score(trial)
      if (beats(s, fit[i]) || (s.v === fit[i].v && s.f === fit[i].f)) {
        pop[i] = trial
        fit[i] = s
        if (beats(s, fit[bestI])) bestI = i
      }
    }
    const best = fit[bestI]
    const improved = best.v < lastBest.v - 1e-12 || (best.v <= 1e-9 && lastBest.f - best.f > opts.precision * Math.max(1, Math.abs(lastBest.f)))
    if (improved) {
      stallGens = 0
      lastBest = best
    } else if (++stallGens > 60) break
    await tick()
    if (opts.hooks?.onProgress?.({ iteration: gen + 1, objective: best.f, seconds: (Date.now() - opts.started) / 1000 }) === false) {
      stopped = true
      break
    }
  }
  return { x: pop[bestI], f: fit[bestI].f, violation: fit[bestI].v, iterations: gen, stopped }
}

// ---------------------------------------------------------------- driver
const ENGINE_NAME = { simplex: 'Simplex LP', grg: 'GRG Nonlinear', evolutionary: 'Evolutionary' } as const

/** Validates a model and expands it into variable cells and scalar constraint rows. */
export function prepareModel(wb: Workbook, sheetId: string, m: SolverModel) {
  const sheetName = wb.meta(sheetId).name
  const obj = parseRefList(m.objective, sheetName)
  if (typeof obj === 'string') return `Objective: ${obj}`
  if (obj.length !== 1 || obj[0].r1 !== obj[0].r2 || obj[0].c1 !== obj[0].c2) return 'The objective must be a single cell.'
  const objCell = { r: obj[0].r1, c: obj[0].c1 }
  if (!wb.isFormula(sheetId, objCell.r, objCell.c)) return 'The objective cell must contain a formula that depends on the variable cells.'
  const vr = parseRefList(m.variables, sheetName)
  if (typeof vr === 'string') return `Variable cells: ${vr}`
  const varCells = cellsOf(vr)
  if (!varCells.length) return 'Choose at least one variable cell.'
  if (varCells.length > 400) return 'Solver supports up to 400 variable cells.'
  const seen = new Set<string>()
  for (const v of varCells) {
    const k = `${v.r},${v.c}`
    if (seen.has(k)) return `Variable cell ${a1(v)} is listed twice.`
    seen.add(k)
    if (wb.isFormula(sheetId, v.r, v.c)) return `Variable cell ${a1(v).replace(/\$/g, '')} contains a formula. Variable cells must hold constants (or be empty).`
    const raw = wb.getValue(sheetId, v.r, v.c)
    if (raw !== null && raw !== '' && typeof raw !== 'number') return `Variable cell ${a1(v).replace(/\$/g, '')} contains text. Variable cells must hold numbers (or be empty).`
  }
  if (m.goal === 'value' && !Number.isFinite(m.targetValue)) return 'Enter the target value for “Value Of”.'
  const index = new Map(varCells.map((v, i) => [`${v.r},${v.c}`, i]))
  const rows: Row[] = []
  const integerVars = new Set<number>()
  const binaryVars = new Set<number>()
  for (const [ci, con] of m.constraints.entries()) {
    const lhs = parseRefList(con.ref, sheetName)
    if (typeof lhs === 'string') return `Constraint ${ci + 1}: ${lhs}`
    const lc = cellsOf(lhs)
    if (con.op === 'int' || con.op === 'bin') {
      for (const c of lc) {
        const j = index.get(`${c.r},${c.c}`)
        if (j === undefined) return `Constraint ${ci + 1}: ${con.op === 'int' ? 'integer' : 'binary'} constraints must refer to variable cells (${a1(c).replace(/\$/g, '')} is not one).`
        integerVars.add(j)
        if (con.op === 'bin') binaryVars.add(j)
      }
      continue
    }
    const rhsText = (con.rhs ?? '').trim()
    if (!rhsText) return `Constraint ${ci + 1}: enter a value or cell for the right-hand side.`
    let rhsFor: (k: number) => Row['rhs']
    const asNum = Number(rhsText)
    if (rhsText !== '' && Number.isFinite(asNum)) rhsFor = () => ({ value: asNum })
    else {
      const rr = parseRefList(rhsText, sheetName)
      if (typeof rr === 'string') return `Constraint ${ci + 1}: ${rr}`
      const rc = cellsOf(rr)
      if (rc.length === 1) rhsFor = () => ({ cell: rc[0] })
      else if (rc.length === lc.length) rhsFor = (k) => ({ cell: rc[k] })
      else return `Constraint ${ci + 1}: the right-hand side must be a number, one cell, or a range the same size as ${con.ref}.`
    }
    lc.forEach((c, k) => rows.push({ ci, label: `${a1(c)} ${con.op} ${rhsText}`, lhs: c, op: con.op as Row['op'], rhs: rhsFor(k) }))
  }
  return { objCell, varCells, rows, integerVars, binaryVars, index }
}

export async function solve(wb: Workbook, sheetId: string, m: SolverModel, hooks?: SolveHooks): Promise<SolverResult> {
  const started = Date.now()
  const engine = ENGINE_NAME[m.method]
  const prep = prepareModel(wb, sheetId, m)
  if (typeof prep === 'string') return { ok: false, message: prep, engine, iterations: 0, seconds: 0 }
  const { objCell, varCells, rows, integerVars, binaryVars, index } = prep
  const n = varCells.length
  const precision = m.options?.precision ?? 1e-6
  const maxTime = (m.options?.maxTimeSec ?? (m.method === 'evolutionary' ? 30 : 60)) * 1000
  const deadline = started + maxTime
  const maxIter = m.options?.maxIterations ?? (m.method === 'evolutionary' ? 1000 : 500)
  const original = varCells.map((v) => {
    const x = wb.getValue(sheetId, v.r, v.c)
    return typeof x === 'number' ? x : 0
  })
  const ov = wb.getValue(sheetId, objCell.r, objCell.c)
  const originalObjective = typeof ov === 'number' ? ov : NaN
  const probe = new Probe(wb, sheetId, varCells)
  const evalAt = (x: number[]): Eval => {
    probe.set(x)
    return {
      f: probe.num(objCell),
      d: rows.map((rw) => probe.num(rw.lhs) - ('value' in rw.rhs ? rw.rhs.value : probe.num(rw.rhs.cell))),
    }
  }
  const finish = (r: Omit<SolverResult, 'engine' | 'seconds' | 'details'> & { details?: Partial<SolverResult['details']> }): SolverResult => {
    probe.destroy()
    return {
      ...r,
      engine,
      seconds: (Date.now() - started) / 1000,
      details: { varCells, original, originalObjective, rows, integerVars, linear: m.method === 'simplex', ...r.details },
    }
  }
  try {
    // bounds: non-negativity, binary, and simple constraints on variable cells with constant right-hand sides
    const lower = new Array(n).fill(m.nonNegative ? 0 : -Infinity)
    const upper = new Array(n).fill(Infinity)
    for (const j of binaryVars) {
      lower[j] = Math.max(lower[j], 0)
      upper[j] = Math.min(upper[j], 1)
    }
    const general: number[] = [] // row indices that are not simple bounds
    rows.forEach((rw, i) => {
      const j = index.get(`${rw.lhs.r},${rw.lhs.c}`)
      const constRhs = 'value' in rw.rhs ? rw.rhs.value : !wb.isFormula(sheetId, rw.rhs.cell.r, rw.rhs.cell.c) && typeof wb.getValue(sheetId, rw.rhs.cell.r, rw.rhs.cell.c) === 'number' ? (wb.getValue(sheetId, rw.rhs.cell.r, rw.rhs.cell.c) as number) : null
      if (j !== undefined && constRhs !== null && m.method !== 'simplex') {
        if (rw.op !== '>=') upper[j] = Math.min(upper[j], constRhs)
        if (rw.op !== '<=') lower[j] = Math.max(lower[j], constRhs)
      } else general.push(i)
    })
    const base = evalAt(original)
    if (!Number.isFinite(base.f)) {
      const v = probe.hf.getCellValue({ sheet: probe.sheet, row: objCell.r, col: objCell.c })
      if (typeof v === 'string') return finish({ ok: false, message: 'The objective cell must contain a number (it currently shows text).', iterations: 0 })
    }
    const sign = m.goal === 'max' ? -1 : 1
    const target = m.goal === 'value' ? (m.targetValue as number) : 0

    // ================================================================ Simplex LP
    if (m.method === 'simplex') {
      // probe linear coefficients: f(x) = f0 + Σ a_j x_j
      const zero = new Array(n).fill(0)
      const e0 = evalAt(zero)
      const cols: Eval[] = []
      for (let j = 0; j < n; j++) {
        const x = [...zero]
        x[j] = 1
        cols.push(evalAt(x))
      }
      const coefF = cols.map((e) => e.f - e0.f)
      const coefD = rows.map((_, i) => cols.map((e) => e.d[i] - e0.d[i]))
      const bad = [e0.f, ...e0.d, ...coefF, ...coefD.flat()].some((v) => !Number.isFinite(v))
      // linearity check at two test points
      let linear = !bad
      for (const pt of [original.map((v, j) => v + 1.37 + j * 0.11), Array.from({ length: n }, (_, j) => 3.1 + ((j * 7) % 5))]) {
        if (!linear) break
        const e = evalAt(pt)
        const near = (got: number, c0: number, cs: number[]) => {
          const pred = c0 + cs.reduce((a, c, j) => a + c * pt[j], 0)
          const scale = Math.max(1, Math.abs(got), Math.abs(pred), ...cs.map((c, j) => Math.abs(c * pt[j])))
          return Math.abs(got - pred) <= 1e-7 * scale
        }
        if (!near(e.f, e0.f, coefF) || rows.some((_, i) => !near(e.d[i], e0.d[i], coefD[i]))) linear = false
      }
      if (!linear) return finish({ ok: false, message: 'The linearity conditions required by this LP Solver are not satisfied. Try the GRG Nonlinear method.', iterations: 0 })
      const lp: LP = {
        c: m.goal === 'value' ? new Array(n).fill(0) : coefF,
        A: coefD.map((r) => [...r]),
        b: rows.map((_, i) => -e0.d[i]),
        ops: rows.map((r) => r.op),
        lower: [...lower],
        upper: [...upper],
        maximize: m.goal === 'max',
      }
      if (m.goal === 'value') {
        lp.A.push([...coefF])
        lp.b.push(target - e0.f)
        lp.ops.push('=')
      }
      let res: { status: string; x: number[]; z: number }
      let iterations = 0
      if (integerVars.size) {
        const bb = await branchAndBound(
          async (lo, up) => {
            const r = solveLP({ ...lp, lower: lo, upper: up })
            iterations += r.iterations
            return r
          },
          lp.lower,
          lp.upper,
          integerVars,
          lp.maximize,
          { nodes: 50_000, deadline },
        )
        res = bb
      } else {
        const r = solveLP(lp)
        iterations = r.iterations
        res = r.status === 'limit' ? { ...r, status: 'limit' } : r
      }
      if (res.status === 'infeasible') return finish({ ok: false, message: 'Solver could not find a feasible solution.', iterations })
      if (res.status === 'unbounded') return finish({ ok: false, message: 'The Objective Cell values do not converge.', iterations })
      if (res.status === 'limit' && !res.x.length) return finish({ ok: false, message: 'Solver stopped at the iteration or time limit without finding a solution.', iterations })
      const x = res.x.map((v) => (Math.abs(v) < 1e-11 ? 0 : +v.toPrecision(15)))
      const fx = evalAt(x).f
      // sensitivity (continuous models only): shadow prices and reduced costs by re-solving
      let shadow: number[] | undefined
      let reduced: number[] | undefined
      if (!integerVars.size && m.goal !== 'value') {
        const z0 = lp.c.reduce((a, c, j) => a + c * x[j], 0)
        shadow = rows.map((_, i) => {
          const dlt = 1e-3 * Math.max(1, Math.abs(lp.b[i]))
          const r = solveLP({ ...lp, b: lp.b.map((b, k) => (k === i ? b + dlt : b)) })
          return r.status === 'optimal' ? +((r.z - z0) / dlt).toPrecision(10) : 0
        })
        reduced = x.map((xv, j) => {
          const atBound = Math.abs(xv - lp.lower[j]) < 1e-9 || Math.abs(xv - lp.upper[j]) < 1e-9
          if (!atBound) return 0
          const dlt = 1e-3
          const atLower = Math.abs(xv - lp.lower[j]) < 1e-9
          const r = solveLP({ ...lp, lower: lp.lower.map((l, k) => (k === j && atLower ? xv + dlt : l)), upper: lp.upper.map((u, k) => (k === j && !atLower ? xv - dlt : u)) })
          return r.status === 'optimal' ? +(((r.z - z0) / dlt) * (atLower ? 1 : -1)).toPrecision(10) : 0
        })
      }
      const msg = res.status === 'limit' ? 'Solver stopped at the time or node limit. The best integer solution found is shown.' : integerVars.size ? 'Solver found an integer solution within tolerance. All Constraints are satisfied.' : 'Solver found a solution. All Constraints and optimality conditions are satisfied.'
      return finish({ ok: true, message: msg, x, objective: fx, feasible: true, iterations, details: { shadow, reduced, objCoef: coefF, linear: true } })
    }

    // ================================================================ GRG / Evolutionary
    const fScale = 1 / Math.max(1, Math.abs(Number.isFinite(base.f) ? base.f - target : 1))
    const genRows = general.map((i) => rows[i])
    const problem: NlpProblem = {
      n,
      lower,
      upper,
      evaluate: (x) => {
        const e = evalAt(x)
        const g: number[] = []
        const h: number[] = []
        general.forEach((ri, k) => {
          const d = e.d[ri]
          const op = genRows[k].op
          if (op === '<=') g.push(d)
          else if (op === '>=') g.push(-d)
          else h.push(d)
        })
        if (m.goal === 'value') h.push((e.f - target) * fScale)
        return { f: m.goal === 'value' ? 0 : sign * e.f * fScale, g, h }
      },
    }
    const start = clamp(original, lower, upper)
    const tol = Math.max(precision, 1e-8)
    let best: { x: number[]; f: number; violation: number; iterations: number; stopped: boolean }
    if (m.method === 'evolutionary') {
      const unb = varCells.find((_, j) => !Number.isFinite(lower[j]) || !Number.isFinite(upper[j]))
      if (unb) return finish({ ok: false, message: `All variable cells need upper and lower bounds for the Evolutionary method (add constraints such as ${a1(unb).replace(/\$/g, '')} <= 100 and ${a1(unb).replace(/\$/g, '')} >= 0).`, iterations: 0 })
      best = await evolve(problem, start, integerVars, { precision: tol, deadline, maxIter, seed: m.options?.seed, hooks, started })
      if (!integerVars.size && !best.stopped) {
        const pol = await minimizeNLP(problem, best.x, { precision: tol, deadline, maxIter: 200, hooks, started })
        if (pol.violation <= Math.max(best.violation, tol) && pol.f <= best.f) best = { ...pol, iterations: best.iterations + pol.iterations }
      }
    } else {
      const runFrom = (x0: number[], lo = lower, up = upper) => minimizeNLP({ ...problem, lower: lo, upper: up }, x0, { precision: tol, deadline, maxIter, hooks, started })
      if (integerVars.size) {
        const bb = await branchAndBound(
          async (lo, up) => {
            const r = await runFrom(clamp(start, lo, up), lo, up)
            return { status: r.violation <= Math.max(1e-6, tol * 10) ? 'optimal' : 'infeasible', x: r.x, z: r.f }
          },
          lower,
          upper,
          integerVars,
          false,
          { nodes: 400, deadline },
        )
        if (!bb.x.length) return finish({ ok: false, message: bb.status === 'limit' ? 'Solver stopped at the time limit before finding an integer solution.' : 'Solver could not find a feasible integer solution.', iterations: bb.nodes })
        const e = problem.evaluate(bb.x)
        best = { x: bb.x, f: e.f, violation: Math.max(0, ...e.g, ...e.h.map(Math.abs)), iterations: bb.nodes, stopped: bb.status === 'limit' }
      } else {
        best = await runFrom(start)
        // ended somewhere infeasible (e.g. a flat spot)? retry from a few nearby points
        for (let k = 0; best.violation > tol * 10 && k < 4 && Date.now() < deadline; k++) {
          const f = [0.5, 1.5, 0.1, 3][k]
          const r = await runFrom(start.map((v, j) => (v === 0 ? (Number.isFinite(upper[j]) ? (lower[j] + upper[j]) / 2 || 1 : 1) * f : v * f)))
          if (r.violation < best.violation) best = { ...r, iterations: best.iterations + r.iterations }
        }
        if (m.options?.multistart) {
          const bounded = lower.every(Number.isFinite) && upper.every(Number.isFinite)
          let seed = (m.options.seed ?? 12345) >>> 0
          const rnd = () => ((seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296)
          for (let k = 0; k < 10 + n && Date.now() < deadline; k++) {
            const x0 = lower.map((l, j) => (bounded ? l + rnd() * (upper[j] - l) : start[j] + (rnd() - 0.5) * 2 * Math.max(10, Math.abs(start[j]) * 2)))
            const r = await runFrom(x0)
            const ok = r.violation <= tol * 10
            if ((ok && (best.violation > tol * 10 || r.f < best.f)) || (!ok && r.violation < best.violation)) best = { ...r, iterations: best.iterations + r.iterations }
          }
        }
      }
    }
    const x = best.x.map((v, j) => (integerVars.has(j) ? Math.round(v) : Math.abs(v) < 1e-12 ? 0 : +v.toPrecision(12)))
    const e = evalAt(x)
    const feasibleTol = Math.max(1e-6, precision * 10)
    const maxViol = Math.max(
      0,
      ...rows.map((rw, i) => {
        const d = e.d[i]
        const scale = Math.max(1, Math.abs(probe.num(rw.lhs)))
        return (rw.op === '<=' ? Math.max(0, d) : rw.op === '>=' ? Math.max(0, -d) : Math.abs(d)) / scale
      }),
      ...x.map((v, j) => Math.max(0, lower[j] - v, v - upper[j])),
      m.goal === 'value' ? Math.abs(e.f - target) / Math.max(1, Math.abs(target)) : 0,
    )
    if (!Number.isFinite(e.f)) return finish({ ok: false, message: 'Solver encountered an error value in the objective or a constraint cell.', iterations: best.iterations })
    if (maxViol > feasibleTol) return finish({ ok: false, message: 'Solver could not find a feasible solution.', x, objective: e.f, feasible: false, iterations: best.iterations })
    let msg: string
    if (best.stopped) msg = 'Solver stopped at the time limit. The best solution found is shown. All Constraints are satisfied.'
    else if (m.method === 'evolutionary') msg = 'Solver cannot improve the current solution. All Constraints are satisfied.'
    else if (integerVars.size) msg = 'Solver found an integer solution within tolerance. All Constraints are satisfied.'
    else msg = 'Solver found a solution. All Constraints and optimality conditions are satisfied.'
    return finish({ ok: true, message: msg, x, objective: e.f, feasible: true, iterations: best.iterations, details: { linear: false } })
  } catch (err) {
    try {
      probe.destroy()
    } catch {
      /* already gone */
    }
    return { ok: false, message: `Solver stopped: ${(err as Error).message}`, engine, iterations: 0, seconds: (Date.now() - started) / 1000 }
  }
}

/** Writes the solution into the variable cells (one undo step) and remembers the model on the sheet. */
export function applySolution(wb: Workbook, sheetId: string, res: SolverResult) {
  if (!res.x || !res.details) return
  wb.setInputs(
    sheetId,
    res.details.varCells.map((v, i) => ({ r: v.r, c: v.c, input: String(res.x![i]) })),
    'Solver',
  )
}

export function saveModel(wb: Workbook, sheetId: string, model: SolverModel) {
  wb.run({ label: 'Save Solver model', dirty: [sheetId], metaScope: [sheetId], metaOps: () => (wb.meta(sheetId).solver = model) })
}

// ---------------------------------------------------------------- reports
/** Excel-style "name" for a cell: nearest text to the left plus nearest header above. */
export function cellLabel(wb: Workbook, sheetId: string, p: Cell): string {
  let left = ''
  for (let c = p.c - 1; c >= 0 && p.c - c <= 30; c--) {
    const v = wb.getValue(sheetId, p.r, c)
    if (typeof v === 'string' && v.trim()) {
      left = v.trim()
      break
    }
  }
  let top = ''
  for (let r = p.r - 1; r >= 0 && p.r - r <= 200; r--) {
    const v = wb.getValue(sheetId, r, p.c)
    if (typeof v === 'string' && v.trim()) {
      top = v.trim()
      break
    }
  }
  return [left, top].filter(Boolean).join(' ')
}

const ref = (p: Cell) => a1(p)

export function answerReport(wb: Workbook, sheetId: string, model: SolverModel, res: SolverResult): OutTable {
  const d = res.details!
  const sheetName = wb.meta(sheetId).name
  const obj = parseRefList(model.objective, sheetName) as Range[]
  const objCell = { r: obj[0].r1, c: obj[0].c1 }
  const t = new TableBuilder()
  t.title('ExcelPro Solver — Answer Report')
  t.row([`Worksheet: [${wb.name}]${sheetName}`])
  t.row([`Report created: ${new Date().toLocaleString()}`])
  t.row([`Result: ${res.message}`])
  t.row([`Solver Engine: ${res.engine}. Solution time ${res.seconds.toFixed(3)} seconds; ${res.iterations} iterations.`])
  t.blank()
  t.title(`Objective Cell (${model.goal === 'max' ? 'Max' : model.goal === 'min' ? 'Min' : `Value Of ${model.targetValue}`})`)
  t.header(['Cell', 'Name', 'Original Value', 'Final Value'])
  t.row([ref(objCell), cellLabel(wb, sheetId, objCell), d.originalObjective, res.objective ?? null]).end()
  t.blank()
  t.title('Variable Cells')
  t.header(['Cell', 'Name', 'Original Value', 'Final Value', 'Integer'])
  d.varCells.forEach((v, i) => t.row([ref(v), cellLabel(wb, sheetId, v), d.original[i], res.x?.[i] ?? null, d.integerVars.has(i) ? 'Integer' : 'Contin']))
  t.end()
  t.blank()
  t.title('Constraints')
  t.header(['Cell', 'Name', 'Cell Value', 'Formula', 'Status', 'Slack'])
  for (const rw of d.rows) {
    const lhs = wb.getValue(sheetId, rw.lhs.r, rw.lhs.c)
    const rhs = 'value' in rw.rhs ? rw.rhs.value : (wb.getValue(sheetId, rw.rhs.cell.r, rw.rhs.cell.c) as number)
    const lv = typeof lhs === 'number' ? lhs : NaN
    const slack = Math.abs(lv - rhs)
    const binding = rw.op === '=' || slack <= 1e-6 * Math.max(1, Math.abs(rhs))
    const rhsText = 'value' in rw.rhs ? String(rw.rhs.value) : ref(rw.rhs.cell)
    t.row([ref(rw.lhs), cellLabel(wb, sheetId, rw.lhs), Number.isFinite(lv) ? lv : null, `${ref(rw.lhs)}${rw.op}${rhsText}`, binding ? 'Binding' : 'Not Binding', binding ? 0 : +slack.toPrecision(12)])
  }
  d.integerVars.forEach((j) => {
    const v = d.varCells[j]
    const isBin = model.constraints.some((c) => c.op === 'bin' && cellsOf(parseRefList(c.ref, sheetName) as Range[]).some((q) => q.r === v.r && q.c === v.c))
    t.row([ref(v), cellLabel(wb, sheetId, v), res.x?.[j] ?? null, `${ref(v)}=${isBin ? 'Binary' : 'Integer'}`, 'Binding', 0])
  })
  t.end()
  return t.build()
}

export function sensitivityReport(wb: Workbook, sheetId: string, res: SolverResult): OutTable | null {
  const d = res.details
  if (!d || !d.linear || !d.shadow || !d.reduced) return null
  const sheetName = wb.meta(sheetId).name
  const t = new TableBuilder()
  t.title('ExcelPro Solver — Sensitivity Report')
  t.row([`Worksheet: [${wb.name}]${sheetName}`])
  t.row([`Report created: ${new Date().toLocaleString()}`])
  t.blank()
  t.title('Variable Cells')
  t.header(['Cell', 'Name', 'Final Value', 'Reduced Cost', 'Objective Coefficient'])
  d.varCells.forEach((v, i) => t.row([ref(v), cellLabel(wb, sheetId, v), res.x?.[i] ?? null, d.reduced![i], d.objCoef?.[i] ?? null]))
  t.end()
  t.blank()
  t.title('Constraints')
  t.header(['Cell', 'Name', 'Final Value', 'Shadow Price', 'Constraint R.H. Side'])
  d.rows.forEach((rw, i) => {
    const lhs = wb.getValue(sheetId, rw.lhs.r, rw.lhs.c)
    const rhs = 'value' in rw.rhs ? rw.rhs.value : wb.getValue(sheetId, rw.rhs.cell.r, rw.rhs.cell.c)
    t.row([ref(rw.lhs), cellLabel(wb, sheetId, rw.lhs), typeof lhs === 'number' ? lhs : null, d.shadow![i], typeof rhs === 'number' ? rhs : null])
  })
  t.end()
  t.blank()
  t.row(['Shadow price: how much the objective improves per unit increase of the constraint’s right-hand side.'])
  t.row(['Reduced cost: how much the objective changes per unit a variable at zero (or its bound) is forced into the solution.'])
  return t.build()
}

export type { SolverConstraint }
