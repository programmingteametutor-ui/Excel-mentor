/**
 * ExcelPro function plugin, part 3 (data analysis): LINEST, LOGEST, TREND, GROWTH, FORECAST,
 * FORECAST.LINEAR, INTERCEPT, KURT, RANK.AVG, PERCENTRANK(.INC/.EXC), TRIMMEAN, MINVERSE, MDETERM,
 * RANDARRAY, PERMUT, PROB.
 */
import { ArraySize, CellError, ErrorType, FunctionArgumentType, FunctionPlugin, HyperFormula, SimpleRangeValue } from 'hyperformula'
import type { ImplementedFunctions } from 'hyperformula'
import { determinant, kurtosis, matInverse, ols, percentRank } from './stats'

type Ast = any // eslint-disable-line @typescript-eslint/no-explicit-any
type State = any // eslint-disable-line @typescript-eslint/no-explicit-any
type V = unknown

const RANGE = { argumentType: FunctionArgumentType.RANGE }
const OPT_RANGE = { argumentType: FunctionArgumentType.RANGE, optionalArg: true }
const NUM = { argumentType: FunctionArgumentType.NUMBER }
const OPT_BOOL_TRUE = { argumentType: FunctionArgumentType.BOOLEAN, optionalArg: true, defaultValue: true }
const OPT_BOOL_FALSE = { argumentType: FunctionArgumentType.BOOLEAN, optionalArg: true, defaultValue: false }

const err = (t: ErrorType, msg?: string) => new CellError(t, msg)
const grid = (r: SimpleRangeValue) => r.data as V[][]
const numbers = (r: SimpleRangeValue) => grid(r).flat().filter((v): v is number => typeof v === 'number')

/** Literal value of an argument when it is written directly in the formula (number or TRUE/FALSE). */
function literal(ast: Ast): number | boolean | undefined {
  if (!ast) return undefined
  if (ast.type === 'NUMBER') return ast.value
  if (ast.type === 'MINUS_UNARY_OP' && ast.value?.type === 'NUMBER') return -ast.value.value
  if (ast.type === 'FUNCTION_CALL' && (ast.procedureName === 'TRUE' || ast.procedureName === 'FALSE') && !ast.args?.length) return ast.procedureName === 'TRUE'
  return undefined
}

/**
 * Splits known_y / known_x into vectors. Returns the y values and one array per x variable.
 * Like Excel: a single-column y pairs with x columns; a single-row y pairs with x rows.
 */
function regressionData(yR: SimpleRangeValue, xR: SimpleRangeValue | undefined): { y: number[]; xs: number[][]; byCol: boolean } | CellError {
  const yd = grid(yR)
  const byCol = yR.width() === 1
  const y = (byCol ? yd.map((r) => r[0]) : yd[0]) as V[]
  if (yR.width() !== 1 && yR.height() !== 1) return err(ErrorType.REF, 'known_y’s must be a single row or column.')
  if (y.some((v) => typeof v !== 'number')) return err(ErrorType.VALUE, 'known_y’s must contain only numbers.')
  const n = y.length
  let xs: V[][]
  if (!xR) xs = [Array.from({ length: n }, (_, i) => i + 1)]
  else {
    const xd = grid(xR)
    if (byCol) {
      if (xR.height() !== n) return err(ErrorType.REF, 'known_x’s must have as many rows as known_y’s.')
      xs = xd[0].map((_, j) => xd.map((r) => r[j]))
    } else {
      if (xR.width() !== n) return err(ErrorType.REF, 'known_x’s must have as many columns as known_y’s.')
      xs = xd
    }
  }
  if (xs.flat().some((v) => typeof v !== 'number')) return err(ErrorType.VALUE, 'known_x’s must contain only numbers.')
  return { y: y as number[], xs: xs as number[][], byCol }
}

/** Number of x variables, from the argument shapes (for sizing LINEST/LOGEST results). */
function xCount(self: FunctionPlugin, ast: Ast, state: State): number {
  const y = (self as any).arraySizeForAst(ast.args[0], state) // eslint-disable-line @typescript-eslint/no-explicit-any
  if (!ast.args[1] || ast.args[1].type === 'EMPTY') return 1
  const x = (self as any).arraySizeForAst(ast.args[1], state) // eslint-disable-line @typescript-eslint/no-explicit-any
  return y.width === 1 ? x.width : x.height
}

export class ExcelProFunctions3 extends FunctionPlugin {
  static implementedFunctions: ImplementedFunctions = {
    LINEST: { method: 'linest', sizeOfResultArrayMethod: 'linestSize', parameters: [RANGE, OPT_RANGE, OPT_BOOL_TRUE, OPT_BOOL_FALSE] },
    LOGEST: { method: 'logest', sizeOfResultArrayMethod: 'linestSize', parameters: [RANGE, OPT_RANGE, OPT_BOOL_TRUE, OPT_BOOL_FALSE] },
    TREND: { method: 'trend', sizeOfResultArrayMethod: 'trendSize', parameters: [RANGE, OPT_RANGE, OPT_RANGE, OPT_BOOL_TRUE] },
    GROWTH: { method: 'growth', sizeOfResultArrayMethod: 'trendSize', parameters: [RANGE, OPT_RANGE, OPT_RANGE, OPT_BOOL_TRUE] },
    FORECAST: { method: 'forecast', parameters: [NUM, RANGE, RANGE] },
    'FORECAST.LINEAR': { method: 'forecast', parameters: [NUM, RANGE, RANGE] },
    INTERCEPT: { method: 'intercept', parameters: [RANGE, RANGE] },
    KURT: { method: 'kurt', parameters: [RANGE], repeatLastArgs: 1 },
    'RANK.AVG': { method: 'rankavg', parameters: [NUM, RANGE, { argumentType: FunctionArgumentType.INTEGER, optionalArg: true, defaultValue: 0 }] },
    'PERCENTRANK.INC': { method: 'percentrankinc', parameters: [RANGE, NUM, { argumentType: FunctionArgumentType.INTEGER, optionalArg: true, defaultValue: 3 }] },
    PERCENTRANK: { method: 'percentrankinc', parameters: [RANGE, NUM, { argumentType: FunctionArgumentType.INTEGER, optionalArg: true, defaultValue: 3 }] },
    'PERCENTRANK.EXC': { method: 'percentrankexc', parameters: [RANGE, NUM, { argumentType: FunctionArgumentType.INTEGER, optionalArg: true, defaultValue: 3 }] },
    TRIMMEAN: { method: 'trimmean', parameters: [RANGE, NUM] },
    MINVERSE: { method: 'minverse', sizeOfResultArrayMethod: 'sameAsFirst', parameters: [RANGE] },
    MDETERM: { method: 'mdeterm', parameters: [RANGE] },
    RANDARRAY: {
      method: 'randarray',
      sizeOfResultArrayMethod: 'randarraySize',
      isVolatile: true,
      parameters: [
        { argumentType: FunctionArgumentType.INTEGER, optionalArg: true, defaultValue: 1 },
        { argumentType: FunctionArgumentType.INTEGER, optionalArg: true, defaultValue: 1 },
        { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 0 },
        { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 1 },
        OPT_BOOL_FALSE,
      ],
    },
    PERMUT: { method: 'permut', parameters: [NUM, NUM] },
    PROB: { method: 'prob', parameters: [RANGE, RANGE, NUM, { argumentType: FunctionArgumentType.NUMBER, optionalArg: true }] },
  }

  // ------------------------------------------------------------ regression
  private regress(ast: Ast, state: State, name: string, log: boolean) {
    return this.runFunction(ast.args, state, this.metadata(name), (yR: SimpleRangeValue, xR: SimpleRangeValue | undefined, constant: boolean, withStats: boolean) => {
      const d = regressionData(yR, xR)
      if (d instanceof CellError) return d
      if (log && d.y.some((v) => v <= 0)) return err(ErrorType.NUM, 'LOGEST needs known_y’s greater than 0.')
      const res = ols(log ? d.y.map(Math.log) : d.y, d.xs, constant)
      if (!res) return err(ErrorType.NUM, 'The regression could not be calculated.')
      const k = d.xs.length
      const tf = (v: number) => (log ? Math.exp(v) : v)
      // Excel order: m_k ... m_1, b
      const row1 = [...res.coef.slice(1).reverse().map(tf), constant ? tf(res.coef[0]) : log ? 1 : 0]
      if (!withStats) return SimpleRangeValue.onlyValues([row1] as never)
      const na = err(ErrorType.NA)
      const pad = (a: V[]) => [...a, ...new Array(k + 1 - a.length).fill(na)]
      const row2 = [...res.se.slice(1).reverse(), constant ? res.se[0] : na]
      return SimpleRangeValue.onlyValues([row1, row2, pad([res.r2, res.seY]), pad([res.F, res.dfRes]), pad([res.ssReg, res.ssRes])] as never)
    })
  }
  linest(ast: Ast, state: State) {
    return this.regress(ast, state, 'LINEST', false)
  }
  logest(ast: Ast, state: State) {
    return this.regress(ast, state, 'LOGEST', true)
  }
  linestSize(ast: Ast, state: State) {
    const s = literal(ast.args[3])
    return new ArraySize(xCount(this, ast, state) + 1, s === true || (typeof s === 'number' && s !== 0) ? 5 : 1)
  }

  private project(ast: Ast, state: State, name: string, log: boolean) {
    return this.runFunction(ast.args, state, this.metadata(name), (yR: SimpleRangeValue, xR: SimpleRangeValue | undefined, newR: SimpleRangeValue | undefined, constant: boolean) => {
      const d = regressionData(yR, xR)
      if (d instanceof CellError) return d
      if (log && d.y.some((v) => v <= 0)) return err(ErrorType.NUM, 'GROWTH needs known_y’s greater than 0.')
      const res = ols(log ? d.y.map(Math.log) : d.y, d.xs, constant)
      if (!res) return err(ErrorType.NUM, 'The regression could not be calculated.')
      const k = d.xs.length
      const predict = (xv: number[]) => {
        const v = res.coef[0] + xv.reduce((a, x, j) => a + res.coef[j + 1] * x, 0)
        return log ? Math.exp(v) : v
      }
      const src = newR ?? xR
      if (!src) return SimpleRangeValue.onlyValues((d.byCol ? d.y.map((_, i) => [predict([i + 1])]) : [d.y.map((_, i) => predict([i + 1]))]) as never)
      const nd = grid(src)
      if (nd.flat().some((v) => typeof v !== 'number')) return err(ErrorType.VALUE, 'new_x’s must contain only numbers.')
      if (k === 1) return SimpleRangeValue.onlyValues(nd.map((r) => r.map((v) => predict([v as number]))) as never)
      // several variables: one observation per row (or per column for row-oriented data)
      if (d.byCol) return SimpleRangeValue.onlyValues(nd.map((r) => [predict(r as number[])]) as never)
      return SimpleRangeValue.onlyValues([nd[0].map((_, j) => predict(nd.map((r) => r[j] as number)))] as never)
    })
  }
  trend(ast: Ast, state: State) {
    return this.project(ast, state, 'TREND', false)
  }
  growth(ast: Ast, state: State) {
    return this.project(ast, state, 'GROWTH', true)
  }
  trendSize(ast: Ast, state: State) {
    const has = (i: number) => ast.args[i] && ast.args[i].type !== 'EMPTY'
    const src = has(2) ? 2 : has(1) ? 1 : 0
    const s = this.arraySizeForAst(ast.args[src], state)
    const k = has(1) ? xCount(this, ast, state) : 1
    if (k === 1 || src === 0) return new ArraySize(s.width, s.height)
    const y = this.arraySizeForAst(ast.args[0], state)
    return y.width === 1 ? new ArraySize(1, s.height) : new ArraySize(s.width, 1)
  }

  forecast(ast: Ast, state: State) {
    const name = ast.procedureName === 'FORECAST.LINEAR' ? 'FORECAST.LINEAR' : 'FORECAST'
    return this.runFunction(ast.args, state, this.metadata(name), (x: number, yR: SimpleRangeValue, xR: SimpleRangeValue) => {
      const p = pairs(yR, xR)
      if (p instanceof CellError) return p
      const res = ols(p.y, [p.x])
      if (!res || res.dropped[1]) return err(ErrorType.DIV_BY_ZERO, 'The known_x’s have no spread.')
      return res.coef[0] + res.coef[1] * x
    })
  }

  intercept(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('INTERCEPT'), (yR: SimpleRangeValue, xR: SimpleRangeValue) => {
      const p = pairs(yR, xR)
      if (p instanceof CellError) return p
      const res = ols(p.y, [p.x])
      if (!res || res.dropped[1]) return err(ErrorType.DIV_BY_ZERO, 'The known_x’s have no spread.')
      return res.coef[0]
    })
  }

  // ------------------------------------------------------------ descriptive
  kurt(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('KURT'), (...ranges: SimpleRangeValue[]) => {
      const xs = ranges.flatMap(numbers)
      if (xs.length < 4) return err(ErrorType.DIV_BY_ZERO, 'KURT needs at least four numbers.')
      const k = kurtosis(xs)
      return isFinite(k) ? k : err(ErrorType.DIV_BY_ZERO, 'The values have no spread.')
    })
  }

  rankavg(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('RANK.AVG'), (x: number, ref: SimpleRangeValue, order: number) => {
      const xs = numbers(ref)
      const ties = xs.filter((v) => v === x).length
      if (!ties) return err(ErrorType.NA, 'The number is not in the list.')
      const better = xs.filter((v) => (order ? v < x : v > x)).length
      return better + (ties + 1) / 2
    })
  }

  private prank(ast: Ast, state: State, name: string, exclusive: boolean) {
    return this.runFunction(ast.args, state, this.metadata(name), (arr: SimpleRangeValue, x: number, sig: number) => {
      if (sig < 1) return err(ErrorType.NUM, 'significance must be at least 1.')
      const r = percentRank(numbers(arr), x, sig, exclusive)
      return r === null ? err(ErrorType.NA, 'x is outside the range of the data.') : r
    })
  }
  percentrankinc(ast: Ast, state: State) {
    return this.prank(ast, state, ast.procedureName === 'PERCENTRANK' ? 'PERCENTRANK' : 'PERCENTRANK.INC', false)
  }
  percentrankexc(ast: Ast, state: State) {
    return this.prank(ast, state, 'PERCENTRANK.EXC', true)
  }

  trimmean(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('TRIMMEAN'), (arr: SimpleRangeValue, pct: number) => {
      if (pct < 0 || pct >= 1) return err(ErrorType.NUM, 'percent must be at least 0 and less than 1.')
      const s = numbers(arr).sort((a, b) => a - b)
      if (!s.length) return err(ErrorType.NUM)
      const cut = Math.floor((s.length * pct) / 2)
      const kept = s.slice(cut, s.length - cut)
      return kept.reduce((a, b) => a + b, 0) / kept.length
    })
  }

  // ------------------------------------------------------------ matrices
  private square(arr: SimpleRangeValue): number[][] | CellError {
    if (arr.width() !== arr.height()) return err(ErrorType.VALUE, 'The array must have as many rows as columns.')
    const d = grid(arr)
    if (d.flat().some((v) => typeof v !== 'number')) return err(ErrorType.VALUE, 'The array must contain only numbers.')
    return d as number[][]
  }
  minverse(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('MINVERSE'), (arr: SimpleRangeValue) => {
      const m = this.square(arr)
      if (m instanceof CellError) return m
      const inv = matInverse(m)
      return inv ? SimpleRangeValue.onlyValues(inv as never) : err(ErrorType.NUM, 'The matrix is singular (it has no inverse).')
    })
  }
  mdeterm(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('MDETERM'), (arr: SimpleRangeValue) => {
      const m = this.square(arr)
      return m instanceof CellError ? m : determinant(m)
    })
  }
  sameAsFirst(ast: Ast, state: State) {
    const s = this.arraySizeForAst(ast.args[0], state)
    return new ArraySize(s.width, s.height)
  }

  // ------------------------------------------------------------ misc
  randarray(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('RANDARRAY'), (rows: number, cols: number, min: number, max: number, whole: boolean) => {
      if (rows < 1 || cols < 1) return err(ErrorType.VALUE, 'rows and columns must be at least 1.')
      if (min > max) return err(ErrorType.VALUE, 'min must not be greater than max.')
      const one = () => (whole ? Math.floor(Math.ceil(min) + Math.random() * (Math.floor(max) - Math.ceil(min) + 1)) : min + Math.random() * (max - min))
      return SimpleRangeValue.onlyValues(Array.from({ length: rows }, () => Array.from({ length: cols }, one)) as never)
    })
  }
  randarraySize(ast: Ast) {
    const r = literal(ast.args[0])
    const c = literal(ast.args[1])
    return new ArraySize(typeof c === 'number' && c >= 1 ? Math.floor(c) : 1, typeof r === 'number' && r >= 1 ? Math.floor(r) : 1)
  }

  permut(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('PERMUT'), (n: number, k: number) => {
      n = Math.trunc(n)
      k = Math.trunc(k)
      if (n < 0 || k < 0 || k > n) return err(ErrorType.NUM, 'Needs 0 ≤ number_chosen ≤ number.')
      let p = 1
      for (let i = 0; i < k; i++) p *= n - i
      return p
    })
  }

  prob(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('PROB'), (xR: SimpleRangeValue, pR: SimpleRangeValue, lo: number, hi?: number) => {
      const xs = grid(xR).flat()
      const ps = grid(pR).flat()
      if (xs.length !== ps.length) return err(ErrorType.NA, 'x_range and prob_range must be the same size.')
      if (ps.some((p) => typeof p !== 'number' || p < 0 || p > 1)) return err(ErrorType.NUM, 'Probabilities must be between 0 and 1.')
      const total = (ps as number[]).reduce((a, b) => a + b, 0)
      if (Math.abs(total - 1) > 1e-9) return err(ErrorType.NUM, 'The probabilities must add up to 1.')
      const top = hi ?? lo
      return xs.reduce<number>((a, x, i) => a + (typeof x === 'number' && x >= lo && x <= top ? (ps[i] as number) : 0), 0)
    })
  }
}

function pairs(yR: SimpleRangeValue, xR: SimpleRangeValue): { y: number[]; x: number[] } | CellError {
  const y = grid(yR).flat()
  const x = grid(xR).flat()
  if (y.length !== x.length) return err(ErrorType.NA, 'known_y’s and known_x’s must be the same size.')
  const py: number[] = []
  const px: number[] = []
  y.forEach((v, i) => {
    if (typeof v === 'number' && typeof x[i] === 'number') {
      py.push(v)
      px.push(x[i] as number)
    }
  })
  if (py.length < 2) return err(ErrorType.DIV_BY_ZERO, 'Needs at least two pairs of numbers.')
  return { y: py, x: px }
}

let registered = false
export function registerExcelProFunctions3() {
  if (registered) return
  registered = true
  const names = Object.keys(ExcelProFunctions3.implementedFunctions)
  HyperFormula.registerFunctionPlugin(ExcelProFunctions3, { enGB: Object.fromEntries(names.map((n) => [n, n])) })
}
