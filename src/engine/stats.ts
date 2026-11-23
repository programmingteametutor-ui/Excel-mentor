/**
 * Numerical routines shared by the statistical functions, the Analysis ToolPak and Solver reports:
 * descriptive statistics, the t / F / normal / chi-square distributions, small dense matrix algebra
 * and ordinary least squares regression (the maths behind LINEST and the Regression tool).
 *
 * Results follow Excel's definitions (sample statistics, Excel's KURT/SKEW, PERCENTILE.INC, ...).
 */

// ---------------------------------------------------------------- descriptive statistics
export const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0)
export const mean = (xs: number[]) => sum(xs) / xs.length
export function devsq(xs: number[]) {
  const m = mean(xs)
  return xs.reduce((a, x) => a + (x - m) ** 2, 0)
}
/** Sample variance (VAR.S). */
export const variance = (xs: number[]) => devsq(xs) / (xs.length - 1)
/** Population variance (VAR.P). */
export const varianceP = (xs: number[]) => devsq(xs) / xs.length
export const stdev = (xs: number[]) => Math.sqrt(variance(xs))
export const sorted = (xs: number[]) => [...xs].sort((a, b) => a - b)

export function median(xs: number[]) {
  const s = sorted(xs)
  const n = s.length
  return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2
}

/** Most frequent value (first one to reach the highest count, like MODE.SNGL); null when no value repeats. */
export function mode(xs: number[]): number | null {
  const counts = new Map<number, number>()
  let best: number | null = null
  let bc = 1
  for (const x of xs) {
    const c = (counts.get(x) ?? 0) + 1
    counts.set(x, c)
    if (c > bc) (bc = c), (best = x)
  }
  return best
}

/** Excel KURT (sample excess kurtosis); needs at least 4 values and non-zero spread. */
export function kurtosis(xs: number[]): number {
  const n = xs.length
  const m = mean(xs)
  const s = stdev(xs)
  const s4 = xs.reduce((a, x) => a + ((x - m) / s) ** 4, 0)
  return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * s4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
}

/** Excel SKEW (sample skewness); needs at least 3 values and non-zero spread. */
export function skewness(xs: number[]): number {
  const n = xs.length
  const m = mean(xs)
  const s = stdev(xs)
  return (n / ((n - 1) * (n - 2))) * xs.reduce((a, x) => a + ((x - m) / s) ** 3, 0)
}

/** PERCENTILE.INC on already-sorted values. */
export function percentileInc(s: number[], p: number) {
  const idx = (s.length - 1) * p
  const lo = Math.floor(idx)
  const hi = Math.ceil(idx)
  return s[lo] + (s[hi] - s[lo]) * (idx - lo)
}

/** Excel PERCENTRANK.INC / .EXC, truncated (not rounded) to `significance` digits like Excel. */
export function percentRank(values: number[], x: number, significance: number, exclusive: boolean): number | null {
  const s = sorted(values)
  const n = s.length
  if (!n || x < s[0] || x > s[n - 1]) return null
  const trunc = (v: number) => Math.floor(v * 10 ** significance + 1e-9) / 10 ** significance
  // position of x among the sorted values (interpolated between neighbours)
  let i = s.findIndex((v) => v >= x)
  let pos: number
  if (s[i] === x) {
    pos = i
  } else {
    i -= 1
    pos = i + (x - s[i]) / (s[i + 1] - s[i])
  }
  return exclusive ? trunc((pos + 1) / (n + 1)) : n === 1 ? 1 : trunc(pos / (n - 1))
}

// ---------------------------------------------------------------- special functions
const LANCZOS = [676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7]

/** ln Γ(x) for x > 0 (Lanczos approximation, ~15 significant digits). */
export function lgamma(x: number): number {
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - lgamma(1 - x)
  x -= 1
  let a = 0.99999999999980993
  const t = x + 7.5
  for (let i = 0; i < 8; i++) a += LANCZOS[i] / (x + i + 1)
  return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a)
}

/** Continued fraction for the incomplete beta function (Numerical Recipes betacf). */
function betacf(a: number, b: number, x: number) {
  const FPMIN = 1e-300
  const qab = a + b
  const qap = a + 1
  const qam = a - 1
  let c = 1
  let d = 1 - (qab * x) / qap
  if (Math.abs(d) < FPMIN) d = FPMIN
  d = 1 / d
  let h = d
  for (let m = 1; m <= 300; m++) {
    const m2 = 2 * m
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2))
    d = 1 + aa * d
    if (Math.abs(d) < FPMIN) d = FPMIN
    c = 1 + aa / c
    if (Math.abs(c) < FPMIN) c = FPMIN
    d = 1 / d
    h *= d * c
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2))
    d = 1 + aa * d
    if (Math.abs(d) < FPMIN) d = FPMIN
    c = 1 + aa / c
    if (Math.abs(c) < FPMIN) c = FPMIN
    d = 1 / d
    const del = d * c
    h *= del
    if (Math.abs(del - 1) < 1e-15) break
  }
  return h
}

/** Regularised incomplete beta I_x(a, b). */
export function ibeta(x: number, a: number, b: number): number {
  if (x <= 0) return 0
  if (x >= 1) return 1
  const bt = Math.exp(lgamma(a + b) - lgamma(a) - lgamma(b) + a * Math.log(x) + b * Math.log(1 - x))
  return x < (a + 1) / (a + b + 2) ? (bt * betacf(a, b, x)) / a : 1 - (bt * betacf(b, a, 1 - x)) / b
}

/** Regularised lower incomplete gamma P(a, x). */
export function igamma(a: number, x: number): number {
  if (x <= 0) return 0
  if (x < a + 1) {
    let ap = a
    let s = 1 / a
    let del = s
    for (let n = 0; n < 500; n++) {
      ap += 1
      del *= x / ap
      s += del
      if (Math.abs(del) < Math.abs(s) * 1e-15) break
    }
    return s * Math.exp(-x + a * Math.log(x) - lgamma(a))
  }
  let b = x + 1 - a
  let c = 1e300
  let d = 1 / b
  let h = d
  for (let i = 1; i < 500; i++) {
    const an = -i * (i - a)
    b += 2
    d = an * d + b
    if (Math.abs(d) < 1e-300) d = 1e-300
    c = b + an / c
    if (Math.abs(c) < 1e-300) c = 1e-300
    d = 1 / d
    const del = d * c
    h *= del
    if (Math.abs(del - 1) < 1e-15) break
  }
  return 1 - Math.exp(-x + a * Math.log(x) - lgamma(a)) * h
}

export function erf(x: number): number {
  // via the incomplete gamma function: erf(x) = P(1/2, x²)
  const v = igamma(0.5, x * x)
  return x < 0 ? -v : v
}

// ---------------------------------------------------------------- distributions
export const normCdf = (z: number) => 0.5 * (1 + erf(z / Math.SQRT2))

/** Inverse standard normal (Acklam's algorithm with one Newton refinement). */
export function normInv(p: number): number {
  if (p <= 0) return -Infinity
  if (p >= 1) return Infinity
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239]
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572]
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783]
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416]
  const pl = 0.02425
  let x: number
  if (p < pl) {
    const q = Math.sqrt(-2 * Math.log(p))
    x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
  } else if (p <= 1 - pl) {
    const q = p - 0.5
    const r = q * q
    x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1)
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p))
    x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
  }
  const e = normCdf(x) - p
  const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2)
  return x - u / (1 + (x * u) / 2)
}

/** Student t cumulative distribution P(T ≤ t). */
export function tCdf(t: number, df: number): number {
  const x = df / (df + t * t)
  const tail = 0.5 * ibeta(x, df / 2, 0.5)
  return t >= 0 ? 1 - tail : tail
}

/** Two-tailed t probability P(|T| ≥ |t|) — Excel T.DIST.2T. */
export const tTwoTail = (t: number, df: number) => ibeta(df / (df + t * t), df / 2, 0.5)

/** Inverse of a monotone increasing CDF on (lo, hi) by bisection + secant polish. */
function invert(cdf: (x: number) => number, p: number, lo: number, hi: number): number {
  // widen the bracket if needed
  while (cdf(hi) < p && hi < 1e12) hi *= 2
  while (cdf(lo) > p && lo > -1e12) lo = lo < 0 ? lo * 2 : lo - 1
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2
    if (cdf(mid) < p) lo = mid
    else hi = mid
    if (hi - lo < 1e-13 * Math.max(1, Math.abs(mid))) break
  }
  return (lo + hi) / 2
}

/** Inverse t (left tail) — Excel T.INV. */
export const tInv = (p: number, df: number) => invert((x) => tCdf(x, df), p, -50, 50)
/** Two-tailed inverse t — Excel T.INV.2T (positive critical value). */
export const tInv2T = (alpha: number, df: number) => tInv(1 - alpha / 2, df)

/** F cumulative distribution P(F ≤ f). */
export const fCdf = (f: number, d1: number, d2: number) => (f <= 0 ? 0 : ibeta((d1 * f) / (d1 * f + d2), d1 / 2, d2 / 2))
/** Right-tail F probability — Excel F.DIST.RT. */
export const fRight = (f: number, d1: number, d2: number) => 1 - fCdf(f, d1, d2)
/** Inverse F (left tail) — Excel F.INV. */
export const fInv = (p: number, d1: number, d2: number) => invert((x) => fCdf(x, d1, d2), p, 0, 100)

/** Chi-square cumulative distribution. */
export const chiCdf = (x: number, k: number) => igamma(k / 2, x / 2)

// ---------------------------------------------------------------- matrices
export type Matrix = number[][]

export const transpose = (a: Matrix): Matrix => a[0].map((_, j) => a.map((r) => r[j]))

export function matMul(a: Matrix, b: Matrix): Matrix {
  const out: Matrix = []
  for (let i = 0; i < a.length; i++) {
    const row = new Array(b[0].length).fill(0)
    for (let k = 0; k < b.length; k++) {
      const v = a[i][k]
      if (v) for (let j = 0; j < b[0].length; j++) row[j] += v * b[k][j]
    }
    out.push(row)
  }
  return out
}

/** Inverse by Gauss–Jordan elimination with partial pivoting; null when (numerically) singular. */
export function matInverse(m: Matrix): Matrix | null {
  const n = m.length
  const a = m.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))])
  const scale = Math.max(1e-300, ...m.flat().map(Math.abs))
  for (let col = 0; col < n; col++) {
    let piv = col
    for (let r = col + 1; r < n; r++) if (Math.abs(a[r][col]) > Math.abs(a[piv][col])) piv = r
    if (Math.abs(a[piv][col]) < 1e-12 * scale) return null
    ;[a[col], a[piv]] = [a[piv], a[col]]
    const p = a[col][col]
    for (let j = 0; j < 2 * n; j++) a[col][j] /= p
    for (let r = 0; r < n; r++) {
      if (r === col) continue
      const f = a[r][col]
      if (f) for (let j = 0; j < 2 * n; j++) a[r][j] -= f * a[col][j]
    }
  }
  return a.map((r) => r.slice(n))
}

/** Determinant by LU elimination with partial pivoting. */
export function determinant(m: Matrix): number {
  const n = m.length
  const a = m.map((r) => [...r])
  let det = 1
  for (let col = 0; col < n; col++) {
    let piv = col
    for (let r = col + 1; r < n; r++) if (Math.abs(a[r][col]) > Math.abs(a[piv][col])) piv = r
    if (a[piv][col] === 0) return 0
    if (piv !== col) {
      ;[a[col], a[piv]] = [a[piv], a[col]]
      det = -det
    }
    det *= a[col][col]
    for (let r = col + 1; r < n; r++) {
      const f = a[r][col] / a[col][col]
      for (let j = col; j < n; j++) a[r][j] -= f * a[col][j]
    }
  }
  return det
}

// ---------------------------------------------------------------- regression
export interface OlsResult {
  /** coefficients: [intercept, b1, b2, ...] (intercept is 0 when `constant` is false) */
  coef: number[]
  /** standard errors in the same order (intercept SE is NaN when `constant` is false) */
  se: number[]
  t: number[]
  p: number[]
  /** columns dropped because they were collinear with earlier ones (their coefficient is 0) */
  dropped: boolean[]
  n: number
  /** regression and residual degrees of freedom */
  dfReg: number
  dfRes: number
  ssReg: number
  ssRes: number
  r2: number
  adjR2: number
  multipleR: number
  seY: number
  F: number
  pF: number
  fitted: number[]
  residuals: number[]
}

/**
 * Ordinary least squares y = b0 + Σ bj·xj. `xs` holds one array per explanatory variable.
 * Matches Excel's LINEST / Regression tool, including the no-constant variants
 * (R² and the F statistic use the uncentred total sum of squares when there is no intercept).
 */
export function ols(y: number[], xs: number[][], constant = true): OlsResult | null {
  const n = y.length
  const k = xs.length
  if (!n || xs.some((x) => x.length !== n)) return null
  // design matrix columns (intercept first)
  const cols: number[][] = constant ? [new Array(n).fill(1), ...xs] : [...xs]
  // drop collinear columns via Gram–Schmidt residual norms
  const kept: number[] = []
  const basis: number[][] = []
  const dropped = new Array(cols.length).fill(false)
  cols.forEach((c, j) => {
    let v = [...c]
    for (const q of basis) {
      const d = v.reduce((a, x, i) => a + x * q[i], 0)
      v = v.map((x, i) => x - d * q[i])
    }
    const norm = Math.sqrt(v.reduce((a, x) => a + x * x, 0))
    const scale = Math.sqrt(c.reduce((a, x) => a + x * x, 0))
    if (norm <= 1e-10 * Math.max(1, scale)) {
      dropped[j] = true
      return
    }
    kept.push(j)
    basis.push(v.map((x) => x / norm))
  })
  const X = Array.from({ length: n }, (_, i) => kept.map((j) => cols[j][i]))
  const Xt = transpose(X.length ? X : [[]])
  const XtX = matMul(Xt, X)
  const inv = kept.length ? matInverse(XtX) : []
  if (kept.length && !inv) return null
  const Xty = Xt.map((r) => r.reduce((a, v, i) => a + v * y[i], 0))
  const bKept = kept.length ? inv!.map((r) => r.reduce((a, v, i) => a + v * Xty[i], 0)) : []
  const fitted = Array.from({ length: n }, (_, i) => kept.reduce((a, _j, q) => a + bKept[q] * X[i][q], 0))
  const residuals = y.map((v, i) => v - fitted[i])
  const ssRes = residuals.reduce((a, r) => a + r * r, 0)
  const ym = mean(y)
  const ssTot = constant ? y.reduce((a, v) => a + (v - ym) ** 2, 0) : y.reduce((a, v) => a + v * v, 0)
  const ssReg = ssTot - ssRes
  const params = kept.length
  const dfRes = n - params
  const dfReg = constant ? params - 1 : params
  const s2 = dfRes > 0 ? ssRes / dfRes : NaN
  const coefAll = new Array(cols.length).fill(0)
  const seAll = new Array(cols.length).fill(0)
  kept.forEach((j, q) => {
    coefAll[j] = bKept[q]
    seAll[j] = Math.sqrt(Math.max(0, s2 * inv![q][q]))
  })
  const coef = constant ? coefAll : [0, ...coefAll]
  const se = constant ? seAll : [NaN, ...seAll]
  const drop = constant ? dropped : [false, ...dropped]
  const t = coef.map((b, i) => (se[i] > 0 ? b / se[i] : NaN))
  const p = t.map((tv) => (isFinite(tv) && dfRes > 0 ? tTwoTail(tv, dfRes) : NaN))
  const r2 = ssTot > 0 ? ssReg / ssTot : 1
  const adjR2 = constant ? 1 - (1 - r2) * ((n - 1) / dfRes) : 1 - (1 - r2) * (n / dfRes)
  const F = dfReg > 0 && dfRes > 0 ? ssReg / dfReg / (ssRes / dfRes) : NaN
  void k
  return {
    coef,
    se,
    t,
    p,
    dropped: drop,
    n,
    dfReg,
    dfRes,
    ssReg,
    ssRes,
    r2,
    adjR2,
    multipleR: Math.sqrt(Math.max(0, r2)),
    seY: Math.sqrt(s2),
    F,
    pF: isFinite(F) ? fRight(F, dfReg, dfRes) : NaN,
    fitted,
    residuals,
  }
}

// ---------------------------------------------------------------- random numbers
/** Small seedable PRNG (mulberry32) so tools with a "Random seed" are reproducible. */
export function seededRandom(seed?: number): () => number {
  if (seed === undefined || !isFinite(seed)) return Math.random
  let a = Math.floor(seed) >>> 0
  return () => {
    a = (a + 0x6d2b79f5) >>> 0
    let t = a
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

export function randNormal(rand: () => number, mu: number, sigma: number) {
  let u = 0
  while (u === 0) u = rand()
  return mu + sigma * Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * rand())
}

export function randPoisson(rand: () => number, lambda: number) {
  if (lambda > 50) return Math.max(0, Math.round(randNormal(rand, lambda, Math.sqrt(lambda))))
  const L = Math.exp(-lambda)
  let k = 0
  let p = 1
  do {
    k++
    p *= rand()
  } while (p > L)
  return k - 1
}

export function randBinomial(rand: () => number, n: number, p: number) {
  let k = 0
  for (let i = 0; i < n; i++) if (rand() < p) k++
  return k
}
