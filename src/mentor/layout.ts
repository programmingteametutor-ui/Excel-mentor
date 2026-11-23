/* Translates [Column] formula templates into A1 formulas for a known table layout. */
import { colToLetters, quoteSheetName } from '../model/address'

export interface TableLayout {
  sheet: string
  /** 0-based header row */
  headerRow: number
  firstRow: number
  lastRow: number
  /** lower-case column name -> 0-based column index */
  cols: Map<string, number>
}

export interface LayoutContext {
  main: TableLayout
  /** other tables by lower-case name (e.g. "lookup") */
  others?: Map<string, TableLayout>
  /** lower-case summary label -> A1 address on the main sheet */
  summaries?: Map<string, string>
  /** sheet the formula will live on */
  fromSheet: string
}

const a1 = (r: number, c: number, abs = false) => (abs ? `$${colToLetters(c)}$${r + 1}` : `${colToLetters(c)}${r + 1}`)

function resolveTable(ctx: LayoutContext, name: string): { t: TableLayout; col: string } | null {
  const dot = name.indexOf('.')
  if (dot > 0) {
    const t = ctx.others?.get(name.slice(0, dot).trim().toLowerCase())
    if (t) return { t, col: name.slice(dot + 1).trim() }
  }
  return { t: ctx.main, col: name.trim() }
}

function prefix(ctx: LayoutContext, t: TableLayout) {
  return t.sheet === ctx.fromSheet ? '' : `${quoteSheetName(t.sheet)}!`
}

/**
 * Tokens:
 *  [Col]            the cell of column Col in the current row
 *  [Col:all]        all data cells of Col (e.g. B2:B11)
 *  [A:all]:[B:all]  the block from column A to column B (for VLOOKUP tables)
 *  [Lookup.Col:all] a column of another table (another sheet)
 *  [[Label]]        the summary cell with that label
 */
export function translateFormula(template: string, ctx: LayoutContext, row?: number): { formula: string } | { error: string } {
  let err: string | null = null
  const colOf = (t: TableLayout, name: string) => {
    const c = t.cols.get(name.toLowerCase())
    if (c === undefined) err = `Unknown column “${name}”${t === ctx.main ? '' : ` in ${t.sheet}`}`
    return c ?? 0
  }
  // protect string literals
  const parts = template.split(/("(?:[^"]|"")*")/)
  const out = parts.map((p, i) => {
    if (i % 2 === 1) return p
    let s = p.replace(/\[([^[\]]+?):all\]\s*:\s*\[([^[\]]+?):all\]/g, (_m, a: string, b: string) => {
      const ra = resolveTable(ctx, a)!
      const rb = resolveTable(ctx, b)!
      const ca = colOf(ra.t, ra.col)
      const cb = colOf(rb.t, rb.col)
      return `${prefix(ctx, ra.t)}${a1(ra.t.firstRow, Math.min(ca, cb), true)}:${a1(ra.t.lastRow, Math.max(ca, cb), true)}`
    })
    s = s.replace(/\[\[([^[\]]+)\]\]/g, (_m, label: string) => {
      const addr = ctx.summaries?.get(label.trim().toLowerCase())
      if (!addr) err = `Unknown summary “${label}”`
      return addr ?? 'A1'
    })
    s = s.replace(/\[([^[\]]+?)(?::(all|first|last))?\]/g, (_m, name: string, mode?: string) => {
      const r = resolveTable(ctx, name)!
      const c = colOf(r.t, r.col)
      const pre = prefix(ctx, r.t)
      if (mode === 'all') return `${pre}${a1(r.t.firstRow, c)}:${a1(r.t.lastRow, c)}`
      if (mode === 'first') return `${pre}${a1(r.t.firstRow, c)}`
      if (mode === 'last') return `${pre}${a1(r.t.lastRow, c)}`
      if (row === undefined) {
        err = `[${name}] needs a row; use [${name}:all] in a summary formula`
        return 'A1'
      }
      return `${pre}${a1(row, c)}`
    })
    return s
  })
  if (err) return { error: err }
  let f = out.join('').trim()
  if (!f.startsWith('=')) f = '=' + f
  return { formula: f }
}

/** Formula functions used (upper case), ignoring text in quotes. */
export function functionsIn(formula: string): string[] {
  const noStr = formula.replace(/"(?:[^"]|"")*"/g, '""')
  const out = new Set<string>()
  for (const m of noStr.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)) out.add(m[1].toUpperCase().replace(/^_XLFN\./, ''))
  return [...out]
}

export interface RefCells {
  /** "sheet|r,c" keys; sheet empty for the formula's own sheet */
  cells: Set<string>
  ranges: string[]
}

/** Cells referenced by a formula (ranges expanded up to a limit). */
export function referencedCells(formula: string, limit = 20000): RefCells {
  const noStr = formula.replace(/"(?:[^"]|"")*"/g, '""')
  const cells = new Set<string>()
  const ranges: string[] = []
  const re = /(?:('(?:[^']|'')+'|[A-Za-z0-9_]+)!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?(?![A-Za-z0-9_(])/g
  for (const m of noStr.matchAll(re)) {
    // skip function-like names (e.g. LOG10) followed by "(" handled by lookahead; skip true/false
    const sheet = m[1] ? m[1].replace(/^'|'$/g, '').replace(/''/g, "'").toLowerCase() : ''
    const c1 = lettersToCol(m[2])
    const r1 = +m[3] - 1
    const c2 = m[4] ? lettersToCol(m[4]) : c1
    const r2 = m[5] ? +m[5] - 1 : r1
    ranges.push((m[1] ? m[1] + '!' : '') + `${m[2].toUpperCase()}${m[3]}${m[4] ? `:${m[4].toUpperCase()}${m[5]}` : ''}`)
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
        if (cells.size >= limit) break
        cells.add(`${sheet}|${r},${c}`)
      }
  }
  return { cells, ranges }
}

export function lettersToCol(s: string): number {
  let n = 0
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}
