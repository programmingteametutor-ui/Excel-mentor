/* What a task cell needs, described in words (sent to the mentor with a hint request). */
import type { Workbook } from '../engine/workbook'
import { colToLetters } from '../model/address'
import { referencedCells } from './layout'
import type { CellTarget } from './types'

/** Plain-language description of the cells a formula needs, using the header names. */
export function describeInputs(wb: Workbook, t: CellTarget): string {
  const sid = wb.sheetByName(t.sheet)?.id
  const refs = referencedCells(t.formula)
  const own = t.sheet.toLowerCase()
  const byCol = new Map<string, { rows: Set<number>; sheet: string }>()
  for (const k of refs.cells) {
    const [sheet, rc] = k.split('|')
    const [r, c] = rc.split(',').map(Number)
    const key = `${sheet}|${c}`
    const e = byCol.get(key) ?? { rows: new Set<number>(), sheet }
    e.rows.add(r)
    byCol.set(key, e)
  }
  const header = (sheet: string, c: number) => {
    const s = sheet && sheet !== own ? wb.sheetByName(sheet)?.id : sid
    const name = s ? wb.getDisplay(s, 0, c).text : ''
    return name || `column ${colToLetters(c)}`
  }
  const cols = [...byCol.entries()].map(([k, v]) => ({ c: +k.split('|')[1], sheet: v.sheet, many: v.rows.size > 1 }))
  if (!cols.length) return 'the right cells'
  const sameRow = cols.every((x) => !x.many)
  const onOther = cols.some((x) => x.sheet && x.sheet !== own)
  const names = cols.map((x) => header(x.sheet, x.c))
  if (sameRow && cols.length > 2) return `the ${names[0]} to ${names[names.length - 1]} values in this row`
  if (sameRow) return `the ${names.join(' and ')} value${names.length > 1 ? 's' : ''} in this row`
  if (cols.length === 1) return `the whole ${names[0]} column${onOther ? ' on the other sheet' : ''}`
  return `the ${names.join(' and ')} columns${onOther ? ' (one is on another sheet)' : ''}`
}
