/* Reads take-home assessment files in the browser: spreadsheets (.xlsx/.xlsm/.csv/.tsv), Word (.docx), PDF and text.
 * Spreadsheets keep their workbook so the real data can become the Task/Data sheets; documents become text for Gemini. */
import JSZip from 'jszip'
import { Workbook } from '../engine/workbook'
import { importDelimited } from '../files/csv'
import { importXlsx } from '../files/xlsx'
import { parseXml, type XmlEl } from '../files/xml'
import type { WorkbookDoc } from '../model/types'
import type { ColumnType, DatasetColumn } from './types'

export interface SheetPreview {
  name: string
  rows: number
  cols: number
  /** first rows as display text, for Gemini to find the header row and understand the data */
  preview: string[][]
}

export interface ExtractedFile {
  name: string
  kind: 'spreadsheet' | 'document'
  /** documents: their text (tables as " | " rows); spreadsheets: a short description */
  text: string
  sheets?: SheetPreview[]
  doc?: WorkbookDoc
  /** something went wrong reading part of it */
  warning?: string
}

export const ACCEPTED_EXTENSIONS = ['.xlsx', '.xlsm', '.csv', '.tsv', '.pdf', '.docx', '.txt', '.md']
export const MAX_FILE_BYTES = 25 * 1024 * 1024
const MAX_DOC_CHARS = 60_000
const PREVIEW_ROWS = 15
const PREVIEW_COLS = 20

const ext = (name: string) => name.split('.').pop()?.toLowerCase() ?? ''

export async function extractFile(file: File): Promise<ExtractedFile> {
  const e = ext(file.name)
  if (file.size > MAX_FILE_BYTES) throw new Error(`${file.name} is larger than 25 MB.`)
  if (e === 'xls' || e === 'doc' || e === 'ods') throw new Error(`.${e} files can't be read in the browser. Save ${file.name} as .${e === 'doc' ? 'docx' : 'xlsx'} and upload it again.`)
  if (e === 'xlsx' || e === 'xlsm') return spreadsheet(file.name, (await importXlsx(await file.arrayBuffer(), file.name)).doc)
  if (e === 'csv' || e === 'tsv') return spreadsheet(file.name, importDelimited(await file.text(), file.name).doc)
  if (e === 'docx') return { name: file.name, kind: 'document', text: clip(await docxText(await file.arrayBuffer())) }
  if (e === 'pdf') return { name: file.name, kind: 'document', text: clip(await pdfText(await file.arrayBuffer())) }
  if (e === 'txt' || e === 'md') return { name: file.name, kind: 'document', text: clip(await file.text()) }
  throw new Error(`${file.name}: unsupported file type. Upload ${ACCEPTED_EXTENSIONS.join(', ')}.`)
}

const clip = (s: string) => (s.length > MAX_DOC_CHARS ? `${s.slice(0, MAX_DOC_CHARS)}\n[…document truncated…]` : s)

export function spreadsheet(name: string, doc: WorkbookDoc): ExtractedFile {
  const wb = Workbook.fromDoc(doc)
  const sheets: SheetPreview[] = []
  for (const sid of wb.order) {
    const m = wb.meta(sid)
    if (m.hidden) continue
    const u = wb.usedSize(sid)
    if (!u.rows || !u.cols) continue
    const preview: string[][] = []
    for (let r = 0; r < Math.min(u.rows, PREVIEW_ROWS); r++) {
      const row: string[] = []
      for (let c = 0; c < Math.min(u.cols, PREVIEW_COLS); c++) row.push(wb.getDisplay(sid, r, c).text)
      preview.push(row)
    }
    sheets.push({ name: m.name, rows: u.rows, cols: u.cols, preview })
  }
  return { name, kind: 'spreadsheet', text: `Spreadsheet with ${sheets.length} sheet(s): ${sheets.map((s) => `${s.name} (${s.rows}×${s.cols})`).join(', ')}`, sheets, doc }
}

// ------------------------------------------------------------------ Word
/** Paragraph and table text from word/document.xml (tables as " | " separated rows). */
export async function docxText(data: ArrayBuffer | Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(data)
  const part = zip.file('word/document.xml')
  if (!part) throw new Error('This .docx has no document body.')
  const root = parseXml(await part.async('string'))
  const body = root.children.find((c) => c.name === 'body') ?? root
  const out: string[] = []
  const runText = (el: XmlEl): string => {
    if (el.name === 't') return el.text
    if (el.name === 'tab') return '\t'
    if (el.name === 'br' || el.name === 'cr') return '\n'
    return el.children.map(runText).join('')
  }
  const walk = (el: XmlEl) => {
    if (el.name === 'p') {
      const t = runText(el).trim()
      const numbered = el.children.some((c) => c.name === 'pPr' && c.children.some((x) => x.name === 'numPr'))
      if (t) out.push(numbered ? `• ${t}` : t)
    } else if (el.name === 'tbl') {
      out.push('[table]')
      for (const tr of el.children.filter((c) => c.name === 'tr')) {
        const cells = tr.children.filter((c) => c.name === 'tc').map((tc) => runText(tc).replace(/\s+/g, ' ').trim())
        out.push(cells.join(' | '))
      }
      out.push('[/table]')
    } else el.children.forEach(walk)
  }
  body.children.forEach(walk)
  return out.join('\n')
}

// ------------------------------------------------------------------ PDF
/** Text of every page, using pdf.js on the main thread (loaded only when a PDF is uploaded). */
export async function pdfText(data: ArrayBuffer): Promise<string> {
  const g = globalThis as { pdfjsWorker?: unknown }
  if (!g.pdfjsWorker) g.pdfjsWorker = await import('pdfjs-dist/build/pdf.worker.mjs')
  const pdfjs = await import('pdfjs-dist')
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(data), isEvalSupported: false, useSystemFonts: true }).promise
  const pages: string[] = []
  for (let i = 1; i <= Math.min(pdf.numPages, 60); i++) {
    const page = await pdf.getPage(i)
    const content = await page.getTextContent()
    let line = ''
    let lastY: number | null = null
    const lines: string[] = []
    for (const item of content.items as { str?: string; transform?: number[]; hasEOL?: boolean }[]) {
      if (typeof item.str !== 'string') continue
      const y = item.transform?.[5] ?? null
      if (lastY !== null && y !== null && Math.abs(y - lastY) > 2 && line.trim()) {
        lines.push(line.trim())
        line = ''
      }
      line += (line && !line.endsWith(' ') && item.str && !item.str.startsWith(' ') ? ' ' : '') + item.str
      if (item.hasEOL) {
        lines.push(line.trim())
        line = ''
      }
      lastY = y
    }
    if (line.trim()) lines.push(line.trim())
    pages.push(`--- page ${i} ---\n${lines.filter(Boolean).join('\n')}`)
  }
  if (pdf.numPages > 60) pages.push(`[…${pdf.numPages - 60} more pages not read…]`)
  const text = pages.join('\n')
  if (text.replace(/--- page \d+ ---/g, '').trim().length < 20) throw new Error('This PDF has no readable text (it may be a scanned image). Upload a text-based PDF or a .docx copy.')
  return text
}

// ------------------------------------------------------------------ tables from an uploaded sheet
const DATE_FMT = /(^|[^"\\])(d{1,4}|m{1,5}|y{2,4})/i

function serialToIso(n: number): string {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000)
  return d.toISOString().slice(0, 10)
}

/** The first row (of the first 40) that looks like a table header: 2+ text cells with a non-empty row under it. */
export function detectHeaderRow(doc: WorkbookDoc, sheetName: string): number | null {
  const wb = Workbook.fromDoc(structuredClone(doc))
  const meta = wb.sheetByName(sheetName) ?? wb.sheets.find((s) => s.name.toLowerCase() === sheetName.trim().toLowerCase())
  if (!meta) return null
  const u = wb.usedSize(meta.id)
  const filled = (r: number) => Array.from({ length: Math.min(u.cols, 60) }, (_, c) => wb.getValue(meta.id, r, c)).filter((v) => v !== null && v !== undefined && v !== '')
  for (let r = 0; r < Math.min(u.rows - 1, 40); r++) {
    const cells = filled(r)
    if (cells.length >= 2 && cells.filter((v) => typeof v === 'string').length >= Math.max(2, cells.length - 1) && filled(r + 1).length >= 2) return r
  }
  return null
}

/** The biggest visible sheet of a workbook (where the data usually is). */
export function largestSheet(file: ExtractedFile): string | null {
  return [...(file.sheets ?? [])].sort((a, b) => b.rows * b.cols - a.rows * a.cols)[0]?.name ?? null
}

/** The table that starts at headerRow on the named sheet: header names, inferred column types and rows of values. */
export function tableFromSheet(doc: WorkbookDoc, sheetName: string, headerRow: number, maxRows = 5000): { columns: DatasetColumn[]; rows: (string | number)[][] } | { error: string } {
  const wb = Workbook.fromDoc(structuredClone(doc))
  const meta = wb.sheetByName(sheetName) ?? wb.sheets.find((s) => s.name.toLowerCase() === sheetName.trim().toLowerCase())
  if (!meta) return { error: `There is no sheet called “${sheetName}”.` }
  const sid = meta.id
  const u = wb.usedSize(sid)
  if (headerRow < 0 || headerRow >= u.rows) return { error: `Row ${headerRow + 1} is outside the data on “${sheetName}”.` }
  // header: contiguous non-empty cells from the first non-empty one
  let c0 = 0
  while (c0 < u.cols && !wb.getDisplay(sid, headerRow, c0).text.trim()) c0++
  const headers: string[] = []
  for (let c = c0; c < u.cols; c++) {
    const h = wb.getDisplay(sid, headerRow, c).text.trim()
    if (!h) break
    headers.push(h)
  }
  if (headers.length < 2) return { error: `Row ${headerRow + 1} on “${sheetName}” doesn't look like a header row (fewer than 2 column names).` }
  const raw: { v: unknown; date: boolean }[][] = []
  for (let r = headerRow + 1; r < u.rows && raw.length < maxRows; r++) {
    const row = headers.map((_, j) => {
      const v = wb.getValue(sid, r, c0 + j)
      const fmt = wb.getStyle(sid, r, c0 + j)?.numFmt ?? ''
      return { v, date: typeof v === 'number' && DATE_FMT.test(fmt) && !/^\[?h/i.test(fmt) }
    })
    if (row.every((x) => x.v === null || x.v === undefined || x.v === '')) break // the table ends at the first empty row
    raw.push(row)
  }
  if (!raw.length) return { error: `The table on “${sheetName}” has no data rows under the header.` }
  const columns: DatasetColumn[] = headers.map((name, j) => {
    const vals = raw.map((r) => r[j]).filter((x) => x.v !== null && x.v !== undefined && x.v !== '')
    let type: ColumnType = 'text'
    if (vals.length && vals.every((x) => typeof x.v === 'number')) {
      if (vals.every((x) => x.date)) type = 'date'
      else type = vals.every((x) => Number.isInteger(x.v)) ? 'integer' : 'number'
    }
    return { name, type }
  })
  const rows = raw.map((row) =>
    row.map((x, j) => {
      if (x.v === null || x.v === undefined) return ''
      if (columns[j].type === 'date') return serialToIso(x.v as number)
      if (typeof x.v === 'number') return x.v
      if (typeof x.v === 'boolean') return x.v ? 'TRUE' : 'FALSE'
      if (typeof x.v === 'object') return '' // error values
      return String(x.v)
    }),
  )
  // unique header names (the engine's column tokens need them)
  const seen = new Map<string, number>()
  for (const col of columns) {
    const k = col.name.toLowerCase()
    const n = (seen.get(k) ?? 0) + 1
    seen.set(k, n)
    if (n > 1) col.name = `${col.name} (${n})`
  }
  return { columns, rows }
}
