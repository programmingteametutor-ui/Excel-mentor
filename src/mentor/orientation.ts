/* Foundational Excel Orientation for students who have never used Excel: tiny steps checked automatically. */
import { emptyMeta, newId, type Workbook } from '../engine/workbook'
import type { WorkbookDoc } from '../model/types'
import { MAX_COLS, MAX_ROWS } from '../model/types'
import type { Selection } from '../grid/selection'
import type { ConceptId } from './types'
import { HEADER_STYLE } from './build'

export interface OrientationContext {
  wb: Workbook
  sheetName: string
  selection: Selection
}

export interface OrientationStep {
  id: string
  title: string
  concept: ConceptId
  explain: string
  task: string
  /** typed answer instead of a workbook action */
  question?: { answers: string[]; placeholder: string }
  check?: (c: OrientationContext) => boolean
}

const val = (c: OrientationContext, sheet: string, r: number, col: number) => {
  const id = c.wb.sheetByName(sheet)?.id
  return id ? c.wb.getValue(id, r, col) : null
}
const text = (c: OrientationContext, sheet: string, r: number, col: number) => String(val(c, sheet, r, col) ?? '').trim().toLowerCase()
const formula = (c: OrientationContext, sheet: string, r: number, col: number) => {
  const id = c.wb.sheetByName(sheet)?.id
  return id ? (c.wb.isFormula(id, r, col) ? c.wb.getEditText(id, r, col).toUpperCase().replace(/\s/g, '') : '') : ''
}
const at = (c: OrientationContext, r: number, col: number) => c.selection.anchor.r === r && c.selection.anchor.c === col && c.selection.focus.r === r && c.selection.focus.c === col && !c.selection.extra.length

export const ORIENTATION: OrientationStep[] = [
  { id: 'what', title: 'What Excel is', concept: 'workbook', explain: 'Excel is a grid for working with numbers and lists. You type data into boxes, and Excel can calculate with it for you — like a very smart calculator with memory.', task: 'Look around the screen: the grid in the middle is where your data goes. Click Continue.' },
  { id: 'workbook', title: 'Workbook vs worksheet', concept: 'workbook', explain: 'This whole file is a workbook. The tabs at the bottom are worksheets — like pages in a notebook.', task: 'Click the “Practice” tab at the bottom of the screen.', check: (c) => c.sheetName === 'Practice' },
  { id: 'rows-cols', title: 'Rows and columns', concept: 'cell-ref', explain: 'Columns go up and down and have letters (A, B, C…). Rows go across and have numbers (1, 2, 3…).', task: 'Click the letter C at the top of the grid to select the whole column C.', check: (c) => c.selection.anchor.c === 2 && c.selection.focus.c === 2 && c.selection.anchor.r === 0 && c.selection.focus.r >= MAX_ROWS - 1 },
  { id: 'cell', title: 'Cells', concept: 'cell-ref', explain: 'Where a column and a row meet, there is a box called a cell. Each cell holds one value.', task: 'Click the cell in column B, row 3.', check: (c) => c.sheetName === 'Practice' && at(c, 2, 1) },
  { id: 'address', title: 'Cell addresses', concept: 'cell-ref', explain: 'Every cell has an address: column letter, then row number. The cell you just clicked is B3 — look at the Name Box at the top left.', task: 'What is the address of the cell in column D, row 5? Type it below.', question: { answers: ['d5'], placeholder: 'e.g. A1' } },
  { id: 'select', title: 'Selecting cells', concept: 'range', explain: 'You can select a block of cells by dragging from one corner to the other. A block like this is called a range, written A1:C3.', task: 'Drag from A1 to C3 to select that block (or click A1, then Shift+click C3).', check: (c) => c.sheetName === 'Practice' && Math.min(c.selection.anchor.r, c.selection.focus.r) === 0 && Math.min(c.selection.anchor.c, c.selection.focus.c) === 0 && Math.max(c.selection.anchor.r, c.selection.focus.r) === 2 && Math.max(c.selection.anchor.c, c.selection.focus.c) === 2 },
  { id: 'enter', title: 'Entering data', concept: 'data-entry', explain: 'To put something in a cell, click it, type, and press Enter.', task: 'Click A1 on the Practice sheet, type Hello and press Enter.', check: (c) => text(c, 'Practice', 0, 0) === 'hello' },
  { id: 'edit', title: 'Editing data', concept: 'data-entry', explain: 'To change a cell, click it and type again — or press F2 (or double-click) to edit what is there.', task: 'Change A1 so it says Hello Excel.', check: (c) => text(c, 'Practice', 0, 0) === 'hello excel' },
  { id: 'copy', title: 'Copying and pasting', concept: 'data-entry', explain: 'Ctrl+C copies the selected cell; Ctrl+V pastes it where you click next.', task: 'Copy A1 and paste it into A2.', check: (c) => text(c, 'Practice', 1, 0) === 'hello excel' },
  { id: 'format', title: 'Basic formatting', concept: 'formatting', explain: 'Formatting changes how a cell looks. Bold (Ctrl+B) makes text stand out.', task: 'Make cell A1 bold.', check: (c) => !!c.wb.getStyle(c.wb.sheetByName('Practice')!.id, 0, 0)?.bold },
  { id: 'arith', title: 'Simple arithmetic', concept: 'formula', explain: 'Excel can calculate. Start with = and then the sum, like =5+3. Press Enter and Excel shows the answer.', task: 'In B1, type =5+3 and press Enter.', check: (c) => formula(c, 'Practice', 0, 1) !== '' && val(c, 'Practice', 0, 1) === 8 },
  { id: 'formula', title: 'What a formula is', concept: 'formula', explain: 'Anything starting with = is a formula. The cell shows the result; the formula bar at the top shows the formula itself. * means multiply.', task: 'In B2, type =B1*2 and press Enter. It doubles the value in B1.', check: (c) => /B1/.test(formula(c, 'Practice', 1, 1)) && val(c, 'Practice', 1, 1) === 16 },
  { id: 'refs', title: 'Formulas use cell references', concept: 'cell-ref', explain: 'Using addresses (like C1) instead of numbers means the answer updates when the data changes.', task: 'Type 10 in C1 and 20 in C2. Then in C3 write a formula that adds C1 and C2.', check: (c) => /C1/.test(formula(c, 'Practice', 2, 2)) && /C2/.test(formula(c, 'Practice', 2, 2)) && val(c, 'Practice', 2, 2) === 30 },
  { id: 'table', title: 'Reading a table', concept: 'range', explain: 'A table has headers in the first row and one record per row. Read across a row to see one item; read down a column to compare items.', task: 'Open the “Welcome” sheet. Which product sold the most units? Type its name below.', question: { answers: ['notebooks', 'notebook'], placeholder: 'Product name' } },
  { id: 'save', title: 'Saving and navigating', concept: 'workbook', explain: 'ExcelPro saves your work automatically (see “Saved” at the top). Ctrl+Home jumps back to cell A1; Ctrl+arrow keys jump to the edge of your data.', task: 'Click any cell, then press Ctrl+Home to jump to A1.', check: (c) => at(c, 0, 0) },
]

export const ORIENTATION_CONCEPTS: ConceptId[] = ['workbook', 'cell-ref', 'range', 'data-entry', 'formatting', 'formula']

export function orientationDoc(): WorkbookDoc {
  const welcome = { ...emptyMeta(newId(), 'Welcome'), cells: [['Product', 'Units sold'], ['Pens', 120], ['Notebooks', 340], ['Bags', 90], ['Pencils', 210]] as (string | number)[][], styles: { '0,0': HEADER_STYLE, '0,1': HEADER_STYLE }, colWidths: { 0: 130, 1: 100 } }
  const practice = { ...emptyMeta(newId(), 'Practice'), cells: [] }
  const now = new Date().toISOString()
  void MAX_COLS
  return { schemaVersion: 1, id: newId('wb'), name: 'Excel Orientation', activeSheetId: welcome.id, sheetOrder: [welcome.id, practice.id], sheets: [welcome, practice], namedRanges: [], createdAt: now, updatedAt: now }
}
