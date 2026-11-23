/* ExcelMentor data model. No React or DOM code here. */

export type ExcelLevel = 'never' | 'beginner' | 'intermediate' | 'advanced'
export type Role = 'Data Analyst' | 'HR' | 'Finance' | 'Sales' | 'Marketing' | 'Operations' | 'Student' | 'Other'
export type Domain = 'Sales' | 'HR' | 'Finance' | 'Marketing' | 'Operations' | 'Healthcare' | 'Education' | 'Customer Support' | 'General'
export type Goal = 'Job preparation' | 'Work skills' | 'Interview preparation' | 'Academic' | 'General learning'
export type PreferredDifficulty = 'easy' | 'normal' | 'challenging'
export type ProblemLevel = 'beginner' | 'intermediate' | 'advanced'

export interface Attachment {
  name: string
  /** text extracted from the file (truncated) */
  text: string
}

export interface Profile {
  level: ExcelLevel
  role: Role
  domain: Domain
  description: string
  attachments: Attachment[]
  goal?: Goal
  difficulty: PreferredDifficulty
  duration: 15 | 30 | 45 | 60
  createdAt: string
  orientationDone?: boolean
}

export type ConceptId =
  | 'workbook'
  | 'cell-ref'
  | 'range'
  | 'data-entry'
  | 'formatting'
  | 'formula'
  | 'sum'
  | 'average'
  | 'count'
  | 'max'
  | 'min'
  | 'round'
  | 'percent'
  | 'abs-ref'
  | 'if'
  | 'countif'
  | 'sumif'
  | 'averageif'
  | 'concat'
  | 'xlookup'
  | 'vlookup'
  | 'index-match'
  | 'sort'
  | 'filter'
  | 'chart'
  | 'cond-format'
  | 'table'
  | 'pivot'

export type MasteryStage = 'locked' | 'introduced' | 'learning' | 'practicing' | 'applying' | 'strong' | 'mastered'

export interface ConceptProgress {
  stage: MasteryStage
  /** think-question answers: correct counts 1, partial 0.5 */
  understandingScore: number
  understandingCount: number
  practiceCorrect: number
  practiceAttempts: number
  /** task cells that used this concept, judged at the student's last check of each problem */
  appliedCorrect: number
  appliedTotal: number
  /** sum of the highest hint level used per applied cell */
  hintLevels: number
  /** problems in which every cell of this concept was correct with hint level <= 2 */
  independentProblems: number
  problems: number
  lastPracticedAt?: string
}

export type ColumnType = 'text' | 'integer' | 'number' | 'currency' | 'percent' | 'date'

export interface DatasetColumn {
  name: string
  type: ColumnType
}

/** The problem as Gemini (or the built-in bank) describes it, before layout. Formula templates use [Column] tokens. */
export interface ProblemDraft {
  title: string
  context: string
  objective: string
  explanation: { situation: string; goal: string; data: string; think: string }
  concepts: ConceptId[]
  estimatedMinutes: number
  currencySymbol?: string
  dataset: { columns: DatasetColumn[]; rows: (string | number)[][] }
  lookup?: { name: string; columns: DatasetColumn[]; rows: (string | number)[][] }
  rowTasks: { header: string; requirement: string; concept: ConceptId; formula: string; format?: ColumnType }[]
  summaryTasks: { label: string; requirement: string; concept: ConceptId; formula: string; format?: ColumnType; underColumn?: string }[]
  operations: (
    | { type: 'sort'; column: string; order: 'asc' | 'desc'; requirement: string }
    | { type: 'chart'; chartType: 'column' | 'bar' | 'line' | 'pie'; labelColumn: string; valueColumn: string; requirement: string }
    | { type: 'format'; column: string; format: 'currency' | 'percent'; requirement: string }
  )[]
  instructions: string[]
}

export type TargetKind = 'cell' | 'sort' | 'chart' | 'format'

export interface CellTarget {
  kind: 'cell'
  id: string
  sheet: string
  cell: string // A1
  r: number
  c: number
  requirement: string
  concept: ConceptId
  /** expected formula in A1 form (hidden from the workbook) */
  formula: string
  /** functions the method must use, e.g. ['SUM'] */
  requireFunctions: string[]
  requireFormula: boolean
  tolerance: number
  format?: ColumnType
}

export interface SortTarget {
  kind: 'sort'
  id: string
  sheet: string
  requirement: string
  concept: 'sort'
  /** data block (header row included) */
  range: string
  column: string // column letter
  columnName: string
  order: 'asc' | 'desc'
}

export interface ChartTarget {
  kind: 'chart'
  id: string
  sheet: string
  requirement: string
  concept: 'chart'
  chartType: string
  labelCol: string
  valueCol: string
  valueName: string
}

export interface FormatTarget {
  kind: 'format'
  id: string
  sheet: string
  requirement: string
  concept: 'formatting' | 'percent'
  range: string
  format: 'currency' | 'percent'
}

export type Target = CellTarget | SortTarget | ChartTarget | FormatTarget

/** How the session was started: a problem built for chosen topics, or one built from an uploaded take-home assessment. */
export type SessionMode = 'topic' | 'takehome'

export interface ProblemPublic {
  problemId: string
  mode: SessionMode
  datasetVersion: string
  title: string
  context: string
  objective: string
  explanation: ProblemDraft['explanation']
  concepts: ConceptId[]
  /** concepts taught in this problem (others are revision) */
  newConcepts: ConceptId[]
  instructions: string[]
  level: ProblemLevel
  domain: Domain
  estimatedMinutes: number
  source: 'gemini' | 'assessment'
  /** take-home only: the mentor's summary of the uploaded brief */
  brief?: string
  /** take-home only: requirements the app can't check automatically (taught in conversation) */
  uncheckedRequirements?: string[]
  /** public description of each target: no expected answers */
  targets: { id: string; kind: TargetKind; sheet: string; cell?: string; requirement: string; concept: ConceptId }[]
}

export type CellState = 'not_required' | 'not_attempted' | 'incorrect' | 'partially_correct' | 'correct_value' | 'correct_formula' | 'correct'

export interface TargetResult {
  id: string
  state: CellState
  student: string
  /** short, student-safe description of what is expected (no answer) */
  expected: string
  message: string
}

export interface CheckRecord {
  at: string
  results: TargetResult[]
  correct: number
  total: number
}

export interface TargetStats {
  hintLevel: number
  hintsUsed: number
  attempts: number
  lastContent?: string
  firstAttemptAt?: string
  solvedAfterHint?: boolean
}

/** The nine stages every session follows (both modes). Learn and Practice repeat once per concept. */
export type StageId = 'problem' | 'roadmap' | 'learn' | 'practice' | 'plan' | 'solve' | 'check' | 'followup' | 'report'

export interface StagePosition {
  stage: StageId
  /** index into problem.newConcepts while in learn/practice */
  concept: number
}

export type TurnVerdict = 'none' | 'shallow' | 'parroting' | 'incorrect' | 'partial' | 'pass'
export type PressureLogic = 'A' | 'B' | 'C' | 'D'

/** One mentor reply as Gemini returns it. The app validates it before the student sees it. */
export interface MentorTurn {
  message_to_student: string
  current_stage: StageId
  concept_index: number
  gate_passed: boolean
  verdict: TurnVerdict
  hint_level: number
  pressure_test: PressureLogic | 'none'
  context_added: string[]
  assumptions_rejected: string[]
  stage_summary: string
  section_score: number
  teach_back_passed: boolean
  session_complete: boolean
  mastery_profile: string
}

export interface ChatMessage {
  id: string
  from: 'mentor' | 'student'
  text: string
  /** event = a note from the app (engine results, hint requests); error = Gemini failed */
  kind?: 'mentor' | 'event' | 'error'
  at: string
}

/** A passed gate: the evidence the report is built from. */
export interface StageRecord {
  stage: StageId
  concept?: ConceptId
  passedAt: string
  summary: string
  score: number
  /** student replies needed to pass */
  attempts: number
}

export interface ConceptLesson {
  concept: ConceptId
  realLife: string
  excel: string
  formulaExplanation: string
  thinkQuestion: string
  thinkAnswer: string
  example: { title: string; headers: string[]; rows: (string | number)[][]; resultLabel: string; formula: string; note: string }
  practice: { prompt: string; headers: string[]; rows: (string | number)[][]; resultLabel: string; formula: string }
  source: 'gemini' | 'built-in'
}

export interface PracticeTarget {
  concept: ConceptId
  cell: string
  r: number
  c: number
  formula: string
  requireFunctions: string[]
  attempts: number
  passed: boolean
}

export interface MentorSession {
  /** 2 = Gemini-driven 9-stage flow; older sessions have no version */
  version?: 2
  wbId: string
  kind: 'problem' | 'orientation'
  mode?: SessionMode
  problem?: ProblemPublic
  position: StagePosition
  /** student replies at the current gate */
  attempts: number
  orientationStep?: number
  messages: ChatMessage[]
  lessons: Partial<Record<ConceptId, ConceptLesson>>
  exampleRow: number
  practiceRow: number
  practice: Partial<Record<ConceptId, PracticeTarget>>
  /** worked example per concept on the Learning Example sheet */
  exampleCells?: Partial<Record<ConceptId, { cell: string; formula: string; value: string; range: string }>>
  /** Context Validation protocol: what the mentor accepted into the solution plan, and what it rejected */
  ledger: { added: string[]; rejected: string[] }
  /** Pressure Test logics already used, per stage/concept key (anti-repetition) */
  pressureTests: Record<string, PressureLogic[]>
  stageLog: StageRecord[]
  /** highest hint level used per task group (solve/check) */
  hints: Record<string, number>
  /** task group waiting for a teach-back after a level-4 reveal */
  teachBack?: string
  /** an app event that got no reply (Gemini failed); Retry resends it */
  pending?: { event: string }
  stats: Record<string, TargetStats>
  checks: CheckRecord[]
  startedAt: string
  lastActivityAt: string
  timeSpentMs: number
  completed?: { at: string; assisted: boolean; report: PerformanceReport }
  /** original workbook for Reset My Work */
  originalDocKey: string
  /** one-off message shown at the top of the panel */
  notice?: string
}

export interface PerformanceReport {
  title: string
  mode: SessionMode
  completion: number
  cells: { correct: number; total: number }
  conceptsLearned: ConceptId[]
  strong: string[]
  needsPractice: string[]
  hints: number
  independence: 'High' | 'Moderate' | 'Low'
  checks: number
  timeMinutes: number
  recommendation: string
  /** performance in each of the nine stages, from the gates passed */
  sections: { area: string; score: number | null; note: string }[]
  /** the mentor's 10–15 sentence mastery profile (final-stage summary) */
  masteryProfile?: string
  nextSteps?: { revise: string[]; problem: string; difficulty: string }
}

export interface ProblemHistoryEntry {
  problemId: string
  wbId: string
  title: string
  level: ProblemLevel
  concepts: ConceptId[]
  mode?: SessionMode
  startedAt: string
  completedAt?: string
  assisted?: boolean
  firstCheckScore?: number
  finalScore?: number
  hints: number
  attempts: number
}

export interface LearnerState {
  concepts: Partial<Record<ConceptId, ConceptProgress>>
  history: ProblemHistoryEntry[]
  currentWbId?: string
  nextLevel?: ProblemLevel
}

export interface MentorSettings {
  connection: 'none' | 'key' | 'proxy'
  apiKey: string
  proxyUrl: string
  model: string
}
