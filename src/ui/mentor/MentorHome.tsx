/* ExcelMentor home: profile setup, the dashboard with the two modes (learn by topic, take-home assessment),
 * the take-home upload page, the concept library and concept review. */
import { useEffect, useMemo, useState, type ReactNode } from 'react'
import type { WorkbookDoc } from '../../model/types'
import { colToLetters } from '../../model/address'
import { CONCEPTS, conceptName, CURRICULUM, FOUNDATIONS, GENERATABLE, type ConceptCategory } from '../../mentor/concepts'
import { createAssessmentProblem, createTopicProblem, builtinLesson, type CreatedProblem } from '../../mentor/ai'
import { ACCEPTED_EXTENSIONS, extractFile, type ExtractedFile } from '../../mentor/extract'
import { newSession } from '../../mentor/flow'
import { GeminiClient } from '../../mentor/gemini'
import { checkBlockAnswer, renderBlock } from '../../mentor/lesson'
import { orientationDoc } from '../../mentor/orientation'
import { conceptProgress, recommend, scores, STAGE_LABEL, topicRecommendation, type Recommendation } from '../../mentor/progress'
import { STAGE_INFO } from '../../mentor/protocol'
import type { MentorStore } from '../../mentor/store'
import type { Attachment, ConceptId, Domain, ExcelLevel, Goal, MentorSession, PreferredDifficulty, Profile, Role } from '../../mentor/types'
import { Button, Field, inputClass } from '../components/ui'
import { Icon, Logo } from '../icons'
import { cn } from '../lib/utils'
import { MentorSettingsDialog } from './Settings'
import { VideoPlayer } from './VideoPlayer'

const LEVELS: [ExcelLevel, string][] = [['never', 'Never used Excel'], ['beginner', 'Beginner'], ['intermediate', 'Intermediate'], ['advanced', 'Advanced']]
const ROLES: Role[] = ['Data Analyst', 'HR', 'Finance', 'Sales', 'Marketing', 'Operations', 'Student', 'Other']
const DOMAINS: Domain[] = ['Sales', 'HR', 'Finance', 'Marketing', 'Operations', 'Healthcare', 'Education', 'Customer Support', 'General']
const GOALS: Goal[] = ['Job preparation', 'Work skills', 'Interview preparation', 'Academic', 'General learning']
const STAGE_COLOR: Record<string, string> = {
  locked: 'bg-[var(--panel-2)] text-[var(--muted)]',
  introduced: 'bg-sky-100 text-sky-800 dark:bg-sky-900/40 dark:text-sky-200',
  learning: 'bg-sky-100 text-sky-800 dark:bg-sky-900/40 dark:text-sky-200',
  practicing: 'bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-200',
  applying: 'bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-200',
  strong: 'bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-200',
  mastered: 'bg-emerald-600 text-white',
}
const MAX_TOPICS = 3

const sel = cn(inputClass, 'h-9')

type Page = 'dashboard' | 'profile' | 'concept' | 'takehome'

export function useMentorRoute(): { page: Page; concept?: ConceptId } {
  const parse = () => {
    const h = location.hash
    const m = /^#\/mentor\/concept\/([\w-]+)/.exec(h)
    if (m && m[1] in CONCEPTS) return { page: 'concept' as const, concept: m[1] as ConceptId }
    if (h.startsWith('#/mentor/profile')) return { page: 'profile' as const }
    if (h.startsWith('#/mentor/takehome')) return { page: 'takehome' as const }
    return { page: 'dashboard' as const }
  }
  const [r, setR] = useState(parse)
  useEffect(() => {
    const f = () => setR(parse())
    window.addEventListener('hashchange', f)
    return () => window.removeEventListener('hashchange', f)
  }, [])
  return r
}

async function readAttachment(f: File): Promise<Attachment> {
  try {
    const x = await extractFile(f)
    if (x.kind === 'document') return { name: f.name, text: x.text.slice(0, 6000) }
    const s = x.sheets?.[0]
    return { name: f.name, text: s ? `First sheet (${s.name}):\n${s.preview.map((r) => r.join('\t')).join('\n')}` : x.text }
  } catch (e) {
    return { name: f.name, text: `(${f.name} couldn't be read: ${e instanceof Error ? e.message : String(e)})` }
  }
}

function Onboarding({ store, initial, onDone, onCancel }: { store: MentorStore; initial: Profile | null; onDone: (p: Profile) => void; onCancel?: () => void }) {
  const [p, setP] = useState<Profile>(initial ?? { level: 'beginner', role: 'Student', domain: 'Sales', description: '', attachments: [], goal: undefined, difficulty: 'normal', duration: 30, createdAt: new Date().toISOString() })
  const [reading, setReading] = useState(false)
  const up = (x: Partial<Profile>) => setP((c) => ({ ...c, ...x }))
  return (
    <div className="mx-auto w-full max-w-2xl px-4 py-6" data-testid="onboarding">
      <h1 className="text-[22px] font-semibold">{initial ? 'Your learning profile' : 'Welcome to ExcelMentor'}</h1>
      <p className="mb-5 mt-1 text-[13px] text-[var(--muted)]">Tell us about yourself. Your mentor uses your level, role and domain to choose topics and build business problems from your world.</p>
      <Field label="Excel level">
        <div className="grid grid-cols-2 gap-2 sm:grid-cols-4" role="radiogroup" aria-label="Excel level">
          {LEVELS.map(([v, l]) => (
            <button key={v} type="button" role="radio" aria-checked={p.level === v} onClick={() => up({ level: v })} className={cn('rounded-lg border px-2 py-2 text-[13px]', p.level === v ? 'border-[var(--accent)] bg-[var(--active)] font-semibold text-[var(--accent)]' : 'border-[var(--border)] hover:border-[var(--border-strong)]')}>
              {l}
            </button>
          ))}
        </div>
      </Field>
      <div className="grid gap-x-3 sm:grid-cols-2">
        <Field label="Role">
          <select className={sel} value={p.role} onChange={(e) => up({ role: e.target.value as Role })} aria-label="Role">
            {ROLES.map((r) => <option key={r}>{r}</option>)}
          </select>
        </Field>
        <Field label="Domain">
          <select className={sel} value={p.domain} onChange={(e) => up({ domain: e.target.value as Domain })} aria-label="Domain">
            {DOMAINS.map((r) => <option key={r}>{r}</option>)}
          </select>
        </Field>
        <Field label="Learning goal (optional)">
          <select className={sel} value={p.goal ?? ''} onChange={(e) => up({ goal: (e.target.value || undefined) as Goal | undefined })} aria-label="Learning goal">
            <option value="">—</option>
            {GOALS.map((r) => <option key={r}>{r}</option>)}
          </select>
        </Field>
        <Field label="Preferred difficulty">
          <select className={sel} value={p.difficulty} onChange={(e) => up({ difficulty: e.target.value as PreferredDifficulty })} aria-label="Preferred difficulty">
            <option value="easy">Easy</option>
            <option value="normal">Normal</option>
            <option value="challenging">Challenging</option>
          </select>
        </Field>
        <Field label="Session duration">
          <select className={sel} value={p.duration} onChange={(e) => up({ duration: +e.target.value as Profile['duration'] })} aria-label="Session duration">
            {[15, 30, 45, 60].map((n) => <option key={n} value={n}>{n} minutes</option>)}
          </select>
        </Field>
      </div>
      <Field label="What do you want to learn? (optional)" hint="e.g. “I need to build monthly attendance reports for my team.”">
        <textarea className={cn(inputClass, 'h-20 py-1.5')} value={p.description} onChange={(e) => up({ description: e.target.value })} aria-label="What do you want to learn" />
      </Field>
      <Field label="Attachments (optional)" hint="A sample of your data or a job description (.txt, .csv, .xlsx, .pdf, .docx). Used only to make problems realistic.">
        <input
          type="file"
          multiple
          accept={ACCEPTED_EXTENSIONS.join(',')}
          aria-label="Attachments"
          className="text-[12.5px]"
          onChange={async (e) => {
            const files = [...(e.target.files ?? [])].slice(0, 3)
            setReading(true)
            const list = await Promise.all(files.map(readAttachment))
            setReading(false)
            up({ attachments: [...p.attachments, ...list].slice(0, 3) })
            e.target.value = ''
          }}
        />
        {p.attachments.length > 0 && (
          <ul className="mt-1 text-[12px]">
            {p.attachments.map((a, i) => (
              <li key={i} className="flex items-center gap-2">
                <Icon name="file" size={13} /> {a.name}
                <button type="button" className="text-[var(--muted)] hover:text-[var(--danger)]" aria-label={`Remove ${a.name}`} onClick={() => up({ attachments: p.attachments.filter((_, j) => j !== i) })}>✕</button>
              </li>
            ))}
          </ul>
        )}
      </Field>
      <div className="mt-4 flex gap-2">
        <Button variant="default" disabled={reading} onClick={() => (store.saveProfile({ ...p, orientationDone: initial?.orientationDone }), onDone(p))} data-testid="save-profile">
          {initial ? 'Save profile' : 'Continue'}
        </Button>
        {onCancel && <Button variant="outline" onClick={onCancel}>Cancel</Button>}
      </div>
    </div>
  )
}

function Card({ title, children, className, testId }: { title: string; children: ReactNode; className?: string; testId?: string }) {
  return (
    <section className={cn('rounded-xl border border-[var(--border)] bg-[var(--panel)] p-4', className)} data-testid={testId}>
      <h2 className="mb-2 text-[13px] font-semibold uppercase tracking-wide text-[var(--muted)]">{title}</h2>
      {children}
    </section>
  )
}

function stageDescription(s: MentorSession): string {
  if (s.completed) return 'Completed'
  if (s.kind === 'orientation') return 'Excel orientation'
  if (s.version !== 2) return 'Started in an older version'
  const info = STAGE_INFO[s.position.stage]
  const c = s.position.stage === 'learn' || s.position.stage === 'practice' ? s.problem?.newConcepts[s.position.concept] : null
  return `Stage ${info.n} of 9 · ${info.area}${c ? ` (${conceptName(c)})` : ''}`
}

function sessionProgress(s: MentorSession): number {
  if (s.completed) return 100
  if (s.kind === 'orientation') return Math.round((100 * (s.orientationStep ?? 0)) / 15)
  if (s.version !== 2) return 0
  return Math.min(95, Math.round((100 * (STAGE_INFO[s.position.stage].n - 1)) / 9) + 3)
}

export function MentorHome({ store, onOpenDoc, onOpenWorkbook, workbookExists }: { store: MentorStore; onOpenDoc: (doc: WorkbookDoc) => void; onOpenWorkbook: (id: string) => void; workbookExists: (id: string) => Promise<boolean> }) {
  const route = useMentorRoute()
  const [profile, setProfile] = useState<Profile | null>(() => store.profile())
  const [learner, setLearner] = useState(() => store.learner())
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [settings, setSettings] = useState(() => store.settings())
  const [busy, setBusy] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [current, setCurrent] = useState<MentorSession | null>(null)
  const [topicMode, setTopicMode] = useState<'mentor' | 'pick'>('mentor')
  const [picked, setPicked] = useState<ConceptId[]>([])
  const client = useMemo(() => new GeminiClient(settings), [settings])

  useEffect(() => {
    const id = learner.currentWbId
    if (!id) return setCurrent(null)
    const s = store.session(id)
    if (!s) return setCurrent(null)
    void workbookExists(id).then((ok) => setCurrent(ok ? s : null))
  }, [learner.currentWbId, store, workbookExists])

  const rec = useMemo(() => (profile ? recommend(profile, learner) : null), [profile, learner])
  // "Start next problem" from the mentor panel lands on #/mentor/start
  useEffect(() => {
    if (location.hash.startsWith('#/mentor/start') && profile && rec && !busy) {
      history.replaceState(null, '', '#/mentor')
      if (client.available) void startTopic(rec)
      else setSettingsOpen(true)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [profile, rec])

  const openSession = (res: CreatedProblem) => {
    const { doc, pub, targets } = res.built
    store.saveAnswers(pub.problemId, targets)
    store.saveOriginal(doc.id, doc)
    store.saveSession(newSession({ wbId: doc.id, problem: pub, notice: res.notice }))
    const l = store.learner()
    const t = new Date().toISOString()
    store.saveLearner({ ...l, currentWbId: doc.id, history: [...l.history, { problemId: pub.problemId, wbId: doc.id, title: pub.title, level: pub.level, concepts: pub.concepts, mode: pub.mode, startedAt: t, hints: 0, attempts: 0 }] })
    onOpenDoc(doc)
  }

  const run = async (label: string, fn: (step: (s: string) => void) => Promise<CreatedProblem>) => {
    if (!client.available) {
      setError('ExcelMentor needs Gemini. Add your Gemini API key (or a proxy URL) in Mentor settings.')
      setSettingsOpen(true)
      return
    }
    setError(null)
    setBusy(label)
    try {
      openSession(await fn((s) => setBusy(s)))
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  const startTopic = (r: Recommendation) => run('Preparing your problem…', (step) => createTopicProblem(client, profile!, store.learner(), r, step))
  const startAssessment = (files: ExtractedFile[], note: string) => run('Reading your assessment…', (step) => createAssessmentProblem(client, profile!, store.learner(), files, note, step))

  const startOrientation = () => {
    const doc = orientationDoc()
    const t = new Date().toISOString()
    store.saveSession({ version: 2, wbId: doc.id, kind: 'orientation', position: { stage: 'problem', concept: 0 }, attempts: 0, orientationStep: 0, messages: [], lessons: {}, exampleRow: 3, practiceRow: 3, practice: {}, ledger: { added: [], rejected: [] }, pressureTests: {}, stageLog: [], hints: {}, stats: {}, checks: [], startedAt: t, lastActivityAt: t, timeSpentMs: 0, originalDocKey: doc.id })
    store.saveOriginal(doc.id, doc)
    store.saveLearner({ ...store.learner(), currentWbId: doc.id })
    onOpenDoc(doc)
  }

  const header = (
    <header className="flex h-12 shrink-0 items-center gap-2 border-b border-[var(--border)] bg-[var(--panel)] px-4">
      <a href="#/mentor" className="flex items-center gap-1.5 text-[14px] font-semibold" aria-label="ExcelMentor dashboard">
        <Logo size={22} /> <span>ExcelMentor</span>
      </a>
      <div className="ml-auto flex items-center gap-2">
        <button type="button" onClick={() => setSettingsOpen(true)} className={cn('rounded-full px-2.5 py-1 text-[11.5px]', client.available ? 'bg-[var(--active)] text-[var(--accent)]' : 'bg-[var(--warn-bg)] text-[var(--warn-text)]')} data-testid="connection-chip">
          {client.available ? `${settings.connection === 'groq' ? 'Groq' : 'Gemini'}: ${client.model}` : 'AI not connected'}
        </button>
        <Button size="icon" aria-label="Mentor settings" tip="Mentor settings" onClick={() => setSettingsOpen(true)}><Icon name="settings" /></Button>
      </div>
    </header>
  )

  const connectCard = !client.available && (
    <Card title="Connect Gemini" testId="connect-card">
      <p className="text-[13.5px]">ExcelMentor runs every session with Google Gemini: it builds your problem, teaches each stage and checks your reasoning. Add your own Gemini API key (it stays in this browser) or the URL of your organisation’s proxy.</p>
      <Button className="mt-2" variant="default" onClick={() => setSettingsOpen(true)} data-testid="open-settings">Open Mentor settings</Button>
    </Card>
  )

  let body: ReactNode
  if (!profile || route.page === 'profile') {
    body = (
      <Onboarding
        store={store}
        initial={profile}
        onCancel={profile ? () => (location.hash = '#/mentor') : undefined}
        onDone={(p) => {
          setProfile(store.profile() ?? p)
          location.hash = '#/mentor'
        }}
      />
    )
  } else if (route.page === 'takehome') {
    body = <TakeHomePage disabled={!!busy || !client.available} connect={connectCard} error={error} onStart={(files, note) => void startAssessment(files, note)} />
  } else if (route.page === 'concept' && route.concept) {
    body = <ConceptReview store={store} concept={route.concept} learner={learner} onRetry={() => void startTopic(topicRecommendation(profile, learner, [route.concept!]))} />
  } else {
    const completed = learner.history.filter((h) => h.completedAt)
    const introduced = CURRICULUM.filter((c) => conceptProgress(learner, c).stage !== 'locked')
    const mastered = introduced.filter((c) => conceptProgress(learner, c).stage === 'mastered')
    const weak = introduced.filter((c) => {
      const p = conceptProgress(learner, c)
      return p.appliedTotal > 0 && scores(p).application < 70
    })
    const checksScore = completed.length ? Math.round((100 * completed.reduce((a, h) => a + (h.finalScore ?? 0), 0)) / completed.length) : null
    const showOrientation = profile.level === 'never' && !profile.orientationDone
    const pickable = CURRICULUM.filter((c) => GENERATABLE.includes(c))
    body = (
      <div className="mx-auto grid w-full max-w-6xl gap-4 px-4 py-5 lg:grid-cols-3" data-testid="mentor-dashboard">
        <div className="space-y-4 lg:col-span-2">
          {connectCard}
          {showOrientation && (
            <Card title="Start here" testId="orientation-card">
              <p className="mb-2 text-[13.5px]">You told us you've never used Excel. Begin with a 15-step <b>Excel Orientation</b>: workbooks, cells, typing, formatting and your first formulas — each step is a tiny hands-on task.</p>
              <Button variant="default" onClick={startOrientation} data-testid="start-orientation">Start Excel Orientation</Button>
            </Card>
          )}
          {current && (
            <Card title="Continue learning" testId="continue-card">
              <div className="flex flex-wrap items-center gap-3">
                <div className="min-w-0 flex-1">
                  <div className="text-[15px] font-semibold">{current.kind === 'orientation' ? 'Excel Orientation' : current.problem?.title}</div>
                  <div className="text-[12.5px] text-[var(--muted)]">
                    {current.mode === 'takehome' ? 'Take-home · ' : ''}{stageDescription(current)} · last activity {new Date(current.lastActivityAt).toLocaleString()}
                  </div>
                  <div className="mt-2 h-1.5 w-full max-w-sm overflow-hidden rounded-full bg-[var(--panel-2)]">
                    <div className="h-full bg-[var(--accent)]" style={{ width: `${sessionProgress(current)}%` }} />
                  </div>
                </div>
                <Button variant="default" onClick={() => onOpenWorkbook(current.wbId)} data-testid="continue">Continue</Button>
              </div>
            </Card>
          )}
          <div className="grid gap-4 md:grid-cols-2">
            <Card title="1 · Learn by topic" testId="topic-card">
              <div className="mb-2 flex gap-1 rounded-lg bg-[var(--panel-2)] p-0.5 text-[12.5px]" role="tablist">
                {(['mentor', 'pick'] as const).map((m) => (
                  <button key={m} type="button" role="tab" aria-selected={topicMode === m} onClick={() => setTopicMode(m)} className={cn('flex-1 rounded-md px-2 py-1', topicMode === m ? 'bg-[var(--bg)] font-semibold shadow-sm' : 'text-[var(--muted)]')} data-testid={`topic-mode-${m}`}>
                    {m === 'mentor' ? 'Mentor decides' : 'I’ll pick topics'}
                  </button>
                ))}
              </div>
              {topicMode === 'mentor' && rec ? (
                <>
                  <div className="text-[14px] font-semibold" data-testid="recommended">{[...rec.newConcepts, ...rec.revision].map(conceptName).join(', ')}</div>
                  <p className="mt-1 text-[12.5px] text-[var(--muted)]">Chosen for a {profile.level === 'never' ? 'new' : profile.level} {profile.role} in {profile.domain}. Difficulty: <span className="capitalize">{rec.level}</span> · about {rec.estimatedMinutes} min.</p>
                  <p className="mt-1 text-[13px]">{rec.reason}</p>
                  <Button className="mt-3" variant="default" disabled={!!busy || !client.available} onClick={() => void startTopic(rec)} data-testid="start-problem">
                    <Icon name="wand" size={15} /> Start my session
                  </Button>
                </>
              ) : (
                <>
                  <p className="mb-2 text-[12.5px] text-[var(--muted)]">Choose up to {MAX_TOPICS} topics. Your mentor builds a {profile.domain} problem around them.</p>
                  <div className="flex flex-wrap gap-1.5" data-testid="topic-picker">
                    {pickable.map((c) => {
                      const on = picked.includes(c)
                      return (
                        <button key={c} type="button" aria-pressed={on} disabled={!on && picked.length >= MAX_TOPICS} onClick={() => setPicked((p) => (on ? p.filter((x) => x !== c) : [...p, c]))} className={cn('rounded-full border px-2.5 py-0.5 text-[12px] disabled:opacity-40', on ? 'border-[var(--accent)] bg-[var(--active)] font-semibold text-[var(--accent)]' : 'border-[var(--border)] hover:border-[var(--border-strong)]')}>
                          {CONCEPTS[c].name}
                        </button>
                      )
                    })}
                  </div>
                  <Button className="mt-3" variant="default" disabled={!!busy || !client.available || !picked.length} onClick={() => void startTopic(topicRecommendation(profile, learner, picked))} data-testid="start-picked">
                    <Icon name="wand" size={15} /> Start with {picked.length ? picked.map(conceptName).join(', ') : 'my topics'}
                  </Button>
                </>
              )}
            </Card>
            <Card title="2 · Take-home assessment" testId="takehome-card">
              <p className="text-[13.5px]">Got an Excel assessment from an employer? Upload the brief and data (Excel, CSV, PDF or Word). Your mentor turns your real tasks into a guided session and teaches every skill they need — without doing the assessment for you.</p>
              <Button className="mt-3" variant="default" disabled={!client.available} onClick={() => (location.hash = '#/mentor/takehome')} data-testid="open-takehome">
                <Icon name="upload" size={15} /> Upload my assessment
              </Button>
            </Card>
          </div>
          {error && <p className="rounded-lg bg-[var(--warn-bg)] px-3 py-2 text-[13px] text-[var(--warn-text)]" role="alert">{error}</p>}
          <Card title="Concept library">
            <ConceptLibrary learner={learner} />
          </Card>
        </div>
        <div className="space-y-4">
          <Card title="My progress" testId="my-progress">
            <dl className="grid grid-cols-2 gap-x-3 gap-y-2 text-[13px]">
              {(
                [
                  ['Sessions completed', completed.length],
                  ['Sessions started', learner.history.length],
                  ['Concepts learned', introduced.length],
                  ['Concepts mastered', mastered.length],
                  ['Needs improvement', weak.length ? weak.map(conceptName).join(', ') : '—'],
                  ['Hints used', learner.history.reduce((a, h) => a + h.hints, 0)],
                  ['Check My Work score', checksScore === null ? '—' : `${checksScore}%`],
                ] as [string, ReactNode][]
              ).map(([k, v]) => (
                <div key={k}>
                  <dt className="text-[11.5px] text-[var(--muted)]">{k}</dt>
                  <dd className="font-semibold">{v}</dd>
                </div>
              ))}
            </dl>
          </Card>
          <Card title="History">
            {learner.history.length === 0 ? (
              <p className="text-[12.5px] text-[var(--muted)]">Your sessions will appear here.</p>
            ) : (
              <ul className="space-y-1.5 text-[12.5px]">
                {[...learner.history].reverse().slice(0, 12).map((h) => (
                  <li key={h.problemId} className="flex items-center gap-2">
                    <span className={cn('h-2 w-2 shrink-0 rounded-full', h.completedAt ? 'bg-emerald-500' : 'bg-amber-400')} />
                    <button type="button" className="min-w-0 flex-1 truncate text-left hover:underline" onClick={() => onOpenWorkbook(h.wbId)}>{h.mode === 'takehome' ? '📎 ' : ''}{h.title}</button>
                    <span className="shrink-0 text-[var(--muted)]">{h.completedAt ? `${Math.round(100 * (h.finalScore ?? 0))}%${h.assisted ? ' · assisted' : ''}` : 'in progress'}</span>
                  </li>
                ))}
              </ul>
            )}
          </Card>
          <Card title="Profile">
            <p className="text-[12.5px]">
              {LEVELS.find((l) => l[0] === profile.level)?.[1]} · {profile.role} · {profile.domain} · {profile.duration} min sessions
            </p>
            <div className="mt-2 flex flex-wrap gap-2">
              <Button size="sm" variant="outline" onClick={() => (location.hash = '#/mentor/profile')}>Edit profile</Button>
              {profile.level === 'never' && profile.orientationDone && <Button size="sm" variant="ghost" onClick={startOrientation}>Repeat orientation</Button>}
              <Button
                size="sm"
                variant="ghost"
                onClick={() => {
                  if (!confirm('Delete your ExcelMentor profile and progress on this device?')) return
                  store.resetAll()
                  setProfile(null)
                  setLearner(store.learner())
                }}
              >
                Reset progress
              </Button>
            </div>
          </Card>
        </div>
      </div>
    )
  }

  return (
    <div className="flex h-full flex-col bg-[var(--bg)]">
      {header}
      <main className="min-h-0 flex-1 overflow-auto">{body}</main>
      {busy && (
        <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/25" role="status" data-testid="mentor-busy">
          <div className="flex max-w-md items-center gap-3 rounded-xl bg-[var(--bg)] px-5 py-4 text-[13px] shadow-2xl">
            <span className="h-4 w-4 shrink-0 animate-spin rounded-full border-2 border-[var(--accent)] border-t-transparent" />
            {busy}
          </div>
        </div>
      )}
      {settingsOpen && <MentorSettingsDialog store={store} onClose={() => setSettingsOpen(false)} onSaved={(s) => setSettings(s)} />}
    </div>
  )
}

// =====================================================================================
interface UploadItem {
  id: string
  name: string
  status: 'reading' | 'ok' | 'error'
  file?: ExtractedFile
  error?: string
}

function TakeHomePage({ disabled, connect, error, onStart }: { disabled: boolean; connect: ReactNode; error: string | null; onStart: (files: ExtractedFile[], note: string) => void }) {
  const [items, setItems] = useState<UploadItem[]>([])
  const [note, setNote] = useState('')
  const [drag, setDrag] = useState(false)
  const add = (files: File[]) => {
    for (const f of files.slice(0, 6)) {
      const id = Math.random().toString(36).slice(2)
      setItems((x) => [...x, { id, name: f.name, status: 'reading' }])
      extractFile(f)
        .then((file) => setItems((x) => x.map((i) => (i.id === id ? { ...i, status: 'ok', file } : i))))
        .catch((e) => setItems((x) => x.map((i) => (i.id === id ? { ...i, status: 'error', error: e instanceof Error ? e.message : String(e) } : i))))
    }
  }
  const ready = items.filter((i) => i.status === 'ok' && i.file).map((i) => i.file!)
  const reading = items.some((i) => i.status === 'reading')
  const summary = (f: ExtractedFile) => (f.kind === 'spreadsheet' ? f.text : `${f.text.length.toLocaleString()} characters of text`)
  return (
    <div className="mx-auto w-full max-w-3xl space-y-4 px-4 py-5" data-testid="takehome-page">
      <a href="#/mentor" className="text-[12.5px] text-[var(--accent)] hover:underline">← Dashboard</a>
      <h1 className="text-[22px] font-semibold">Take-home assessment</h1>
      <p className="text-[13.5px] text-[var(--muted)]">Upload everything the employer sent: the brief (PDF, Word or text) and the data (Excel or CSV). Your mentor reads the brief, puts your real data on the Task sheet with a yellow cell for each answer, and takes you through the same nine stages as any session — teaching the skills on separate practice data. It won’t write your answers.</p>
      {connect}
      <label
        className={cn('flex cursor-pointer flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed px-4 py-8 text-center text-[13px]', drag ? 'border-[var(--accent)] bg-[var(--active)]' : 'border-[var(--border-strong)] hover:border-[var(--accent)]')}
        onDragOver={(e) => (e.preventDefault(), setDrag(true))}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => {
          e.preventDefault()
          setDrag(false)
          add([...e.dataTransfer.files])
        }}
      >
        <Icon name="upload" size={22} />
        <span><b>Choose files</b> or drop them here</span>
        <span className="text-[12px] text-[var(--muted)]">{ACCEPTED_EXTENSIONS.join(' ')} · up to 6 files, 25 MB each</span>
        <input type="file" multiple accept={ACCEPTED_EXTENSIONS.join(',')} className="sr-only" aria-label="Assessment files" data-testid="takehome-files" onChange={(e) => (add([...(e.target.files ?? [])]), (e.target.value = ''))} />
      </label>
      {items.length > 0 && (
        <ul className="space-y-1.5 text-[13px]" data-testid="takehome-list">
          {items.map((i) => (
            <li key={i.id} className="flex items-start gap-2 rounded-lg border border-[var(--border)] px-3 py-2">
              <Icon name={i.file?.kind === 'spreadsheet' ? 'sheet' : 'file'} size={15} className="mt-0.5 shrink-0" />
              <div className="min-w-0 flex-1">
                <div className="truncate font-semibold">{i.name}</div>
                <div className={cn('text-[12px]', i.status === 'error' ? 'text-[var(--danger)]' : 'text-[var(--muted)]')}>
                  {i.status === 'reading' ? 'Reading…' : i.status === 'error' ? i.error : summary(i.file!)}
                </div>
              </div>
              <button type="button" className="text-[var(--muted)] hover:text-[var(--danger)]" aria-label={`Remove ${i.name}`} onClick={() => setItems((x) => x.filter((y) => y.id !== i.id))}>✕</button>
            </li>
          ))}
        </ul>
      )}
      <Field label="Anything your mentor should know? (optional)" hint="e.g. the role you applied for, the deadline, or parts of the brief you find confusing.">
        <textarea className={cn(inputClass, 'h-20 py-1.5')} value={note} onChange={(e) => setNote(e.target.value)} aria-label="Note for your mentor" />
      </Field>
      <p className="text-[12px] text-[var(--muted)]">Your files are read in this browser. Their text and a sample of the data are sent to Gemini to build the session.</p>
      {error && <p className="rounded-lg bg-[var(--warn-bg)] px-3 py-2 text-[13px] text-[var(--warn-text)]" role="alert">{error}</p>}
      <Button variant="default" disabled={disabled || reading || !ready.length} onClick={() => onStart(ready, note)} data-testid="start-takehome">
        <Icon name="wand" size={15} /> Build my mentoring session
      </Button>
    </div>
  )
}

function ConceptLibrary({ learner }: { learner: ReturnType<MentorStore['learner']> }) {
  const groups: ConceptCategory[] = ['Excel Foundations', 'Functions', 'Data Analysis']
  const all = [...FOUNDATIONS, ...CURRICULUM.filter((c) => !FOUNDATIONS.includes(c))]
  return (
    <div className="grid gap-4 sm:grid-cols-3">
      {groups.map((g) => (
        <div key={g}>
          <div className="mb-1.5 text-[12.5px] font-semibold">{g}</div>
          <ul className="space-y-1">
            {all
              .filter((c) => CONCEPTS[c].category === g)
              .map((c) => {
                const p = conceptProgress(learner, c)
                const locked = p.stage === 'locked'
                return (
                  <li key={c}>
                    <a href={`#/mentor/concept/${c}`} className={cn('flex items-center justify-between gap-2 rounded-md px-2 py-1 text-[12.5px] hover:bg-[var(--hover)]', locked && 'opacity-60')} title={locked ? 'Not introduced yet — preview' : STAGE_LABEL[p.stage]}>
                      <span className="truncate">{locked && <Icon name="lock" size={11} className="mr-1 inline align-[-1px]" />}{CONCEPTS[c].name}</span>
                      <span className={cn('shrink-0 rounded-full px-1.5 py-0.5 text-[10.5px]', STAGE_COLOR[p.stage])}>{locked ? 'Future' : STAGE_LABEL[p.stage]}</span>
                    </a>
                  </li>
                )
              })}
          </ul>
        </div>
      ))}
    </div>
  )
}

function MiniGrid({ rows }: { rows: { text: string; formula?: string; target?: boolean }[][] }) {
  const width = Math.max(...rows.map((r) => r.length))
  return (
    <div className="overflow-auto rounded-md border border-[var(--border)]">
      <table className="border-collapse text-[12px]">
        <thead>
          <tr>
            <th className="w-8 border border-[var(--border)] bg-[var(--panel-2)]" />
            {Array.from({ length: width }, (_, c) => <th key={c} className="min-w-16 border border-[var(--border)] bg-[var(--panel-2)] px-2 font-normal text-[var(--muted)]">{colToLetters(c)}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, r) => (
            <tr key={r}>
              <td className="border border-[var(--border)] bg-[var(--panel-2)] px-1 text-center text-[var(--muted)]">{r + 2}</td>
              {row.map((cell, c) => (
                <td key={c} className={cn('border border-[var(--border)] px-2 py-0.5', r === 0 && 'bg-[#0f766e] font-semibold text-white', cell.target && 'bg-[#fff7d1] text-[#3d2c00]')} title={cell.formula}>
                  {cell.formula ? <span><span className="font-semibold">{cell.text}</span> <code className="ml-1 text-[11px] text-[var(--muted)]">{cell.formula}</code></span> : cell.text}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function ConceptReview({ store, concept, learner, onRetry }: { store: MentorStore; concept: ConceptId; learner: ReturnType<MentorStore['learner']>; onRetry: () => void }) {
  const def = CONCEPTS[concept]
  const lesson = builtinLesson(concept)
  const ex = renderBlock(lesson.example, true)
  const pr = renderBlock(lesson.practice, false)
  const p = conceptProgress(learner, concept)
  const sc = scores(p)
  const [answer, setAnswer] = useState('')
  const [fb, setFb] = useState<{ ok: boolean; message: string } | null>(null)
  return (
    <div className="mx-auto w-full max-w-4xl space-y-4 px-4 py-5" data-testid="concept-review">
      <a href="#/mentor" className="text-[12.5px] text-[var(--accent)] hover:underline">← Dashboard</a>
      <div className="flex flex-wrap items-center gap-3">
        <h1 className="text-[22px] font-semibold">{def.name}</h1>
        <span className={cn('rounded-full px-2 py-0.5 text-[11.5px]', STAGE_COLOR[p.stage])}>{p.stage === 'locked' ? 'Not introduced yet' : STAGE_LABEL[p.stage]}</span>
      </div>
      <p className="text-[14px]">{def.summary}</p>
      <div className="grid gap-4 md:grid-cols-2">
        <Card title="Simple explanation">
          <p className="text-[13.5px]">{lesson.realLife}</p>
          <p className="mt-2 text-[13.5px]">{lesson.excel}</p>
          {def.prereqs.length > 0 && <p className="mt-2 text-[12px] text-[var(--muted)]">Builds on: {def.prereqs.map(conceptName).join(', ')}</p>}
        </Card>
        <Card title="Reference video">
          <VideoPlayer concept={concept} store={store} />
        </Card>
      </div>
      <Card title="Workbook example">
        <p className="mb-2 text-[13px]">{lesson.example.title}</p>
        {ex && <MiniGrid rows={ex.rows} />}
        <p className="mt-2 text-[13px]">{lesson.formulaExplanation}</p>
      </Card>
      <Card title="Mini practice" testId="mini-practice">
        <p className="mb-2 text-[13px]">{lesson.practice.prompt} Write the formula for the yellow cell {pr?.resultCell} using the addresses shown.</p>
        {pr && <MiniGrid rows={pr.rows} />}
        <form className="mt-2 flex gap-2" onSubmit={(e) => (e.preventDefault(), setFb(checkBlockAnswer(lesson.practice, concept, answer)))}>
          <input className={cn(inputClass, 'flex-1 font-mono')} value={answer} onChange={(e) => (setAnswer(e.target.value), setFb(null))} placeholder="=" aria-label="Your formula" />
          <Button variant="outline" type="submit" disabled={!answer.trim()}>Check</Button>
        </form>
        {fb && <p className={cn('mt-2 text-[13px]', fb.ok ? 'text-[var(--accent)]' : 'text-[var(--warn-text)]')} role="status">{fb.ok ? '✅ ' : ''}{fb.message}</p>}
      </Card>
      <Card title="Your previous performance">
        {p.stage === 'locked' ? (
          <p className="text-[13px] text-[var(--muted)]">You haven't learned this concept yet.</p>
        ) : (
          <div className="grid grid-cols-2 gap-2 text-[13px] sm:grid-cols-4">
            <div><div className="text-[11.5px] text-[var(--muted)]">Understanding</div><b>{sc.understanding}%</b></div>
            <div><div className="text-[11.5px] text-[var(--muted)]">Practice</div><b>{sc.practice}%</b></div>
            <div><div className="text-[11.5px] text-[var(--muted)]">Application</div><b>{sc.application}%</b></div>
            <div><div className="text-[11.5px] text-[var(--muted)]">Mastery</div><b>{STAGE_LABEL[p.stage]}</b></div>
          </div>
        )}
        <Button className="mt-3" variant="default" onClick={onRetry} disabled={!GENERATABLE.includes(concept)} data-testid="retry-concept">
          Practise {def.name} with a new problem
        </Button>
      </Card>
    </div>
  )
}
