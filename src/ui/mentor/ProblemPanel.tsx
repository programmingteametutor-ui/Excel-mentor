/* The mentor for a problem (topic or take-home). Gemini drives the conversation through the nine stages using the
 * teaching protocol; the app supplies the evidence (formula-engine checks), enforces the gates and blocks answer leaks:
 * 1 problem → 2 roadmap → 3 learn ⇄ 4 practice (per concept) → 5 plan → 6 solve → 7 check → 8 follow-up → 9 report. */
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import type { EditorBridge } from '../Editor'
import { colToLetters } from '../../model/address'
import { exportXlsx } from '../../files/xlsx'
import { CONCEPTS, conceptName } from '../../mentor/concepts'
import { computeExpected, SHEETS } from '../../mentor/build'
import { GeminiClient } from '../../mentor/gemini'
import { generateLesson, mentorTurn } from '../../mentor/ai'
import { checkPractice, writeExample, writePractice } from '../../mentor/lesson'
import { describeInputs } from '../../mentor/hints'
import { checkWork, diagnoseCell, simpleState } from '../../mentor/validate'
import { buildReport, markIntroduced, nextLevel, performanceScore, recommend, recordPractice, recordProblem, recordThink } from '../../mentor/progress'
import { applyTurn, currentConcept, findLeak, groupOf, MAX_HINT, missingEvidence, nextPosition, problemBrief, stageIndex, stateText, taskGroups, type LeakGuard } from '../../mentor/flow'
import { chunkParagraphs, CLOSING_PHRASE, STAGE_INFO, systemPrompt } from '../../mentor/protocol'
import type { MentorStore } from '../../mentor/store'
import type { CellState, CellTarget, ChatMessage, CheckRecord, ConceptId, MentorSession, MentorTurn, Profile, StagePosition, Target, TargetResult } from '../../mentor/types'
import { Button, Dialog, inputClass, MenuItem, MenuSep } from '../components/ui'
import { Icon } from '../icons'
import { cn, download, safeFileName } from '../lib/utils'
import { MentorSettingsDialog } from './Settings'
import { Bubble, now, PanelHeader, RichText, StageStepper, STATE_ICON, STATE_LABEL, statOf, uid } from './parts'

const START_EVENT = '[App] The session has started. Begin STAGE 1: present the business problem — the situation, why it matters, the data available, and the functional and non-functional requirements. Teach only that in this response and end with exactly one question.'

export function ProblemPanel({ bridge, store, initial, onCollapse, onDashboard, onNextProblem }: { bridge: EditorBridge; store: MentorStore; initial: MentorSession; onCollapse: () => void; onDashboard: () => void; onNextProblem: () => void }) {
  const { wb } = bridge
  const [session, setSession] = useState<MentorSession>(initial)
  const sessionRef = useRef(session)
  sessionRef.current = session
  const pub = session.problem!
  const targets = useMemo<Target[]>(() => store.answers(pub.problemId) ?? [], [store, pub.problemId])
  const cellTargets = useMemo(() => targets.filter((t): t is CellTarget => t.kind === 'cell'), [targets])
  const original = useMemo(() => store.original(wb.id), [store, wb.id])
  const groups = useMemo(() => taskGroups(targets), [targets])
  const expected = useMemo(() => (original ? computeExpected(original, cellTargets) : {}), [original, cellTargets])
  const profile = useMemo<Profile>(() => store.profile() ?? { level: 'beginner', role: 'Other', domain: pub.domain, description: '', attachments: [], difficulty: 'normal', duration: 30, createdAt: now() }, [store, pub.domain])
  const beginner = profile.level === 'never' || profile.level === 'beginner'
  const system = useMemo(() => systemPrompt({ profile, problem: pub, problemBrief: problemBrief(pub, targets, original) }), [profile, pub, targets, original])
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [settingsVersion, setSettingsVersion] = useState(0)
  const client = useMemo(() => new GeminiClient(store.settings()), [store, settingsVersion]) // eslint-disable-line react-hooks/exhaustive-deps
  const [busy, setBusy] = useState<string | null>(null)
  const [view, setView] = useState<'lesson' | 'results' | 'report'>(session.completed ? 'report' : 'lesson')
  const [answer, setAnswer] = useState('')
  const [dockNote, setDockNote] = useState<{ id: string; text: string } | null>(null)
  const [confirm, setConfirm] = useState<null | 'sheet' | 'problem'>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const inFlight = useRef(false)

  const update = useCallback((fn: (s: MentorSession) => MentorSession) => {
    const n = fn(sessionRef.current)
    n.lastActivityAt = now()
    sessionRef.current = n
    setSession(n)
  }, [])
  useEffect(() => {
    const t = setTimeout(() => store.saveSession(session), 250)
    return () => clearTimeout(t)
  }, [session, store])
  useEffect(() => () => store.saveSession(sessionRef.current), [store])
  useEffect(() => {
    listRef.current?.scrollTo({ top: listRef.current.scrollHeight, behavior: 'smooth' })
  }, [session.messages.length, busy, view])
  useEffect(() => {
    const t = setInterval(() => {
      if (document.visibilityState === 'visible') update((s) => ({ ...s, timeSpentMs: s.timeSpentMs + 15000 }))
    }, 15000)
    return () => clearInterval(t)
  }, [update])

  const post = useCallback((text: string, from: ChatMessage['from'] = 'mentor', kind?: ChatMessage['kind']) => update((s) => ({ ...s, messages: [...s.messages, { id: uid(), from, text, kind, at: now() }] })), [update])
  const saveLearner = (fn: (l: ReturnType<typeof store.learner>) => ReturnType<typeof store.learner>) => store.saveLearner(fn(store.learner()))

  // ----------------------------------------------------------------- attempts per task cell (for hints and the report)
  useEffect(() => {
    const changes: Record<string, string> = {}
    for (const t of cellTargets) {
      const sid = wb.sheetByName(t.sheet)?.id
      if (!sid) continue
      const content = wb.getEditText(sid, t.r, t.c)
      const prev = sessionRef.current.stats[t.id]?.lastContent
      if (prev === undefined ? content !== '' : content !== prev) changes[t.id] = content
    }
    if (!Object.keys(changes).length) return
    update((s) => {
      const stats = { ...s.stats }
      for (const [id, content] of Object.entries(changes)) {
        const st = statOf(stats[id])
        st.lastContent = content
        if (content !== '') {
          st.attempts += 1
          st.firstAttemptAt ??= now()
        }
        stats[id] = st
      }
      return { ...s, stats }
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [bridge.version, cellTargets])

  // ----------------------------------------------------------------- example and practice sheets, prepared just in time
  const prepareConcept = async (c: ConceptId) => {
    let s = sessionRef.current
    if (s.exampleCells?.[c] && s.practice[c]) return
    let lesson = s.lessons[c]
    if (!lesson) {
      setBusy(`Preparing a worked example and practice for ${conceptName(c)}…`)
      lesson = await generateLesson(client, c, pub, profile)
      update((x) => ({ ...x, lessons: { ...x.lessons, [c]: lesson! } }))
      s = sessionRef.current
    }
    if (!s.exampleCells?.[c]) {
      const ex = writeExample(wb, lesson, s.exampleRow, true)
      if (!('error' in ex)) {
        const sid = wb.sheetByName(SHEETS.example)!.id
        const cell = `${colToLetters(ex.cell.c)}${ex.cell.r + 1}`
        const range = `A${ex.firstDataRow}:${colToLetters(Math.max(ex.headers.length - 1, ex.cell.c))}${ex.nextRow - 2}`
        update((x) => ({ ...x, exampleRow: ex.nextRow, exampleCells: { ...x.exampleCells, [c]: { cell, formula: ex.formula, value: wb.getDisplay(sid, ex.cell.r, ex.cell.c).text, range } } }))
      }
    }
    if (!sessionRef.current.practice[c]) {
      const res = writePractice(wb, lesson, sessionRef.current.practiceRow)
      if (!('error' in res)) update((x) => ({ ...x, practiceRow: res.nextRow, practice: { ...x.practice, [c]: res.target } }))
    }
  }

  /** Side effects of arriving at a position (sheets to look at, progress records). */
  const arrive = (p: StagePosition) => {
    const c = p.stage === 'learn' || p.stage === 'practice' ? pub.newConcepts[p.concept] : null
    if (p.stage === 'learn' && c) {
      saveLearner((l) => markIntroduced(l, c))
      const ex = sessionRef.current.exampleCells?.[c]
      if (ex) bridge.select(SHEETS.example, Number(ex.cell.replace(/\D+/g, '')) - 1, ex.cell.replace(/\d+/g, '').split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1)
    }
    if (p.stage === 'practice' && c) {
      const pt = sessionRef.current.practice[c]
      if (pt) bridge.select(SHEETS.practice, pt.r, pt.c)
    }
    if (p.stage === 'problem') bridge.select(SHEETS.data, 0, 0)
    if (p.stage === 'plan' || p.stage === 'solve') bridge.select(SHEETS.task, 0, 0)
  }

  // ----------------------------------------------------------------- one exchange with the mentor
  const leakGuard = (authorize?: string): LeakGuard => {
    const s = sessionRef.current
    const c = currentConcept(s)
    const pt = c ? s.practice[c] : undefined
    return {
      targets: cellTargets,
      expected,
      allowGroups: new Set(pub.mode === 'topic' && authorize ? [authorize] : []),
      practice: pt ? { cell: pt.cell, formula: pt.formula, level: s.hints[`practice-${c}`] ?? 0 } : undefined,
    }
  }

  const send = async (event: string, opts: { show?: { text: string; from: ChatMessage['from']; kind?: ChatMessage['kind'] }; authorize?: string; hint?: { key: string; level: number } } = {}) => {
    if (inFlight.current) return
    inFlight.current = true
    const history = sessionRef.current.messages.filter((m) => m.kind !== 'error')
    if (opts.show) post(opts.show.text, opts.show.from, opts.show.kind)
    if (opts.show?.from === 'student') update((s) => ({ ...s, attempts: s.attempts + 1 }))
    update((s) => ({ ...s, pending: undefined, messages: s.messages.filter((m) => m.kind !== 'error') }))
    try {
      // the next concept's example and practice exist before the reply that moves there
      const s0 = sessionRef.current
      const nxt = nextPosition(s0.position, pub.newConcepts.length)
      const cur = currentConcept(s0)
      if (cur) await prepareConcept(cur)
      if ((nxt.stage === 'learn' || nxt.stage === 'practice') && s0.position.stage !== 'report') await prepareConcept(pub.newConcepts[nxt.concept])
      setBusy('Your mentor is thinking…')
      const s = sessionRef.current
      const guard = leakGuard(opts.authorize)
      const res = await mentorTurn({
        client,
        system,
        history,
        state: stateText(s, targets),
        event,
        stage: s.position.stage,
        concept: s.position.concept,
        check: {
          next: nextPosition(s.position, pub.newConcepts.length),
          missingEvidence: missingEvidence(s),
          appDriven: s.position.stage === 'solve',
          leaks: (text, level) => findLeak(text, guard, level),
          closingPhrase: CLOSING_PHRASE[pub.mode],
        },
      })
      let turn: MentorTurn = res.turn
      const hard = res.problems.filter((p) => p.hard)
      if (hard.some((p) => p.rule.startsWith('No Spoiling'))) {
        post('Your mentor’s reply gave away part of an answer, so it was hidden. Retry, or rephrase what you asked.', 'mentor', 'error')
        update((x) => ({ ...x, pending: { event } }))
        return
      }
      if (hard.length) {
        // still breaking a gate rule after the correction round: keep the student where they are
        turn = { ...turn, current_stage: s.position.stage, concept_index: s.position.concept, gate_passed: s.position.stage === 'report' ? turn.gate_passed : false, session_complete: hard.some((p) => /session_complete|Mastery Profile|final hand-off/.test(p.rule)) ? false : turn.session_complete }
      }
      post(chunkParagraphs(turn.message_to_student))
      const before = sessionRef.current
      let after = applyTurn(before, turn, now())
      if (opts.hint) after = { ...after, hints: { ...after.hints, [opts.hint.key]: Math.max(after.hints[opts.hint.key] ?? 0, opts.hint.level) } }
      if (opts.hint && opts.hint.level >= MAX_HINT && !opts.hint.key.startsWith('practice-')) after = { ...after, teachBack: opts.hint.key }
      if (turn.teach_back_passed) after = { ...after, teachBack: undefined }
      update(() => after)
      // learner progress from the gates
      const c = currentConcept(before)
      if (c && before.position.stage === 'learn' && after.position.stage !== 'learn') saveLearner((l) => recordThink(l, c, turn.section_score / 100))
      if (after.position.stage !== before.position.stage || after.position.concept !== before.position.concept) arrive(after.position)
      if (turn.session_complete && before.position.stage === 'report') complete(turn.mastery_profile)
    } catch (e) {
      post(`The mentor couldn't reply: ${e instanceof Error ? e.message : String(e)}`, 'mentor', 'error')
      update((x) => ({ ...x, pending: { event } }))
    } finally {
      setBusy(null)
      inFlight.current = false
    }
  }

  // first visit
  useEffect(() => {
    if (!sessionRef.current.messages.length) {
      arrive(sessionRef.current.position)
      void send(START_EVENT)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const retry = () => {
    const p = sessionRef.current.pending
    if (p) void send(p.event)
  }

  const submitAnswer = () => {
    const text = answer.trim()
    if (!text || busy) return
    setAnswer('')
    void send(`Student says:\n"""${text}"""`, { show: { text, from: 'student' } })
  }

  // ----------------------------------------------------------------- practice (stage 4)
  const curConcept = currentConcept(session)
  const practiceTarget = curConcept ? session.practice[curConcept] : undefined
  const doCheckPractice = () => {
    if (!curConcept || !practiceTarget) return
    const res = checkPractice(wb, practiceTarget)
    const sid = wb.sheetByName(SHEETS.practice)!.id
    const entry = wb.getEditText(sid, practiceTarget.r, practiceTarget.c)
    const first = practiceTarget.attempts === 0
    update((s) => ({ ...s, practice: { ...s.practice, [curConcept]: { ...practiceTarget, attempts: practiceTarget.attempts + 1, passed: practiceTarget.passed || res.ok } } }))
    if (res.state !== 'not_attempted') saveLearner((l) => recordPractice(l, curConcept, res.ok && first))
    bridge.select(SHEETS.practice, practiceTarget.r, practiceTarget.c)
    const note = res.ok ? `Practice check: ${practiceTarget.cell} is correct.` : `Practice check: ${practiceTarget.cell} — ${res.message}`
    void send(`[Engine] Practice check for ${conceptName(curConcept)} in ${SHEETS.practice}!${practiceTarget.cell}: ${res.ok ? 'CORRECT' : `NOT CORRECT (${res.state.replace(/_/g, ' ')})`}. Engine message: ${res.message} Student's entry: ${entry || '(empty)'}. ${res.ok ? 'Now apply the practice rubric (justification tied to a requirement + one rejected alternative) before the gate can pass.' : 'Point to the problem without giving the formula; use the hint protocol if they are stuck.'}`, { show: { text: note, from: 'mentor', kind: 'event' } })
  }
  const practiceHint = () => {
    if (!curConcept || !practiceTarget) return
    const key = `practice-${curConcept}`
    const level = Math.min(MAX_HINT, (session.hints[key] ?? 0) + 1)
    const sid = wb.sheetByName(SHEETS.practice)!.id
    const entry = wb.getEditText(sid, practiceTarget.r, practiceTarget.c)
    void send(`[Hint request] The student asked for a hint on the practice cell ${SHEETS.practice}!${practiceTarget.cell} (${conceptName(curConcept)}). Their entry: ${entry || '(empty)'}. Give Hint Level ${level} of ${MAX_HINT} of the Adaptive Pivot protocol${level === MAX_HINT ? ' (Reveal & Reflect: show the practice formula with full reasoning, then require the teach-back)' : ''}.`, { show: { text: `Hint ${level} of ${MAX_HINT} requested for ${practiceTarget.cell}.`, from: 'mentor', kind: 'event' }, hint: { key, level } })
  }

  // ----------------------------------------------------------------- hints for the selected task cell (stages 6–7)
  const sheetName = wb.meta(bridge.sheetId).name
  const { r: ar, c: ac } = bridge.selection.anchor
  const selTarget = cellTargets.find((t) => t.sheet === sheetName && t.r === ar && t.c === ac)
  useEffect(() => {
    if (dockNote && dockNote.id !== selTarget?.id) setDockNote(null)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selTarget?.id])

  const requestHint = (t: CellTarget) => {
    const key = groupOf(t.id)
    const level = Math.min(MAX_HINT, (session.hints[key] ?? 0) + 1)
    const st = statOf(session.stats[t.id])
    if ((session.hints[key] ?? 0) >= MAX_HINT) return setDockNote({ id: t.id, text: 'You have had every hint level for this task. Ask your mentor a question in the chat if something is still unclear.' })
    if (level === MAX_HINT && st.attempts < 1) return setDockNote({ id: t.id, text: 'Try writing a formula in this cell first using the hints so far. Level 4 unlocks after a real attempt.' })
    const sid = wb.sheetByName(t.sheet)!.id
    const entry = wb.getEditText(sid, t.r, t.c)
    const diag = entry ? diagnoseCell(wb, t, expected[t.id]) : null
    update((s) => ({ ...s, stats: { ...s.stats, [t.id]: { ...statOf(s.stats[t.id]), hintLevel: Math.max(statOf(s.stats[t.id]).hintLevel, level), hintsUsed: statOf(s.stats[t.id]).hintsUsed + 1 } } }))
    const reveal = level === MAX_HINT ? (pub.mode === 'topic' ? ' (Reveal & Reflect: you may now give the formula for this task with the full reasoning, then require the teach-back)' : ' (Reveal & Reflect for a take-home: show the full reasoning and a worked formula on a PARALLEL example with different data — never this cell, never their numbers — then require the teach-back)') : ''
    void send(`[Hint request] Cell ${t.sheet}!${t.cell} — task group ${key}: "${t.requirement}" (concept: ${conceptName(t.concept)}). It needs ${describeInputs(wb, t)}. Student's entry: ${entry || '(empty)'}. Engine diagnosis: ${diag ? diag.message : 'no entry yet'}. Attempts on this cell: ${st.attempts}. Give Hint Level ${level} of ${MAX_HINT}${reveal}.`, { show: { text: `Hint ${level} of ${MAX_HINT} requested for ${t.cell}.`, from: 'mentor', kind: 'event' }, authorize: level === MAX_HINT ? key : undefined, hint: { key, level } })
  }
  const checkThisCell = (t: CellTarget) => {
    const d = diagnoseCell(wb, t, expected[t.id])
    setDockNote({ id: t.id, text: d.state === 'correct' ? '✅ This cell is correct.' : `${STATE_ICON(d.state)} ${d.message}` })
  }

  // ----------------------------------------------------------------- Check My Work (stage 7)
  const runCheck = () => {
    const results = checkWork(wb, targets, original)
    const correct = results.filter((r) => r.state === 'correct').length
    const rec: CheckRecord = { at: now(), results, correct, total: results.length }
    const checks = [...sessionRef.current.checks, rec]
    // Solve ends when the student checks their work; its score is the first check's accuracy
    const hintsUsed = Object.entries(sessionRef.current.hints).filter(([k]) => !k.startsWith('practice-')).reduce((a, [, v]) => a + v, 0)
    update((s) =>
      s.position.stage === 'solve'
        ? { ...s, checks, position: { stage: 'check' as const, concept: 0 }, attempts: 0, stageLog: [...s.stageLog, { stage: 'solve' as const, passedAt: now(), summary: `First Check My Work: ${correct}/${results.length} correct, with ${hintsUsed} hint level${hintsUsed === 1 ? '' : 's'} used.`, score: Math.round((100 * correct) / Math.max(1, results.length)), attempts: s.attempts }] }
        : { ...s, checks },
    )
    const wrong = results.filter((r) => r.state !== 'correct')
    const keys = wrong.map((r) => targets.find((t) => t.id === r.id)).filter((t): t is CellTarget => t?.kind === 'cell').map((t) => `${t.r},${t.c}`)
    bridge.setHighlight(keys.length ? { sheet: SHEETS.task, keys } : null)
    if (wrong.length) setView('results')
    const describe = (r: TargetResult) => {
      const t = targets.find((x) => x.id === r.id)
      return `${t?.kind === 'cell' ? t.cell : t?.requirement ?? r.id} — ${r.state.replace(/_/g, ' ')}: ${r.message}`
    }
    void send(`[Engine] Check My Work #${checks.length}: ${correct}/${results.length} correct.${wrong.length ? `\nNot correct:\n${wrong.slice(0, 25).map(describe).join('\n')}${wrong.length > 25 ? `\n…and ${wrong.length - 25} more` : ''}` : ' Every Task cell is correct — begin the deep dive (7A).'}`, { show: { text: `Check My Work #${checks.length}: ${correct}/${results.length} correct.`, from: 'mentor', kind: 'event' } })
  }

  const jumpTo = (t: Target) => {
    if (t.kind === 'cell') {
      bridge.select(t.sheet, t.r, t.c)
      bridge.setHighlight({ sheet: t.sheet, keys: [`${t.r},${t.c}`] })
    } else bridge.select(t.sheet, 1, 0)
    setView('lesson')
  }

  // ----------------------------------------------------------------- completion (stage 9)
  const complete = (masteryProfile: string) => {
    const s = sessionRef.current
    const checks = s.checks
    const perf = performanceScore(pub, checks, s.stats)
    let l = checks.length ? recordProblem(store.learner(), pub, checks[checks.length - 1], s.stats) : store.learner()
    l = {
      ...l,
      nextLevel: nextLevel(pub.level, perf),
      history: l.history.map((h) => (h.problemId === pub.problemId ? { ...h, completedAt: now(), assisted: perf.assisted, firstCheckScore: perf.firstCheck, finalScore: perf.final, hints: Object.values(s.stats).reduce((a, x) => a + x.hintsUsed, 0), attempts: Object.values(s.stats).reduce((a, x) => a + x.attempts, 0) } : h)),
    }
    store.saveLearner(l)
    const next = recommend(store.profile() ?? profile, l)
    const report = buildReport(pub, checks, s.stats, s.timeSpentMs, next, s.stageLog, masteryProfile)
    update((x) => ({ ...x, completed: { at: now(), assisted: perf.assisted, report } }))
    setView('report')
  }

  // ----------------------------------------------------------------- reset
  const resetSheet = (name: string) => {
    if (!original) return
    const orig = original.sheets.find((s) => s.name === name)
    const s = wb.sheetByName(name)
    if (!orig || !s) return
    const sid = s.id
    const hfId = wb.hfId(sid)
    const { cells, id: _id, name: _n, ...meta } = structuredClone(orig)
    void _id
    void _n
    wb.run({
      label: `Reset ${name}`,
      dirty: [sid],
      metaScope: [sid],
      hfOps: () => wb.hf.setSheetContent(hfId, cells as never),
      metaOps: () => {
        const m = wb.meta(sid)
        for (const k of Object.keys(m)) if (k !== 'id' && k !== 'name') delete (m as unknown as Record<string, unknown>)[k]
        Object.assign(m, meta)
      },
    })
    // the mentor's examples and exercises are written again when needed
    if (name === SHEETS.example) update((x) => ({ ...x, exampleRow: 3, exampleCells: {} }))
    if (name === SHEETS.practice) update((x) => ({ ...x, practiceRow: 3, practice: {} }))
  }
  const doReset = (scope: 'sheet' | 'problem') => {
    if (scope === 'sheet') resetSheet(sheetName)
    else {
      for (const s of original?.sheets ?? []) resetSheet(s.name)
      for (const s of [...wb.sheets]) if (!original?.sheets.some((o) => o.name === s.name)) wb.removeSheet(s.id)
      update((x) => ({ ...x, stats: Object.fromEntries(Object.entries(x.stats).map(([k, v]) => [k, { ...v, lastContent: '' }])) }))
    }
    bridge.setHighlight(null)
    bridge.notify(scope === 'sheet' ? `${sheetName} was reset. Undo (Ctrl+Z) brings it back.` : 'The problem was reset to the original data. Your conversation and progress are kept.')
    const c = currentConcept(sessionRef.current)
    if (c) void prepareConcept(c).finally(() => setBusy(null))
  }

  // ----------------------------------------------------------------- render
  const pos = session.position
  const stageNo = session.completed ? 8 : stageIndex(pos.stage)
  const nC = Math.max(1, pub.newConcepts.length)
  const progress = session.completed ? 100 : Math.min(97, Math.round((100 * stageNo) / 9 + (pos.stage === 'learn' || pos.stage === 'practice' ? (11 * (pos.concept + (pos.stage === 'practice' ? 0.5 : 0))) / nC : 0)))
  const subtitle = `Stage ${STAGE_INFO[pos.stage].n}/9 · ${STAGE_INFO[pos.stage].name}${curConcept ? `: ${conceptName(curConcept)}` : ''}`
  const hintsOpen = pos.stage === 'solve' || pos.stage === 'check' || pos.stage === 'followup'
  const lastCheck = session.checks[session.checks.length - 1]

  const stageActions: ReactNode = (() => {
    if (busy || session.completed) return null
    if (pos.stage === 'practice' && practiceTarget)
      return (
        <div className="flex flex-wrap gap-1.5">
          <Button variant="default" size="sm" onClick={doCheckPractice} data-testid="check-practice"><Icon name="check" size={14} /> Check practice ({practiceTarget.cell})</Button>
          <Button variant="outline" size="sm" onClick={practiceHint} disabled={(session.hints[`practice-${curConcept}`] ?? 0) >= MAX_HINT} data-testid="practice-hint"><Icon name="help" size={14} /> Hint {Math.min(MAX_HINT, (session.hints[`practice-${curConcept}`] ?? 0) + 1)} of {MAX_HINT}</Button>
          <Button variant="ghost" size="sm" onClick={() => bridge.select(SHEETS.practice, practiceTarget.r, practiceTarget.c)}>Go to cell</Button>
        </div>
      )
    if (pos.stage === 'solve' || pos.stage === 'check')
      return (
        <div className="space-y-1.5" data-testid="task-guide">
          <ol className="space-y-0.5 text-[12px]">
            {groups.map((g) => {
              const done = g.targets.every((t) => (t.kind === 'cell' ? !!wb.getEditText(wb.sheetByName(t.sheet)!.id, t.r, t.c) : false))
              const res = lastCheck?.results.filter((r) => g.targets.some((t) => t.id === r.id))
              const ok = res?.length && res.every((r) => r.state === 'correct')
              return (
                <li key={g.id}>
                  <button type="button" className="flex w-full items-start gap-1.5 rounded px-1.5 py-0.5 text-left hover:bg-[var(--hover)]" onClick={() => { const t = g.targets[0]; if (t.kind === 'cell') bridge.select(t.sheet, t.r, t.c); else bridge.select(t.sheet, 1, 0) }}>
                    <span className="mt-0.5 shrink-0">{ok ? '✅' : done ? '☑' : '☐'}</span>
                    <span className="min-w-0 flex-1">{g.requirement} <span className="text-[var(--muted)]">· {g.where}</span></span>
                  </button>
                </li>
              )
            })}
          </ol>
        </div>
      )
    return null
  })()

  const menu = (
    <>
      <MenuItem icon="home" onSelect={onDashboard}>Mentor dashboard</MenuItem>
      <MenuItem icon="settings" onSelect={() => setSettingsOpen(true)}>Mentor settings</MenuItem>
      <MenuSep />
      <MenuItem icon="download" onSelect={() => void exportXlsx(wb).then((b) => download(b, `${safeFileName(pub.title)}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))}>Download workbook (.xlsx)</MenuItem>
      <MenuSep />
      <MenuItem icon="undo" onSelect={() => setConfirm('sheet')}>Reset current sheet…</MenuItem>
      <MenuItem icon="undo" danger onSelect={() => setConfirm('problem')}>Reset entire problem…</MenuItem>
    </>
  )

  return (
    <>
      <PanelHeader title={pub.title} subtitle={subtitle} progress={progress} onCollapse={onCollapse} menu={menu} />
      <StageStepper current={stageNo} />
      <div className="flex shrink-0 gap-1 border-b border-[var(--border)] px-2 pt-1 text-[12.5px]" role="tablist">
        {(['lesson', 'results', 'report'] as const).map((v) =>
          v === 'report' && !session.completed ? null : (
            <button key={v} role="tab" aria-selected={view === v} type="button" onClick={() => setView(v)} className={cn('-mb-px border-b-2 px-2.5 py-1.5 capitalize', view === v ? 'border-[var(--accent)] font-semibold text-[var(--accent)]' : 'border-transparent text-[var(--muted)]')}>
              {v === 'lesson' ? 'Mentor' : v === 'results' ? `Results${lastCheck ? ` ${lastCheck.correct}/${lastCheck.total}` : ''}` : 'Report'}
            </button>
          ),
        )}
      </div>
      {session.notice && (
        <div className="flex shrink-0 items-start gap-2 bg-[var(--warn-bg)] px-3 py-1.5 text-[12px] text-[var(--warn-text)]" role="status">
          <Icon name="info" size={14} className="mt-0.5 shrink-0" />
          <span className="flex-1">{session.notice}</span>
          <button type="button" aria-label="Dismiss" onClick={() => update((s) => ({ ...s, notice: undefined }))}>✕</button>
        </div>
      )}

      {view === 'lesson' && (
        <div ref={listRef} className="min-h-0 flex-1 space-y-2.5 overflow-auto px-3 py-3" data-testid="mentor-messages">
          {session.messages.map((m, i) => <Bubble key={m.id} m={m} onRetry={m.kind === 'error' && i === session.messages.length - 1 && session.pending ? retry : undefined} />)}
          {busy && (
            <div className="flex items-center gap-2 text-[12px] text-[var(--muted)]" role="status">
              <span className="h-3 w-3 animate-spin rounded-full border-2 border-[var(--accent)] border-t-transparent" /> {busy}
            </div>
          )}
          {stageActions && <div className="pt-1">{stageActions}</div>}
        </div>
      )}
      {view === 'results' && <ResultsView pub={pub} targets={targets} checks={session.checks} beginner={beginner} onJump={jumpTo} onRecheck={runCheck} />}
      {view === 'report' && session.completed && <ReportView report={session.completed.report} assisted={session.completed.assisted} onNext={onNextProblem} onDashboard={onDashboard} />}

      {!session.completed && (
        <div className="shrink-0 border-t border-[var(--border)] bg-[var(--panel)] px-3 py-2" data-testid="cell-dock">
          {hintsOpen && (selTarget ? (
            <div className="mb-2">
              <div className="text-[12px]">
                <span className="font-semibold" data-testid="selected-cell">Selected cell: {selTarget.cell}</span>
                <span className="text-[var(--muted)]"> · {selTarget.requirement} · {conceptName(selTarget.concept)}</span>
              </div>
              {dockNote?.id === selTarget.id && <div className="mt-1 rounded-md border border-amber-300 bg-[#fffbe6] px-2.5 py-1.5 text-[12.5px] text-[#3d3200]" data-testid="hint-box"><RichText text={dockNote.text} /></div>}
              <div className="mt-1.5 flex flex-wrap gap-1.5">
                <Button size="sm" variant="outline" disabled={!!busy} onClick={() => requestHint(selTarget)} data-testid="need-hint"><Icon name="help" size={14} /> Hint {Math.min(MAX_HINT, (session.hints[groupOf(selTarget.id)] ?? 0) + 1)} of {MAX_HINT}</Button>
                <Button size="sm" variant="ghost" onClick={() => checkThisCell(selTarget)}>Check this cell</Button>
              </div>
            </div>
          ) : (
            <p className="mb-2 text-[12px] text-[var(--muted)]" data-testid="dock-note">Select a yellow cell on the Task sheet to ask for a hint on it.</p>
          ))}
          <form className="flex flex-col gap-1.5" onSubmit={(e) => (e.preventDefault(), submitAnswer())}>
            <textarea
              className={cn(inputClass, 'h-16 resize-none py-1.5')}
              value={answer}
              onChange={(e) => setAnswer(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault()
                  submitAnswer()
                }
              }}
              placeholder="Reply to your mentor… (Enter to send, Shift+Enter for a new line)"
              aria-label="Your reply"
              data-testid="think-input"
            />
            <div className="flex gap-1.5">
              <Button variant="default" size="sm" type="submit" disabled={!answer.trim() || !!busy} data-testid="submit-answer">Send</Button>
              {(pos.stage === 'solve' || pos.stage === 'check' || session.checks.length > 0) && (
                <Button className="ml-auto" variant="default" size="sm" disabled={!!busy} onClick={runCheck} data-testid="check-my-work">
                  <Icon name="check2" size={15} /> {session.checks.length ? 'Check my work again' : 'Check my work'}
                </Button>
              )}
            </div>
          </form>
        </div>
      )}

      {settingsOpen && <MentorSettingsDialog store={store} onClose={() => setSettingsOpen(false)} onSaved={() => setSettingsVersion((v) => v + 1)} />}
      {confirm && (
        <Dialog
          title={confirm === 'sheet' ? `Reset “${sheetName}”?` : 'Reset the entire problem?'}
          onClose={() => setConfirm(null)}
          footer={
            <>
              <Button variant="outline" onClick={() => setConfirm(null)}>Cancel</Button>
              <Button variant="danger" onClick={() => (doReset(confirm), setConfirm(null))} data-testid="confirm-reset">Reset</Button>
            </>
          }
        >
          <p className="text-[13px]">
            {confirm === 'sheet'
              ? `This puts the ${sheetName} sheet back to how it was generated: your entries on it are removed. You can undo with Ctrl+Z.`
              : 'This restores every sheet to the original data with blank task cells. Your conversation, stage and check history are kept.'}
          </p>
        </Dialog>
      )}
    </>
  )
}

// =====================================================================================
function ResultsView({ pub, targets, checks, beginner, onJump, onRecheck }: { pub: NonNullable<MentorSession['problem']>; targets: Target[]; checks: CheckRecord[]; beginner: boolean; onJump: (t: Target) => void; onRecheck: () => void }) {
  const last = checks[checks.length - 1]
  if (!last) return <div className="flex-1 p-4 text-[13px] text-[var(--muted)]">No checks yet. Fill in the yellow cells on the Task sheet and click Check my work.</div>
  const count = (f: (r: TargetResult) => boolean) => last.results.filter(f).length
  const byConcept = new Map<ConceptId, { ok: number; total: number }>()
  for (const r of last.results) {
    const t = pub.targets.find((x) => x.id === r.id)
    if (!t) continue
    const e = byConcept.get(t.concept) ?? { ok: 0, total: 0 }
    e.total++
    if (r.state === 'correct') e.ok++
    byConcept.set(t.concept, e)
  }
  const label = (s: CellState) => (beginner ? simpleState(s) : STATE_LABEL[s])
  return (
    <div className="min-h-0 flex-1 overflow-auto px-3 py-3 text-[12.5px]" data-testid="results">
      <div className="text-[15px] font-semibold" data-testid="results-overall">{last.correct} of {last.total} correct</div>
      <div className="mt-1 flex gap-3">
        <span>✅ {count((r) => r.state === 'correct')} Correct</span>
        <span>❌ {count((r) => r.state !== 'correct' && r.state !== 'not_attempted')} {beginner ? 'Wrong' : 'Need work'}</span>
        <span>⚠️ {count((r) => r.state === 'not_attempted')} Not attempted</span>
      </div>
      {pub.mode === 'takehome' && <p className="mt-1 text-[11.5px] text-[var(--muted)]">Checks follow your mentor’s reading of the brief. If a task seems to be read differently from what the brief says, tell your mentor.</p>}
      {checks.length > 1 && <div className="mt-1 text-[11.5px] text-[var(--muted)]">{checks.map((c, i) => `#${i + 1}: ${c.correct}/${c.total}`).join(' · ')}</div>}
      <div className="mt-3 font-semibold">By concept</div>
      <div className="mt-1 flex flex-wrap gap-1.5">
        {[...byConcept].map(([c, e]) => (
          <span key={c} className={cn('rounded-full px-2 py-0.5 text-[11.5px]', e.ok === e.total ? 'bg-[var(--active)] text-[var(--accent)]' : 'bg-[var(--warn-bg)] text-[var(--warn-text)]')}>
            {conceptName(c)}: {e.ok === e.total ? (CONCEPTS[c].functions.length ? 'Strong' : 'Completed correctly') : e.ok / e.total >= 0.5 ? 'Almost there' : 'Needs practice'}
          </span>
        ))}
      </div>
      <table className="mt-3 w-full border-collapse text-left text-[12px]">
        <thead>
          <tr className="border-b border-[var(--border)] text-[11px] uppercase text-[var(--muted)]">
            <th className="py-1 pr-1">Cell</th>
            <th className="py-1 pr-1">Your entry</th>
            <th className="py-1 pr-1">Expected</th>
            <th className="py-1">Result</th>
          </tr>
        </thead>
        <tbody>
          {last.results.map((r) => {
            const t = targets.find((x) => x.id === r.id)
            if (!t) return null
            return (
              <tr key={r.id} className={cn('cursor-pointer border-b border-[var(--border)] align-top hover:bg-[var(--hover)]', r.state !== 'correct' && 'bg-[var(--warn-bg)]/40')} onClick={() => onJump(t)} title={r.message} data-testid={`result-${t.kind === 'cell' ? t.cell : t.id}`}>
                <td className="py-1 pr-1 font-mono">{t.kind === 'cell' ? t.cell : t.kind === 'sort' ? 'Sort' : t.kind === 'chart' ? 'Chart' : 'Format'}</td>
                <td className="max-w-[110px] truncate py-1 pr-1 font-mono">{r.student || 'blank'}</td>
                <td className="py-1 pr-1">{r.expected}</td>
                <td className="whitespace-nowrap py-1">{STATE_ICON(r.state)} {label(r.state)}</td>
              </tr>
            )
          })}
        </tbody>
      </table>
      <p className="mt-2 text-[11.5px] text-[var(--muted)]">Click a row to jump to that cell. Ask your mentor (or use a hint) about anything marked wrong.</p>
      <Button className="mt-3 w-full" variant="outline" onClick={onRecheck}>Check my work again</Button>
    </div>
  )
}

function ReportView({ report, assisted, onNext, onDashboard }: { report: NonNullable<MentorSession['completed']>['report']; assisted: boolean; onNext: () => void; onDashboard: () => void }) {
  const row = (k: string, v: ReactNode) => (
    <div className="flex gap-2 border-b border-[var(--border)] py-1.5">
      <span className="w-32 shrink-0 text-[var(--muted)]">{k}</span>
      <span className="min-w-0 flex-1">{v}</span>
    </div>
  )
  const section = (t: string) => <div className="mb-1 mt-4 text-[11px] font-semibold uppercase tracking-wide text-[var(--accent)]">{t}</div>
  return (
    <div className="min-h-0 flex-1 overflow-auto px-3 py-3 text-[12.5px]" data-testid="report">
      <div className="mb-2 text-[15px] font-semibold">Performance report</div>
      {row('Problem', report.title)}
      {row('Mode', report.mode === 'takehome' ? 'Take-home assessment' : 'Topic learning')}
      {row('Completion', `${report.completion}%${assisted ? ' — completed with assistance' : ''}`)}
      {row('Task cells', `${report.cells.correct}/${report.cells.total} correct after ${report.checks} check${report.checks === 1 ? '' : 's'}`)}
      {row('Hints used', `${report.hints}`)}
      {row('Independence', report.independence)}
      {row('Time', `${report.timeMinutes} min`)}
      {section('The nine stages')}
      <div className="space-y-1.5" data-testid="report-sections">
        {report.sections.map((s) => (
          <div key={s.area} className="border-b border-[var(--border)] pb-1.5">
            <div className="flex justify-between gap-2"><span className="font-semibold">{s.area}</span><span>{s.score === null ? '—' : `${s.score}%`}</span></div>
            {s.note && s.note !== 'Not assessed' && <p className="mt-0.5 text-[var(--muted)]">{s.note}</p>}
          </div>
        ))}
      </div>
      {section('Concepts')}
      {row('In this problem', report.conceptsLearned.map(conceptName).join(', '))}
      {row('Strong', report.strong.join(', ') || '—')}
      {row('Needs practice', report.needsPractice.join(', ') || '—')}
      {report.masteryProfile && (
        <>
          {section('Mastery profile')}
          <p className="rounded-md bg-[var(--active)] p-2.5 text-[13px] leading-[1.5]" data-testid="mastery-profile">{report.masteryProfile}</p>
        </>
      )}
      {report.nextSteps && (
        <div className="mt-4 rounded-md border border-[var(--accent)] p-2.5">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-[var(--accent)]">Recommended next steps</div>
          <ul className="mt-1 list-disc space-y-0.5 pl-4">
            {report.nextSteps.revise.length > 0 && <li>Revise: {report.nextSteps.revise.join(', ')}</li>}
            <li>Next problem: {report.nextSteps.problem}</li>
            <li className="capitalize">Suggested difficulty: {report.nextSteps.difficulty}</li>
          </ul>
        </div>
      )}
      <div className="mt-3 flex gap-2">
        <Button variant="default" onClick={onNext} data-testid="next-problem">Start next problem</Button>
        <Button variant="outline" onClick={onDashboard}>Dashboard</Button>
      </div>
    </div>
  )
}
