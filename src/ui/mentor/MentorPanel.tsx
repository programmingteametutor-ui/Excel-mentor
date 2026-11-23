/* The AI Mentor panel shown beside the workbook (problem sessions and the Excel orientation). */
import { useEffect, useMemo, useState } from 'react'
import type { EditorBridge } from '../Editor'
import { markIntroduced } from '../../mentor/progress'
import type { MentorStore } from '../../mentor/store'
import type { MentorSession } from '../../mentor/types'
import { ORIENTATION, ORIENTATION_CONCEPTS } from '../../mentor/orientation'
import { Button, inputClass, MenuItem } from '../components/ui'
import { Icon } from '../icons'
import { cn } from '../lib/utils'
import { now, PanelHeader, RichText } from './parts'
import { ProblemPanel } from './ProblemPanel'

export { RichText }

export function MentorSide({ bridge, store, onDashboard, onNextProblem }: { bridge: EditorBridge; store: MentorStore; onDashboard: () => void; onNextProblem: () => void }) {
  const [collapsed, setCollapsed] = useState(() => {
    try {
      return localStorage.getItem('excelmentor.panelCollapsed') === '1' || window.innerWidth < 720
    } catch {
      return false
    }
  })
  const toggle = () =>
    setCollapsed((c) => {
      try {
        localStorage.setItem('excelmentor.panelCollapsed', c ? '0' : '1')
      } catch {
        /* ignore */
      }
      return !c
    })
  // the mentor takes 40% of the screen by default (workbook 60%); drag the left edge to change it
  const [widthPct, setWidthPct] = useState(() => {
    try {
      const v = Number(localStorage.getItem('excelmentor.panelWidth'))
      return v >= 25 && v <= 65 ? v : 40
    } catch {
      return 40
    }
  })
  const setWidth = (v: number) => {
    const w = Math.max(25, Math.min(65, v))
    setWidthPct(w)
    try {
      localStorage.setItem('excelmentor.panelWidth', String(w))
    } catch {
      /* ignore */
    }
  }
  const session = useMemo(() => store.session(bridge.wb.id), [store, bridge.wb.id])
  if (!session) return null
  if (collapsed)
    return (
      <aside className="flex w-10 shrink-0 flex-col items-center gap-2 border-l border-[var(--border)] bg-[var(--panel)] pt-2">
        <button type="button" className="flex h-8 w-8 items-center justify-center rounded-md bg-[var(--accent)] text-[var(--accent-contrast)]" onClick={toggle} aria-label="Open mentor panel" title="Open ExcelMentor">
          <Icon name="wand" size={16} />
        </button>
      </aside>
    )
  return (
    <aside className="relative flex h-full shrink-0 flex-col border-l border-[var(--border)] bg-[var(--bg)]" style={{ width: `${widthPct}%`, minWidth: 320 }} aria-label="AI Mentor" data-testid="mentor-panel">
      <div
        role="separator"
        aria-orientation="vertical"
        aria-label="Resize mentor panel"
        title="Drag to resize · double-click for 40%"
        className="absolute -left-1 top-0 z-10 h-full w-2 cursor-col-resize hover:bg-[var(--accent)]/30"
        onDoubleClick={() => setWidth(40)}
        onPointerDown={(e) => {
          e.preventDefault()
          const el = e.currentTarget
          el.setPointerCapture(e.pointerId)
          const move = (ev: PointerEvent) => setWidth(Math.round((100 * (window.innerWidth - ev.clientX)) / window.innerWidth))
          const up = () => {
            el.removeEventListener('pointermove', move)
            el.removeEventListener('pointerup', up)
          }
          el.addEventListener('pointermove', move)
          el.addEventListener('pointerup', up)
        }}
      />
      {session.kind === 'orientation' ? (
        <OrientationPanel bridge={bridge} store={store} initial={session} onCollapse={toggle} onDashboard={onDashboard} onNextProblem={onNextProblem} />
      ) : session.version !== 2 ? (
        <LegacyPanel onCollapse={toggle} onDashboard={onDashboard} onNextProblem={onNextProblem} />
      ) : (
        <ProblemPanel bridge={bridge} store={store} initial={session} onCollapse={toggle} onDashboard={onDashboard} onNextProblem={onNextProblem} />
      )}
    </aside>
  )
}

// =====================================================================================
function OrientationPanel({ bridge, store, initial, onCollapse, onDashboard, onNextProblem }: { bridge: EditorBridge; store: MentorStore; initial: MentorSession; onCollapse: () => void; onDashboard: () => void; onNextProblem: () => void }) {
  const [session, setSession] = useState(initial)
  const [answer, setAnswer] = useState('')
  const [feedback, setFeedback] = useState<string | null>(null)
  const [done, setDone] = useState(false)
  const i = Math.min(session.orientationStep ?? 0, ORIENTATION.length)
  const finished = i >= ORIENTATION.length
  const step = ORIENTATION[i]
  useEffect(() => {
    store.saveSession(session)
  }, [session, store])
  useEffect(() => {
    setDone(false)
    setFeedback(null)
    setAnswer('')
  }, [i])
  const sheetName = bridge.wb.meta(bridge.sheetId).name
  // automatic checks on every selection/edit
  useEffect(() => {
    if (!step?.check || done) return
    try {
      if (step.check({ wb: bridge.wb, sheetName, selection: bridge.selection })) {
        setDone(true)
        setFeedback('✅ Well done!')
      }
    } catch {
      /* not yet */
    }
  }, [bridge.version, bridge.selection, sheetName, step, done, bridge.wb])
  const next = () => {
    const n = i + 1
    setSession((s) => ({ ...s, orientationStep: n, lastActivityAt: now() }))
    if (n >= ORIENTATION.length) {
      const p = store.profile()
      if (p) store.saveProfile({ ...p, orientationDone: true })
      let l = store.learner()
      for (const c of ORIENTATION_CONCEPTS) l = markIntroduced(l, c)
      store.saveLearner(l)
    }
  }
  const submit = () => {
    if (!step?.question) return
    const ok = step.question.answers.includes(answer.trim().toLowerCase().replace(/\s+/g, ''))
    setDone(ok)
    setFeedback(ok ? '✅ Correct!' : 'Not quite — have another look and try again.')
  }
  return (
    <>
      <PanelHeader title="Excel Orientation" progress={Math.round((100 * i) / ORIENTATION.length)} onCollapse={onCollapse} menu={<MenuItem icon="home" onSelect={onDashboard}>Mentor dashboard</MenuItem>} />
      <div className="min-h-0 flex-1 overflow-auto px-3 py-3 text-[13px]" data-testid="orientation">
        {finished ? (
          <div className="space-y-3">
            <div className="text-[15px] font-semibold">🎉 You know the basics!</div>
            <p>You can now find your way around a workbook, enter and edit data, copy, format and write simple formulas. Time to use them on a real business problem.</p>
            <Button variant="default" onClick={onNextProblem} data-testid="first-problem">Start my first problem</Button>
          </div>
        ) : (
          <>
            <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-[var(--muted)]">Step {i + 1} of {ORIENTATION.length}</div>
            <div className="mb-2 text-[15px] font-semibold" data-testid="orientation-title">{step.title}</div>
            <p className="mb-3 leading-[1.5]">{step.explain}</p>
            <div className="rounded-md border border-[var(--accent)] bg-[var(--active)] p-2.5">
              <div className="text-[11px] font-semibold uppercase tracking-wide text-[var(--accent)]">Your turn</div>
              <div>{step.task}</div>
            </div>
            {step.question && (
              <form className="mt-2 flex gap-1.5" onSubmit={(e) => (e.preventDefault(), submit())}>
                <input className={cn(inputClass, 'flex-1')} value={answer} onChange={(e) => setAnswer(e.target.value)} placeholder={step.question.placeholder} aria-label="Your answer" />
                <Button size="sm" variant="outline" type="submit">Check</Button>
              </form>
            )}
            {feedback && <p className="mt-2 text-[13px]" role="status" data-testid="orientation-feedback">{feedback}</p>}
            {!step.check && !step.question && <p className="mt-2 text-[12px] text-[var(--muted)]">Read this, then continue.</p>}
            <div className="mt-3 flex gap-1.5">
              <Button variant="default" size="sm" disabled={!done && !!(step.check || step.question)} onClick={next} data-testid="orientation-next">Next <Icon name="chevronRight" size={14} /></Button>
              {(step.check || step.question) && !done && <Button variant="ghost" size="sm" onClick={next}>Skip</Button>}
            </div>
          </>
        )}
      </div>
    </>
  )
}


/** Sessions saved by the earlier, scripted mentor can't continue in the Gemini-driven flow. */
function LegacyPanel({ onCollapse, onDashboard, onNextProblem }: { onCollapse: () => void; onDashboard: () => void; onNextProblem: () => void }) {
  return (
    <>
      <PanelHeader title="Earlier session" progress={0} onCollapse={onCollapse} menu={<MenuItem icon="home" onSelect={onDashboard}>Mentor dashboard</MenuItem>} />
      <div className="space-y-3 px-3 py-3 text-[13px]" data-testid="legacy-session">
        <p>This problem was started with an earlier version of ExcelMentor, so the mentor can't continue it. Your workbook is still here and you can keep working in it.</p>
        <Button variant="default" onClick={onNextProblem}>Start a new session</Button>
      </div>
    </>
  )
}
