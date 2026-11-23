/* Small shared pieces of the mentor panel. */
import type { ReactNode } from 'react'
import type { CellState, ChatMessage, TargetStats } from '../../mentor/types'
import { STAGE_INFO, STAGE_ORDER } from '../../mentor/protocol'
import { Button, Popover } from '../components/ui'
import { Icon } from '../icons'
import { cn } from '../lib/utils'

export const statOf = (s: TargetStats | undefined): TargetStats => ({ hintLevel: 0, hintsUsed: 0, attempts: 0, ...s })
export const uid = () => Math.random().toString(36).slice(2, 10)
export const now = () => new Date().toISOString()

/** Inline formatting: `code`, **bold** and *italic*. */
export function RichText({ text }: { text: string }) {
  const parts = text.split(/(`[^`]+`|\*\*[^*]+\*\*|\*[^*\s][^*]*\*)/g)
  return (
    <>
      {parts.map((p, i) =>
        p.startsWith('`') && p.endsWith('`') && p.length > 1 ? (
          <code key={i} className="rounded bg-[var(--panel)] px-1 font-mono text-[12.5px]">{p.slice(1, -1)}</code>
        ) : p.startsWith('**') && p.endsWith('**') && p.length > 4 ? (
          <b key={i}>{p.slice(2, -2)}</b>
        ) : p.startsWith('*') && p.endsWith('*') && p.length > 2 ? (
          <i key={i}>{p.slice(1, -1)}</i>
        ) : (
          <span key={i}>{p}</span>
        ),
      )}
    </>
  )
}

/** The mentor's Markdown subset: ### headings, --> / - bullets, 1. lists, > quotes, ``` code, blank-line paragraphs. */
export function Markdown({ text }: { text: string }) {
  const lines = text.replace(/\r/g, '').split('\n')
  const out: ReactNode[] = []
  let i = 0
  let key = 0
  const BULLET = /^\s*(-->|[-*•])\s+/
  const NUMBER = /^\s*(\d+)[.)]\s+/
  while (i < lines.length) {
    const line = lines[i]
    if (!line.trim()) {
      i++
      continue
    }
    if (line.trim().startsWith('```')) {
      const body: string[] = []
      i++
      while (i < lines.length && !lines[i].trim().startsWith('```')) body.push(lines[i++])
      i++
      out.push(<pre key={key++} className="my-1 overflow-auto rounded-md bg-[var(--panel)] px-2.5 py-2 font-mono text-[12.5px] leading-[1.45]">{body.join('\n')}</pre>)
      continue
    }
    const h = /^(#{1,6})\s+(.*)$/.exec(line)
    if (h) {
      out.push(<div key={key++} className={cn('mt-1 font-semibold', h[1].length <= 3 ? 'text-[15px]' : 'text-[14px]')}><RichText text={h[2]} /></div>)
      i++
      continue
    }
    if (/^\s*>/.test(line)) {
      const body: string[] = []
      while (i < lines.length && /^\s*>/.test(lines[i])) body.push(lines[i++].replace(/^\s*>\s?/, ''))
      out.push(<blockquote key={key++} className="border-l-[3px] border-[var(--accent)] bg-[var(--active)] px-2.5 py-1.5 text-[13.5px]">{body.map((b, j) => <div key={j}><RichText text={b} /></div>)}</blockquote>)
      continue
    }
    if (BULLET.test(line) || NUMBER.test(line)) {
      const ordered = NUMBER.test(line)
      const items: string[] = []
      while (i < lines.length && (ordered ? NUMBER.test(lines[i]) : BULLET.test(lines[i]))) {
        let item = lines[i++].replace(ordered ? NUMBER : BULLET, '')
        // continuation lines (indented, not a new item or block)
        while (i < lines.length && lines[i].trim() && /^\s{2,}/.test(lines[i]) && !BULLET.test(lines[i]) && !NUMBER.test(lines[i])) item += ' ' + lines[i++].trim()
        items.push(item)
        while (i < lines.length && !lines[i].trim() && i + 1 < lines.length && (ordered ? NUMBER.test(lines[i + 1]) : BULLET.test(lines[i + 1]))) i++
      }
      const Tag = ordered ? 'ol' : 'ul'
      out.push(
        <Tag key={key++} className={cn('space-y-1', ordered ? 'list-decimal pl-5' : 'list-none pl-1')}>
          {items.map((it, j) =>
            ordered ? (
              <li key={j}><RichText text={it} /></li>
            ) : (
              <li key={j} className="flex gap-1.5"><span aria-hidden className="shrink-0 text-[var(--accent)]">→</span><span className="min-w-0"><RichText text={it} /></span></li>
            ),
          )}
        </Tag>,
      )
      continue
    }
    const body: string[] = []
    while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|\s*>|\s*```)/.test(lines[i]) && !BULLET.test(lines[i]) && !NUMBER.test(lines[i])) body.push(lines[i++])
    out.push(<p key={key++}>{body.map((b, j) => <span key={j}>{j > 0 && <br />}<RichText text={b} /></span>)}</p>)
  }
  return <div className="space-y-2.5">{out}</div>
}

export function Bubble({ m, onRetry }: { m: ChatMessage; onRetry?: () => void }) {
  if (m.kind === 'event')
    return (
      <div className="flex justify-center" data-testid="event-message">
        <div className="max-w-[94%] rounded-lg border border-dashed border-[var(--border-strong)] px-3 py-1.5 text-[12.5px] text-[var(--muted)]"><RichText text={m.text} /></div>
      </div>
    )
  if (m.kind === 'error')
    return (
      <div className="flex justify-start" role="alert" data-testid="error-message">
        <div className="max-w-[94%] rounded-xl bg-[var(--warn-bg)] px-3.5 py-2.5 text-[13px] text-[var(--warn-text)]">
          <RichText text={m.text} />
          {onRetry && <div className="mt-1.5"><Button size="sm" variant="outline" onClick={onRetry} data-testid="retry">Retry</Button></div>}
        </div>
      </div>
    )
  const mentor = m.from === 'mentor'
  return (
    <div className={cn('flex', mentor ? 'justify-start' : 'justify-end')}>
      <div className={cn('max-w-[94%] rounded-xl px-3.5 py-2.5 text-[14px] leading-[1.55]', mentor ? 'rounded-tl-sm bg-[var(--panel-2)] text-[var(--text)]' : 'whitespace-pre-wrap rounded-tr-sm bg-[var(--accent)] text-[var(--accent-contrast)]')} data-testid={mentor ? 'mentor-message' : 'student-message'}>
        {mentor ? <Markdown text={m.text} /> : m.text}
      </div>
    </div>
  )
}

export const STATE_LABEL: Record<CellState, string> = {
  not_required: 'Not required',
  not_attempted: 'Not attempted',
  incorrect: 'Incorrect',
  partially_correct: 'Partially correct',
  correct_value: 'Correct value',
  correct_formula: 'Correct formula',
  correct: 'Correct',
}
export const STATE_ICON = (s: CellState) => (s === 'correct' ? '✅' : s === 'not_attempted' ? '⚠️' : s === 'correct_value' || s === 'correct_formula' || s === 'partially_correct' ? '🟡' : '❌')

export function PanelHeader({ title, subtitle, progress, onCollapse, menu }: { title: string; subtitle?: string; progress: number; onCollapse: () => void; menu: ReactNode }) {
  return (
    <div className="shrink-0 border-b border-[var(--border)] bg-[var(--panel)] px-3 pb-2 pt-2">
      <div className="flex items-center gap-2">
        <span className="flex h-6 w-6 items-center justify-center rounded-md bg-[var(--accent)] text-[var(--accent-contrast)]">
          <Icon name="wand" size={14} />
        </span>
        <div className="min-w-0 flex-1">
          <div className="truncate text-[11px] font-semibold uppercase tracking-wide text-[var(--accent)]">ExcelMentor{subtitle ? ` · ${subtitle}` : ''}</div>
          <div className="truncate text-[13px] font-semibold" title={title} data-testid="mentor-title">{title}</div>
        </div>
        <Popover align="end" trigger={({ toggle }) => <Button size="icon" aria-label="Mentor menu" onClick={toggle}><Icon name="more" size={16} /></Button>}>
          {menu}
        </Popover>
        <Button size="icon" aria-label="Collapse mentor panel" onClick={onCollapse}><Icon name="chevronRight" size={16} /></Button>
      </div>
      <div className="mt-2 h-1.5 overflow-hidden rounded-full bg-[var(--panel-2)]" role="progressbar" aria-valuenow={progress} aria-valuemin={0} aria-valuemax={100} aria-label="Progress">
        <div className="h-full rounded-full bg-[var(--accent)] transition-all" style={{ width: `${progress}%` }} />
      </div>
    </div>
  )
}

/** The nine stages as a compact stepper. */
export function StageStepper({ current }: { current: number }) {
  return (
    <ol className="flex shrink-0 flex-wrap gap-0.5 border-b border-[var(--border)] px-2 py-1.5 text-[11px]" aria-label="Learning stages" data-testid="stage-stepper">
      {STAGE_ORDER.map((s, i) => (
        <li
          key={s}
          aria-current={i === current ? 'step' : undefined}
          className={cn(
            'shrink-0 rounded-full px-1.5 py-0.5',
            i < current && 'text-[var(--accent)]',
            i === current && 'bg-[var(--accent)] font-semibold text-[var(--accent-contrast)]',
            i > current && 'text-[var(--muted)]',
          )}
        >
          {i < current ? '✓ ' : `${i + 1}. `}
          {STAGE_INFO[s].name}
        </li>
      ))}
    </ol>
  )
}
