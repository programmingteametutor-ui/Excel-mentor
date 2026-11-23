import { useCallback, useEffect, useMemo, useState } from 'react'
import { Workbook } from '../engine/workbook'
import type { DateOrder } from '../engine/inputParser'
import type { ImportReport } from '../files/xlsx'
import { Autosave } from '../persistence/autosave'
import { IndexedDbStore, MemoryStore, type WorkbookStore } from '../persistence/store'
import type { ViewState } from './commands'
import { Editor } from './Editor'
import { MentorHome } from './mentor/MentorHome'
import { MentorSide } from './mentor/MentorPanel'
import { MentorStore } from '../mentor/store'
import { Logo } from './icons'
import { storageGet, storageSet } from './lib/utils'

let sessionNonce = 0

interface Session {
  nonce: number
  wb: Workbook
  autosave: Autosave
  report?: { report: ImportReport; fileName: string } | null
}

const DEFAULT_VIEW: ViewState = { zoom: 1, showGrid: true, showHeadings: true, showFormulaBar: true, showFormulas: false, dark: false }

function readView(): ViewState {
  try {
    const v = JSON.parse(storageGet('excelpro.view') ?? '{}')
    const hostTheme = document.documentElement.dataset.theme
    const dark =
      typeof v.dark === 'boolean'
        ? v.dark
        : hostTheme === 'dark' || (hostTheme !== 'light' && typeof matchMedia !== 'undefined' && matchMedia('(prefers-color-scheme: dark)').matches)
    return { ...DEFAULT_VIEW, ...v, showFormulas: false, dark }
  } catch {
    return DEFAULT_VIEW
  }
}

const route = () => {
  const m = /^#\/wb\/([^/?]+)/.exec(location.hash)
  return m ? decodeURIComponent(m[1]) : null
}

export function App() {
  const [store, setStore] = useState<WorkbookStore | null>(null)
  const [view, setViewState] = useState<ViewState>(readView)
  const [dateOrder, setDateOrderState] = useState<DateOrder>(() => (storageGet('excelpro.dateOrder') === 'DMY' ? 'DMY' : 'MDY'))
  const [wbId, setWbId] = useState<string | null>(route)
  const [session, setSession] = useState<Session | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const mentorStore = useMemo(() => new MentorStore(), [])

  const setView = useCallback((v: Partial<ViewState>) => {
    setViewState((cur) => {
      const next = { ...cur, ...v }
      storageSet('excelpro.view', JSON.stringify({ ...next, showFormulas: undefined }))
      return next
    })
  }, [])
  const setDateOrder = (d: DateOrder) => {
    setDateOrderState(d)
    storageSet('excelpro.dateOrder', d)
  }

  useEffect(() => {
    document.documentElement.classList.toggle('dark', view.dark)
  }, [view.dark])

  // storage: IndexedDB, falling back to memory when the browser blocks it
  useEffect(() => {
    let cancelled = false
    ;(async () => {
      try {
        if (typeof indexedDB === 'undefined') throw new Error('IndexedDB unavailable')
        const s = new IndexedDbStore()
        await s.list()
        if (!cancelled) setStore(s)
      } catch {
        if (!cancelled) {
          setStore(new MemoryStore())
          setError('This browser is blocking local storage (private mode or site settings), so your sessions only last until you close the tab. Download .xlsx copies to keep your work.')
        }
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    const h = () => {
      setWbId(route())
    }
    window.addEventListener('hashchange', h)
    return () => window.removeEventListener('hashchange', h)
  }, [])

  const go = (id: string | null) => {
    location.hash = id ? `#/wb/${encodeURIComponent(id)}` : '#/mentor'
  }

  // open the workbook named in the URL
  useEffect(() => {
    if (!store) return
    if (!wbId) {
      if (session) {
        void session.autosave.flush().finally(() => session.autosave.dispose())
        setSession(null)
      }
      return
    }
    if (session && session.wb.id === wbId) return
    let cancelled = false
    setBusy('Opening workbook…')
    store
      .load(wbId)
      .then((res) => {
        if (cancelled) return
        if (!res) {
          setError('That workbook was not found in this browser.')
          go(null)
          return
        }
        session?.autosave.dispose()
        const wb = Workbook.fromDoc(res.doc)
        setSession({ nonce: ++sessionNonce, wb, autosave: new Autosave(wb, store, { revision: res.revision, persisted: true }) })
      })
      .catch((e) => setError(`Couldn't open the workbook: ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => !cancelled && setBusy(null))
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [store, wbId])

  const startSession = (wb: Workbook, report?: Session['report']) => {
    if (!store) return
    session?.autosave.dispose()
    const autosave = new Autosave(wb, store, { revision: 0, persisted: false })
    setSession({ nonce: ++sessionNonce, wb, autosave, report })
    go(wb.id)
  }

  const reloadFromStore = async () => {
    if (!store || !session) return
    const res = await store.load(session.wb.id)
    if (!res) return
    session.autosave.dispose()
    const wb = Workbook.fromDoc(res.doc)
    setSession({ nonce: ++sessionNonce, wb, autosave: new Autosave(wb, store, { revision: res.revision, persisted: true }) })
  }

  const overlay = useMemo(
    () =>
      busy && (
        <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/20">
          <div className="flex items-center gap-3 rounded-xl bg-[var(--bg)] px-5 py-4 text-[13px] shadow-2xl">
            <span className="h-4 w-4 animate-spin rounded-full border-2 border-[var(--accent)] border-t-transparent" />
            {busy}
          </div>
        </div>
      ),
    [busy],
  )

  const mentorOpenDoc = (doc: Parameters<typeof Workbook.fromDoc>[0]) => startSession(Workbook.fromDoc(doc))
  const workbookExists = useCallback(async (id: string) => !!(await store?.list(true))?.some((h) => h.id === id), [store])

  if (!store) {
    return (
      <div className="flex h-full items-center justify-center gap-2 text-[var(--muted)]">
        <Logo /> Loading ExcelMentor…
      </div>
    )
  }

  return (
    <>
      {!wbId ? (
        <MentorHome store={mentorStore} onOpenDoc={mentorOpenDoc} onOpenWorkbook={(id) => go(id)} workbookExists={workbookExists} />
      ) : session && wbId === session.wb.id ? (
        <Editor
          key={session.nonce}
          wb={session.wb}
          autosave={session.autosave}
          view={view}
          setView={setView}
          dateOrder={dateOrder}
          setDateOrder={setDateOrder}
          onHome={() => go(null)}
          onImportFile={() => setError('Opening other files isn’t available inside a mentor session. To work on a take-home assessment, upload it from the dashboard.')}
          initialReport={session.report}
          onReloadFromStore={() => void reloadFromStore()}
          renderSide={mentorStore.session(session.wb.id) ? (bridge) => <MentorSide bridge={bridge} store={mentorStore} onDashboard={() => (location.hash = '#/mentor')} onNextProblem={() => (location.hash = '#/mentor/start')} /> : undefined}
        />
      ) : null}
      {overlay}
      {error && (
        <div role="alert" className="fixed bottom-6 left-1/2 z-[150] flex max-w-[92vw] -translate-x-1/2 items-start gap-3 rounded-lg border border-[var(--border)] bg-[var(--bg)] px-4 py-3 text-[13px] shadow-2xl">
          <span className="text-[var(--danger)]">⚠</span>
          <span className="max-w-xl">{error}</span>
          <button type="button" className="ml-2 text-[var(--muted)] hover:text-[var(--text)]" onClick={() => setError(null)} aria-label="Dismiss">
            ✕
          </button>
        </div>
      )}
    </>
  )
}
