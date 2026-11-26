/* Mentor settings: how ExcelMentor talks to Google Gemini. */
import { useState } from 'react'
import { DEFAULT_GROQ_MODEL, DEFAULT_MODEL, GeminiClient, GeminiError, pickGroqModel, pickModel } from '../../mentor/gemini'
import type { MentorStore } from '../../mentor/store'
import type { MentorSettings } from '../../mentor/types'
import { Button, Dialog, Field, inputClass } from '../components/ui'
import { cn } from '../lib/utils'

export function MentorSettingsDialog({ store, onClose, onSaved }: { store: MentorStore; onClose: () => void; onSaved?: (s: MentorSettings) => void }) {
  const [s, setS] = useState<MentorSettings>(store.settings())
  const [models, setModels] = useState<string[]>([])
  const [status, setStatus] = useState<{ ok: boolean; text: string } | null>(null)
  const [busy, setBusy] = useState(false)
  const groq = s.connection === 'groq'
  const modelValue = groq ? (s.groqModel ?? '') : s.model
  const withModel = (x: MentorSettings, m: string): MentorSettings => (x.connection === 'groq' ? { ...x, groqModel: m } : { ...x, model: m })
  const pick = (x: MentorSettings, list: string[]) => (x.connection === 'groq' ? pickGroqModel(list) : pickModel(list))
  const up = (p: Partial<MentorSettings>) => (setS((x) => ({ ...x, ...p })), setStatus(null))
  const save = (x: MentorSettings) => {
    store.saveSettings(x)
    onSaved?.(x)
  }
  const ping = async (x: MentorSettings) => {
    const c = new GeminiClient(x)
    const reply = await c.generate({ system: 'You are a connection test.', turns: [{ role: 'user', text: 'Reply with the single word OK.' }], timeoutMs: 30_000 })
    return { reply, used: c.model }
  }
  /** Tests the connection; if the model isn't available, switches to the best model the key can use. Saves on success. */
  const test = async () => {
    const cur = { ...s, apiKey: s.apiKey.trim(), proxyUrl: s.proxyUrl.trim(), groqKey: (s.groqKey ?? '').trim() }
    if (cur.connection === 'groq' && !cur.groqKey) return setStatus({ ok: false, text: 'Paste your Groq API key first (from console.groq.com/keys).' })
    if (cur.connection === 'key' && !cur.apiKey) return setStatus({ ok: false, text: 'Paste your Gemini API key first (from aistudio.google.com/apikey).' })
    if (cur.connection === 'proxy' && !cur.proxyUrl) return setStatus({ ok: false, text: 'Enter the proxy URL first.' })
    setBusy(true)
    setStatus(null)
    try {
      let x: MentorSettings = cur
      let reply: string
      let used = cur.connection === 'groq' ? (cur.groqModel ?? '') : cur.model
      try {
        ;({ reply, used } = await ping(x))
      } catch (e) {
        if (!(e instanceof GeminiError) || e.kind !== 'model') throw e
        const list = await new GeminiClient(x).listModels()
        setModels(list)
        const m = pick(x, list)
        if (!m) throw e
        x = withModel(x, m)
        ;({ reply, used } = await ping(x))
      }
      setS(x)
      save(x)
      const xm = x.connection === 'groq' ? (x.groqModel ?? '') : x.model
      const cm = cur.connection === 'groq' ? (cur.groqModel ?? '') : cur.model
      setStatus({ ok: true, text: `Connected and saved. Model: ${xm}${xm !== cm ? ` (“${cm}” isn't available for this key, so I switched)` : ''}.${used !== xm ? ` “${xm}” is overloaded at the moment, so the mentor will use “${used}” until it recovers.` : ''} Gemini replied: “${reply.trim().slice(0, 40)}”.` })
    } catch (e) {
      setStatus({ ok: false, text: e instanceof GeminiError || e instanceof Error ? e.message : String(e) })
    } finally {
      setBusy(false)
    }
  }
  const loadModels = async () => {
    setBusy(true)
    try {
      const list = await new GeminiClient(s).listModels()
      setModels(list)
      const best = list.includes(modelValue) ? null : pick(s, list)
      if (best) setS((x) => withModel(x, best))
      setStatus({ ok: true, text: `${list.length} models available for this key.${best ? ` “${modelValue}” isn't one of them, so I picked ${best}.` : ''}` })
    } catch (e) {
      setStatus({ ok: false, text: e instanceof Error ? e.message : String(e) })
    } finally {
      setBusy(false)
    }
  }
  return (
    <Dialog
      title="Mentor settings"
      onClose={onClose}
      width={540}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button
            variant="default"
            onClick={() => {
              save({ ...s, apiKey: s.apiKey.trim(), proxyUrl: s.proxyUrl.trim(), groqKey: (s.groqKey ?? '').trim() })
              onClose()
            }}
          >
            Save
          </Button>
        </>
      }
    >
      <p className="mb-3 text-[12.5px] text-[var(--muted)]">ExcelMentor needs an AI service (Google Gemini or Groq): it builds every problem and runs the whole lesson. Paste your key, then click Test connection — it saves automatically when it works.</p>
      <div className="mb-3 space-y-1.5 text-[13px]" role="radiogroup" aria-label="Connection">
        {(
          [
            ['key', 'My Gemini API key', 'The key is stored only in this browser and sent directly to Google.'],
            ['groq', 'My Groq API key', 'Free and fast alternative. The key is stored only in this browser and sent directly to Groq.'],
            ['proxy', 'Proxy server', 'Your organisation’s proxy keeps the key on a server (see proxy/README in the source).'],
          ] as const
        ).map(([v, label, hint]) => (
          <label key={v} className={cn('flex cursor-pointer gap-2 rounded-md border p-2', s.connection === v ? 'border-[var(--accent)] bg-[var(--active)]' : 'border-[var(--border)]')}>
            <input type="radio" name="conn" className="mt-0.5 accent-[var(--accent)]" checked={s.connection === v} onChange={() => up({ connection: v })} />
            <span>
              <span className="block font-medium">{label}</span>
              <span className="block text-[11.5px] text-[var(--muted)]">{hint}</span>
            </span>
          </label>
        ))}
      </div>
      {s.connection === 'key' && (
        <Field label="Gemini API key" hint="Create one at aistudio.google.com/apikey. Anyone using this browser profile can read it, so don't use a shared computer.">
          <input className={inputClass} type="password" autoComplete="off" value={s.apiKey} onChange={(e) => up({ apiKey: e.target.value })} aria-label="Gemini API key" placeholder="AIza…" />
        </Field>
      )}
      {groq && (
        <Field label="Groq API key" hint="Create one at console.groq.com/keys. Anyone using this browser profile can read it, so don't use a shared computer.">
          <input className={inputClass} type="password" autoComplete="off" value={s.groqKey ?? ''} onChange={(e) => up({ groqKey: e.target.value })} aria-label="Groq API key" placeholder="gsk_…" />
        </Field>
      )}
      {s.connection === 'proxy' && (
        <Field label="Proxy URL" hint="Requests go to <proxy>/models/<model>:generateContent">
          <input className={inputClass} value={s.proxyUrl} onChange={(e) => up({ proxyUrl: e.target.value })} aria-label="Proxy URL" placeholder="https://mentor-proxy.example.workers.dev" />
        </Field>
      )}
      {(

        <>
          <Field label="Model" hint={`Default: ${groq ? DEFAULT_GROQ_MODEL : DEFAULT_MODEL}. Use “Load models” to see what your key can use.`}>
            <div className="flex gap-2">
              <input className={cn(inputClass, 'flex-1 font-mono')} list="gemini-models" value={modelValue} onChange={(e) => up(groq ? { groqModel: e.target.value } : { model: e.target.value })} aria-label="Model" />
              <datalist id="gemini-models">{models.map((m) => <option key={m} value={m} />)}</datalist>
              <Button variant="outline" size="sm" disabled={busy} onClick={() => void loadModels()}>Load models</Button>
            </div>
          </Field>
          <Button variant="outline" size="sm" disabled={busy} onClick={() => void test()}>{busy ? 'Working…' : 'Test connection'}</Button>
        </>
      )}
      {status && (
        <p className={cn('mt-3 rounded-md px-2 py-1.5 text-[12.5px]', status.ok ? 'bg-[var(--active)] text-[var(--accent)]' : 'bg-[var(--warn-bg)] text-[var(--warn-text)]')} role="status">
          {status.text}
        </p>
      )}
    </Dialog>
  )
}
