/* Minimal Google Gemini client (REST generateContent) usable directly with a key or through a proxy. */
import type { MentorSettings } from './types'

export const DEFAULT_MODEL = 'gemini-3.6-flash'
const API = 'https://generativelanguage.googleapis.com/v1beta'

declare const __MENTOR_PROXY__: string | undefined
export const BUILT_IN_PROXY: string = typeof __MENTOR_PROXY__ === 'string' ? __MENTOR_PROXY__ : ''

export class GeminiError extends Error {
  constructor(
    message: string,
    public status = 0,
    public kind: 'auth' | 'quota' | 'model' | 'network' | 'blocked' | 'format' | 'busy' | 'other' = 'other',
  ) {
    super(message)
  }
}

export interface GenerateOptions {
  system: string
  /** conversation turns; the last must be from the user */
  turns: { role: 'user' | 'model'; text: string }[]
  /** JSON schema (subset) for structured output */
  schema?: Record<string, unknown>
  signal?: AbortSignal
  timeoutMs?: number
}

export type Fetcher = typeof fetch

/** Which structured-output request shape the endpoint accepted last (remembered per page load). */
let schemaVariant = 0
/** A model used instead of the chosen one after it stayed overloaded (per page load). */
const busyFallback = new Map<string, string>()

const sleep = (ms: number, signal?: AbortSignal) =>
  new Promise<void>((resolve, reject) => {
    if (!ms) return resolve()
    const t = setTimeout(resolve, ms)
    signal?.addEventListener('abort', () => (clearTimeout(t), reject(new GeminiError('Cancelled.', 0, 'network'))))
  })

export class GeminiClient {
  constructor(
    private settings: MentorSettings,
    private fetcher: Fetcher = (...a) => fetch(...a),
    /** waits before retrying an overloaded/temporarily failing request */
    private retryDelays: number[] = [2000, 5000, 10000],
  ) {}

  get available(): boolean {
    return (this.settings.connection === 'key' && !!this.settings.apiKey.trim()) || (this.settings.connection === 'proxy' && !!this.settings.proxyUrl.trim())
  }

  /** The chosen model (from settings). */
  get chosenModel(): string {
    return this.settings.model.trim() || DEFAULT_MODEL
  }

  /** The model requests go to: the chosen one, or a stand-in while it is overloaded. */
  get model(): string {
    return busyFallback.get(this.chosenModel) ?? this.chosenModel
  }

  private endpoint(path: string): { url: string; headers: Record<string, string> } {
    if (this.settings.connection === 'proxy') {
      const base = this.settings.proxyUrl.replace(/\/+$/, '')
      return { url: `${base}/${path}`, headers: { 'Content-Type': 'application/json' } }
    }
    return { url: `${API}/${path}`, headers: { 'Content-Type': 'application/json', 'x-goog-api-key': this.settings.apiKey.trim() } }
  }

  private async post(path: string, body: unknown, opts: { signal?: AbortSignal; timeoutMs?: number }): Promise<unknown> {
    if (!this.available) throw new GeminiError('The mentor is not connected to Gemini. Add an API key in Mentor settings.', 0, 'auth')
    const { url, headers } = this.endpoint(path)
    const ctl = new AbortController()
    const timer = setTimeout(() => ctl.abort(), opts.timeoutMs ?? 60_000)
    opts.signal?.addEventListener('abort', () => ctl.abort())
    let res: Response
    try {
      res = await this.fetcher(url, { method: 'POST', headers, body: JSON.stringify(body), signal: ctl.signal })
    } catch (e) {
      throw new GeminiError(ctl.signal.aborted ? 'Gemini took too long to answer. Try again.' : `Couldn't reach Gemini (${e instanceof Error ? e.message : 'network error'}). Check your connection.`, 0, 'network')
    } finally {
      clearTimeout(timer)
    }
    const text = await res.text()
    let json: unknown = null
    try {
      json = text ? JSON.parse(text) : null
    } catch {
      /* non-JSON error page */
    }
    if (!res.ok) {
      const msg = (json as { error?: { message?: string } } | null)?.error?.message ?? (text.slice(0, 200) || res.statusText)
      const google = ` (Google says: ${msg.slice(0, 220)})`
      if (res.status === 400 && /api key/i.test(msg)) throw new GeminiError(`Gemini rejected the API key. Check you copied the whole key from aistudio.google.com/apikey.${google}`, 400, 'auth')
      if (res.status === 401 || res.status === 403)
        throw new GeminiError(
          /referer|referrer|ip address|android|ios/i.test(msg)
            ? `This API key is restricted and doesn't allow this website. In Google Cloud Console → Credentials, allow this site's address (${typeof location !== 'undefined' ? location.origin : 'this site'}) or remove the restriction.${google}`
            : /not been used|disabled|enable it/i.test(msg)
              ? `The Generative Language API isn't enabled for this key's Google Cloud project. Create the key in Google AI Studio (aistudio.google.com/apikey) or enable the API.${google}`
              : `Gemini refused the request for this key.${google}`,
          res.status,
          'auth',
        )
      if (res.status === 404) throw new GeminiError(`The model “${this.model}” isn't available for this key. Click “Load models” in Mentor settings and pick one.${google}`, 404, 'model')
      if (res.status === 429 && /high demand|overloaded|try again later/i.test(msg)) throw new GeminiError(`Gemini is overloaded right now.${google}`, 429, 'busy')
      if (res.status === 429) throw new GeminiError(`Gemini usage limit reached for this key. Wait a minute and try again.${google}`, 429, 'quota')
      if ([500, 502, 503, 504].includes(res.status)) throw new GeminiError(`Gemini (${this.model}) is overloaded or temporarily unavailable.${google}`, res.status, 'busy')
      throw new GeminiError(`Gemini error ${res.status}: ${msg}`, res.status, res.status === 400 ? 'format' : 'other')
    }
    return json
  }

  /** Plain text (or JSON text when a schema is given). Overload errors are retried with waits, then once on another model. */
  async generate(o: GenerateOptions): Promise<string> {
    let triedFallback = false
    for (let attempt = 0; ; attempt++) {
      try {
        return await this.generateOnce(o)
      } catch (e) {
        if (!(e instanceof GeminiError) || e.kind !== 'busy' || o.signal?.aborted) throw e
        if (attempt < this.retryDelays.length) {
          await sleep(this.retryDelays[attempt], o.signal)
          continue
        }
        // still overloaded: switch to the best other model this key can use, once
        if (!triedFallback && this.settings.connection === 'key') {
          triedFallback = true
          const other = pickModel((await this.listModels().catch(() => [] as string[])).filter((m) => m !== this.model))
          if (other) {
            busyFallback.set(this.chosenModel, other)
            attempt = this.retryDelays.length - 2
            continue
          }
        }
        throw new GeminiError(`Gemini is overloaded right now (high demand on ${this.model}). This is on Google's side and usually passes within a few minutes — try again shortly, or pick another model in Mentor settings.`, e.status, 'busy')
      }
    }
  }

  private async generateOnce(o: GenerateOptions): Promise<string> {
    const contents = o.turns.map((t) => ({ role: t.role, parts: [{ text: t.text }] }))
    const base: Record<string, unknown> = { systemInstruction: { parts: [{ text: o.system }] }, contents }
    const attempts: Record<string, unknown>[] = o.schema
      ? [
          { ...base, generationConfig: { responseMimeType: 'application/json', responseJsonSchema: o.schema } },
          { ...base, generationConfig: { responseMimeType: 'application/json', responseSchema: toOpenApiSchema(o.schema) } },
          { ...base, generationConfig: { responseMimeType: 'application/json' } },
        ]
      : [base]
    let lastErr: unknown
    const start = o.schema ? Math.min(schemaVariant, attempts.length - 1) : 0
    for (let i = start; i < attempts.length; i++) {
      const body = attempts[i]
      try {
        const json = (await this.post(`models/${encodeURIComponent(this.model)}:generateContent`, body, o)) as GenerateResponse
        const cand = json?.candidates?.[0]
        const text = cand?.content?.parts?.map((p) => p.text ?? '').join('') ?? ''
        if (!text) {
          const reason = json?.promptFeedback?.blockReason ?? cand?.finishReason
          throw new GeminiError(reason ? `Gemini returned no answer (${reason}).` : 'Gemini returned an empty answer.', 200, reason ? 'blocked' : 'other')
        }
        if (o.schema) schemaVariant = i
        return text
      } catch (e) {
        lastErr = e
        // only fall back to a simpler request shape when the request format was refused
        if (!(e instanceof GeminiError) || e.kind !== 'format') throw e
      }
    }
    throw lastErr
  }

  async generateJson<T>(o: GenerateOptions & { schema: Record<string, unknown> }): Promise<T> {
    const text = await this.generate(o)
    return parseJsonLoose<T>(text)
  }

  /** Models that support generateContent (for the settings picker). */
  async listModels(): Promise<string[]> {
    if (!this.available) return []
    const { url, headers } = this.endpoint('models?pageSize=200')
    const res = await this.fetcher(url, { headers })
    if (!res.ok) throw new GeminiError(`Couldn't list models (${res.status}).`, res.status, res.status === 401 || res.status === 403 ? 'auth' : 'other')
    const json = (await res.json()) as { models?: { name: string; supportedGenerationMethods?: string[] }[] }
    return (json.models ?? [])
      .filter((m) => !m.supportedGenerationMethods || m.supportedGenerationMethods.includes('generateContent'))
      .map((m) => m.name.replace(/^models\//, ''))
      .filter((n) => /gemini/i.test(n))
      .sort()
  }
}

interface GenerateResponse {
  candidates?: { content?: { parts?: { text?: string }[] }; finishReason?: string }[]
  promptFeedback?: { blockReason?: string }
}

/** The best general-purpose model in a list: newest "flash" first, skipping lite, image, audio, embedding and experimental variants. */
export function pickModel(models: string[]): string | null {
  const ok = models.filter((m) => /^gemini-/.test(m) && !/lite|image|tts|audio|live|embed|vision|thinking-exp|exp-|-exp\b|learnlm|nano|robotics|computer-use/i.test(m))
  const version = (m: string) => Number(/gemini-(\d+(?:\.\d+)?)/.exec(m)?.[1] ?? 0)
  const score = (m: string) => version(m) * 10 + (/flash/.test(m) ? 3 : /pro/.test(m) ? 1 : 0) + (/preview/.test(m) ? -2 : 0) + (/latest/.test(m) ? 0.5 : 0)
  return [...ok].sort((a, b) => score(b) - score(a))[0] ?? models[0] ?? null
}

/** Parses JSON even when wrapped in ``` fences or followed by text. */
export function parseJsonLoose<T>(text: string): T {
  let t = text.trim()
  const fence = /```(?:json)?\s*([\s\S]*?)```/i.exec(t)
  if (fence) t = fence[1].trim()
  try {
    return JSON.parse(t) as T
  } catch {
    const start = t.search(/[[{]/)
    const end = Math.max(t.lastIndexOf('}'), t.lastIndexOf(']'))
    if (start >= 0 && end > start) {
      try {
        return JSON.parse(t.slice(start, end + 1)) as T
      } catch {
        /* fallthrough */
      }
    }
    throw new GeminiError('Gemini returned data in an unexpected format.', 200, 'format')
  }
}

/** JSON Schema -> the OpenAPI-style schema older endpoints expect (type names upper case). */
function toOpenApiSchema(s: unknown): unknown {
  if (Array.isArray(s)) return s.map(toOpenApiSchema)
  if (!s || typeof s !== 'object') return s
  const out: Record<string, unknown> = {}
  for (const [k, v] of Object.entries(s as Record<string, unknown>)) {
    if (k === 'additionalProperties') continue
    if (k === 'type') {
      const t = Array.isArray(v) ? v.find((x) => x !== 'null') : v
      out.type = String(t).toUpperCase()
      if (Array.isArray(v) && v.includes('null')) out.nullable = true
    } else out[k] = toOpenApiSchema(v)
  }
  return out
}
