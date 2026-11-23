/* Mentor persistence in browser storage. Answers are stored apart from the workbook and obfuscated. */
import type { WorkbookDoc } from '../model/types'
import { BUILT_IN_PROXY, DEFAULT_MODEL } from './gemini'
import type { LearnerState, MentorSession, MentorSettings, Profile, Target } from './types'

export interface KV {
  get(k: string): string | null
  set(k: string, v: string): void
  remove(k: string): void
}

export const browserKV: KV = {
  get: (k) => {
    try {
      return localStorage.getItem(k)
    } catch {
      return null
    }
  },
  set: (k, v) => {
    try {
      localStorage.setItem(k, v)
    } catch {
      /* storage full or blocked */
    }
  },
  remove: (k) => {
    try {
      localStorage.removeItem(k)
    } catch {
      /* ignore */
    }
  },
}

export function memoryKV(): KV {
  const m = new Map<string, string>()
  return { get: (k) => m.get(k) ?? null, set: (k, v) => void m.set(k, v), remove: (k) => void m.delete(k) }
}

const P = 'excelmentor.'

function read<T>(kv: KV, key: string): T | null {
  const s = kv.get(P + key)
  if (!s) return null
  try {
    return JSON.parse(s) as T
  } catch {
    return null
  }
}

// XOR + base64 so answers aren't readable at a glance in dev tools (not a security boundary).
function scramble(text: string, key: string): string {
  const k = [...key].map((c) => c.charCodeAt(0))
  const bytes = new TextEncoder().encode(text).map((b, i) => b ^ k[i % k.length])
  let bin = ''
  bytes.forEach((b) => (bin += String.fromCharCode(b)))
  return btoa(bin)
}
function unscramble(data: string, key: string): string {
  const k = [...key].map((c) => c.charCodeAt(0))
  const bin = atob(data)
  const bytes = Uint8Array.from(bin, (c, i) => c.charCodeAt(0) ^ k[i % k.length])
  return new TextDecoder().decode(bytes)
}

export class MentorStore {
  constructor(private kv: KV = browserKV) {}

  profile(): Profile | null {
    return read<Profile>(this.kv, 'profile')
  }
  saveProfile(p: Profile) {
    this.kv.set(P + 'profile', JSON.stringify(p))
  }

  learner(): LearnerState {
    return read<LearnerState>(this.kv, 'learner') ?? { concepts: {}, history: [] }
  }
  saveLearner(s: LearnerState) {
    this.kv.set(P + 'learner', JSON.stringify(s))
  }

  settings(): MentorSettings {
    const s = read<MentorSettings>(this.kv, 'settings')
    // "none" (the old built-in mentor) no longer exists: every session uses Gemini
    const connection = s?.connection === 'proxy' || (!s?.connection && BUILT_IN_PROXY) ? 'proxy' : 'key'
    return { connection, apiKey: s?.apiKey ?? '', proxyUrl: s?.proxyUrl ?? BUILT_IN_PROXY, model: s?.model || DEFAULT_MODEL }
  }
  saveSettings(s: MentorSettings) {
    this.kv.set(P + 'settings', JSON.stringify(s))
  }

  session(wbId: string): MentorSession | null {
    return read<MentorSession>(this.kv, `session.${wbId}`)
  }
  saveSession(s: MentorSession) {
    const copy = { ...s, messages: s.messages.slice(-150) }
    this.kv.set(P + `session.${s.wbId}`, JSON.stringify(copy))
  }
  removeSession(wbId: string) {
    const s = this.session(wbId)
    this.kv.remove(P + `session.${wbId}`)
    if (s) {
      this.kv.remove(P + `original.${wbId}`)
      if (s.problem) this.kv.remove(P + `answers.${s.problem.problemId}`)
    }
  }

  saveAnswers(problemId: string, targets: Target[]) {
    this.kv.set(P + `answers.${problemId}`, scramble(JSON.stringify(targets), problemId))
  }
  answers(problemId: string): Target[] | null {
    const s = this.kv.get(P + `answers.${problemId}`)
    if (!s) return null
    try {
      return JSON.parse(unscramble(s, problemId)) as Target[]
    } catch {
      return null
    }
  }

  saveOriginal(wbId: string, doc: WorkbookDoc) {
    this.kv.set(P + `original.${wbId}`, JSON.stringify(doc))
  }
  original(wbId: string): WorkbookDoc | null {
    return read<WorkbookDoc>(this.kv, `original.${wbId}`)
  }

  badVideos(): Set<string> {
    return new Set(read<string[]>(this.kv, 'videos.unavailable') ?? [])
  }
  markVideoBad(id: string) {
    const s = this.badVideos()
    s.add(id)
    this.kv.set(P + 'videos.unavailable', JSON.stringify([...s]))
  }

  /** Mentor sessions stored in this browser (wbId list). */
  sessionIds(): string[] {
    const out: string[] = []
    try {
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i)
        if (k?.startsWith(P + 'session.')) out.push(k.slice((P + 'session.').length))
      }
    } catch {
      /* ignore */
    }
    return out
  }

  resetAll() {
    for (const id of this.sessionIds()) this.removeSession(id)
    for (const k of ['profile', 'learner']) this.kv.remove(P + k)
  }
}
