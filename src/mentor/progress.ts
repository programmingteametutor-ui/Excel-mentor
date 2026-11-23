/* Concept mastery, adaptive difficulty and problem recommendations. */
import { allPrereqs, CONCEPTS, conceptName, CURRICULUM, FOUNDATIONS, GENERATABLE } from './concepts'
import { STAGE_INFO, STAGE_ORDER } from './protocol'
import type { CheckRecord, ConceptId, ConceptProgress, LearnerState, MasteryStage, PerformanceReport, ProblemLevel, ProblemPublic, Profile, Role, StageRecord, TargetStats } from './types'

export const STAGES: MasteryStage[] = ['locked', 'introduced', 'learning', 'practicing', 'applying', 'strong', 'mastered']
export const STAGE_LABEL: Record<MasteryStage, string> = { locked: 'Locked', introduced: 'Introduced', learning: 'Learning', practicing: 'Practicing', applying: 'Applying', strong: 'Strong', mastered: 'Mastered' }

export function emptyProgress(): ConceptProgress {
  return { stage: 'locked', understandingScore: 0, understandingCount: 0, practiceCorrect: 0, practiceAttempts: 0, appliedCorrect: 0, appliedTotal: 0, hintLevels: 0, independentProblems: 0, problems: 0 }
}

export function conceptProgress(s: LearnerState, id: ConceptId): ConceptProgress {
  return s.concepts[id] ?? emptyProgress()
}

const pct = (a: number, b: number) => (b ? Math.round((100 * a) / b) : 0)

export function scores(p: ConceptProgress) {
  return {
    understanding: p.understandingCount ? Math.round((100 * p.understandingScore) / p.understandingCount) : 0,
    practice: pct(p.practiceCorrect, p.practiceAttempts),
    application: pct(p.appliedCorrect, p.appliedTotal),
    hintAverage: p.appliedTotal ? p.hintLevels / p.appliedTotal : 0,
  }
}

/** Evidence-based stage: never “mastered” from a single correct answer. */
export function computeStage(p: ConceptProgress): MasteryStage {
  const s = scores(p)
  if (p.problems >= 2 && p.independentProblems >= 2 && s.application >= 90 && s.hintAverage <= 2) return 'mastered'
  if (p.problems >= 1 && s.application >= 80 && s.hintAverage <= 3) return 'strong'
  if (p.appliedTotal > 0) return 'applying'
  if (p.practiceAttempts > 0) return 'practicing'
  if (p.understandingCount > 0) return 'learning'
  if (p.stage !== 'locked') return 'introduced'
  return 'locked'
}

function update(s: LearnerState, id: ConceptId, fn: (p: ConceptProgress) => void): LearnerState {
  const p = { ...conceptProgress(s, id) }
  fn(p)
  p.stage = computeStage(p)
  return { ...s, concepts: { ...s.concepts, [id]: p } }
}

export const markIntroduced = (s: LearnerState, id: ConceptId) =>
  update(s, id, (p) => {
    if (p.stage === 'locked') p.stage = 'introduced'
  })
export const recordThink = (s: LearnerState, id: ConceptId, score: number) =>
  update(s, id, (p) => {
    if (p.stage === 'locked') p.stage = 'introduced'
    p.understandingScore += Math.max(0, Math.min(1, score))
    p.understandingCount += 1
  })
export const recordPractice = (s: LearnerState, id: ConceptId, correct: boolean) =>
  update(s, id, (p) => {
    if (p.stage === 'locked') p.stage = 'introduced'
    p.practiceAttempts += 1
    if (correct) p.practiceCorrect += 1
    p.lastPracticedAt = new Date().toISOString()
  })

/** Applies the final check of a problem to every concept it used. */
export function recordProblem(s: LearnerState, pub: ProblemPublic, last: CheckRecord, stats: Record<string, TargetStats>): LearnerState {
  const byConcept = new Map<ConceptId, { ok: number; total: number; hints: number; independent: boolean }>()
  for (const t of pub.targets) {
    const res = last.results.find((r) => r.id === t.id)
    const e = byConcept.get(t.concept) ?? { ok: 0, total: 0, hints: 0, independent: true }
    const hint = stats[t.id]?.hintLevel ?? 0
    e.total++
    e.hints += hint
    if (res?.state === 'correct') e.ok++
    if (res?.state !== 'correct' || hint > 2) e.independent = false
    byConcept.set(t.concept, e)
  }
  let out = s
  for (const [id, e] of byConcept)
    out = update(out, id, (p) => {
      if (p.stage === 'locked') p.stage = 'introduced'
      p.appliedCorrect += e.ok
      p.appliedTotal += e.total
      p.hintLevels += e.hints
      p.problems += 1
      if (e.independent && e.ok === e.total) p.independentProblems += 1
    })
  return out
}

/** Overall result of a problem: drives the next difficulty. */
export function performanceScore(pub: ProblemPublic, checks: CheckRecord[], stats: Record<string, TargetStats>): { firstCheck: number; final: number; hintRate: number; assisted: boolean } {
  const total = pub.targets.length || 1
  const firstCheck = checks[0] ? checks[0].correct / total : 0
  const final = checks.length ? checks[checks.length - 1].correct / total : 0
  const hints = pub.targets.reduce((a, t) => a + (stats[t.id]?.hintLevel ?? 0), 0)
  const hintRate = hints / total
  const fullAnswers = pub.targets.filter((t) => (stats[t.id]?.hintLevel ?? 0) >= 5).length
  return { firstCheck, final, hintRate, assisted: fullAnswers / total > 0.25 || hintRate > 2.5 }
}

const LEVELS: ProblemLevel[] = ['beginner', 'intermediate', 'advanced']

export function startLevel(profile: Profile): ProblemLevel {
  if (profile.level === 'advanced') return profile.difficulty === 'easy' ? 'intermediate' : 'advanced'
  if (profile.level === 'intermediate') return profile.difficulty === 'challenging' ? 'advanced' : 'intermediate'
  return profile.difficulty === 'challenging' && profile.level === 'beginner' ? 'intermediate' : 'beginner'
}

/** Next level from evidence: harder after a clean solve, same after a struggle. */
export function nextLevel(cur: ProblemLevel, perf: ReturnType<typeof performanceScore>): ProblemLevel {
  const i = LEVELS.indexOf(cur)
  if (perf.firstCheck >= 0.85 && perf.hintRate <= 1 && !perf.assisted) return LEVELS[Math.min(2, i + 1)]
  if (perf.assisted || perf.firstCheck < 0.4) return LEVELS[Math.max(0, i - 1)]
  return cur
}

export interface Recommendation {
  level: ProblemLevel
  newConcepts: ConceptId[]
  revision: ConceptId[]
  estimatedMinutes: number
  reason: string
}

/** Picks the next concepts by curriculum order, respecting prerequisites and weak areas. */
/** Concepts each role leans on most, used when the mentor chooses the topics. */
export const ROLE_FOCUS: Partial<Record<Role, ConceptId[]>> = {
  'Data Analyst': ['sum', 'average', 'countif', 'sumif', 'averageif', 'xlookup', 'index-match', 'chart'],
  HR: ['count', 'countif', 'if', 'average', 'xlookup', 'concat'],
  Finance: ['sum', 'round', 'percent', 'sumif', 'abs-ref', 'if', 'vlookup'],
  Sales: ['sum', 'sort', 'chart', 'sumif', 'percent', 'max', 'xlookup'],
  Marketing: ['sum', 'percent', 'average', 'countif', 'chart', 'averageif'],
  Operations: ['sum', 'max', 'min', 'if', 'countif', 'sumif', 'xlookup'],
}

/** A problem for topics the student chose (at most 3), at the level their progress suggests. */
export function topicRecommendation(profile: Profile, s: LearnerState, chosen: ConceptId[]): Recommendation {
  const level = s.nextLevel ?? startLevel(profile)
  const newConcepts = CURRICULUM.filter((c) => chosen.includes(c) && GENERATABLE.includes(c)).slice(0, 3)
  return { level, newConcepts, revision: [], estimatedMinutes: Math.min(90, 15 + 10 * newConcepts.length), reason: `You chose ${newConcepts.map(conceptName).join(', ')}.` }
}

export function recommend(profile: Profile, s: LearnerState): Recommendation {
  const level = s.nextLevel ?? startLevel(profile)
  const stage = (id: ConceptId) => conceptProgress(s, id).stage
  const known = (id: ConceptId) => ['applying', 'strong', 'mastered'].includes(stage(id))
  const weak = CURRICULUM.filter((id) => GENERATABLE.includes(id) && (stage(id) === 'applying' || (conceptProgress(s, id).appliedTotal > 0 && scores(conceptProgress(s, id)).application < 70)))
  const ready = (id: ConceptId) => [...allPrereqs(id)].every((p) => FOUNDATIONS.includes(p) || known(p) || !GENERATABLE.includes(p))
  const maxLevel = LEVELS.indexOf(level)
  // concepts a person in this role uses most come first (the curriculum order breaks ties)
  const focus = ROLE_FOCUS[profile.role] ?? []
  const rank = (id: ConceptId) => (focus.includes(id) ? focus.indexOf(id) - 100 : CURRICULUM.indexOf(id))
  const candidates = CURRICULUM.filter((id) => GENERATABLE.includes(id) && !known(id) && ready(id) && LEVELS.indexOf(CONCEPTS[id].level) <= maxLevel).sort((a, b) => rank(a) - rank(b))
  const last = s.history[s.history.length - 1]
  const struggling = last && last.completedAt && (last.assisted || (last.firstCheckScore ?? 1) < 0.5)
  const nNew = struggling ? (weak.length ? 0 : 1) : profile.difficulty === 'challenging' ? 3 : profile.duration <= 15 ? 1 : 2
  let newConcepts = candidates.slice(0, nNew)
  if (!newConcepts.length && !weak.length) newConcepts = CURRICULUM.filter((id) => GENERATABLE.includes(id) && !known(id)).slice(0, 1)
  const revision = [...weak.slice(0, 2), ...CURRICULUM.filter((id) => known(id) && !weak.includes(id)).slice(-1)].filter((id) => !newConcepts.includes(id)).slice(0, 2)
  if (!newConcepts.length && !revision.length) newConcepts = ['sum']
  const all = [...newConcepts, ...revision]
  const reason = struggling
    ? `More practice with ${all.map(conceptName).join(', ')} before moving on.`
    : newConcepts.length
      ? `Next step: ${newConcepts.map(conceptName).join(' and ')}${revision.length ? `, with a quick revision of ${revision.map(conceptName).join(', ')}` : ''}.`
      : `Strengthen ${revision.map(conceptName).join(', ')}.`
  return { level, newConcepts, revision, estimatedMinutes: Math.min(profile.duration, 10 + 8 * newConcepts.length + 4 * revision.length), reason }
}

export function buildReport(pub: ProblemPublic, checks: CheckRecord[], stats: Record<string, TargetStats>, timeMs: number, next: Recommendation, log: StageRecord[] = [], masteryProfile?: string): PerformanceReport {
  const last = checks[checks.length - 1]
  const perf = performanceScore(pub, checks, stats)
  const hints = Object.values(stats).reduce((a, s) => a + s.hintsUsed, 0)
  const byConcept = new Map<ConceptId, { ok: number; total: number; hints: number }>()
  for (const t of pub.targets) {
    const e = byConcept.get(t.concept) ?? { ok: 0, total: 0, hints: 0 }
    e.total++
    e.hints += stats[t.id]?.hintLevel ?? 0
    const first = checks[0]?.results.find((r) => r.id === t.id)
    if (first?.state === 'correct') e.ok++
    byConcept.set(t.concept, e)
  }
  const strong: string[] = []
  const needs: string[] = []
  for (const [id, e] of byConcept) {
    if (e.ok / e.total >= 0.8 && e.hints / e.total <= 2) strong.push(conceptName(id))
    else needs.push(conceptName(id))
  }
  // concepts whose learn/practice gate scored low also need practice
  for (const r of log) if (r.concept && (r.stage === 'learn' || r.stage === 'practice') && r.score < 60 && !needs.includes(conceptName(r.concept))) needs.push(conceptName(r.concept))
  const independence = perf.hintRate <= 0.5 && perf.firstCheck >= 0.8 ? 'High' : perf.assisted ? 'Low' : 'Moderate'
  const sections = STAGE_ORDER.map((stage) => {
    const recs = log.filter((r) => r.stage === stage)
    if (!recs.length) return { area: STAGE_INFO[stage].area, score: null, note: 'Not assessed' }
    const score = Math.round(recs.reduce((a, r) => a + r.score, 0) / recs.length)
    return { area: STAGE_INFO[stage].area, score, note: recs.map((r) => r.summary).filter((x) => x && x !== '(no summary)').join(' ') || '—' }
  })
  return {
    title: pub.title,
    mode: pub.mode,
    completion: last ? Math.round((100 * last.correct) / Math.max(1, last.total)) : 0,
    cells: { correct: last?.correct ?? 0, total: last?.total ?? pub.targets.length },
    conceptsLearned: pub.concepts,
    strong: strong.filter((x) => !needs.includes(x)),
    needsPractice: needs,
    hints,
    independence,
    checks: checks.length,
    timeMinutes: Math.max(1, Math.round(timeMs / 60000)),
    recommendation: next.reason,
    sections,
    ...(masteryProfile ? { masteryProfile } : {}),
    nextSteps: { revise: needs, problem: next.reason, difficulty: next.level },
  }
}
