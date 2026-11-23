/**
 * ExcelMentor Gemini proxy for Cloudflare Workers.
 * Keeps your Gemini API key on the server so students don't need their own key.
 *
 * Secrets / variables (Workers → Settings → Variables):
 *   GEMINI_API_KEY   (secret)  your key from aistudio.google.com/apikey
 *   ALLOWED_ORIGINS  e.g. "https://yourname.github.io,http://localhost:5173"
 *   ALLOWED_MODELS   optional, comma-separated, e.g. "gemini-3.6-flash,gemini-3.5-flash-lite"
 *
 * The app calls:  POST <worker-url>/models/<model>:generateContent
 *                 GET  <worker-url>/models            (model list for the settings screen)
 */
const UPSTREAM = 'https://generativelanguage.googleapis.com/v1beta'
const MAX_BODY = 200_000 // bytes

export default {
  async fetch(request, env) {
    const origin = request.headers.get('Origin') ?? ''
    const allowed = (env.ALLOWED_ORIGINS ?? '').split(',').map((s) => s.trim()).filter(Boolean)
    const okOrigin = allowed.length === 0 || allowed.includes(origin)
    const cors = {
      'Access-Control-Allow-Origin': okOrigin ? origin || '*' : 'null',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '86400',
      Vary: 'Origin',
    }
    if (request.method === 'OPTIONS') return new Response(null, { status: okOrigin ? 204 : 403, headers: cors })
    if (!okOrigin) return json({ error: { message: 'Origin not allowed' } }, 403, cors)
    if (!env.GEMINI_API_KEY) return json({ error: { message: 'Proxy is missing GEMINI_API_KEY' } }, 500, cors)

    const url = new URL(request.url)
    const models = (env.ALLOWED_MODELS ?? '').split(',').map((s) => s.trim()).filter(Boolean)

    if (request.method === 'GET' && url.pathname.replace(/\/+$/, '') === '/models') {
      const res = await fetch(`${UPSTREAM}/models?pageSize=200`, { headers: { 'x-goog-api-key': env.GEMINI_API_KEY } })
      const data = await res.json()
      if (models.length && data.models) data.models = data.models.filter((m) => models.includes(m.name.replace(/^models\//, '')))
      return json(data, res.status, cors)
    }

    const m = /^\/models\/([\w.\-]+):generateContent$/.exec(url.pathname)
    if (request.method !== 'POST' || !m) return json({ error: { message: 'Not found' } }, 404, cors)
    if (models.length && !models.includes(m[1])) return json({ error: { message: `Model ${m[1]} is not enabled on this proxy` } }, 404, cors)
    const body = await request.text()
    if (body.length > MAX_BODY) return json({ error: { message: 'Request too large' } }, 413, cors)
    const res = await fetch(`${UPSTREAM}/models/${m[1]}:generateContent`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-goog-api-key': env.GEMINI_API_KEY },
      body,
    })
    return new Response(res.body, { status: res.status, headers: { ...cors, 'Content-Type': 'application/json' } })
  },
}

function json(data, status, headers) {
  return new Response(JSON.stringify(data), { status, headers: { ...headers, 'Content-Type': 'application/json' } })
}
