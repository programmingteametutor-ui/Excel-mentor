# ExcelMentor Gemini proxy (optional)

ExcelPro is a static site, so a Gemini API key placed in the page would be visible to everyone.
There are two ways to connect the mentor:

1. **Each student uses their own key** — Mentor settings → “My Gemini API key”. The key stays in that browser and is sent only to Google. Nothing to deploy.
2. **One key for everybody, kept secret** — deploy this small proxy and point the app at it.

## Deploy on Cloudflare Workers (free tier is enough to start)

1. Create a Worker at dash.cloudflare.com → Workers & Pages → Create → “Hello World”, then replace its code with `gemini-worker.js`.
2. Settings → Variables:
   - `GEMINI_API_KEY` — **Secret**, your key from aistudio.google.com/apikey
   - `ALLOWED_ORIGINS` — your site, e.g. `https://programmingteametutor-ui.github.io`
   - `ALLOWED_MODELS` — optional, e.g. `gemini-3.6-flash`
3. Deploy and copy the Worker URL (e.g. `https://excelmentor-proxy.<you>.workers.dev`).

## Point the app at it

- Per browser: Mentor settings → “Proxy server” → paste the URL.
- For everyone: build with the URL baked in. In GitHub, add a repository **variable** `EXCELMENTOR_PROXY_URL` (Settings → Secrets and variables → Actions → Variables). The Pages workflow passes it to the build, and the mentor then connects automatically.

## Notes

- Add rate limiting (Cloudflare → Security → WAF → Rate limiting rules) before sharing the site widely: every visitor's requests use your quota.
- The proxy forwards only `models/<model>:generateContent` and the model list; nothing is stored.
