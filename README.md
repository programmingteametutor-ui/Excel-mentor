# ExcelMentor

An AI Excel tutor that runs in the browser. Students learn by solving real business problems in a real spreadsheet. Google Gemini does the teaching, and the app checks every answer with its own formula engine.

There are two ways to start a session. Both build a problem first, then teach it in the same nine stages.

| Mode | How the problem is made |
|---|---|
| **1 · Learn by topic** | The student picks up to 3 topics, **or** lets the mentor decide from their Excel level, role, domain and progress. Gemini writes a business problem and dataset for those topics. |
| **2 · Take-home assessment** | The student uploads what an employer sent: Excel/CSV data and a PDF, Word or text brief. The real data becomes the Task sheet. Each question that can be checked becomes a yellow answer cell. The rest (PivotTables, written commentary…) is listed as "not checked automatically" and covered in conversation. The mentor teaches the skills on separate practice data and never writes the student's answers. |

## Quick start

```bash
npm install
npm run dev          # http://localhost:5173 (rebuilds on change)
npm test             # unit and integration tests
npm run build        # dist/index.html + dist/assets (the PDF reader is a separate chunk, loaded only when needed)
npm run build:single # dist/excelpro.html, one self-contained file
```

ExcelMentor **requires Gemini**. Open Mentor settings and paste a Gemini API key, which is stored only in that browser. You can instead deploy the proxy in [`proxy/`](proxy/README.md) to keep one key on a server. There is no offline mode: no built-in problem bank and no keyword-based answer judging.

## The nine stages

Every session goes through the same stages. Learn and Practice repeat once for each concept.

1. **Problem**: the business situation and its functional and non-functional requirements. The mentor demonstrates clarifying questions, then the student asks their own. The checklist has 7 areas and the student must cover at least 4. Then the "Official Context" is given.
2. **Concepts**: building versus typing (the analyst mindset), and the roadmap of concepts for this problem.
3. **Learn**: one concept per step, with a worked example on the *Learning Example* sheet, a rotating delivery style and a unique Pressure Test.
4. **Practice**: the same concept with new data on the *Practice* sheet. The engine must confirm the formula, then the student justifies it and names an alternative.
5. **Plan**: a mini-plan prelude, then Place-Label-Link for each task, sanity checks, and a narration of one data row's journey through the sheet.
6. **Solve**: the student fills in the yellow Task cells. Hints follow the 4-level Adaptive Pivot protocol.
7. **Check**: Check My Work runs in the engine, followed by a deep dive into failure modes (blanks, new rows, missing lookups).
8. **Follow-up**: trade-off cards, decision points where the student would switch approach ("flip points"), and a teach-back to a non-technical manager.
9. **Report**: the master blueprint, a reflection, the closing hand-off, and a 10–15 sentence Mastery Profile. The report scores each stage.

## How the teaching rules are enforced

The teaching rules live in [`src/mentor/protocol.ts`](src/mentor/protocol.ts). They are the ArchMentor protocols adapted from system design to Excel: Atomic Teaching, Context Validation, Pressure Tests (Logic A–D), Human-First style, Build-in-the-Sheet, the Rigorous Gatekeeper with Three-Pillar validation, Anti-Parrot, Teach-Then-Guide, Adaptive Pivot hints, the Mastery Profile, and the readability rules.

Every Gemini reply is JSON: `message_to_student`, `current_stage`, `gate_passed`, `verdict`, `hint_level`, `pressure_test`, `context_added`, `stage_summary`, `section_score`, `session_complete`, `mastery_profile` and a few more. The app checks each reply before the student sees it ([`checkTurn`](src/mentor/protocol.ts)):

- **State Lockdown and Wait for Signal.** A reply may only stay where it is or move to the single next position, and only with `gate_passed` and `verdict: "pass"`. Skipping ahead and going back are rejected.
- **Evidence.** Practice can't pass until the formula engine marks the practice cell correct. Check can't pass until every Task cell is correct. A pending teach-back after a Level-4 reveal blocks the gate. Moving from Solve to Check happens only when the student clicks Check My Work.
- **No Spoiling.** A reply that contains a Task formula, or the exact result of a summary cell, is hidden. Level 4 may reveal one task's formula in a topic session, but never in a take-home.
- **Banned phrases.** Banned template phrases, "Does this make sense?", and superlative praise for an answer that hasn't passed are all flagged.
- **Final hand-off.** It needs the exact closing sentence and a Mastery Profile of 10–15 sentences.
- **Chunking.** Prose paragraphs longer than 3 sentences are split automatically.

A reply that breaks a rule is sent back once with the list of broken rules. If it still breaks a gate rule, the student stays where they are. If it still leaks an answer, it is hidden and a Retry button is shown.

The closing sentence is set per mode in `CLOSING_PHRASE` (`protocol.ts`). Topic sessions end with "You are now ready for the AI Mock Interview."

## Architecture

```
src/mentor/protocol.ts  system prompt (teaching rules), reply schema, reply checks, chunking
src/mentor/flow.ts      stage order, evidence per gate, the state sent each turn, applying a reply, leak detection
src/mentor/ai.ts        Gemini: topic problems, take-home problems, lesson tables, one mentor turn (with a correction round)
src/mentor/extract.ts   reads uploads: .xlsx/.xlsm/.csv/.tsv (real tables with header row and types), .docx, .pdf (pdf.js), .txt/.md
src/mentor/build.ts     turns a problem draft into a workbook plus hidden answers; guardrails (real uploaded data is allowed blanks and repeats)
src/mentor/validate.ts  Check My Work, cell by cell
src/mentor/progress.ts  mastery per concept, role-aware topic choice, the 9-stage report
src/ui/mentor/          dashboard, take-home upload page, mentor panel, settings
src/engine, src/grid    the spreadsheet engine (HyperFormula) and grid the student works in
proxy/                  optional Cloudflare Worker that keeps a Gemini key secret
```

The app opens on the ExcelMentor dashboard. The spreadsheet appears only inside a mentor session (Instructions, Task, Data, Learning Example and Practice sheets). Sessions and workbooks are saved in the visitor's own browser.

## Browser test

```bash
npm run build && node tests/e2e/serve.mjs dist 4173 &
PLAYWRIGHT_MODULE=playwright CHROME_PATH=/path/to/chrome node tests/e2e/mentor.mjs
```

[`tests/e2e/mentor.mjs`](tests/e2e/mentor.mjs) mocks Gemini and runs both modes. It goes through all nine stages and checks gate enforcement, practice and Check My Work in the engine, hint requests, leak hiding, and CSV, Word and PDF uploads.

## Publish on GitHub Pages

The build is fully static, so it runs from any sub-path. Push to `main` and set **Settings → Pages → Source: GitHub Actions**. The workflow type-checks, tests, builds and deploys. HyperFormula is GPLv3, so a public deployment should keep this source public (or use a HyperFormula commercial licence).

## Honest limits

- Without a server, the answer key lives in the browser. It is kept out of the workbook and obscured, but a determined student could find it.
- Take-home checks follow Gemini's reading of the brief. The mentor and the Results tab say so.
- Scanned (image-only) PDFs can't be read. Upload a text PDF or a .docx copy instead.
- `.xls`, `.doc` and `.ods` files must be re-saved as `.xlsx` or `.docx` first.
