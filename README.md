# wOS Sheets

Spreadsheets your team owns. Formulas, tabs, formatting, frozen rows, sort and filter, charts, pivot tables, comments, editing together live, version history, and Excel and CSV in and out.

Sheets has no AI of its own and never sells you AI credits. Your own AI (Claude, ChatGPT, Claude Code, Codex, or anything that speaks MCP) connects with one click and works the sheet through the same tools the buttons use: it reads exact values and formulas, writes cells, builds charts and pivots, and pulls records from the wOS CRM.

Try it: **https://sheets.waronsaas.com**. The examples open without an account (a sales pipeline, a budget and a hiring tracker, all made up). Sign in to make your own.

| Host it yourself, free | Host it with us |
|---|---|
| `docker compose up` runs Sheets and Postgres on any server. One person can run it with Node and a SQLite file. No licence check, ever. | We run it and charge what it costs us times two, shown openly. Move to your own server any time with one export. |

Licence: AGPL-3.0-only. Status: early (v0.1). Decisions and licences of what it is built on: [`docs/DECISIONS.md`](docs/DECISIONS.md).

## Host it yourself

**On a server (a team):**

```sh
git clone https://github.com/warOnSaaS/sheets && cd sheets
cp .env.example .env            # set OAUTH_SECRET (openssl rand -hex 32) and how people sign in
docker compose up -d            # Sheets + Postgres, on http://localhost:3994
```

Put HTTPS in front (Caddy, nginx, or your host's). The first person to sign in owns the server and adds everyone else in Settings. Database changes apply themselves when the server starts.

**On your own computer (one person, nothing else to install):**

```sh
npm install
npm start                       # http://localhost:3994, SQLite in ./data/sheets.db
```

Without email set up, the sign-in link is printed in the terminal.

**What a self-hoster needs** (more than a database):

| Piece | Needed for | Free option |
|---|---|---|
| A machine that stays on, with Docker or Node 20+ | everything (live editing holds a connection open) | a small VPS, a home server, your laptop |
| Postgres (`DATABASE_URL`) or nothing (SQLite) | storage | the bundled Postgres, Neon, Supabase |
| HTTPS | sign-in callbacks, connecting Claude or ChatGPT | Caddy or your host |
| A way to sign in | people | an email link (SMTP from any mailbox you have, or Resend), GitHub (a free OAuth app), or a warOnSaaS account |
| The wOS CRM (optional) | pulling records into a sheet | self-host it too, or connect ours |

Every setting: [`.env.example`](.env.example). Your data, any time: Settings, Export everything (`sheets.export_data`) gives every spreadsheet as .xlsx plus all of it as JSON.

## Connect your AI

Signed in, open **Connect your AI**. Claude: one button. ChatGPT: open it and paste the copied address. Claude Code and Codex: paste one line in a terminal (`curl -fsSL https://<your sheets>/connect/claude | sh`). Anything else: the MCP address is `https://<your sheets>/mcp` (standard MCP sign-in: discovery, dynamic registration, PKCE). Plain HTTPS works too: `POST /api/tools/<name>`, described in `/openapi.json`.

Things to ask: "Build a pipeline sheet from my CRM deals with weighted value by stage and a chart." "Which month in the Budget has the lowest margin? Show the formula."

## What it does today, and what Google Sheets does that it does not yet

| Works | Not yet (honest list) |
|---|---|
| 430 functions (HyperFormula plus ours: TEXT and DATEDIF as Excel has them, CONCAT, AVERAGEIFS, RANK, REGEXMATCH, REGEXEXTRACT, REGEXREPLACE), cross-tab references, named ranges, FILTER, SORT, UNIQUE, XLOOKUP | INDIRECT, LET, LAMBDA, XMATCH, QUERY, IMPORTRANGE, GOOGLEFINANCE, and Google's ARRAYFORMULA |
| Several tabs, rename (formulas follow), move, delete | Hiding tabs, protected ranges, tab colours in the screen |
| Number formats (Excel codes), bold, italic, underline, strike, text and fill colours, alignment, wrap, column widths | Borders, merged cells, fonts other than the app's, rich text inside one cell, custom row heights, conditional formatting, data validation (drop-downs) |
| Freeze rows and columns; sort (formulas move with their rows); a shared filter by values or conditions | Filter views (a private filter per person), slicers |
| Column, bar, line, area, pie and scatter charts, moved and resized on the grid | Combo charts, chart styling, trendlines, charts in .xlsx files |
| Pivot tables written as live formulas (SUMIFS and friends), refreshable | Pivot tables you drag fields around in |
| Comments with replies, resolve | @mentions that notify, assigning a comment |
| Editing together live (Yjs), who is here and where | Each person's own undo across sessions (undo here covers your edits, formats, fills and sorts in this window; version history covers the rest), offline editing |
| Version history: automatic every ten minutes of editing, named versions, look at one, restore | Seeing who changed which cell |
| .xlsx in (with a report of what did not carry over, and of any formula that works out differently from what Excel saved) and out (formulas with their results); CSV in and out | .xls, .ods and Google Sheets files; charts and images inside .xlsx; files over about 3 MB from the browser (larger ones through the API) |
| View-only links anyone can open; everyone on the team can edit | Sharing one spreadsheet with someone outside the team as an editor; commenter-only access |
| Pull CRM records onto a linked tab, refresh it | Writing back to the CRM from a sheet |
| Every action is a tool, for people and AI alike | Apps Script, macros |

## How it is built

| Part | Where | What |
|---|---|---|
| Server | `server.mjs`, `lib/` | One Node process: pages, `/api/tools/<name>` and `/mcp` from the same handlers (`lib/tools.mjs`), sign-in (`lib/auth.mjs`), live editing over `/ws`, downloads |
| A workbook | `lib/book.mjs` | One Yjs document: tabs, cells (`"r,c"` to what was typed), formats, comments, charts, named ranges |
| Formulas | `lib/engine.mjs`, `lib/functions.mjs` | HyperFormula kept in step with the document, the same on the server (what agents read) and in the browser (what people see) |
| Operations | `lib/ops.mjs` | Every change a tool makes: write, fill, sort, filter, insert and delete rows, format, charts, pivots, comments |
| Storage | `lib/store.mjs`, `migrations/` | Postgres or SQLite: each workbook's folded state plus the changes since, merged by Yjs, so any number of server copies can write at once; versions; view-only links |
| Excel and CSV | `lib/xlsx.mjs` | ExcelJS for the workbook, fflate to count what ExcelJS does not read, our own report |
| The CRM | `lib/crm.mjs` | Calls the CRM's own `crm.query` tool (in the suite, `call.callTool`; standalone, its REST address) |
| Screens | `public/app/` | The grid (`grid.mjs`), list, Connect your AI and Settings (`sheets.mjs`), charts as SVG (`chart.mjs`), on the ui-design kit (`public/ui`, copied by `npm run sync-kit`) |
| In the wOS suite | `wos-app.json`, `tools.json`, `server.mjs` default export, `screens.mjs` | The suite contract: `register(ctx)` returns the handlers; `screens.mjs` is built by `npm run build` |

## Agent parity

Everything a person can do is a tool, and the screens only call tools. The build checks it:

| Check | Command | Fails when |
|---|---|---|
| Screen to tool | `npm test` (`test/parity.test.mjs`, Playwright) | a button, menu item, select, form or file picker on any screen, menu, dialog or popover, at 1440 and 390 wide, names no tool or a tool not in the catalogue. Today: 100% (1,159 actions on 34 screens) |
| Every tool works | `test/tools.test.mjs` | a tool in the catalogue was never called, or a journey breaks |
| No side doors | `test/parity.test.mjs` | screen code calls anything but `/api/tools/*`, downloads under `/files/sheets/`, and `/ws` |
| An agent over MCP only | `test/agent.test.mjs` | an agent cannot build a pipeline with formulas and a chart, or an .xlsx does not round-trip with the right values (checked with openpyxl) |
| Suite contract | `node ~/wos-suite/packages/manifest/bin/check.mjs .` | `wos-app.json` or `tools.json` is off the contract |

## Development

```sh
npm install
npm run build                  # the browser bundle (public/app/page.js) and screens.mjs
npm start                      # http://localhost:3994
npm test                       # everything; the openpyxl checks need: python3 -m venv .venv && .venv/bin/pip install openpyxl
npm run test:pg                # the tools and two server copies on Postgres (SHEETS_TEST_PG)
npm run shots                  # screenshots at 1440 and 390, light and dark, into .shots/
npm run check                  # all of the above checks, plus the private-names scan
```

Tool names: `sheets.verb_noun` in `tools.json` and the code, `sheets_verb_noun` on the wire (MCP, OpenAPI). Both are accepted.
