# Decisions

Plain answers first; the detail is in the tables.

## 1. The grid and the formula engine: HyperFormula plus our own grid, not Univer

**Decision:** formulas are worked out by **HyperFormula** (GPL-3.0-only, compatible with our AGPL-3.0-only) on the server and in the browser. The grid is our own, drawn with the ui-design kit. Univer was evaluated first, as the brief asked, and does not fit.

Checked on npm on 2026-10-07 (version 1.0.3 of every Univer package):

| Univer part | Licence | What it gives | Needed for our MVP? |
|---|---|---|---|
| `@univerjs/core`, `sheets`, `sheets-ui`, `engine-formula`, `sheets-formula`, `sheets-filter`, `sheets-sort`, `sheets-thread-comment`, `presets` | Apache-2.0 | The grid, formulas, filter, sort, comments, number formats | Yes |
| `@univerjs-pro/collaboration-client` (and the collaboration server) | No open licence (no `license` field; needs `@univerjs-pro/license`, a paid key) | Editing together live | **Yes** |
| `@univerjs-pro/sheets-chart` | Same, paid | Charts | **Yes** |
| `@univerjs-pro/sheets-pivot` | Same, paid | Pivot tables | **Yes** |
| `@univerjs-pro/exchange-client` | Same, paid | xlsx import and export | **Yes** |
| `@univerjs-pro/edit-history-viewer` | Same, paid | Version history | **Yes** |

Five of the MVP's needs are in Univer's paid part. The free part would still leave us building live editing (on Univer's internal command format), charts, pivots, xlsx and history ourselves, inside someone else's UI framework (`@univerjs/design`, React and RxJS, about 9 MB unpacked for `sheets-ui` alone) that the ui-design kit cannot style. Agents also need exact values on the server with no browser, and Univer's engine is built around its UI stack.

HyperFormula is only an engine: no UI, runs the same in Node and the browser, 423 functions with Excel-compatible results, named ranges, cross-tab references that follow a tab rename, and undo-free batch updates. The same raw cell contents give the same values on the server (what agents read and what export writes) and in the browser (what people see).

| Piece | Choice | Licence | Why |
|---|---|---|---|
| Formula engine | HyperFormula 3.4 (`licenseKey: 'gpl-v3'`) | GPL-3.0-only | Exact Excel-style results in Node and the browser; GPL-3.0 code may be combined with AGPL-3.0 (GPL-3.0 section 13) |
| Missing functions | Our own HyperFormula plugin: `CONCAT`, `AVERAGEIFS`, `RANK`, `XMATCH` (and more as people ask) | ours, AGPL | HyperFormula 3.4 lacks them |
| Grid | Our own, virtualised DOM grid on ui-design | ours, AGPL | Kit styles, 390px and 1440px, every action is a tool |
| Live editing | Yjs 13 (one document per workbook), WebSockets plus Postgres LISTEN/NOTIFY between server copies, polling as the fallback | MIT | Concurrent edits merge without locks, any number of server copies can append changes, and a reconnecting screen catches up from its state vector |
| Number formats | SSF (SheetJS number format library) | Apache-2.0 | Excel format codes (`$#,##0.00`, `0%`, `yyyy-mm-dd`), the same in Node and the browser |
| xlsx import and export | ExcelJS 4.4 | MIT | Reads and writes formulas with their cached results, styles, number formats, freeze panes, column widths and notes. SheetJS CE (Apache-2.0) was the other option, but its free build does not write styles. |
| Unzip for the import report | fflate | MIT | Counts the charts, pivots, images and macros inside an xlsx that ExcelJS does not read, so the import report can name them |
| Charts | Our own SVG charts in kit colours | ours, AGPL | Small, themeable, no canvas library |

## 2. How a change flows

Every change is a tool call (`sheets.write_range`, `sheets.sort`, ...), from a person's screen or an agent alike. The server applies it to the workbook's Yjs document, stores the Yjs update in `sheets_updates`, and sends it to every open screen. Screens apply it and work the formulas out locally with the same engine. A typed value shows at once (the screen sets it in its own engine) and is confirmed when the server's update arrives.

Moves that shift cells (insert or delete rows and columns) are done by HyperFormula on the server, which rewrites every reference; the resulting cell contents are written into the document. Sorting moves whole rows and keeps each moved formula's relative references, as Excel and Google Sheets do.

## 3. Sign-in

`AUTH_PROVIDER=waronsaas|github|local`:

| Value | How people sign in |
|---|---|
| `waronsaas` | The shared warOnSaaS account (OpenID Connect at `ACCOUNT_URL`, default `https://account.waronsaas.com`), checked against its `/jwks.json` |
| `github` | GitHub, plus an email link |
| `local` | An email link only |

When `AUTH_PROVIDER` is not set: `github` if `GITHUB_OAUTH_CLIENT_ID` is set, otherwise `local`. Agents connect to `/mcp` with standard MCP OAuth (discovery, dynamic registration, PKCE), and the person signs in the same way.

## 4. Look freely, sign in to use

Anyone can open the example workbooks read-only, through their view-only share links. Editing, making a workbook and connecting an AI need sign-in. On the hosted demo (`SHEETS_OPEN_SIGNUP=1`) a new person gets their own workspace with copies of the examples. A self-hosted server is one team: the first person to sign in owns it and adds the others.
