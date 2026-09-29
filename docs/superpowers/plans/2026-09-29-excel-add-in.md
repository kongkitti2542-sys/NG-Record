# NG Record Excel Add-in Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a desktop Excel task pane that appends NG records to the open workbook and reads the workbook table for history and exports.

**Architecture:** Keep the current single-page app and IndexedDB standalone mode. Add a small Office.js adapter selected only when the app runs inside Excel; it creates and appends to `NGRecords` on `NG_Record`, and serves history from that table. Add a minimal local HTTPS server and XML manifest for development sideloading.

**Tech Stack:** Existing HTML/CSS/JavaScript, Office.js Excel JavaScript API (ExcelApi 1.4), Node.js built-in HTTPS/static serving, Windows PowerShell for local certificate setup, Node.js built-in test runner, and bundled Playwright/Chrome for browser smoke checks.

**Spec:** `docs/superpowers/specs/2026-09-29-excel-add-in-design.md`

## Global Constraints

- No database, account system, cloud architecture, or ERP connection.
- Add-in workbook table name: `NGRecords`; worksheet name: `NG_Record`.
- Table columns, in order: `Record Date`, `Record Time`, `Part No.`, `Job No.`, `Cause`, `NG Qty`.
- Cause remains optional; Part No., Job No., Record Date, and positive integer NG Qty are required.
- Existing IndexedDB records are not migrated; standalone browser mode continues using IndexedDB.
- Support desktop Excel on Windows only in this implementation; do not sideload or install certificates automatically.
- If `NGRecords` already exists with incompatible headers, fail with an actionable message and do not overwrite it.

## Review Focus

- Workbook contains a table named `NGRecords` with incompatible headers: reject writes without altering the table.
- Workbook has no `NG_Record` sheet or table: create the sheet and correctly headed table on first save.
- Table has headers but no data rows: History must show zero rows and zero total without failing.
- Part No. or Job No. contains leading zeroes or text beginning with `=`: preserve it as text in Excel.
- Cause contains quotes, ampersands, or Thai text: History and CSV/XLSX exports preserve the text.
- Office.js is unavailable or rejects a workbook operation: do not show a success message or silently save only to IndexedDB.

---

### Task 1: Local Add-in hosting and manifest

**Files:**
- Create: `manifest.xml`
- Create: `server.cjs`
- Create: `package.json`
- Create: `setup-addin.ps1`
- Create: `README-ADDIN.md`
- Create: `.gitignore` (ignore local certificate/key artifacts)
- Test: `tests/server.test.cjs`

**Interfaces:**
- `server.cjs` exports `createRequestHandler(root)` for tests and serves workspace files over HTTPS at `https://localhost:3000/` using `localhost.pfx` and `NG_ADDIN_CERT_PASSWORD` when run directly.
- `package.json` defines `npm start` and `npm test` scripts; Node.js is a local development prerequisite.
- `manifest.xml` loads `https://localhost:3000/index.html` as an Excel task pane and requests `ReadWriteDocument` permission.
- `setup-addin.ps1` creates/trusts a current-user localhost development certificate and exports its PFX locally; it never sideloads the manifest.

- [x] **Step 1: Write failing server tests**
  - Add Node `node:test` checks that the server serves `/index.html`, returns 404 for an unknown path, and rejects traversal paths.
- [x] **Step 2: Run tests and confirm the expected failures**
  - Run: `node --test tests/server.test.cjs`
  - Expected: FAIL because the server module does not exist.
- [x] **Step 3: Implement the HTTPS host and manifest**
  - Serve only files under the project root, set HTML/JS/XML MIME types, and bind to `localhost:3000`.
  - Set a stable GUID in the XML manifest and the Excel task pane permission to `ReadWriteDocument`.
  - Add setup instructions for certificate setup, starting the server, and uploading `manifest.xml` through Excel's Add-ins > Upload My Add-in flow.
- [x] **Step 4: Verify hosting and manifest**
  - Run: `node --test tests/server.test.cjs`
  - Expected: all server tests pass; XML parses; HTTPS returns the task pane with status 200 when a local certificate is configured.
- [x] **Step 5: Commit**
  - Commit message: `feat: add local Excel add-in manifest and host`.

### Task 2: Workbook adapter

**Files:**
- Create: `excel-adapter.js`
- Create: `tests/excel-adapter.test.cjs`

**Interfaces:**
- `NGExcel.isAvailable()` returns `true` only when Office reports the Excel host and `Excel.run` exists.
- `NGExcel.appendRecord(record)` accepts `{date, part, job, cause, qty, time}` and resolves after the row is committed.
- `NGExcel.readRecords()` resolves to records with `{date, part, job, cause, qty, time}`.
- The task pane calls `Office.onReady()` before checking `NGExcel.isAvailable()`.
- Adapter functions use Office.js `Excel.run`; they never fall back to IndexedDB.

- [x] **Step 1: Write failing adapter tests**
  - Test `isAvailable()` for Excel host, non-Excel host, and missing Office.js.
  - Mock workbook worksheet/table collections and test creation of the worksheet/table, appending one row, reading rows, empty table, and incompatible headers.
  - Test that `Part No.` and `Job No.` values with leading zeroes or a leading `=` remain text and are never written as formulas.
- [x] **Step 2: Run tests and confirm the expected failures**
  - Run: `node --test tests/excel-adapter.test.cjs`
  - Expected: FAIL because `excel-adapter.js` does not exist.
- [x] **Step 3: Implement the Office.js adapter**
  - Create `NG_Record` and `NGRecords` only when absent; set table headers in the exact spec order.
  - Compare existing table headers to the expected six headers before reading or writing; reject mismatches without editing the table.
  - Set Part No., Job No., and Cause table columns to text format before appending so barcodes with leading zeroes or formula-like prefixes remain literal text.
  - Append a single row and read table body values, handling the table with header only as an empty list.
- [x] **Step 4: Run adapter tests**
  - Run: `node --test tests/excel-adapter.test.cjs`
  - Expected: all host detection, create, append, read, empty-table, and mismatch tests pass.
- [x] **Step 5: Commit**
  - Commit message: `feat: add Excel workbook record adapter`.

### Task 3: Task pane integration

**Files:**
- Modify: `index.html`
- Create: `app.js` (extract the existing inline browser application script)
- Create: `record-store.js`
- Modify: `excel-adapter.js` only if Task 2 tests expose a contract gap
- Test: `tests/app-mode.test.cjs`

**Interfaces:**
- `index.html` loads `excel-adapter.js`, `record-store.js`, and `app.js`; `app.js` loads Microsoft's Office.js library only for the manifest URL with `?addin=1`.
- `NGRecordStore.create({excel, isExcelHost, indexedDb})` returns `{mode, add(record), all()}`; `mode` is `"excel"` or `"browser"`.
- Excel host mode delegates to `NGExcel`; standalone mode uses the existing IndexedDB operations. An Excel write error is propagated and never retried into IndexedDB.

- [x] **Step 1: Write failing mode tests**
  - Verify Excel mode delegates `add` and `all` only to `NGExcel`; browser mode delegates only to IndexedDB.
  - Verify Excel adapter failures propagate and do not trigger a browser-store write.
- [x] **Step 2: Run tests and confirm the expected failures**
  - Run: `node --test tests/app-mode.test.cjs`
  - Expected: FAIL because app mode integration is not implemented.
- [x] **Step 3: Wire the task pane to the adapter**
  - Show a small connected-to-workbook status in Excel and a browser-local status outside Excel.
  - Wait for `Office.onReady()` before selecting Excel or standalone mode.
  - Append date, local time, Part No., Job No., Cause, and NG Qty to the workbook table on Save; disable Save during the write and report success only after completion.
  - Read Excel table records for history/filter/summary/CSV/XLSX; continue using IndexedDB outside Excel.
  - Preserve the existing record form, filters, cause field, and filtered export behavior.
- [x] **Step 4: Run mode tests**
  - Run: `node --test tests/app-mode.test.cjs`
  - Expected: Excel and standalone data paths pass without crossing stores.
- [ ] **Step 5: Commit**
  - Commit message: `feat: connect NG Record task pane to Excel`.

### Task 4: End-to-end verification and setup guide

**Files:**
- Modify: `README-ADDIN.md`

**Interfaces:**
- The smoke test runs the task pane in Chrome with an Office.js/Excel mock; it does not claim a real Excel client was exercised.
- Manual Excel acceptance steps remain documented for the user to run after sideloading.

- [x] **Step 1: Exercise the add-in path in Chrome with a mocked Excel host**
  - Load the task pane with a mocked `Office.onReady`/`Excel.run`, save Part `PART-TEST`, Job `JOB-TEST`, Cause `เครื่องเสีย`, Qty `10`, reload using the mock workbook's persistent table state, filter to the date, verify one row and total 10, export XLSX, and inspect the sheet XML for every value.
- [x] **Step 2: Exercise standalone mode in Chrome**
  - Save a sample record outside Office, reload, and confirm the record remains in IndexedDB without an Excel-connected status.
- [x] **Step 3: Complete installation and manual verification instructions**
  - Document trusted localhost certificate setup, server startup, manifest upload, Excel restart if needed, and uninstall steps.
  - Explain that records are stored in the active workbook, while standalone IndexedDB records remain separate.
- [x] **Step 4: Run unit checks and inspect browser errors**
  - Run: `npm test`
  - Expected: all unit tests pass; both browser flows have no JavaScript/page errors.
- [ ] **Step 5: Commit**
  - Commit message: `test: verify Excel Add-in record and export flow`.

## Manual Excel acceptance

After local certificate setup, start the HTTPS server, sideload the XML manifest, and open the task pane in desktop Excel. Save a sample row, confirm it appears in `NG_Record` > `NGRecords`, save and reopen the workbook, filter History to the record date, verify NG Qty total 10, and export XLSX. The automated mock-browser test does not replace this real-Excel acceptance check.
