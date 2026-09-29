# NG Record Excel Add-in Design

## Goal

Let an operator enter NG records from a task pane inside Excel and write each saved record directly into the workbook that is currently open. Keep the existing small scope: record, search history, and export.

## Approved direction

Build an Excel task pane add-in using Office.js. On Windows Excel, the add-in uses the open workbook as the source of truth. When opened as a normal web page outside Office, the existing app continues using browser IndexedDB; these two stores remain separate.

No server database, account system, cloud architecture, or ERP connection is added.

## Workbook data layout

On first save, create a worksheet named `NG_Record` if it does not exist, then create a table named `NGRecords` with these columns:

1. Record Date
2. Record Time
3. Part No.
4. Job No.
5. Cause
6. NG Qty

Each successful save appends one row to this table. The record time is generated automatically. Existing browser IndexedDB records are not migrated automatically; they remain available in the standalone browser app.

The History view reads the current workbook table and applies the existing date, Part No., and Job No. filters. It calculates row count and NG Qty total from the filtered rows. CSV and XLSX exports contain the same filtered rows, including Cause.

## Add-in and local setup

Add an Office Add-in XML manifest, a task pane entry point, and a small local HTTPS static server. The manifest points to the local task pane URL for development sideloading. Include setup instructions for trusting a localhost development certificate and uploading the manifest into Excel. Do not install certificates or sideload the add-in automatically.

The add-in requires a supported desktop Excel version with the required Office.js Excel APIs and access to the local HTTPS server. Excel Online is outside this first implementation.

## User experience

Keep the current Thai, mobile-sized entry form and history layout. In Excel, show that the current workbook is connected. Save writes only to `NGRecords`; History and exports read from that table. Outside Excel, retain the current IndexedDB mode and label it as browser-local storage.

## Failure handling

- Validate required Part No., Job No., Record Date, and positive integer NG Qty before calling Office.js.
- Disable Save while a workbook write is in progress to prevent duplicate clicks.
- Show a success message only after the Excel table append succeeds.
- If the workbook is unavailable, a sheet/table cannot be accessed, or the write fails, show an actionable error and do not report success.
- Cause remains optional, matching the existing form.

## Acceptance checks

1. Open the manifest-backed task pane in desktop Excel and verify it identifies the workbook connection.
2. Enter a sample Part No., Job No., Cause, NG Qty 10, and a chosen Record Date; save and confirm the row appears in the `NG_Record` worksheet table.
3. Close and reopen that workbook; confirm the row remains.
4. Open History, filter to the saved date, and confirm the record and total of 10 appear.
5. Export the filtered history to XLSX and confirm the record and Cause are present.
6. Confirm standalone browser mode still saves to IndexedDB and has no false Excel-connected status.

## Constraints and risks

- Local development sideloading requires a trusted HTTPS localhost certificate and an Excel installation that permits sideloading. This workspace can prepare the manifest and setup steps but does not grant organizational permission to install the add-in.
- Workbook data is available only in that workbook. Separate workbooks do not share records, which is consistent with the selected direct-to-Excel design.
- The workbook must not already contain an incompatible table named `NGRecords`; show an error rather than overwrite unrelated data.
