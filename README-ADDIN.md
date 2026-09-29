# NG Record Excel Add-in (local development)

This add-in writes NG entries into the workbook currently open in **Excel Desktop for Windows**. The workbook is the persistent store. The standalone web page continues to store data in browser IndexedDB; the two stores are separate and are not migrated automatically.

## Prerequisites

- Excel Desktop for Windows with Office Add-ins enabled by the organization.
- Excel JavaScript API requirement set 1.4 or newer.
- Node.js 20 or newer.
- A network connection to load Microsoft's Office.js library.

## Start the local HTTPS host

Open PowerShell in this folder and dot-source the certificate setup script. It creates a local `localhost.pfx`, trusts its certificate for the current Windows user, and keeps its password only in the current PowerShell session. This setup is manual; it does not install or sideload the add-in.

```powershell
. .\setup-addin.ps1
npm start
```

Keep this PowerShell window open while using the add-in. Confirm that `https://localhost:3000/index.html` opens without a certificate warning before continuing. The manifest and certificate file are development-only; do not publish the self-signed certificate or PFX.

## Upload the manifest to Excel

1. Open a workbook in Excel Desktop.
2. Open the Office Add-ins management view and choose **Upload My Add-in**.
3. Select `manifest.xml` from this folder and add the NG Record task pane.
4. Launch NG Record from the Add-ins menu.

The exact Add-ins menu location can vary by Excel build and organization policy. If upload is disabled, ask the Microsoft 365 administrator whether sideloading Office Add-ins is permitted.

## Use and remove

Save creates the `NG_Record` worksheet and `NGRecords` table in the current workbook when needed. New records are appended there. Save the workbook to keep the rows. The History tab reads that table and exports its filtered rows.

Close Excel or remove NG Record from the Office Add-ins management view to stop using it. Stop the local host with Ctrl+C. To discard local development credentials, remove `localhost.pfx` and delete the corresponding NG Record certificate from the current user's Personal and Trusted Root Certification Authorities stores.

This is a local development sideload. It is not deployed for other computers; other users would need an HTTPS-hosted version and an approved organization deployment method.
