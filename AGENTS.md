# AGENTS.md

## Project overview

This is a **Google Apps Script** project (`InvoicePdfToSheet.gs`) that parses Russian invoices (УПД / счета-фактуры) from PDF files in a Google Drive folder and writes structured data into a Google Spreadsheet. It is a single `.gs` file — there is no backend, no containers, and no build system.

The script runs entirely in Google's Apps Script cloud environment, bound to a Google Spreadsheet.

## Cursor Cloud specific instructions

### Development tooling

- **Linting:** `npm run lint` runs ESLint on all `.gs` files. The ESLint config (`.eslintrc.json`) recognizes Google Apps Script globals (`SpreadsheetApp`, `DriveApp`, `Logger`, etc.) via `eslint-plugin-googleappsscript`.
- **Syntax validation:** The `.gs` file is standard JavaScript (ES2020, script mode). Node.js can parse it for syntax checking.
- **clasp:** `@google/clasp` is installed for pushing/pulling code to Google Apps Script. Requires Google authentication (`npx clasp login`) and a `.clasp.json` config pointing to a script ID. The `.clasp.json` file is gitignored since it contains project-specific IDs.

### Key caveats

- **Cannot run locally:** The code depends on Google Apps Script built-in services (`SpreadsheetApp`, `DriveApp`, `DocumentApp`, `PropertiesService`, `UrlFetchApp`, `Utilities`, `Drive` advanced service). These are only available in the Apps Script runtime. There is no local emulator.
- **No automated tests:** The repository has no test infrastructure. Validation is limited to linting and syntax checking.
- **Entry points:** Functions `onOpen`, `runProcessFolder`, `runProcessFolderForSpreadsheet`, and `showRecognitionSetupHelp` are called by the Apps Script runtime (triggers, menu items). They appear unused in static analysis but are real entry points.
- **Optional APIs:** The script optionally calls Gemini (`GEMINI_API_KEY`) and OCR.space (`OCR_SPACE_API_KEY`) for scanned PDF text extraction. These keys are stored in Apps Script's Script Properties, not in environment variables.

### Available npm scripts

| Script | Command | Description |
|--------|---------|-------------|
| `lint` | `npm run lint` | ESLint all `.gs` files |
| `lint:fix` | `npm run lint:fix` | ESLint with auto-fix |
| `push` | `npm run push` | Push code to Apps Script (requires clasp auth) |
| `pull` | `npm run pull` | Pull code from Apps Script (requires clasp auth) |
