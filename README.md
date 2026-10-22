# Case Management Tool

A private client dashboard and intake app. Locally, it reads the `MasterData` sheet in the Excel workbook. On Vercel, it uses Google Sheets; the local workbook is excluded from Git.

## Local development

Requirements: Node.js 20 or newer.

```sh
npm ci
npm run dev
```

Open `http://127.0.0.1:5173`. Local development uses the Excel workbook and does not need Google credentials.

Excel mode expects `Prototype of Client's Tracker.xlsx` and `Prototype of Reports.xlsx` in the project folder. The client API reads `MasterData`; court dates come from the `COURT DATE LIST` tab in the reports workbook. Workbooks are ignored by Git. If your files have different paths, set `WORKBOOK_PATH` and `COURT_WORKBOOK_PATH` in the local environment.

## Publish to GitHub

Install Git for Windows and sign in to GitHub from VS Code. Create a **private** empty repository on GitHub, then run these commands from this folder:

```sh
git init
git add -A
git status --short
```

Check that `Prototype of Client's Tracker .xlsx`, `.env` files, and `node_modules` are not staged. Then commit and push, replacing the URL with your private repository URL:

```sh
git commit -m "Initial case management app"
git branch -M main
git remote add origin https://github.com/YOUR-ACCOUNT/YOUR-PRIVATE-REPO.git
git push -u origin main
```

The `.gitignore` excludes Excel workbooks, dependencies, build output, and environment files. Do not force-add the client workbook or service-account credentials.

## Prepare Google Sheets

1. Upload the client tracker to Google Drive and convert it to a Google Sheet. Keep the `MasterData` tab and its header row.
2. Upload `Prototype of Reports.xlsx` separately and convert it to a second Google Sheet. Rename the court tab to `COURT DATE LIST` (no trailing space) and keep its header row.
3. In Google Cloud, enable the Google Sheets API and create a service account. Share both Google Sheets with that account and grant Editor access.
4. Copy each spreadsheet ID from its Google Sheets URL. Keep the service-account key private.

## Deploy on Vercel

1. Import the private GitHub repository into Vercel. The included `vercel.json` configures the Vite build and serverless API.
2. Add these environment variables to the Vercel project for Production and Preview as appropriate:
	- `DATA_SOURCE` = `google-sheets`
	- `GOOGLE_SHEET_ID` = the spreadsheet ID
	- `GOOGLE_SHEET_NAME` = `MasterData`
	- `GOOGLE_SERVICE_ACCOUNT_EMAIL` = the service-account email
	- `GOOGLE_PRIVATE_KEY` = the service-account private key
	- `GOOGLE_COURT_SHEET_ID` = the court-date spreadsheet ID
	- `GOOGLE_COURT_SHEET_NAME` = `COURT DATE LIST`
	- `APP_PASSWORD` = a strong private workspace password
	- `SESSION_SECRET` = a separate random secret used to sign sessions
3. Redeploy after saving the environment variables. The app requires sign-in and refuses to serve the bundled Excel workbook in production.

Vercel redeploys from GitHub pushes. The app checks client and court data every 20 seconds and when a tab regains focus. New intakes append to `MasterData`; staff edits appear after the sheet is saved. Court dates are matched to clients by normalized name; a reschedule date takes precedence over the original court date. Staff also need their own Google Sheet permissions.

The login is a shared password for a private pilot, not individual employee accounts or an audit trail. The dashboard clock and follow-up calculations use `America/Chicago`, including daylight saving time.
