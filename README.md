# Buki Tracker

Daily tracker for Buki clients, Naulo Stick and Cig transactions, dues, payment methods, and end-of-day stock carry-forward.

## Run locally

```bash
npm install
npm run dev
```

## Export/Import

Use **Load Excel** to load the workbook you use for daily records. After a workbook is loaded, **Save Excel** downloads an updated copy using that same filename. **Save As New** downloads a date-stamped workbook such as `buki_tracker_2026-07-25.xlsx`.

Browsers do not allow a website to silently overwrite an arbitrary local Excel file for safety reasons, so saving still happens through the browser's download flow. If your browser asks what to do with the file, choose replace/overwrite to keep one workbook.

The app also saves automatically in the current browser with `localStorage`.

## Google Sheets sync

1. Create a Google Sheet.
2. Open **Extensions > Apps Script**.
3. Paste the code from `google-apps-script.js`.
4. Deploy as a web app.
5. Set access to the intended users.
6. Copy the web app URL.
7. In Vercel, add:

```bash
VITE_GOOGLE_SHEET_WEBAPP_URL=https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
```

Then redeploy the app.

The script writes all cloud records into one tab named `Buki Records`.
