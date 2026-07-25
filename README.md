# Buki Tracker

Daily tracker for Buki clients, Naulo Stick and Cig transactions, dues, payment methods, and end-of-day stock carry-forward.

## Run locally

```bash
npm install
npm run dev
```

## Save And Export

Use **Save Excel** to sync the current local records into the configured Google Sheet. The Google Sheet script writes into one tab named `Buki Records` and updates existing rows by a stable unique key.

Use **Export Excel** or **Download** to download an `.xlsx` file from the selected date, client, payment, and record-type filters.

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

The script writes all cloud records into one tab named `Buki Records`. It may remove extra blank tabs, but it does not delete non-empty old tabs.
