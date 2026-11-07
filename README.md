# Meezan Investment Tracker (MIF)

Maria & Maeda – Al Meezan Islamic Fund. Web app (`index.html`) on GitHub Pages, data stored online in a Google Sheet through a small Apps Script (`Code.gs`).

Rule: **Friday's rate is applied automatically to Saturday and Sunday.**

## One-time setup (about 5 minutes)
1. Open the Google Sheet "Maria-Meezan Investment Tracker (MIF) v2" → **Extensions → Apps Script**.
2. Delete the sample code, paste all of `Code.gs`, click **Save**.
3. Choose the function **setup** → **Run** → approve the permissions (this renames the sheet tab to "Rates", creates "Plans", adds 18 & 19 Jul weekend rows).
4. **Deploy → New deployment → type: Web app** → Execute as: **Me** → Who has access: **Anyone** → Deploy → copy the **Web app URL** (ends in `/exec`).
5. On GitHub open `index.html` → pencil (edit) → find `const API_URL='';` → paste the URL between the quotes → Commit.
6. Repo **Settings → Pages** → Source `main` / root → Save → open https://haseebmeerza.github.io/meezaninvestment/
   The banner shows "● Online – Google Sheet" when linked.

## How it works
- Add/delete rates and plans in the app → saved straight into the Sheet (any device sees them).
- Edit the Sheet directly (Rates tab: columns A Date, B NAV) → app picks it up on Refresh/Update (and every 60 s).
- Update = pull latest · Refresh = reload view · Reset = restore original Excel data (also in the Sheet).

## Note
Anyone who has the web-app URL can change the data (no password, as requested). Keep the URL private if that matters.
Forecasts are statistical projections of past NAV only, not financial advice.
