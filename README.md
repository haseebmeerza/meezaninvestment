# Meezan Investment Tracker – v3.0.0 (private)

## Update steps (10 minutes)
1. Apps Script: replace ALL code with `Code.gs` -> Save.
2. In `setupPasswords()` type the two real passwords (Maria's and your admin one), Run it ONCE, then change them back to the CHANGE-ME placeholders and Save.
3. Run `setup` once, then `installTrigger` once.
4. Deploy -> Manage deployments -> Edit -> New version -> Deploy (Execute as: Me, Access: Anyone). The /exec URL stays the same.
5. GitHub: replace `index.html` -> Commit. Open the site and log in with the password only.

## Roles
- Admin password: everything (including Reset and deleting rates).
- Maria password: view, add rates/plans/events, Check now.

## Security
All figures (investment, units, bank statement, portfolio no.) are stored only in Code.gs and are sent after login. The public page contains only the login box. Keep the Google Sheet sharing on "Restricted".
