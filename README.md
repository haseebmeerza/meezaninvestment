# Meezan Investment Tracker – v3.2.1 (private)

## Update steps (5 minutes)
1. Apps Script: replace ALL code with `Code.gs` -> Save.
2. Run `setup` once (first time only), then `installTrigger` once (first time only).
3. Deploy -> Manage deployments -> Edit -> New version -> Deploy (Execute as: Me, Access: Anyone). The /exec URL stays the same.
4. GitHub: replace `index.html` -> Commit. Open the site and log in with the password.

## Password
Login password: `maria123` (set at the top of `Code.gs`: `USER_PW` / `ADMIN_PW`). To change it, edit those two words, Save, and Deploy a New version. Nothing else to run.
If you give `ADMIN_PW` a different password, that one can also Reset and delete rates; the other can view and add data only.
