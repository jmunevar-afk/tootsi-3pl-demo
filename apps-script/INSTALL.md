# PO Hub · Google Workspace version · install guide

Free pilot. It runs on Tootsi's Google Workspace account. No new software, no monthly bill.

- **Database:** one Google Sheet (one tab per table).
- **Files:** one Drive folder ("PO Hub files"), one sub-folder per PO.
- **App:** Google Apps Script web app, attached to the Sheet.
- **Emails:** sent from the account that deploys the app.

Time to install: about 30 minutes.

---

## 1. Create the Sheet and paste the code

1. In Google Drive, create a new Google Sheet. Name it **PO Hub (pilot)**.
2. In the Sheet, open **Extensions › Apps Script**.
3. Delete the empty `Code.gs` content.
4. Create each file from this folder with the same name:
   - Script files (`+` › Script): `Code`, `Store`, `Logic`, `Api`, `Setup`, `Jobs`
   - HTML files (`+` › HTML): `Hub`, `Portal`, `Receive`, `Client`, `Styles`, `Lang`, `Logo`, `Message`
5. Open **Project settings** (gear icon). Tick **Show "appsscript.json" manifest file**. Paste the content of `appsscript.json`.
6. Click **Save**.

Faster option for developers: use `clasp push` from this folder.

## 2. Run the setup

1. In the Apps Script editor, pick **setup** in the function menu. Click **Run**.
2. Google asks for permissions. Accept them. The app needs Sheets, Drive, email and triggers.
3. Back in the Sheet, you now have the tabs: POs, Submissions, Issues, Notes, Delays, Recounts, Docs, SupplierQty, Users, Config, Log.
4. In the Config tab, paste the ID of the "UPCOMING FTL FCL" Google Sheet in `UPCOMING_SHEET_ID` (the long code in its URL, between `/d/` and `/edit`).
5. Optional, for a test run: run **loadSampleData**.

## 3. Deploy the two links

Both deployments use the same code. Only "who has access" changes.

**Internal hub (Tootsi staff)**
1. **Deploy › New deployment › Web app**.
2. Execute as: **Me**. Who has access: **Anyone within Tootsi Impex**.
3. Copy the URL. Paste it in the Config tab, key `HUB_URL`.

**3PL portal (TP and DL)**
1. **Deploy › New deployment › Web app** again.
2. Execute as: **Me**. Who has access: **Anyone**.
3. Copy the URL. Paste it in the Config tab, key `PORTAL_URL`.

Anyone opening the portal URL without signing in sees only the sign-in page. The hub only opens for a signed-in @tootsi.com Google account.

## 4. Create the 3PL logins

Run once per 3PL in the editor (edit the values first):

```
addThreePL('TP', 'Trans-Plus', 'receiving@transplus.example', 'transplus')
addThreePL('DL', 'Deslauriers', 'receiving@deslauriers.example', 'deslauriers')
```

The execution log shows the portal link, the username and a temporary password. Give the password by phone, not by email.

How sign-in works:
- **3PLs** open the portal link and sign in with username and password. They stay signed in for one shift (12 h, `SESSION_HOURS` in Config). Then they sign in again.
- **5 wrong passwords** lock that username for 15 minutes.
- **Tootsi staff** don't get a password. Google already knows who they are. Any @tootsi.com account opening either link goes straight to the hub. To see the 3PL screen yourself, add `?view=3pl` to the portal link.
- Passwords are never stored. The Users tab keeps only a salted hash. Hide the Users and Sessions tabs anyway (right-click › Hide sheet).

Managing logins:
- Forgot password or a receiver left: `resetThreePLPassword('transplus')`. It signs out every open session.
- Stop a 3PL completely: `disableThreePL('transplus')`.

## 5. Turn on the automatic jobs

Run **installTriggers** once. It adds:

- **importUpcoming** every 2 hours. It reads the "UPCOMING FTL FCL" sheet and adds new POs. It never overwrites what people edited in the hub.
- **dailyJob** every morning at 7:40. It imports, updates the colors, nudges late 3PL POs and emails the daily digest.

## 6. Emails (start safe)

All emails are off at first. In the Config tab:

| Key | Start with | Meaning |
|---|---|---|
| `EMAIL_ON` | `yes` | Turn alerts on |
| `PILOT_ONLY_ME` | `yes` | Every alert goes only to `FALLBACK_EMAIL` (you) during the test |
| `EMAIL_3PL` | `no` | No emails to the 3PLs yet |
| `TEAM_Purchasing`, `TEAM_Logistics`, … | emails | Who gets the issues each team owns |

When the pilot looks right, set `PILOT_ONLY_ME` to `no`, then `EMAIL_3PL` to `yes`.

## 7. Daily use

- **Expected cases:** the upcoming sheet has no cases column. Open the hub, filter **More filters › Missing expected cases**, and fill them in. The 3PL can't be checked against an empty number.
- **POs not in the sheet** (LTL, small orders): use **＋ Add upcoming PO**.
- **Holidays:** list them in Config, key `HOLIDAYS` (yyyy-mm-dd, comma-separated). The late clock skips them.

## Limits to know

| Limit | Value | Is it a problem? |
|---|---|---|
| Google Sheet size | 10 million cells | Years of POs |
| Drive storage | Tootsi's Workspace quota | Photos are compressed to about 300 KB each |
| Emails per day | 1,500 (Workspace) | Far above need |
| Script run time | 6 minutes per call | Each action takes 1 to 3 seconds |
| Speed | 1 to 3 s per save | Slower than a paid app, fine for a pilot |

## Privacy rules built in

- The 3PL page only receives an allow-list of fields: PO, item, description, date, container, pallets, load type, 3PL instructions, colour level. No supplier, price, sales order, allocation, notes or Tootsi documents.
- Each 3PL signs in with its own username and password, and only sees its own POs. The server checks the session on every call.
- The count check only answers "matches" or "doesn't match". It never shows the expected quantity.
- Files go to Tootsi's Drive. The 3PL never gets a Drive link.
