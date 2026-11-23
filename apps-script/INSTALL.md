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

Anyone opening the portal URL without a valid private code sees "access denied". The hub page also checks for a @tootsi.com account.

## 4. Create the 3PL links

In the editor, run once per 3PL (edit the values first, or run from a small helper):

```
addThreePL('TP', 'Trans-Plus', 'receiving@transplus.example')
addThreePL('DL', 'Deslauriers', 'receiving@deslauriers.example')
```

The private link appears in the execution log and in the **Users** tab. Send it to the 3PL receiving lead. They bookmark it on the dock phone. No Google account needed.

To cut access later: set `active` to FALSE in the Users tab, or run `disableThreePL(token)`.

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
- Each 3PL only sees its own POs. The server checks the private code on every call.
- The count check only answers "matches" or "doesn't match". It never shows the expected quantity.
- Files go to Tootsi's Drive. The 3PL never gets a Drive link.
