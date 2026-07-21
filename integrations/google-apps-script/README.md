# Google Sheets + Calendar activation

The website integration is already built. Complete these account-level steps once to activate direct lead storage.

1. Create a Google Sheet for website leads and copy the spreadsheet ID from its URL.
2. Open **Extensions → Apps Script** from that Sheet.
3. Replace the default script with the contents of `Code.gs` in this folder.
4. In Apps Script, open **Project Settings → Script properties** and add:
   - `SPREADSHEET_ID`: the spreadsheet ID
   - `SHEET_NAME`: `Website Leads`
   - `NOTIFICATION_EMAIL`: `andrew@akrandall.com`
   - `CALENDAR_ID`: optional; leave out to use the default calendar
5. Run `installBookingSyncTrigger` once and approve the requested Sheet, Calendar, and email permissions.
6. Choose **Deploy → New deployment → Web app**.
7. Set **Execute as** to yourself and access to **Anyone**.
8. Copy the deployed URL ending in `/exec`.
9. Paste that URL into `appsScriptUrl` in the site-root `lead-config.js` file.
10. Submit both a standard inquiry and a completed free assessment. Confirm that each creates a row, sends the owner notification, and can be matched to a later calendar booking by email.

Do not place spreadsheet IDs, OAuth tokens, passwords, or other private credentials in browser-facing JavaScript. Only the public Apps Script web-app URL belongs in `lead-config.js`.
