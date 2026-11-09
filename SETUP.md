# Setup

## 1. Backend (Google Apps Script)

1. Open your existing Apps Script project (or create a new one at script.google.com).
2. Replace **all** of the code in `Code.gs` with the `Code.gs` from this folder. Save.
3. **Project Settings** (gear icon) → set the time zone to **(GMT+05:30) India Standard Time**.
4. In the editor, choose `setup` from the function dropdown and click **Run**. Allow the permissions it asks for (Sheets, Calendar, Gmail send).
   The log prints the link to your bookings sheet.
5. **Deploy → New deployment** → type **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
   Copy the web app URL (ends in `/exec`).
6. **Deploy → Manage deployments** → archive the **old** deployment, so the old spammed URL stops working.

## 2. Website

1. In `index.html`, set:
   - `SCRIPT_URL` = the `/exec` URL from step 5
   - `UPI_ID` = your UPI ID (optional; adds an "Open UPI app" button on phones with ₹70 filled in)
2. Save your UPI QR image as `upi-qr.jpg` next to `index.html`.
3. Upload `index.html` and `upi-qr.jpg` to your GitHub repo (replace the old `index.html`).

## 3. Day to day

- Each booking request emails you. Check the UTR and ₹70 in your UPI app, open the **Review** link, press **Confirm** (adds it to Google Calendar) or **Reject** (frees the date).
- Requests never expire. They hold the date (shown as "pending") until you decide.
- Reject also works on confirmed bookings, as a cancel: it removes the calendar event.
- Don't forward the emails: the Review link works for anyone who has it.
- If the link ever leaks, delete `ADMIN_KEY` under Project Settings → Script properties and run `setup` again. Old links stop working.

## Settings you can change (top of Code.gs)

- `MAX_PENDING` (10): new requests are refused while this many are waiting for you.
- `MAX_DAYS_AHEAD` (60): how far ahead people can book. It's also set in `index.html`.
- `PICKUP_LOCATIONS`: must match the options in `index.html`.
