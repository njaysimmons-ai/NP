# NP Appointment Requests

A simple appointment-request page for a nurse practitioner. It runs entirely inside **your Google account**:

- **Patients** open a link, see the open times, pick one, and enter their name, contact info, and what they need. They don't need an account or password.
- **You (the NP)** manage everything in a **Google Sheet**: add open times, see requests, and mark them Confirmed or Declined. You sign in with the Google account you already use. There are no extra usernames or passwords.

There's no server to run and nothing to install. The patient page is a Google Apps Script web app attached to the Sheet, and both live in your Google Drive.

```
Patient ──link──▶  Request page (Apps Script, runs as you)
                         │  picks a time + fills form
                         ▼
               Google Sheet in your Drive
               ├── Availability  ← you add open times here
               └── Requests      ← new requests appear here (Pending)
```

---

## ⚠️ Read first: HIPAA requirements

The information patients type here (name + reason for visit) is **protected health information (PHI)**. This app is built to protect it, but no app is "HIPAA compliant" on its own. Compliance depends on the account it runs in and how you use it.

**Before you take real patient requests, you must have:**

1. **A paid Google Workspace account** (Business Starter or higher), **not** a free personal `@gmail.com` account. Google doesn't sign a HIPAA agreement for free Gmail accounts, so you **cannot** store PHI in them.
2. **A signed Business Associate Agreement (BAA) with Google.** A Workspace admin accepts it in the Admin console under *Account → Account settings → Legal and compliance*.
3. **Confirmation that the Google services used here are covered by that BAA.** This app uses Sheets, Drive, Apps Script, and (optionally) Gmail. Check Google's current list of covered services: <https://support.google.com/a/answer/3407054>. If a service isn't listed, don't use it for PHI.

If you only have a personal Gmail account today, you can still set this up and **test it with made-up data**. Upgrade to Workspace and sign the BAA before real patients use it.

---

## Setup (about 15 minutes, one time)

### 1. Create the Sheet
1. Go to <https://sheets.new> while signed in to your **Workspace** account.
2. Name it something like **"Appointment Requests (PHI – do not share)"**.

### 2. Add the code
1. In the Sheet, open **Extensions → Apps Script**.
2. Delete what's in `Code.gs` and paste in the contents of [`apps-script/Code.gs`](apps-script/Code.gs).
3. Near the top, edit `CONFIG` with your practice name, your name, and phone number. To get an email alert for each new request, set `NOTIFY_EMAIL`. The alert never includes patient details, only a link to the Sheet.
4. Click **＋ (Add a file) → HTML**, name it `Index` (no `.html`), and paste in [`apps-script/Index.html`](apps-script/Index.html).
5. Do the same again: **＋ → HTML**, name it `Sidebar`, and paste in [`apps-script/Sidebar.html`](apps-script/Sidebar.html).
6. Click the ⚙️ **Project Settings** gear and check **"Show appsscript.json manifest file in editor"**. Back in the editor, replace `appsscript.json` with [`apps-script/appsscript.json`](apps-script/appsscript.json). Change `"timeZone"` to your time zone if you're not in US Eastern.
7. Click 💾 **Save**.

### 3. Set up the tabs
1. In the Apps Script editor, choose the `setup` function in the dropdown and click **▶ Run**.
2. Google will ask you to authorize the script. Review it and allow it. The script only accesses this spreadsheet and sends the optional alert email.
3. Go back to the Sheet and reload it. You'll see a new **Appointments** menu, plus the **Availability** and **Requests** tabs.
4. Make sure the **Sheet's** time zone matches yours: **File → Settings → Time zone**.

### 4. Publish the patient page
1. In the Apps Script editor, click **Deploy → New deployment → ⚙️ → Web app**.
2. Set **Execute as:** *Me* and **Who has access:** *Anyone*. Patients don't have Google accounts for your organization, so they need this access level. They can only see open times and submit the form. They can never see the Sheet.
3. Click **Deploy** and copy the **Web app URL**. That's the link you give patients. You can put it on your website, in your email signature, or turn it into a QR code.

> If "Anyone" isn't offered, your Workspace admin has restricted sharing outside your organization. Ask them to allow it for your account. It's in the Admin console's Drive sharing settings.

> **When you change the code later:** use **Deploy → Manage deployments → ✏️ Edit → Version: New version**. That keeps the same patient link.

---

## Day-to-day use

**Add open times:** **Appointments → Add open time slots…** opens a panel. Pick a date, a time range, and the appointment length. You can repeat it weekly, for example "Tuesdays 9–12, 30-minute visits, 8 weeks". You can also type rows directly into the **Availability** tab. Put the date in column B and times like `2:00 PM` in columns C and D. An ID fills in automatically.

**Remove an open time:** delete its row in **Availability**. Don't delete rows that have a Request ID in column E. Decline or cancel that request instead.

**Handle requests:** new requests show up in the **Requests** tab with status **Pending**. Contact the patient, then change the status:

| Status | What happens |
|---|---|
| Confirmed | The time stays booked. |
| Declined / Cancelled | The time **automatically reopens** for other patients. |
| Completed | Visit happened. The time stays booked. |

Patients get a reference number such as `REQ-7F3A9C21` when they submit. Use it to find their row if they call. Don't change a Declined/Cancelled request back to Pending/Confirmed: its time may already have been given to someone else.

**Clean up:** **Appointments → Remove past unbooked slots** clears out old open times.

---

## Built-in protections

| Safeguard | How |
|---|---|
| Access control | Only Google accounts you share the Sheet with can see requests. Patients can only submit; they can't read anything back. |
| No new passwords | You sign in with Google, so you get your organization's 2-Step Verification and sign-in policies. |
| Minimum necessary | Only collects name, email/phone, contact preference, and reason. The form tells patients not to enter SSNs or insurance numbers. |
| Encryption | Google encrypts data in transit (HTTPS) and at rest. |
| No PHI in email | The optional alert email only says "new request" and links to the Sheet. |
| No PHI on the patient's device | Nothing is saved in the patient's browser. The confirmation shows only the time and a reference number. |
| Double-booking prevention | A script lock makes sure two people can't book the same time. |
| Spreadsheet-injection protection | Text that starts with `=`, `+`, `-`, or `@` is stored as plain text, never run as a formula. |
| Spam / flooding | A hidden "honeypot" field blocks bots, and submissions are capped per hour (`MAX_REQUESTS_PER_HOUR`). |
| Error logging off | `exceptionLogging: NONE`, so error details aren't sent to logging. |
| Audit trail | Google Workspace keeps Drive audit logs (who opened or changed the Sheet) and version history. |

## Your compliance checklist

- [ ] Google Workspace (paid) with a **signed BAA**, and you've confirmed the services used here are covered.
- [ ] **2-Step Verification** turned on for every account that can open the Sheet.
- [ ] The Sheet is shared **only** with people who need it. Never use "Anyone with the link".
- [ ] Don't download or export the Sheet to personal computers or phones. If you must, use encrypted, managed devices.
- [ ] Devices used to open the Sheet have screen locks and are kept up to date.
- [ ] Decide how long to keep requests (your retention policy). Delete or archive old rows on that schedule.
- [ ] Note this tool in your **HIPAA risk assessment** and policies.
- [ ] Optional: in the Admin console, turn on Drive audit-log alerts for the Sheet.

## Alternative with no code

Google Calendar has a built-in **Appointment schedule** feature (Calendar → Create → Appointment schedule). Patients pick a time from a booking page, and paid Workspace plans can add custom questions such as "reason for visit". It's covered by the same Workspace BAA. It's simpler, but you get less control over the form and how requests are approved. This app always makes the visit a *request* that you confirm.

---

## For developers

`npm test` runs the booking logic in `apps-script/Code.gs` against in-memory fakes of the Google services. You only need Node 20+, no Google account. The tests cover slot creation, booking, double-booking, declining and reopening slots, formula injection, the honeypot, and the rate limit.
