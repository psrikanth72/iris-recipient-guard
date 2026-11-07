# IRIS Recipient Guard (v1.0.0)

An Outlook add-in that checks every email when **Send** is clicked: who it goes to, which client it's for, and whether attachments are safe to send. It's a separate project from Send Check, with its own add-in ID, so the two don't interfere with each other.

## What it does

| Situation | Rule |
|---|---|
| Internal addresses | Addresses ending in `@iriskpo.in` or `@iris.co.uk` are always allowed. |
| New email | Typed external addresses are removed. Start the subject with the client name (e.g. `Sample Traders Ltd - March payroll`) and that client's contacts are added to To. Change the name and the contacts swap; remove it and they disappear. |
| Reply / Reply all | Only contacts of the thread's client can stay or be added. Anyone not in the client list, or belonging to another client, is removed and blocked. |
| Forward | Internal addresses only. |
| One client per email | Recipients from two clients: blocked. |
| Unknown or inactive contacts | Blocked. |
| Attachments to clients | Every file name must contain the client name, and every file must be password protected (ZIP, 7z, PDF, Word, Excel, PowerPoint). Other file types must go inside a password-protected ZIP. Shared links and attached emails are blocked. Images in the body (e.g. signature logos) are skipped. |
| Fail closed | If the add-in can't load, Outlook won't send (manifest `SendMode="Block"`). If the client list can't be downloaded, emails to clients are blocked; internal emails still go. |
| Status pane | **Recipient Guard** button on the ribbon: protection status, client-list connection, client picker, and "Check this email now". |

Matching is by the **full email address**, not the domain, so shared domains like gmail.com are safe. Company names in subjects and file names ignore case, spaces and punctuation (`SampleTradersLtd_March.zip` matches `Sample Traders Ltd`).

## Files

| Folder | Goes to | Contents |
|---|---|---|
| `iris-recipient-guard/` | GitHub repository **psrikanth72/iris-recipient-guard** (Pages on) | Add-in code, manifest, icons, sample client list |
| `user-site/` | GitHub repository **psrikanth72/psrikanth72.github.io** (Pages on) | `.well-known/microsoft-officeaddins-allowed.json` – needed by classic Outlook (see below) |
| `it-server/` | IT – the network-drive folder | `clients.csv` (your real list, 116 clients / 756 contacts) and `web.config` for IIS |

**Never upload `clients.csv` to GitHub** – GitHub Pages sites are public.

## Step 1 – Put the code on GitHub

1. Create the repository `iris-recipient-guard` (Public) under psrikanth72 and upload everything in `iris-recipient-guard/`, including the `assets` folder and the empty `.nojekyll` file.
2. Settings → Pages → Source: **Deploy from a branch**, branch `main`, folder `/ (root)` → Save.
3. Create a second repository named exactly `psrikanth72.github.io` (Public), upload everything in `user-site/` (including the `.well-known` folder and `.nojekyll`), and turn Pages on the same way.
4. After a few minutes these should open in a browser:
   - `https://psrikanth72.github.io/iris-recipient-guard/guard.js`
   - `https://psrikanth72.github.io/.well-known/microsoft-officeaddins-allowed.json`

> Why step 3? Classic Outlook on Windows runs the send check in a restricted engine. Microsoft requires the add-in's website to publish this file at the root of its address before that engine may download data (the client list) from another server.

## Step 2 – IT: publish the client list from the network drive

The add-in can only read `https://` addresses, not `\\server\share` paths. IT keeps the file on the network drive and publishes that folder through an internal web server:

1. Copy `clients.csv` and `web.config` from `it-server/` into the network-drive folder.
2. In IIS, add a **Virtual Directory** (e.g. `/recipient-guard`) pointing at that folder, on a site with a valid HTTPS certificate.
3. `web.config` already sets the headers the add-in needs:
   - `Access-Control-Allow-Origin: https://psrikanth72.github.io`
   - no caching, so every save is picked up straight away.
4. Check `https://<server>/recipient-guard/clients.csv` opens from **every place staff use Outlook** (office, VPN, home). If it doesn't, emails to clients will be blocked for those users.
5. Edit `guard.js`, line `REPOSITORY_URL:` – replace the sample address with the real one, upload `guard.js` to GitHub again.

While the sample address is in place, the status pane shows a red **Sample client list in use** banner.

## Step 3 – IT: deploy through the Microsoft 365 admin center

1. admin.microsoft.com → Settings → **Integrated apps** → **Upload custom apps** → *Office Add-in* → upload `manifest.xml`.
2. Assign to the pilot users first, then everyone.
3. Recommended policies for classic Outlook (Microsoft's Smart Alerts guidance):
   - Configure the Group Policy setting **"Running Outlook for Simple MAPI Mail Sending"** as described in Microsoft's Smart Alerts article, so emails started from Word/Excel "Send as attachment" are also checked.
4. Remove the old Send Check add-in from the same users once the pilot is signed off, so they don't get two sets of prompts.

Changes to `guard.js` / `taskpane.html` go live without redeploying. Changes to `manifest.xml` need IT to click **Update** in the admin center.

## Keeping the client list up to date (real time)

- Open `clients.csv` **from the network drive** in Excel, edit, and press **Ctrl+S** → keep **CSV** format when Excel asks.
- Columns: `Company Name`, `Email ID`, `Status`. Status is `Active` or `Inactive` (blank = Active).
- Several addresses can go in one cell, separated by `;`.
- **Every Send downloads the latest file**, so a saved change applies to the very next email. (The as-you-type helpers reuse a copy for up to 60 seconds.)
- To stop a client: set all its rows to `Inactive`. The company then can't be picked, its contacts are blocked, and any thread with them is blocked.
- If the same address appears under two clients, it's treated as unknown (blocked) and the status pane shows a warning.

## Using it (for staff)

- **New email to a client:** type the client name at the start of the subject, or click **Recipient Guard → Pick company**. Contacts are added to To. Remove any you don't need.
- **Pin the pane** (pin icon at the top) so subject-based filling happens instantly. Without it, contacts are added when you type in To or when you click Send.
- **Attachments:** name the file with the client name and password-protect it. For 7-Zip, tick **Encrypt file names**.
- **Blocked?** The message lists exactly what to fix. Click **Open Recipient Guard** to see the details.

## Supported Outlook versions

Smart Alerts (the Send check) requires Mailbox requirement set 1.12:

- Outlook on the web, new Outlook for Windows – supported.
- Classic Outlook for Windows – recent Microsoft 365 builds (around version 2206 or later).
- New Outlook for Mac – supported.
- **Not covered:** Outlook mobile apps (iOS/Android) and older perpetual versions (Outlook 2016/2019). IT should block or restrict those for client email, or users could send without checks.
- Meeting invites are not checked (messages only).

## Limits worth knowing

- Outlook has no "subject changed" event, so subject-based filling happens through the pinned pane, on recipient changes, or at Send. The Send check always enforces the rules regardless.
- 7z archives are confirmed only when file names are encrypted.
- Password detection checks the file's encryption markers; it can't confirm the password is strong or was shared separately.
- Classic Outlook's restricted engine has reported network quirks on some builds. Test classic Outlook in the pilot first. If it can't reach the client list, the safe option is to host the add-in files on the same internal server as `clients.csv` (find/replace `https://psrikanth72.github.io/iris-recipient-guard/` in `manifest.xml`, and put the `.well-known` file at that server's root).

## Testing done

The rules were tested against a simulated Outlook (41 scenarios: new/reply/forward, company switching and removal, cross-client, unknown, inactive, multi-client, lookalike domains, unreachable client list, attachments) and real encrypted/unencrypted ZIP, AES-ZIP, 7z, PDF, DOCX, XLSX and XLS files. All 116 company names in your list are recognised from a subject with and without punctuation. Live Outlook testing in the pilot is still needed.

## Pilot test checklist

| # | Test | Expected |
|---|---|---|
| 1 | New email, type a client address with no subject | Address removed, message shown |
| 2 | Subject `Adams Beeny Limited - test` | Adams Beeny contacts appear in To (pane pinned) |
| 3 | Change subject to another client | Contacts swap |
| 4 | Delete the client name from subject | Client contacts removed |
| 5 | Add a contact of another client | Removed |
| 6 | Send to client with no attachment | Sent |
| 7 | Attach unprotected Excel named with client | Blocked: not password protected |
| 8 | Attach protected PDF without client name | Blocked: file name must contain… |
| 9 | Attach protected file named correctly | Sent |
| 10 | Reply to a client email, add a colleague | Sent |
| 11 | Reply, add another client's contact | Removed / blocked |
| 12 | Forward a client email to a client | Removed / blocked |
| 13 | Set a contact Inactive in clients.csv, save, send to them | Blocked |
| 14 | Status pane | Send check **Active**, client list **Connected** |
| 15 | Ask IT to stop the client-list site briefly, send to client | Blocked; internal email still sends |
| 16 | Repeat 1–15 in Outlook on the web and classic Outlook | Same results |
