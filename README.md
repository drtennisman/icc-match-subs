# Match Sub App

A Progressive Web App for finding subs across multiple tennis teams. Captains post a
match, hand-pick who gets emailed, and the first sub to accept gets the spot.

## How it works

- **Players sign themselves up** — name, email, and **one question: their soft court
  level**. A confirmation email has to be clicked before they receive anything, which
  catches typo'd addresses before they cost anyone a match.
- **Captains post a match** and then choose exactly who to notify. The list shows
  everyone eligible, grouped by level with the strongest first, and within each level
  sorted by who has subbed least for that team — so the same reliable person doesn't
  get asked every single time.
- **First to accept wins.** Subs can accept straight from the email or in the app. The
  captain gets an email with the sub's name, email and phone.
- **If nobody bites**, the captain gets a nudge email after 24 hours with a reminder to
  widen the net.

## The two rules that decide who a captain sees

**You can sub up, never down.** A match at 3.5 can use anyone rated 3.5 or lower. A 4.0
player is not eligible and does not appear at all. This is why signup only needs one
question — eligibility falls out of a single number.

**Three matches per team, per season.** Once someone has subbed three times for Miller
4.0 they're locked for that team until the season changes, but stay fully available for
every other team. Maxed-out players are still shown in the picker, greyed out with a
reason, so a captain hunting for a specific name sees *why* they can't pick them rather
than finding them mysteriously absent.

The cap is enforced **on the server**, not just in the picker — the one-tap claim link
in a notification email bypasses the UI entirely, so the backend is the only place that
can actually stop it. Backing out of a match gives the slot back.

Seasons are **Spring and Fall only**. Rolling over is a one-cell edit: change `Season`
on the Config tab and every count reads as zero. Nothing parses the value, so any
consistent naming works so long as it changes between seasons.

## Architecture

| Piece | Where | Notes |
|-------|-------|-------|
| `index.html` | Vercel | The entire frontend — vanilla JS/CSS, no build step |
| `google-apps-script.js` | Google Apps Script web app | All reads/writes; deployed manually (see below) |
| Google Sheet | Google Drive | The database — 5 tabs, all self-installing |
| `sw.js` / `manifest.json` | Vercel | PWA install + offline shell |

## Google Sheet tabs

All five are created automatically the first time the script runs. **Teams** and
**Config** are the only ones filled in by hand.

| Tab | Headers |
|-----|---------|
| **Teams** | `TeamID, Team Name, Level, Captain, Captain Email, Active` |
| **Subs** | `SubID, Name, Email, Phone, Level, Verified, Token, Active, Added By, Signed Up At, Sub Count, Last Sub, Season, Team Subs` |
| **Requests** | `ID, TeamID, Level, Date, Time, Location, Opponent, Line, Notes, Posted By, Posted At, Notified, Status, Claimed By, Claimed At, Nudged` |
| **History** | `Request ID, Team, Match Date, Sub Name, Sub Email, Claimed At, Posted By` |
| **Config** | `Key, Value` |

### Filling in Teams

One row per team. Leave `TeamID` blank and the script fills it in for you. `Active` can
be `TRUE`/`FALSE` (blank counts as active).

| TeamID | Team Name | Level | Captain | Captain Email | Active |
|--------|-----------|-------|---------|---------------|--------|
| *(blank)* | ICC Ladies 4.0 — Thursday | 4.0 | Anne Powell | anne@example.com | TRUE |

The **Captain Email** is where claim notifications and no-sub nudges go, so it has to be
right or that captain hears nothing.

### Config keys

| Key | What it does |
|-----|--------------|
| `AppUrl` | The public app URL (the Vercel one). Used for links inside emails. |
| `CaptainPIN` | Shared PIN that unlocks posting and the sub list. **Change it from `1234`.** |
| `ManagerEmail` | Optional — CC'd on every claim confirmation. Leave blank to skip. |
| `NudgeHours` | How long before a captain gets nudged about an unclaimed match. Default `24`. |
| `Season` | Current season, e.g. `2026 Fall`. Changing it resets every sub's per-team count. |
| `MaxSubsPerTeam` | How many times one sub may play for one team per season. Default `3`. |

The `Team Subs` column on the Subs tab stores counts as `miller40:2, ray45:1` — readable
if you need to check or correct someone by hand. It's paired with `Season`, so a count
from a previous season is ignored rather than deleted.

## Why there's a PIN

The app URL is public, and the `Subs` tab holds members' personal emails and phone
numbers. Without a gate, anyone with the link could read all of it. So:

- **No PIN needed** to sign up as a sub or to accept a match.
- **PIN required** to post a match, see the sub list, or add someone manually.

Contact details are stripped from the API response entirely unless the PIN is supplied —
they aren't just hidden in the UI.

## Deploying changes

**Frontend** (`index.html`, `sw.js`): commit and push. Vercel deploys automatically. The
service worker is network-first for `index.html`, so a new version shows up on the next
load without any cache bumping.

**Backend** (`google-apps-script.js`):
1. Bump the `VERSION` constant at the top of the file **and** `APP_VERSION` in
   `index.html` (keep them equal).
2. Copy the whole file and paste it over the code in the Apps Script editor
   (Extensions → Apps Script from the Sheet).
3. Deploy → Manage deployments → ✏️ edit → **Version: New version** → Deploy.
   *Skipping "New version" keeps the old code running — this is the #1 gotcha.*
4. Paste the Web App URL into the app's **Me → Backend** field, or bake it into
   `DEFAULT_SCRIPT_URL` at the top of `index.html` and push.

## First-time setup

1. Create a Google Sheet called "Match Subs Worksheet for App".
2. Extensions → Apps Script, paste in `google-apps-script.js`, save.
3. Run `setupSheets` once — approve the permissions prompt. All five tabs appear.
4. Fill in the **Teams** tab and set `CaptainPIN` + `AppUrl` on **Config**.
5. Deploy as a Web App (Execute as: **Me**, Access: **Anyone**). Copy the URL.
6. Run `setupNudgeTrigger` once to schedule the hourly no-sub check.
7. Put the Web App URL into `DEFAULT_SCRIPT_URL` in `index.html`, push to GitHub.

## Demo mode

While `DEFAULT_SCRIPT_URL` is empty and no URL is saved on the device, the app runs on
built-in sample data with a purple banner across the top. Nothing saves and no email is
sent — useful for showing captains how it works before any of it is real.

## Email volume

Google caps Apps Script at **100 recipients per day** on a consumer Gmail account, and it
fails quietly at the limit. Because captains hand-pick a handful of subs per match rather
than blasting everyone, normal use stays far below that. Worth remembering before anyone
adds a "notify all 60 subs" button.
