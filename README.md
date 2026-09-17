# Match Sub App

A Progressive Web App for finding subs across multiple tennis teams. Captains post a
match, hand-pick who gets emailed, and the first sub to accept gets the spot.

**Live:** [icc-match-subs.vercel.app](https://icc-match-subs.vercel.app)

## How it works

- **Players sign themselves up** — name, email, cell phone, and **one question: their
  soft court level**. A confirmation email has to be clicked before they receive
  anything, which catches typo'd addresses before they cost anyone a match.
- **Captains post a match** — date, singles/doubles/both, location, and optionally a
  start time, opponent and notes. The team and its level come from whoever is posting.
- **Captains choose exactly who to notify.** The list shows everyone eligible, grouped
  by level with the strongest first, and within each level sorted by who has subbed
  least for that team — so the same reliable person doesn't get asked every time.
- **First to accept wins.** Subs accept straight from the email or in the app. The
  captain gets an email with the sub's name, email and phone, and the sub gets a receipt
  saying the captain will text them.
- **If nobody bites**, the captain gets a nudge email after 24 hours with a reminder to
  widen the net.

## The two rules that decide who a captain sees

**You can sub up, never down.** A match at 3.5 can use anyone rated 3.5 or lower. A 4.0
player is not eligible and does not appear at all. This is why signup only needs one
level question — eligibility falls out of a single number.

**Three matches per team, per season.** Once someone has subbed three times for Miller
4.0 they're locked for that team until the season changes, but stay fully available for
every other team. Maxed-out players are still shown in the picker, greyed out with a
reason, so a captain hunting for a specific name sees *why* they can't pick them.

The cap is enforced **on the server**, not just in the picker — the claim button in an
email bypasses the app's screens entirely, so the backend is the only place that can
actually stop it. Backing out of a match gives the slot back.

Seasons are **Spring and Fall only**. Rolling over is a one-cell edit: change `Season`
on the Config tab and every count reads as zero. Nothing parses the value, so any
consistent naming works as long as it changes between seasons.

## Emails

Every email goes out as HTML with a real button, plus a plain-text version.

**Who it appears to come from.** Apps Script can only send from the account that
deployed it, and putting a captain's Gmail address in From would be treated as spoofing.
So the captain goes in the display name and Reply-To instead:

| Email | Shows as from | Reply goes to |
|-------|---------------|---------------|
| Sub needed / No longer needed (to subs) | *Jane Miller via ICC Match Subs* | the captain who posted |
| You're subbing (receipt) / Confirm you're backing out | *Jane Miller via ICC Match Subs* | the captain who posted |
| You were added to the sub list | *[captain] via ICC Match Subs* | the captain who added them |
| Someone is subbing / Sub dropped out (to the captain) | ICC Match Subs | the sub |
| Confirm your spot / your changes, Still no sub | ICC Match Subs | the sending account |

**Buttons open the app, never `script.google.com`.** A browser signed into more than one
Google account rewrites script.google.com links to `/macros/u/1/…`, which fails with
Google Drive's "unable to open the file" page. Email buttons go to the app instead
(`?claim=`, `?verify=`, `?release=`, `?optout=` plus a per-sub token `t`), and the app
hands the request to the backend with `fetch`, which carries no Google sign-in. It also
means nothing happens until a person opens the link — a mail scanner that pre-fetches
URLs gets a page, not a claimed match. Old-style links in emails sent before this change
still work.

## Only the sub can change their own things

The app knows a sub by an ID that anyone can see, so anything that changes a sub's
record is confirmed from **their own inbox** with a tokenized link:

- **One person, one spot.** A signup with an email or cell that's already on the list is
  refused ("You're already signed up"). The app offers **Update my details instead**: the
  new values are parked in the `Pending …` columns and a "Confirm your changes" email goes
  to that address, applied only when its button is tapped. Someone who signed up but never
  confirmed gets their link re-sent (so a typo'd first try can't lock them out), and
  someone who opted out can rejoin. Captains can't manually add a duplicate either.
- **The signup form always starts blank.** A phone may have been used by a captain or
  someone else, so no name is ever carried into it.
- **Backing out of a match** ("I can no longer play") emails the sub a confirm button.
  The spot stays theirs until they tap it — then the captain is told and it reopens.
- **Every claim sends the sub a receipt** with a "Release your spot" link. Beyond being
  useful, it means a claim made in someone else's name lands in the real person's inbox.
- **Opting out** is a token link in the footer of every email.

Claiming a match *inside the app* still isn't tied to a token — the receipt email is what
catches misuse there.

## Architecture

| Piece | Where | Notes |
|-------|-------|-------|
| `index.html` | Vercel | The entire frontend — vanilla JS/CSS, no build step |
| `google-apps-script.js` | Google Apps Script web app | All reads/writes; deployed manually (see below) |
| Google Sheet | Google Drive | The database — 5 tabs, all self-installing |
| `sw.js` / `manifest.json` | Vercel | PWA install + offline shell |

## Google Sheet tabs

All five are created automatically the first time the script runs. **Teams** and
**Config** are the only ones filled in by hand. Sheets also **self-heal**: if a new
version expects a column an existing tab doesn't have, the script adds it. Columns it no
longer uses (like `Line` on older `Requests` tabs) are left alone — nothing is deleted.
Rows are written by matching header names, so column order doesn't matter.

| Tab | Headers |
|-----|---------|
| **Teams** | `TeamID, Team Name, Level, Captain, Captain Email, Active` |
| **Subs** | `SubID, Name, Email, Phone, Level, Verified, Token, Active, Added By, Signed Up At, Sub Count, Last Sub, Season, Team Subs, Pending Name, Pending Phone, Pending Level` |
| **Requests** | `ID, TeamID, Level, Date, Time, Match Type, Location, Opponent, Notes, Posted By, Posted At, Notified, Status, Claimed By, Claimed At, Nudged` |
| **History** | `Request ID, Team, Match Date, Sub Name, Sub Email, Claimed At, Posted By` |
| **Config** | `Key, Value` |

### Filling in Teams

One row per team. Leave `TeamID` blank and the script fills it in for you. `Active` can
be `TRUE`/`FALSE` (blank counts as active).

| TeamID | Team Name | Level | Captain | Captain Email | Active |
|--------|-----------|-------|---------|---------------|--------|
| *(blank)* | Miller 4.0 | 4.0 | Jane Miller | jane@example.com | TRUE |

- **Captain Email** is where claim notifications and nudges go, and what subs reach when
  they hit Reply — if it's blank, that captain hears nothing and replies go to the
  sending account.
- **Captain** must match the name the captain picks in the app, or their replies fall
  back to the team's listed captain.
- **TeamID is permanent.** Rename a team freely when a captain hands off, but never change
  its TeamID — it's what ties subs' per-team counts and past matches to the team.
- **Levels:** Google Sheets stores a typed `3.0` as the number `3`. The script normalises
  every level to one decimal on read, so don't try to "fix" the column by hand.

### Config keys

| Key | What it does |
|-----|--------------|
| `CaptainPIN` | Shared PIN that unlocks posting and the sub list. **Change it from `1234`.** |
| `AppUrl` | Optional. The app's address for links in emails; defaults to `https://icc-match-subs.vercel.app`. |
| `ManagerEmail` | Optional — CC'd on every claim confirmation. Leave blank to skip. |
| `NudgeHours` | How long before a captain gets nudged about an unclaimed match. Default `24`. |
| `Season` | Current season, e.g. `2026 Fall`. Changing it resets every sub's per-team count. |
| `MaxSubsPerTeam` | How many times one sub may play for one team per season. Default `3`. |

Config is never rebuilt once it exists, so a key added in a later version has to be typed
in by hand.

The `Team Subs` column on the Subs tab stores counts as `miller4.0:2, ray4.5:1` —
readable if you need to check or correct someone by hand. It's paired with `Season`, so a
count from a previous season is ignored rather than deleted.

## Why there's a PIN

The app URL is public, and the `Subs` tab holds members' personal emails and phone
numbers. Without a gate, anyone with the link could read all of it. So:

- **No PIN needed** to sign up as a sub or to accept a match.
- **PIN required** to post a match, see the sub list, or add someone manually.

Contact details are stripped from the API response entirely unless the PIN is supplied —
they aren't just hidden in the UI. That's only as good as the PIN, so change it from the
default.

## Deploying changes

**Frontend** (`index.html`, `sw.js`): commit and push. Vercel deploys automatically. The
service worker is network-first for `index.html`, so a new version shows up on the next
load without any cache bumping.

**Backend** (`google-apps-script.js`):
1. Bump the `VERSION` constant at the top of the file **and** `APP_VERSION` in
   `index.html` (keep them equal).
2. Paste the whole file over the code in the Apps Script editor (Extensions → Apps
   Script from the Sheet). Keep what gets pasted **pure ASCII** — the clipboard mangles
   em-dashes and curly quotes.
3. Deploy → Manage deployments → ✏️ edit → **Version: New version** → Deploy.
   *Skipping "New version" keeps the old code running — this is the #1 gotcha.* Don't use
   "New deployment" either: that creates a different URL.
4. Check `<web app URL>?action=ping` reports the new `version`.

When a change touches both sides and alters what email links look like, **push the
frontend first**, so emails never point at an app that doesn't understand them yet.

## First-time setup

1. Create a Google Sheet called "Match Subs Worksheet for App".
2. Extensions → Apps Script, paste in `google-apps-script.js`, save.
3. Run `setupSheets` once — approve the permissions prompt. All five tabs appear.
4. Fill in the **Teams** tab and set `CaptainPIN` on **Config**.
5. Deploy as a Web App (Execute as: **Me**, Access: **Anyone**). Copy the URL.
6. Run `setupNudgeTrigger` once to schedule the hourly no-sub check.
7. Put the Web App URL into `DEFAULT_SCRIPT_URL` in `index.html`, push to GitHub.

## Demo mode

While `DEFAULT_SCRIPT_URL` is empty and no URL is saved on the device, the app runs on
built-in sample data with a purple banner across the top. Nothing saves and no email is
sent — useful for showing captains how it works before any of it is real.

## Email volume

Google caps Apps Script at **100 recipients per day** on a consumer Gmail account, and it
fails quietly at the limit. Each claim now sends two emails (the captain's notice and the
sub's receipt), but because captains hand-pick a handful of subs per match rather than
blasting everyone, normal use stays well below that. Worth remembering before anyone adds
a "notify all 60 subs" button.
