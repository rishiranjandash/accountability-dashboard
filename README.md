# Warehouse Leaders — Accountability Board

A small dashboard for the "Warehouse Leaders" daily task-accountability system.
It reads the same **Accountability Playbook — Poll Options & Follow-ups**
Google Sheet that already drives the WhatsApp poll automation, and shows
today's task board, the points leaderboard, the leave tracker, and the
slippage report — live, mobile-friendly, no login required to view.

Live sheet: https://docs.google.com/spreadsheets/d/1chXS-4BZjJDuZ91jfvSo99aNUmp3TkAjwZJC31f-L14/edit

## How it fits with the existing WhatsApp automation

This is **additive, not a replacement** for v1. The WhatsApp group polls
keep working exactly as they do today. This dashboard is a read-mostly
viewer on top of the same Sheet, so anyone can check status on their phone
without scrolling WhatsApp. It also exposes a `markDone` write endpoint
(see below) that can optionally be wired up later if you want people to be
able to tick tasks off directly from the page instead of (or in addition
to) the WhatsApp poll — nothing in v1 requires that.

## Structure

```
apps-script/Code.gs   — the backend. Deployed via Google Apps Script as a
                         Web App; reads/writes the Sheet, serves JSON.
docs/index.html        — the frontend. A single self-contained static page,
                         served by GitHub Pages from /docs on the main branch.
```

## Setup (one-time)

### 1. Deploy the backend (Google Apps Script)

1. Open the Sheet above → **Extensions → Apps Script**.
2. Delete the boilerplate `Code.gs` content and paste in
   [`apps-script/Code.gs`](apps-script/Code.gs) from this repo.
3. (Optional but recommended) **Project Settings → Script properties → Add
   property**: name it `WRITE_TOKEN`, value = any random string you pick.
   This is the shared secret that protects the `markDone` write endpoint —
   without it, anyone with the URL could write rows.
4. **Deploy → New deployment → type: Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click Deploy, authorize the requested scopes (it only touches this one
   Sheet), and copy the `.../exec` URL it gives you.

### 2. Point the frontend at the backend

In `docs/index.html`, find:
```js
const APPS_SCRIPT_URL = 'PASTE_YOUR_APPS_SCRIPT_/exec_URL_HERE';
```
and replace it with the `/exec` URL from step 1. Commit the change.

### 3. Turn on GitHub Pages

Repo → **Settings → Pages → Source: Deploy from a branch → Branch: main,
folder: /docs → Save**. GitHub will give you a URL like
`https://<your-username>.github.io/<repo-name>/` — that's the dashboard.

## Notes

- The `Points` and `Slippage Report` tabs still contain data from the
  original pilot group ("Hello world - night shift - LW"), not the live
  Warehouse Leaders group — those two tabs need to be migrated/repopulated
  with real Warehouse Leaders data for the Leaderboard and Slippage tabs in
  the app to show anything meaningful. The Today board does not depend on
  them and works off the current data already in the Sheet.
- The `Daily Log` tab is currently empty, so the Today board's per-task
  status will show "Pending" for everything until something starts writing
  rows to it — either the WhatsApp-checkpoint automation (writing a row
  each time it confirms a task done, so this dashboard mirrors what's
  already being tracked) or people using the page's own mark-done action,
  once that's wired up. Not required for the initial deploy.
