# BeeAware by LPP

Field-awareness app for Lawns Plants & Pests: before treating a property, check
where registered apiaries (bees) and sting-hypersensitive people are.

- **Live app:** https://lawnsplantspests-ui.github.io/lpp-search/
- **Data:** one Google Sheet (the "Master Sheet"), read/written through a Google
  Apps Script API. A backup of that script is in [`apps-script/Code.gs`](apps-script/Code.gs).

## ⚠️ DO NOT upload old copies of these files

**The files in this GitHub repo are the only source of truth.**

Old copies of `map.html`, `browse.html`, etc. saved on the Desktop are STALE.
Uploading them here ("Add files via upload") has **erased finished features
before** (June 2026: it wiped the unregistered-hives layer, the satellite
toggle, and the BeeAware rename, which then had to be restored from git
history).

If a feature ever disappears from the app, check the recent commits for an
"Add files via upload" commit — that is almost always the cause, and the fix
is restoring the file from the commit before it.

## The two halves of the app

1. **These HTML files** — the screens. Pushing to `main` updates the live app
   automatically (give it ~1 minute, then fully reload the app).
2. **The Apps Script** — lives *inside the Google Sheet* (Extensions → Apps
   Script). It is the only thing that reads/writes the Sheet. Changing it
   requires pasting the code there and redeploying (Deploy → Manage
   deployments → pencil → New version → Deploy). Keep
   [`apps-script/Code.gs`](apps-script/Code.gs) in sync with it as a backup.

## Pages

| Page | What it is |
|---|---|
| `index.html` | Sign-in + dashboard |
| `search.html` | Address check before a treatment |
| `map.html` | Live map: apiaries, hypers, swarms, unregistered hives, measure tool, spot check |
| `browse.html` | Admin: search all records by name, inline results map |
| `admin.html` | Admin: users, activity, sync tools, PA Plants importer |
| `import.html` | Converts pasted PA Plants results into sheet rows |
| `recensus.html` | Bulk re-geocode via US Census (preferred) |
| `geocode.html` | Bulk geocode via OSM (legacy) |
| `edit.html` / `verify.html` / `restore-gps.html` | Location fix-up tools |
| `log.html` | Contact logs |
