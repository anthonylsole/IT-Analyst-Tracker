# DELTA IT Resource and Project Allocation — Self-Hosted Version

A React app (Vite) served by a Cloudflare Worker, storing data in Cloudflare
D1, gated behind a single shared password. This replaces the Claude.ai
"artifact" version — same app, same code, but hosted somewhere you fully
control.

## What you need first

- A [GitHub](https://github.com) account
- A [Cloudflare](https://dash.cloudflare.com/sign-up) account (free tier is enough)
- [Node.js](https://nodejs.org) 18+ installed locally
- The `wrangler` CLI (installed automatically as a dev dependency below)

## 1. Put this on GitHub

```bash
cd capacity-tracker
git add -A
git commit -m "Initial commit"
```

Then create a new empty repo on GitHub (no README/license, so it doesn't
conflict), and push:

```bash
git remote add origin https://github.com/<your-username>/<your-repo-name>.git
git branch -M main
git push -u origin main
```

## 2. Install dependencies

```bash
npm install
```

## 3. Log in to Cloudflare

```bash
npx wrangler login
```

This opens a browser window to authorize the CLI against your Cloudflare
account.

## 4. Create the D1 database

```bash
npx wrangler d1 create capacity_tracker_db
```

This prints something like:

```
[[d1_databases]]
binding = "DB"
database_name = "capacity_tracker_db"
database_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

Copy that `database_id` value and paste it into `wrangler.toml`, replacing
`REPLACE_WITH_YOUR_D1_DATABASE_ID`.

## 5. Create the database table

```bash
npm run db:migrate:remote
```

This runs `schema.sql` against your live D1 database and creates the single
table the app needs (`kv_store`).

(If you also want to test locally with `wrangler dev`'s local D1 emulation,
run `npm run db:migrate:local` too.)

## 6. Set your password

Pick a password and a random session secret, then set them as Worker secrets
(these are encrypted by Cloudflare, never committed to git):

```bash
npx wrangler secret put SITE_PASSWORD
# paste your chosen password when prompted

npx wrangler secret put SESSION_SECRET
# paste any long random string when prompted, e.g. output of:
# openssl rand -hex 32
```

Anyone visiting the site will need `SITE_PASSWORD` to get in. `SESSION_SECRET`
is just an internal value used to sign the login cookie — it should be
random and doesn't need to be memorable.

## 7. Build and deploy

```bash
npm run deploy
```

This builds the React app (`vite build` → `dist/`) and deploys the Worker,
which serves that build and handles the API + login. Wrangler will print the
`*.workers.dev` URL your app is live at.

## 8. Using it day to day

- Visit your Worker's URL. You'll be asked for the password once; after
  that, a cookie keeps you signed in for 30 days.
- All the features from the Claude.ai version work as before: View
  resource/project dropdowns, the Team table, Update Allocations self-serve
  form, Export/Import, and automatic + manual backups.
- Data lives in your D1 database now — not in Claude.ai — so it's under your
  control and isn't affected by anything happening on the Claude.ai side.

## Making future changes

If you ask Claude to change the tool later, just replace
`src/CapacityTracker.jsx` with the updated file Claude gives you, then:

```bash
npm run deploy
```

## A note on the password protection

This is intentionally simple, as requested: one shared password, checked by
the Worker itself (not just hidden in client-side JavaScript, so it can't be
bypassed by viewing page source). It's appropriate for a small internal tool
where everyone with the password is trusted. It is **not** per-user
authentication — everyone shares one password and one session cookie type,
and there's no audit trail of who did what. If you ever need real per-person
logins, that would require a bigger authentication layer (e.g. Cloudflare
Access, or a proper auth provider) on top of this.

## Project structure

```
capacity-tracker/
├── src/
│   ├── CapacityTracker.jsx   # the app itself (unchanged from Claude.ai)
│   ├── storagePolyfill.js    # replaces Claude's window.storage with D1/localStorage
│   └── main.jsx              # React entry point
├── worker/src/index.js       # Cloudflare Worker: auth gate + API + static serving
├── schema.sql                # D1 table definition
├── wrangler.toml             # Cloudflare Worker + D1 + assets config
└── index.html / vite.config.js / package.json
```
