# Metadata Decoder (web app)

A web version of the **Metadata Sheet Decoder** desktop tool. Upload a music
metadata spreadsheet (`.xlsx`), and the app finds and helps you fix four classes
of problems — all inside the browser, with no Excel round-trips:

1. **Artist-name typos** — fuzzy-clusters near-duplicate artist spellings.
2. **Duplicate ISRCs** — the same ISRC on rows with different artists.
3. **Missing required fields** — blank cells that must be filled.
4. **Format issues** — bad ISRC/UPC patterns and master splits that don't sum to 100.

You edit the suggested corrections inline, click **Apply**, and the app rewrites
the workbook, re-scans, and updates the counts. When you're done you can download
the annotated copy and the issues report (see [Using the app](#using-the-app)).

This is the deployable successor to the `Metadata Sheet Decoder/` desktop scripts
(kept untouched in the parent directory as the reference implementation). The core
analysis engine is vendored verbatim under `backend/engine/`.

---

## Architecture

One Cloud Run service serves everything:

```
            ┌──────────────────────── Cloud Run service ────────────────────────┐
  Browser ─▶│  FastAPI  ──/api/*──▶  EngineService ──▶ vendored engine (pandas)   │
  (React)   │     │                      │   │                                    │
            │     └── serves built SPA    │   ├──▶ Cloud Storage  (xlsx + results)│
            │         (static/)           │   └──▶ Firestore      (scans + LEAVE) │
            └────────────────────────────────────────────────────────────────────┘
```

- **Frontend**: React + Vite + TypeScript + Tailwind. Built to static files and
  served by FastAPI in production.
- **Backend**: FastAPI. Wraps the original engine and exposes a JSON API.
- **Storage**: Google Cloud Storage holds each scan's `original.xlsx`,
  `annotated.xlsx`, `issues.xlsx`, and `results.json`.
- **Database**: Firestore holds scan metadata (filename, counts, timestamps) and
  the org-wide **LEAVE** decisions (artist clusters / ISRC pairs you've confirmed
  are intentional, so future scans stop flagging them).
- **Auth**: Google sign-in, restricted to a company email domain. The frontend
  sends the Google ID token as a bearer token; the backend verifies it.

Both storage and database have **local fallbacks** (filesystem + JSON file) so the
whole app runs with zero GCP setup during development.

---

## Using the app

The day-to-day flow for an end user (e.g. someone at CMG):

1. **Sign in** with their Google account.
2. **Upload** a metadata `.xlsx` on the home page. The app scans it and shows
   the issue counts.
3. **Open the scan** and work the four tabs — Artist names, Duplicate ISRCs,
   Missing fields, Formats. Edit the suggested corrections inline (changed
   fields are highlighted), then **Apply** (a confirmation summarizes what will
   change). Each Apply rewrites the working copy, re-scans, and updates counts.
   - **Leave** (artist) / **Confirm OK** (ISRC) mark something as *intentional*
     so it's remembered org-wide and future scans stop flagging it.
4. **Download** the results when done.

### What the two downloads produce

| Download | File | What it is |
| --- | --- | --- |
| **Download annotated** | `<name>_annotated.xlsx` | A copy of the **original sheet with all applied fixes baked into the cells**, plus color highlights + Excel comments on anything *still* flagged, and a "Decoder Summary" tab. This is the **corrected master spreadsheet** you hand off / ingest. |
| **Download report** | `<name>_issues.xlsx` | A **separate report of the problems** (tabs: Issues list, Artist Clusters, ISRC Conflicts, Missing, Format, Splits). After fixing + re-scanning it shows what's *left*. It's the audit/proof sheet — it does not contain the corrected track data. |

Both always reflect the latest state (every Apply regenerates them). When fully
resolved, the annotated file has no highlights left except items intentionally
kept as Leave/Confirm OK. If the downstream ingestion system is picky about extra
tabs/formatting, delete the "Decoder Summary" tab (the corrected *values* are
what matter; colors/comments are ignored by most systems).

---

## Local development

Requirements: Python 3.11+, Node 20+.

### 1. Backend

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Runs in local mode: filesystem storage, JSON DB, auth OFF.
uvicorn app.main:app --reload --port 8000
```

By default (`ENV=local`) the backend stores files under `backend/.local_data/`
and disables auth, so you can click around without signing in. Copy
`.env.example` to `.env` if you want to tweak anything.

### 2. Frontend

```bash
cd frontend
npm install
npm run dev          # http://localhost:5173
```

Vite proxies `/api/*` to the backend on port 8000 (see `vite.config.ts`). Open
http://localhost:5173 and upload a spreadsheet.

### Build the SPA (optional, to mimic prod locally)

```bash
cd frontend && npm run build      # outputs frontend/dist
```

In production the Docker build copies `frontend/dist` into `backend/static`, and
FastAPI serves it for all non-`/api` routes.

---

## Deployment (Google Cloud Run)

The whole thing deploys with one idempotent script.

### Prerequisites

- The [`gcloud` CLI](https://cloud.google.com/sdk/docs/install), authenticated
  (`gcloud auth login`).
- A GCP project with billing enabled.

### 1. Deploy

```bash
cd decoder-app
PROJECT_ID=your-project ./deploy.sh
```

`deploy.sh` is safe to run repeatedly. It will:

1. Enable the required APIs (Run, Cloud Build, Artifact Registry, Firestore, Storage).
2. Create the Firestore `(default)` database (native mode) if missing.
3. Create the GCS bucket (`<project>-metadata-decoder`) if missing.
4. Create a least-privilege runtime service account (`roles/datastore.user` +
   `roles/storage.objectAdmin` on the bucket only).
5. Build the container from the `Dockerfile` and deploy to Cloud Run.
6. Print the service URL.

Useful overrides (all optional):

```bash
PROJECT_ID=your-project \
REGION=us-central1 \
SERVICE=metadata-decoder \
BUCKET=your-project-metadata-decoder \
AUTH_ENABLED=true \
ALLOWED_EMAIL_DOMAIN=createmusicgroup.com \
ALLOWED_EMAILS="contractor@example.com" \
OAUTH_CLIENT_ID=xxxx.apps.googleusercontent.com \
./deploy.sh
```

### 2. Authentication setup (one-time)

Auth is **on by default** in production. Access is controlled by **two gates** —
an email must pass both:

1. **Google's gate** — the OAuth consent screen decides who Google will let
   authenticate at all.
2. **The app's gate** — the backend then only admits `@${ALLOWED_EMAIL_DOMAIN}`
   accounts plus anyone in `ALLOWED_EMAILS`. It verifies the ID token on every
   request, so the API is protected even though the URL is public.

#### a. Create the OAuth client

1. GCP Console → **APIs & Services → Credentials → Create credentials → OAuth
   client ID → Web application**.
2. Under **Authorized JavaScript origins**, add the Cloud Run URL printed by
   `deploy.sh` (e.g. `https://metadata-decoder-xxxx.run.app`) — exact, `https`,
   no trailing slash.
3. Copy the **Client ID** and redeploy with it:

   ```bash
   PROJECT_ID=your-project OAUTH_CLIENT_ID=xxxx.apps.googleusercontent.com ./deploy.sh
   ```

#### b. Configure the consent screen (Google's gate)

Go to **APIs & Services → OAuth consent screen → Audience**. Choose the model
that fits who owns the project:

- **Internal** (recommended for the final CMG handoff): only possible if the GCP
  project lives inside CMG's Google Workspace org. Every `@createmusicgroup.com`
  user can sign in immediately — **no test-user list, no verification, no extra
  steps**.
- **External** (used when the project is personal / outside the org): then either
  - **Testing mode** — only emails added under **Test users** can sign in. Add
    each tester there *and* to `ALLOWED_EMAILS` (both gates). Or
  - **In production** — click **Publish app**. Because this app only requests
    basic identity scopes (`openid`, `email`, `profile`, which are
    *non-sensitive*), publishing needs **no Google verification and shows no
    "unverified app" warning**, and the **test-user list disappears**. Access is
    then controlled entirely by the app (`ALLOWED_EMAIL_DOMAIN` /
    `ALLOWED_EMAILS`) — one source of truth. This is the cleanest External setup.

> To run prod-mode without auth temporarily (e.g. a private demo), deploy with
> `AUTH_ENABLED=false`. Not recommended for anything with real data.

### 3. Managing who can sign in

CMG staff get in automatically via `ALLOWED_EMAIL_DOMAIN`. To grant access to
individual outside emails, use `ALLOWED_EMAILS` (comma-separated).

`ALLOWED_EMAILS` is a single env var, so you always provide the **whole list**
(setting it replaces the previous value — it doesn't append). To change it
**without rebuilding the container** (a few seconds vs. several minutes):

```bash
gcloud run services update metadata-decoder \
  --region us-central1 \
  --update-env-vars "^##^ALLOWED_EMAILS=a@gmail.com,b@gmail.com,c@outlook.com"
```

(The `^##^` prefix makes gcloud split on `##` so the commas inside the list are
safe.) A full `./deploy.sh` also works but rebuilds the image. After changing the
list, affected users should sign out / hard-refresh to pick up a fresh token.

> If the project is **External + Testing**, remember to also add the email under
> **OAuth consent screen → Audience → Test users**. In **Internal** or
> **External + production**, `ALLOWED_EMAILS` is the only place to manage.

---

## Handing off to the client

Decide who *owns* the running app:

**Option A — you host it on CMG's GCP (simplest for them).**
Run the deploy + auth steps above using a **CMG-owned** project (so you can use
an **Internal** consent screen). Then hand over:
1. The **live URL**.
2. A one-liner: "Sign in with your `@createmusicgroup.com` Google account."
3. **Admin access**: add their IT/admin as **Owner** on the GCP project so they
   control it long-term.
4. The **code** (this `decoder-app/` folder or a Git repo) as the source of record.

**Option B — they deploy it themselves.** Hand over:
1. The **`decoder-app/` folder** (zip or Git repo — `.gitignore` keeps secrets out).
2. This **README**.
3. Prereqs: a GCP project with billing, the `gcloud` CLI, ~10 minutes. They run
   `PROJECT_ID=... ./deploy.sh`, create the OAuth client, and set the consent screen.

> Do **not** ship the old desktop tool's `credentials.json` / `.sheets_token.json`.
> They're unrelated secrets from the Google Sheets workflow and this app never
> uses them.

---

## Continuous deployment (GitHub Actions)

Push to `main` → it builds and deploys to Cloud Run automatically. Auth is
**keyless** via Workload Identity Federation (no service-account JSON key stored
in GitHub). Files: `.github/workflows/deploy.yml` + `scripts/setup-github-deploy.sh`.

> Make the **git repo = this `decoder-app/` folder** (not the parent
> `meta-decoder/`). That keeps the old desktop tool's secrets out of the repo and
> puts the `Dockerfile` at the repo root.

### One-time setup

1. **Deploy once manually first** so the project, APIs, bucket, Firestore, and
   runtime service account exist:
   ```bash
   PROJECT_ID=decoder-app-500118 OAUTH_CLIENT_ID=... ./deploy.sh
   ```

2. **Init git and push to GitHub** (from inside `decoder-app/`):
   ```bash
   cd decoder-app
   git init -b main
   git add .
   git commit -m "Metadata Decoder web app"
   git remote add origin https://github.com/<owner>/<repo>.git
   git push -u origin main
   ```

3. **Wire up keyless auth** (creates a deployer SA + Artifact Registry repo +
   Workload Identity pool locked to your repo):
   ```bash
   GITHUB_REPO=<owner>/<repo> PROJECT_ID=decoder-app-500118 ./scripts/setup-github-deploy.sh
   ```
   It prints two values. Add them as **GitHub repo secrets** (repo → Settings →
   Secrets and variables → Actions → New repository secret):
   - `WIF_PROVIDER`
   - `WIF_SERVICE_ACCOUNT`

4. Done. Every push to `main` now deploys. You can also trigger it manually from
   the **Actions** tab → *Deploy to Cloud Run* → *Run workflow*.

### What the pipeline does (and doesn't) touch

- It **builds the image and deploys the code** only.
- It **does not** set env vars — config (`OAUTH_CLIENT_ID`, `ALLOWED_EMAILS`,
  domain, etc.) is **inherited from the running revision**. So your one-time
  `./deploy.sh` config and any live `gcloud run services update` changes survive
  every CI deploy. Manage access the same way as before (see
  [Managing who can sign in](#3-managing-who-can-sign-in)).
- If you change the project/region/service names, update the `env:` block at the
  top of `.github/workflows/deploy.yml`.

---

## Configuration reference

All configuration is via environment variables (see `backend/app/config.py`).

| Variable               | Default                 | Purpose                                                        |
| ---------------------- | ----------------------- | -------------------------------------------------------------- |
| `ENV`                  | `local`                 | `local` or `prod`. Controls CORS + auth defaults.              |
| `AUTH_ENABLED`         | `false` local / on prod | Require Google sign-in.                                        |
| `OAUTH_CLIENT_ID`      | —                       | Google OAuth Web client ID (token audience).                  |
| `ALLOWED_EMAIL_DOMAIN` | `createmusicgroup.com`  | Domain allowed to sign in.                                     |
| `ALLOWED_EMAILS`       | —                       | Comma-separated extra allowlist (bypasses the domain check).   |
| `GOOGLE_CLOUD_PROJECT` | —                       | Set → use Firestore. Empty → local JSON DB.                    |
| `GCS_BUCKET`           | —                       | Set → use Cloud Storage. Empty → local filesystem.             |
| `FIRESTORE_DATABASE`   | `(default)`             | Firestore database id.                                          |
| `DATA_DIR`             | `./.local_data`         | Local-mode storage + DB location.                              |
| `MAX_UPLOAD_BYTES`     | `26214400` (25 MB)      | Upload size limit.                                             |
| `STATIC_DIR`           | `backend/static`        | Where the built SPA lives (set by the Docker build).           |

---

## API reference

All routes are under `/api` and require a valid bearer token when auth is on.

| Method   | Path                                    | Purpose                                  |
| -------- | --------------------------------------- | ---------------------------------------- |
| `GET`    | `/api/config`                           | Public: auth on/off + OAuth client ID.   |
| `GET`    | `/api/healthz`                          | Liveness check.                          |
| `GET`    | `/api/me`                               | Current signed-in user.                  |
| `POST`   | `/api/scans`                            | Upload a workbook → run a scan.          |
| `GET`    | `/api/scans`                            | List scans (dashboard).                  |
| `GET`    | `/api/scans/{id}`                       | Scan detail + results.                   |
| `DELETE` | `/api/scans/{id}`                       | Delete a scan and its files.             |
| `GET`    | `/api/scans/{id}/files/{which}`         | Download `original` / `annotated` / `issues`. |
| `POST`   | `/api/scans/{id}/corrections/artist`    | Apply artist-name fixes, re-scan.        |
| `POST`   | `/api/scans/{id}/corrections/isrc`      | Apply ISRC fixes/confirmations, re-scan. |
| `POST`   | `/api/scans/{id}/corrections/missing`   | Fill missing fields, re-scan.            |
| `POST`   | `/api/scans/{id}/corrections/format`    | Apply format fixes, re-scan.             |

---

## Desktop → web app mapping

How the original tool's pieces map onto this app:

| Desktop tool                                  | Web app equivalent                                                            |
| --------------------------------------------- | ----------------------------------------------------------------------------- |
| `scan_metadata.py` (`analyze`)                | Called by `EngineService.create_scan` on upload and after every fix.          |
| `apply_corrections.py` (artist)               | `POST /corrections/artist` → `EngineService.apply_artist`.                    |
| `apply_isrc_corrections.py`                   | `POST /corrections/isrc` → `EngineService.apply_isrc`.                        |
| `apply_missing_corrections.py`                | `POST /corrections/missing` → `EngineService.apply_missing`.                  |
| `apply_format_corrections.py`                 | `POST /corrections/format` → `EngineService.apply_format`.                    |
| `*_issues.xlsx` (edit corrections in Excel)   | Inline-editable tables in the scan detail view (no Excel round-trip).         |
| `*_annotated.xlsx`, `*_issues.xlsx` outputs   | Downloadable from the scan detail page; stored per-scan in Cloud Storage.     |
| `Decoder Dashboard.html`                      | The React dashboard (`/`) and scan detail page.                               |
| `.artist_leave.json` / `.isrc_leave.json`     | LEAVE decisions stored org-wide in Firestore; "Leave" / "Confirm OK" buttons. |
| `.command` launchers                          | Not needed — it's a hosted web app.                                           |
| Google Sheets upload + `credentials.json`     | Dropped. Auth is Google sign-in; files live in Cloud Storage. No secrets in the repo. |

The one change to the vendored engine: `scan_metadata.analyze()` accepts an
optional `project_dir` so the backend can point LEAVE-record lookups at a temp
directory materialized from Firestore instead of the module's own folder. See
`backend/engine/__init__.py` for details.

---

## Troubleshooting

**Sign-in seems to do nothing / bounces back to the login screen.**
The account isn't allowed. Google signed them in, but the backend rejected the
email (wrong domain and not in `ALLOWED_EMAILS`). The sign-in screen now shows
the reason (e.g. *"Access is restricted to @createmusicgroup.com accounts."*).
Fix: use a domain account, or add the email to `ALLOWED_EMAILS` (see
[Managing who can sign in](#3-managing-who-can-sign-in)).

**Console shows CORS errors for `play.google.com/log` or `accounts.google.com/gsi/log`.**
Harmless. These are Google Identity telemetry pings that the browser blocks; they
do **not** affect sign-in. Ignore them.

**`deploy.sh` fails with "Service account ... does not exist" right after creating it.**
A new service account is eventually consistent and can briefly 404 in IAM. The
script already waits + retries these bindings; if you hit it on an older copy,
just **re-run `./deploy.sh`** (it's idempotent and the account will exist by then).

**"The given origin is not allowed for the given client ID."**
The Cloud Run URL isn't in the OAuth client's **Authorized JavaScript origins**.
Add the exact `https://...run.app` URL (no trailing slash) and retry.

**Changed `ALLOWED_EMAILS` but a user still can't get in.**
Have them sign out / hard-refresh so a fresh token is issued. Confirm the new
revision is serving (`gcloud run services describe metadata-decoder --region us-central1`).

---

## Project layout

```
decoder-app/
├── Dockerfile            # multi-stage: build SPA, serve from FastAPI
├── deploy.sh             # idempotent GCP provisioning + Cloud Run deploy
├── .github/workflows/    # deploy.yml: push-to-main auto-deploy (GitHub Actions)
├── scripts/              # setup-github-deploy.sh: one-time keyless CI auth
├── .env.example          # local config template
├── backend/
│   ├── app/              # FastAPI app (config, storage, db, auth, engine_service, api)
│   ├── engine/           # vendored desktop engine (scan + apply_* + threaded_comments)
│   └── requirements.txt
└── frontend/             # React + Vite + Tailwind SPA
    └── src/
        ├── pages/        # HomePage (upload + dashboard), ScanDetailPage
        │   └── tabs/     # Artist / Isrc / Missing / Format correction tables
        ├── components/   # UI primitives, Table, Layout
        └── lib/          # api client, auth context, types
```
