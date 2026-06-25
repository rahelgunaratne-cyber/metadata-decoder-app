#!/usr/bin/env bash
#
# Provision + deploy the Metadata Decoder to Google Cloud Run.
#
# This script is IDEMPOTENT: run it as many times as you like. It enables the
# required APIs, creates the Firestore database, the GCS bucket, and a runtime
# service account (with least-privilege roles), then builds the container from
# the Dockerfile and deploys it to Cloud Run.
#
# Usage:
#   ./deploy.sh                       # uses your current gcloud project
#   PROJECT_ID=my-proj ./deploy.sh    # explicit project
#
# Common overrides (env vars):
#   REGION=us-central1                Cloud Run + Firestore + bucket region
#   SERVICE=metadata-decoder          Cloud Run service name
#   BUCKET=<project>-metadata-decoder GCS bucket for uploaded/generated files
#   AUTH_ENABLED=true                 Require Google sign-in (recommended)
#   ALLOWED_EMAIL_DOMAIN=createmusicgroup.com
#   ALLOWED_EMAILS=                   Comma-separated extra allowlist
#   OAUTH_CLIENT_ID=                  Google OAuth Web client ID (see README)
#
set -euo pipefail

# Always run from the directory that holds the Dockerfile.
cd "$(dirname "$0")"

# ---- Configuration ---------------------------------------------------------
PROJECT_ID="${PROJECT_ID:-$(gcloud config get-value project 2>/dev/null || true)}"
REGION="${REGION:-us-central1}"
SERVICE="${SERVICE:-metadata-decoder}"
BUCKET="${BUCKET:-${PROJECT_ID}-metadata-decoder}"
SA_NAME="${SA_NAME:-metadata-decoder-run}"

AUTH_ENABLED="${AUTH_ENABLED:-true}"
ALLOWED_EMAIL_DOMAIN="${ALLOWED_EMAIL_DOMAIN:-createmusicgroup.com}"
ALLOWED_EMAILS="${ALLOWED_EMAILS:-}"
OAUTH_CLIENT_ID="${OAUTH_CLIENT_ID:-}"

SA_EMAIL="${SA_NAME}@${PROJECT_ID}.iam.gserviceaccount.com"

# ---- Helpers ---------------------------------------------------------------
bold() { printf '\033[1m%s\033[0m\n' "$*"; }
step() { printf '\n\033[1;36m==> %s\033[0m\n' "$*"; }
die()  { printf '\033[1;31mERROR:\033[0m %s\n' "$*" >&2; exit 1; }

# Retry a command a few times. New IAM resources (service accounts) are
# eventually consistent, so a binding can briefly 400 with "does not exist"
# right after creation; retrying clears it.
retry() {
  local n=0 max="${RETRY_MAX:-8}" delay="${RETRY_DELAY:-5}"
  until "$@"; do
    n=$((n + 1))
    if [ "$n" -ge "$max" ]; then
      return 1
    fi
    echo "  ...transient failure, retrying ($n/$max) in ${delay}s"
    sleep "$delay"
  done
}

command -v gcloud >/dev/null 2>&1 || die "gcloud CLI not found. Install: https://cloud.google.com/sdk/docs/install"
[ -n "$PROJECT_ID" ] || die "No project. Run 'gcloud config set project <id>' or pass PROJECT_ID=<id>."

bold "Project:        $PROJECT_ID"
bold "Region:         $REGION"
bold "Service:        $SERVICE"
bold "Bucket:         gs://$BUCKET"
bold "Service acct:   $SA_EMAIL"
bold "Auth enabled:   $AUTH_ENABLED"
[ -n "$OAUTH_CLIENT_ID" ] && bold "OAuth client:   ${OAUTH_CLIENT_ID:0:24}..." || true

gcloud config set project "$PROJECT_ID" >/dev/null

if [ "$AUTH_ENABLED" = "true" ] && [ -z "$OAUTH_CLIENT_ID" ]; then
  printf '\033[1;33mWARNING:\033[0m AUTH_ENABLED=true but no OAUTH_CLIENT_ID set.\n'
  printf '         Sign-in will be broken until you create an OAuth client and re-run with\n'
  printf '         OAUTH_CLIENT_ID=... (see README "Authentication setup"). Continuing.\n'
fi

# ---- 1. Enable APIs --------------------------------------------------------
step "Enabling required APIs (idempotent)"
gcloud services enable \
  run.googleapis.com \
  cloudbuild.googleapis.com \
  artifactregistry.googleapis.com \
  firestore.googleapis.com \
  storage.googleapis.com

# ---- 2. Firestore (native mode) -------------------------------------------
step "Ensuring Firestore database exists"
if gcloud firestore databases describe --database="(default)" >/dev/null 2>&1; then
  echo "  Firestore (default) already exists."
else
  gcloud firestore databases create --database="(default)" --location="$REGION" --type=firestore-native
fi

# ---- 3. GCS bucket ---------------------------------------------------------
step "Ensuring GCS bucket exists"
if gcloud storage buckets describe "gs://$BUCKET" >/dev/null 2>&1; then
  echo "  Bucket gs://$BUCKET already exists."
else
  gcloud storage buckets create "gs://$BUCKET" \
    --location="$REGION" \
    --uniform-bucket-level-access
fi

# ---- 4. Runtime service account + IAM -------------------------------------
step "Ensuring runtime service account exists"
if gcloud iam service-accounts describe "$SA_EMAIL" >/dev/null 2>&1; then
  echo "  Service account already exists."
else
  gcloud iam service-accounts create "$SA_NAME" \
    --display-name="Metadata Decoder (Cloud Run runtime)"
  # Wait for the new account to become visible to IAM before binding roles.
  echo "  Waiting for service account to propagate..."
  retry gcloud iam service-accounts describe "$SA_EMAIL" >/dev/null 2>&1 \
    || die "Service account $SA_EMAIL did not become available."
fi

step "Granting least-privilege roles"
# Firestore read/write. (Retry: IAM propagation can lag account creation.)
retry gcloud projects add-iam-policy-binding "$PROJECT_ID" \
  --member="serviceAccount:$SA_EMAIL" \
  --role="roles/datastore.user" \
  --condition=None >/dev/null
# Read/write objects in the one bucket only (scoped to the bucket, not project).
retry gcloud storage buckets add-iam-policy-binding "gs://$BUCKET" \
  --member="serviceAccount:$SA_EMAIL" \
  --role="roles/storage.objectAdmin" >/dev/null

# ---- 5. Build + deploy -----------------------------------------------------
step "Building container and deploying to Cloud Run"
# Use a custom '##' delimiter so values containing commas (ALLOWED_EMAILS) are safe.
ENV_VARS="^##^ENV=prod"
ENV_VARS="${ENV_VARS}##GOOGLE_CLOUD_PROJECT=${PROJECT_ID}"
ENV_VARS="${ENV_VARS}##GCS_BUCKET=${BUCKET}"
ENV_VARS="${ENV_VARS}##AUTH_ENABLED=${AUTH_ENABLED}"
ENV_VARS="${ENV_VARS}##ALLOWED_EMAIL_DOMAIN=${ALLOWED_EMAIL_DOMAIN}"
ENV_VARS="${ENV_VARS}##ALLOWED_EMAILS=${ALLOWED_EMAILS}"
ENV_VARS="${ENV_VARS}##OAUTH_CLIENT_ID=${OAUTH_CLIENT_ID}"

gcloud run deploy "$SERVICE" \
  --source . \
  --region "$REGION" \
  --service-account "$SA_EMAIL" \
  --allow-unauthenticated \
  --memory 1Gi \
  --cpu 1 \
  --timeout 300 \
  --max-instances 5 \
  --set-env-vars "$ENV_VARS"

# ---- 6. Prune old images (keep last 5) ------------------------------------
step "Pruning old images from Artifact Registry (keeping last 5)"
IMAGE_PATH="${REGION}-docker.pkg.dev/${PROJECT_ID}/cloud-run-source-deploy/${SERVICE}"
if gcloud artifacts repositories describe cloud-run-source-deploy \
    --location="$REGION" --project "$PROJECT_ID" >/dev/null 2>&1; then
  gcloud artifacts docker images list "$IMAGE_PATH" \
    --sort-by=~CREATE_TIME --format='get(digest)' 2>/dev/null \
    | tail -n +6 \
    | while read -r digest; do
        [ -n "$digest" ] || continue
        echo "  Deleting ${IMAGE_PATH}@${digest}"
        gcloud artifacts docker images delete "${IMAGE_PATH}@${digest}" \
          --quiet --async --project "$PROJECT_ID" 2>/dev/null || true
      done
else
  echo "  No cloud-run-source-deploy repo yet — skipping."
fi

# ---- 7. Report -------------------------------------------------------------
URL="$(gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.url)')"
step "Done"
bold "Service URL: $URL"
echo
echo "Next steps:"
echo "  1. If using auth, add this URL to your OAuth client's Authorized JavaScript origins:"
echo "       $URL"
echo "  2. Re-run with OAUTH_CLIENT_ID=... if you haven't set it yet."
echo "  3. Open $URL and sign in."
