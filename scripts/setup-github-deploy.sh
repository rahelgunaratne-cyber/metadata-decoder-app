#!/usr/bin/env bash
#
# One-time setup so GitHub Actions can deploy to Cloud Run with NO long-lived
# keys, using Workload Identity Federation (GitHub OIDC -> GCP).
#
# It creates a dedicated "deployer" service account (separate from the app's
# runtime SA), grants it the minimum roles to build + push an image and deploy
# Cloud Run, creates an Artifact Registry repo for the images, and wires up a
# Workload Identity pool/provider locked to YOUR GitHub repo. At the end it
# prints the two values to paste into GitHub repo secrets.
#
# Run this AFTER you've deployed once with ./deploy.sh (so the project, APIs,
# bucket, Firestore, and runtime SA already exist).
#
# Usage:
#   GITHUB_REPO=your-org/your-repo PROJECT_ID=decoder-app-500118 ./scripts/setup-github-deploy.sh
#
set -euo pipefail

PROJECT_ID="${PROJECT_ID:-$(gcloud config get-value project 2>/dev/null || true)}"
GITHUB_REPO="${GITHUB_REPO:?Set GITHUB_REPO=owner/repo (e.g. cmg/metadata-decoder)}"
REGION="${REGION:-us-central1}"
REPOSITORY="${REPOSITORY:-decoder}"
RUNTIME_SA="${RUNTIME_SA:-metadata-decoder-run@${PROJECT_ID}.iam.gserviceaccount.com}"
DEPLOYER_SA_NAME="${DEPLOYER_SA_NAME:-github-deployer}"
POOL="${POOL:-github-pool}"
PROVIDER="${PROVIDER:-github-provider}"

[ -n "$PROJECT_ID" ] || { echo "ERROR: no PROJECT_ID"; exit 1; }

DEPLOYER_SA="${DEPLOYER_SA_NAME}@${PROJECT_ID}.iam.gserviceaccount.com"
PROJECT_NUMBER="$(gcloud projects describe "$PROJECT_ID" --format='value(projectNumber)')"

step() { printf '\n\033[1;36m==> %s\033[0m\n' "$*"; }

step "Enabling APIs for keyless auth"
gcloud services enable iamcredentials.googleapis.com sts.googleapis.com \
  artifactregistry.googleapis.com run.googleapis.com --project "$PROJECT_ID"

step "Creating Artifact Registry repo '${REPOSITORY}' (idempotent)"
if gcloud artifacts repositories describe "$REPOSITORY" --location="$REGION" --project "$PROJECT_ID" >/dev/null 2>&1; then
  echo "  Repo already exists."
else
  gcloud artifacts repositories create "$REPOSITORY" \
    --repository-format=docker --location="$REGION" \
    --description="Metadata Decoder container images" --project "$PROJECT_ID"
fi

step "Creating deployer service account (idempotent)"
if gcloud iam service-accounts describe "$DEPLOYER_SA" --project "$PROJECT_ID" >/dev/null 2>&1; then
  echo "  Service account already exists."
else
  gcloud iam service-accounts create "$DEPLOYER_SA_NAME" \
    --display-name="GitHub Actions deployer" --project "$PROJECT_ID"
fi

step "Granting least-privilege deploy roles"
for ROLE in roles/run.admin roles/artifactregistry.writer; do
  gcloud projects add-iam-policy-binding "$PROJECT_ID" \
    --member="serviceAccount:${DEPLOYER_SA}" --role="$ROLE" --condition=None >/dev/null
done
# Needed to deploy a service that RUNS AS the runtime SA.
gcloud iam service-accounts add-iam-policy-binding "$RUNTIME_SA" \
  --member="serviceAccount:${DEPLOYER_SA}" \
  --role="roles/iam.serviceAccountUser" --project "$PROJECT_ID" >/dev/null

step "Creating Workload Identity pool + provider (locked to ${GITHUB_REPO})"
if ! gcloud iam workload-identity-pools describe "$POOL" \
  --location=global --project "$PROJECT_ID" >/dev/null 2>&1; then
  gcloud iam workload-identity-pools create "$POOL" \
    --location=global --display-name="GitHub Actions" --project "$PROJECT_ID"
fi

if ! gcloud iam workload-identity-pools providers describe "$PROVIDER" \
  --location=global --workload-identity-pool="$POOL" --project "$PROJECT_ID" >/dev/null 2>&1; then
  gcloud iam workload-identity-pools providers create-oidc "$PROVIDER" \
    --location=global --workload-identity-pool="$POOL" \
    --display-name="GitHub OIDC" \
    --issuer-uri="https://token.actions.githubusercontent.com" \
    --attribute-mapping="google.subject=assertion.sub,attribute.repository=assertion.repository,attribute.repository_owner=assertion.repository_owner" \
    --attribute-condition="assertion.repository=='${GITHUB_REPO}'" \
    --project "$PROJECT_ID"
fi

step "Allowing ${GITHUB_REPO} to impersonate the deployer SA"
gcloud iam service-accounts add-iam-policy-binding "$DEPLOYER_SA" \
  --role="roles/iam.workloadIdentityUser" \
  --member="principalSet://iam.googleapis.com/projects/${PROJECT_NUMBER}/locations/global/workloadIdentityPools/${POOL}/attribute.repository/${GITHUB_REPO}" \
  --project "$PROJECT_ID" >/dev/null

PROVIDER_RESOURCE="projects/${PROJECT_NUMBER}/locations/global/workloadIdentityPools/${POOL}/providers/${PROVIDER}"

cat <<EOF

============================================================================
Done. Add these as GitHub repo secrets:
  GitHub repo -> Settings -> Secrets and variables -> Actions -> New repository secret

  WIF_PROVIDER         = ${PROVIDER_RESOURCE}
  WIF_SERVICE_ACCOUNT  = ${DEPLOYER_SA}

Then push to 'main' (or run the workflow manually) and it will deploy.
============================================================================
EOF
