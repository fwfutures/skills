#!/bin/bash
set -euo pipefail

# GCP Cloud Run deploy-from-source script
# Usage: deploy.sh [SERVICE_NAME]
# Env vars: GCP_REGION (default: australia-southeast1), GCP_PROJECT (default: current gcloud project)

SERVICE_NAME="${1:-my-service}"
REGION="${GCP_REGION:-australia-southeast1}"
PROJECT_ID="${GCP_PROJECT:-$(gcloud config get-value project 2>/dev/null)}"

if [ -z "$PROJECT_ID" ]; then
  echo "Error: No GCP project set. Run: gcloud config set project PROJECT_ID"
  exit 1
fi

echo "Deploying '$SERVICE_NAME' to $PROJECT_ID ($REGION)..."

gcloud run deploy "$SERVICE_NAME" \
  --source . \
  --region "$REGION" \
  --project "$PROJECT_ID" \
  --allow-unauthenticated \
  --quiet

echo ""
echo "Deployed. URL:"
gcloud run services describe "$SERVICE_NAME" \
  --region "$REGION" \
  --project "$PROJECT_ID" \
  --format "value(status.url)"
