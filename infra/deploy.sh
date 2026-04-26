#!/usr/bin/env bash
# One-shot deploy: provision Azure infra via Bicep, then publish the Function App.
#
# Prereqs:
#   - az login completed
#   - .env populated (incl. GRAPH_REFRESH_TOKEN from `npm run auth:setup`)
#   - npm install + npm run build run cleanly
#
# Usage:
#   FUNCTION_APP_NAME=fn-billhunter-scott LOCATION=eastus2 RG=rg-billhunter \
#     bash infra/deploy.sh

set -euo pipefail

# Add winget-installed tooling to PATH (Git Bash may not pick these up automatically)
export PATH="/c/Program Files/Microsoft SDKs/Azure/CLI2/wbin:/c/Program Files/Microsoft/Azure Functions Core Tools:$PATH"

cd "$(dirname "$0")/.."

if [ ! -f .env ]; then
  echo "ERROR: .env not found. cp .env.example .env first." >&2
  exit 1
fi

# shellcheck disable=SC2046
read_env() { grep -E "^$1=" .env | cut -d'=' -f2- | tr -d '\r'; }

: "${FUNCTION_APP_NAME:?Set FUNCTION_APP_NAME (must be globally unique)}"
: "${RG:=rg-billhunter}"
: "${LOCATION:=eastus2}"

AZURE_TENANT_ID="$(read_env AZURE_TENANT_ID)"
AZURE_CLIENT_ID="$(read_env AZURE_CLIENT_ID)"
AZURE_CLIENT_SECRET="$(read_env AZURE_CLIENT_SECRET)"
ANTHROPIC_API_KEY="$(read_env ANTHROPIC_API_KEY)"
WEBHOOK_CLIENT_STATE_SECRET="$(read_env WEBHOOK_CLIENT_STATE_SECRET)"
GRAPH_REFRESH_TOKEN="$(read_env GRAPH_REFRESH_TOKEN)"
DISCORD_WEBHOOK_URL="$(read_env DISCORD_WEBHOOK_URL || true)"

for v in AZURE_TENANT_ID AZURE_CLIENT_ID AZURE_CLIENT_SECRET ANTHROPIC_API_KEY \
         WEBHOOK_CLIENT_STATE_SECRET GRAPH_REFRESH_TOKEN; do
  if [ -z "${!v}" ]; then
    echo "ERROR: $v is empty in .env" >&2
    exit 1
  fi
done

echo "==> Ensuring resource group '$RG' in '$LOCATION'..."
az group create -n "$RG" -l "$LOCATION" -o none

echo "==> Deploying Bicep (Function App + Storage + App Insights)..."
az deployment group create \
  --resource-group "$RG" \
  --template-file infra/main.bicep \
  --parameters \
    functionAppName="$FUNCTION_APP_NAME" \
    azureTenantId="$AZURE_TENANT_ID" \
    azureClientId="$AZURE_CLIENT_ID" \
    azureClientSecret="$AZURE_CLIENT_SECRET" \
    anthropicApiKey="$ANTHROPIC_API_KEY" \
    webhookClientStateSecret="$WEBHOOK_CLIENT_STATE_SECRET" \
    graphRefreshToken="$GRAPH_REFRESH_TOKEN" \
    discordWebhookUrl="$DISCORD_WEBHOOK_URL" \
  -o table

WEBHOOK_URL=$(az deployment group show -g "$RG" -n main --query properties.outputs.webhookUrl.value -o tsv)
echo
echo "==> Building TypeScript..."
npm run build

echo "==> Publishing to Function App '$FUNCTION_APP_NAME'..."
AZURE_FUNCTIONAPP_NAME="$FUNCTION_APP_NAME" npm run deploy

# Update WEBHOOK_BASE_URL in local .env so subscribe script targets the deployed URL
BASE_URL="${WEBHOOK_URL%/api/webhook}"
if grep -q '^WEBHOOK_BASE_URL=' .env; then
  sed -i.bak "s|^WEBHOOK_BASE_URL=.*|WEBHOOK_BASE_URL=$BASE_URL|" .env && rm -f .env.bak
else
  printf '\nWEBHOOK_BASE_URL=%s\n' "$BASE_URL" >> .env
fi

echo
echo "✓ Deployed."
echo "  Webhook URL: $WEBHOOK_URL"
echo "  Updated .env with WEBHOOK_BASE_URL=$BASE_URL"
echo
echo "Next: npm run subscribe"
