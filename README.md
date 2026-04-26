# Bill Hunter + Tax Archivist

Real-time bookkeeping infrastructure for Outlook 365 + OneDrive Excel.

## What it does

Six pipelines, one Azure Function App:

| Pipeline | Trigger | Output |
|---|---|---|
| **Bill Hunter** | Email webhook (real-time) | `bills.xlsx` — unpaid bills with due dates |
| **Tax Archivist** | Email webhook (real-time) | `taxes_YYYY.xlsx` — categorized expenses |
| **Statement Parser** | OneDrive folder watcher (15 min) | `transactions_YYYY.xlsx` — every CC/bank line item |
| **Rental Importer** | OneDrive folder watcher (30 min) | `rentals_YYYY.xlsx` — Airbnb/Vrbo/Guesty bookings |
| **Reconciliation** | Nightly timer | matches transactions ↔ receipts |
| **Daily Summary** | Daily timer | Discord post with bills due this week |

Plus on-demand:

- **P&L builder** — per-entity P&Ls (NHP-1941, NHDPM, Sol Haus, Personal)
- **CPA packet** — full year-end bundle (CSVs + P&Ls + receipts + 1099 list)

## Architecture

```
Outlook 365 ──webhook──▶ Azure Function ──▶ bills.xlsx + taxes_YYYY.xlsx
                                            │
OneDrive: statements/inbox  ──watcher──────▶│
OneDrive: rentals/inbox     ──watcher──────▶│
                                            │
                                            ▼
                              transactions / rentals
                                            │
                              Reconciler (nightly)
                                            │
                                            ▼
                              P&L builder (on demand)
                                            │
                                            ▼
                              CPA packet (year-end)
```

## Setup

### 0. Prerequisites

- Node 20+
- Azure CLI + Azure Functions Core Tools v4 (`npm i -g azure-functions-core-tools@4 --unsafe-perm true`)
- An Azure subscription
- Anthropic API key
- Microsoft 365 personal/business account (the one you want monitored)

### 1. Azure AD app registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name: `BillHunter-TaxArchivist`
3. Supported accounts: **Single tenant**
4. Redirect URI (Web): `http://localhost:8000/auth/callback`
5. Click Register
6. Copy the **Application (client) ID** and **Directory (tenant) ID**
7. **Certificates & secrets** → New client secret → copy the value (shown once)
8. **API permissions** → Add → Microsoft Graph → **Delegated**:
   - `Mail.Read`
   - `Mail.ReadBasic`
   - `Files.ReadWrite`
   - `User.Read`
   - `offline_access`
9. Click **Grant admin consent**

### 2. Local install + auth

```bash
git clone <this repo>
cd bill-hunter
npm install
cp .env.example .env
cp local.settings.json.example local.settings.json
```

Edit `.env`:
- `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` — from step 1
- `ANTHROPIC_API_KEY` — from console.anthropic.com
- `WEBHOOK_CLIENT_STATE_SECRET` — generate with `openssl rand -hex 32`

First-time auth (browser flow, ~30 seconds):

```bash
npm run auth:setup
```

Verify it works:

```bash
npm run test:fetch
```

You should see your last 10 emails. ✓

### 3. Create OneDrive folders

In OneDrive, create these (or run `mkdir` via Files app):

```
/Documents/BillHunter/
  attachments/
  statements/
    inbox/
    processed/
  rentals/
    inbox/
    processed/
  pl/
```

The .xlsx files are auto-created on first write.

### 4. Deploy to Azure

```bash
# Create resources (one-time)
az login
az group create -n rg-billhunter -l eastus2
az storage account create -n stbillhunter$RANDOM -g rg-billhunter -l eastus2 --sku Standard_LRS
az functionapp create \
  --resource-group rg-billhunter \
  --consumption-plan-location eastus2 \
  --runtime node --runtime-version 20 \
  --functions-version 4 \
  --name fn-billhunter-$USER \
  --storage-account <storage-account-name>

# Push env vars to Function App
az functionapp config appsettings set \
  --name fn-billhunter-$USER \
  --resource-group rg-billhunter \
  --settings @<(grep -v '^#' .env | grep -v '^$' | sed 's/^/"/;s/=/=/;s/$/"/')

# Build + deploy
npm run build
export AZURE_FUNCTIONAPP_NAME=fn-billhunter-$USER
npm run deploy
```

Update `.env` with the deployed URL:

```
WEBHOOK_BASE_URL=https://fn-billhunter-yourname.azurewebsites.net
```

### 5. Subscribe to email notifications

```bash
npm run subscribe
```

Microsoft Graph posts a validation handshake to your function during creation; you should see it in the function logs immediately.

That's it. New emails now flow through both agents in real-time.

## Daily use

### Drop statements

Save credit card / bank statement PDFs to `/Documents/BillHunter/statements/inbox/`. Within 15 minutes, transactions land in `transactions_YYYY.xlsx`.

### Drop rental exports

Export from Airbnb/Vrbo/Guesty as CSV → drop into `/Documents/BillHunter/rentals/inbox/`. Within 30 minutes, bookings land in `rentals_YYYY.xlsx`.

### Mark bills paid

Open `bills.xlsx`, change `status` from `unpaid` to `paid`. The daily summary respects this.

### Daily summary

If `DISCORD_WEBHOOK_URL` is set, you get a daily post with bills due this week.

## Periodic tasks

### Backfill historical email for taxes

```bash
npm run backfill:tax -- --year=2025
npm run backfill:tax -- --year=2024 --limit=500   # test with a small batch first
npm run backfill:tax -- --year=2024 --dry-run     # see what would happen
```

Cost: ~$0.05–$0.15 per email on Opus 4.7. ~5000 emails ≈ $25-75. Switch to Sonnet 4.6 in `tax-archivist.ts` for ~5x cheaper.

### Generate P&Ls

```bash
npm run build:pl -- --year=2026
npm run build:pl -- --year=2026 --entity="Sol Haus"
```

### Build CPA packet

```bash
npm run build:cpa -- --year=2025
```

Output: `/Documents/BillHunter/cpa_packet_2025/` — share this folder with your CPA.

## File map

```
src/
  auth/
    tokens.ts              OAuth refresh-token manager
  graph/
    client.ts              Graph API wrapper (messages, attachments, subscriptions)
    excel.ts               OneDrive Excel writer (4 table types + helpers)
  agents/
    bill-hunter.ts         Email → invoice extraction
    tax-archivist.ts       Email → Schedule C classification
    statement-parser.ts    PDF statement → transaction line items
  handlers/
    webhook.ts             Real-time email handler (Bill Hunter + Tax Archivist)
    statements.ts          Statement folder watcher
    rentals.ts             Rental CSV folder watcher
    reconcile.ts           Nightly reconciliation timer
    summary.ts             Daily Discord summary
    renew.ts               Subscription renewal timer
  lib/
    empty-xlsx.ts          Bootstrap blank xlsx
    reconcile.ts           Match transactions ↔ receipts
    pl-builder.ts          Per-entity P&L generation
    cpa-packet.ts          Year-end CPA bundle
    rental-import.ts       Airbnb/Vrbo/Guesty CSV parser

scripts/
  auth-setup.ts            One-time browser OAuth flow
  create-subscription.ts   Subscribe to Graph webhooks
  renew-subscriptions.ts   Manual renewal
  test-fetch.ts            Smoke test
  backfill-tax.ts          Historical email sweep
  reconcile.ts             Manual reconciliation run
  build-pl.ts              On-demand P&L
  build-cpa-packet.ts      Year-end packet
  build-empty-xlsx.ts      Regenerate empty xlsx constant

agents/
  bill-hunter.md           Claude Code subagent definition
  tax-archivist.md         Claude Code subagent definition
```

## Cost notes

- Azure Functions Consumption plan: free for low volume; you'll pay <$5/mo
- OneDrive storage: comes with M365
- Anthropic API: ~$0.05–0.15 per email processed (Opus 4.7); statement parsing similar
- Estimate: $20-50/mo at moderate volume, dominated by API costs

## Troubleshooting

**Webhook not firing**: check subscription is alive — `az functionapp log tail -n fn-billhunter-X -g rg-billhunter`. Subscriptions expire every ~3 days; the renewal timer should keep them current.

**"Token refresh failed"**: refresh token expired (90 days inactivity). Re-run `npm run auth:setup`.

**"Missing env var"**: in production, env vars must be set as Function App settings, not in `.env`. See deploy step 4.

**Statement parsing wrong**: check the prompt in `src/agents/statement-parser.ts`. PDF quality varies wildly; some scanned statements need OCR upstream.

**Bills double-logged**: dedup uses `vendor + amount + invoice_number` hash. If your vendor sends invoices without unique numbers, dedupe falls back to due_date. Adjust `buildInvoiceId` in `src/agents/bill-hunter.ts` if needed.

## Next steps for v2

When you move this into Claude Code for iteration:

- Plaid integration for real-time transactions (replace statement parser)
- 1099-NEC form pre-fill from candidates list
- Monthly close email summary (P&L diff vs prior month, anomalies)
- Categorization confidence threshold tuning (auto-accept >0.85, queue 0.7-0.85, reject <0.7)
- Multi-property support if NHP-1941 expands
- Mileage tracker integration (Schedule C line 9)
