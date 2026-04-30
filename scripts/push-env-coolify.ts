/**
 * Push relevant .env keys to a Coolify application.
 *
 * Usage:
 *   COOLIFY_TOKEN=... COOLIFY_APP_UUID=... ts-node scripts/push-env-coolify.ts
 *
 * Writes one env entry per call (Coolify supports POST /envs and PATCH /envs).
 * Idempotent: tries POST first; on 409/duplicate, falls back to PATCH by key.
 */

import 'dotenv/config';

const KEYS_TO_PUSH = [
  'AZURE_TENANT_ID',
  'AZURE_CLIENT_ID',
  'AZURE_CLIENT_SECRET',
  'GRAPH_REFRESH_TOKEN',
  'ANTHROPIC_API_KEY',
  'WEBHOOK_BASE_URL',
  'WEBHOOK_CLIENT_STATE_SECRET',
  'DISCORD_WEBHOOK_URL',
  'ONEDRIVE_BILLS_PATH',
  'ONEDRIVE_TAX_PATH_PREFIX',
  'ONEDRIVE_ATTACHMENTS_DIR',
  'ONEDRIVE_STATEMENTS_INBOX',
  'ONEDRIVE_STATEMENTS_PROCESSED',
  'ONEDRIVE_TRANSACTIONS_PATH_PREFIX',
  'ONEDRIVE_RENTALS_INBOX',
  'ONEDRIVE_RENTALS_PROCESSED',
  'ONEDRIVE_RENTALS_PATH_PREFIX',
  'ONEDRIVE_PL_DIR',
  'ONEDRIVE_CPA_PACKETS_DIR',
  'ENTITIES',
];

const COOLIFY_BASE = process.env.COOLIFY_BASE || 'http://142.93.182.236:8000/api/v1';
const TOKEN = required('COOLIFY_TOKEN');
const APP = required('COOLIFY_APP_UUID');

function required(name: string): string {
  const v = process.env[name];
  if (!v) {
    console.error(`Missing ${name}`);
    process.exit(1);
  }
  return v;
}

async function pushOne(key: string, value: string): Promise<void> {
  const isSecret =
    /TOKEN|SECRET|KEY|PASSWORD/i.test(key) && key !== 'AZURE_TENANT_ID';
  const body = JSON.stringify({
    key,
    value,
    is_preview: false,
    is_literal: false,
    is_multiline: false,
    is_shown_once: false,
  });
  const headers = {
    Authorization: `Bearer ${TOKEN}`,
    'Content-Type': 'application/json',
  };

  // Try POST; if it's a duplicate, PATCH instead.
  let res = await fetch(`${COOLIFY_BASE}/applications/${APP}/envs`, {
    method: 'POST',
    headers,
    body,
  });
  if (res.status === 409 || res.status === 422) {
    res = await fetch(`${COOLIFY_BASE}/applications/${APP}/envs`, {
      method: 'PATCH',
      headers,
      body,
    });
  }
  const text = await res.text();
  if (!res.ok) {
    console.error(`  ✗ ${key}: HTTP ${res.status} — ${text}`);
    return;
  }
  console.log(`  ✓ ${key} (${isSecret ? '***' : value.slice(0, 40)})`);
}

async function main() {
  console.log(`Pushing ${KEYS_TO_PUSH.length} env vars to Coolify app ${APP}\n`);
  for (const key of KEYS_TO_PUSH) {
    const value = process.env[key];
    if (value === undefined || value === '') {
      console.log(`  · ${key} (empty, skipping)`);
      continue;
    }
    await pushOne(key, value);
  }
  console.log('\n✓ Done.');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
