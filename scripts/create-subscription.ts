/**
 * Create the Graph webhook subscription pointing at the deployed Function.
 *
 * Run AFTER:
 *  1. `npm run auth:setup` (have refresh token)
 *  2. Function App is deployed and reachable at WEBHOOK_BASE_URL
 *
 * Run: `npm run subscribe`
 */

import 'dotenv/config';
import { createSubscription, listSubscriptions } from '../src/graph/client';

async function main() {
  const baseUrl = required('WEBHOOK_BASE_URL');
  const clientState = required('WEBHOOK_CLIENT_STATE_SECRET');
  const notificationUrl = `${baseUrl.replace(/\/$/, '')}/api/webhook`;

  console.log('Existing subscriptions:');
  const existing = await listSubscriptions();
  existing.forEach((s) =>
    console.log(`  - ${s.id} → ${s.resource} (expires ${s.expirationDateTime})`),
  );

  const expirationDateTime = new Date(Date.now() + 70 * 60 * 60 * 1000).toISOString();

  const result = await createSubscription({
    changeType: 'created',
    notificationUrl,
    resource: "me/mailFolders('inbox')/messages",
    expirationDateTime,
    clientState,
  });

  console.log('\n✓ Subscription created:');
  console.log(`  id: ${result.id}`);
  console.log(`  resource: ${result.resource}`);
  console.log(`  expires: ${result.expirationDateTime}`);
  console.log(`  notificationUrl: ${notificationUrl}`);
  console.log('\nThe Function will receive a validation handshake during creation.');
  console.log('Renewal handler runs daily at 06:00 UTC.');
}

function required(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Missing env var: ${name}`);
  return v;
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
