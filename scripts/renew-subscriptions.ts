/**
 * Manually renew Graph subscriptions.
 *
 * Normally runs automatically via the timer-triggered handler, but useful
 * for the first deploy or recovery.
 *
 * Run: npm run renew
 */

import 'dotenv/config';
import { listSubscriptions, renewSubscription } from '../src/graph/client';

async function main() {
  const subs = await listSubscriptions();
  console.log(`Found ${subs.length} subscription(s).`);

  const newExpiration = new Date(Date.now() + 70 * 60 * 60 * 1000).toISOString();

  for (const sub of subs) {
    try {
      await renewSubscription(sub.id, newExpiration);
      console.log(`✓ ${sub.id} (${sub.resource}) → ${newExpiration}`);
    } catch (err) {
      console.error(`✗ ${sub.id}: ${(err as Error).message}`);
    }
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
