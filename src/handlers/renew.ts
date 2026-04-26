/**
 * Timer trigger: renew Graph subscriptions before they expire.
 *
 * Mail subscriptions have a max expiration of ~71 hours (3 days).
 * We renew daily to be safe.
 */

import { app, InvocationContext, Timer } from '@azure/functions';
import { listSubscriptions, renewSubscription } from '../graph/client';

export async function renewHandler(_t: Timer, ctx: InvocationContext): Promise<void> {
  const subs = await listSubscriptions();
  const newExpiration = new Date(Date.now() + 70 * 60 * 60 * 1000).toISOString();

  for (const sub of subs) {
    try {
      await renewSubscription(sub.id, newExpiration);
      ctx.log(`Renewed ${sub.id} (${sub.resource}) until ${newExpiration}`);
    } catch (err) {
      ctx.error(`Failed to renew ${sub.id}`, err);
    }
  }
}

if (process.env.FUNCTIONS_WORKER_RUNTIME) {
  app.timer('renewSubscriptions', {
    schedule: '0 0 6 * * *', // Daily at 06:00 UTC
    handler: renewHandler,
  });
}
