/**
 * Timer trigger: nightly reconciliation pass.
 *
 * Walks the current and prior tax year, matching transactions to
 * email receipts/bills and updating reconciliation_status.
 */

import { app, InvocationContext, Timer } from '@azure/functions';
import { reconcileYear } from '../lib/reconcile';

export async function reconcileHandler(
  _t: Timer,
  ctx: InvocationContext,
): Promise<void> {
  const now = new Date();
  const years = [now.getFullYear(), now.getFullYear() - 1];

  for (const year of years) {
    try {
      const result = await reconcileYear(year);
      ctx.log(
        `Reconciled ${year}: ${result.matched} matched, ${result.unmatched} unmatched, ${result.transfer} transfers, ${result.needsReview} need review (${result.alreadyMatched} already done)`,
      );
    } catch (err) {
      ctx.error(`Reconcile failed for ${year}`, err);
    }
  }
}

if (process.env.FUNCTIONS_WORKER_RUNTIME) {
  app.timer('reconcile', {
    schedule: '0 30 7 * * *', // 07:30 UTC daily
    handler: reconcileHandler,
  });
}
