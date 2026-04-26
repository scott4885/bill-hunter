/**
 * Manually run reconciliation (normally runs nightly).
 *
 * Run: npm run reconcile -- --year=2026
 */

import 'dotenv/config';
import { reconcileYear } from '../src/lib/reconcile';

async function main() {
  let year = new Date().getFullYear();
  for (const arg of process.argv.slice(2)) {
    if (arg.startsWith('--year=')) year = parseInt(arg.split('=')[1], 10);
  }

  console.log(`Reconciling ${year}...`);
  const r = await reconcileYear(year);
  console.log(
    `\nResults: ${r.matched} matched, ${r.unmatched} unmatched, ${r.transfer} transfers, ${r.needsReview} need review (${r.alreadyMatched} already done)`,
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
