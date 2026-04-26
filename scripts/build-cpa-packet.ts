/**
 * Build the CPA export packet for a tax year.
 *
 * Run: npm run build:cpa -- --year=2025
 */

import 'dotenv/config';
import { buildCpaPacket } from '../src/lib/cpa-packet';

async function main() {
  let year = new Date().getFullYear() - 1; // default: prior year
  for (const arg of process.argv.slice(2)) {
    if (arg.startsWith('--year=')) year = parseInt(arg.split('=')[1], 10);
  }

  console.log(`Building CPA packet for ${year}...`);
  const path = await buildCpaPacket(year);
  console.log(`\n✓ Packet ready: ${path}`);
  console.log(`  Open it in OneDrive and share with your CPA.`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
