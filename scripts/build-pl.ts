/**
 * Build per-entity P&L workbooks.
 *
 * Run: npm run build:pl -- --year=2026
 *      npm run build:pl -- --year=2026 --entity="Sol Haus"
 */

import 'dotenv/config';
import { buildPL } from '../src/lib/pl-builder';

type Entity = 'NHP-1941' | 'NHDPM' | 'Sol Haus' | 'Personal';

async function main() {
  let year = new Date().getFullYear();
  let entity: Entity | null = null;

  for (const arg of process.argv.slice(2)) {
    if (arg.startsWith('--year=')) year = parseInt(arg.split('=')[1], 10);
    else if (arg.startsWith('--entity=')) entity = arg.split('=')[1] as Entity;
  }

  const entities: Entity[] = entity
    ? [entity]
    : ['NHP-1941', 'NHDPM', 'Sol Haus', 'Personal'];

  for (const e of entities) {
    try {
      const path = await buildPL({ entity: e, year });
      console.log(`✓ ${e}: ${path}`);
    } catch (err) {
      console.error(`✗ ${e}: ${(err as Error).message}`);
    }
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
