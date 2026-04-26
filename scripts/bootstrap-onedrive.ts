/**
 * Idempotently create the OneDrive folder structure used by Bill Hunter.
 *
 * Run once after `npm run auth:setup`:
 *   npm run bootstrap:onedrive
 *
 * Reads ONEDRIVE_* env vars and ensures the matching folders exist.
 * The xlsx files themselves are auto-created on first write by src/graph/excel.ts.
 */

import 'dotenv/config';
import { getGraphClient } from '../src/graph/client';

const ROOT = (process.env.ONEDRIVE_ROOT || '/Documents/BillHunter').replace(/\/$/, '');

const FOLDERS = [
  ROOT,
  `${ROOT}/attachments`,
  `${ROOT}/statements`,
  `${ROOT}/statements/inbox`,
  `${ROOT}/statements/processed`,
  `${ROOT}/rentals`,
  `${ROOT}/rentals/inbox`,
  `${ROOT}/rentals/processed`,
  `${ROOT}/pl`,
  `${ROOT}/cpa_packet`,
];

async function ensureFolder(absPath: string): Promise<void> {
  const client = getGraphClient();
  const segments = absPath.split('/').filter(Boolean);
  const name = segments[segments.length - 1];
  const parentPath = '/' + segments.slice(0, -1).join('/');

  const parentRef =
    parentPath === '/' ? '/me/drive/root' : `/me/drive/root:${parentPath}:`;

  try {
    await client.api(`${parentRef}/children`).post({
      name,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'fail',
    });
    console.log(`  + created ${absPath}`);
  } catch (err: any) {
    const code = err?.statusCode ?? err?.code;
    if (code === 409 || err?.code === 'nameAlreadyExists') {
      console.log(`  · exists  ${absPath}`);
      return;
    }
    throw err;
  }
}

async function main() {
  console.log(`Bootstrapping OneDrive folder structure under ${ROOT}\n`);
  for (const folder of FOLDERS) {
    await ensureFolder(folder);
  }
  console.log('\n✓ OneDrive folder structure ready.');
  console.log('\nDrop credit-card / bank statements into:');
  console.log(`  ${ROOT}/statements/inbox`);
  console.log('Drop Airbnb / Vrbo / Guesty CSVs into:');
  console.log(`  ${ROOT}/rentals/inbox`);
}

main().catch((e) => {
  console.error('Bootstrap failed:', e?.message ?? e);
  if (e?.body) console.error(e.body);
  process.exit(1);
});
