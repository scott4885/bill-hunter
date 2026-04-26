/**
 * Timer trigger: scan ONEDRIVE_RENTALS_INBOX for new CSV exports
 * (Airbnb, Vrbo, Guesty), normalize, and append to rentals_YYYY.xlsx.
 */

import { app, InvocationContext, Timer } from '@azure/functions';
import {
  listFolder,
  downloadFile,
  moveFile,
  appendRentalRow,
  rentalExists,
} from '../graph/excel';
import { detectFormat, parseCsv, normalize } from '../lib/rental-import';

export async function rentalsHandler(
  _t: Timer,
  ctx: InvocationContext,
): Promise<void> {
  const inbox = process.env.ONEDRIVE_RENTALS_INBOX;
  const processed = process.env.ONEDRIVE_RENTALS_PROCESSED;
  if (!inbox || !processed) {
    ctx.warn('ONEDRIVE_RENTALS_INBOX or _PROCESSED not set; skipping');
    return;
  }

  const files = await listFolder(inbox);
  const csvs = files.filter((f) => f.name.toLowerCase().endsWith('.csv'));
  if (csvs.length === 0) {
    ctx.log('No new rental CSVs to process.');
    return;
  }

  for (const file of csvs) {
    const filePath = `${inbox}/${file.name}`;
    try {
      const buf = await downloadFile(filePath);
      const content = buf.toString('utf-8');
      const { headers, rows } = parseCsv(content);
      const format = detectFormat(headers);

      if (format === 'unknown') {
        ctx.warn(
          `Unknown format for ${file.name}; headers: ${headers.slice(0, 5).join(',')}...`,
        );
        continue;
      }

      let inserted = 0;
      let skipped = 0;
      for (const row of rows) {
        const norm = normalize(format, headers, row);
        if (!norm || !norm.booking_id) continue;
        const year = parseInt(norm.payout_date.slice(0, 4), 10);
        if (!year) continue;

        if (await rentalExists(year, norm.booking_id)) {
          skipped++;
          continue;
        }

        await appendRentalRow(year, norm as unknown as Record<string, unknown>);
        inserted++;
      }
      ctx.log(`${file.name} (${format}): ${inserted} inserted, ${skipped} dup`);

      await moveFile(filePath, processed);
    } catch (err) {
      ctx.error(`Failed processing ${file.name}: ${(err as Error).message}`);
    }
  }
}

if (process.env.FUNCTIONS_WORKER_RUNTIME) {
  app.timer('processRentals', {
    schedule: '0 */30 * * * *', // every 30 min
    handler: rentalsHandler,
  });
}
