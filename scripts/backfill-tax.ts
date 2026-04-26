/**
 * Tax backfill: sweep historical email by year and run Tax Archivist on each.
 *
 * Run: `npm run backfill -- --year=2024`
 *      `npm run backfill -- --year=2025 --start=2025-01-01 --end=2025-12-31`
 *
 * Designed for occasional/quarterly use. Pages through email in chunks,
 * skips messages already classified (by message id), and writes to the
 * appropriate taxes_YYYY.xlsx file.
 *
 * Cost note: each email = 1 Claude call. ~5000 emails = ~$15-25 on Opus 4.7.
 * For pure backfill, consider switching the model to claude-sonnet-4-6 in
 * src/agents/tax-archivist.ts to cut cost ~5x.
 */

import 'dotenv/config';
import { listMessagesByDateRange, getMessage, getAttachments } from '../src/graph/client';
import { classifyForTax } from '../src/agents/tax-archivist';
import { appendTaxRow, uploadAttachment } from '../src/graph/excel';

interface Args {
  year: number;
  start?: string;
  end?: string;
  dryRun: boolean;
  limit?: number;
}

function parseArgs(): Args {
  const args: Partial<Args> = { dryRun: false };
  for (const arg of process.argv.slice(2)) {
    if (arg.startsWith('--year=')) args.year = parseInt(arg.split('=')[1], 10);
    else if (arg.startsWith('--start=')) args.start = arg.split('=')[1];
    else if (arg.startsWith('--end=')) args.end = arg.split('=')[1];
    else if (arg.startsWith('--limit=')) args.limit = parseInt(arg.split('=')[1], 10);
    else if (arg === '--dry-run') args.dryRun = true;
  }
  if (!args.year) throw new Error('Required: --year=YYYY');
  return args as Args;
}

async function main() {
  const args = parseArgs();
  const start = args.start ?? `${args.year}-01-01T00:00:00Z`;
  const end = args.end ?? `${args.year}-12-31T23:59:59Z`;

  console.log(`Backfill: ${start} → ${end}${args.dryRun ? ' (dry run)' : ''}`);

  const messages = await listMessagesByDateRange(start, end);
  console.log(`Found ${messages.length} messages to process.`);

  const limit = args.limit ?? messages.length;
  let processed = 0;
  let logged = 0;
  let skipped = 0;
  let errors = 0;

  for (const stub of messages.slice(0, limit)) {
    processed++;
    try {
      // Re-fetch to get full body
      const email = await getMessage(stub.id);
      const attachments = email.hasAttachments ? await getAttachments(stub.id) : [];

      const classification = await classifyForTax(email, attachments);

      if (!classification.is_tax_relevant) {
        skipped++;
        continue;
      }

      if (args.dryRun) {
        console.log(
          `  [dry] ${classification.vendor} $${classification.amount} → ${classification.schedule_c_category} / ${classification.entity}`,
        );
        logged++;
        continue;
      }

      const txDate = classification.date ?? email.receivedDateTime.slice(0, 10);
      const year = parseInt(txDate.slice(0, 4), 10);

      let receiptPath = '';
      for (const att of attachments) {
        if (att.contentBytes && att.contentType === 'application/pdf') {
          receiptPath = await uploadAttachment(att.name, att.contentBytes);
          break;
        }
      }

      await appendTaxRow(year, {
        date: txDate,
        vendor: classification.vendor ?? '',
        amount: classification.amount ?? '',
        schedule_c_category: classification.schedule_c_category,
        entity: classification.entity,
        deductible: classification.deductible,
        business_purpose: classification.business_purpose,
        receipt_path: receiptPath,
        email_link: email.webLink,
        is_1099_candidate: classification.is_1099_candidate ? 'yes' : 'no',
        confidence: classification.confidence,
        notes: classification.flags.join(','),
      });
      logged++;
      if (logged % 20 === 0) console.log(`  ... ${logged} logged, ${processed} processed`);
    } catch (err) {
      errors++;
      console.error(`  ✗ Error on ${stub.id}: ${(err as Error).message}`);
    }

    // Light rate limit — Claude API + Graph
    await sleep(250);
  }

  console.log('\n=== Backfill complete ===');
  console.log(`Processed: ${processed}`);
  console.log(`Tax-logged: ${logged}`);
  console.log(`Skipped (not tax-relevant): ${skipped}`);
  console.log(`Errors: ${errors}`);
}

function sleep(ms: number) {
  return new Promise((r) => setTimeout(r, ms));
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
