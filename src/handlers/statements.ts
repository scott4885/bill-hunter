/**
 * Timer trigger: scan ONEDRIVE_STATEMENTS_INBOX every 15 minutes for new
 * PDFs. For each, run Statement Parser, append transactions, and move
 * the PDF to the processed folder.
 *
 * Drop your monthly statements into the inbox folder; this handles them.
 */

import { app, InvocationContext, Timer } from '@azure/functions';
import {
  listFolder,
  downloadFile,
  moveFile,
  appendTransactionRow,
  transactionExists,
} from '../graph/excel';
import { parseStatement, buildTxnId } from '../agents/statement-parser';

export async function statementsHandler(
  _t: Timer,
  ctx: InvocationContext,
): Promise<void> {
  const inbox = process.env.ONEDRIVE_STATEMENTS_INBOX;
  const processed = process.env.ONEDRIVE_STATEMENTS_PROCESSED;
  if (!inbox || !processed) {
    ctx.warn('ONEDRIVE_STATEMENTS_INBOX or _PROCESSED not set; skipping');
    return;
  }

  const files = await listFolder(inbox);
  const pdfs = files.filter((f) => f.name.toLowerCase().endsWith('.pdf'));
  if (pdfs.length === 0) {
    ctx.log('No new statements to process.');
    return;
  }
  ctx.log(`Processing ${pdfs.length} statement(s)`);

  for (const file of pdfs) {
    const filePath = `${inbox}/${file.name}`;
    try {
      const buf = await downloadFile(filePath);
      const b64 = buf.toString('base64');
      const result = await parseStatement(b64, file.name);

      ctx.log(
        `Parsed ${file.name}: ${result.account_label}, ${result.transactions.length} txns`,
      );

      let inserted = 0;
      let skipped = 0;
      for (const txn of result.transactions) {
        const year = parseInt(txn.date.slice(0, 4), 10);
        const txnId = buildTxnId(result.account_label, txn);

        if (await transactionExists(year, txnId)) {
          skipped++;
          continue;
        }

        await appendTransactionRow(year, {
          txn_id: txnId,
          date: txn.date,
          description: txn.description,
          amount: txn.amount,
          account: result.account_label,
          statement_path: filePath,
          entity: txn.entity,
          schedule_c_category: txn.schedule_c_category,
          deductible: txn.deductible,
          matched_receipt_path: '',
          matched_bill_id: '',
          reconciliation_status: 'unmatched',
          confidence: txn.confidence,
          notes: txn.flags.join(','),
        });
        inserted++;
      }
      ctx.log(`  ${inserted} inserted, ${skipped} dedup-skipped`);

      // Move the processed PDF out of the inbox
      await moveFile(filePath, processed);
    } catch (err) {
      ctx.error(`Failed processing ${file.name}: ${(err as Error).message}`);
    }
  }
}

if (process.env.FUNCTIONS_WORKER_RUNTIME) {
  app.timer('processStatements', {
    schedule: '0 */15 * * * *', // every 15 min
    handler: statementsHandler,
  });
}
