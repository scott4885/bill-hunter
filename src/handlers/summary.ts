/**
 * Timer trigger: daily summary of bills due this week.
 * Reads from bills.xlsx via Graph and posts to Discord webhook.
 */

import { app, InvocationContext, Timer } from '@azure/functions';
import { getGraphClient } from '../graph/client';
import { addDays, format } from 'date-fns';

export async function summaryHandler(_t: Timer, ctx: InvocationContext): Promise<void> {
  const billsPath = process.env.ONEDRIVE_BILLS_PATH;
  const discordUrl = process.env.DISCORD_WEBHOOK_URL;
  if (!billsPath || !discordUrl) {
    ctx.warn('Missing ONEDRIVE_BILLS_PATH or DISCORD_WEBHOOK_URL; skipping summary');
    return;
  }

  const client = getGraphClient();
  const res = await client
    .api(`/me/drive/root:${billsPath}:/workbook/tables/Bills/rows`)
    .get();

  const rows: any[] = res.value || [];
  const today = new Date();
  const weekOut = addDays(today, 7);

  // Headers in order from excel.ts BILLS_TABLE
  const HEADERS = [
    'invoice_id',
    'received_date',
    'vendor',
    'amount',
    'due_date',
    'invoice_number',
    'entity',
    'status',
    'email_link',
    'attachment_path',
    'confidence',
    'notes',
  ];

  const dueSoon = rows
    .map((r) => Object.fromEntries(HEADERS.map((h, i) => [h, r.values[0][i]])))
    .filter((r) => r.status === 'unpaid' && r.due_date)
    .filter((r) => {
      const d = new Date(r.due_date);
      return d >= today && d <= weekOut;
    })
    .sort((a, b) => new Date(a.due_date).getTime() - new Date(b.due_date).getTime());

  const total = dueSoon.reduce((s, r) => s + (Number(r.amount) || 0), 0);

  const lines = dueSoon.length
    ? dueSoon
        .map(
          (r) =>
            `• **${r.vendor}** — $${Number(r.amount).toFixed(2)} due ${r.due_date} (${r.entity})`,
        )
        .join('\n')
    : '_No bills due in the next 7 days._';

  const content = `**Bill Hunter — ${format(today, 'EEE, MMM d')}**\n${dueSoon.length} bill${dueSoon.length === 1 ? '' : 's'} due this week, $${total.toFixed(2)} total.\n\n${lines}`;

  await fetch(discordUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ content }),
  });
  ctx.log(`Posted summary: ${dueSoon.length} bills, $${total.toFixed(2)}`);
}

if (process.env.FUNCTIONS_WORKER_RUNTIME) {
  app.timer('dailySummary', {
    schedule: '0 0 13 * * *', // 13:00 UTC = 8am ET / 7am CT (adjust for PCB / FL = ET)
    handler: summaryHandler,
  });
}
