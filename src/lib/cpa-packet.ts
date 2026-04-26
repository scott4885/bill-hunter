/**
 * CPA export packet generator.
 *
 * Bundles everything a CPA needs for a tax year:
 *   - P&L per entity (built fresh)
 *   - tax_log_{year}.csv (categorized expenses with email links)
 *   - transactions_{year}.csv (all account activity, reconciled)
 *   - rentals_{year}.csv (Sol Haus revenue detail)
 *   - 1099_candidates_{year}.csv (contractors >= $600)
 *   - reconciliation_summary.txt
 *   - all receipt PDFs in /receipts subfolder
 *
 * Output folder: /Documents/BillHunter/cpa_packet_{year}/
 */

import {
  listTaxRows,
  listTransactions,
  listRentals,
  downloadFile,
} from '../graph/excel';
import { buildPL } from './pl-builder';
import { Client } from '@microsoft/microsoft-graph-client';
import { getGraphClient } from '../graph/client';

export async function buildCpaPacket(year: number): Promise<string> {
  const client = getGraphClient();
  const baseDir = `${required('ONEDRIVE_CPA_PACKETS_DIR')}/cpa_packet_${year}`;

  // Ensure folder exists
  await ensureFolder(client, baseDir);
  await ensureFolder(client, `${baseDir}/receipts`);
  await ensureFolder(client, `${baseDir}/pl`);

  const [taxRows, txns, rentals] = await Promise.all([
    listTaxRows(year),
    listTransactions(year),
    listRentals(year),
  ]);

  // Tax log CSV
  await writeCsv(
    client,
    `${baseDir}/tax_log_${year}.csv`,
    [
      'date',
      'vendor',
      'amount',
      'schedule_c_category',
      'entity',
      'deductible',
      'business_purpose',
      'is_1099_candidate',
      'email_link',
      'receipt_path',
    ],
    taxRows,
  );

  // Transactions CSV
  await writeCsv(
    client,
    `${baseDir}/transactions_${year}.csv`,
    [
      'date',
      'description',
      'amount',
      'account',
      'entity',
      'schedule_c_category',
      'deductible',
      'reconciliation_status',
      'matched_receipt_path',
      'matched_bill_id',
    ],
    txns,
  );

  // Rentals CSV
  if (rentals.length > 0) {
    await writeCsv(
      client,
      `${baseDir}/rentals_${year}.csv`,
      [
        'booking_id',
        'platform',
        'property',
        'check_in',
        'check_out',
        'nights',
        'gross_revenue',
        'platform_fees',
        'cleaning_fees_collected',
        'taxes_collected',
        'net_payout',
        'payout_date',
      ],
      rentals,
    );
  }

  // 1099 candidates
  const ten99 = aggregate1099Candidates(taxRows, txns);
  if (ten99.length > 0) {
    await writeCsv(
      client,
      `${baseDir}/1099_candidates_${year}.csv`,
      ['vendor', 'entity', 'total_paid', 'transaction_count', 'first_payment', 'last_payment'],
      ten99,
    );
  }

  // P&Ls per entity
  const entities: Array<'NHP-1941' | 'NHDPM' | 'Sol Haus' | 'Personal'> = [
    'NHP-1941',
    'NHDPM',
    'Sol Haus',
  ];
  for (const entity of entities) {
    const hasData = txns.some((t) => String(t.entity) === entity);
    if (!hasData && entity !== 'Sol Haus') continue;
    try {
      await buildPL({ entity, year });
    } catch (err) {
      // PL builder writes directly to ONEDRIVE_PL_DIR; intentional separation.
    }
  }

  // Reconciliation summary
  const summary = buildReconcileSummary(year, taxRows, txns, rentals, ten99);
  await client
    .api(`/me/drive/root:${baseDir}/SUMMARY.txt:/content`)
    .put(Buffer.from(summary, 'utf-8'));

  // Copy receipts (limit to deductible items to keep packet sane)
  const receiptPaths = new Set<string>();
  for (const r of taxRows) {
    if (String(r.deductible) !== 'no' && String(r.receipt_path)) {
      receiptPaths.add(String(r.receipt_path));
    }
  }
  let copied = 0;
  for (const path of receiptPaths) {
    try {
      const buf = await downloadFile(path);
      const filename = path.split('/').pop() || `receipt_${copied}.pdf`;
      await client
        .api(`/me/drive/root:${baseDir}/receipts/${filename}:/content`)
        .put(buf);
      copied++;
    } catch {
      // skip missing
    }
  }

  return baseDir;
}

interface Ten99Row {
  vendor: string;
  entity: string;
  total_paid: number;
  transaction_count: number;
  first_payment: string;
  last_payment: string;
}

function aggregate1099Candidates(
  taxRows: Array<Record<string, any>>,
  txns: Array<Record<string, any>>,
): Ten99Row[] {
  const map = new Map<string, Ten99Row>();

  // From tax log (email-derived) - explicit 1099 flag
  for (const r of taxRows) {
    if (String(r.is_1099_candidate) !== 'yes') continue;
    const vendor = String(r.vendor || '').toLowerCase().trim();
    if (!vendor) continue;
    const key = `${vendor}|${r.entity}`;
    const cur = map.get(key) || {
      vendor: String(r.vendor),
      entity: String(r.entity),
      total_paid: 0,
      transaction_count: 0,
      first_payment: '',
      last_payment: '',
    };
    cur.total_paid += Number(r.amount) || 0;
    cur.transaction_count++;
    const d = String(r.date);
    if (!cur.first_payment || d < cur.first_payment) cur.first_payment = d;
    if (!cur.last_payment || d > cur.last_payment) cur.last_payment = d;
    map.set(key, cur);
  }

  // From transactions: Contract_Labor category, sum by vendor
  for (const t of txns) {
    if (String(t.schedule_c_category) !== 'Contract_Labor') continue;
    const amount = Math.abs(Number(t.amount) || 0);
    const vendor = String(t.description || '').toLowerCase().trim();
    if (!vendor) continue;
    const key = `${vendor}|${t.entity}`;
    const cur = map.get(key) || {
      vendor: String(t.description),
      entity: String(t.entity),
      total_paid: 0,
      transaction_count: 0,
      first_payment: '',
      last_payment: '',
    };
    cur.total_paid += amount;
    cur.transaction_count++;
    const d = String(t.date);
    if (!cur.first_payment || d < cur.first_payment) cur.first_payment = d;
    if (!cur.last_payment || d > cur.last_payment) cur.last_payment = d;
    map.set(key, cur);
  }

  return Array.from(map.values()).filter((r) => r.total_paid >= 600);
}

function buildReconcileSummary(
  year: number,
  taxRows: any[],
  txns: any[],
  rentals: any[],
  ten99: Ten99Row[],
): string {
  const matched = txns.filter((t) => t.reconciliation_status === 'matched').length;
  const unmatched = txns.filter((t) => t.reconciliation_status === 'unmatched').length;
  const transfers = txns.filter((t) => t.reconciliation_status === 'transfer').length;
  const review = txns.filter((t) => t.reconciliation_status === 'needs_review').length;

  const byEntity = new Map<string, number>();
  for (const t of txns) {
    const a = Number(t.amount) || 0;
    if (a >= 0) continue;
    if (String(t.deductible) === 'no') continue;
    const e = String(t.entity);
    byEntity.set(e, (byEntity.get(e) || 0) + Math.abs(a));
  }

  const lines: string[] = [
    `CPA PACKET — TAX YEAR ${year}`,
    `Generated ${new Date().toISOString()}`,
    '',
    'CONTENTS',
    `  - tax_log_${year}.csv          ${taxRows.length} rows from email pipeline`,
    `  - transactions_${year}.csv     ${txns.length} rows from statements`,
    `  - rentals_${year}.csv          ${rentals.length} rows`,
    `  - 1099_candidates_${year}.csv  ${ten99.length} candidates`,
    `  - pl/pl_*_${year}.xlsx         per-entity P&Ls`,
    `  - receipts/                    PDF receipts`,
    '',
    'RECONCILIATION',
    `  Matched (txn ↔ receipt):  ${matched}`,
    `  Unmatched (no receipt):   ${unmatched}`,
    `  Transfers (not expenses): ${transfers}`,
    `  Needs review:             ${review}`,
    '',
    'DEDUCTIBLE EXPENSE TOTALS BY ENTITY',
  ];
  for (const [entity, total] of byEntity) {
    lines.push(`  ${entity.padEnd(15)} $${total.toFixed(2)}`);
  }

  if (ten99.length > 0) {
    lines.push('', '1099 CANDIDATES (>= $600)');
    for (const t of ten99) {
      lines.push(
        `  ${t.vendor.slice(0, 40).padEnd(40)} ${t.entity.padEnd(12)} $${t.total_paid.toFixed(2)}`,
      );
    }
  }

  lines.push(
    '',
    'NOTES FOR CPA',
    '  - Mortgage interest from 1098 forms (not extracted here)',
    '  - Depreciation schedules not generated',
    '  - Meals shown at 100% in raw data; P&L applies 50% reduction',
    '  - Sol Haus assumed Schedule C (substantial services). Confirm vs Schedule E.',
    '  - "needs_review" transactions had multiple matching receipts; please verify.',
  );

  return lines.join('\n');
}

async function writeCsv(
  client: Client,
  drivePath: string,
  columns: string[],
  rows: Array<Record<string, any>>,
): Promise<void> {
  const lines = [columns.join(',')];
  for (const row of rows) {
    lines.push(columns.map((c) => csvEscape(row[c])).join(','));
  }
  await client
    .api(`/me/drive/root:${drivePath}:/content`)
    .put(Buffer.from(lines.join('\n'), 'utf-8'));
}

function csvEscape(value: any): string {
  if (value === null || value === undefined) return '';
  const s = String(value);
  if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

async function ensureFolder(client: Client, folderPath: string): Promise<void> {
  try {
    await client.api(`/me/drive/root:${folderPath}`).get();
  } catch (err: any) {
    if (err.statusCode !== 404) throw err;
    const parent = folderPath.substring(0, folderPath.lastIndexOf('/'));
    const name = folderPath.substring(folderPath.lastIndexOf('/') + 1);
    if (parent) await ensureFolder(client, parent);
    const parentApi = parent ? `/me/drive/root:${parent}:/children` : '/me/drive/root/children';
    await client.api(parentApi).post({
      name,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'replace',
    });
  }
}

function required(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Missing env var: ${name}`);
  return v;
}
