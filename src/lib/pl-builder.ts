/**
 * P&L builder.
 *
 * Generates a per-entity profit & loss spreadsheet for a given year.
 * Reads transactions + rentals, groups by category × month, writes a
 * formatted P&L workbook to OneDrive.
 *
 * Output: /Documents/BillHunter/pl/pl_{entity}_{year}.xlsx
 *
 * NOTE: Mortgage interest, depreciation, and home office allocations
 * are typically calculated by the CPA. This produces the income/expense
 * picture needed as INPUT to those calculations.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { getGraphClient } from '../graph/client';
import { listTransactions, listRentals } from '../graph/excel';
import { EMPTY_XLSX_BASE64 } from './empty-xlsx';

const MONTHS = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

const EXPENSE_CATEGORIES = [
  'Advertising',
  'Car_Truck',
  'Contract_Labor',
  'Insurance',
  'Legal_Professional',
  'Office_Expense',
  'Repairs_Maintenance',
  'Supplies',
  'Travel',
  'Meals',
  'Utilities',
  'Software_Subscriptions',
  'Education',
  'Mortgage_Interest',
  'Property_Tax',
  'Bank_Fees',
  'Other',
];

interface PLOptions {
  entity: 'NHP-1941' | 'NHDPM' | 'Sol Haus' | 'Personal';
  year: number;
}

export async function buildPL(opts: PLOptions): Promise<string> {
  const { entity, year } = opts;
  const client = getGraphClient();
  const drivePath = `${required('ONEDRIVE_PL_DIR')}/pl_${slug(entity)}_${year}.xlsx`;

  // Bootstrap empty file at drivePath (overwrite each run)
  const emptyBytes = Buffer.from(EMPTY_XLSX_BASE64, 'base64');
  await client.api(`/me/drive/root:${drivePath}:/content`).put(emptyBytes);

  // Fetch source data
  const [txns, rentals] = await Promise.all([
    listTransactions(year),
    entity === 'Sol Haus' ? listRentals(year) : Promise.resolve([]),
  ]);

  const entityTxns = txns.filter((t) => String(t.entity) === entity);

  // Build expense pivot: category × month
  const pivot: Record<string, number[]> = {};
  for (const cat of EXPENSE_CATEGORIES) pivot[cat] = new Array(12).fill(0);

  for (const t of entityTxns) {
    const amount = Number(t.amount) || 0;
    if (amount >= 0) continue; // outflows only
    if (String(t.deductible) === 'no') continue;

    const cat = String(t.schedule_c_category || 'Other');
    const month = parseInt(String(t.date).slice(5, 7), 10) - 1;
    if (month < 0 || month > 11) continue;

    const expense = Math.abs(amount);
    const adjusted = String(t.deductible) === 'partial' ? expense * 0.5 : expense;
    if (!pivot[cat]) pivot[cat] = new Array(12).fill(0);
    pivot[cat][month] += adjusted;
  }

  // Revenue rows (Sol Haus only — pulls from rentals)
  const revenueRows: Array<{ label: string; monthly: number[] }> = [];
  if (entity === 'Sol Haus' && rentals.length > 0) {
    const platforms = ['Airbnb', 'Vrbo', 'Direct', 'Other'];
    for (const platform of platforms) {
      const monthly = new Array(12).fill(0);
      for (const r of rentals) {
        if (String(r.platform).toLowerCase() !== platform.toLowerCase()) continue;
        const month = parseInt(String(r.payout_date).slice(5, 7), 10) - 1;
        if (month < 0 || month > 11) continue;
        monthly[month] += Number(r.gross_revenue) || 0;
      }
      if (monthly.some((m) => m > 0)) {
        revenueRows.push({ label: `${platform} bookings`, monthly });
      }
    }
  }

  // Write to Sheet1
  const rows: any[][] = [];
  rows.push([`${entity} — ${year} Profit & Loss`, ...new Array(12).fill(''), '']);
  rows.push([]);
  rows.push(['', ...MONTHS, 'YTD']);

  // Revenue section
  if (revenueRows.length > 0) {
    rows.push(['REVENUE']);
    let revTotals = new Array(12).fill(0);
    for (const r of revenueRows) {
      rows.push([`  ${r.label}`, ...r.monthly, sum(r.monthly)]);
      revTotals = revTotals.map((v, i) => v + r.monthly[i]);
    }
    rows.push(['  Total Revenue', ...revTotals, sum(revTotals)]);
    rows.push([]);
  }

  // Expense section
  rows.push(['EXPENSES']);
  const expenseTotals = new Array(12).fill(0);
  for (const cat of EXPENSE_CATEGORIES) {
    const monthly = pivot[cat] || new Array(12).fill(0);
    if (monthly.every((v) => v === 0)) continue;
    rows.push([`  ${cat.replace(/_/g, ' ')}`, ...monthly, sum(monthly)]);
    monthly.forEach((v, i) => (expenseTotals[i] += v));
  }
  rows.push(['  Total Expenses', ...expenseTotals, sum(expenseTotals)]);
  rows.push([]);

  // Net
  if (revenueRows.length > 0) {
    const revTotals = revenueRows.reduce(
      (acc, r) => acc.map((v, i) => v + r.monthly[i]),
      new Array(12).fill(0),
    );
    const net = revTotals.map((r, i) => r - expenseTotals[i]);
    rows.push(['NET INCOME (pre-depreciation)', ...net, sum(net)]);
  } else {
    rows.push([
      'TOTAL DEDUCTIBLE EXPENSES',
      ...expenseTotals,
      sum(expenseTotals),
    ]);
  }

  rows.push([]);
  rows.push(['Notes:']);
  rows.push(['  - Mortgage interest: from 1098, CPA enters separately']);
  rows.push(['  - Depreciation: CPA calculates']);
  rows.push(['  - Meals shown at 50% (Schedule C standard)']);
  rows.push([
    `  - Generated ${new Date().toISOString().slice(0, 10)} from transactions_${year}.xlsx`,
  ]);

  // Pad rows to 14 columns (label + 12 months + YTD)
  const padded = rows.map((r) => {
    while (r.length < 14) r.push('');
    return r.slice(0, 14);
  });

  const lastCol = colIndexToLetter(14);
  const range = `A1:${lastCol}${padded.length}`;

  await client
    .api(
      `/me/drive/root:${drivePath}:/workbook/worksheets/Sheet1/range(address='${range}')`,
    )
    .patch({ values: padded });

  // Format header row bold
  await client
    .api(
      `/me/drive/root:${drivePath}:/workbook/worksheets/Sheet1/range(address='A1:${lastCol}1')/format/font`,
    )
    .patch({ bold: true, size: 14 });

  return drivePath;
}

function sum(arr: number[]): number {
  return Math.round(arr.reduce((a, b) => a + b, 0) * 100) / 100;
}

function slug(s: string): string {
  return s.replace(/[^A-Za-z0-9]/g, '');
}

function colIndexToLetter(n: number): string {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function required(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Missing env var: ${name}`);
  return v;
}
