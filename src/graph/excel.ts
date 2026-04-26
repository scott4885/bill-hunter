/**
 * OneDrive Excel writer via Graph workbook API.
 *
 * Auto-creates the file + table on first write. Supports four table types:
 *   - Bills (bills.xlsx)
 *   - TaxLog (taxes_YYYY.xlsx)
 *   - Transactions (transactions_YYYY.xlsx)
 *   - Rentals (rentals_YYYY.xlsx)
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { getGraphClient } from './client';
import { EMPTY_XLSX_BASE64 } from '../lib/empty-xlsx';

interface TableSpec {
  name: string;
  headers: string[];
}

const BILLS_TABLE: TableSpec = {
  name: 'Bills',
  headers: [
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
  ],
};

const TAX_TABLE: TableSpec = {
  name: 'TaxLog',
  headers: [
    'date',
    'vendor',
    'amount',
    'schedule_c_category',
    'entity',
    'deductible',
    'business_purpose',
    'receipt_path',
    'email_link',
    'is_1099_candidate',
    'confidence',
    'notes',
  ],
};

const TRANSACTIONS_TABLE: TableSpec = {
  name: 'Transactions',
  headers: [
    'txn_id',
    'date',
    'description',
    'amount',
    'account',
    'statement_path',
    'entity',
    'schedule_c_category',
    'deductible',
    'matched_receipt_path',
    'matched_bill_id',
    'reconciliation_status',
    'confidence',
    'notes',
  ],
};

const RENTALS_TABLE: TableSpec = {
  name: 'Rentals',
  headers: [
    'booking_id',
    'platform',
    'property',
    'guest_name',
    'check_in',
    'check_out',
    'nights',
    'gross_revenue',
    'platform_fees',
    'cleaning_fees_collected',
    'taxes_collected',
    'net_payout',
    'payout_date',
    'entity',
    'notes',
  ],
};

async function ensureFileWithTable(
  client: Client,
  drivePath: string,
  table: TableSpec,
): Promise<void> {
  // Create the file if it doesn't exist.
  try {
    await client.api(`/me/drive/root:${drivePath}`).get();
  } catch (err: any) {
    if (err.statusCode !== 404) throw err;
    const emptyBytes = Buffer.from(EMPTY_XLSX_BASE64, 'base64');
    await client.api(`/me/drive/root:${drivePath}:/content`).put(emptyBytes);
  }

  // Create the named table if it doesn't exist.
  try {
    await client
      .api(`/me/drive/root:${drivePath}:/workbook/tables/${table.name}`)
      .get();
    return;
  } catch (err: any) {
    if (err.statusCode !== 404) throw err;
  }

  const colLetter = colIndexToLetter(table.headers.length);
  const range = `A1:${colLetter}1`;

  await client
    .api(
      `/me/drive/root:${drivePath}:/workbook/worksheets/Sheet1/range(address='${range}')`,
    )
    .patch({ values: [table.headers] });

  await client
    .api(`/me/drive/root:${drivePath}:/workbook/worksheets/Sheet1/tables/add`)
    .post({ address: `Sheet1!${range}`, hasHeaders: true });

  const tables = await client
    .api(`/me/drive/root:${drivePath}:/workbook/tables`)
    .get();
  const newest = tables.value[tables.value.length - 1];
  if (newest.name !== table.name) {
    await client
      .api(`/me/drive/root:${drivePath}:/workbook/tables/${newest.name}`)
      .patch({ name: table.name });
  }
}

async function appendRow(
  drivePath: string,
  table: TableSpec,
  row: Record<string, unknown>,
): Promise<void> {
  const client = getGraphClient();
  await ensureFileWithTable(client, drivePath, table);
  const values = [table.headers.map((h) => row[h] ?? '')];
  await client
    .api(`/me/drive/root:${drivePath}:/workbook/tables/${table.name}/rows/add`)
    .post({ values });
}

async function listTableRows(
  drivePath: string,
  table: TableSpec,
): Promise<Array<Record<string, any>>> {
  const client = getGraphClient();
  try {
    const res = await client
      .api(`/me/drive/root:${drivePath}:/workbook/tables/${table.name}/rows`)
      .get();
    return (res.value || []).map((r: any) =>
      Object.fromEntries(table.headers.map((h, i) => [h, r.values[0][i]])),
    );
  } catch (err: any) {
    if (err.statusCode === 404) return [];
    throw err;
  }
}

// ---------- Bills ----------

export async function appendBillRow(row: Record<string, unknown>): Promise<void> {
  return appendRow(required('ONEDRIVE_BILLS_PATH'), BILLS_TABLE, row);
}

export async function billExists(invoiceId: string): Promise<boolean> {
  const path = required('ONEDRIVE_BILLS_PATH');
  const client = getGraphClient();
  try {
    const res = await client
      .api(
        `/me/drive/root:${path}:/workbook/tables/${BILLS_TABLE.name}/columns/invoice_id/values`,
      )
      .get();
    const flat: string[] = (res.values || []).flat().map(String);
    return flat.includes(invoiceId);
  } catch {
    return false;
  }
}

export async function listBills() {
  return listTableRows(required('ONEDRIVE_BILLS_PATH'), BILLS_TABLE);
}

// ---------- Tax ----------

export async function appendTaxRow(
  year: number,
  row: Record<string, unknown>,
): Promise<void> {
  const path = `${required('ONEDRIVE_TAX_PATH_PREFIX')}${year}.xlsx`;
  return appendRow(path, TAX_TABLE, row);
}

export async function listTaxRows(year: number) {
  const path = `${required('ONEDRIVE_TAX_PATH_PREFIX')}${year}.xlsx`;
  return listTableRows(path, TAX_TABLE);
}

// ---------- Transactions ----------

export async function appendTransactionRow(
  year: number,
  row: Record<string, unknown>,
): Promise<void> {
  const path = `${required('ONEDRIVE_TRANSACTIONS_PATH_PREFIX')}${year}.xlsx`;
  return appendRow(path, TRANSACTIONS_TABLE, row);
}

export async function listTransactions(year: number) {
  const path = `${required('ONEDRIVE_TRANSACTIONS_PATH_PREFIX')}${year}.xlsx`;
  return listTableRows(path, TRANSACTIONS_TABLE);
}

export async function transactionExists(
  year: number,
  txnId: string,
): Promise<boolean> {
  const path = `${required('ONEDRIVE_TRANSACTIONS_PATH_PREFIX')}${year}.xlsx`;
  const client = getGraphClient();
  try {
    const res = await client
      .api(
        `/me/drive/root:${path}:/workbook/tables/${TRANSACTIONS_TABLE.name}/columns/txn_id/values`,
      )
      .get();
    const flat: string[] = (res.values || []).flat().map(String);
    return flat.includes(txnId);
  } catch {
    return false;
  }
}

export async function updateTransactionMatch(
  year: number,
  txnId: string,
  patch: Partial<Record<string, unknown>>,
): Promise<void> {
  const path = `${required('ONEDRIVE_TRANSACTIONS_PATH_PREFIX')}${year}.xlsx`;
  const client = getGraphClient();
  const res = await client
    .api(`/me/drive/root:${path}:/workbook/tables/${TRANSACTIONS_TABLE.name}/rows`)
    .get();
  const rows: any[] = res.value || [];
  const headers = TRANSACTIONS_TABLE.headers;
  const idIdx = headers.indexOf('txn_id');
  const targetIdx = rows.findIndex((r) => r.values[0][idIdx] === txnId);
  if (targetIdx < 0) return;

  const current = rows[targetIdx].values[0];
  const updated = [...current];
  for (const [k, v] of Object.entries(patch)) {
    const i = headers.indexOf(k);
    if (i >= 0) updated[i] = v ?? '';
  }
  await client
    .api(
      `/me/drive/root:${path}:/workbook/tables/${TRANSACTIONS_TABLE.name}/rows/itemAt(index=${targetIdx})`,
    )
    .patch({ values: [updated] });
}

// ---------- Rentals ----------

export async function appendRentalRow(
  year: number,
  row: Record<string, unknown>,
): Promise<void> {
  const path = `${required('ONEDRIVE_RENTALS_PATH_PREFIX')}${year}.xlsx`;
  return appendRow(path, RENTALS_TABLE, row);
}

export async function listRentals(year: number) {
  const path = `${required('ONEDRIVE_RENTALS_PATH_PREFIX')}${year}.xlsx`;
  return listTableRows(path, RENTALS_TABLE);
}

export async function rentalExists(
  year: number,
  bookingId: string,
): Promise<boolean> {
  const rows = await listRentals(year);
  return rows.some((r) => String(r.booking_id) === bookingId);
}

// ---------- Attachments / files ----------

export async function uploadAttachment(
  filename: string,
  contentBytesBase64: string,
  subdir = '',
): Promise<string> {
  const dir = required('ONEDRIVE_ATTACHMENTS_DIR');
  const safeName = filename.replace(/[^\w.\-]/g, '_');
  const stamped = `${Date.now()}_${safeName}`;
  const drivePath = subdir ? `${dir}/${subdir}/${stamped}` : `${dir}/${stamped}`;
  const client = getGraphClient();
  await client
    .api(`/me/drive/root:${drivePath}:/content`)
    .put(Buffer.from(contentBytesBase64, 'base64'));
  return drivePath;
}

export async function downloadFile(drivePath: string): Promise<Buffer> {
  const client = getGraphClient();
  const stream: ArrayBuffer = await client
    .api(`/me/drive/root:${drivePath}:/content`)
    .responseType('arraybuffer' as any)
    .get();
  return Buffer.from(stream);
}

export async function listFolder(
  folderPath: string,
): Promise<Array<{ name: string; id: string; size: number }>> {
  const client = getGraphClient();
  try {
    const res = await client
      .api(`/me/drive/root:${folderPath}:/children`)
      .get();
    return (res.value || []).map((f: any) => ({
      name: f.name,
      id: f.id,
      size: f.size,
    }));
  } catch (err: any) {
    if (err.statusCode === 404) return [];
    throw err;
  }
}

export async function moveFile(
  fromPath: string,
  toFolderPath: string,
): Promise<void> {
  const client = getGraphClient();
  // Get destination folder id
  const folder = await client.api(`/me/drive/root:${toFolderPath}`).get();
  await client.api(`/me/drive/root:${fromPath}`).patch({
    parentReference: { id: folder.id },
  });
}

// ---------- Helpers ----------

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

export { BILLS_TABLE, TAX_TABLE, TRANSACTIONS_TABLE, RENTALS_TABLE };
