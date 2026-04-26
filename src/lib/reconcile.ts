/**
 * Reconciliation engine.
 *
 * Joins:
 *   - transactions (from statement parser)
 *   - tax log rows / bills (from email pipeline)
 *
 * Match logic:
 *   - Same absolute amount within $0.01
 *   - Date within ±5 days
 *   - Vendor fuzzy match (token overlap >= 0.5 OR substring)
 *
 * Outcomes written to transactions.reconciliation_status:
 *   - "matched"       transaction has a receipt/bill in email pipeline
 *   - "unmatched"     no receipt found (might need to chase)
 *   - "transfer"      account-to-account / CC payment, not a real expense
 *   - "needs_review"  ambiguous (multiple candidates)
 */

import { differenceInDays, parseISO } from 'date-fns';
import {
  listTransactions,
  listTaxRows,
  listBills,
  updateTransactionMatch,
} from '../graph/excel';

interface Candidate {
  source: 'tax' | 'bill';
  date: string;
  vendor: string;
  amount: number;
  receipt_path: string;
  email_link: string;
  bill_id?: string;
}

export interface ReconcileResult {
  total: number;
  matched: number;
  unmatched: number;
  transfer: number;
  needsReview: number;
  alreadyMatched: number;
}

export async function reconcileYear(year: number): Promise<ReconcileResult> {
  const [txns, taxRows, bills] = await Promise.all([
    listTransactions(year),
    listTaxRows(year),
    listBills(),
  ]);

  // Build candidate pool from tax rows + bills.
  const candidates: Candidate[] = [
    ...taxRows.map<Candidate>((r) => ({
      source: 'tax',
      date: String(r.date),
      vendor: String(r.vendor || ''),
      amount: Number(r.amount) || 0,
      receipt_path: String(r.receipt_path || ''),
      email_link: String(r.email_link || ''),
    })),
    ...bills.map<Candidate>((b) => ({
      source: 'bill',
      date: String(b.due_date || b.received_date),
      vendor: String(b.vendor || ''),
      amount: Number(b.amount) || 0,
      receipt_path: String(b.attachment_path || ''),
      email_link: String(b.email_link || ''),
      bill_id: String(b.invoice_id || ''),
    })),
  ];

  const result: ReconcileResult = {
    total: txns.length,
    matched: 0,
    unmatched: 0,
    transfer: 0,
    needsReview: 0,
    alreadyMatched: 0,
  };

  for (const txn of txns) {
    const txnId = String(txn.txn_id);
    const currentStatus = String(txn.reconciliation_status || '');

    // Skip already-resolved ones
    if (currentStatus === 'matched' || currentStatus === 'transfer') {
      result.alreadyMatched++;
      continue;
    }

    const amount = Number(txn.amount) || 0;
    const desc = String(txn.description || '').toLowerCase();
    const date = String(txn.date);

    // Heuristic: classify obvious transfers/payments before searching candidates
    if (isTransferLike(desc, amount)) {
      await updateTransactionMatch(year, txnId, {
        reconciliation_status: 'transfer',
      });
      result.transfer++;
      continue;
    }

    // Only match outflows (charges/withdrawals) — receipts pertain to expenses.
    if (amount >= 0) {
      await updateTransactionMatch(year, txnId, {
        reconciliation_status: 'unmatched',
      });
      result.unmatched++;
      continue;
    }

    const txnAbs = Math.abs(amount);
    const matches = candidates.filter((c) => {
      if (Math.abs(c.amount - txnAbs) > 0.01) return false;
      if (!c.date) return false;
      try {
        const d = Math.abs(differenceInDays(parseISO(date), parseISO(c.date)));
        if (d > 5) return false;
      } catch {
        return false;
      }
      return vendorMatch(desc, c.vendor.toLowerCase());
    });

    if (matches.length === 1) {
      const m = matches[0];
      await updateTransactionMatch(year, txnId, {
        reconciliation_status: 'matched',
        matched_receipt_path: m.receipt_path,
        matched_bill_id: m.bill_id || '',
      });
      result.matched++;
    } else if (matches.length > 1) {
      await updateTransactionMatch(year, txnId, {
        reconciliation_status: 'needs_review',
        notes: `${matches.length} candidate matches`,
      });
      result.needsReview++;
    } else {
      await updateTransactionMatch(year, txnId, {
        reconciliation_status: 'unmatched',
      });
      result.unmatched++;
    }
  }

  return result;
}

function isTransferLike(desc: string, amount: number): boolean {
  const transferKeywords = [
    'payment thank you',
    'autopay payment',
    'mobile payment',
    'online payment',
    'transfer to',
    'transfer from',
    'zelle to self',
    'internal transfer',
    'ach credit',
    'credit card payment',
  ];
  const d = desc.toLowerCase();
  return transferKeywords.some((kw) => d.includes(kw));
}

function vendorMatch(txnDesc: string, vendor: string): boolean {
  if (!vendor) return false;
  if (txnDesc.includes(vendor)) return true;
  if (vendor.includes(txnDesc.split(/\s+/)[0])) return true;

  // Token overlap
  const txnTokens = new Set(
    txnDesc.split(/[\s\-_*#]+/).filter((t) => t.length >= 3),
  );
  const vendorTokens = new Set(
    vendor.split(/[\s\-_*#]+/).filter((t) => t.length >= 3),
  );
  if (vendorTokens.size === 0) return false;
  let overlap = 0;
  for (const t of vendorTokens) if (txnTokens.has(t)) overlap++;
  return overlap / vendorTokens.size >= 0.5;
}
