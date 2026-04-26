/**
 * Statement Parser agent.
 *
 * Reads credit card or bank statement PDFs and extracts every transaction
 * line item. Output rows are written to transactions_YYYY.xlsx and later
 * reconciled with email receipts/bills.
 *
 * Trigger: drop a PDF into ONEDRIVE_STATEMENTS_INBOX. The watcher invokes
 * this agent, which writes rows + moves the PDF to a "processed" folder.
 */

import Anthropic from '@anthropic-ai/sdk';
import { z } from 'zod';
import { createHash } from 'crypto';

const StatementTransaction = z.object({
  date: z.string(), // YYYY-MM-DD
  description: z.string(),
  amount: z.number(), // negative = outflow / charge; positive = payment/credit/deposit
  schedule_c_category: z
    .enum([
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
      'Not_Applicable',
    ])
    .default('Other'),
  entity: z.enum(['NHP-1941', 'NHDPM', 'Sol Haus', 'Personal', 'Unknown']),
  deductible: z.enum(['yes', 'no', 'partial']),
  confidence: z.number().min(0).max(1),
  flags: z.array(z.string()).default([]),
});

const StatementResult = z.object({
  account_label: z.string(), // e.g. "Chase Sapphire 5678" or "Regions Checking 1234"
  account_type: z.enum(['credit_card', 'checking', 'savings', 'unknown']),
  statement_period_start: z.string(), // YYYY-MM-DD
  statement_period_end: z.string(),
  opening_balance: z.number().nullable(),
  closing_balance: z.number().nullable(),
  transactions: z.array(StatementTransaction),
});

export type StatementResult = z.infer<typeof StatementResult>;
export type StatementTransaction = z.infer<typeof StatementTransaction>;

const SYSTEM_PROMPT = `You are Statement Parser. You extract every transaction from a credit card or bank statement PDF.

ENTITIES (the user's businesses):
- NHP-1941: real estate holding (mortgage payments, property tax, property insurance, repairs on rental property)
- NHDPM: dental practice management activities
- Sol Haus: short-term rental at Panama City Beach (utilities for STR property, cleaning vendors, Guesty/Airbnb/Vrbo deposits, supplies for STR)
- Personal: personal spending (groceries, personal subscriptions, personal restaurants)
- Unknown: cannot determine from description alone

CATEGORIES (Schedule C / Schedule E lines):
- Advertising, Car_Truck, Contract_Labor, Insurance, Legal_Professional
- Office_Expense, Repairs_Maintenance, Supplies, Travel, Meals
- Utilities (FPL, water, gas, internet at business properties)
- Software_Subscriptions (Claude API, Anthropic, AWS, Azure, hosting, productivity SaaS used for business)
- Education, Mortgage_Interest, Property_Tax, Bank_Fees, Other, Not_Applicable

INSTRUCTIONS:
1. Extract every transaction. Use the post date if both trans and post are shown.
2. Sign convention: charges/withdrawals = NEGATIVE, payments/deposits/credits = POSITIVE.
3. For each transaction, classify entity + category + deductible.
4. Be conservative: if you can't confidently attribute to a business, mark entity=Personal and deductible=no, OR entity=Unknown if it's clearly business but you can't tell which one.
5. Bank fees, finance charges, foreign transaction fees → Bank_Fees, deductible=yes if on a business account, otherwise no.
6. Account-to-account transfers, payments to credit card from bank → Not_Applicable, deductible=no (these are not expenses).
7. Set confidence < 0.7 if uncertain on category or entity.

Respond with valid JSON only. No prose outside JSON.`;

export async function parseStatement(
  pdfBytesBase64: string,
  filename: string,
): Promise<StatementResult> {
  const client = new Anthropic();

  const response = await client.messages.create({
    model: 'claude-opus-4-7',
    max_tokens: 8192,
    system: SYSTEM_PROMPT,
    messages: [
      {
        role: 'user',
        content: [
          {
            type: 'document',
            source: {
              type: 'base64',
              media_type: 'application/pdf',
              data: pdfBytesBase64,
            },
          },
          {
            type: 'text',
            text: `Filename: ${filename}

Extract all transactions and return JSON matching this schema:
{
  "account_label": "string (e.g. 'Chase Sapphire 5678')",
  "account_type": "credit_card" | "checking" | "savings" | "unknown",
  "statement_period_start": "YYYY-MM-DD",
  "statement_period_end": "YYYY-MM-DD",
  "opening_balance": number | null,
  "closing_balance": number | null,
  "transactions": [
    {
      "date": "YYYY-MM-DD",
      "description": "merchant or memo as printed",
      "amount": number (negative = charge/withdrawal, positive = payment/deposit),
      "schedule_c_category": "...",
      "entity": "...",
      "deductible": "yes" | "no" | "partial",
      "confidence": 0.0-1.0,
      "flags": []
    }
  ]
}

Return JSON only, no prose, no markdown fences.`,
          },
        ],
      },
    ],
  });

  const text = response.content
    .filter((b): b is Anthropic.TextBlock => b.type === 'text')
    .map((b) => b.text)
    .join('\n')
    .replace(/```json|```/g, '')
    .trim();

  return StatementResult.parse(JSON.parse(text));
}

/**
 * Stable hash for transaction dedupe.
 * Same date + amount + description + account = same transaction.
 */
export function buildTxnId(
  accountLabel: string,
  txn: StatementTransaction,
): string {
  const key = [
    accountLabel,
    txn.date,
    txn.amount.toFixed(2),
    txn.description.toLowerCase().replace(/\s+/g, ' ').trim(),
  ].join('|');
  return createHash('sha256').update(key).digest('hex').slice(0, 16);
}
