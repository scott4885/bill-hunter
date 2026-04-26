/**
 * Bill Hunter agent.
 *
 * Given an email + attachments, classifies and extracts invoice/bill data
 * using Claude. Returns structured output for the bills.xlsx schema.
 */

import Anthropic from '@anthropic-ai/sdk';
import { z } from 'zod';
import { createHash } from 'crypto';
import type { EmailMessage, Attachment } from '../graph/client';

const BillExtraction = z.object({
  is_bill: z.boolean(),
  vendor: z.string().nullable(),
  amount: z.number().nullable(),
  currency: z.string().default('USD'),
  due_date: z.string().nullable(), // ISO date
  invoice_number: z.string().nullable(),
  entity: z.enum(['NHP-1941', 'NHDPM', 'Sol Haus', 'Personal', 'Unknown']),
  confidence: z.number().min(0).max(1),
  reasoning: z.string(),
  flags: z.array(z.string()).default([]),
});

export type BillExtraction = z.infer<typeof BillExtraction>;

const SYSTEM_PROMPT = `You are Bill Hunter. You analyze emails to identify and extract bills/invoices that require payment.

A BILL/INVOICE is something that requires the recipient to pay money. Examples:
- Utility bills, credit card statements with balance due, vendor invoices
- Subscription renewals with charge upcoming
- Service invoices (cleaning, repairs, professional services)

NOT bills (do not flag is_bill=true):
- Receipts confirming payment already made
- Order confirmations for purchases
- Marketing emails, newsletters, promotions
- Bank statements showing transactions (no action needed)
- Welcome emails, account notifications

ENTITY ATTRIBUTION (the user's businesses):
- NHP-1941: real estate holding entity (property taxes, insurance, HOA, mortgage statements for investment property)
- NHDPM: dental practice activities (likely not in personal inbox; flag if seen)
- Sol Haus: short-term rental in Panama City Beach (utilities for that address, Guesty/Airbnb/Vrbo invoices, cleaning services, supplies for the STR)
- Personal: anything else the user pays personally (personal utilities, subscriptions, credit cards)
- Unknown: cannot determine; user should review

If confidence < 0.7, set the relevant fields to your best guess but flag for review in the "flags" array.

Always respond with valid JSON matching the schema. No prose outside JSON.`;

export async function extractBill(
  email: EmailMessage,
  attachments: Attachment[],
): Promise<BillExtraction> {
  const client = new Anthropic();

  const content: Anthropic.MessageCreateParams['messages'][0]['content'] = [
    {
      type: 'text',
      text: buildEmailContext(email, attachments),
    },
  ];

  // Attach PDFs/images directly so Claude can read them.
  for (const att of attachments) {
    if (!att.contentBytes) continue;
    if (att.contentType === 'application/pdf') {
      content.push({
        type: 'document',
        source: {
          type: 'base64',
          media_type: 'application/pdf',
          data: att.contentBytes,
        },
      });
    } else if (att.contentType.startsWith('image/')) {
      content.push({
        type: 'image',
        source: {
          type: 'base64',
          media_type: att.contentType as
            | 'image/jpeg'
            | 'image/png'
            | 'image/gif'
            | 'image/webp',
          data: att.contentBytes,
        },
      });
    }
  }

  const response = await client.messages.create({
    model: 'claude-opus-4-7',
    max_tokens: 1024,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content }],
  });

  const text = response.content
    .filter((b): b is Anthropic.TextBlock => b.type === 'text')
    .map((b) => b.text)
    .join('\n')
    .replace(/```json|```/g, '')
    .trim();

  const parsed = BillExtraction.parse(JSON.parse(text));
  return parsed;
}

function buildEmailContext(email: EmailMessage, attachments: Attachment[]): string {
  const attachmentSummary = attachments.length
    ? `\n\nAttachments:\n${attachments.map((a) => `- ${a.name} (${a.contentType}, ${a.size} bytes)`).join('\n')}`
    : '\n\n(no attachments)';

  // Strip HTML to text for the prompt; full content sent as document if PDF.
  const bodyText =
    email.body.contentType === 'html'
      ? stripHtml(email.body.content)
      : email.body.content;

  return `Analyze this email and extract bill/invoice information.

From: ${email.from.emailAddress.name} <${email.from.emailAddress.address}>
Subject: ${email.subject}
Received: ${email.receivedDateTime}
${attachmentSummary}

Body:
${bodyText.slice(0, 8000)}

Respond with JSON only matching this schema:
{
  "is_bill": boolean,
  "vendor": string | null,
  "amount": number | null,
  "currency": "USD",
  "due_date": "YYYY-MM-DD" | null,
  "invoice_number": string | null,
  "entity": "NHP-1941" | "NHDPM" | "Sol Haus" | "Personal" | "Unknown",
  "confidence": 0.0-1.0,
  "reasoning": "one sentence explaining your classification",
  "flags": ["needs_review", "ambiguous_amount", ...]
}`;
}

function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

export function buildInvoiceId(extraction: BillExtraction, email: EmailMessage): string {
  // Stable hash for dedupe: vendor + amount + invoice_number (or due_date as fallback)
  const key = [
    extraction.vendor ?? 'unknown',
    extraction.amount ?? 0,
    extraction.invoice_number ?? extraction.due_date ?? email.id,
  ].join('|');
  return createHash('sha256').update(key).digest('hex').slice(0, 16);
}
