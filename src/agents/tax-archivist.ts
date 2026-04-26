/**
 * Tax Archivist agent.
 *
 * Classifies emails as tax-relevant business expenses, maps to Schedule C
 * categories, and attributes to entities. Used both for real-time tagging
 * (on every new email) and historical backfill.
 */

import Anthropic from '@anthropic-ai/sdk';
import { z } from 'zod';
import type { EmailMessage, Attachment } from '../graph/client';

const TaxClassification = z.object({
  is_tax_relevant: z.boolean(),
  vendor: z.string().nullable(),
  amount: z.number().nullable(),
  date: z.string().nullable(), // ISO date of transaction
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
      'Other',
      'Not_Applicable',
    ])
    .default('Not_Applicable'),
  entity: z.enum(['NHP-1941', 'NHDPM', 'Sol Haus', 'Personal', 'Unknown']),
  deductible: z.enum(['yes', 'no', 'partial']),
  business_purpose: z.string(),
  is_1099_candidate: z.boolean().default(false),
  confidence: z.number().min(0).max(1),
  flags: z.array(z.string()).default([]),
});

export type TaxClassification = z.infer<typeof TaxClassification>;

const SYSTEM_PROMPT = `You are Tax Archivist. You analyze emails to identify business expenses for Schedule C tax preparation.

The user has these business entities:
- NHP-1941: real estate holding (rental property expenses, property tax, insurance, repairs, mortgage interest)
- NHDPM: dental practice management activities
- Sol Haus: short-term rental in Panama City Beach (Schedule E or C depending on services - assume C with substantial services for now)
  - Sol Haus expenses: Guesty/Airbnb/Vrbo commissions, cleaning fees (contract labor), utilities at the STR address, supplies, repairs, advertising
- Personal: not deductible

SCHEDULE C CATEGORIES (mapped to IRS line items):
- Advertising: ads, marketing, listing fees that aren't platform commissions
- Car_Truck: vehicle expenses, mileage, fuel for business use
- Contract_Labor: cleaners, contractors, freelancers (also flag is_1099_candidate=true if individual/sole prop)
- Insurance: business insurance, liability, property insurance for business assets
- Legal_Professional: lawyers, accountants, consultants
- Office_Expense: general office costs
- Repairs_Maintenance: property repairs, equipment maintenance
- Supplies: consumables for the business
- Travel: business travel (lodging, transport - excluding meals)
- Meals: business meals (50% deductible by default)
- Utilities: electric, water, gas, internet for business property (Sol Haus utilities qualify)
- Software_Subscriptions: SaaS, software licenses (Claude API, hosting, productivity tools used for business)
- Education: courses, books, conferences for business skills
- Other: business expense that doesn't fit
- Not_Applicable: not a business expense (personal, or not an expense at all)

DEDUCTIBLE FIELD:
- "yes": clearly fully deductible business expense
- "no": clearly personal or not deductible
- "partial": mixed-use (personal + business), business meals (50%), or uncertain split

1099 candidates: payments to individuals or sole proprietors >= $600/year for services.
Platforms (Stripe, Airbnb, Guesty) are not 1099 candidates - they issue their own forms.

CONSERVATIVE BIAS: If unsure whether deductible, mark "partial" and explain. Never invent business purposes.

Respond with valid JSON only. No prose outside JSON.`;

export async function classifyForTax(
  email: EmailMessage,
  attachments: Attachment[],
): Promise<TaxClassification> {
  const client = new Anthropic();

  const content: Anthropic.MessageCreateParams['messages'][0]['content'] = [
    { type: 'text', text: buildContext(email, attachments) },
  ];

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
          media_type: att.contentType as 'image/jpeg' | 'image/png' | 'image/gif' | 'image/webp',
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

  return TaxClassification.parse(JSON.parse(text));
}

function buildContext(email: EmailMessage, attachments: Attachment[]): string {
  const attachmentSummary = attachments.length
    ? `\n\nAttachments:\n${attachments.map((a) => `- ${a.name} (${a.contentType})`).join('\n')}`
    : '';

  const body =
    email.body.contentType === 'html'
      ? stripHtml(email.body.content)
      : email.body.content;

  return `Analyze this email for tax/business expense relevance.

From: ${email.from.emailAddress.name} <${email.from.emailAddress.address}>
Subject: ${email.subject}
Received: ${email.receivedDateTime}${attachmentSummary}

Body:
${body.slice(0, 8000)}

Respond with JSON only:
{
  "is_tax_relevant": boolean,
  "vendor": string | null,
  "amount": number | null,
  "date": "YYYY-MM-DD" | null,
  "schedule_c_category": "Advertising" | "Car_Truck" | "Contract_Labor" | "Insurance" | "Legal_Professional" | "Office_Expense" | "Repairs_Maintenance" | "Supplies" | "Travel" | "Meals" | "Utilities" | "Software_Subscriptions" | "Education" | "Other" | "Not_Applicable",
  "entity": "NHP-1941" | "NHDPM" | "Sol Haus" | "Personal" | "Unknown",
  "deductible": "yes" | "no" | "partial",
  "business_purpose": "one sentence justification",
  "is_1099_candidate": boolean,
  "confidence": 0.0-1.0,
  "flags": []
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
