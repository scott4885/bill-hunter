/**
 * Run Bill Hunter + Tax Archivist on a single Graph message id.
 *
 * Extracted from the deprecated webhook handler so both push (legacy) and
 * poll (Paperclip-driven) entrypoints share one implementation.
 */

import type { InvocationContext } from '@azure/functions';
import { getMessage, getAttachments } from '../graph/client';
import { extractBill, buildInvoiceId } from '../agents/bill-hunter';
import { classifyForTax } from '../agents/tax-archivist';
import {
  appendBillRow,
  appendTaxRow,
  billExists,
  uploadAttachment,
} from '../graph/excel';

export interface ProcessResult {
  messageId: string;
  receivedDateTime: string;
  billLogged: boolean;
  taxLogged: boolean;
  duplicate: boolean;
}

export async function processMessage(
  messageId: string,
  ctx: InvocationContext,
): Promise<ProcessResult> {
  const email = await getMessage(messageId);
  const rawAttachments = email.hasAttachments ? await getAttachments(messageId) : [];

  const [billResult, taxResult] = await Promise.allSettled([
    extractBill(email, rawAttachments),
    classifyForTax(email, rawAttachments),
  ]);

  if (billResult.status === 'rejected') ctx.error('extractBill failed', billResult.reason);
  if (taxResult.status === 'rejected') ctx.error('classifyForTax failed', taxResult.reason);

  if (billResult.status === 'fulfilled') {
    ctx.log(
      `bill: is_bill=${billResult.value.is_bill} vendor=${billResult.value.vendor} amount=${billResult.value.amount} confidence=${billResult.value.confidence}`,
    );
  }
  if (taxResult.status === 'fulfilled') {
    ctx.log(
      `tax: is_tax_relevant=${taxResult.value.is_tax_relevant} vendor=${taxResult.value.vendor} amount=${taxResult.value.amount}`,
    );
  }

  let billLogged = false;
  let duplicate = false;

  if (billResult.status === 'fulfilled' && billResult.value.is_bill) {
    const ext = billResult.value;
    const invoiceId = buildInvoiceId(ext, email);

    if (await billExists(invoiceId)) {
      ctx.log(`Skipping duplicate bill ${invoiceId}`);
      duplicate = true;
    } else {
      const attachmentPaths: string[] = [];
      for (const att of rawAttachments) {
        if (att.contentBytes && att.contentType === 'application/pdf') {
          const p = await uploadAttachment(att.name, att.contentBytes);
          attachmentPaths.push(p);
        }
      }

      await appendBillRow({
        invoice_id: invoiceId,
        received_date: email.receivedDateTime.slice(0, 10),
        vendor: ext.vendor ?? '',
        amount: ext.amount ?? '',
        due_date: ext.due_date ?? '',
        invoice_number: ext.invoice_number ?? '',
        entity: ext.entity,
        status: 'unpaid',
        email_link: email.webLink,
        attachment_path: attachmentPaths.join('; '),
        confidence: ext.confidence,
        notes: `${ext.reasoning}${ext.flags.length ? ` | flags: ${ext.flags.join(',')}` : ''}`,
      });
      ctx.log(`Logged bill ${invoiceId} from ${ext.vendor} for $${ext.amount}`);
      billLogged = true;
    }
  }

  let taxLogged = false;
  if (taxResult.status === 'fulfilled' && taxResult.value.is_tax_relevant) {
    const tax = taxResult.value;
    const txDate = tax.date ?? email.receivedDateTime.slice(0, 10);
    const year = parseInt(txDate.slice(0, 4), 10);

    let receiptPath = '';
    for (const att of rawAttachments) {
      if (att.contentBytes && att.contentType === 'application/pdf') {
        receiptPath = await uploadAttachment(att.name, att.contentBytes);
        break;
      }
    }

    await appendTaxRow(year, {
      date: txDate,
      vendor: tax.vendor ?? '',
      amount: tax.amount ?? '',
      schedule_c_category: tax.schedule_c_category,
      entity: tax.entity,
      deductible: tax.deductible,
      business_purpose: tax.business_purpose,
      receipt_path: receiptPath,
      email_link: email.webLink,
      is_1099_candidate: tax.is_1099_candidate ? 'yes' : 'no',
      confidence: tax.confidence,
      notes: tax.flags.join(','),
    });
    ctx.log(`Tax-logged ${tax.vendor} (${tax.schedule_c_category}, ${tax.entity})`);
    taxLogged = true;
  }

  return {
    messageId,
    receivedDateTime: email.receivedDateTime,
    billLogged,
    taxLogged,
    duplicate,
  };
}
