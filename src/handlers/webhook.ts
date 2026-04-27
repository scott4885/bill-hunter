/**
 * Azure Function HTTP trigger for Microsoft Graph webhook notifications.
 *
 * Graph sends a POST with notifications when emails arrive in the inbox.
 * The handshake (validationToken) must be acknowledged within 10 seconds.
 *
 * Subscription resource: me/mailFolders('inbox')/messages
 * Change types: created
 */

import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getMessage, getAttachments } from '../graph/client';
import { extractBill, buildInvoiceId } from '../agents/bill-hunter';
import { classifyForTax } from '../agents/tax-archivist';
import { appendBillRow, appendTaxRow, billExists, uploadAttachment } from '../graph/excel';

interface GraphNotification {
  subscriptionId: string;
  clientState?: string;
  changeType: string;
  resource: string; // e.g. "Users/{id}/Messages/{messageId}"
  resourceData: { id: string; '@odata.type': string };
  tenantId: string;
}

export async function webhookHandler(
  req: HttpRequest,
  ctx: InvocationContext,
): Promise<HttpResponseInit> {
  // Handshake: Graph appends ?validationToken=... on subscription creation
  const url = new URL(req.url);
  const validationToken = url.searchParams.get('validationToken');
  if (validationToken) {
    return {
      status: 200,
      headers: { 'Content-Type': 'text/plain' },
      body: validationToken,
    };
  }

  // Real notification
  const body = (await req.json()) as { value: GraphNotification[] };
  const expectedClientState = process.env.WEBHOOK_CLIENT_STATE_SECRET;

  // ACK quickly; do work async via setImmediate so we return 202 within 10s.
  setImmediate(async () => {
    for (const notification of body.value || []) {
      try {
        if (notification.clientState !== expectedClientState) {
          ctx.warn(`Bad clientState on notification ${notification.subscriptionId}`);
          continue;
        }
        await processNotification(notification, ctx);
      } catch (err) {
        ctx.error('Notification processing failed', err);
      }
    }
  });

  return { status: 202 };
}

async function processNotification(n: GraphNotification, ctx: InvocationContext) {
  if (n.changeType !== 'created') return;
  const messageId = n.resourceData.id;
  ctx.log(`Processing message ${messageId}`);

  const email = await getMessage(messageId);
  const rawAttachments = email.hasAttachments ? await getAttachments(messageId) : [];

  // Run both agents in parallel.
  const [billResult, taxResult] = await Promise.allSettled([
    extractBill(email, rawAttachments),
    classifyForTax(email, rawAttachments),
  ]);

  if (billResult.status === 'rejected') {
    ctx.error('extractBill failed', billResult.reason);
  }
  if (taxResult.status === 'rejected') {
    ctx.error('classifyForTax failed', taxResult.reason);
  }
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

  // Bill Hunter
  if (billResult.status === 'fulfilled' && billResult.value.is_bill) {
    const ext = billResult.value;
    const invoiceId = buildInvoiceId(ext, email);

    if (await billExists(invoiceId)) {
      ctx.log(`Skipping duplicate bill ${invoiceId}`);
    } else {
      // Save attachments to OneDrive for the record
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
    }
  }

  // Tax Archivist
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
  }
}

if (process.env.FUNCTIONS_WORKER_RUNTIME) {
  app.http('webhook', {
    methods: ['POST', 'GET'],
    authLevel: 'anonymous', // Graph posts here; security via clientState
    handler: webhookHandler,
  });
}
