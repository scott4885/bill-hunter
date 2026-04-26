/**
 * Microsoft Graph client wrapper.
 *
 * Handles email fetching, attachment download, and webhook subscription management.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import { getAccessToken } from '../auth/tokens';

export function getGraphClient(): Client {
  return Client.init({
    authProvider: async (done) => {
      try {
        const token = await getAccessToken();
        done(null, token);
      } catch (err) {
        done(err as Error, null);
      }
    },
  });
}

export interface EmailMessage {
  id: string;
  subject: string;
  from: { emailAddress: { name: string; address: string } };
  receivedDateTime: string;
  bodyPreview: string;
  body: { contentType: string; content: string };
  hasAttachments: boolean;
  webLink: string;
  conversationId: string;
}

export interface Attachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
  contentBytes?: string; // base64
}

export async function getMessage(messageId: string): Promise<EmailMessage> {
  const client = getGraphClient();
  return client
    .api(`/me/messages/${messageId}`)
    .select(
      'id,subject,from,receivedDateTime,bodyPreview,body,hasAttachments,webLink,conversationId',
    )
    .get();
}

export async function getAttachments(messageId: string): Promise<Attachment[]> {
  const client = getGraphClient();
  const res = await client.api(`/me/messages/${messageId}/attachments`).get();
  return res.value || [];
}

export async function listMessagesByDateRange(
  startIso: string,
  endIso: string,
  pageSize = 50,
): Promise<EmailMessage[]> {
  const client = getGraphClient();
  const messages: EmailMessage[] = [];
  let url:
    | string
    | null = `/me/messages?$filter=receivedDateTime ge ${startIso} and receivedDateTime le ${endIso}&$top=${pageSize}&$select=id,subject,from,receivedDateTime,bodyPreview,hasAttachments,webLink,conversationId`;

  while (url) {
    const res: { value: EmailMessage[]; '@odata.nextLink'?: string } = await client
      .api(url)
      .get();
    messages.push(...res.value);
    url = res['@odata.nextLink'] ?? null;
  }
  return messages;
}

export interface SubscriptionInput {
  changeType: 'created' | 'updated' | 'deleted' | string; // comma-sep allowed
  notificationUrl: string;
  resource: string; // e.g. "me/mailFolders('inbox')/messages"
  expirationDateTime: string; // ISO
  clientState: string;
}

export async function createSubscription(input: SubscriptionInput) {
  const client = getGraphClient();
  return client.api('/subscriptions').post(input);
}

export async function listSubscriptions() {
  const client = getGraphClient();
  const res = await client.api('/subscriptions').get();
  return res.value as Array<{ id: string; expirationDateTime: string; resource: string }>;
}

export async function renewSubscription(id: string, expirationDateTime: string) {
  const client = getGraphClient();
  return client.api(`/subscriptions/${id}`).patch({ expirationDateTime });
}

export async function deleteSubscription(id: string) {
  const client = getGraphClient();
  return client.api(`/subscriptions/${id}`).delete();
}
