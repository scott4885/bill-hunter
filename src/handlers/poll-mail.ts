/**
 * Poll Microsoft Graph for new inbox messages and run them through
 * Bill Hunter + Tax Archivist.
 *
 * Replaces the push-driven webhook flow. Driven by Paperclip's http adapter
 * on a 60s heartbeat.
 *
 * Cursor state is persisted to BH_STATE_FILE (default /state/cursor.json):
 *   {
 *     "lastSeenIso": "2026-04-29T20:30:00Z",  // greatest receivedDateTime seen
 *     "recentIds":   ["AAMk...", ...]         // last 200 message ids, defensive
 *   }
 *
 * Crash safety: state is written after every processed message. If the
 * sidecar dies mid-batch, on restart we re-list since lastSeenIso (which
 * may be slightly stale) and skip anything already in recentIds.
 */

import type { InvocationContext } from '@azure/functions';
import { promises as fs } from 'node:fs';
import { dirname } from 'node:path';
import { getGraphClient, type EmailMessage } from '../graph/client';
import { processMessage } from '../lib/process-message';

const STATE_FILE = process.env.BH_STATE_FILE || '/state/cursor.json';
const RECENT_ID_CAP = 200;
const PAGE_SIZE = 25;

interface CursorState {
  lastSeenIso: string;
  recentIds: string[];
}

async function loadState(): Promise<CursorState> {
  try {
    const raw = await fs.readFile(STATE_FILE, 'utf8');
    const parsed = JSON.parse(raw) as Partial<CursorState>;
    return {
      lastSeenIso: parsed.lastSeenIso ?? bootstrapTs(),
      recentIds: Array.isArray(parsed.recentIds) ? parsed.recentIds : [],
    };
  } catch (err: any) {
    if (err?.code !== 'ENOENT') throw err;
    return { lastSeenIso: bootstrapTs(), recentIds: [] };
  }
}

async function saveState(state: CursorState): Promise<void> {
  await fs.mkdir(dirname(STATE_FILE), { recursive: true });
  await fs.writeFile(STATE_FILE, JSON.stringify(state, null, 2), 'utf8');
}

function bootstrapTs(): string {
  // First-run cursor = "now". Anything received before sidecar startup is
  // assumed to have been handled by the prior pipeline (webhook or backfill).
  // Avoids tax-row duplication during the webhook→poll cutover.
  return new Date().toISOString();
}

interface ListedMessage {
  id: string;
  receivedDateTime: string;
}

async function listSince(sinceIso: string, ctx: InvocationContext): Promise<ListedMessage[]> {
  const client = getGraphClient();
  const out: ListedMessage[] = [];
  let url:
    | string
    | null = `/me/mailFolders('inbox')/messages?$filter=receivedDateTime gt ${sinceIso}&$orderby=receivedDateTime asc&$top=${PAGE_SIZE}&$select=id,receivedDateTime`;

  while (url) {
    const res: { value: ListedMessage[]; '@odata.nextLink'?: string } = await client
      .api(url)
      .get();
    out.push(...(res.value ?? []));
    url = res['@odata.nextLink'] ?? null;
    if (out.length >= 200) {
      ctx.warn(`pollMail: capping batch at ${out.length}; remaining will be picked up next tick`);
      break;
    }
  }
  return out;
}

export async function pollMailHandler(ctx: InvocationContext): Promise<{
  processed: number;
  skipped: number;
  bills: number;
  tax: number;
  duplicates: number;
  lastSeenIso: string;
}> {
  const state = await loadState();
  ctx.log(`pollMail: cursor=${state.lastSeenIso} recent=${state.recentIds.length}`);

  const messages = await listSince(state.lastSeenIso, ctx);
  if (messages.length === 0) {
    return { processed: 0, skipped: 0, bills: 0, tax: 0, duplicates: 0, lastSeenIso: state.lastSeenIso };
  }
  ctx.log(`pollMail: ${messages.length} message(s) since ${state.lastSeenIso}`);

  const recent = new Set(state.recentIds);
  let processed = 0;
  let skipped = 0;
  let bills = 0;
  let tax = 0;
  let duplicates = 0;

  for (const msg of messages) {
    if (recent.has(msg.id)) {
      skipped += 1;
      // Still advance cursor so we don't keep re-listing this row.
      if (msg.receivedDateTime > state.lastSeenIso) {
        state.lastSeenIso = msg.receivedDateTime;
        await saveState(state);
      }
      continue;
    }

    try {
      const result = await processMessage(msg.id, ctx);
      processed += 1;
      if (result.billLogged) bills += 1;
      if (result.taxLogged) tax += 1;
      if (result.duplicate) duplicates += 1;
    } catch (err) {
      ctx.error(`pollMail: processMessage(${msg.id}) failed`, err);
      // Don't update cursor for this message — try again next tick.
      continue;
    }

    // Commit progress: cursor advances, id is remembered.
    state.recentIds.push(msg.id);
    if (state.recentIds.length > RECENT_ID_CAP) {
      state.recentIds.splice(0, state.recentIds.length - RECENT_ID_CAP);
    }
    if (msg.receivedDateTime > state.lastSeenIso) {
      state.lastSeenIso = msg.receivedDateTime;
    }
    await saveState(state);
  }

  ctx.log(
    `pollMail: processed=${processed} skipped=${skipped} bills=${bills} tax=${tax} dup=${duplicates} cursor=${state.lastSeenIso}`,
  );
  return { processed, skipped, bills, tax, duplicates, lastSeenIso: state.lastSeenIso };
}
