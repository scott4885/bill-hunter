/**
 * Hourly tick driven by Paperclip's http adapter (intervalSec=3600).
 *
 * Replaces three separate node-cron jobs. On each fire, checks whether
 * the daily reconcile (07:00 UTC) and daily summary (13:00 UTC) have
 * already run for the current UTC date — if not and the gate hour has
 * passed, runs them and records the date.
 *
 * State persisted next to the poll-mail cursor:
 *   {
 *     "lastReconcileDate": "2026-04-29",
 *     "lastSummaryDate":   "2026-04-29"
 *   }
 *
 * Idempotent: rerunning within the same UTC day is a no-op.
 */

import type { InvocationContext, Timer } from '@azure/functions';
import { promises as fs } from 'node:fs';
import { dirname } from 'node:path';
import { reconcileHandler } from './reconcile';
import { summaryHandler } from './summary';

const STATE_FILE = process.env.BH_DAILY_STATE_FILE || '/state/daily-tick.json';
const RECONCILE_HOUR_UTC = 7;
const SUMMARY_HOUR_UTC = 13;

interface DailyState {
  lastReconcileDate: string;
  lastSummaryDate: string;
}

async function loadState(): Promise<DailyState> {
  try {
    const raw = await fs.readFile(STATE_FILE, 'utf8');
    const parsed = JSON.parse(raw) as Partial<DailyState>;
    return {
      lastReconcileDate: parsed.lastReconcileDate ?? '',
      lastSummaryDate: parsed.lastSummaryDate ?? '',
    };
  } catch (err: any) {
    if (err?.code !== 'ENOENT') throw err;
    return { lastReconcileDate: '', lastSummaryDate: '' };
  }
}

async function saveState(state: DailyState): Promise<void> {
  await fs.mkdir(dirname(STATE_FILE), { recursive: true });
  await fs.writeFile(STATE_FILE, JSON.stringify(state, null, 2), 'utf8');
}

function utcDateString(d: Date): string {
  return d.toISOString().slice(0, 10);
}

function fakeTimer(): Timer {
  return {
    schedule: { adjustForDST: false },
    scheduleStatus: { last: '', next: '', lastUpdated: '' },
    isPastDue: false,
  } as Timer;
}

export async function dailyTickHandler(ctx: InvocationContext): Promise<{
  ranReconcile: boolean;
  ranSummary: boolean;
  utcDate: string;
  utcHour: number;
}> {
  const state = await loadState();
  const now = new Date();
  const utcDate = utcDateString(now);
  const utcHour = now.getUTCHours();

  let ranReconcile = false;
  let ranSummary = false;

  if (utcHour >= RECONCILE_HOUR_UTC && state.lastReconcileDate !== utcDate) {
    try {
      await reconcileHandler(fakeTimer(), ctx);
      state.lastReconcileDate = utcDate;
      await saveState(state);
      ranReconcile = true;
    } catch (err) {
      ctx.error('dailyTick: reconcile failed', err);
    }
  }

  if (utcHour >= SUMMARY_HOUR_UTC && state.lastSummaryDate !== utcDate) {
    try {
      await summaryHandler(fakeTimer(), ctx);
      state.lastSummaryDate = utcDate;
      await saveState(state);
      ranSummary = true;
    } catch (err) {
      ctx.error('dailyTick: summary failed', err);
    }
  }

  ctx.log(
    `dailyTick: utc=${utcDate}T${String(utcHour).padStart(2, '0')} reconcile=${ranReconcile} summary=${ranSummary}`,
  );
  return { ranReconcile, ranSummary, utcDate, utcHour };
}
