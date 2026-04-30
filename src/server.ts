/**
 * Express entrypoint for the Paperclip-orchestrated deploy.
 *
 * No internal scheduler — all cadence is owned by Paperclip's http adapter,
 * which posts to the /jobs/* routes below on a per-agent intervalSec.
 *
 * Routes:
 *   GET  /healthz             — unauthenticated liveness
 *   POST /jobs/poll-mail      — Graph poll → Bill Hunter + Tax Archivist
 *   POST /jobs/statements     — OneDrive statements/inbox sweep
 *   POST /jobs/rentals        — OneDrive rentals/inbox sweep
 *   POST /jobs/daily-tick     — hourly heartbeat → reconcile (07 UTC) + summary (13 UTC)
 *
 * All /jobs/* routes require `Authorization: Bearer ${BH_JOB_TOKEN}`.
 */

import 'dotenv/config';
import express from 'express';
import type { InvocationContext, Timer } from '@azure/functions';

import { statementsHandler } from './handlers/statements';
import { rentalsHandler } from './handlers/rentals';
import { pollMailHandler } from './handlers/poll-mail';
import { dailyTickHandler } from './handlers/daily-tick';

const PORT = Number(process.env.PORT || 8080);
const JOB_TOKEN = process.env.BH_JOB_TOKEN;
if (!JOB_TOKEN) {
  console.error('BH_JOB_TOKEN is required');
  process.exit(1);
}

function makeCtx(name: string): InvocationContext {
  const prefix = `[${name}]`;
  const log: any = (...args: unknown[]) => console.log(prefix, ...args);
  log.info = log;
  log.warn = (...args: unknown[]) => console.warn(prefix, ...args);
  log.error = (...args: unknown[]) => console.error(prefix, ...args);
  log.verbose = log;
  log.debug = log;
  log.trace = log;
  return {
    invocationId: Math.random().toString(36).slice(2),
    functionName: name,
    extraInputs: { get: () => undefined, set: () => undefined },
    extraOutputs: { get: () => undefined, set: () => undefined },
    log,
    info: log,
    warn: log.warn,
    error: log.error,
    debug: log.debug,
    trace: log.trace,
    options: {} as never,
  } as unknown as InvocationContext;
}

function fakeTimer(): Timer {
  return {
    schedule: { adjustForDST: false },
    scheduleStatus: { last: '', next: '', lastUpdated: '' },
    isPastDue: false,
  } as Timer;
}

const app = express();
app.use(express.json({ limit: '20mb' }));

app.get('/healthz', (_req, res) => {
  res.json({ ok: true, ts: new Date().toISOString() });
});

function requireBearer(req: express.Request, res: express.Response): boolean {
  const header = req.header('authorization') || '';
  const expected = `Bearer ${JOB_TOKEN}`;
  if (header !== expected) {
    res.status(401).json({ error: 'unauthorized' });
    return false;
  }
  return true;
}

type JobName = 'pollMail' | 'statements' | 'rentals' | 'dailyTick';

async function runJob(name: JobName): Promise<unknown> {
  const ctx = makeCtx(name);
  switch (name) {
    case 'pollMail':
      return pollMailHandler(ctx);
    case 'statements':
      await statementsHandler(fakeTimer(), ctx);
      return { ok: true };
    case 'rentals':
      await rentalsHandler(fakeTimer(), ctx);
      return { ok: true };
    case 'dailyTick':
      return dailyTickHandler(ctx);
  }
}

function mountJob(path: string, name: JobName): void {
  app.post(path, async (req, res) => {
    if (!requireBearer(req, res)) return;
    try {
      const result = await runJob(name);
      res.json({ job: name, result });
    } catch (err) {
      console.error(`[${name}] crash`, err);
      res.status(500).json({ job: name, error: String(err) });
    }
  });
}

mountJob('/jobs/poll-mail', 'pollMail');
mountJob('/jobs/statements', 'statements');
mountJob('/jobs/rentals', 'rentals');
mountJob('/jobs/daily-tick', 'dailyTick');

app.listen(PORT, () => {
  console.log(`bill-hunter listening on :${PORT} (${Object.keys({ pollMail: 0, statements: 0, rentals: 0, dailyTick: 0 }).join(', ')})`);
});
