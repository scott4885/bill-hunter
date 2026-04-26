/**
 * Express + node-cron entrypoint for self-hosted deploy (Coolify, Docker, etc).
 *
 * Mirrors the Azure Functions v4 setup:
 *   - HTTP webhook on POST /api/webhook (Graph notifications)
 *   - Daily timer: subscription renewal (06:00 UTC)
 *   - Daily timer: reconciliation (07:30 UTC)
 *   - Daily timer: Discord summary (13:00 UTC)
 *   - Periodic: scan statements inbox (every 15 min)
 *   - Periodic: scan rentals inbox (every 30 min)
 *
 * Azure registrations in handlers/*.ts are gated behind FUNCTIONS_WORKER_RUNTIME
 * so importing them here is a no-op for the SDK; we wire Express + cron ourselves.
 */

import 'dotenv/config';
import express from 'express';
import cron from 'node-cron';
import type { HttpRequest, HttpResponseInit, InvocationContext, Timer } from '@azure/functions';

import { webhookHandler } from './handlers/webhook';
import { summaryHandler } from './handlers/summary';
import { statementsHandler } from './handlers/statements';
import { rentalsHandler } from './handlers/rentals';
import { reconcileHandler } from './handlers/reconcile';
import { renewHandler } from './handlers/renew';

const PORT = Number(process.env.PORT || 8080);

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

const httpApp = express();
httpApp.use(express.json({ limit: '20mb' }));

httpApp.get('/healthz', (_req, res) => {
  res.json({ ok: true, ts: new Date().toISOString() });
});

httpApp.all('/api/webhook', async (req, res) => {
  const ctx = makeCtx('webhook');
  const fullUrl = `https://${req.headers.host ?? 'localhost'}${req.originalUrl}`;
  const httpReq = {
    url: fullUrl,
    method: req.method,
    headers: req.headers,
    json: async () => req.body,
  } as unknown as HttpRequest;

  try {
    const result: HttpResponseInit = await webhookHandler(httpReq, ctx);
    res.status(result.status ?? 200);
    if (result.headers) {
      for (const [k, v] of Object.entries(result.headers)) {
        res.setHeader(k, v as string);
      }
    }
    res.send(result.body ?? '');
  } catch (err) {
    console.error('[webhook] crash', err);
    res.status(500).send('internal error');
  }
});

interface CronJob {
  name: string;
  expr: string; // 5-field node-cron syntax (Azure has 6, dropped seconds)
  fn: (t: Timer, c: InvocationContext) => Promise<void>;
}

const JOBS: CronJob[] = [
  { name: 'renewSubscriptions', expr: '0 6 * * *', fn: renewHandler },
  { name: 'reconcile', expr: '30 7 * * *', fn: reconcileHandler },
  { name: 'dailySummary', expr: '0 13 * * *', fn: summaryHandler },
  { name: 'processStatements', expr: '*/15 * * * *', fn: statementsHandler },
  { name: 'processRentals', expr: '*/30 * * * *', fn: rentalsHandler },
];

for (const job of JOBS) {
  cron.schedule(
    job.expr,
    async () => {
      const ctx = makeCtx(job.name);
      try {
        await job.fn(fakeTimer(), ctx);
      } catch (err) {
        console.error(`[${job.name}] crash`, err);
      }
    },
    { timezone: 'UTC' },
  );
  console.log(`scheduled ${job.name}: ${job.expr} UTC`);
}

httpApp.listen(PORT, () => {
  console.log(`bill-hunter listening on :${PORT}`);
});
