/**
 * Quick smoke test: fetch the 10 most recent emails to verify auth works.
 * Run: `npm run test:fetch`
 */

import 'dotenv/config';
import { getGraphClient } from '../src/graph/client';

async function main() {
  const client = getGraphClient();
  const me = await client.api('/me').get();
  console.log(`Authenticated as: ${me.displayName} <${me.userPrincipalName}>`);

  const res = await client
    .api('/me/messages')
    .top(10)
    .select('subject,from,receivedDateTime,hasAttachments')
    .get();

  console.log(`\nLast 10 messages:`);
  for (const m of res.value) {
    console.log(
      `  ${m.receivedDateTime} | ${m.from?.emailAddress?.address ?? '?'} | ${m.subject ?? '(no subject)'}${m.hasAttachments ? ' 📎' : ''}`,
    );
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
