/**
 * First-time auth setup.
 *
 * Run: `npm run auth:setup`
 * - Spins up a localhost listener on :8000
 * - Prints a URL to open in your browser
 * - You sign in to your Microsoft account, consent to the scopes
 * - Microsoft redirects back to localhost with a code
 * - We exchange code → refresh_token, save to .env
 *
 * After this, the agent runs entirely on the refresh token (silent).
 */

import 'dotenv/config';
import http from 'http';
import { randomBytes } from 'crypto';
import { exchangeCodeForTokens, getAuthorizeUrl } from '../src/auth/tokens';
import { promises as fs } from 'fs';
import path from 'path';

const PORT = 8000;
const REDIRECT_URI = `http://localhost:${PORT}/auth/callback`;

async function main() {
  const state = randomBytes(16).toString('hex');
  const url = getAuthorizeUrl(REDIRECT_URI, state);

  console.log('\nOpen this URL in your browser to authorize:\n');
  console.log(url);
  console.log('\nWaiting for callback on http://localhost:8000/auth/callback ...\n');

  const code: string = await new Promise((resolve, reject) => {
    const server = http.createServer(async (req, res) => {
      if (!req.url?.startsWith('/auth/callback')) {
        res.writeHead(404).end();
        return;
      }
      const params = new URL(req.url, `http://localhost:${PORT}`).searchParams;
      const returnedState = params.get('state');
      const code = params.get('code');
      const error = params.get('error');

      if (error) {
        res.writeHead(400).end(`Error: ${error} - ${params.get('error_description')}`);
        server.close();
        reject(new Error(error));
        return;
      }
      if (returnedState !== state) {
        res.writeHead(400).end('State mismatch');
        server.close();
        reject(new Error('state mismatch'));
        return;
      }
      if (!code) {
        res.writeHead(400).end('No code');
        server.close();
        reject(new Error('no code'));
        return;
      }

      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end('<h1>Authorized.</h1><p>You can close this tab.</p>');
      server.close();
      resolve(code);
    });
    server.listen(PORT);
  });

  const tokens = await exchangeCodeForTokens(code, REDIRECT_URI);
  console.log('\n✓ Got tokens. Refresh token expires in ~90 days of inactivity.');

  const envPath = path.join(process.cwd(), '.env');
  let content = '';
  try {
    content = await fs.readFile(envPath, 'utf-8');
  } catch {
    content = await fs.readFile(path.join(process.cwd(), '.env.example'), 'utf-8');
  }

  content = upsert(content, 'GRAPH_REFRESH_TOKEN', tokens.refreshToken);
  content = upsert(content, 'GRAPH_ACCESS_TOKEN', tokens.accessToken);
  content = upsert(content, 'TOKEN_EXPIRES_AT', String(tokens.expiresAt));
  await fs.writeFile(envPath, content);
  console.log(`✓ Wrote tokens to ${envPath}`);
  console.log('\nNext: run `npm run subscribe` to create the webhook subscription.');
}

function upsert(content: string, key: string, value: string): string {
  const line = `${key}=${value}`;
  const re = new RegExp(`^${key}=.*$`, 'm');
  return re.test(content) ? content.replace(re, line) : `${content.trimEnd()}\n${line}\n`;
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
